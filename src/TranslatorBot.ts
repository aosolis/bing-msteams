import * as config from "config";
import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as winston from "winston";
import * as utils from "./utils";
import { TranslatorApi, TranslationResult } from "./TranslatorApi";
import { Strings } from "./locale/locale";

// =========================================================
// Bot Setup
// =========================================================

const maxTranslationHistory = 5;

export class TranslatorBot extends builder.UniversalBot {

    private loadSessionAsync: {(address: builder.IEvent): Promise<builder.Session>};
    private translator: TranslatorApi;

    constructor(
        public _connector: builder.IConnector,
        private botSettings: any,
    )
    {
        super(_connector, botSettings);
        this.set("persistConversationData", true);

        this.translator = botSettings.translator as TranslatorApi;

        // Handle invoke events
        this.loadSessionAsync = (event) => {
            return new Promise((resolve, reject) => {
                this.loadSession(event.address, (err: any, session: builder.Session) => {
                    if (err) {
                        winston.error("Failed to load session", { error: err, address: event.address });
                        reject(err);
                    } else if (!session) {
                        winston.error("Loaded null session", { address: event.address });
                        reject(new Error("Failed to load session"));
                    } else {
                        let locale = utils.getLocale(event);
                        if (locale) {
                            (session as any)._locale = locale;
                            session.localizer.load(locale, (err2) => {
                                // Log but resolve session anyway
                                winston.error(`Failed to load localizer for ${locale}`, err2);
                                resolve(session);
                            });
                        } else {
                            resolve (session);
                        }
                    }
                });
            });
        };

        // Handle compose extension invokes
        let teamsConnector = this._connector as msteams.TeamsChatConnector;
        if (teamsConnector.onQuery) {
            teamsConnector.onQuery("translate", async (event, query, cb) => {
                try {
                    await this.handleTranslateQuery(event, query, cb);
                } catch (e) {
                    winston.error("Translate handler failed", e);
                    cb(e, null, 500);
                }
            });
        }
        if (teamsConnector.onQuerySettingsUrl) {
            teamsConnector.onQuerySettingsUrl(async (event, query, cb) => {
                try {
                    await this.handleQuerySettingsUrl(event, query, cb);
                } catch (e) {
                    winston.error("Query settings url handler failed", e);
                    cb(e, null, 500);
                }
            });
        }
        if (teamsConnector.onSettingsUpdate) {
            teamsConnector.onSettingsUpdate(async (event, query, cb) => {
                try {
                    await this.handleSettingsUpdate(event, query, cb);
                } catch (e) {
                    winston.error("Settings update handler failed", e);
                    cb(e, null, 500);
                }
            });
        }

        // Handle generic invokes
        teamsConnector.onInvoke(async (event, cb) => {
            try {
                await this.onInvoke(event, cb);
            } catch (e) {
                winston.error("Invoke handler failed", e);
                cb(e, null, 500);
            }
        });

        // Register default dialog for testing
        this.dialog("/", async (session) => {
            let result = await this.translator.translateText(session.message.text, "it");
            session.endDialog(result[0].translatedText);
        });
    }

    // Handle compose extension query invocation
    private async handleTranslateQuery(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let session = await this.loadSessionAsync(event);

        let text = this.getQueryParameter(query, "text");
        let initialRun = !!this.getQueryParameter(query, "initialRun");

        // Handle settings coming in part of a query, as happens when we return a configuration response
        let incomingSettings = query.state;
        if (incomingSettings) {
            this.updateSettings(session, incomingSettings);
            text = "";
        }

        let translationLanguages = this.getTranslationLanguages(session);

        if ((text === "settings") && config.get("features.allowConfigurationViaQuery")) {
            // Provide a way to get to settings for client versions that don't yet support canUpdateConfiguration
            cb(null, this.createConfigurationResponse(session, translationLanguages));
        } else if (text) {
            // We got text, translate it
            try {
                let translations = await this.translator.translateText(text, translationLanguages);
                let response = msteams.ComposeExtensionResponse.result("list")
                    .attachments(translations
                        .filter(translation => translation.from !== translation.to)
                        .map(translation => this.createTranslationResult(session, translation)));
                cb(null, response.toResponse());
            } catch (e) {
                winston.error("Failed to get translations", e);
                cb(null, this.createMessageResponse(session, Strings.error_translation));
            }
        } else if (initialRun) {
            // Show the last few items in translation history, if present, otherwise instruction text
            let translationHistory = session.userData.translationHistory || [];
            if (translationHistory.length) {
                let response = msteams.ComposeExtensionResponse.result("list")
                    .attachments(translationHistory
                        .filter(translation => translation.from !== translation.to)
                        .map(translation => this.createTranslationHistoryResult(session, translation)));
                cb(null, response.toResponse());
            } else {
                cb(null, this.createMessageResponse(session, Strings.error_notext));
            }
        } else {
            // Default to showing instruction text
            cb(null, this.createMessageResponse(session, Strings.error_notext));
        }
    }

    // Handle compose extension query settings url callback
    private async handleQuerySettingsUrl(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let session = await this.loadSessionAsync(event);
        cb(null, this.createConfigurationResponse(session));
    }

    // Handle compose extension settings update callback
    private async handleSettingsUpdate(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let session = await this.loadSessionAsync(event);
        let incomingSettings = query.state;
        if (incomingSettings) {
            this.updateSettings(session, incomingSettings);
        }

        // Return a response. MS Teams doesn't care what the response is, so just return an empty message.
        cb(null, msteams.ComposeExtensionResponse.message().text("").toResponse());
    }

    // Handle compose extension item selected callback
    private async handleSelectItem(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let session = await this.loadSessionAsync(event);

        let invokeEvent = event as msteams.IInvokeEvent;
        let translation = invokeEvent.value as TranslationResult;

        // Store last few translations so we can show them in initial run
        let translationHistory: TranslationResult[] = session.userData.translationHistory || [];
        let existingItemIndex = translationHistory.findIndex(val =>
            (val.text.toLocaleUpperCase() === translation.text.toLocaleUpperCase()) &&
            (val.translatedText.toLocaleUpperCase() === translation.translatedText.toLocaleUpperCase()));
        if (existingItemIndex >= 0) {
            translationHistory.splice(existingItemIndex, 1);
        }
        translationHistory.unshift(translation);
        if (translationHistory.length > maxTranslationHistory) {
            translationHistory = translationHistory.slice(0, maxTranslationHistory);
        }
        session.userData.translationHistory = translationHistory;
        session.save().sendBatch();

        // This callback can only return a card -- right now it does not support returning a message response
        // To show an error message, consider returning a card with error text.
        cb(null, msteams.ComposeExtensionResponse.result("list")
            .attachments([{
                ...this.createTranslationCard(session, translation),
                // Current FE implementation still expects a title property for the preview, so return a dummy preview item
                // We can remove the line below when the issue is resolved
                preview: new builder.ThumbnailCard()
                    .title(" ")
                    .text(" ")
                    .toAttachment(),
            }])
            .toResponse());
    }

    // Handle other invokes
    private async onInvoke(event: builder.IEvent, cb: (err: Error, body: any, statusCode?: number) => void): Promise<void> {
        let invokeEvent = event as msteams.IInvokeEvent;
        let eventName = invokeEvent.name;

        switch (eventName) {
            case "composeExtension/selectItem":
                await this.handleSelectItem(event, null, (err, result, statusCode) => {
                    cb(err, result, statusCode);
                });
                break;

            default:
                let unrecognizedEvent = `Unrecognized event name: ${eventName}`;
                winston.error(unrecognizedEvent);
                cb(new Error(unrecognizedEvent), null, 500);
                break;
        }
    }

    // Get compose extension response that lets the user configurfe Translator
    private createConfigurationResponse(session: builder.Session, translationLanguages?: string[]): msteams.IComposeExtensionResponse {
        translationLanguages = translationLanguages || this.getTranslationLanguages(session);
        let baseUri = config.get("app.baseUri");
        let configPage = session.gettext(Strings.configure_page);
        let languages = translationLanguages.join(",");
        let response = msteams.ComposeExtensionResponse.config().actions([
            builder.CardAction.openUrl(session, `${baseUri}/html/${configPage}?languages=${languages}`, Strings.configure_text),
        ]);
        return response.toResponse();
    }

    // Create compose extension response that shows a message
    private createMessageResponse(session: builder.Session, text: string): msteams.IComposeExtensionResponse {
        let response = msteams.ComposeExtensionResponse.message()
            .text(session.gettext(text));
        return response.toResponse();
    }

    // Create a compose extension result from a translation
    private createTranslationResult(session: builder.Session, translation: TranslationResult): msteams.ComposeExtensionAttachment {
        let card: msteams.ComposeExtensionAttachment = this.createTranslationCard(session, translation);
        card.preview = new builder.ThumbnailCard()
            .title(translation.translatedText)
            .text(session.gettext(translation.to))
            // Attach a tap action to the preview card, so we get a selectItem callback
            .tap(new builder.CardAction(session)
                .type("invoke")
                .value(JSON.stringify(translation)))
            .toAttachment();
        return card;
    }

    // Create a compose extension result from a translation history item
    private createTranslationHistoryResult(session: builder.Session, translation: TranslationResult): msteams.ComposeExtensionAttachment {
        let card: msteams.ComposeExtensionAttachment = this.createTranslationCard(session, translation);
        card.preview = new builder.ThumbnailCard()
            .title(translation.translatedText)
            .text(translation.text)
            .tap(new builder.CardAction(session)
                .type("invoke")
                .value(JSON.stringify(translation)))
            .toAttachment();
        return card;
    }

    // Create the translation card that will be dropped into the conversation
    private createTranslationCard(session: builder.Session, translation: TranslationResult): builder.IAttachment {
        let fromLanguage = session.gettext(translation.from);
        let originalLabel = session.gettext(fromLanguage);
        let cardText =
            `<div style="font-size:1.6rem;font-weight:600;">${translation.translatedText}</div>
             <div style="margin-top:1.4rem;"><span style="text-decoration:underline;">${originalLabel}</span><br/>${translation.text}</div>`;
        return new builder.ThumbnailCard()
            .text(cardText)
            .toAttachment();
    }

    // Update Translator settings
    private updateSettings(session: builder.Session, state: string): void {
        // State is a comma-separated list of languages
        state = state || "";

        let supportedLangs = this.translator.getSupportedLanguages();
        let langs = state.split(",")
            .filter(lang => supportedLangs.find(i => i === lang));
        if (langs.length === 0) {
            langs = this.translator.getDefaultLanguages();
        }

        session.userData.languages = langs;
        session.save().sendBatch();
    }

    // Get the languages we should translate into
    private getTranslationLanguages(session: builder.Session): string[] {
        return session.userData.languages || this.translator.getDefaultLanguages();
    }

    // Return the value of the specified query parameter
    private getQueryParameter(query: msteams.ComposeExtensionQuery, name: string): string {
        let matchingParams = (query.parameters || []).filter(p => p.name === name);
        return matchingParams.length ? matchingParams[0].value : "";
    }
}
