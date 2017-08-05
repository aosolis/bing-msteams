import * as config from "config";
import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as winston from "winston";
import * as moment from "moment";
import * as escapeHtml from "escape-html";
import * as utils from "./utils";
import { BingSearchApi } from "./BingSearchApi";
import { Strings } from "./locale/locale";

// =========================================================
// Bot Setup
// =========================================================

export class BingSearchBot extends builder.UniversalBot {

    private bingSearch: BingSearchApi;

    constructor(
        public _connector: builder.IConnector,
        private botSettings: any,
    )
    {
        super(_connector, botSettings);
        this.set("persistConversationData", true);

        this.bingSearch = botSettings.bingSearch as BingSearchApi;

        // Handle compose extension invokes
        let teamsConnector = this._connector as msteams.TeamsChatConnector;
        if (teamsConnector.onQuery) {
            teamsConnector.onQuery("searchNews", async (event, query, cb) => {
                try {
                    await this.handleNewsSearchQuery(event, query, cb);
                } catch (e) {
                    winston.error("News search handler failed", e);
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
            session.endDialog("Hi!");
        });
    }

    // Handle compose extension query invocation
    private async handleNewsSearchQuery(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let session = await utils.loadSessionAsync(this, event);

        let text = this.getQueryParameter(query, "text");
        let initialRun = !!this.getQueryParameter(query, "initialRun");

        // Handle settings coming in part of a query, as happens when we return a configuration response
        let incomingSettings = query.state;
        if (incomingSettings) {
            this.updateSettings(session, incomingSettings);
            text = "";
        }

        if (text) {
            let searchResult = await this.bingSearch.searchNewsAsync(text, session.userData.clientId);
            if (searchResult.clientId && (searchResult.clientId !== session.userData.clientId)) {
                session.userData.clientId = searchResult.clientId;
            }

            let response = msteams.ComposeExtensionResponse.result("list")
                .attachments(searchResult.articles.map(article => {
                    // Build the attributions line
                    let attributions = [];
                    if (article.provider && article.provider.length) {
                        attributions.push(article.provider.map(provider => provider.name).join(", "));
                    }
                    if (article.datePublished) {
                        attributions.push(moment.utc(article.datePublished).fromNow());
                    }

                    let card = new builder.ThumbnailCard(session)
                        .title(`<a href="${escapeHtml(article.url)}">${escapeHtml(article.name)}</a>`)
                        .text(`<p>${escapeHtml(article.description)}</p><p>${attributions.join(" | ")}</p>`);
                    let previewCard = new builder.ThumbnailCard(session)
                        .title(article.name)
                        .text(article.description);

                    // Add images if available
                    if (article.image) {
                        card.images([ new builder.CardImage(session).url(article.image.thumbnail.contentUrl) ]);
                        previewCard.images([ new builder.CardImage(session).url(article.image.thumbnail.contentUrl) ]);
                    }

                    return {
                        ...card.toAttachment(),
                        preview: previewCard.toAttachment(),
                    };
                }));
            cb(null, response.toResponse());
        } else if (initialRun) {
            cb(null, this.createMessageResponse(session, Strings.error_notext));
        } else {
            cb(null, this.createMessageResponse(session, Strings.error_notext));
        }
    }

    // Handle compose extension query settings url callback
    private async handleQuerySettingsUrl(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let session = await utils.loadSessionAsync(this, event);
        cb(null, this.createConfigurationResponse(session));
    }

    // Handle compose extension settings update callback
    private async handleSettingsUpdate(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let session = await utils.loadSessionAsync(this, event);
        let incomingSettings = query.state;
        if (incomingSettings) {
            this.updateSettings(session, incomingSettings);
        }

        // Return a response. MS Teams doesn't care what the response is, so just return an empty message.
        cb(null, msteams.ComposeExtensionResponse.message().text("").toResponse());
    }

    // Handle other invokes
    private async onInvoke(event: builder.IEvent, cb: (err: Error, body: any, statusCode?: number) => void): Promise<void> {
        let invokeEvent = event as msteams.IInvokeEvent;
        let eventName = invokeEvent.name;

        switch (eventName) {
            default:
                let unrecognizedEvent = `Unrecognized event name: ${eventName}`;
                winston.error(unrecognizedEvent);
                cb(new Error(unrecognizedEvent), null, 500);
                break;
        }
    }

    // Get compose extension response that lets the user configure Bing Search
    private createConfigurationResponse(session: builder.Session, translationLanguages?: string[]): msteams.IComposeExtensionResponse {
        let baseUri = config.get("app.baseUri");
        let configPage = session.gettext(Strings.configure_page);
        let response = msteams.ComposeExtensionResponse.config().actions([
            builder.CardAction.openUrl(session, `${baseUri}/html/${configPage}`, Strings.configure_text),
        ]);
        return response.toResponse();
    }

    // Create compose extension response that shows a message
    private createMessageResponse(session: builder.Session, text: string): msteams.IComposeExtensionResponse {
        let response = msteams.ComposeExtensionResponse.message()
            .text(session.gettext(text));
        return response.toResponse();
    }

    // Update Bing Search settings
    private updateSettings(session: builder.Session, state: string): void {
        state = state || "";
        session.save().sendBatch();
    }

    // Return the value of the specified query parameter
    private getQueryParameter(query: msteams.ComposeExtensionQuery, name: string): string {
        let matchingParams = (query.parameters || []).filter(p => p.name === name);
        return matchingParams.length ? matchingParams[0].value : "";
    }
}
