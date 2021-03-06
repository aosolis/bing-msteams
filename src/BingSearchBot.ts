import * as config from "config";
import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as winston from "winston";
import * as moment from "moment";
import { sprintf } from "sprintf-js";
import * as escapeHtml from "escape-html";
import * as consts from "./constants";
import * as utils from "./utils";
import * as bing from "./BingSearchApi";
import { Strings } from "./locale/locale";

// =========================================================
// Bot Setup
// =========================================================

export class BingSearchBot extends builder.UniversalBot {

    private bingSearch: bing.BingSearchApi;

    constructor(
        public _connector: builder.IConnector,
        private botSettings: any,
    )
    {
        super(_connector, botSettings);
        this.set("persistConversationData", true);

        this.bingSearch = botSettings.bingSearch as bing.BingSearchApi;

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

            teamsConnector.onQuery("searchVideos", async (event, query, cb) => {
                try {
                    await this.handleVideosSearchQuery(event, query, cb);
                } catch (e) {
                    winston.error("Videos search handler failed", e);
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

    // Handle searching for news
    private async handleNewsSearchQuery(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let text = this.getQueryParameter(query, "text");
        let initialRun = !!this.getQueryParameter(query, "initialRun");
        utils.trackEvent(consts.TelemetryEvent.Compose,
            { type: "query", command: query.commandId, initialRun, skip: query.queryOptions.skip },
            event);

        let session = await utils.loadSessionAsync(this, event);

        // Handle settings coming in part of a query, as happens when we return a configuration response
        let incomingSettings = query.state;
        if (incomingSettings) {
            this.updateSettings(session, incomingSettings);
            text = "";
        }

        if (text) {
            let options = {} as bing.news.NewsSearchOptions;
            options.count = query.queryOptions.count;
            options.offset = query.queryOptions.skip;

            let searchResult = await this.bingSearch.searchNewsAsync(text, session.userData.clientId, options);
            if (searchResult.clientId && (searchResult.clientId !== session.userData.clientId)) {
                session.userData.clientId = searchResult.clientId;
            }

            let response = msteams.ComposeExtensionResponse.result("list")
                .attachments(searchResult.articles.map(article => this.createNewsResult(session, article)));
            cb(null, response.toResponse());
        } else if (initialRun) {
            let searchResult = await this.bingSearch.getNewsAsync(session.userData.clientId);
            if (searchResult.clientId && (searchResult.clientId !== session.userData.clientId)) {
                session.userData.clientId = searchResult.clientId;
            }

            let response = msteams.ComposeExtensionResponse.result("list")
                .attachments(searchResult.articles.map(article => this.createNewsResult(session, article)));
            cb(null, response.toResponse());
        } else {
            cb(null, this.createMessageResponse(session, Strings.error_news_notext));
        }
    }

    // Handle searching for videos
    private async handleVideosSearchQuery(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let text = this.getQueryParameter(query, "text");
        let initialRun = !!this.getQueryParameter(query, "initialRun");
        utils.trackEvent(consts.TelemetryEvent.Compose,
            { type: "query", command: query.commandId, initialRun, skip: query.queryOptions.skip },
            event);

        let session = await utils.loadSessionAsync(this, event);

        // Handle settings coming in part of a query, as happens when we return a configuration response
        let incomingSettings = query.state;
        if (incomingSettings) {
            this.updateSettings(session, incomingSettings);
            text = "";
        }

        if (text) {
            let options = {} as bing.videos.VideoSearchOptions;
            options.count = query.queryOptions.count;
            options.offset = query.queryOptions.skip;

            let searchResult = await this.bingSearch.searchVideosAsync(text, session.userData.clientId, options);
            if (searchResult.clientId && (searchResult.clientId !== session.userData.clientId)) {
                session.userData.clientId = searchResult.clientId;
            }

            let response = msteams.ComposeExtensionResponse.result("list")
                .attachments(searchResult.videos.map(video => this.createVideoResult(session, video)));
            cb(null, response.toResponse());
        } else if (initialRun) {
            let trendingVideos = await this.bingSearch.getTrendingVideosAsync(session.userData.clientId);
            if (trendingVideos.clientId && (trendingVideos.clientId !== session.userData.clientId)) {
                session.userData.clientId = trendingVideos.clientId;
            }

            let response = msteams.ComposeExtensionResponse.result("list")
                .attachments(trendingVideos.bannerTiles.map(video => this.createTrendingVideoResult(session, video)));
            cb(null, response.toResponse());
        } else {
            cb(null, this.createMessageResponse(session, Strings.error_videos_notext));
        }
    }

    // Handle compose extension query settings url callback
    private async handleQuerySettingsUrl(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        utils.trackEvent(consts.TelemetryEvent.Compose, { type: "querySettingsUrl" }, event);

        let session = await utils.loadSessionAsync(this, event);
        cb(null, this.createConfigurationResponse(session));
    }

    // Handle compose extension settings update callback
    private async handleSettingsUpdate(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        utils.trackEvent(consts.TelemetryEvent.Compose, { type: "settingsUpdate" }, event);

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

    private createNewsResult(session: builder.Session, article: bing.news.NewsArticle): msteams.ComposeExtensionAttachment {
        let card = new builder.ThumbnailCard()
            .title(`<a href="${escapeHtml(article.url)}">${escapeHtml(article.name)}</a>`)
            .text(`<p>${escapeHtml(article.description)}</p><p>${this.createNewsAttribution(article)}</p>`);
        let previewCard = new builder.ThumbnailCard()
            .title(`<span style="font-weight:600">${article.name}</span>`)
            .text(this.createNewsAttribution(article, true /*highlightRecent*/));

        // Add images if available
        if (article.image) {
            card.images([ new builder.CardImage().url(article.image.thumbnail.contentUrl) ]);
            previewCard.images([ new builder.CardImage().url(article.image.thumbnail.contentUrl) ]);
        }

        return {
            ...card.toAttachment(),
            preview: previewCard.toAttachment(),
        };
    }

    private createNewsAttribution(article: bing.news.NewsArticle, highlightRecent?: boolean): string {
        // Build the attributions line
        let info = [];
        if (article.provider && article.provider.length) {
            info.push(article.provider.map(provider => provider.name).join(", "));
        }
        if (article.datePublished) {
            let datePublished = moment.utc(article.datePublished);
            if (datePublished.isBefore(moment().subtract(1, "day"))) {
                info.push(datePublished.format("l"));
            } else if (datePublished.isBefore(moment().subtract(6, "hours"))) {
                info.push(datePublished.fromNow());
            } else {
                info.push(`<span style="color:#c50e2e">${datePublished.fromNow()}</span>`);
            }
        }
        return info.join(" | ");
    }

    private createVideoResult(session: builder.Session, video: bing.videos.Video): msteams.ComposeExtensionAttachment {
        // Build the attributions line
        let info = [];
        if (video.publisher && video.publisher.length) {
            info.push(video.publisher.map(publisher => publisher.name).join(", "));
        }
        if (video.duration) {
            let duration = moment.duration(video.duration);
            if (duration.hours() > 0) {
                info.push(sprintf("%d:%02d:%02d", duration.hours(), duration.minutes(), duration.seconds()));
            } else {
                info.push(sprintf("%02d:%02d", duration.minutes(), duration.seconds()));
            }
        }
        if (video.datePublished) {
            let datePublished = moment.utc(video.datePublished);
            if (datePublished.isBefore(moment().subtract(1, "days"))) {
                info.push(datePublished.format("l"));
            } else {
                info.push(datePublished.fromNow());
            }
        }

        let card = new builder.ThumbnailCard()
            .title(`<a href="${escapeHtml(video.hostPageUrl)}">${escapeHtml(video.name)}</a>`)
            .text(`<p>${escapeHtml(video.description)}</p><p>${info.join(" | ")}</p>`);
        let previewCard = new builder.ThumbnailCard()
            .title(`<span style="font-weight:600">${video.name}</span>`)
            .text(video.description);

        // Add images if available
        if (video.thumbnailUrl) {
            card.images([ new builder.CardImage().url(video.thumbnailUrl) ]);
            previewCard.images([ new builder.CardImage().url(video.thumbnailUrl) ]);
        }

        return {
            ...card.toAttachment(),
            preview: previewCard.toAttachment(),
        };
    }

    private createTrendingVideoResult(session: builder.Session, bannerTile: bing.videos.BannerTile): msteams.ComposeExtensionAttachment {
        let card = new builder.ThumbnailCard()
            .title(bannerTile.query.displayText)
            .text(`<p>${escapeHtml(bannerTile.image.headLine)}</p><p><a href="${escapeHtml(bannerTile.query.webSearchUrl)}">See more on Bing</a></p>`)
            .images([ new builder.CardImage().url(bannerTile.image.thumbnailUrl) ]);
        let previewCard = new builder.ThumbnailCard()
            .title(`<span style="font-weight:600">${escapeHtml(bannerTile.query.displayText)}</span>`)
            .text(bannerTile.image.headLine)
            .images([ new builder.CardImage().url(bannerTile.image.thumbnailUrl) ]);

        return {
            ...card.toAttachment(),
            preview: previewCard.toAttachment(),
        };
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
