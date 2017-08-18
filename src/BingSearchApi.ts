import * as http from "http";
import * as request from "request";
import * as querystring from "querystring";

// =========================================================
// Bing Search API
// =========================================================

export type SafeSearch = "off" | "moderate" | "strict";

export namespace news {
    export interface Image {
        url: string;
        thumbnail: Thumbnail;
    }

    export interface Thumbnail {
        contentUrl: string;
        width: number;
        height: number;
    }

    export interface Organization {
        name: string;
    }

    export interface NewsArticle {
        name: string;
        url: string;
        description: string;
        image: Image;
        datePublished: string;
        provider: Organization[];
    }

    export interface NewsSearchResult {
        totalEstimatedMatches?: number;
        articles: NewsArticle[];
        clientId: string;
    }

    export interface NewsSearchOptions {
        count?: number;
        offset?: number;
        mkt?: string;
        safeSearch?: SafeSearch;
    }
}

export namespace videos {
    export interface Video {
        videoId: string;
        name: string;
        description: string;
        publisher: Publisher[];
        thumbnailUrl: string;
        hostPageUrl: string;
        contentUrl: string;
        webSearchUrl: string;
        duration: string;
        datePublished: string;
        creator: Publisher;
    }

    export interface Publisher {
        name: string;
    }

    export interface VideoSearchResult {
        totalEstimatedMatches?: number;
        nextOffsetAddCount?: number;
        videos: Video[];
        clientId: string;
    }

    export interface TileImage {
        contentUrl: string;
        headLine: string;
        thumbnailUrl: string;
    }

    export interface VideoQuery {
        displayText: string;
        text: string;
        webSearchUrl: string;
    }

    export interface BannerTile {
        image: TileImage;
        query: VideoQuery;
    }

    export interface TrendingVideosResult {
        clientId: string;
        bannerTiles: BannerTile[];
    }

    type VideoFreshness = "day" | "week" | "month";

    export interface VideoSearchOptions {
        count?: number;
        offset?: number;
        mkt?: string;
        safeSearch?: SafeSearch;
        freshness?: VideoFreshness;
        id?: string;
    }
}

// Service endpoints
const topNewsEndpoint = "https://api.cognitive.microsoft.com/bing/v5.0/news";
const newsSearchEndpoint = "https://api.cognitive.microsoft.com/bing/v5.0/news/search";
const videosSearchEndpoint = "https://api.cognitive.microsoft.com/bing/v5.0/videos/search";
const trendingVideosEndpoint = "https://api.cognitive.microsoft.com/bing/v5.0/videos/trending";

export class BingSearchApi {

    constructor(
        private accessKey: string,
    )
    {
    }

    public async searchNewsAsync(query: string, clientId: string, searchOptions?: news.NewsSearchOptions): Promise<news.NewsSearchResult> {
        return new Promise<news.NewsSearchResult>((resolve, reject) => {
            let qsp: any = { q: query };
            if (searchOptions) {
                qsp = { ...searchOptions, ...qsp };
            }

            let options = {
                url: `${newsSearchEndpoint}?${querystring.stringify(qsp)}`,
                headers: {
                    "Ocp-Apim-Subscription-Key": this.accessKey,
                    "X-MSEdge-ClientID": clientId,
                },
                json: true,
            };
            request.get(options, (err, res: http.IncomingMessage, body) => {
                if (err) {
                    reject(err);
                } else if (res.statusCode !== 200) {
                    reject(new Error(res.statusMessage));
                } else {
                    resolve({
                        totalEstimatedMatches: body.totalEstimatedMatches,
                        clientId: res.headers["X-MSEdge-ClientID"],
                        articles: body.value,
                    });
                }
            });
        });
    }

    public async getNewsAsync(clientId: string): Promise<news.NewsSearchResult> {
        return new Promise<news.NewsSearchResult>((resolve, reject) => {
            let options = {
                url: `${topNewsEndpoint}`,
                headers: {
                    "Ocp-Apim-Subscription-Key": this.accessKey,
                    "X-MSEdge-ClientID": clientId,
                },
                json: true,
            };
            request.get(options, (err, res: http.IncomingMessage, body) => {
                if (err) {
                    reject(err);
                } else if (res.statusCode !== 200) {
                    reject(new Error(res.statusMessage));
                } else {
                    resolve({
                        clientId: res.headers["X-MSEdge-ClientID"],
                        articles: body.value,
                    });
                }
            });
        });
    }

    public async searchVideosAsync(query: string, clientId: string, searchOptions?: videos.VideoSearchOptions): Promise<videos.VideoSearchResult> {
        return new Promise<videos.VideoSearchResult>((resolve, reject) => {
            let qsp: any = { q: query };
            if (searchOptions) {
                qsp = { ...searchOptions, ...qsp };
            }

            let options = {
                url: `${videosSearchEndpoint}?${querystring.stringify(qsp)}`,
                headers: {
                    "Ocp-Apim-Subscription-Key": this.accessKey,
                    "X-MSEdge-ClientID": clientId,
                },
                json: true,
            };
            request.get(options, (err, res: http.IncomingMessage, body) => {
                if (err) {
                    reject(err);
                } else if (res.statusCode !== 200) {
                    reject(new Error(res.statusMessage));
                } else {
                    resolve({
                        totalEstimatedMatches: body.totalEstimatedMatches,
                        nextOffsetAddCount: body.nextOffsetAddCount,
                        clientId: res.headers["X-MSEdge-ClientID"],
                        videos: body.value,
                    });
                }
            });
        });
    }

    public async getTrendingVideosAsync(clientId: string): Promise<videos.TrendingVideosResult> {
        return new Promise<videos.TrendingVideosResult>((resolve, reject) => {
            let options = {
                url: `${trendingVideosEndpoint}`,
                headers: {
                    "Ocp-Apim-Subscription-Key": this.accessKey,
                    "X-MSEdge-ClientID": clientId,
                },
                json: true,
            };
            request.get(options, (err, res: http.IncomingMessage, body) => {
                if (err) {
                    reject(err);
                } else if (res.statusCode !== 200) {
                    reject(new Error(res.statusMessage));
                } else {
                    resolve({
                        clientId: res.headers["X-MSEdge-ClientID"],
                        bannerTiles: body.bannerTiles,
                    });
                }
            });
        });
    }
}
