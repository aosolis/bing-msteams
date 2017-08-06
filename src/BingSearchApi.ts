import * as http from "http";
import * as request from "request";
import * as querystring from "querystring";

// =========================================================
// Bing Search API
// =========================================================

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
    }
}

// Service endpoints
const topNewsEndpoint = "https://api.cognitive.microsoft.com/bing/v5.0/news";
const newsSearchEndpoint = "https://api.cognitive.microsoft.com/bing/v5.0/news/search";

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
}
