import * as http from "http";
import * as request from "request";
import * as querystring from "querystring";

// =========================================================
// Bing Search API
// =========================================================

const newsSearchEndpoint = "https://api.cognitive.microsoft.com/bing/v5.0/news/search";

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
    totalEstimatedMatches: number;
    articles: NewsArticle[];
    clientId: string;
}

export interface NewsSearchOptions {
    count?: number;
    offset?: number;
    mkt?: string;
}

export class BingSearchApi {

    constructor(
        private accessKey: string,
    )
    {
    }

    public async searchNewsAsync(query: string, clientId: string, searchOptions?: NewsSearchOptions): Promise<NewsSearchResult> {
        return new Promise<NewsSearchResult>((resolve, reject) => {
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

}
