import * as _ from "lodash";
import * as request from "request";
import * as xml2js from "xml2js";

// =========================================================
// Translator API
// =========================================================

// Access tokens last 10 minutes, but refresh every 9 minutes to be safe
const accessTokenLifetimeMs = 9 * 60 * 1000;

// Supported languages
const supportedLanguages: string[] = [
    "af",
    "ar",
    "bn",
    "bs-Latn",
    "bg",
    "ca",
    "zh-CHS",
    "zh-CHT",
    "hr",
    "cs",
    "da",
    "nl",
    "en",
    "et",
    "fj",
    "fil",
    "fi",
    "fr",
    "de",
    "el",
    "ht",
    "he",
    "hi",
    "hu",
    "id",
    "it",
    "ja",
    "tlh",
    "ko",
    "lv",
    "lt",
    "mg",
    "ms",
    "mt",
    "no",
    "fa",
    "pl",
    "pt",
    "ro",
    "ru",
    "sm",
    "sr-Cyrl",
    "sr-Latn",
    "sk",
    "sl",
    "es",
    "sv",
    "ty",
    "th",
    "to",
    "tr",
    "uk",
    "ur",
    "vi",
    "cy",
];

// Default languages
const defaultLanguages: string[] = [
    "en",
    "es",
    "fr",
    "it",
    "ar",
];

export interface TranslationResult {
    from?: string;
    text: string;
    to: string;
    translatedText: string;
}

export class TranslatorApi {

    private accessToken: string;
    private accessTokenExpiryTime: number;

    constructor(
        private accessKey: string,
    )
    {
    }

    // Translate text to the specified language
    public async translateText(text: string, to: string|string[], from?: string): Promise<TranslationResult[]> {
        if (!Array.isArray(to)) {
            to = [ to ];
        }

        return Promise.all(to.map((toLang) => {
            return this.translateTextWorker(text, toLang, from);
        }));
    }

    // Return supported languages
    public getSupportedLanguages(): string[] {
        return supportedLanguages.slice();
    }

    // Return default languages
    public getDefaultLanguages(): string[] {
        return defaultLanguages.slice();
    }

    // Translate text to the specified language
    private async translateTextWorker(text: string, to: string, from?: string): Promise<TranslationResult> {
        // Escape parameters
        let escapedText = text ? _.escape(text) : "";
        let escapedTo = to ? _.escape(to) : "en";
        let escapedFrom = from ? _.escape(from) : "";

        let body = `
<TranslateArrayRequest>
  <AppId />
  <From>${escapedFrom}</From>
  <Texts>
    <string xmlns="http://schemas.microsoft.com/2003/10/Serialization/Arrays">${escapedText}</string>
  </Texts>
  <To>${escapedTo}</To>
</TranslateArrayRequest>`;

        let url = "https://api.microsofttranslator.com/v2/http.svc/TranslateArray";
        let authHeader = await this.getAuthorizationHeader();
        let options: request.Options = {
            url: url,
            headers: {
                "Content-Type": "application/xml",
                "Authorization": authHeader,
            },
            body: body,
        };

        return new Promise<TranslationResult>((resolve, reject) => {
            request.post(options, (error, response, responseBody) => {
                if (error) {
                    reject(error);
                } else if (response.statusCode !== 200) {
                    reject(new Error(response.statusMessage));
                } else {
                    xml2js.parseString(responseBody as string, (parseError, result) => {
                        if (parseError) {
                            reject(parseError);
                        } else {
                            resolve({
                                from: result.ArrayOfTranslateArrayResponse.TranslateArrayResponse[0].From[0],
                                text: text,
                                to: to,
                                translatedText: result.ArrayOfTranslateArrayResponse.TranslateArrayResponse[0].TranslatedText[0],
                            });
                        }
                    });
                }
            });
        });
    }

    private async getAuthorizationHeader(): Promise<string> {
        if (!this.accessToken || this.isAccessTokenExpired()) {
            await this.refreshAccessToken();
        }

        return "Bearer " + this.accessToken;
    }

    private isAccessTokenExpired(): boolean {
        return new Date().valueOf() > this.accessTokenExpiryTime;
    }

    private async refreshAccessToken(): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            this.accessToken = null;
            this.accessTokenExpiryTime = 0;

            let options: request.Options = {
                url: "https://api.cognitive.microsoft.com/sts/v1.0/issueToken",
                headers: {
                    "Ocp-Apim-Subscription-Key": this.accessKey,
                },
                body: "",
            };

            request.post(options, (error, response, body) => {
                if (error) {
                    reject(error);
                } else if (response.statusCode !== 200) {
                    reject(new Error(response.statusMessage));
                } else {
                    this.accessToken = body as string;
                    this.accessTokenExpiryTime = new Date().valueOf() + accessTokenLifetimeMs;
                    resolve();
                }
            });
        });
    }

}
