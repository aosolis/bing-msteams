import * as http from "http";
import * as request from "request";
import * as querystring from "querystring";

// =========================================================
// Bing Entity Search API
// =========================================================

export type SafeSearch = "off" | "moderate" | "strict";

export namespace entities {
    /** Defines additional information about an entity such as type hints. */
    export interface EntityPresentationInfo {
        /** The supported scenario. */
        entityScenario: string;
        /** A display version of the entity hint. For example, if entityTypeHints is Artist, this field may be set to American Singer. */
        entityTypeDisplayHint: string;
        /** A list of hints that indicate the entity's type. The list could contain a single hint such as Movie or a list of hints such as Place, LocalBusiness, Restaurant. Each successive hint in the array narrows the entity's type. */
        entityTypeHint: string[];
    }

    export interface Image {
        height: number;
        hostPageUrl: string;
        name: string;
        provider: Organization[];
        thumbnailUrl: string;
        width: number;
    }

    /** Defines a publisher. */
    export interface Organization {
        name: string;
        url: string;
    }

    /** Defines the license under which the text or photo may be used. */
    export interface License {
        name: string;
        url: string;
    }

    /** Defines a contractual rule for license attribution. */
    export interface LicenseAttribution {
        /** A type hint, which is set to LicenseAttribution. */
        _type: "LicenseAttribution";
        /** The license under which the content may be used. */
        license: License;
        /** The license to display next to the targeted field. For example, "Text under CC-BY-SA license". */
        licenseNotice: string;
        /** A Boolean value that determines whether the contents of the rule must be placed in close proximity to the field that the rule applies to. */
        mustBeCloseToContent: boolean;
        /** The name of the field that the rule applies to. */
        targetPropertyName: string;
    }

    /** Defines a contractual rule for link attribution. */
    export interface LinkAttribution {
        /** A type hint, which is set to LinkAttribution. */
        _type: "LinkAttribution";
        /** The attribution text. */
        text: License;
        /** The URL to the provider's website. */
        url: string;
        /** A Boolean value that determines whether the contents of the rule must be placed in close proximity to the field that the rule applies to. */
        mustBeCloseToContent: boolean;
        /** The name of the field that the rule applies to. */
        targetPropertyName: string;
    }

    /** Defines a contractual rule for media attribution. */
    export interface MediaAttribution {
        /** A type hint, which is set to MediaAttribution. */
        _type: "MediaAttribution";
        /** The URL to the provider's website. */
        url: string;
        /** A Boolean value that determines whether the contents of the rule must be placed in close proximity to the field that the rule applies to. */
        mustBeCloseToContent: boolean;
        /** The name of the field that the rule applies to. */
        targetPropertyName: string;
    }

    /** Defines a contractual rule for plain text attribution. */
    export interface TextAttribution {
        /** A type hint, which is set to TextAttribution. */
        _type: "TextAttribution";
        /** The attribution text. Text attribution applies to the entity as a whole and should be displayed immediately following the entity presentation. If there are multiple text or link attribution rules that do not specify a target, you should concatenate them and display them using a "Data from: " label.  */
        text: string;
    }

    /** Defines an entity such as a person, place, or thing. */
    export interface Entity {
        bingId: string;
        contractualRules: (LicenseAttribution|LinkAttribution|MediaAttribution|TextAttribution)[];
        description: string;
        entityPresentationInfo: EntityPresentationInfo;
        image: Image;
        name: string;
        webSearchUrl: string;
    }

    /** Defines the query context that Bing used for the request. */
    export interface QueryContext {
        adultIntent?: boolean;
        alteredQuery?: string;
        alterationOverrideQuery?: string;
        askUserForLocation?: boolean;
        originalQuery: string;
    }

    /** Defines an entity answer. */
    export interface EntityAnswer {
        /** The supported query scenario. This field is set to DominantEntity or DisambiguationItem. The field is set to DominantEntity if Bing determines that only a single entity satisfies the request. For example, a book, movie, person, or attraction. If multiple entities could satisfy the request, the field is set to DisambiguationItem. For example, if the request uses the generic title of a movie franchise, the entity's type would likely be DisambiguationItem. But, if the request specifies a specific title from the franchise, the entity's type would likely be DominantEntity. */
        queryScenario: string;
        /** A list of entities */
        entities: Entity[];
    }

    /** Defines a local entity answer. */
    export interface LocalEntityAnswer {
        /** Type hint */
        _type: string;
        /** A list of local entities, such as local restaurants or hotels. */
        value: Place[];
    }

    /** Defines information about a local entity, such as a restaurant or hotel. */
    export interface Place {
        /** Type hint */
        _type: "Hotel"|"LocalBusiness"|"Restaurant";
        /** The postal address of where the entity is located. */
        address: PostalAddress;
        /** Additional information about the entity such as hints that you can use to determine the entity's type. */
        entityPresentationInfo: EntityPresentationInfo;
        /** The entity's name. */
        name: string;
        /** The entity's telephone number. */
        telephone: string;
        /** The URL to the entity's website. */
        url: string;
        /** The URL to Bing's search result for this place. */
        webSearchUrl: string;
    }

    /** Defines a postal address. */
    export interface PostalAddress {
        addressCountry: string;
        addressLocality: string;
        addressRegion: string;
        neighborhood: string;
        postalCode: string;
        text: string;
    }

    export interface EntitySearchResult {
        /** A list of entities that are relevant to the search query. */
        entities: EntityAnswer[];
        /** A list of local entities such as restaurants or hotels that are relevant to the search query. */
        places: LocalEntityAnswer[];
        /** An object that contains the query string that Bing used for the request. */
        queryContext: QueryContext;
        clientId: string;
    }

    export interface SearchLocation {
        /** The latitude of the client's location, in degrees. The latitude must be greater than or equal to -90.0 and less than or equal to +90.0. Negative values indicate southern latitudes and positive values indicate northern latitudes. */
        latitude: number;
        /** The longitude of the client's location, in degrees. The longitude must be greater than or equal to -180.0 and less than or equal to +180.0. Negative values indicate western longitudes and positive values indicate eastern longitudes. */
        longitude: number;
        /** The radius, in meters, which specifies the horizontal accuracy of the coordinates. Pass the value returned by the device's location service. Typical values might be 22m for GPS/Wi-Fi, 380m for cell tower triangulation, and 18,000m for reverse IP lookup. */
        horizontalAccuracy: number;
        /** The UTC time when the client was at the location. */
        timestamp?: Date;
    }

    export interface EntitySearchOptions {
        location?: SearchLocation;
        mkt?: string;
        safeSearch?: SafeSearch;
    }
}

// Service endpoints
const entitySearchEndpoint = "https://api.cognitive.microsoft.com/bing/v7.0/entities";

export class BingEntitySearchApi {

    constructor(
        private accessKey: string,
    )
    {
    }

    public async searchEntitiesAsync(query: string, clientId: string, searchOptions?: entities.EntitySearchOptions): Promise<entities.EntitySearchResult> {
        return new Promise<entities.EntitySearchResult>((resolve, reject) => {
            let qsp: any = { q: query };
            if (searchOptions) {
                qsp = { ...searchOptions, ...qsp };
            }

            let options = {
                url: `${entitySearchEndpoint}?${querystring.stringify(qsp)}`,
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
                        entities: body.entities,
                        places: body.places,
                        queryContext: body.queryContext,
                        clientId: res.headers["X-MSEdge-ClientID"],
                    });
                }
            });
        });
    }
}
