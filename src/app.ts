let appInsights = require("applicationinsights");
let express = require("express");
import { Request, Response } from "express";
let bodyParser = require("body-parser");
let favicon = require("serve-favicon");
let http = require("http");
let path = require("path");
let logger = require("morgan");
let config = require("config");
import * as msteams from "botbuilder-teams";
import * as winston from "winston";
import { BingSearchApi } from "./BingSearchApi";
import { BingSearchBot } from "./BingSearchBot";
import { MongoDbBotStorage } from "./storage/MongoDbBotStorage";
import * as utils from "./utils";
import * as certs from "./windows-certs";
import * as jwt from "jsonwebtoken";

// Configure instrumentation
let instrumentationKey = config.get("app.instrumentationKey");
if (instrumentationKey) {
    appInsights.setup(instrumentationKey)
        .setAutoDependencyCorrelation(true)
        .start();
    winston.add(utils.ApplicationInsightsTransport as any);
    appInsights.client.addTelemetryProcessor(utils.stripQueryFromTelemetryUrls);
}

// Configure Key Vault
if (config.get("keyVault.enabled")) {
    // Fetch the certificate
    winston.info("Fetching certificate for KeyVault");
    new Promise<certs.X509Certificate>((resolve, reject) => {
        certs.get({ storeLocation: "CurrentUser" }, (err, certs) => {
            if (err) {
                winston.error("FATAL: Failed to find certificate for KeyVault", err);
                reject(err);
            } else {
                let thumbprint = config.get("keyVault.certificateThumbprint");
                let cert = certs.find(c => c.thumbprint === thumbprint);
                winston.info("Found certificate with thumbprint " + cert.thumbprint);
                resolve(cert);
            }
        });
    }).then(cert => {
        // Mint a token so we can connect to Key Vault
        let options = {
            algorithm: "RS256",
            expiresIn: "1h",
            notBefore: 0,
            audience: "https://login.microsoftonline.com/<tenant_id>/oauth2/token",
            issuer: "<client_id>",
            subject: "<client_id>",
            jwtid: "guid",
            header: {
                x5t: cert.thumbprint,
            },
        };
        let token = jwt.sign({}, cert.pem, options);
        console.log(token);
    }).catch((e) => {
        winston.error("FATAL: Failed to set up Azure Key Vault", e);
        process.exit(1);
    });
}

let app = express();

app.set("port", process.env.PORT || 3978);
app.use(logger("dev"));
app.use(express.static(path.join(__dirname, "../../public")));
app.use(favicon(path.join(__dirname, "../../public/assets", "favicon.ico")));
app.use(bodyParser.json());

// Configure storage
let botStorage = null;
if (config.get("storage") === "mongoDb") {
    botStorage = new MongoDbBotStorage(config.get("mongoDb.botStateCollection"), config.get("mongoDb.connectionString"));
}

// Create chat bot
let connector = new msteams.TeamsChatConnector({
    appId: config.get("bot.appId"),
    appPassword: config.get("bot.appPassword"),
});
let botSettings = {
    storage: botStorage,
    bingSearch: new BingSearchApi(config.get("bing.accessKey")),
};
let bot = new BingSearchBot(connector, botSettings);

// Log bot errors
bot.on("error", (error: Error) => {
    winston.error(error.message, error);
});

// Configure bot routes
app.post("/api/messages", connector.listen());

// Configure ping route
app.get("/ping", (req, res) => {
    res.status(200).send("OK");
});

// catch 404 and forward to error handler
app.use((req: Request, res: Response, next: Function) => {
    let err: any = new Error("Not Found");
    err.status = 404;
    next(err);
});

// error handlers

// development error handler
// will print stacktrace
if (app.get("env") === "development") {
    app.use(function(err: any, req: Request, res: Response, next: Function): void {
        winston.error("Failed request", err);
        res.status(err.status || 500);
        res.render("error", {
            message: err.message,
            error: err,
        });
    });
}

// production error handler
// no stacktraces leaked to user
app.use(function(err: any, req: Request, res: Response, next: Function): void {
    winston.error("Failed request", err);
    res.status(err.status || 500);
    res.render("error", {
        message: err.message,
        error: {},
    });
});

http.createServer(app).listen(app.get("port"), function (): void {
    winston.verbose("Express server listening on port " + app.get("port"));
});
