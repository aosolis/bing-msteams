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
console.log("* Fetching local machine certs");
certs.get({ storeLocation: "LocalMachine" }, (err, certs) => {
   if (!err) {
       certs.forEach(cert => console.log(JSON.stringify(cert)));
//        let cert = certs.find(c => c.thumbprint === "C6448EF954225D147DBB091469F8151C5DFB6ECE");
//        console.log(cert);
   }
});
console.log("* Fetching current user certs");
certs.get({ storeLocation: "CurrentUser" }, (err, certs) => {
   if (!err) {
       certs.forEach(cert => console.log(JSON.stringify(cert)));
//        let cert = certs.find(c => c.thumbprint === "C6448EF954225D147DBB091469F8151C5DFB6ECE");
//        console.log(cert);
   }
});

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
