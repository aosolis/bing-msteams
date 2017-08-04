# Translator
This contains the source for the LinkedIn compose extension for Microsoft Teams.

### Compile
To compile the Typescript files, run `gulp build`.
To package up the app for Azure deployment, run `gulp package`.

#### Fiddler
To use Fiddler, add the following lines to the `env` section:
```
    "http_proxy": "http://localhost:8888",
    "no_proxy": "login.microsoftonline.com",
    "NODE_TLS_REJECT_UNAUTHORIZED": "0"
```
Change `http_proxy` to the Fiddler endpoint.

### Bot state
Per-user and per-conversation state goes to the Bot Framework state store (https://docs.botframework.com/en-us/core-concepts/userdata/).

### Configuration 
 - Default configuration is in `config\default.json`
 - Production overrides are in `config\production.json`
 - Environment variables overrides are in `config\custom-environment-variables.json`.
 - Specify local overrides in `config\local.json`. **Do not check in this file.**
