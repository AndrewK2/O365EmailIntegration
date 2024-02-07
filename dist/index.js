"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const dotenv_1 = __importDefault(require("dotenv"));
const node_path_1 = __importDefault(require("node:path"));
const express_ejs_layouts_1 = __importDefault(require("express-ejs-layouts"));
const querystring = __importStar(require("querystring"));
const nocache_1 = __importDefault(require("nocache"));
const tokenEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
const authEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
dotenv_1.default.config();
const app = (0, express_1.default)();
const port = process.env.PORT || 3000;
app.use((0, nocache_1.default)());
app.use(express_ejs_layouts_1.default);
app.set("views", node_path_1.default.join(__dirname, 'views'));
app.set("view engine", 'ejs');
app.get("/", (req, res) => res.redirect("/step1"));
app.get("/step1", (req, res) => {
    const authParameters = {
        response_type: "code",
        client_id: process.env.OAUTH_CLIENT_ID,
        scope: [
            "https://outlook.office.com/IMAP.AccessAsUser.All",
            "https://outlook.office.com/SMTP.Send",
            "offline_access"
        ].join(" "),
        redirect_uri: createOAuthRedirectUrl(req)
        //state: 'your-auth-session-id-here-if-needed',
        //login_hint: 'john@example.local' //to pre-populate the email field in the auth form
    };
    const office365AuthorizationUrl = authEndpoint + "?" + querystring.stringify(authParameters);
    res.render('pages/step1', {
        authUrl: office365AuthorizationUrl,
    });
});
app.get("/step2", (req, res) => __awaiter(void 0, void 0, void 0, function* () {
    const authorizationCode = req.query['code'];
    console.debug("Authorization code received: ", authorizationCode);
    let error, refreshToken;
    if (authorizationCode) {
        console.debug("Redeeming authorization code: ", authorizationCode);
        try {
            refreshToken = yield retrieveRefreshToken(authorizationCode.toString(), req);
            console.debug("Refresh token received: ", refreshToken);
            return res.redirect("/step3?" + querystring.stringify({ refresh_token: refreshToken }));
        }
        catch (e) {
            console.error(`Failed to redeem authorization code "${authorizationCode}": ${e}`);
            error = e === null || e === void 0 ? void 0 : e.toString();
        }
    }
    res.render('pages/step2', {
        authorizationCode,
        refreshToken,
        error
    });
}));
app.listen(port, () => {
    console.log(`[server]: Server is running at http://localhost:${port}`);
});
function createOAuthRedirectUrl(req) {
    return `https://${req.get('host')}/step2`;
}
function retrieveRefreshToken(authorizationCode, req) {
    return __awaiter(this, void 0, void 0, function* () {
        const tokenParameters = {
            grant_type: "authorization_code",
            code: authorizationCode,
            client_id: process.env.OAUTH_CLIENT_ID,
            client_secret: process.env.OAUTH_CLIENT_SECRET,
            redirect_uri: createOAuthRedirectUrl(req)
        };
        const response = yield fetch(tokenEndpoint, {
            method: 'POST',
            body: querystring.stringify(tokenParameters),
            headers: { 'Content-type': 'application/x-www-form-urlencoded' },
        });
        console.debug("HTTP status: ", response.status);
        const json = yield response.text();
        console.debug("Response JSON: ", json);
        let parsedResponse;
        if (json) {
            parsedResponse = JSON.parse(json);
        }
        else {
            parsedResponse = undefined;
        }
        if (!parsedResponse) {
            throw new Error(`Empty response received, HTTP status "${response.status}"`);
        }
        const refreshToken = parsedResponse['refresh_token'];
        if (refreshToken) {
            return refreshToken;
        }
        console.error("Refresh token not found in the response", parsedResponse);
        const tokenError = parsedResponse["error"];
        const tokenErrorDescr = parsedResponse["error_description"];
        if (tokenError || tokenErrorDescr) {
            throw new Error("Failed to redeem auth code: " + [tokenError, tokenErrorDescr].join("\n\n"));
        }
        throw new Error("No refresh token received from server");
    });
}
