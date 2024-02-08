import express, {Express, Request, Response} from "express";
import dotenv from "dotenv";
import path from "node:path";
import expressLayouts from 'express-ejs-layouts';
import * as querystring from "querystring";
import nocache from "nocache";
import * as fs from "fs";
import {jwtDecode} from "jwt-decode";

const tokenEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
const authEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"

const appConsole = new console.Console(
  {
    stdout: fs.createWriteStream("./logger.txt", {flags: "a+"}),
    stderr: fs.createWriteStream("./error.txt", {flags: "a+"}),
    colorMode: false
  }
);

dotenv.config();

const port = process.env.PORT || 3000;

const app: Express = express();


app.use(nocache());
app.use(expressLayouts);
app.set("views", path.join(__dirname, 'views'));
app.set("view engine", 'ejs');

app.get("/", (req, res) => res.redirect("/step1"));

app.get("/step1", (req: Request, res: Response) => {
  const authParameters = {
    response_type: "code",
    client_id: process.env.OAUTH_CLIENT_ID,
    scope: [
      "https://outlook.office.com/IMAP.AccessAsUser.All",
      "https://outlook.office.com/SMTP.Send",
      "offline_access",
      "openid",
    ].join(" "),
    redirect_uri: createOAuthRedirectUrl(req)
    //state: 'your-auth-session-id-here-if-needed',
    //login_hint: 'john@example.local' //to pre-populate the email field in the auth form
  };

  const office365AuthorizationUrl = authEndpoint + "?" + querystring.stringify(authParameters);

  res.render('pages/step1', {authUrl: office365AuthorizationUrl});
});

app.get("/step2", async (req: Request, res: Response) => {
  const authorizationCode = req.query['code'];
  appConsole.debug("Authorization code received: ", authorizationCode)

  let error, refreshToken: string | undefined;
  error = [req.query['error'], req.query['error_description']].filter(s => !!s).join("\n\n")

  if(authorizationCode) {
    appConsole.debug("Redeeming authorization code: ", authorizationCode)

    try {
      refreshToken = await retrieveRefreshToken(authorizationCode.toString(), req);

      appConsole.debug("Refresh token received: ", refreshToken)

      return res.redirect("/step3?" + querystring.stringify({refresh_token: refreshToken}));
    } catch (e) {
      appConsole.error(`Failed to redeem authorization code "${authorizationCode}": ${e}`)
      error = e?.toString()
    }
  }

  res.render('pages/step2', {authorizationCode, error});
});

app.get("/step3", async (req: Request, res: Response) => {
  const refreshToken = req.query['refresh_token'];
  appConsole.debug("Refresh token received: ", refreshToken)

  res.render('pages/step3', {
    refreshToken
  });
});

app.get("/fetch/inbox.json", async (req: Request, res: Response) => {
  const refreshToken = req.query['refresh_token']?.toString();

  const fail = (error: string) => res.json({error: "Failed to fetch mailbox: " + error});

  try {
    appConsole.debug("Fetching inbox using refresh token: ", refreshToken)

    if(!refreshToken) {
      return fail("No refresh token");
    }

    const [accessToken, emailAddress] = await retrieveAccessToken(refreshToken);

    appConsole.log("Using email: ", emailAddress)
    appConsole.log("Using access token: ", accessToken)

    res.json([1, 2, 3]);
  } catch (e) {
    return fail("Failed to fetch mailbox: " + e);
  }
});

app.listen(port, () => {
  appConsole.log(`[server]: Server is running at http://localhost:${port}`);
});

function createOAuthRedirectUrl(req: Request): string {
  return `https://${req.get('host')}/step2`;
}

async function callTokenEndpoint(tokenParameters: any): Promise<any> {
  const response = await fetch(tokenEndpoint, {
    method: 'POST',
    body: querystring.stringify(tokenParameters),
    headers: {'Content-type': 'application/x-www-form-urlencoded'},
  })

  appConsole.debug("HTTP status: ", response.status, response.statusText)

  const json = await response.text();

  let parsedResponse: any;
  try {
    parsedResponse = json ? JSON.parse(json) : undefined;
  } catch (error) {
    appConsole.debug("Raw response JSON: ", json)
    throw new Error("Failed to parse response JSON: " + json)
  }

  if(!parsedResponse) {
    throw new Error(`Empty response received, HTTP status "${response.status}"`)
  }

  appConsole.debug("Parsed response: ", JSON.stringify(parsedResponse, null, "\t"))

  const tokenError = parsedResponse["error"];
  const tokenErrorDescr = parsedResponse["error_description"];

  if(tokenError || tokenErrorDescr) {
    throw new Error("Failed to call token endpoint: " + [tokenError, tokenErrorDescr].filter(s => !!s).join("\n\n"))
  }

  return parsedResponse;
}

async function retrieveRefreshToken(authorizationCode: string, req: Request): Promise<string | undefined> {
  const tokenParameters = {
    grant_type: "authorization_code",
    code: authorizationCode,
    client_id: process.env.OAUTH_CLIENT_ID,
    client_secret: process.env.OAUTH_CLIENT_SECRET,
    redirect_uri: createOAuthRedirectUrl(req)
  };

  let parsedResponse = await callTokenEndpoint(tokenParameters);

  const refreshToken = parsedResponse['refresh_token'];
  if(refreshToken) {
    return refreshToken;
  }

  appConsole.error("Refresh token not found in the response", parsedResponse)
  throw new Error("No refresh token received from server");
}

async function retrieveAccessToken(refreshToken: string): Promise<[string, string]> {
  const tokenParameters = {
    client_id: process.env.OAUTH_CLIENT_ID,
    client_secret: process.env.OAUTH_CLIENT_SECRET,
    refresh_token: refreshToken,
    grant_type: "refresh_token",
  };

  const parsedResponse = await callTokenEndpoint(tokenParameters);

  const idToken = parsedResponse["id_token"];
  appConsole.debug("Raw id_token: ", idToken);

  const decodedIdToken: any = jwtDecode(idToken);
  appConsole.debug("Decoded id_token: ", decodedIdToken);

  return [parsedResponse["access_token"], decodedIdToken['email']];
}
