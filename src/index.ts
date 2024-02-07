import express, {Express, Request, Response} from "express";
import dotenv from "dotenv";
import path from "node:path";
import expressLayouts from 'express-ejs-layouts';
import * as querystring from "querystring";
import nocache from "nocache";

const tokenEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
const authEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"

dotenv.config();

const app: Express = express();
const port = process.env.PORT || 3000;

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

app.get("/step2", async (req: Request, res: Response) => {
  const authorizationCode = req.query['code'];
  console.debug("Authorization code received: ", authorizationCode)

  let error, refreshToken: string | undefined;

  if(authorizationCode) {
    console.debug("Redeeming authorization code: ", authorizationCode)

    try {
      refreshToken = await retrieveRefreshToken(authorizationCode.toString(), req);

      console.debug("Refresh token received: ", refreshToken)

      return res.redirect("/step3?" + querystring.stringify({refresh_token: refreshToken}));
    } catch (e) {
      console.error(`Failed to redeem authorization code "${authorizationCode}": ${e}`)
      error = e?.toString()
    }
  }

  res.render('pages/step2', {
    authorizationCode,
    refreshToken,
    error
  });
});

app.listen(port, () => {
  console.log(`[server]: Server is running at http://localhost:${port}`);
});


function createOAuthRedirectUrl(req: Request): string {
  return `https://${req.get('host')}/step2`;
}

async function retrieveRefreshToken(authorizationCode: string, req: Request): Promise<string | undefined> {
  const tokenParameters = {
    grant_type: "authorization_code",
    code: authorizationCode,
    client_id: process.env.OAUTH_CLIENT_ID,
    client_secret: process.env.OAUTH_CLIENT_SECRET,
    redirect_uri: createOAuthRedirectUrl(req)
  };

  const response = await fetch(tokenEndpoint, {
    method: 'POST',
    body: querystring.stringify(tokenParameters),
    headers: {'Content-type': 'application/x-www-form-urlencoded'},
  })

  console.debug("HTTP status: ", response.status)

  const json = await response.text();
  console.debug("Response JSON: ", json)

  let parsedResponse: any | undefined;
  if(json) {
    parsedResponse = JSON.parse(json);
  } else {
    parsedResponse = undefined;
  }

  if(!parsedResponse) {
    throw new Error(`Empty response received, HTTP status "${response.status}"`)
  }

  const refreshToken = parsedResponse['refresh_token'];
  if(refreshToken) {
    return refreshToken;
  }

  console.error("Refresh token not found in the response", parsedResponse)

  const tokenError = parsedResponse["error"];
  const tokenErrorDescr = parsedResponse["error_description"];

  if(tokenError || tokenErrorDescr) {
    throw new Error("Failed to redeem auth code: " + [tokenError, tokenErrorDescr].join("\n\n"))
  }

  throw new Error("No refresh token received from server");
}
