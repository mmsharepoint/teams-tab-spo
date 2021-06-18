import Axios from "axios";
import debug = require("debug");
import express = require("express");
import passport = require("passport");
import { BearerStrategy, IBearerStrategyOption, ITokenPayload, VerifyCallback } from "passport-azure-ad";
import qs = require("qs");
import { IUser } from "../../model/IUser";
const log = debug('graphRouter');

export const spoRouter = (options: any): express.Router => {
  const router = express.Router();

  // Set up the Bearer Strategy
  const bearerStrategy = new BearerStrategy({
      identityMetadata: "https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration",
      clientID: process.env.SPORESTAPI_APP_ID as string,
      audience: `api://${process.env.HOSTNAME}/${process.env.SPORESTAPI_APP_ID}`,
      loggingLevel: "warn",
      validateIssuer: false,
      passReqToCallback: false
  } as IBearerStrategyOption,
      (token: ITokenPayload, done: VerifyCallback) => {
          done(null, { tid: token.tid, name: token.name, upn: token.upn }, token);
      }
  );
  const pass = new passport.Passport();
  router.use(pass.initialize());
  pass.use(bearerStrategy);

  /**
   * This function creates an access token to be used with SPO Rest Api
   * @param tenantName Name
   * @param scope scope such as https://domain.sharepoint.com/sites.read.all
   * @param refreshToken Refresh Token of the Graph access token
   * @returns accessToken as string
   */
   const getSPOToken = async (tenantName: string, scope: string, refreshToken: string): Promise<string> => {
    return new Promise((resolve, reject) => {
        const url = `https://login.microsoftonline.com/${tenantName}/oauth2/v2.0/token`;
        const params = {
            client_id: process.env.SPORESTAPI_APP_ID,
            client_secret: process.env.SPORESTAPI_APP_SECRET,
            grant_type: "refresh_token",
            refresh_token: refreshToken,
            scope: scope
        };

        Axios.post(url,
            qs.stringify(params), {
            headers: {
                "Accept": "application/json",
                "Content-Type": "application/x-www-form-urlencoded"
            }
        }).then(result => {
            if (result.status !== 200) {
                reject(result);
                log(result.statusText);
            } else {
              resolve(result.data.access_token);
            }
        }).catch(err => {
            log(err.response.data);
            reject(err);
        });
    });
  };
  // Define a method used to exhchange the identity token to an access token
  const exchangeForToken = (tid: string, token: string, scopes: string[]): Promise<{accessToken: string,refreshToken: string}> => {
      return new Promise((resolve, reject) => {
          const url = `https://login.microsoftonline.com/${tid}/oauth2/v2.0/token`;
          const params = {
              client_id: process.env.SPORESTAPI_APP_ID,
              client_secret: process.env.SPORESTAPI_APP_SECRET,
              grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
              assertion: token,
              requested_token_use: "on_behalf_of",
              scope: scopes.join(" ")
          };

          Axios.post(url,
              qs.stringify(params), {
              headers: {
                  "Accept": "application/json",
                  "Content-Type": "application/x-www-form-urlencoded"
              }
          }).then(result => {
              if (result.status !== 200) {
                reject(result);
              } else {
                resolve({accessToken: result.data.access_token, refreshToken: result.data.refresh_token});
              }
          }).catch(err => {
              // error code 400 likely means you have not done an admin consent on the app
              reject(err);
          });
      });
    };

    const ensureSPOUserByLogin = async (spoAccessToken: string, userEmail: string, siteUrl: string): Promise<IUser> => {
      const requestUrl: string = `${siteUrl}/_api/web/ensureuser`;      
      const userLogin = {
        logonName: userEmail
      };
      return Axios.post(requestUrl, userLogin,
        {
          headers: {          
            Authorization: `Bearer ${spoAccessToken}`
          }
      })
      .then(response => {
          const userLookupID = response.data.Id;
          const userTitle = response.data.Title;
          const user: IUser = { login: userEmail, lookupID: userLookupID, displayName: userTitle };
          return user;      
      });
    };

    router.post(
      "/ensureuser",
      pass.authenticate("oauth-bearer", { session: false }),
      async (req: express.Request, res: express.Response, next: express.NextFunction) => {
        const user: any = req.user;
        
        try {
            const tokenResult = await exchangeForToken(user.tid,
                req.header("Authorization")!.replace("Bearer ", "") as string,
                ["https://graph.microsoft.com/sites.readwrite.all","offline_access"]);
            const accessToken = tokenResult.accessToken;
            const refreshToken = tokenResult.refreshToken;
            const teamSiteDomain = req.body.domain;
            const spoAccessToken = await getSPOToken(teamSiteDomain.toLowerCase().replace('sharepoint', 'onmicrosoft'), 
                                                      `https://${teamSiteDomain}/Sites.ReadWrite.All`, 
                                                      refreshToken);
            const teamSiteUrl = req.body.siteUrl;
            const spouser = await ensureSPOUserByLogin(spoAccessToken, user.upn, teamSiteUrl);
            res.send(spouser);
        } catch (err) {
            if (err.status) {
                res.status(err.status).send(err.message);
            } else {
                res.status(500).send(err);
            }
        }
    });
    return router;
};