import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import Axios from "axios";
import { IUser } from "../../../model/IUser";

/**
 * Implementation of the SPO Rest Api  content page
 */
export const SpoRestApiTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [user, setUser] = useState<IUser>();
    const [error, setError] = useState<string>();

    useEffect(() => {
        if (context) {
            microsoftTeams.authentication.getAuthToken({
                successCallback: (token: string) => {
                    ensureUser(token, context?.teamSiteDomain!, context?.teamSiteUrl!);
                    microsoftTeams.appInitialization.notifySuccess();
                },
                failureCallback: (message: string) => {
                    setError(message);
                    microsoftTeams.appInitialization.notifyFailure({
                        reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                        message
                    });
                },
                resources: [process.env.SPORESTAPI_APP_URI as string]
            });
        }
    }, [context]);

    const ensureUser = (token: string, domain: string, siteUrl: string) => {
        if (token) {
            const requestBody = {
                domain,
                siteUrl
            };
            Axios.post(`https://${process.env.HOSTNAME}/api/ensureuser`, requestBody, {
                          responseType: "json",
                          headers: {
                            Authorization: `Bearer ${token}`
                          }
              }).then(result => {
                const user: IUser = result.data;
                
                setUser(user);
              })
              .catch((error) => {
                console.log(error);
              });
        }
    };
    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <Flex.Item>
                    <Header content="Demo calling SPO Rest Api from Teams" />
                </Flex.Item>
                <Flex.Item>
                    <div>

                        <div>
                            <Text content={`Hello ${user?.displayName}`} />                            
                        </div>
                        <div>
                            <Text content={`Your LookupID in this site is: ${user?.lookupID}`} />
                        </div>
                        {error && <div><Text content={`An SSO error occurred ${error}`} /></div>}

                        <div>
                            <Button onClick={() => alert("It worked!")}>A sample button</Button>
                        </div>
                    </div>
                </Flex.Item>
                <Flex.Item styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Text size="smaller" content="(C) Copyright Markus Moeller" />
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
