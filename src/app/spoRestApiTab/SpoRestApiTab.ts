import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/spoRestApiTab/index.html")
@PreventIframe("/spoRestApiTab/config.html")
@PreventIframe("/spoRestApiTab/remove.html")
export class SpoRestApiTab {
}
