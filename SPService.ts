import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import {
  MSGraphClientFactory,
  SPHttpClient,
  SPHttpClientResponse,
} from "@microsoft/sp-http";

export class SPService {
  public static readonly serviceKey: ServiceKey<SPService> = ServiceKey.create(
    "AaaS:SPService",
    SPService
  );
  private _spHttpClient: SPHttpClient;
  private _webUrl: string;
  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      const pageContext = serviceScope.consume(PageContext.serviceKey);
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
      this._webUrl = pageContext.web.absoluteUrl;
    });
  }

  public async getEvents(
    webUrl: string = this._webUrl,
    listId: string,
    viewXml: string,
    datesInUTC: boolean = true,
    renderOptions: number = 2
  ): Promise<any[]> {
    const endpoint = `${webUrl}/_api/web/lists('${listId}')/RenderListDataAsStream`;
    const body = {
      parameters: {
        RenderOptions: renderOptions,
        ViewXml: viewXml,
        DatesInUtc: datesInUTC,
      },
    };
    const response: SPHttpClientResponse = await this._spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=verbose",
        },
        body: JSON.stringify(body),
      }
    );
    if (!response.ok) {
      throw new Error(`Failed to get events: ${response.statusText}`);
    }
    const data = await response.json();
    return data.Row || [];
  }
}
