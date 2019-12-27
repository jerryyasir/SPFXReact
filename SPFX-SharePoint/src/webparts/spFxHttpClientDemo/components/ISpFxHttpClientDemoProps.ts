import { ICountryListItem, ButtonClickCallback } from "../../../models/index";
import { SPHttpClient } from "@microsoft/sp-http";

export interface ISpFxHttpClientDemoProps {
  spHttpClient: SPHttpClient;
  listName: string;
  siteUrl: string;
}
