import { ICountryListItem, ButtonClickCallback } from "../../../models/index";

export interface ISpFxHttpClientDemoProps {
  spListItems: ICountryListItem[];
  onGetListItems?: ButtonClickCallback;
  onAddListItem?: ButtonClickCallback;
  onUpdateListItem?: ButtonClickCallback;
  onDeleteListItem?: ButtonClickCallback;
  title: string;
  cost: string;
  category: string;
}
