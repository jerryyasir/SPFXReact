import { ICountryListItem } from "../../../models";

export interface ISpFxHttpClientDemoState {
  PerformanceListItems: ICountryListItem[];
  Title: string;
  Cost: string;
  SelectedId: string;
  SelectedItem: any;
  Performance_x0020_Category: string;
  AddButtonHidden: boolean;
  UpdateButtonHidden: boolean;
  Status: string;
  showDiv: boolean;
}
