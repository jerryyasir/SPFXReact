import * as React from "react";
import styles from "./SpFxHttpClientDemo.module.scss";
import { ISpFxHttpClientDemoProps } from "./ISpFxHttpClientDemoProps";
import { ISpFxHttpClientDemoState } from "./ISpFxHttpClientDemoState";
import { escape } from "@microsoft/sp-lodash-subset";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ICountryListItem } from "../../../models";

import { Environment, EnvironmentType } from "@microsoft/sp-core-library";

export default class SpFxHttpClientDemo extends React.Component<
  ISpFxHttpClientDemoProps,
  ISpFxHttpClientDemoState,
  {}
> {
  //private PerformanceListItems: ICountryListItem[] = [];
  constructor(
    props: ISpFxHttpClientDemoProps,
    state: ISpFxHttpClientDemoState
  ) {
    super(props);
    this.state = {
      Title: "",
      Cost: "",
      Category: "",
      SelectedId: "",
      PerformanceListItems: []
    };
    this.onTitleChange = this.onTitleChange.bind(this);
    this.onCostChange = this.onCostChange.bind(this);
    this.onCategoryChange = this.onCategoryChange.bind(this);
    this.onUpdateListItem = this.onUpdateListItem.bind(this);
  }

  private get _isSharePoint(): boolean {
    return (
      Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint
    );
  }

  private _getListItems(): Promise<ICountryListItem[]> {
    return this.props.spHttpClient
      .get(
        `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$select=Id,Title,Cost,Performance_x0020_Category&$Filter=Cost le '10'`,
        SPHttpClient.configurations.v1
      )
      .then(response => {
        return response.json();
      })
      .then(jsonResponse => {
        return jsonResponse.value;
      }) as Promise<ICountryListItem[]>;
  }

  private _onGetListItems = (): void => {
    if (!this._isSharePoint) {
      if (this.state.PerformanceListItems.length === 0) {
        this.setState({
          PerformanceListItems: [
            {
              Id: "1",
              Title: "Test Title 1",
              Cost: "200",
              Performance_x0020_Category: "Category 1"
            },
            {
              Id: "2",
              Title: "Test Title 2",
              Cost: "300",
              Performance_x0020_Category: "Category 2"
            },
            {
              Id: "3",
              Title: "Test Title 3",
              Cost: "400",
              Performance_x0020_Category: "Category 3"
            },
            {
              Id: "4",
              Title: "Test Title 4 ",
              Cost: "500",
              Performance_x0020_Category: "Category 4"
            }
          ]
        });
      }
    } else {
      this._getListItems().then(response => {
        //        this._countries = response;
        this.setState({ PerformanceListItems: response });
      });
    }
    this.render();
  };

  private _onAddListItem = (): void => {
    if (!this._isSharePoint) {
      var newItem = {
        Id: (Math.floor(Math.random() * 10000) + 1).toString(),
        Title: this.state.Title,
        Cost: this.state.Cost,
        Performance_x0020_Category: this.state.Category
      };
      let listItems = this.state.PerformanceListItems;
      listItems.push(newItem);
      this.setState({ PerformanceListItems: listItems });
      this.render();
    } else {
      this._addListItem().then(() => {
        this._getListItems().then(response => {
          this.setState({ PerformanceListItems: response });
          this.render();
        });
      });
    }
  };

  private _getItemEntityType(): Promise<string> {
    return this.props.spHttpClient
      .get(
        `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1
      )
      .then(response => {
        return response.json();
      })
      .then(jsonResponse => {
        return jsonResponse.ListItemEntityTypeFullName;
      }) as Promise<string>;
  }

  private _addListItem(): Promise<SPHttpClientResponse> {
    return this._getItemEntityType().then(spEntityType => {
      const request: any = {};
      request.body = JSON.stringify({
        Title: this.state.Title,
        Cost: this.state.Cost,
        Performance_x0020_Category: this.state.Category,
        "@odata.type": spEntityType
      });

      return this.props.spHttpClient.post(
        `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`,
        SPHttpClient.configurations.v1,
        request
      );
    });
  }
  private onGetListItemsClicked = (
    event: React.MouseEvent<HTMLAnchorElement>
  ): void => {
    event.preventDefault();
    this._onGetListItems();
  };

  private onTitleChange = event => {
    this.setState({ Title: event.target.value });
  };
  private onCostChange = event => {
    this.setState({ Cost: event.target.value });
  };
  private onCategoryChange = event => {
    this.setState({ Category: event.target.value });
  };

  private onUpdateListItem = (event, id) => {
    let Id = event.target.getAttribute("data-item");
    this.setState({ SelectedId: Id });
    console.log(Id);
  };

  private onAddListItemClicked = (
    event: React.MouseEvent<HTMLAnchorElement>
  ): void => {
    event.preventDefault();

    let title = this.state.Title;
    let cost = this.state.Cost;
    let category = this.state.Category;
    let listItem = { title: title, cost: cost, Category: category };
    console.log(
      "Button was clicked...:" + title + " : " + cost + " : " + category
    );

    this._onAddListItem();
  };

  componentDidMount() {
    //setInterval(() => this.props.onGetListItems(), 1000000);
    this._onGetListItems();
  }

  public render(): React.ReactElement<ISpFxHttpClientDemoProps> {
    //console.log(this.state.PerformanceListItems);

    return (
      <div className={styles.panelStyle}>
        <div className={styles.tableCaptionStyle}>
          Fetch Proejct Details from SharePointList using SPFx,RESTAPI,React JS
          Data on page changes with change in the SharePointList{" "}
        </div>

        <div className={styles.headerCaptionStyle}>Course Details</div>
        <div className={styles.tableStyle}>
          <div className={styles.headerStyle}>
            <div className={styles.IdCellStyle}>Id</div>
            <div className={styles.titleCellStyle}>Title </div>
            <div className={styles.titleCellStyle}>Performance Category</div>
            <div className={styles.costCellStyle}>Cost</div>
            <div className={styles.costCellStyle}>Action</div>
          </div>
          {this.state.PerformanceListItems.map(function(item, key) {
            return (
              <div className={styles.rowStyle} key={key}>
                <div className={styles.IdCellStyle}>{item.Id}</div>
                <div className={styles.titleCellStyle}>
                  {escape(item.Title)}
                </div>
                <div className={styles.titleCellStyle}>
                  {escape(item.Performance_x0020_Category)}
                </div>
                <div className={styles.costCellStyle}>{escape(item.Cost)}</div>
                <div className={styles.costCellStyle} key={item.Id}>
                  <button
                    key={item.Id}
                    onClick={() => this.updateListItem(item.Id)}
                    data-item={item.Id}
                  >
                    Update
                  </button>
                  <a href="#">Delete</a>
                </div>
              </div>
            );
          })}
        </div>

        <div>
          <h3>SPFX Add Record</h3>
          <form>
            Title:
            <input
              type="Text"
              id="title"
              placeholder="enter title here"
              value={this.state.Title}
              onChange={this.onTitleChange}
            />
            <br />
            Cost:
            <input
              type="Text"
              id="cost"
              placeholder="enter cost here"
              value={this.state.Cost}
              onChange={this.onCostChange}
            />
            <br />
            Performance Category:
            <select
              value={this.state.Category}
              onChange={this.onCategoryChange}
            >
              <option key="Empty" value="Empty">
                Please Select
              </option>
              <option key="Cost" value="Cost">
                Cost
              </option>
              <option key="Quality">Quality</option>
              <option key="Reliability">Reliability</option>
              <option key="Responsiveness">Responsiveness</option>
            </select>
            <br />
            <a
              href="#"
              className={styles.button}
              onClick={this.onAddListItemClicked}
            >
              <span className={styles.label}>Add List Item</span>
            </a>
          </form>
        </div>
      </div>
    );
  }
}
