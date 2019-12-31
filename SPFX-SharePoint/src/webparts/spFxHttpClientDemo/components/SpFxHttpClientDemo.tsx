import * as React from "react";
import styles from "./SpFxHttpClientDemo.module.scss";
import { ISpFxHttpClientDemoProps } from "./ISpFxHttpClientDemoProps";
import { ISpFxHttpClientDemoState } from "./ISpFxHttpClientDemoState";
import { escape } from "@microsoft/sp-lodash-subset";
import * as bootstrap from "bootstrap";

import CustomDialog from "../components/CustomDialog";

require("../../../../node_modules/bootstrap/dist/css/bootstrap.css");
require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ICountryListItem } from "../../../models";

import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { Dialog } from "@microsoft/sp-dialog";

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
      Performance_x0020_Category: "",
      SelectedId: "",
      SelectedItem: null,
      AddButtonHidden: false,
      UpdateButtonHidden: true,
      PerformanceListItems: [],
      Status: "",
      showDiv: false
    };
    this.onTitleChange = this.onTitleChange.bind(this);
    this.onCostChange = this.onCostChange.bind(this);
    this.onCategoryChange = this.onCategoryChange.bind(this);
    this.onItemSelected = this.onItemSelected.bind(this);
    this.onItemSelectedForDeletion = this.onItemSelectedForDeletion.bind(this);
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
  private _getListItem(id): Promise<ICountryListItem[]> {
    return this.props.spHttpClient
      .get(
        `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$select=Id,Title,Cost,Performance_x0020_Category&$Filter=Id eq '${id}'`,
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
              Performance_x0020_Category: "Cost"
            },
            {
              Id: "2",
              Title: "Test Title 2",
              Cost: "300",
              Performance_x0020_Category: "Quality"
            },
            {
              Id: "3",
              Title: "Test Title 3",
              Cost: "400",
              Performance_x0020_Category: "Reliability"
            },
            {
              Id: "4",
              Title: "Test Title 4 ",
              Cost: "500",
              Performance_x0020_Category: "Responsiveness"
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
    //this.setState({ AddButtonHidden: false, UpdateButtonHidden: false });
    this.render();
  };

  private _updateListItem = (id): void => {
    if (!this._isSharePoint) {
      if (this.state.PerformanceListItems.length > 0) {
        let item = this.state.PerformanceListItems.filter(
          item => item.Id === id
        );
        if (item.length > 0) {
          item[0].Cost = this.state.Cost;
          item[0].Performance_x0020_Category = this.state.Performance_x0020_Category;
          item[0].Title = this.state.Title;
          this.setState({
            Cost: "",
            Title: "",
            AddButtonHidden: true,
            showDiv: false,
            UpdateButtonHidden: false
          });
        }
      }
    } else {
      //this._getListItems().then(response => {
      this._updateListItemToSP();
      //});
    }
    this.render();
  };

  private onItemSelected = event => {
    var selectedId = event.target.getAttribute("data-item");
    var title = event.target.getAttribute("title");

    this.setState({ SelectedId: selectedId, Title: title, showDiv: true });
    this.getListItemOnPage(selectedId);
    console.log(selectedId);
  };

  private onItemSelectedForDeletion = event => {
    var selectedId = event.target.getAttribute("data-item");
    var title = event.target.getAttribute("title");
    this.setState({ SelectedId: selectedId, Title: title });

    const dialog: CustomDialog = new CustomDialog();
    this._deleteItem(this.state.SelectedId);
    // dialog.show().then(() => {
    //   console.log(dialog.paramFromDailog);

    //   if (
    //     dialog.paramFromDailog != undefined &&
    //     dialog.paramFromDailog === "yes"
    //   ) {
    //     Dialog.alert(`Record has been Deleted!`);
    //     this._deleteItem(this.state.SelectedId);
    //     console.log(this.state.SelectedId);
    //   }
    // });
  };

  private arrayRemove(arr, value) {
    return arr.filter(function(ele) {
      return ele.Id != value;
    });
  }
  private _deleteItem = (id): void => {
    if (!this._isSharePoint) {
      if (this.state.PerformanceListItems.length > 0) {
        let items = this.state.PerformanceListItems;
        let result = this.arrayRemove(items, id);
        console.log(this.state.PerformanceListItems.length);
        console.log(result.length);
        this.setState({
          PerformanceListItems: result
        });
      }
    } else {
      this._deleteListItemFromSP();
    }
    console.log("Record has been Deleted!");
    this.setState({ AddButtonHidden: true, UpdateButtonHidden: false });
  };

  private getListItemOnPage = (id): void => {
    if (!this._isSharePoint) {
      if (this.state.PerformanceListItems.length > 0) {
        let item = this.state.PerformanceListItems.filter(
          item => item.Id === id
        );
        if (item.length > 0) {
          this.setState({
            Cost: item[0].Cost,
            Performance_x0020_Category: item[0].Performance_x0020_Category
          });
        }
      }
    } else {
      this._getListItem(id).then(response => {
        let Id = response[0].Id;
        let Title = response[0].Title;
        let Cost = response[0].Cost;
        let Performance_x0020_Category = response[0].Performance_x0020_Category;
        this.setState({
          SelectedId: Id,
          Performance_x0020_Category: Performance_x0020_Category,
          Cost: Cost,
          Title: Title
        });
      });
    }
    this.setState({ AddButtonHidden: true, UpdateButtonHidden: false });
    this.render();
  };

  private _onAddListItem = (): void => {
    if (!this._isSharePoint) {
      var newItem = {
        Id: (Math.floor(Math.random() * 10000) + 1).toString(),
        Title: this.state.Title,
        Cost: this.state.Cost,
        Performance_x0020_Category: this.state.Performance_x0020_Category
      };
      let listItems = this.state.PerformanceListItems;
      listItems.push(newItem);
      this.setState({
        PerformanceListItems: listItems,
        UpdateButtonHidden: true,
        AddButtonHidden: false
      });
      this.render();
    } else {
      this._addListItem().then(() => {
        this._getListItems().then(response => {
          this.setState({
            PerformanceListItems: response,
            showDiv: false,
            Title: "",
            Cost: "",
            Performance_x0020_Category: ""
          });
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
        Performance_x0020_Category: this.state.Performance_x0020_Category,
        "@odata.type": spEntityType
      });

      return this.props.spHttpClient.post(
        `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`,
        SPHttpClient.configurations.v1,
        request
      );
    });
  }

  private _deleteListItemFromSP(): void {
    let etag: string = undefined;
    let listItemEntityTypeName: string = undefined;
    this._getItemEntityType()
      .then(spEntityType => {
        listItemEntityTypeName = spEntityType;
        return this.props.spHttpClient.get(
          `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${this.state.SelectedId})?$select=Id`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "odata-version": ""
            }
          }
        );
      })
      .then(
        (response: SPHttpClientResponse): Promise<ICountryListItem> => {
          etag = response.headers.get("ETag");
          return response.json();
        }
      )
      .then(
        (item: ICountryListItem): Promise<SPHttpClientResponse> => {
          return this.props.spHttpClient.post(
            `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${item.Id})`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=nometadata",
                "Content-type": "application/json;odata=verbose",
                "odata-version": "",
                "IF-MATCH": etag,
                "X-HTTP-Method": "DELETE"
              }
            }
          );
        }
      )
      .then(
        (response: SPHttpClientResponse): void => {
          this.setState({
            Status: `Item with ID: ${this.state.SelectedId} successfully updated`,
            UpdateButtonHidden: true,
            AddButtonHidden: false,
            Title: "",
            Cost: "",
            Performance_x0020_Category: "",
            SelectedId: ""
          });
          this._onGetListItems();
        },
        (error: any): void => {
          this.setState({
            Status: `Error updating item: ${error}`
          });
        }
      );
  }

  private _updateListItemToSP(): void {
    let etag: string = undefined;
    let listItemEntityTypeName: string = undefined;
    this._getItemEntityType()
      .then(spEntityType => {
        listItemEntityTypeName = spEntityType;
        return this.props.spHttpClient.get(
          `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${this.state.SelectedId})?$select=Id`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "odata-version": ""
            }
          }
        );
      })
      .then(
        (response: SPHttpClientResponse): Promise<ICountryListItem> => {
          etag = response.headers.get("ETag");
          return response.json();
        }
      )
      .then(
        (item: ICountryListItem): Promise<SPHttpClientResponse> => {
          const body = JSON.stringify({
            __metadata: {
              type: listItemEntityTypeName
            },
            Title: this.state.Title,
            Cost: this.state.Cost,
            Performance_x0020_Category: this.state.Performance_x0020_Category,
            "@odata.type": listItemEntityTypeName
          });

          return this.props.spHttpClient.post(
            `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${item.Id})`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=nometadata",
                "Content-type": "application/json;odata=verbose",
                "odata-version": "",
                "IF-MATCH": etag,
                "X-HTTP-Method": "MERGE"
              },
              body: body
            }
          );
        }
      )
      .then(
        (response: SPHttpClientResponse): void => {
          console.log("Item updated...");
          this.setState({
            Status: `Item with ID: ${this.state.SelectedId} successfully updated`,
            UpdateButtonHidden: true,
            AddButtonHidden: false,
            Title: "",
            showDiv: false,
            Cost: "",
            Performance_x0020_Category: "",
            SelectedId: ""
          });
          this._onGetListItems();
        },
        (error: any): void => {
          this.setState({
            Status: `Error updating item: ${error}`
          });
        }
      );
  }

  private onTitleChange = event => {
    this.setState({ Title: event.target.value });
  };
  private onCostChange = event => {
    this.setState({ Cost: event.target.value });
  };
  private onCategoryChange = event => {
    this.setState({ Performance_x0020_Category: event.target.value });
  };

  private onAddListItemClicked = event => {
    event.preventDefault();

    let title = this.state.Title;
    let cost = this.state.Cost;
    let category = this.state.Performance_x0020_Category;
    console.log(
      "Button was clicked...:" + title + " : " + cost + " : " + category
    );

    this._onAddListItem();
  };

  OnUpdateClicked = event => {
    event.preventDefault();

    let title = this.state.Title;
    let cost = this.state.Cost;
    let category = this.state.Performance_x0020_Category;

    this._updateListItem(this.state.SelectedId);
  };

  componentDidMount() {
    //setInterval(() => this.props.onGetListItems(), 1000000);
    this._onGetListItems();
  }

  public render(): React.ReactElement<ISpFxHttpClientDemoProps> {
    let tableRows = this.state.PerformanceListItems.map(
      function(item, key) {
        return (
          <div className={styles.rowStyle} key={key}>
            <div className={styles.IdCellStyle}>{item.Id}</div>
            <div className={styles.titleCellStyle}>{escape(item.Title)}</div>
            <div className={styles.titleCellStyle}>
              {escape(item.Performance_x0020_Category)}
            </div>
            <div className={styles.costCellStyle}>{escape(item.Cost)}</div>
            <div className={styles.costCellStyle} key={item.Id}>
              <button
                key={item.Id}
                data-item={item.Id}
                title={item.Title}
                className="btn btn-secondary"
                onClick={this.onItemSelected}
              >
                Update
              </button>
              &nbsp;
              <button
                data-item={item.Id}
                title={item.Title}
                className="btn btn-danger"
                onClick={this.onItemSelectedForDeletion}
              >
                Delete
              </button>
            </div>
          </div>
        );
      }.bind(this)
    );

    return (
      <div className={styles.panelStyle}>
        <div className={styles.tableCaptionStyle}>
          Server Side Web Part to SPFX Web Part Migration{" "}
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

          {tableRows}
        </div>
        <br />
        <button
          hidden={this.state.AddButtonHidden}
          id="addNewrecord"
          className="btn btn-primary"
          onClick={() => this.setState({ showDiv: !this.state.showDiv })}
        >
          Add New Item
        </button>
        {this.state.showDiv && (
          <div id="addItemDev">
            <form>
              <div className="form-group">
                <label className="text-left medium">Title</label>
                <input
                  className="form-control"
                  type="Text"
                  id="title"
                  name="title"
                  placeholder="enter title here"
                  value={this.state.Title}
                  onChange={this.onTitleChange}
                />
              </div>
              <div className="form-group">
                <label>Cost</label>
                <input
                  className="form-control"
                  type="Text"
                  id="cost"
                  name="cost"
                  placeholder="enter cost here"
                  value={this.state.Cost}
                  onChange={this.onCostChange}
                />
              </div>
              <div className="form-group">
                <label>Performance Categogry</label>
                <select
                  className="form-control"
                  id="categorydrop"
                  name="categorydrop"
                  value={this.state.Performance_x0020_Category}
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
              </div>
              <button
                hidden={this.state.AddButtonHidden}
                id="addrecord"
                className="btn btn-primary"
                onClick={this.onAddListItemClicked}
              >
                Save
              </button>
              &nbsp;
              <a
                href="#"
                hidden={this.state.UpdateButtonHidden}
                id="updaterecord"
                className="btn btn-primary"
                onClick={this.OnUpdateClicked}
              >
                <span className={styles.label}>Update Record</span>
              </a>
              <button
                id="cancell"
                className="btn btn-secondary"
                onClick={() => this.setState({ showDiv: false })}
              >
                Cancel
              </button>
            </form>
          </div>
        )}
      </div>
    );
  }
}
