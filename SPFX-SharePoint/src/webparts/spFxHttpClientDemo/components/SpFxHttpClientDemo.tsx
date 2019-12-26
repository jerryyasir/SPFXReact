import * as React from "react";
import styles from "./SpFxHttpClientDemo.module.scss";
import { ISpFxHttpClientDemoProps } from "./ISpFxHttpClientDemoProps";
import { ISpFxHttpClientDemoState } from "./ISpFxHttpClientDemoState";
import { escape } from "@microsoft/sp-lodash-subset";

export default class SpFxHttpClientDemo extends React.Component<
  ISpFxHttpClientDemoProps,
  ISpFxHttpClientDemoState,
  {}
> {
  constructor(
    props: ISpFxHttpClientDemoProps,
    state: ISpFxHttpClientDemoState
  ) {
    super(props);
    this.state = { Title: "", Cost: "", Category: "" };
    this.onTitleChange = this.onTitleChange.bind(this);
    this.onCostChange = this.onCostChange.bind(this);
    this.onCategoryChange = this.onCategoryChange.bind(this);
  }
  private onGetListItemsClicked = (
    event: React.MouseEvent<HTMLAnchorElement>
  ): void => {
    event.preventDefault();
    this.props.onGetListItems();
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

    this.props.onAddListItem();
  };
  componentDidMount() {
    //setInterval(() => this.props.onGetListItems(), 1000000);
    this.props.onGetListItems();
  }
  public render(): React.ReactElement<ISpFxHttpClientDemoProps> {
    return (
      <div className={styles.panelStyle}>
        <div className={styles.tableCaptionStyle}>
          Fetch Proejct Details from SharePointList using SPFx,RESTAPI,React JS
          Data on page changes with change in the SharePointList{" "}
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

        <div className={styles.headerCaptionStyle}>Course Details</div>
        <div className={styles.tableStyle}>
          <div className={styles.headerStyle}>
            <div className={styles.IdCellStyle}>Id</div>
            <div className={styles.titleCellStyle}>Title </div>
            <div className={styles.titleCellStyle}>Performance Category</div>
            <div className={styles.costCellStyle}>Cost</div>
          </div>

          {this.props.spListItems.map(function(item, key) {
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
              </div>
            );
          })}
        </div>
      </div>
    );
  }
}
