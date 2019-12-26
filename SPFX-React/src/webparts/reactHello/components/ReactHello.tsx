import * as React from "react";
import styles from "./ReactHello.module.scss";
import { IReactHelloProps } from "./IReactHelloProps";
import { escape } from "@microsoft/sp-lodash-subset";

interface IReactReadWebPartState {
  Property: string;
  Value1: number;
  Value2: number;
  Value3: number;
}

export default class ReactHello extends React.Component<
  IReactHelloProps,
  IReactReadWebPartState,
  {}
> {
  constructor(props) {
    super(props);
    this.state = { Property: "", Value1: 0, Value2: 0, Value3: 0 };
    this.ButtonClick = this.ButtonClick.bind(this);
  }
  ButtonClick = () => {
    let DateTime = new Date();
    let value = DateTime.toISOString() + " at Button clicked.";
    this.setState({ Property: value });
    console.log(value);
  };

  Value1Change = event => {
    console.log(event.target.value);
    this.setState({ Value1: event.target.value });
  };

  Value2Change = event => {
    let aValue = event.target.value;
    console.log(aValue);
    this.setState({ Value2: aValue });
  };

  CountTotal = event => {
    console.log("Total");
    let Total =
      parseInt(this.state.Value1.toString()) +
      parseInt(this.state.Value2.toString());
    this.setState({ Value3: Total });
    event.preventDefault();
  };

  public render(): React.ReactElement<IReactHelloProps> {
    return (
      <div className={styles.reactHello}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>
                Customize SharePoint experiences using Web Parts.
              </p>
              <p className={styles.description}>
                {escape(this.props.description)}
              </p>

              <p className="${ styles.description}">
                Continents where I reside: {escape(this.props.myContinent)}
              </p>
              <p className="${ styles.description}">
                Number of contents I'have visted:
                {this.props.numContinentsVisited}
              </p>

              <p className={styles.description}>
                {escape(this.props.areYouGoodWithReact)}
              </p>
              <input
                type="button"
                onClick={this.ButtonClick}
                value="Click Me for Logging"
              />
              <br />

              <input
                type="text"
                value={this.state.Property}
                id="StateProperty"
              />
              <form onSubmit={this.CountTotal}>
                <p>Value 1:</p>
                <input
                  type="text"
                  id="txtValue1"
                  value={this.state.Value1}
                  onChange={this.Value1Change}
                />
                <p>Value 2:</p>
                <input
                  type="text"
                  id="txtValue2"
                  value={this.state.Value2}
                  onChange={this.Value2Change}
                />
                <p>Total:</p>
                <input type="text" value={this.state.Value3} id="txtValue3" />
                <input type="Submit" value="Submit" />
              </form>
              <a href="#" className={styles.button}>
                <span className={styles.label}>Click Me</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
