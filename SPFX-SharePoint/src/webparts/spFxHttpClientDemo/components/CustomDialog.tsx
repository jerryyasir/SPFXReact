import {
  BaseDialog,
  IDialogConfiguration
} from "../../../../node_modules/@microsoft/sp-dialog/lib";
import * as React from "react";
import * as ReactDom from "react-dom";

export default class CustomDialog extends BaseDialog {
  public paramFromDailog: string;

  public render(): void {
    var html: string = "";

    html += `<div class="modal-content">`;
    html += `<div class="modal-header">
        <h5 class="modal-title">Confirm Deletion</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>`;
    html += `<div class="modal-body">
    <p>Are you sure you want to delete this item?</p>
  </div>`;
    html += `<div class="modal-footer"><button id="yesButton" type="button" class="btn btn-danger">Yes</button>
    <button type="button" class="btn btn-primary" id="noButton">No</button>`;
    html += `</div>`;
    this.domElement.innerHTML += html;
    this._setButtonEventHandlers();
  }

  // METHOD TO BIND EVENT HANDLER TO BUTTON CLICK
  private _setButtonEventHandlers(): void {
    const webPart: CustomDialog = this;
    this.domElement
      .querySelector("#yesButton")
      .addEventListener("click", () => {
        console.log("Yes Buttin is Clicked...");
        this.paramFromDailog = "yes";
        this.close();
      });
    this.domElement.querySelector("#noButton").addEventListener("click", () => {
      this.paramFromDailog = "no";
      this.close();
    });
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }

  protected onAfterClose(): void {
    super.onAfterClose();
    // Clean up the element for the next dialog
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
