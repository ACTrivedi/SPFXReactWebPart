import * as React from "react";
import styles from "./ReactWebPart.module.scss";
import type { IReactWebPartProps } from "./IReactWebPartProps";
import type { IReactWebPartState } from "./IReactWebPartState";
import { SP_OPerations } from "../../Services/SpServices";
import { IDropdownOption } from "office-ui-fabric-react";
import { Dropdown } from "office-ui-fabric-react/lib/Dropdown";
import { PrimaryButton } from "office-ui-fabric-react";

// import { escape } from '@microsoft/sp-lodash-subset';

export default class ReactWebPart extends React.Component<
  IReactWebPartProps,
  IReactWebPartState,
  {}
> {
  public _spOPs: SP_OPerations;
  public selectedListTitle: string;
  constructor(props: IReactWebPartProps) {
    super(props);
    this._spOPs = new SP_OPerations();
    this.state = {
      listTitles: [],
      status: "",
    };
  }

  public getListTitle = (event: any, data: any) => {
    this.selectedListTitle = data.text;
  };

  public componentDidMount(): void {
    this._spOPs
      .GetAllList(this.props.context)
      .then((result: IDropdownOption[]) => {
        this.setState({ listTitles: result });
      });
  }

  public render(): React.ReactElement<IReactWebPartProps> {
    return (
      <div className={styles.SpFxCrud}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <div className={styles.titleDiv}>
              <span className={styles.title}>Welcome to SPFXCrud!</span>
              </div>
              <p className={styles.subTitle}>
                DEMO: SharePoint CRUD Operations Using Rest Api (spHTTPClient)
              </p>
            </div>
            <div className={styles.column}>
              <Dropdown
                options={this.state.listTitles}
                placeholder="Select Your List"
                onChange={this.getListTitle}
              ></Dropdown>
              <PrimaryButton
                className={styles.button}
                text="Create Item"
                onClick={() =>
                  this._spOPs
                    .createListItem(this.props.context, this.selectedListTitle)
                    .then((result: string) => {
                      this.setState({ status: result });
                    })
                }
              ></PrimaryButton>

              <PrimaryButton
                className={styles.button}
                text="Update Item"  
                onClick={() =>
                  this._spOPs
                    .UpdateListItem(this.props.context, this.selectedListTitle)
                    .then((result: string) => {
                      this.setState({ status: result });
                    })}                
              ></PrimaryButton>

              <PrimaryButton
                className={styles.button}
                text="Delete Item"  
                onClick={() =>
                  this._spOPs
                    .DeleteListItem(this.props.context, this.selectedListTitle)
                    .then((result: string) => {
                      this.setState({ status: result });
                    })}             
              ></PrimaryButton>
              <div className={styles.column}>{this.state.status}</div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
