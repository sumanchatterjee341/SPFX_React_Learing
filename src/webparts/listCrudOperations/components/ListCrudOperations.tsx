import * as React from "react";
import styles from "./ListCrudOperations.module.scss";
import { IListCrudOperationsProps } from "./IListCrudOperationsProps";
import { escape, fromPairs } from "@microsoft/sp-lodash-subset";
import { IListCrudOperationsState } from "./IListCrudOperationsState";
import { SPOperations } from "../Services/SPServices";
import {
  DefaultButton,
  Dropdown,
  IDropdownOption,
} from "office-ui-fabric-react";

export default class ListCrudOperations extends React.Component<
  IListCrudOperationsProps,
  IListCrudOperationsState,
  {}
> {
  public _SPOps: SPOperations;
  public selectedListTitle: string;
  //Constructors
  constructor(props: IListCrudOperationsProps) {
    super(props);
    this._SPOps = new SPOperations();
    this.state = { ListTitles: [], status: "" };
  }

  public GetListTitle = (event: any, data: any) => {
    this.selectedListTitle = data.text;
  }

  //Sets the data in the dropdown
  public componentDidMount() {
    this._SPOps
      .GetAllListFromWeb(this.props.Context)
      .then((results: IDropdownOption[]) => {
        this.setState({ ListTitles: results });
      });
  }

  // Responsible to Render components in the page
  public render(): React.ReactElement<IListCrudOperationsProps> {    
    return (
      <div className={styles.listCrudOperations}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
            </div>
            <div id="dv_parent" className={styles.CustomStyles}>
              <Dropdown
                className={styles.CustomDropdown}
                options={this.state.ListTitles}
                onChange={this.GetListTitle}
                placeholder="----Select Your List----"
              ></Dropdown>

              <DefaultButton
                className={styles.myButton}
                text="Create List Item"
                onClick={() =>
                  this._SPOps
                    .CreateListItem(this.props.Context, this.selectedListTitle)
                    .then((result: string) => {
                      console.log("Status is: " + result);
                      this.setState({ status: result });
                    })
                }
              ></DefaultButton>
              <DefaultButton
                className={styles.myButton}
                text="Update List Item"
                onClick={()=>{
                  this._SPOps.UpdateItemByLatestItemInList(this.props.Context,this.selectedListTitle).then((result:any)=>{
                    this.setState({status:result});
                  });
                }}
              ></DefaultButton>
              <DefaultButton
                className={styles.myButton}
                text="Delete List Item"
                onClick={() =>
                  this._SPOps
                    .DeleteItemByLatestItemIDinList(
                      this.props.Context,
                      this.selectedListTitle
                    )
                    .then((result: any) => {
                      this.setState({ status: result });
                    })
                }
              ></DefaultButton>
              <div className={styles.myStatusBar}>{this.state.status}</div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
