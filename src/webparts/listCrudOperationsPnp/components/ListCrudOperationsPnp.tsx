import * as React from "react";
import styles from "./ListCrudOperationsPnp.module.scss";
import { IListCrudOperationsPnpProps } from "./IListCrudOperationsPnpProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  Dropdown,
  IDropdownOption,
  DefaultButton,
} from "office-ui-fabric-react";
import { IListCrudOperationsPNPState } from "./IListCrudOperationsPnpStates";
import { SPListOperations } from "../Services/SPServices/SPListOperationPNP";

export default class ListCrudOperationsPnp extends React.Component<
  IListCrudOperationsPnpProps,
  IListCrudOperationsPNPState,
  {}
> {
  private _SPOps: SPListOperations;
  public selectedListTitle: string;
  //Constructors
  constructor(props: IListCrudOperationsPnpProps) {
    super(props);
    this._SPOps = new SPListOperations();
    this.state = { ListTitles: [], status: "" };
  }

  public componentDidMount() {
    this._SPOps.GetListTitles(this.props.Context).then((result: any) => {
      console.log(result);
      this.setState({ ListTitles: result });
    });
  }
  public GetListTitle = (event: any, data: any) => {
    this.selectedListTitle = data.text;
  }

  public render(): React.ReactElement<IListCrudOperationsPnpProps> {
    let option: IDropdownOption[] = [];
    return (
      <div className={styles.listCrudOperationsPnp}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.title}>List Operarions using PNP JSOM</p>
            </div>
            <div className={styles.CustomStyles}>
              <Dropdown
                className={styles.CustomDropdown}
                options={this.state.ListTitles}
                onChange={this.GetListTitle}
                placeholder="---- Select Your List ----"
              ></Dropdown>
              <DefaultButton
                className={styles.myButton}
                text="Create List Item"
                onClick={()=>{
                  this._SPOps.CreateListItem(this.props.Context,this.selectedListTitle).then((result:any)=>{
                    this.setState({status:result});
                  });
                }}
              ></DefaultButton>

              <DefaultButton
                className={styles.myButton}
                text="Update List Item"
              ></DefaultButton>

              <DefaultButton
                className={styles.myButton}
                text="Delete List Item"
              ></DefaultButton>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
