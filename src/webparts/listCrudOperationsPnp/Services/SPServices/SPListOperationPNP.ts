import { WebPartContext } from "@microsoft/sp-webpart-base";
//import { sp } from "@pnp/sp/presets/all";
import { IDropdownOption, TooltipHost } from "office-ui-fabric-react";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";
export class SPListOperations {
  /**
   * Gets all the list titles from the current site and returns an array of DropdownOptions
   */
  public GetListTitles(context: WebPartContext): Promise<IDropdownOption[]> {
    sp.setup(context);
    let listTitles: IDropdownOption[] = [];
    return new Promise<IDropdownOption[]>(async (resolved, rejected) => {
      sp.web.lists
        .select("Title")()
        .then(
          (results: any) => {
            console.log(results);
            results.map((result: any) => {
              listTitles.push({ key: result.Title, text: result.Title });
            });
            resolved(listTitles);
          },
          (error: any) => {
            rejected("error occurred" + error);
          }
        );
    });
  }

  public async CreateListItem(
    Context: WebPartContext,
    ListTitle: string
  ):Promise<string>{
    sp.setup({
      spfxContext: Context,
      sp:{
        headers:{
          "Accept":"application/json;odata=verbose"
        }
      }
    });

    return new Promise<string>(async(resolved,rejected)=>{
      sp.web.lists.getByTitle(ListTitle).items.add({
        Title:"New PNPItem"
      }).then((result:any)=>{
        console.log(result);
        resolved("Item Created");
      },(err:any)=>{
        rejected("Error Occurred: "+err);
      });
    });
  }
}
