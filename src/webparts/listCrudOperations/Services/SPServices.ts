import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { IDropdownOption } from "office-ui-fabric-react";
import * as strings from "ListCrudOperationsWebPartStrings";

export class SPOperations {
  /**
   * GetAllListFromWeb
   */
  public GetAllListFromWeb(
    context: WebPartContext
  ): Promise<IDropdownOption[]> {
    let restUrl: string =
      context.pageContext.web.absoluteUrl + "/_api/web/lists?select=Title";
    console.log(restUrl);
    var listTitles: IDropdownOption[] = [];
    return new Promise<IDropdownOption[]>(async (resolve, reject) => {
      context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1).then(
        (response: SPHttpClientResponse) => {
          response.json().then((results: any) => {
            console.log(results);
            results.value.map((result: any) => {
              listTitles.push({
                key: result.Title,
                text: result.Title,
              });
            });
          });
          resolve(listTitles);
        },
        (error: any): void => {
          reject("Error Occurred" + error);
        }
      );
    });
  }

  /**
        * CreateListItem
    context:WebPartContext,listTitle:string     */
  public CreateListItem(
    context: WebPartContext,
    listTitle: string
  ): Promise<string> {
    let restUrl: string =
      context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getByTitle('" +
      listTitle +
      "')/items";
    const body: string = JSON.stringify({ Title: "New List Title" });
    const options: ISPHttpClientOptions = {
      body: body,
      headers: {
        Accept: "Application/json;odata=nometadata",
        "content-type": "Application/json;odata=nometadata",
        "odata-version": "",
      },
    };
    return new Promise<string>(async (resolved, rejected) => {
      context.spHttpClient
        .post(restUrl, SPHttpClient.configurations.v1, options)
        .then((response: SPHttpClientResponse) => {
          response.json().then(
            (result: any) => {
              resolved(
                "Item with ID: " +
                  result.ID +
                  " got Created Succcessfully in List: " +
                  listTitle +
                  " !!!"
              );
            },
            (error: any) => {
              rejected("Error Occurred!!!");
            }
          );
        });
    });
  }

  /**
   * GetLatestItemIDFromList
context:WebPartContext,listTitle:string   */
  public GetLatestItemIDFromList(
    context: WebPartContext,
    listTitle: string
  ): Promise<number> {
    let restUrl =
      context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getByTitle('" +
      listTitle +
      "')/items/?$orderby=Id desc&$top=1&select=Id";
    return new Promise<number>(async (resolved, rejected) => {
      context.spHttpClient
        .get(restUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          response.json().then((result: any) => {
            console.log(result);
            resolved(result.value[0].Id),
              (error: any) => {
                rejected("Error Occurred while fetching the ID.." + error);
              };
          });
        });
    });
  }

  /**
   * DeleteItemByLatestItemIDinList
context:WebPartContext,listTitle:string   */
  public DeleteItemByLatestItemIDinList(
    context: WebPartContext,
    listTitle: string
  ): Promise<string> {
    let restUrl =
      context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getByTitle('" +
      listTitle +
      "')/items";

    return new Promise<string>(async (resolved, rejected) => {
      this.GetLatestItemIDFromList(context, listTitle).then(
        (itemId: Number) => {
          context.spHttpClient
            .post(
              restUrl + "('" + itemId + "')",
              SPHttpClient.configurations.v1,
              {
                headers: {
                  Accept: "Application/json;odata=nometadata",
                  "content-type": "Application/json;odata=nometadata",
                  "odata-version": "",
                  "IF-MATCH": "*",
                  "X-HTTP-METHOD": "DELETE",
                },
              }
            )
            .then(() => {
              resolved(
                "Item with ID: " + itemId + " is deleted successfully.."
              );
              (error: any) => {
                rejected(
                  "Error Occurred on deleting item with ID: " +
                    itemId +
                    "Error Message: " +
                    error
                );
              };
            });
        }
      );
    });
  }

  /**
   * UpdateItemByLatestItemInList
context:WebPartContext,listTitle:string   */
  public UpdateItemByLatestItemInList(
    context: WebPartContext,
    listTitle: string
  ): Promise<string> {
    let restUrl =
      context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getByTitle('" +
      listTitle +
      "')/items";
    const body: string = JSON.stringify({ Title: "Updated Item" });
    return new Promise<string>(async (resolved, rejected) => {
      this.GetLatestItemIDFromList(context, listTitle).then(
        (itemId: number) => {
          console.log(
            "Updated Item Rest Url: " + restUrl + "('" + itemId + "')"
          );
          context.spHttpClient
            .post(
              restUrl + "('" + itemId + "')",
              SPHttpClient.configurations.v1,
              {
                headers: {
                  Accept: "Application/json;odata=nometadata",
                  "content-type": "Application/json;odatada=nometadata",
                  "odata-version": "",
                  "IF-MATCH": "*",
                  "X-HTTP-METHOD": "MERGE",
                },
                body: body,
              }
            )
            .then((response: SPHttpClientResponse) => {
              resolved(
                "Item with ID: " + itemId + "got Updated Successfully!!"
              );
              (error: any) => {
                rejected(
                  "Error occurred in Updating item with ID: " +
                    itemId +
                    "Error Message: " +
                    error
                );
              };
            });
        }
      );
    });
  }
}
