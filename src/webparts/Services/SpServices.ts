import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";

export class SP_OPerations {

  public GetAllList(context: WebPartContext): Promise<IDropdownOption[]> {
    let restApiUrl: string =
      // context.pageContext.web.absoluteUrl + "/_api/web/Lists?$select=Title";
      context.pageContext.web.absoluteUrl + "/sites/SPFxLearningSite/_api/web/Lists?$select=Title";

    var listTitles: IDropdownOption[] = [];

    return new Promise<IDropdownOption[]>(async (resolve, reject) => {
      context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1).then(
        (response: SPHttpClientResponse) => {
          response.json().then((results: any) => {
            console.log(results);
            results.value.map((result: any) => {
              listTitles.push({ key: result.Title, text: result.Title });
            });
          });
          resolve(listTitles);
        },
        (error: any): void => {
          reject("Error Occured" + error);
        }
      );
    });
  }

  public createListItem(
    context: WebPartContext,
    listTitle: string
  ): Promise<string> {
    let restApiUrl: string =
      context.pageContext.web.absoluteUrl +
      "/sites/SPFxLearningSite/_api/web/Lists/getByTitle('" +
      listTitle +
      "')/items";
    const body: string = JSON.stringify({ Title: "New Item Created" });
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=nometadata",
        "content-type": "application/json;odata=nometadata",
        "odata-version":"",
      },
      body: body,
    };
    return new Promise<string>(async (resolve, reject) => {
      context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options).then((response:SPHttpClientResponse)=>{
        resolve("Item created Successfully");
      }),(error:any)=>{
        reject("Error occured");
      }
    });
  }

  public DeleteListItem(context:WebPartContext,listTitle: string):
    Promise<string> {
    let restApiUrl: string =
      context.pageContext.web.absoluteUrl +
      "/sites/SPFxLearningSite/_api/web/Lists/getByTitle('" +
      listTitle +
      "')/items";
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=nometadata",
        "content-type": "application/json;odata=nometadata",
        "odata-version":"",
        "IF-MATCH":"*",
        "X-HTTP-METHOD":"DELETE",
      },
    };
    return new Promise<string>(async (resolve, reject) => {
      this.getLatestItemId(context,listTitle).then((itemId:number)=>{
        context.spHttpClient.post(restApiUrl+"("+itemId+")",SPHttpClient.configurations.v1,options).then((response:SPHttpClientResponse)=>{
          resolve("Item with Id"+itemId+"deleted successfully");
        }),(error:any)=>{
          reject("Error occured");
        }
      })
    });

  }

  public getLatestItemId(context: WebPartContext, listTitle: string): Promise<number> {
    let restApiUrl: string =
      context.pageContext.web.absoluteUrl +
      "/sites/SPFxLearningSite/_api/web/Lists/getByTitle('" +
      listTitle +
      "')/items?$orderby=Id desc&$top=1&$select=id";
  
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=nometadata",
        "odata-version": "",
      },
    };
  
    return new Promise<number>(async (resolve, reject) => {
      try {
        const response: SPHttpClientResponse = await context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1, options);
        
        if (response.ok) {
          const result: any = await response.json();
          if (result.value && result.value.length > 0) {
            resolve(result.value[0].Id);
          } else {
            reject("No data found in the response.");
          }
        } else {
          reject(`Error: ${response.status} - ${response.statusText}`);
        }
      } catch (error) {
        reject(`An error occurred: ${error}`);
      }
    });
  }

  public UpdateListItem(context:WebPartContext,listTitle: string):
    Promise<string> {
    let restApiUrl: string =
      context.pageContext.web.absoluteUrl +
      "/sites/SPFxLearningSite/_api/web/Lists/getByTitle('" +
      listTitle +
      "')/items";
      const body: string = JSON.stringify({ Title: "Updated Item" });

    const options: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=nometadata",
        "content-type": "application/json;odata=nometadata",
        "odata-version":"",
        "IF-MATCH":"*",
        "X-HTTP-METHOD":"MERGE",
      },
      body:body
    };
    return new Promise<string>(async (resolve, reject) => {
      this.getLatestItemId(context,listTitle).then((itemId:number)=>{
        context.spHttpClient.post(restApiUrl+"("+itemId+")",SPHttpClient.configurations.v1,options).then((response:SPHttpClientResponse)=>{
          resolve("Item with Id"+itemId+"updated successfully");
        }),(error:any)=>{
          reject("Error occured");
        }
      })
    });

  }
  
}
