/*eslint-disable*/
import { MSGraphClientFactory, MSGraphClient} from "@microsoft/sp-http";
import { BatchRequestContent, BatchResponseContent} from '@microsoft/microsoft-graph-client';  
import {
    SPHttpClient,
    SPHttpClientResponse,
    ISPHttpClientOptions
} from "@microsoft/sp-http";
import { MSGraphClientV3 } from '@microsoft/sp-http';


import { WebPartContext } from '@microsoft/sp-webpart-base';


export class GraphService{
    public context: WebPartContext;
    private _msGraphClientFactory: MSGraphClientFactory;
    private _siteUrl: string;
    private _spHttpClient: SPHttpClient;
    private _userEmail: string;
    public _graphClient: MSGraphClientV3;
    private _spHttpOptions: any = {
        getMetaData: <ISPHttpClientOptions>{
            headers: {
                'ACCEPT': 'application/json; odata.metadata=full'
            }
        },
        getNoMetaData: <ISPHttpClientOptions>{
            headers: {
                'ACCEPT': 'application/json; odata.metadata=none'
            }
        },
        updateNoMetaData: <ISPHttpClientOptions>{
            headers: {
                'ACCEPT': 'application/json; odata.metadata=none',
                'CONTENT-TYPE': 'application/json',
                'X-HTTP-METHOD': 'MERGE'
            }
        },
        postVerboseMetaData: <ISPHttpClientOptions>{
            headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-type': 'application/json;odata=verbose'
            }
        },
        postNoMetaData: <ISPHttpClientOptions>{
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=nometadata'
            }
        },
        postMinimalMetaData: <ISPHttpClientOptions>{
            headers: {
                'Accept': 'application/json;odata.metadata=none;charset=utf-8',
                'Content-type': 'application/json;odata.metadata=none;charset=utf-8'
            }
        }
    };
   
  /**
   *
   */
  constructor(context: WebPartContext) {
    this.context = context;
    this._siteUrl = this.context.pageContext.web.absoluteUrl;
    this._spHttpClient = this.context.spHttpClient;
    this._userEmail = this.context.pageContext.user.email;
    this._msGraphClientFactory = this.context.msGraphClientFactory;
    
  }
public async GetUserByEmailWithManagers(userEmail:string){
    this._graphClient = await this._msGraphClientFactory.getClient("3");
    return await this._graphClient.api(`https://graph.microsoft.com/beta/users/${userEmail}`)
    // .expand("directReports")
    .expand("manager")
    .header("ConsistencyLevel", "eventual")
    .get()    
}
// public async GetUserByEmailWithDirectReports(userEmail:string){
//     this._graphClient = await this._msGraphClientFactory.getClient("3");
//     return await this._graphClient.api(`/users/${userEmail}`)
//     // .expand("directReports")
//     .expand("directReports")
//     .header("ConsistencyLevel", "eventual")
//     .get()    
// }
public async GetUserExtendedManager(userEmail:string){
    this._graphClient = await this._msGraphClientFactory.getClient("3");
    try{
        return await this._graphClient.api(`https://graph.microsoft.com/beta/users/${userEmail}?&$expand=manager($levels=max;$select=id,displayName,mail)`)
       .header("ConsistencyLevel", "eventual")
       .get()  

    }catch {
    }
}

public async GetUserDirectReports(userEmail:string){
    this._graphClient = await this._msGraphClientFactory.getClient("3");
     return await this._graphClient.api(`https://graph.microsoft.com/beta/users/${userEmail}/directReports`)
    .header("ConsistencyLevel", "eventual")
    .get()  
}
// public async GetDirectReportsInBatch(userNames:Array<string>) {
//     this._graphClient = await this._msGraphClientFactory.getClient("3");
//     let batchRequestContent = new BatchRequestContent();
  
//     // Loop through the array of user IDs and add a GET request for each user
//     userNames.forEach((userId, i) => {
//       let index = i +1;
//       batchRequestContent.addRequest({
//         id: index.toString(),
//         request: new Request(`/users/${userId}`, {
//             method: "GET"
//           })
//       });
//     });

  
//     let content = await batchRequestContent.getContent();
    
//     // POST the batch request content to the /$batch endpoint
//     let batchResponse = await this._graphClient
//       .api('/$batch')
      
//       .post(content);
  
//     let batchResponseContent = new BatchResponseContent(batchResponse);
//     let i;    
//     let responsePromises :Array<Promise<any>> = [];
//     for(i=1; i<= userNames.length; i++){
//        let temp = batchResponseContent.getResponseById(i.toString());
    
//        if(temp.ok) {
//         let data = temp.json().then(val => val);
//         responsePromises.push(data);
//        }
//     }
  
//     let responses = await Promise.all(responsePromises);
//     return responses;
//   }
  public async GetUsersInBatch(userNames:Array<string>) {
    this._graphClient = await this._msGraphClientFactory.getClient("3");
    let batchRequestContent = new BatchRequestContent();
    // Loop through the array of user IDs and add a GET request for each user
    userNames.forEach((userId, i) => {
      let index = i +1;
      batchRequestContent.addRequest({
        id: index.toString(),
        request: new Request(`/users/${userId}?&expand=manager`, {
            method: "GET"
          })
      });
    });

  
    let content = await batchRequestContent.getContent();
    
    // POST the batch request content to the /$batch endpoint
    let batchResponse = await this._graphClient
      .api('/$batch')
      .post(content);
  
    let batchResponseContent = new BatchResponseContent(batchResponse);
    let i;    
    let responsePromises :Array<Promise<any>> = [];
    for(i=1; i<= userNames.length; i++){
       let temp = batchResponseContent.getResponseById(i.toString());
       if(temp.ok) {
        let data = temp.json().then(val => val);
        responsePromises.push(data);
       }
    }
    
    let responses = await Promise.all(responsePromises);
   
    return responses;
  }

}
