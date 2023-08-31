
/* eslint-disable*/
import { IUser } from "./IUser";
export interface IUserInfo extends  IUser {
  title:string;
  pictureUrl?: string;
  // userUrl?:string;
  id?:string;
  directReports?:Array<any>
  hasDirectReports?:boolean;
  numberDirectReports?:number;
  // hasPeers?:boolean;
  numberPeers?:number;
  manager?:string;
  department?:string;
  workPhone?:string;
  cellPhone?:string;
  location?:string;
  office?: string;
  userPrincipalName?: string;
  onPremisesExtensionAttributes?: {
    extensionAttribute1:string;
    extensionAttribute2: string
    extensionAttribute3: string
extensionAttribute4: string
extensionAttribute5: string
extensionAttribute6:string
extensionAttribute7: string
extensionAttribute8: string
extensionAttribute9: string
extensionAttribute10: string
extensionAttribute11: string
extensionAttribute12: string
extensionAttribute13: string
extensionAttribute14: string
extensionAttribute15: string
  }
 }
