/*eslint-disable*/
export interface IPersonProperties {
    AccountName: string;
    directReport: Array<any>;
    displayName: string;
    mail: string;
    manager: manager
    PictureUrl: string;
    jobTitle: string;
    officeLocation?: string;
    UserProfileProperties: { Key: string; Value: string; ValueType: string }[];
    UserUrl: string;
    userPrincipalName: string;
    onPremisesExtensionAttributes?: {
    extensionAttribute1: string;
    extensionAttribute2: string;
    extensionAttribute3: string;
    extensionAttribute4: string;
    extensionAttribute5: string;
    extensionAttribute6: string;
    extensionAttribute7: string;
    extensionAttribute8: string;
    extensionAttribute9: string;
    extensionAttribute10: string;
    extensionAttribute11: string;
    extensionAttribute12: string;
    extensionAttribute13: string;
    extensionAttribute14: string;
    extensionAttribute15: string;
    }

  }
  export interface manager {
   displayName:string,id?:string,mail?:string, manager?:manager, userPrincipalName?: string;  
  } 
  export interface directReport {
    displayName:string,id?:string,mail?:string, manager?:manager, userPrincipalName?: string; 
  }