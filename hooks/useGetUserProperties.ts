
/*eslint-disable*/
import { IUserInfo } from "../models/IUserInfo";
import * as React from "react";
import { IPersonProperties } from "../models/IPersonProperties";
import { OrgChartContext } from "../context/OrgChartContext";
import { GraphService } from "../services/GraphService";

/*************************************************************************************/
// Hook to get user profile information
// *************************************************************************************/
type returnProfileData =  { managersList:IUserInfo[], reportsLists:IUserInfo[], currentUserProfile :IPersonProperties} ;

type getUserProfileFunc = ( 
  currentUser: string,
  startUser?: string,
  showAllManagers?: boolean,
  filterByCompanyName?:string
  ) => Promise<returnProfileData>;



export  const useGetUserProperties = (): { getUserProfile:getUserProfileFunc } => {
  const {context} = React.useContext(OrgChartContext);
  const adress = context.pageContext.web.absoluteUrl.replace(context.pageContext.web.serverRelativeUrl, "")
  const _graphService = new GraphService(context)
  const getUserProfile = React.useCallback(async (currentUser: string,startUser?: string,showAllManagers?: boolean, filterByCompanyName? : string): Promise<returnProfileData> => {
  const companyFilter = filterByCompanyName !== '' ? filterByCompanyName : '';
      if (!currentUser) return;

      const loginName = currentUser;
      const loginNameStartUser: string = startUser && startUser;
      let currentUserProfile:IPersonProperties   = undefined;

      currentUserProfile = await _graphService.GetUserExtendedManager(loginName);
      console.log(currentUserProfile)
      let reports:IPersonProperties = await _graphService.GetUserDirectReports(loginName);
      const filterCompanyName = reports["value"].filter((value: any) => {
        // If the companyFilter is empty, return true to include all items
        if (!companyFilter) {
          return true;
        }
      
        // Check if value's companyName includes companyFilter
        return value.companyName.includes(companyFilter);
      });
      console.log(filterCompanyName);
      if(reports)
      {
        currentUserProfile.directReport = filterCompanyName
      }

      const reportsLists: IUserInfo[] = [];
      const managersList: IUserInfo[] = [];
      
      const wDirectReports: string[] = currentUserProfile && currentUserProfile.directReport ? getEmailsForDirectReports(currentUserProfile.directReport) : [];
      const wExtendedManagers: string[] = currentUserProfile && currentUserProfile.manager ? getEmailsForManagers(currentUserProfile) : [];
      /*eslint-disable*/
  
      if (wDirectReports && wDirectReports.length > 0) {
         for(const dReports of currentUserProfile.directReport){
          let reports = await _graphService.GetUserDirectReports(dReports.id);
          if(reports)
          {
            dReports.directReport = reports.value.map((v) => manpingUserProperties(v, adress, currentUserProfile.mail))
          }
         reportsLists.push( manpingUserProperties(dReports, adress, currentUserProfile.userPrincipalName))
         }
      }
      // Get Managers if exists
      if (loginNameStartUser && wExtendedManagers && wExtendedManagers.length > 0) {
  

    
     for (let manager of wExtendedManagers)
     {
       if (!showAllManagers && manager !== loginNameStartUser) {
         continue;
        }

       let nManager = await _graphService.GetUserExtendedManager(manager)
        if(!nManager?.manager)
        {
          continue;
        }
       let mDIrectReports =  await _graphService.GetUserDirectReports(manager)
       nManager.directReport = mDIrectReports["value"]; 
       managersList.push(manpingUserProperties(nManager, adress, ""))
     }
     managersList.reverse()
} 
      return   { managersList, reportsLists, currentUserProfile } ;
    },
    []
  );
  return   { getUserProfile }  ;
};

function getEmailsForManagers(employee) {
  if (!employee || !employee.manager) {
    return [];
  }
  const managerEmails = [];
  const managers = Array.isArray(employee.manager) ? employee.manager : [employee.manager];
  managers.forEach(manager => {
    if (manager.id) {
      managerEmails.push(manager.id);
    }
    const subManagerEmails = getEmailsForManagers(manager);
    managerEmails.push(...subManagerEmails);
  });
  
  console.log(managerEmails)
  return managerEmails;
}
function getEmailsForDirectReports(directReports) {
  if (!Array.isArray(directReports)) {
    return [];
  }
  const directReportEmails = [];
  directReports.forEach(report => {
    if (report.id) {
      directReportEmails.push(report.id);
    }
    
  });

  return directReportEmails;
}


export function manpingUserProperties(userProperties: IPersonProperties, adress?: string, backupDataProperty?: string): IUserInfo{

 try{
   const newObject: IUserInfo = {
     displayName: userProperties.displayName as string,
     mail:userProperties.mail,
     title: userProperties.jobTitle as string,
     pictureUrl: `${adress}/_layouts/15/userphoto.aspx?size=L&accountname=${userProperties.userPrincipalName}`,
     numberDirectReports: userProperties.directReport && userProperties.directReport.length > 0 ?  userProperties.directReport.length : 0,
     hasDirectReports: userProperties.directReport && userProperties.directReport.length > 0 ? true : false,
     office: userProperties.officeLocation ?? '',
     manager: userProperties?.manager?.displayName ? userProperties?.manager?.displayName : backupDataProperty,
     loginName: userProperties.userPrincipalName,
     onPremisesExtensionAttributes: userProperties?.onPremisesExtensionAttributes
   };
   return newObject

 } catch (error){

 }
}