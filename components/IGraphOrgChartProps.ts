import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls";

export interface IGraphOrgChartProps {
  title: string;
  defaultUser: string;
  context: WebPartContext;
  startFromUser: IPropertyFieldGroupOrPerson[];
  showAllManagers: boolean;
  showActionsBar:boolean;
  extensionAttributeValue: string;
  filterByCompanyName: string;
}
