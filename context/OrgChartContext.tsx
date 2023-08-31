import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IOrgChartContext {
  context: WebPartContext;
  extensionAttributeValue: string
  // Add any other global data here
}

export const OrgChartContext = React.createContext<IOrgChartContext>({
  context: undefined,
  extensionAttributeValue: ""
  // Set any other default values here
});