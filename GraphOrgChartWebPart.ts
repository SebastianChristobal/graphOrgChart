import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import * as strings from 'GraphOrgChartWebPartStrings';
import GraphOrgChart from './components/GraphOrgChart';
import { IGraphOrgChartProps } from './components/IGraphOrgChartProps';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls';

export interface IGraphOrgChartWebPartProps {
  description: string;
  title: string;
  currentUser: string;
  selectedUser: IPropertyFieldGroupOrPerson[];
  showAllManagers: boolean;
  showActionsBar: boolean;
  filterByCompanyName: string;
  dropdownFieldChoice:string;
  
}
/*eslint-disable*/
export default class GraphOrgChartWebPart extends BaseClientSideWebPart<IGraphOrgChartWebPartProps> {
  public dropdownOptions: IPropertyPaneDropdownOption[] = [
    { key: 'extensionAttribute1', text: 'extensionAttribute1' },
    { key: 'extensionAttribute2', text: 'extensionAttribute2' },  
    { key: 'extensionAttribute3', text: 'extensionAttribute3' },
    { key: 'extensionAttribute4', text: 'extensionAttribute4' },
    { key: 'extensionAttribute5', text: 'extensionAttribute5' },
    { key: 'extensionAttribute6', text: 'extensionAttribute6' },
    { key: 'extensionAttribute7', text: 'extensionAttribute7' },
    { key: 'extensionAttribute8', text: 'extensionAttribute8' },
    { key: 'extensionAttribute9', text: 'extensionAttribute9' },
    { key: 'extensionAttribute10', text: 'extensionAttribute10' },
    { key: 'extensionAttribute11', text: 'extensionAttribute11' },
    { key: 'extensionAttribute12', text: 'extensionAttribute12' },
    { key: 'extensionAttribute13', text: 'extensionAttribute13' },
    { key: 'extensionAttribute14', text: 'extensionAttribute14' },
    { key: 'extensionAttribute15', text: 'extensionAttribute15' },
  ];
  
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';


  public render(): void {
    const element: React.ReactElement<IGraphOrgChartProps> = React.createElement(
      GraphOrgChart,
      {
        // graphService: this._graphClient,
        title: this.properties.title,
        defaultUser: this.context.pageContext.user.email,
        startFromUser: this.properties.selectedUser,
        showAllManagers: this.properties.showAllManagers,
        context: this.context,
        showActionsBar: this.properties.showActionsBar,
        extensionAttributeValue: this.properties.dropdownFieldChoice,
        filterByCompanyName: this.properties.filterByCompanyName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }


protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
  if (propertyPath === 'dropdownFieldOptions') {
   this.properties.dropdownFieldChoice = newValue
  }
  if(propertyPath === 'filterByCompanyName'){
    this.properties.filterByCompanyName === newValue
  }
}

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("title", {
                  label: "Titel",
                }),
                PropertyFieldPeoplePicker("selectedUser", {
                  context: this.context as any,
                  label: "Start from",
                  initialData: this.properties.selectedUser,
                  key: "peopleFieldId",
                  multiSelect: false,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  onGetErrorMessage: null,
                }),
                PropertyPaneToggle("showAllManagers", {
                  label: "Show all managers",
                }),
                PropertyPaneToggle("showActionsBar", {
                  label: "Show actionbar",
                }),
                PropertyPaneDropdown('dropdownFieldOptions', {
                  label: 'Select property to search for',
                  options: this.dropdownOptions,
                  
              }),
              PropertyPaneTextField("filterByCompanyName", {
                label: "Filtrera på företagsnamn",
              }),
              ]
            }
          ]
        }
      ]
    };
  }
}
