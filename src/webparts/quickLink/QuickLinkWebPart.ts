import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "QuickLinkWebPartStrings";
import QuickLink from "./components/QuickLink";
import { IQuickLinkProps } from "./components/IQuickLinkProps";

export interface IQuickLinkWebPartProps {
  NumbrtofColumnsToShow: any;
  componentTitle: string;
  emptyMessage: string;
  DropDown: any;
  viewType: boolean;
  listName: string;
  description: string;
}

export default class QuickLinkWebPart extends BaseClientSideWebPart<IQuickLinkWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  public render(): void {
    const element: React.ReactElement<IQuickLinkProps> = React.createElement(
      QuickLink,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        SiteURL: this.context.pageContext.web.absoluteUrl,
        listName: this.properties.listName,
        context: this.context,
        viewType: this.properties.viewType,
        NumbrtofColumnsToShow: this.properties.NumbrtofColumnsToShow,
        emptyMessage: this.properties.emptyMessage,
        componentTitle: this.properties.componentTitle,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error("Unknown host");
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          
          groups: [
            {
             
              groupFields: [
                PropertyPaneTextField("componentTitle", {
                  label: "Component Title",
                }),
                PropertyPaneTextField("listName", {
                  label: "Enter list Name",
                }),
                PropertyPaneTextField("emptyMessage", {
                  label: "Text, If no data available",
                }),
                PropertyPaneToggle("viewType", {
                  label: "View Type",
                  onText: "Theme 1",
                  offText: "Theme",
                }),
                PropertyPaneDropdown("NumbrtofColumnsToShow", {
                  label: "Number of columns in row ",
                  options: [
                  
                    {
                      key: "4",
                      text: "4 Columns",
                    },
                    {
                      key: "3",
                      text: "3 Columns",
                    },
                    {
                      key: "2",
                      text: "2 Columns",
                    },
                    {
                      key: "1",
                      text: "1 Columns",
                    },
                  ],
                  selectedKey: "4",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
