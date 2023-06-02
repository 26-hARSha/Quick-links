import * as React from "react";
//import styles from "./QuickLink.module.scss";
import { IQuickLinkProps } from "./IQuickLinkProps";
//mport { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import styles from "./QuickLink.module.scss";
//single item
interface ICardlink {
  Title: string;
  RedirectLink: any;
  Logo: string;
  Number: number;
  IsActive:string;
}
//multiple items
interface IAllCardLinks {
  AllLinks: ICardlink[];
}

export default class QuickLinkWebpart extends React.Component<
  IQuickLinkProps,
  IAllCardLinks
> {
  properties: any;
  constructor(props: IQuickLinkProps, state: IAllCardLinks) {
    super(props);
    this.state = {
      AllLinks: [],
    };
  }

  componentDidMount() {
    //alert ("Componenet Did Mount Called...");
    //console.log("First Call.....");
    this.getAllCardLinks();
  }

  public getAllCardLinks = () => {
    console.log("This is link Detail function");
    //api call
    let listurl = `${this.props.SiteURL}/_api/lists/GetByTitle('${this.props.listName}')/items`;
    console.log(listurl);

    this.props.context.spHttpClient
      .get(listurl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: any) => {
          //console.log(responseJSON);
          this.setState({ AllLinks: responseJSON.value });
        });
        console.log(this.state.AllLinks);
      });
  };
  public render(): React.ReactElement<IQuickLinkProps> {
    let SelectedNumberofColumns = parseInt(this.props.NumbrtofColumnsToShow);
    let columnWidth: string;

    if (SelectedNumberofColumns == 4) {
      columnWidth = "23%";
    } else if (SelectedNumberofColumns == 3) {
      columnWidth = "30%";
    } else if (SelectedNumberofColumns == 2) {
      columnWidth = "48%";
    } else columnWidth = "100%";

    return (
      /*  Component Title */
      <div className={styles["cquicklinks"]}>
        <div>
          <p className={styles["componentTitle"]}>
            {this.props.componentTitle}
          </p>
        </div>

        {/*   Empty Message */}
        <div
          style={{ display: this.state.AllLinks.length === 0 ? "" : "none" }}
        >
          <p>{this.props.emptyMessage}</p>
        </div>
        <div className={styles["allLinksContainer"]}>
          {this.state.AllLinks.map((link) => {
            return (
              <div
                className={styles["quicklink-card"]}
                style={{
                  backgroundColor: this.props.viewType
                    ? "  rgb(196, 54, 252)"
                    : "  rgb(217, 159, 240)",
                  width: columnWidth,
                }}
                onClick={() => window.open(link.RedirectLink.SiteURL, "_blank")}
              >
                <img
                  src={
                    link.Logo == null
                      ? require("./Image/nologo.png")
                      : window.location.origin +
                        JSON.parse(link.Logo).serverRelativeUrl
                  }
                  alt=""
                />
                <p>{link.Title}</p>
                {/*  <p>{link.Number}</p> */}
              </div>
            );
          })}
        </div>
      </div>
    );
  }
}
