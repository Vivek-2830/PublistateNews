import * as React from "react";
import styles from "./PublistatNews.module.scss";
import { IPublistatNewsProps } from "./IPublistatNewsProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { IFieldInfo, Social, sp } from "@pnp/sp/presets/all";
import {
  DatePicker,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dialog,
  DialogFooter,
  IColumn,
  Icon,
  IIconProps,
  PrimaryButton,
  SearchBox,
  Selection,
  SelectionMode,
  TextField,
} from "office-ui-fabric-react";
import * as moment from "moment";
import * as Excel from "exceljs";
import { saveAs } from "file-saver";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
  IHttpClientOptions,
  HttpClient,
} from "@microsoft/sp-http";

require("../assets/css/fabric.min.css");
require("../assets/css/style.css");

export interface IPublistatNewsState {
  AllNews: any;
  AddTagDialog: boolean;
  CurrentUserName: any;
  AllUsers: any;
  CurrentEmail: any;
  AddFormTag: any;
  AddFormSubscribed: boolean;
  AddFormSendNotifications: boolean;
  MyNewsTags: any;
  MySubscribedTags: any;
  MyNews: any;
  MyNewsFilterData: any;
  MySavedNews: any;
  FilterDialog: boolean;
  ExportData: any;
  FilteredExportData: any;
  searchText: string;
  startDate: any;
  endDate: any;
  selectedItems: any;
  selectionDetails: any;
  EmailDialog: boolean;
  RecevierEmailID: any;
  EmailSuccessDialog: boolean;
}

const dialogContentProps = {
  title: "Add News Tag",
};
const FilterdialogContentProps = {
  title: "Export News",
};
const SendEmaildialogContentProps = {
  title: "Send Mail",
};
const EmailSuccessDialogContentProps = {
  title: "Mail Sent",
  subText: "The Mail has been sent successfully.",
};

const follow: IIconProps = { iconName: "Accept" };
const unfollow: IIconProps = { iconName: "Cancel" };

const NotifyTrue: IIconProps = { iconName: "RingerSolid" };
const NotifyFalse: IIconProps = { iconName: "RingerOff" };

const Export: IIconProps = { iconName: "DownloadDocument" };
const SendMail: IIconProps = { iconName: "MailLowImportance" };

const FlowURL = {
  SendMail:
    "https://prod-161.westeurope.logic.azure.com:443/workflows/82e57b7dd0404f2a90cce36397e0ccb1/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=c2w2VR_Z54FydTdWmab5DkRNtudapFfFB1a5S_n9zTc",
};

const columns: IColumn[] = [
  {
    key: "Title",
    name: "Title",
    fieldName: "Title",
    minWidth: 50,
    maxWidth: 350,
    isResizable: true,
    onRender: (item) => {
      return (
        <a
          href={item.Link}
          style={{ textDecoration: "none", color: "#006eb5" }}
        >
          <div>
            <span>
              {item.Source}: {item.Title}
            </span>
          </div>
        </a>
      );
    },
  },
  {
    key: "Source",
    name: "Source",
    fieldName: "Source",
    minWidth: 50,
    maxWidth: 120,
    isResizable: true,
  },
  {
    key: "Pubdate",
    name: "Publish Date",
    fieldName: "Pubdate",
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
    onRender: (item) => {
      return <span>{moment(new Date(item.Pubdate)).format("Do MMM")}</span>;
    },
  },
];

let XLcolums = [
  { header: "News Title", key: "Title" },
  { header: "Source", key: "Source" },
  { header: "Publish Date", key: "Pubdate" },
  { header: "URL", key: "Link" },
];

export default class PublistatNews extends React.Component<
  IPublistatNewsProps,
  IPublistatNewsState
> {
  private _selection: Selection;
  constructor(props: IPublistatNewsProps, state: IPublistatNewsState) {
    super(props);

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails(),
        });
      },
      getKey: this._getKey,
    });

    this.state = {
      AllNews: [],
      AddTagDialog: true,
      CurrentUserName: "",
      AllUsers: "",
      CurrentEmail: "",
      AddFormTag: "",
      AddFormSubscribed: true,
      AddFormSendNotifications: true,
      MyNewsTags: [],
      MySubscribedTags: [],
      MyNews: [],
      MyNewsFilterData: [],
      MySavedNews: [],
      FilterDialog: true,
      ExportData: [],
      FilteredExportData: [],
      searchText: "",
      startDate: "",
      endDate: "",
      selectedItems: [],
      selectionDetails: this._getSelectionDetails(),
      EmailDialog: true,
      RecevierEmailID: [],
      EmailSuccessDialog: true,
    };
  }

  public render(): React.ReactElement<IPublistatNewsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section id="PublistatNews">
        <div className="ms-Grid">
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg8 ms-xl8">
              <div className="d-flex-header">
                <h2 className="Publistat-header">Your News</h2>
                <PrimaryButton
                  text="Export"
                  iconProps={Export}
                  onClick={() => this.setState({ FilterDialog: false })}
                ></PrimaryButton>
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg6 ms-xl6">
                <SearchBox
                  placeholder="Search"
                  className="new-search"
                  onChange={(e) => this.SearchMyNews(e.target.value)}
                  onClear={() =>
                    this.setState({ MyNews: this.state.ExportData })
                  }
                />
              </div>
            </div>
          </div>

          <div className="ms-Grid-row flex-wrap-m">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg8 ms-xl8">
              <h4 className="Mynewstitle">Your News</h4>
              {this.state.MyNews.length == 0 ? (
                <div
                  style={{
                    textAlign: "center",
                    backgroundColor: "#ffffff",
                    paddingBottom: "40px",
                  }}
                >
                  <img
                    className="NewsError-img"
                    src={require("../assets/Images/newspaper.png")}
                  />
                  <h4 className="NewsError-msg">
                    Welcome to your personalized news page.
                  </h4>
                </div>
              ) : (
                <></>
              )}
              {this.state.MyNews.length > 0 &&
                this.state.MyNews.map((item) => {
                  const words = item.Title.split(" ");
                  const firstLetters = words[0][0] + words[1][0];

                  return (
                    <>
                      <div className="Publistat-Newcard">
                        <div className="Title">
                          <div className="Title-Avatar">
                            <p>{firstLetters.toUpperCase()}</p>
                          </div>
                          <div className="News-Title">
                            <p className="Publistat-Newsdate">
                              {moment(new Date(item.Pubdate)).format("Do MMM")}{" "}
                              <span>- {item.Category}</span>{" "}
                            </p>

                            <a
                              className="NewsLink"
                              href={item.Link}
                              data-interception="off"
                              target="_blank"
                            >
                              <h4 className="Publistat-Newstitle">
                                <span>{item.Source}</span>: {item.Title}
                              </h4>
                            </a>
                          </div>
                          <div>
                            <Icon
                              className="SaveNews-Icon"
                              onClick={() =>
                                this.MarkAsSave(
                                  item.Title,
                                  item.Link,
                                  item.Date,
                                  item.Source
                                )
                              }
                              iconName="Pinned"
                            ></Icon>
                          </div>
                        </div>
                      </div>
                    </>
                  );
                })}
            </div>
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg4 ms-xl4">
              <div className="Subscription-area Subscription-area-position">
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <h4 className="areatitle">Subscribe News</h4>
                  </div>
                  <div className="ms-Grid-col ms-sm9 ms-md9 ms-lg9">
                    <TextField
                      className="AddTag-textfield"
                      placeholder="Add new topics"
                      onChange={(value) =>
                        this.setState({ AddFormTag: value.target["value"] })
                      }
                      value={this.state.AddFormTag}
                    />
                  </div>
                  <div className="ms-Grid-col ms-sm3 ms-md3 ms-lg3">
                    <div className="text-right mb-20">
                      <PrimaryButton
                        text="Add"
                        onClick={() => this.AddTags()}
                      ></PrimaryButton>
                    </div>
                  </div>
                </div>

                {this.state.MyNewsTags.length == 0 ? (
                  <div style={{ textAlign: "center" }}>
                    <img
                      className="MyNewsTag-img"
                      src={require("../assets/Images/bell.png")}
                    />
                    <h4 className="MyNewsTag-msg">
                      Easily add and manage the topics that interest you.
                    </h4>
                  </div>
                ) : (
                  <></>
                )}
                {this.state.MyNewsTags.length > 0 &&
                  this.state.MyNewsTags.map((item) => {
                    return (
                      <>
                        <div className="ms-Grid-row Mytag-wrapper">
                          <p className="ms-Grid-col ms-sm6 ms-md6 ms-lg6 MyTag-title">
                            {item.NewsTag}
                          </p>
                          <div className="ms-Grid-col ms-sm4 ms-md4 ms-lg4">
                            <DefaultButton
                              toggle
                              checked={item.Subscribed == true ? true : false}
                              text={
                                item.Subscribed == true ? "Unfollow" : "Follow"
                              }
                              iconProps={
                                item.Subscribed == true ? unfollow : follow
                              }
                              onClick={() =>
                                this.UpdateSubscription(
                                  item.ID,
                                  item.Subscribed
                                )
                              }
                              allowDisabledFocus
                              className="SubscriptionBtn"
                            />
                          </div>
                          <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2 text-center">
                            <DefaultButton
                              toggle
                              checked={
                                item.SendNotifications == true ? true : false
                              }
                              iconProps={
                                item.SendNotifications == true
                                  ? NotifyTrue
                                  : NotifyFalse
                              }
                              onClick={() =>
                                this.UpdateNotifications(
                                  item.ID,
                                  item.SendNotifications
                                )
                              }
                              allowDisabledFocus
                              title={
                                item.SendNotifications == true
                                  ? "Don't Notify"
                                  : "Notify me"
                              }
                              className="NotificationBtn"
                            />
                          </div>
                        </div>
                      </>
                    );
                  })}
              </div>

              <div className="Subscription-area">
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <h4 className="areatitle">Saved News</h4>
                  </div>
                  <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    {this.state.MySavedNews.length > 0 &&
                      this.state.MySavedNews.map((item) => {
                        return (
                          <>
                            <div className="Saved-Newcard">
                              <p className="Saved-Newsdate">
                                {moment(new Date(item.Pubdate)).format(
                                  "Do MMM"
                                )}
                              </p>
                              <Icon
                                iconName="Cancel"
                                onClick={() => this.Unsave(item.ID)}
                              ></Icon>
                              <a
                                className="NewsLink"
                                href={item.Link}
                                data-interception="off"
                                target="_blank"
                              >
                                <h4 className="Saved-Newstitle">
                                  <span>{item.Source}</span>: {item.Title}
                                </h4>
                              </a>
                            </div>
                          </>
                        );
                      })}
                  </div>
                </div>
              </div>

              <div className="UserName">
                <h3>AllUser - Details</h3>
                {this.state.AllUsers.length > 0 &&
                  this.state.AllUsers.map((item) => {
                    return (
                      <>
                        <div className="User-card">
                          <div className="AllUser">
                            <img
                              src={
                                this.props.context.pageContext.web.absoluteUrl +
                                `/_layouts/15/userphoto.aspx?UserName=${item.Email}&size=L`
                              }
                              draggable="false"
                            />
                            {/* <div className="User-Name">
                                  <span>{this.props.context.item.Title}</span>
                                </div> */}
                          </div>
                          <div className="User-Details">
                            <p>{item.Title}</p>
                            <p className="mail">{item.Email}</p>
                          </div>
                        </div>
                      </>
                    );
                  })}
              </div>
            </div>
          </div>

          <Dialog
            hidden={this.state.AddTagDialog}
            onDismiss={() => this.setState({ AddTagDialog: true })}
            dialogContentProps={dialogContentProps}
            minWidth={450}
          >
            <div>
              <TextField
                label="Tag"
                onChange={(value) =>
                  this.setState({ AddFormTag: value.target["value"] })
                }
                value={this.state.AddFormTag}
              />
            </div>
            <DialogFooter>
              <PrimaryButton text="Add" onClick={() => this.AddTags()} />
              <DefaultButton
                onClick={() => this.setState({ AddTagDialog: true })}
                text="Cancel"
              />
            </DialogFooter>
          </Dialog>

          <Dialog
            hidden={this.state.FilterDialog}
            onDismiss={() =>
              this.setState({
                FilterDialog: true,
                searchText: "",
                startDate: "",
                endDate: "",
                FilteredExportData: this.state.ExportData,
              })
            }
            dialogContentProps={FilterdialogContentProps}
            minWidth={800}
          >
            <div>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm4 ms-md4 ms-lg4">
                  <SearchBox
                    placeholder="Search"
                    onChange={this.handleSearchChange}
                  />
                </div>
                <div className="ms-Grid-col ms-sm4 ms-md4 ms-lg4">
                  <DatePicker
                    placeholder="Select Start date..."
                    ariaLabel="Select a date"
                    onSelectDate={this.handleStartDateChange}
                    value={this.state.startDate}
                  />
                </div>
                <div className="ms-Grid-col ms-sm4 ms-md4 ms-lg4">
                  <DatePicker
                    placeholder="Select End date..."
                    ariaLabel="Select a date"
                    onSelectDate={this.handleEndDateChange}
                    value={this.state.endDate}
                  />
                </div>
                <div
                  className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 mt-10"
                  style={{ textAlign: "right" }}
                >
                  <PrimaryButton
                    text="Export"
                    onClick={() => this.saveExcel()}
                    iconProps={Export}
                  />
                  <PrimaryButton
                    text="Send Mail"
                    className="ml-15"
                    onClick={() => this.setState({ EmailDialog: false })}
                    iconProps={SendMail}
                  />
                </div>
              </div>
              <div>{this.state.selectionDetails}</div>
              <DetailsList
                items={this.state.FilteredExportData}
                // compact={isCompactMode}
                columns={columns}
                selectionMode={SelectionMode.multiple}
                setKey="multiple"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                selection={this._selection}
                selectionPreservedOnEmptyClick={true}
                // onItemInvoked={this._onItemInvoked}
                enterModalSelectionOnTouch={true}
                ariaLabelForSelectionColumn="Toggle selection"
                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                checkButtonAriaLabel="select row"
              />
            </div>
          </Dialog>

          <Dialog
            hidden={this.state.EmailDialog}
            // onDismiss={toggleEmailDialog}
            dialogContentProps={SendEmaildialogContentProps}
            minWidth={400}
          >
            <TextField
              label="Recipient's Email"
              onChange={(e) =>
                this.setState({ RecevierEmailID: e.target["value"] })
              }
            />
            <DialogFooter>
              <PrimaryButton
                onClick={() =>
                  this.triggerFlow(FlowURL.SendMail, this.state.selectedItems)
                }
                text="Send"
              />
              <DefaultButton
                onClick={() => this.setState({ EmailDialog: true })}
                text="Don't send"
              />
            </DialogFooter>
          </Dialog>
          <Dialog
            hidden={this.state.EmailSuccessDialog}
            // onDismiss={toggleHideDialog}
            dialogContentProps={EmailSuccessDialogContentProps}
            minWidth={400}
          >
            <DialogFooter>
              <PrimaryButton
                onClick={() => this.setState({ EmailSuccessDialog: true })}
                text="Ok"
              />
            </DialogFooter>
          </Dialog>
        </div>

        {/* <div id="avatar" className="avatar">
              <span id="avatar-initials" className="avatar-initials"></span>
              <span id="title" className="title"></span>
          </div> */}
      </section>
    );
  }

  public async componentDidMount(): Promise<void> {
    await this.HideNavigation();
    await this.GetCurrentUser();
    await this.GetMySubscribedTags();
    await this.GetMyTags();
    await this.GetSavedNews();
    await this.GetAllUser();
    // this.generateAvatar("Title");
  }

  public async GetCurrentUser() {
    let user = await sp.web.currentUser.get();
    this.setState({ CurrentUserName: user.Title, CurrentEmail: user.Email });
  }

  public async GetAllUser() {
    let users = await sp.web.siteUsers();
    const filteredUsers = users.filter(
      (user) => user.PrincipalType == 1 && user.Email
    ); //&& user.UserId == null
    this.setState({ AllUsers: filteredUsers });
    console.log(this.state.AllUsers);
  }

  public async GetMySubscribedTags() {
    sp.web.lists
      .getByTitle("User Prefrence")
      .items.select("NewsTags")
      .filter(`Title eq '${this.state.CurrentUserName}' and Subscribed eq 1`)
      .get()
      .then((data) => {
        console.log(data);

        let tags = data
          .map((item) => item.NewsTags)
          .filter((tag) => tag !== undefined && tag !== null);
        this.setState({ MySubscribedTags: tags });
        console.log(this.state.MySubscribedTags);

        this.GetNews();
      })
      .catch((err) => {
        console.log(err);
      });
  }

  public async GetNews() {
    let items = [];
    let position = 0;
    const pageSize = 2000;
    let AllData = [];

    try {
      while (true) {
        const response = await sp.web.lists
          .getByTitle("News")
          .items.select(
            "Title",
            "Link",
            "Pubdate",
            "Description",
            "Date",
            "Source",
            "Newsgroup",
            "Category",
            "Newsguid"
          )
          .orderBy("Date", false)
          .top(pageSize)
          .skip(position)
          .get();
        if (response.length === 0) {
          break;
        }
        items = items.concat(response);
        position += pageSize;
      }
      console.log(`Total items retrieved: ${items.length}`);
      if (items.length > 0) {
        items.forEach((item, i) => {
          AllData.push({
            ID: item.Id ? item.Id : "",
            Title: item.Title ? item.Title : "",
            Link: item.Link ? item.Link : "",
            Pubdate: item.Pubdate
              ? new Date(
                  new Date(item.Date).setHours(
                    new Date(item.Pubdate).getHours() + 2
                  )
                )
                  .toISOString()
                  .split("T")[0]
              : "",
            Description: item.Description ? item.Description : "",
            Date: item.Date
              ? new Date(
                  new Date(item.Date).setHours(
                    new Date(item.Date).getHours() + 2
                  )
                )
                  .toISOString()
                  .split("T")[0]
              : "",
            Source: item.Source ? item.Source : "",
            Newsgroup: item.Newsgroup ? item.Newsgroup : "",
            Category: item.Category ? item.Category : "",
          });
        });
        this.setState({ AllNews: AllData });

        // let MySubTags = this.state.MySubscribedTags.map(tag => tag.toLowerCase());

        // let filteredData = this.state.AllNews.filter((x) => {
        //   let Title = x.Title;
        //   let Description = x.Description;
        //   let Category = x.Category;
        //   let Source = x.Source;
        //   let Newsgroup = x.Newsgroup;

        //   if (this.state.MySubscribedTags) {
        //     return MySubTags.some(tag => Title.includes(tag) || Description.includes(tag) || Category.includes(tag) || Source.includes(tag) || Newsgroup.includes(tag));
        //   }
        // });

        // console.log(filteredData);

        let MySubTags = this.state.MySubscribedTags.map(
          (tag) => `\\b${tag.toLowerCase()}\\b`
        );

        let filteredData = this.state.AllNews.filter((x) => {
          let Title = x.Title.toLowerCase();
          let Description = x.Description.toLowerCase();
          let Category = x.Category.toLowerCase();
          let Source = x.Source.toLowerCase();
          let Newsgroup = x.Newsgroup.toLowerCase();

          if (this.state.MySubscribedTags) {
            return MySubTags.some(
              (tag) =>
                new RegExp(tag, "i").test(Title) ||
                new RegExp(tag, "i").test(Description) ||
                new RegExp(tag, "i").test(Category) ||
                new RegExp(tag, "i").test(Source) ||
                new RegExp(tag, "i").test(Newsgroup)
            );
          }
        });
        this.setState({ MyNews: filteredData, MyNewsFilterData: filteredData });
        this.setState({
          ExportData: filteredData,
          FilteredExportData: filteredData,
        });
      }
    } catch (error) {
      console.error(error);
    }

    // sp.web.lists.getByTitle('News').items.select('Title', 'Link', 'Pubdate', 'Description', 'Date', 'Source', 'Newsgroup', 'Category', 'Newsguid').orderBy('Date',false).get()
    //   .then((data) => {
    //     let AllData = [];
    //     if (data.length > 0) {
    //       data.forEach((item, i) => {
    //         AllData.push({
    //           ID: item.Id ? item.Id : "",
    //           Title: item.Title ? item.Title : "",
    //           Link: item.Link.Url ? item.Link.Url : "",
    //           Pubdate: item.Pubdate ? item.Pubdate.split("T")[0] : "",
    //           Description: item.Description ? item.Description : "",
    //           Date: item.Date ? item.Date.split("T")[0]  : "",
    //           Source: item.Source ? item.Source : "",
    //           Newsgroup: item.Newsgroup ? item.Newsgroup : "",
    //           Category: item.Category ? item.Category : "",
    //         });
    //       });
    //       this.setState({ AllNews: AllData });

    //       let MySubTags = this.state.MySubscribedTags.map(tag => tag.toLowerCase());

    //       let filteredData = this.state.AllNews.filter((x) => {
    //         let Title = x.Title.toLowerCase();
    //         let Description = x.Description.toLowerCase();
    //         let Category = x.Category.toLowerCase();
    //         let Source = x.Source.toLowerCase();
    //         let Newsgroup = x.Newsgroup.toLowerCase();

    //         if (this.state.MySubscribedTags) {
    //           return MySubTags.some(tag => Title.includes(tag) || Description.includes(tag) || Category.includes(tag) || Source.includes(tag) || Newsgroup.includes(tag));
    //         }
    //       });

    //       console.log(filteredData);
    //       this.setState({ MyNews: filteredData });

    //     }
    //   }
    //   )
    //   .catch((err) => {
    //     console.log(err);
    //   });
  }

  public SearchMyNews(searchText: string): void {
    const { MyNewsFilterData } = this.state;

    const filteredItems = MyNewsFilterData.filter(
      (item) =>
        item.Title.toLowerCase().includes(searchText.toLowerCase()) ||
        item.Source.toLowerCase().includes(searchText.toLowerCase()) ||
        item.Category.toLowerCase().includes(searchText.toLowerCase())
    );

    // Update the state with the filtered items
    this.setState({ MyNews: filteredItems });
  }

  public async UpdateSubscription(ItemID, subscription) {
    await sp.web.lists
      .getByTitle("User Prefrence")
      .items.getById(ItemID)
      .delete();
    //  const Tags = await sp.web.lists.getByTitle("User Prefrence").items.getById(ItemID).update({
    //     Subscribed: subscription == true ? false : true ,
    //   }).catch((err) => {
    //     console.log(err);
    //   });
    await this.GetMyTags();
    await this.GetMySubscribedTags();
    // this.componentDidMount();
  }

  public async UpdateNotifications(ItemID, SendNotifications) {
    const Tags = await sp.web.lists
      .getByTitle("User Prefrence")
      .items.getById(ItemID)
      .update({
        SendNotifications: SendNotifications == true ? false : true,
      })
      .catch((err) => {
        console.log(err);
      });
    this.GetMyTags();
  }

  public async GetMyTags() {
    sp.web.lists
      .getByTitle("User Prefrence")
      .items.select(
        "Title",
        "Email",
        "NewsTags",
        "Subscribed",
        "SendNotifications",
        "Id"
      )
      .filter(`Title eq '${this.state.CurrentUserName}'`)
      .get()
      .then((data) => {
        let AllData = [];
        console.log(data);
        if (data.length > 0) {
          data.forEach((item, i) => {
            AllData.push({
              ID: item.Id ? item.Id : "",
              NewsTag: item.NewsTags ? item.NewsTags : "",
              Subscribed: item.Subscribed ? item.Subscribed : "",
              SendNotifications: item.SendNotifications
                ? item.SendNotifications
                : "",
            });
          });
        }
        this.setState({ MyNewsTags: AllData });
        console.log(this.state.MyNewsTags);
      })
      .catch((err) => {
        console.log(err);
      });
  }

  public async AddTags() {
    if (this.state.AddFormTag.length == 0) {
      alert("Please enter tag.");
    } else {
      const Tags = await sp.web.lists
        .getByTitle("User Prefrence")
        .items.add({
          Title: this.state.CurrentUserName,
          Email: this.state.CurrentEmail,
          NewsTags: this.state.AddFormTag,
          Subscribed: this.state.AddFormSubscribed,
          SendNotifications: this.state.AddFormSendNotifications,
        })
        .catch((err) => {
          console.log(err);
        });
      this.setState({ AddFormTag: "" });
      this.GetMySubscribedTags();
      this.GetMyTags();
    }
  }

  public async MarkAsSave(Title, URL, Date, Source) {
    await sp.web.lists
      .getByTitle("Saved News")
      .items.add({
        Title: Title,
        Link: URL,
        Source: Source,
        Pubdate: Date,
      })
      .catch((err) => {
        console.log(err);
      });
    this.GetSavedNews();
  }

  public async GetSavedNews() {
    sp.web.lists
      .getByTitle("Saved News")
      .items.select("Title", "Link", "Pubdate", "Author/Title", "Id", "Source")
      .expand("Author")
      .filter(`Author/Title eq '${this.state.CurrentUserName}'`)
      .orderBy("Pubdate", false)
      .get()
      .then((data) => {
        let AllData = [];
        console.log(data);
        if (data.length > 0) {
          data.forEach((item, i) => {
            AllData.push({
              ID: item.Id ? item.Id : "",
              Title: item.Title ? item.Title : "",
              Link: item.Link ? item.Link : "",
              Pubdate: item.Pubdate
                ? new Date(
                    new Date(item.Pubdate).setHours(
                      new Date(item.Pubdate).getHours() + 2
                    )
                  )
                    .toISOString()
                    .split("T")[0]
                : "",
              Source: item.Source ? item.Source : "",
            });
          });
        }
        this.setState({ MySavedNews: AllData });
        console.log(this.state.MySavedNews);
      })
      .catch((err) => {
        console.log(err);
      });
  }

  public async Unsave(ID) {
    await sp.web.lists.getByTitle("Saved News").items.getById(ID).delete();
    this.GetSavedNews();
  }

  public normalizeDate = (date: Date): Date => {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  };

  handleSearchChange = (
    event: React.ChangeEvent<HTMLInputElement>,
    newValue?: string
  ) => {
    const searchText = newValue || "";
    this.setState({ searchText }, this.applyFilters);
  };

  handleStartDateChange = (date: Date | null) => {
    this.setState({ startDate: date }, this.applyFilters);
  };

  handleEndDateChange = (date: Date | null) => {
    this.setState({ endDate: date }, this.applyFilters);
  };

  applyFilters = () => {
    const { ExportData, searchText, startDate, endDate } = this.state;
    const FilteredExportData = ExportData.filter((item) => {
      const Title = item.Title || "";
      const Source = item.Source || "";
      const date = this.normalizeDate(new Date(item.Pubdate));

      const matchesSearch =
        !searchText ||
        Title.toLowerCase().includes(searchText.toLowerCase()) ||
        Source.toLowerCase().includes(searchText.toLowerCase());
      const matchesStartDate =
        !startDate || date >= this.normalizeDate(startDate);
      const matchesEndDate = !endDate || date <= this.normalizeDate(endDate);

      return matchesSearch && matchesStartDate && matchesEndDate;
    });
    this.setState({ FilteredExportData });
    console.log(this.state.FilteredExportData);
  };

  private _getSelectionDetails() {
    const selectionCount = this._selection.getSelectedCount();

    let selecteditems = this._selection.getSelection();
    console.log(selecteditems);

    this.setState({ selectedItems: selecteditems });
    // console.log(this.state.selectedItems);
  }

  private _getKey(item: any, index?: number): string {
    return item.Title;
  }

  private saveExcel = async () => {
    const web = sp.web;
    const siteTitle = await web.select("Title").get();

    const workbook = new Excel.Workbook();

    if (this.state.selectedItems.length > 0) {
      try {
        const fileName =
          moment().format("DD/MM/YYYY HH:MM") +
          " Publistat Excel Overview " +
          siteTitle.Title;
        const worksheet = workbook.addWorksheet();

        // add worksheet columns
        // each columns contains header and its mapping key from data
        worksheet.columns = XLcolums;

        // updated the font for first row.
        worksheet.getRow(1).font = { bold: true, color: { argb: "00000000" } };
        worksheet.getRow(1).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "ffffffff" }, // Blue color#002e6d
        };

        // loop through all of the columns and set the alignment with width.
        // worksheet.columns.forEach(column => {
        //   column.width = 20;
        //   column.alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };
        // });

        worksheet.columns = [
          { width: 70 },
          { width: 30 },
          { width: 15 },
          { width: 50 },
        ];
        worksheet.getColumn(1).alignment = {
          horizontal: "left",
          wrapText: true,
          vertical: "middle",
        };
        worksheet.getColumn(2).alignment = {
          horizontal: "left",
          wrapText: true,
          vertical: "middle",
        };
        worksheet.getColumn(3).alignment = {
          horizontal: "left",
          wrapText: true,
          vertical: "middle",
        };
        worksheet.getColumn(4).alignment = {
          horizontal: "left",
          wrapText: true,
          vertical: "middle",
        };

        const oddRowColor = "FFFFFF"; // Lighter shade
        const evenRowColor = "fbfbfb"; // Darker shade
        const borderColor = "aaaaaa"; // Dark border color FFBFDEF7

        // Loop through data and add each one to worksheet
        this.state.selectedItems.forEach((singleData: any, index: number) => {
          const row = worksheet.addRow(singleData);

          // Set fill color based on odd or even row
          const fillColor =
            index % 2 === 0 ? { argb: oddRowColor } : { argb: evenRowColor };
          row.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: fillColor,
          };

          // Set border color for each cell in the row
          row.eachCell((cell, colNumber) => {
            cell.border = {
              top: { style: "thin", color: { argb: borderColor } },
              left: { style: "thin", color: { argb: borderColor } },
              bottom: { style: "thin", color: { argb: borderColor } },
              right: { style: "thin", color: { argb: borderColor } },
            };
          });
        });

        // write the content using writeBuffer
        const buf = await workbook.xlsx.writeBuffer();

        // download the processed file
        saveAs(new Blob([buf]), `${fileName}.xlsx`);
      } catch (error) {
        console.error("Something Went Wrong", error.message);
      }
    } else {
      alert(
        "Please select News you want to export, then click the 'Export' button."
      );
    }
  };

  public triggerFlow = (postURL, data) => {
    if (this.state.RecevierEmailID.length > 0) {
      if (this.state.selectedItems.length > 0) {
        this.setState({ EmailDialog: true, RecevierEmailID: "" });

        const mail = this.state.RecevierEmailID;
        const data1 = JSON.stringify({ data, mail });
        const body: string = data1;

        const requestHeaders: Headers = new Headers();
        requestHeaders.append("Content-type", "application/json");

        const httpClientOptions: IHttpClientOptions = {
          body: body,
          headers: requestHeaders,
        };

        return this.props.context.httpClient
          .post(postURL, HttpClient.configurations.v1, httpClientOptions)
          .then((response) => {
            console.log("Flow Triggered Successfully...");
            this.setState({ EmailDialog: true, EmailSuccessDialog: false });
          })
          .catch((error) => {
            console.log(error);
          });
      } else {
        alert(
          "Please select News you want to export, then click the 'Send Mail' button."
        );
      }
    } else {
      alert("Please Add the recipient's email address.");
    }
  };

  public async HideNavigation() {
    try {
      // Get current user's groups
      const userGroups = await sp.web.currentUser.groups();

      // Check if the user is in the Owners or Admins group
      const isAdmin = userGroups.some(
        (group) =>
          group.Title.indexOf("Owners") !== -1 ||
          group.Title.indexOf("Admins") !== -1
      );

      if (!isAdmin) {
        // Hide the navigation bar for non-admins
        const navBar = document.querySelector("#SuiteNavWrapper");
        if (navBar) {
          navBar.setAttribute("style", "display: none;");
        }
      } else {
        // Show the navigation bar for admins
        const navBar = document.querySelector("#SuiteNavWrapper");
        if (navBar) {
          navBar.setAttribute("style", "display: block;");
        }
      }
    } catch (error) {
      console.error("Error checking user permissions: ", error);
    }
  }
}
