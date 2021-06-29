import * as ko from 'knockout';
import styles from './TaskList.module.scss';
import { ITaskListWebPartProps } from './TaskListWebPart';
import * as moment from 'moment';
import { Web } from 'sp-pnp-js';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as strings from 'TaskListWebPartStrings';
require('./TaskList.scss');
require('datatables.net-responsive');
const $ = require('jquery');

export interface ITaskListBindingContext extends ITaskListWebPartProps {
  shouter: KnockoutSubscribable<{}>;
}

export interface IListViewHeader {
  Title: string;
}
export interface IListViewHeaders {
  value: IListViewHeader[];
}

export interface IDict {
  key: string;
  value: string;
  url: string;
}

export class ISiteUser {
  Id: string;
  UserTitle: string;
}

export default class TaskListViewModel {
  public selectedList: KnockoutObservable<string> = ko.observable('');
  public selectedView: KnockoutObservable<string> = ko.observable('');
  private BaseUrl: KnockoutObservable<string> = ko.observable('');
  public viewName: KnockoutObservable<string> = ko.observable('');
  public tblViewName: KnockoutObservable<string> = ko.observable('');
  public listViewID: KnockoutObservable<string> = ko.observable('');
  public listViewHeaders: KnockoutObservableArray<IListViewHeader> = ko.observableArray([]);
  private taskListItems: KnockoutObservableArray<any> = ko.observableArray([]);
  public meetingMinutestypeId: string = "";

  public taskListClass: string = styles.taskList;
  public containerClass: string = styles.container;
  public rowClass: string = styles.row;
  public columnClass: string = styles.column;
  public titleClass: string = styles.title;
  public subTitleClass: string = styles.subTitle;
  public descriptionClass: string = styles.description;
  public buttonClass: string = styles.button;
  public labelClass: string = styles.label;

  constructor(bindings: ITaskListBindingContext) {
    SPComponentLoader.loadCss("https://fonts.googleapis.com/css?family=Lora:400,400i,700,700i");
    SPComponentLoader.loadCss("https://fonts.googleapis.com/css?family=Open+Sans:300,400,600,700");
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css");
    SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.13/css/jquery.dataTables.min.css');
    SPComponentLoader.loadCss('https://cdn.datatables.net/responsive/2.1.1/css/responsive.dataTables.min.css');

    this.selectedList(bindings.listTitle);
    this.selectedView(bindings.listViewName);

    var url = bindings.context.pageContext.web.absoluteUrl;
    this.BaseUrl(url);

    bindings.shouter.subscribe((value: string) => {
      this.selectedList(value);
    }, this, 'listTitle');
    bindings.shouter.subscribe((value: string) => {
      this.selectedView(value);
    }, this, 'listViewName');

    // Call of the wild
    if (this.selectedList() && this.selectedList() !== '' && this.selectedView() && this.selectedView() !== '') {
      this.LoadAllSiteUsers();
    }
    else {
      $('#spnMessage').css('display', 'block');
    }
  }

  /**
   * DisplayListView
   */
  private async DisplayListView(web: Web, listId: string, viewId: string, userColl: Array<ISiteUser>) {
    let contentTypes = await web.contentTypes.select("Name, Id/StringValue").expand("Id").get();
    contentTypes.forEach(contentType => {
      if(contentType["Name"] == "Meeting Minutes"){
        this.meetingMinutestypeId = contentType.Id.StringValue;
      }
    });

    let columnsHeaderFields= [];
    let val = await web.lists.getById(listId).views.getById(viewId).fields.get();
    let viewHeaders = val ? val.Items : val;
    
    if (viewHeaders) {
      viewHeaders.forEach((element: any) => {
        let header: string;
        switch (element) {
          case "AssignedTo": header = strings.AssignedTo;
          break;
          case "DueDate": header = strings.DueDate;
          break;
          case "Goedkeuring_budget_aanvragen": header = strings.Goedkeuring;
          break;
          case "ID": header = element;
          break;
          case "Kostenplaats_x003_kostenplaatsv": header = strings.Kostenplaats;
          break;  
          case "OpmerkingBudgetBeheerder":header = strings.Opmerking;
          break;                  
          case "LinkTitleNoMenu": header = strings.LinkTitle;
          break;
          case "LinkTitle": header = strings.LinkTitle;
          break;
          case "Terugkoppeling_x0020_CO": header = strings.TerugkoppelingCO;
          break;          
          case "StartDate": header = strings.StartDate;
          break;
          case "Status_x0020_WPLU": header = strings.StatusWPLU;
          break;
          default:
            if (element.split('_x0020_').length > 1) {
              header = element.split('_x0020_').join(" ");
            } else if (element.split('_').length > 1) {
              header = element.split('_').join(" ");
            } else {
              header = element;
            }
          break;
        }
        
        let ch: IListViewHeader = {
          Title: header,
        };
        this.listViewHeaders.push(ch);

        columnsHeaderFields.push(element);
      });
    }

    let _viewXML: string = "";
    let xx = await web.lists.getById(listId).views.getById(viewId).get();
    _viewXML = xx.ListViewXml;
    _viewXML = _viewXML.replace('<FieldRef Name="LinkTitleNoMenu" />', '<FieldRef Name="LinkTitle" />');

    let items = await web.lists.getById(listId).getItemsByCAMLQuery({ ViewXml: _viewXML });
    items.forEach((element: { [x: string]: any; }) => {
      let itmValue: Array<IDict> = new Array<IDict>();
      columnsHeaderFields.forEach((x,) => {
        switch (x) {
          case "AssignedTo":
            //this is to get user info based on user id
            var result = userColl.filter(IUser => IUser.Id == element['AssignedToId']);
            if (result && result.length !== 0) {
              itmValue.push({ key: x, value: result[0]['UserTitle'], url: null });
            }
            else {
              itmValue.push({ key: strings.AssignedTo, value: '', url: null });
            }
            break;
          case "Predecessors":
            itmValue.push({ key: x, value: element['PredecessorsId'], url: null });                  
            break;
          case "WorkflowLink":
            if(element['WorkflowLink']) {
              itmValue.push({ key: x, value: element['WorkflowLink']['Description'], url: element['WorkflowLink']['Url'] });
            }else {
              itmValue.push({ key: x, value: "", url: "" });
            }
            break;
          case "LinkTitle":
            let actionUrl = "";
            if(element["ContentTypeId"] ? element["ContentTypeId"].indexOf(this.meetingMinutestypeId)>-1 : false){
              actionUrl = this.BaseUrl() + '/Lists/Tasks/EditForm.aspx?ID=' + element['Id'] + '&Source=' + this.BaseUrl() + '/SitePages/Dashboard.aspx';
            }
            else{
              actionUrl = this.BaseUrl() + '/_layouts/15/WrkTaskIP.aspx?List=' + this.selectedList() + '&ID=' + element['Id'] + '&Source=' + this.BaseUrl() + '/SitePages/Dashboard.aspx' + '&ContentTypeId=' + element['ContentTypeId'];
            }
            itmValue.push({ key: strings.LinkTitle, value: element['Title'], url: actionUrl });
            break;
          case "LinkTitleNoMenu":
            if(element["ContentTypeId"] ? element["ContentTypeId"].indexOf(this.meetingMinutestypeId)>-1 : false){
              actionUrl = this.BaseUrl() + '/Lists/Tasks/EditForm.aspx?ID=' + element['Id'] + '&Source=' + this.BaseUrl() + '/SitePages/Dashboard.aspx';
            }
            else{
              actionUrl = this.BaseUrl() + '/_layouts/15/WrkTaskIP.aspx?List=' + this.selectedList() + '&ID=' + element['Id'] + '&Source=' + this.BaseUrl() + '/SitePages/Dashboard.aspx' + '&ContentTypeId=' + element['ContentTypeId'];
            }
            itmValue.push({ key: strings.LinkTitle, value: element['Title'], url: actionUrl });
            break;
          case "DueDate":
            if (element['DueDate']) {
              itmValue.push({ key: strings.DueDate, value: moment(new Date(element['DueDate'])).format("DD/MM/YYYY"), url: null });
            } else {
              itmValue.push({ key: strings.DueDate, value: "", url: null });
            }                                  
            break;
          case "StartDate":
            if (element['StartDate']) {
              itmValue.push({ key: strings.StartDate, value: moment(new Date(element['StartDate'])).format("DD/MM/YYYY"), url: null });
            } else {
              itmValue.push({ key: strings.StartDate, value: "", url: null });
            }                                  
            break;
          default:
            itmValue.push({ key: x, value: element[x], url: null });
            break;
        }
      });
      this.taskListItems.push(itmValue);
    });

    let tableId: string = '#tbl' + this.tblViewName() + this.listViewID();
    $(tableId).DataTable({
      responsive: true,
      "lengthMenu": [[5, 10, 25], [5, 10, 25]],
      initComplete: () => {
      }
    });
    $('#span' + this.tblViewName()).css('display', 'none');
  }

  /**
   * LoadAllSiteUsers
   */
  public LoadAllSiteUsers() {
    var web = new Web(this.BaseUrl());

    web.siteUsers.get().then(u => {
      let siteUsersCollection: Array<ISiteUser> = new Array<ISiteUser>();
      u.forEach(el => {
        let iSiteUser: ISiteUser = new ISiteUser();
        iSiteUser.Id = el.Id;
        iSiteUser.UserTitle = el.Title;
        siteUsersCollection.push(iSiteUser);
      });
      return siteUsersCollection;
    }).then(_userColl => {
      let ListID = this.selectedList();
      let ViewID = this.selectedView();
      // get view Name
      web.lists.getById(ListID).views.getById(ViewID).get().then(v => {
        this.viewName(v.Title);
        this.tblViewName(v.Title.replace(/ /g, ''));
        this.listViewID(ViewID);
      });
      //
      this.DisplayListView(web, ListID, ViewID, _userColl);
    });
  }
}
