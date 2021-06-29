"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var ko = require("knockout");
var TaskList_module_scss_1 = require("./TaskList.module.scss");
var moment = require("moment");
var sp_pnp_js_1 = require("sp-pnp-js");
var sp_loader_1 = require("@microsoft/sp-loader");
var strings = require("TaskListWebPartStrings");
require('./TaskList.scss');
require('datatables.net-responsive');
var $ = require('jquery');
var ISiteUser = (function () {
    function ISiteUser() {
    }
    return ISiteUser;
}());
exports.ISiteUser = ISiteUser;
var TaskListViewModel = (function () {
    function TaskListViewModel(bindings) {
        var _this = this;
        this.selectedList = ko.observable('');
        this.selectedView = ko.observable('');
        this.BaseUrl = ko.observable('');
        this.viewName = ko.observable('');
        this.tblViewName = ko.observable('');
        this.listViewID = ko.observable('');
        this.listViewHeaders = ko.observableArray([]);
        this.taskListItems = ko.observableArray([]);
        this.meetingMinutestypeId = "";
        this.taskListClass = TaskList_module_scss_1.default.taskList;
        this.containerClass = TaskList_module_scss_1.default.container;
        this.rowClass = TaskList_module_scss_1.default.row;
        this.columnClass = TaskList_module_scss_1.default.column;
        this.titleClass = TaskList_module_scss_1.default.title;
        this.subTitleClass = TaskList_module_scss_1.default.subTitle;
        this.descriptionClass = TaskList_module_scss_1.default.description;
        this.buttonClass = TaskList_module_scss_1.default.button;
        this.labelClass = TaskList_module_scss_1.default.label;
        sp_loader_1.SPComponentLoader.loadCss("https://fonts.googleapis.com/css?family=Lora:400,400i,700,700i");
        sp_loader_1.SPComponentLoader.loadCss("https://fonts.googleapis.com/css?family=Open+Sans:300,400,600,700");
        sp_loader_1.SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css");
        sp_loader_1.SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.13/css/jquery.dataTables.min.css');
        sp_loader_1.SPComponentLoader.loadCss('https://cdn.datatables.net/responsive/2.1.1/css/responsive.dataTables.min.css');
        this.selectedList(bindings.listTitle);
        this.selectedView(bindings.listViewName);
        var url = bindings.context.pageContext.web.absoluteUrl;
        this.BaseUrl(url);
        bindings.shouter.subscribe(function (value) {
            _this.selectedList(value);
        }, this, 'listTitle');
        bindings.shouter.subscribe(function (value) {
            _this.selectedView(value);
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
    TaskListViewModel.prototype.DisplayListView = function (web, listId, viewId, userColl) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var contentTypes, columnsHeaderFields, val, viewHeaders, _viewXML, xx, items, tableId;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, web.contentTypes.select("Name, Id/StringValue").expand("Id").get()];
                    case 1:
                        contentTypes = _a.sent();
                        contentTypes.forEach(function (contentType) {
                            if (contentType["Name"] == "Meeting Minutes") {
                                _this.meetingMinutestypeId = contentType.Id.StringValue;
                            }
                        });
                        columnsHeaderFields = [];
                        return [4 /*yield*/, web.lists.getById(listId).views.getById(viewId).fields.get()];
                    case 2:
                        val = _a.sent();
                        viewHeaders = val ? val.Items : val;
                        if (viewHeaders) {
                            viewHeaders.forEach(function (element) {
                                var header;
                                switch (element) {
                                    case "AssignedTo":
                                        header = strings.AssignedTo;
                                        break;
                                    case "DueDate":
                                        header = strings.DueDate;
                                        break;
                                    case "Goedkeuring_budget_aanvragen":
                                        header = strings.Goedkeuring;
                                        break;
                                    case "ID":
                                        header = element;
                                        break;
                                    case "Kostenplaats_x003_kostenplaatsv":
                                        header = strings.Kostenplaats;
                                        break;
                                    case "OpmerkingBudgetBeheerder":
                                        header = strings.Opmerking;
                                        break;
                                    case "LinkTitleNoMenu":
                                        header = strings.LinkTitle;
                                        break;
                                    case "LinkTitle":
                                        header = strings.LinkTitle;
                                        break;
                                    case "Terugkoppeling_x0020_CO":
                                        header = strings.TerugkoppelingCO;
                                        break;
                                    case "StartDate":
                                        header = strings.StartDate;
                                        break;
                                    case "Status_x0020_WPLU":
                                        header = strings.StatusWPLU;
                                        break;
                                    default:
                                        if (element.split('_x0020_').length > 1) {
                                            header = element.split('_x0020_').join(" ");
                                        }
                                        else if (element.split('_').length > 1) {
                                            header = element.split('_').join(" ");
                                        }
                                        else {
                                            header = element;
                                        }
                                        break;
                                }
                                var ch = {
                                    Title: header,
                                };
                                _this.listViewHeaders.push(ch);
                                columnsHeaderFields.push(element);
                            });
                        }
                        _viewXML = "";
                        return [4 /*yield*/, web.lists.getById(listId).views.getById(viewId).get()];
                    case 3:
                        xx = _a.sent();
                        _viewXML = xx.ListViewXml;
                        _viewXML = _viewXML.replace('<FieldRef Name="LinkTitleNoMenu" />', '<FieldRef Name="LinkTitle" />');
                        return [4 /*yield*/, web.lists.getById(listId).getItemsByCAMLQuery({ ViewXml: _viewXML })];
                    case 4:
                        items = _a.sent();
                        items.forEach(function (element) {
                            var itmValue = new Array();
                            columnsHeaderFields.forEach(function (x) {
                                switch (x) {
                                    case "AssignedTo":
                                        //this is to get user info based on user id
                                        var result = userColl.filter(function (IUser) { return IUser.Id == element['AssignedToId']; });
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
                                        if (element['WorkflowLink']) {
                                            itmValue.push({ key: x, value: element['WorkflowLink']['Description'], url: element['WorkflowLink']['Url'] });
                                        }
                                        else {
                                            itmValue.push({ key: x, value: "", url: "" });
                                        }
                                        break;
                                    case "LinkTitle":
                                        var actionUrl = "";
                                        if (element["ContentTypeId"] ? element["ContentTypeId"].indexOf(_this.meetingMinutestypeId) > -1 : false) {
                                            actionUrl = _this.BaseUrl() + '/Lists/Tasks/EditForm.aspx?ID=' + element['Id'] + '&Source=' + _this.BaseUrl() + '/SitePages/Dashboard.aspx';
                                        }
                                        else {
                                            actionUrl = _this.BaseUrl() + '/_layouts/15/WrkTaskIP.aspx?List=' + _this.selectedList() + '&ID=' + element['Id'] + '&Source=' + _this.BaseUrl() + '/SitePages/Dashboard.aspx' + '&ContentTypeId=' + element['ContentTypeId'];
                                        }
                                        itmValue.push({ key: strings.LinkTitle, value: element['Title'], url: actionUrl });
                                        break;
                                    case "LinkTitleNoMenu":
                                        if (element["ContentTypeId"] ? element["ContentTypeId"].indexOf(_this.meetingMinutestypeId) > -1 : false) {
                                            actionUrl = _this.BaseUrl() + '/Lists/Tasks/EditForm.aspx?ID=' + element['Id'] + '&Source=' + _this.BaseUrl() + '/SitePages/Dashboard.aspx';
                                        }
                                        else {
                                            actionUrl = _this.BaseUrl() + '/_layouts/15/WrkTaskIP.aspx?List=' + _this.selectedList() + '&ID=' + element['Id'] + '&Source=' + _this.BaseUrl() + '/SitePages/Dashboard.aspx' + '&ContentTypeId=' + element['ContentTypeId'];
                                        }
                                        itmValue.push({ key: strings.LinkTitle, value: element['Title'], url: actionUrl });
                                        break;
                                    case "DueDate":
                                        if (element['DueDate']) {
                                            itmValue.push({ key: strings.DueDate, value: moment(new Date(element['DueDate'])).format("DD/MM/YYYY"), url: null });
                                        }
                                        else {
                                            itmValue.push({ key: strings.DueDate, value: "", url: null });
                                        }
                                        break;
                                    case "StartDate":
                                        if (element['StartDate']) {
                                            itmValue.push({ key: strings.StartDate, value: moment(new Date(element['StartDate'])).format("DD/MM/YYYY"), url: null });
                                        }
                                        else {
                                            itmValue.push({ key: strings.StartDate, value: "", url: null });
                                        }
                                        break;
                                    default:
                                        itmValue.push({ key: x, value: element[x], url: null });
                                        break;
                                }
                            });
                            _this.taskListItems.push(itmValue);
                        });
                        tableId = '#tbl' + this.tblViewName() + this.listViewID();
                        $(tableId).DataTable({
                            responsive: true,
                            "lengthMenu": [[5, 10, 25], [5, 10, 25]],
                            initComplete: function () {
                            }
                        });
                        $('#span' + this.tblViewName()).css('display', 'none');
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * LoadAllSiteUsers
     */
    TaskListViewModel.prototype.LoadAllSiteUsers = function () {
        var _this = this;
        var web = new sp_pnp_js_1.Web(this.BaseUrl());
        web.siteUsers.get().then(function (u) {
            var siteUsersCollection = new Array();
            u.forEach(function (el) {
                var iSiteUser = new ISiteUser();
                iSiteUser.Id = el.Id;
                iSiteUser.UserTitle = el.Title;
                siteUsersCollection.push(iSiteUser);
            });
            return siteUsersCollection;
        }).then(function (_userColl) {
            var ListID = _this.selectedList();
            var ViewID = _this.selectedView();
            // get view Name
            web.lists.getById(ListID).views.getById(ViewID).get().then(function (v) {
                _this.viewName(v.Title);
                _this.tblViewName(v.Title.replace(/ /g, ''));
                _this.listViewID(ViewID);
            });
            //
            _this.DisplayListView(web, ListID, ViewID, _userColl);
        });
    };
    return TaskListViewModel;
}());
exports.default = TaskListViewModel;

//# sourceMappingURL=TaskListViewModel.js.map
