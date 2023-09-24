import * as React from 'react';
import styles from './ReportingPortal.module.scss';
import { IReportingPortalProps } from './IReportingPortalProps';
import { IReportingPortalState } from './IReportingPortalState';
import { padZeroesLeft, SortArrayByPeriod, includes } from '../../../utils/helper';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Row, Container, Col } from 'react-bootstrap';
import 'bootstrap/dist/css/bootstrap.min.css';
import Select from 'react-select';
import Form from 'react-bootstrap/Form';
import Button from 'react-bootstrap/Button';
import {
  Modal, ModalBody, ModalFooter, ModalHeader
} from 'reactstrap';
import DataTable from 'react-data-table-component';
import ITableHeader from './ITableHeader';
import { TeachingBubble } from 'office-ui-fabric-react/lib/TeachingBubble';
import "@pnp/polyfill-ie11";
import { sp } from '@pnp/sp';
import '@pnp/sp/items';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import { WebAPIHelper, APISource } from '@common/WebAPIHelper';
import { EventType, Logging } from '@common/logging/Logging';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Environment } from '../../../common/Environment';
import { saveAs } from 'file-saver';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Modal as FabricModal } from 'office-ui-fabric-react';
import browser from 'browser-detect';
import { FontIcon } from '@fluentui/react/lib/Icon';
import * as moment from 'moment';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import AsyncSelect from 'react-select/async';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { AppInsightsContext } from '@microsoft/applicationinsights-react-js';


const tableStyle = {
  header: {
    style: {
      paddingTop: '0px !important',
      paddingBottom: '0px !important',
      minHeight: '40px !important',

    }
  },
  subHeader: {
    style: {
      minHeight: '40px !important',
      padding: '0px 0px 0px 0px !important'
    },
  },
  headRow: {
    style: {
      minHeight: '25px !important',
    },
  },
  rows: {
    style: {
      fontSize: '12px',
    }
  },
  pagination: {
    style: {
      'div > select': { minWidth: '30px !important' }

    }
  },

  headCells: {
    style: {
      paddingRight: '1% !important',
    }
  },
  cells: {
    style: {
      paddingRight: '1% !important',
    },
  },
};
const ddlStyle = {
  control: (provided) => ({
    ...provided,
    height: '35px',
    borderColor: 'black'
  }),
  container: (provided) => ({
    ...provided,
    height: '35px',
  }),
  indicatorContainer: (provided) => ({
    ...provided,
    paddingTop: '0px',
    paddingBottom: '0px',
    height: '35px'
  }),
};

const reportLevelValueDDlStyle = {
  control: (provided) => ({
    ...provided,
    borderColor: 'black'
  }),
  indicatorContainer: (provided) => ({
    ...provided,
    paddingLeft: '0px !important'
  }),
  valueContainer: (provided) => ({
    ...provided,
    padding: '0px !important',
    fontSize: '9.5px'
  })
};

const indexValueDDlStyle = {
  control: (provided) => ({
    ...provided,
    borderColor: 'black'
  }),
  indicatorContainer: (provided) => ({
    ...provided,
    paddingLeft: '0px !important'
  }),
  valueContainer: (provided) => ({
    ...provided,
    padding: '0px !important',
    fontSize: '12px'
  })
};


const browserDetails = browser();

let intervalDialogue = setInterval(() => {
  if (browserDetails && browserDetails.name.indexOf("ie") > -1) {
    var searchBoxes = document.querySelectorAll("input[type=search]");
    for (var j = 0; j < searchBoxes.length; j++) {
      if (searchBoxes[j].className.indexOf("_3X3KIHRvQlB_k1KQr3703K Its0LEoVfS3AbQY9MWygN") > -1 && searchBoxes[j].className.indexOf("pageTopSearchBox") < 0) {
        searchBoxes[j].classList.add("pageTopSearchBox");
      }
    }
  }
}, 1000);

export default class ReportingPortal extends React.Component<IReportingPortalProps, IReportingPortalState> {
  private webAPI: WebAPIHelper;
  private EVENTNAME: string;
  private indexQueryArray: any = [];
  private parsedIndexQuery: any = [];

  public constructor(props: IReportingPortalProps) {
    super(props);

    this.state = {
      ReportCategories: [],
      ReportCategory: "",
      ReportTypes: [],
      ReportType: "",
      ReportIndexes: [
        { value: '', label: 'Choose an Index' },
      ],
      IndexValue: "",
      CurrentReportIndex: {
        value: '', label: 'Choose an Index',
      },
      ReportData: [],
      FilteredReportData: [],
      columns: [],
      HCMColumns: [],
      GLReportColumns: [],
      IndexQuery: "",
      ParsedIndexQuery: "",
      ConditionCount: 0,
      ShowBubble: false,
      BubbleTarget: "",
      BubbleMessage: "",
      Condition: "",
      addedCondition: "",
      manageReportViewModal: false,
      CurrentReportEmbedURL: "",
      UserPermissions: [],
      UserRoles: [],
      Countries: [],
      Regions: [],
      Areas: [],
      Branches: [],
      Departments: [],
      ReportLevels: [],
      SelectedCountry: { value: "-1", label: "Select" },
      SelectedReportLevel: "",
      SelectedReportLevelValue: { value: "All", label: "All" },
      ReportLevelValues: [],
      IsAccessAvailable: true,
      IsDataLoading: false,
      IsReportLoading: false,
      SelectedReports: [],
      confirmReportDownloadModal: false,
      IsDownloadAll: false,
      manageHelpDocumentViewModal: false,
      manageLoadingModal: false,
      ScreenLoadingMsg: "Please wait...",
      IsConnectionFailed: false,
      ShowParentOnly: false,
      CurrentReportFile: "",
      CurrentHelpDocumentBlob: "",
      managMessageModal: false,
      indexDateValue: false,
      startDateValue: null,
      endDateValue: null,
      indexNameValue: false,
      ReportLevelIndexName: [],
      SelectedReportLevelIndexName: { value: "-1", label: "Search with 3 char or more" },
      value: '',
      indexValues: [],
      fetchedindexValuesArray: [],
      lastSearchIndexName: "",
      indexQueryArray: [],
      allowAddIndex: true,
      NoOfIndexValueChar: 0,
      invoiceSTDColumns: [],
      invoiceCONColumns: [],
      indexQueryDateFormat: "",
      parseQueryDateFormat: "",
      isDeptChanged: true,
      ReportIndexValues: [],
      previousSelectedDept: [],
      selectedDept: [],
      previousSelectedIndexValue: { Name: "", Dept: "" },
      reportsPermission: "",
      networkID: "",
      Sensitive: "",
      Selector: [],
      searchInput: null,
      isGetInvoiceFetchRecord: true,
      totalCountDisplayMsg: "",
      totalRecordCount: "",
      totalRecordCountNumber: 0,
      invoiceNumberLimitation: 0,
      isAllRowSelect: true,
      isItemSelect: false,
      ReportTypesValues: [],
      downloadFilesCount: 0,
      downloadBatchSize: 0
    };

    sp.setup({
      sp: {
        headers: {
          "Accept": "application/json; odata=verbose",
          "ContentType": "application/json; odata=verbose",
          "User-Agent": "NONISV|Securitas|NAReportingPortalWebPart/1.0",
          "X-ClientService-ClientTag": "NONISV|Securitas|NAReportingPortalWebPart/1.0"
        }
      },
      // set ie 11 mode
      ie11: true,
      spfxContext: props.context,
    });
    this.webAPI = new WebAPIHelper(props.context);
    Environment.initialize();
    this.EVENTNAME = "ReportingPortal";
    this.getConfigValues();
  }

  public componentDidMount() {
    Logging.AppInsightsTrackEvent(this.EVENTNAME, EventType.Information, "componentDidMount", "page load triggered ");
    this.HandleContainerUI();
    this.GetUserPermissions();
    window.addEventListener('resize', this.handleWindowResize);

    if (document.getElementById("spSiteHeader") != null) {
      if (document.getElementById("spSiteHeader").querySelectorAll("span[role='heading']")[0].getElementsByTagName('a').length > 0) {
        document.getElementById("spSiteHeader").querySelectorAll("span[role='heading']")[0].getElementsByTagName('a')[0].innerHTML += " (" + this.props.context.pageContext.user.displayName + ")";
      }
    }
  }
  public componentDidUpdate() {
    this.HandleContainerUI();
  }

  private getConfigListData() {
    return this.props.context.spHttpClient.get(this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('ReportingPortalConfig')/Items", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private getColumnListData() {
    return this.props.context.spHttpClient.get(this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('ReportCategoryColumn')/Items?$filter=IsActive eq 1", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }


  private getConfigValues(): void {
    var noOfIndexValueChar: number = 0;
    var indexQueryDateFormat: string = "";
    var parseQueryDateFormat: string = "";
    var resultAll: any = [];
    var ReportCategoryCOlumns: any = [];
    var ReportInvoiceSTDColumns: any = [];
    var ReportInvoiceCONColumns: any = [];
    var ReportHCMColumns: any = [];
    var selector: any = [];
    var totalCountDisplayMsg: string = "";
    var invoiceNumberLimitation: number = 0;
    var downloadFilesCount: number = 0;
    var downloadBatchSize: number = 0;

    this.getConfigListData()
      .then((response) => {
        var result = response.value;
        result.map((res) => {
          switch (res.Title) {
            case 'NoOfIndexValueChar':
              noOfIndexValueChar = Number(res.Value);
              break;
            case 'IndexQueryDateFormat':
              indexQueryDateFormat = res.Value;
              break;
            case 'ParseQueryDateFormat':
              parseQueryDateFormat = res.Value;
              break;
            case 'TotalCountDisplayMSG':
              totalCountDisplayMsg = res.Value1;
              break;
            case 'InvoiceNumberLimitation':
              invoiceNumberLimitation = res.Value;
              break;
            case 'DownloadFilesCount':
              downloadFilesCount = parseInt(res.Value, 10);
              break;
            case 'DownloadBatchSize':
              downloadBatchSize = parseInt(res.Value, 10);
              break;
          }
        });
        this.setState({
          NoOfIndexValueChar: noOfIndexValueChar,
          indexQueryDateFormat: indexQueryDateFormat,
          parseQueryDateFormat: parseQueryDateFormat,
          totalCountDisplayMsg: totalCountDisplayMsg,
          invoiceNumberLimitation: invoiceNumberLimitation,
          downloadFilesCount: downloadFilesCount,
          downloadBatchSize: downloadBatchSize
        });
      });

    this.getColumnListData()
      .then((response) => {
        var result: any = response.value;
        resultAll = (result.sort((a, b) => a.Order0 < b.Order0 ? -1 : a.Order0 > b.Order0 ? 1 : 0));
        var column: any = [];

        if (resultAll) {
          resultAll.map((res) => {
            switch (res.columnType) {
              case 'Name':
                selector.push({ name: res.name, value: res.selector });
                column = {
                  name: res.name,
                  selector: res.selector,
                  sortable: res.sortable,
                  wrap: res.wrap,
                  width: res.width,
                  compact: res.compact,
                  cell: (row) => <span>{row[res.selector] ? row[res.selector] : "--"}</span>

                };

                switch (res.ReportCategory.length) {
                  case 3:
                    ReportCategoryCOlumns.push(column);
                    ReportInvoiceSTDColumns.push(column);
                    ReportInvoiceCONColumns.push(column);
                    ReportHCMColumns.push(column);
                    break;
                  case 2:
                    if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "BillingInvoice" || res.ReportCategory[1] == "BillingInvoice")) {
                      ReportCategoryCOlumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportCategoryCOlumns.push(column);
                      ReportHCMColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "BillingInvoice " || res.ReportCategory[1] == "BillingInvoice") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportHCMColumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    break;
                  case 1:
                    if (res.ReportCategory[0] == "BillingInvoice") {
                      if (res.ReportType && res.ReportType.length > 0) {

                        switch (res.ReportType.length) {
                          case 2:
                            ReportInvoiceSTDColumns.push(column);
                            ReportInvoiceCONColumns.push(column);
                            break;
                          case 1:
                            if (res.ReportType[0] == "StandardInvoice") {
                              ReportInvoiceSTDColumns.push(column);
                            } else {
                              ReportInvoiceCONColumns.push(column);
                            }
                        }
                      }
                    }
                    else if (res.ReportCategory[0] == "HCMReports") {
                      ReportHCMColumns.push(column);
                    }
                    else {
                      ReportCategoryCOlumns.push(column);
                    }
                    break;
                }
                break;
              case 'String':
                var startPath = (res.selector).split('-')[0];
                var endPath = (res.selector).split('-')[1];
                selector.push({ name: res.name, value: startPath + "-" + endPath });
                column = {
                  name: res.name,
                  selector: startPath,
                  sortable: res.sortable,
                  wrap: res.wrap,
                  compact: res.compact,
                  cell: (row) => <span>{(row[startPath] || row[endPath]) ? row[startPath] + (row[startPath] && row[endPath] ? "-" : "") + row[endPath] : "--"}</span>

                };

                switch (res.ReportCategory.length) {
                  case 3:
                    ReportCategoryCOlumns.push(column);
                    ReportInvoiceSTDColumns.push(column);
                    ReportInvoiceCONColumns.push(column);
                    ReportHCMColumns.push(column);
                    break;
                  case 2:
                    if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "BillingInvoice" || res.ReportCategory[1] == "BillingInvoice")) {
                      ReportCategoryCOlumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportCategoryCOlumns.push(column);
                      ReportHCMColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "BillingInvoice " || res.ReportCategory[1] == "BillingInvoice") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportHCMColumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    break;
                  case 1:
                    if (res.ReportCategory[0] == "BillingInvoice") {
                      if (res.ReportType && res.ReportType.length > 0) {

                        switch (res.ReportType.length) {
                          case 2:
                            ReportInvoiceSTDColumns.push(column);
                            ReportInvoiceCONColumns.push(column);
                            break;
                          case 1:
                            if (res.ReportType[0] == "StandardInvoice") {
                              ReportInvoiceSTDColumns.push(column);
                            } else {
                              ReportInvoiceCONColumns.push(column);
                            }
                        }
                      }
                    }
                    else if (res.ReportCategory[0] == "HCMReports") {
                      ReportHCMColumns.push(column);
                    }
                    else {
                      ReportCategoryCOlumns.push(column);
                    }
                    break;
                }
                break;
              case 'FilePath':
                selector.push({ name: res.name, value: res.selector });
                column = {
                  name: res.name,
                  selector: res.selector,
                  sortable: res.sortable,
                  width: res.width,
                  wrap: res.wrap,
                  compact: res.compact,
                  cell: (row) => <span>{row[res.selector] ? row[res.selector].split('/')[row[res.selector].split('/').length - 1] : "--"}</span>

                };
                switch (res.ReportCategory.length) {
                  case 3:
                    ReportCategoryCOlumns.push(column);
                    ReportInvoiceSTDColumns.push(column);
                    ReportInvoiceCONColumns.push(column);
                    ReportHCMColumns.push(column);
                    break;
                  case 2:
                    if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "BillingInvoice" || res.ReportCategory[1] == "BillingInvoice")) {
                      ReportCategoryCOlumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportCategoryCOlumns.push(column);
                      ReportHCMColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "BillingInvoice " || res.ReportCategory[1] == "BillingInvoice") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportHCMColumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    break;
                  case 1:
                    if (res.ReportCategory[0] == "BillingInvoice") {
                      if (res.ReportType && res.ReportType.length > 0) {

                        switch (res.ReportType.length) {
                          case 2:
                            ReportInvoiceSTDColumns.push(column);
                            ReportInvoiceCONColumns.push(column);
                            break;

                          case 1:
                            if (res.ReportType[0] == "StandardInvoice") {
                              ReportInvoiceSTDColumns.push(column);
                            } else {
                              ReportInvoiceCONColumns.push(column);
                            }
                        }
                      }
                    }
                    else if (res.ReportCategory[0] == "HCMReports") {
                      ReportHCMColumns.push(column);
                    }
                    else {
                      ReportCategoryCOlumns.push(column);
                    }
                    break;
                }
                break;
              case 'Date':
                selector.push({ name: res.name, value: res.selector });
                column = {
                  name: res.name,
                  selector: res.selector,
                  sortable: res.sortable,
                  width: res.width,
                  wrap: res.wrap,
                  compact: res.compact,
                  cell: (row) => <span>{row[res.selector] ? moment(new Date(row[res.selector])).format('MM-DD-YYYY') : "--"}</span>

                };
                switch (res.ReportCategory.length) {
                  case 3:
                    ReportCategoryCOlumns.push(column);
                    ReportInvoiceSTDColumns.push(column);
                    ReportInvoiceCONColumns.push(column);
                    ReportHCMColumns.push(column);
                    break;
                  case 2:
                    if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "BillingInvoice" || res.ReportCategory[1] == "BillingInvoice")) {
                      ReportCategoryCOlumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportCategoryCOlumns.push(column);
                      ReportHCMColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "BillingInvoice " || res.ReportCategory[1] == "BillingInvoice") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportHCMColumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    break;
                  case 1:
                    if (res.ReportCategory[0] == "BillingInvoice") {
                      if (res.ReportType && res.ReportType.length > 0) {

                        switch (res.ReportType.length) {
                          case 2:
                            ReportInvoiceSTDColumns.push(column);
                            ReportInvoiceCONColumns.push(column);
                            break;

                          case 1:
                            if (res.ReportType[0] == "StandardInvoice") {
                              ReportInvoiceSTDColumns.push(column);
                            } else {
                              ReportInvoiceCONColumns.push(column);
                            }
                        }
                      }
                    }
                    else if (res.ReportCategory[0] == "HCMReports") {
                      ReportHCMColumns.push(column);
                    }
                    else {
                      ReportCategoryCOlumns.push(column);
                    }
                    break;
                }
                break;
              case 'Period':
                var Year = (res.selector).split('-')[0];
                var Month = (res.selector).split('-')[1];
                selector.push({ name: res.name, value: Month + "-" + Year });
                column = {
                  name: res.name,
                  selector: Year,
                  sortable: res.sortable,
                  width: res.width,
                  compact: res.compact,
                  center: res.center,
                  cell: (row) => <span>{row[Year] && row[Month] ? row[Month] + "-" + row[Year] : "--"}</span>,
                  sortFunction: (a, b) => {
                    return ((new Date(parseInt(a[Year]), parseInt(a[Month]) - 1, 1) < new Date(parseInt(b[Year]), parseInt(b[Month]) - 1, 1)) ? -1 : ((new Date(parseInt(a[Year]), parseInt(a[Month]) - 1, 1) > new Date(parseInt(b[Year]), parseInt(b[Month]) - 1, 1)) ? 1 : 0));
                  }
                };
                switch (res.ReportCategory.length) {
                  case 3:
                    ReportCategoryCOlumns.push(column);
                    ReportInvoiceSTDColumns.push(column);
                    ReportInvoiceCONColumns.push(column);
                    ReportHCMColumns.push(column);
                    break;
                  case 2:
                    if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "BillingInvoice" || res.ReportCategory[1] == "BillingInvoice")) {
                      ReportCategoryCOlumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportCategoryCOlumns.push(column);
                      ReportHCMColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "BillingInvoice " || res.ReportCategory[1] == "BillingInvoice") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportHCMColumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    break;
                  case 1:
                    if (res.ReportCategory[0] == "BillingInvoice") {
                      if (res.ReportType && res.ReportType.length > 0) {

                        switch (res.ReportType.length) {
                          case 2:
                            ReportInvoiceSTDColumns.push(column);
                            ReportInvoiceCONColumns.push(column);
                            break;

                          case 1:
                            if (res.ReportType[0] == "StandardInvoice") {
                              ReportInvoiceSTDColumns.push(column);
                            } else {
                              ReportInvoiceCONColumns.push(column);
                            }
                        }
                      }
                    }
                    else if (res.ReportCategory[0] == "HCMReports") {
                      ReportHCMColumns.push(column);
                    }
                    else {
                      ReportCategoryCOlumns.push(column);
                    }
                    break;
                }
                break;
              case 'View':
                selector.push({ name: res.name, value: res.selector });
                column = {
                  name: <FontIcon iconName="RedEye" className={styles.cursorPointer + " " + styles.rowIcon} title="View" />,
                  selector: res.selector,
                  sortable: res.sortable,
                  width: res.width,
                  center: true,
                  compact: res.compact,
                  cell: (row) =>
                    <FontIcon iconName="RedEye" className={styles.cursorPointer + " " + styles.rowIcon} title="View" onClick={(e) => { this.OpenReportViewModal(row); }} />
                };
                switch (res.ReportCategory.length) {
                  case 3:
                    ReportCategoryCOlumns.push(column);
                    ReportInvoiceSTDColumns.push(column);
                    ReportInvoiceCONColumns.push(column);
                    ReportHCMColumns.push(column);
                    break;
                  case 2:
                    if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "BillingInvoice" || res.ReportCategory[1] == "BillingInvoice")) {
                      ReportCategoryCOlumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportCategoryCOlumns.push(column);
                      ReportHCMColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "BillingInvoice " || res.ReportCategory[1] == "BillingInvoice") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportHCMColumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    break;
                  case 1:
                    if (res.ReportCategory[0] == "BillingInvoice") {
                      if (res.ReportType && res.ReportType.length > 0) {

                        switch (res.ReportType.length) {
                          case 2:
                            ReportInvoiceSTDColumns.push(column);
                            ReportInvoiceCONColumns.push(column);
                            break;

                          case 1:
                            if (res.ReportType[0] == "StandardInvoice") {
                              ReportInvoiceSTDColumns.push(column);
                            } else {
                              ReportInvoiceCONColumns.push(column);
                            }
                        }
                      }
                    }
                    else if (res.ReportCategory[0] == "HCMReports") {
                      ReportHCMColumns.push(column);
                    }
                    else {
                      ReportCategoryCOlumns.push(column);
                    }
                    break;
                }
                break;

              case 'Download':
                selector.push({ name: res.name, value: res.selector });
                column = {
                  name: <FontIcon iconName="Download" className={styles.cursorPointer + " " + styles.rowIcon} title="Download" />,
                  selector: res.selector,
                  sortable: res.sortable,
                  width: res.width,
                  center: true,
                  compact: res.compact,
                  cell: (row) =>
                    <FontIcon iconName="Download" className={styles.cursorPointer + " " + styles.rowIcon} title="Download" onClick={(e) => { this.DownloadFile(row); }} />

                };
                switch (res.ReportCategory.length) {
                  case 3:
                    ReportCategoryCOlumns.push(column);
                    ReportInvoiceSTDColumns.push(column);
                    ReportInvoiceCONColumns.push(column);
                    ReportHCMColumns.push(column);
                    break;
                  case 2:
                    if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "BillingInvoice" || res.ReportCategory[1] == "BillingInvoice")) {
                      ReportCategoryCOlumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "GLReport" || res.ReportCategory[1] == "GLReport") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportCategoryCOlumns.push(column);
                      ReportHCMColumns.push(column);
                    }
                    else if ((res.ReportCategory[0] == "BillingInvoice " || res.ReportCategory[1] == "BillingInvoice") && (res.ReportCategory[0] == "HCMReports" || res.ReportCategory[1] == "HCMReports")) {
                      ReportHCMColumns.push(column);
                      ReportInvoiceSTDColumns.push(column);
                      ReportInvoiceCONColumns.push(column);
                    }
                    break;
                  case 1:
                    if (res.ReportCategory[0] == "BillingInvoice") {
                      if (res.ReportType && res.ReportType.length > 0) {

                        switch (res.ReportType.length) {
                          case 2:
                            ReportInvoiceSTDColumns.push(column);
                            ReportInvoiceCONColumns.push(column);
                            break;

                          case 1:
                            if (res.ReportType[0] == "StandardInvoice") {
                              ReportInvoiceSTDColumns.push(column);
                            } else {
                              ReportInvoiceCONColumns.push(column);
                            }
                        }
                      }
                    }
                    else if (res.ReportCategory[0] == "HCMReports") {
                      ReportHCMColumns.push(column);
                    }
                    else {
                      ReportCategoryCOlumns.push(column);
                    }
                    break;
                }
                break;
            }
          });
        }

        this.setState({
          GLReportColumns: ReportCategoryCOlumns,
          invoiceSTDColumns: ReportInvoiceSTDColumns,
          invoiceCONColumns: ReportInvoiceCONColumns,
          HCMColumns: ReportHCMColumns,
          Selector: selector
        });
      });
  }

  public HandleContainerUI = () => {
    if (document.getElementsByClassName("CanvasZone") && document.getElementsByClassName("CanvasZone").length > 0) {
      for (var i = 0; i < document.getElementsByClassName("CanvasZone").length; i++) {
        document.getElementsByClassName("CanvasZone")[i].setAttribute("style", "max-width:100vw !important");
      }
    }

    if (document.getElementById("O365_MainLink_TenantLogo") != null) {
      // document.getElementById("O365_MainLink_TenantLogo").classList.remove('hover');
      document.getElementById("O365_MainLink_TenantLogo").setAttribute("style", "pointer-events:none !important");
    }

  }

  public GetReportTypes = async () => {
    try {
      var fileStart: any[] = this.state.UserRoles;
      fileStart = fileStart.filter((v, i) => {
         return fileStart.map((val) => val.ReportType).indexOf(v.ReportType) == i ;
        });
      let reportTypes: any[] = [];
      let rptTypes: any[] = [];
      for (let i = 0; i < fileStart.length; i++) {
        rptTypes.push(await sp.web.lists.getByTitle("ReportTypeMaster").items.select("Country", "Title", "ReportCategory", "FileNameStarts", "UserRole").filter("IsActive eq 'Yes' and FileNameStarts eq '" + fileStart[i].ReportType.trim() + "'").get());
      }

      rptTypes.map((reportType) => {
        if (reportType.length > 0) {
          (reportType.map((rptType) => {
            var reports = { Country: rptType.Country.results.length == 2 ? rptType.Country.results[0] + "," + rptType.Country.results[1] : rptType.Country.results[0], Title: rptType.Title, ReportCategory: rptType.ReportCategory, FileNameStarts: rptType.FileNameStarts, UserRole: rptType.UserRole };
            reportTypes.push(reports);
          }
          ));
        }
      });

      this.setState({ ReportTypesValues: reportTypes }, () => {
        this.toggleLoadingModal();
        if (this.state.Countries.length == 1) {
          this.GetReportCategories();
        }
      });
    }
    catch (ex) {
      this.HandleException("GetReportTypes", ex);
    }
  }


  public GetReportCategories = async () => {
    try {
      var uniqueRptCategories = [];
      var distinctRptCategories = [];
      var reportCategory = { value: '-1', label: 'Select', userRole: '' };
      let reportCategories: any[] = this.state.ReportTypesValues.filter(x => x.Country.includes(this.state.SelectedCountry.value));
      var rptCategories = this.state.ReportCategories;

      var reportPermission = (this.state.reportsPermission).split(',');

      rptCategories = rptCategories.concat(reportCategories.map((rptCategory) => { return { value: rptCategory.ReportCategory, label: rptCategory.ReportCategory, userRole: rptCategory.UserRole }; }));

      for (let i = 0; i < rptCategories.length; i++) {
        if (!uniqueRptCategories[rptCategories[i].value]) {
          var reportCat = (rptCategories[i].userRole);
          var uniqueRptPerm = [];
          reportPermission.forEach(reportPerm => {
            if (!uniqueRptPerm[reportPerm.trim()]) {
              if (reportCat == reportPerm.trim()) {
                distinctRptCategories.push(rptCategories[i]);
                uniqueRptPerm[reportPerm.trim()] = 1;
                uniqueRptCategories[rptCategories[i].value] = 1;

              }
            }
          });

        }
      }

      distinctRptCategories = (distinctRptCategories.map((rptCategory) => { return { value: rptCategory.value, label: rptCategory.label }; }));

      this.setState({
        ReportCategories: distinctRptCategories,
        ReportCategory: reportCategory
      }, function () {
        if (distinctRptCategories.length == 1) {
          this.LoadReportTypeDDLByCategory(reportCategory);
        }
      });
    }
    catch (ex) {
      this.HandleException("GetReportCategories", ex);
    }

  }
  public GetUserPermissions = () => {
    try {
      this.toggleLoadingModal();
      this.webAPI.GetWebAPI(APISource.WebAPI, "AuthUser")
        .then((response: any): void => {
          Logging.AppInsightsTrackEvent(this.EVENTNAME, EventType.Information, "GetUserPermissions", "Auth Response count:" + response.length);
          var result = response;

          if (result && result.Table2 && result.Table2.length > 0 && result.Table[0] && result.Table1.length > 0) {
            this.setState({
              UserPermissions: result.Table2,
              reportsPermission: result.Table[0].ReportCat,
              networkID: result.Table[0].NetworkId,
              Sensitive: result.Table[0].Sensitive,
              UserRoles: result.Table1,
              IsDataLoading: false,
              IsAccessAvailable: true,
            }, function () {
              this.BindPermissionsData();
              this.handleWindowResize();
              this.GetReportTypes();
            });
          }
          else {
            this.setState({
              IsDataLoading: false,
              IsAccessAvailable: false,
            }, function () {
              this.toggleLoadingModal();
            });
          }
        })
        .catch((error: any) => {
          // console.error(error);
          Logging.AppInsightsTrackEvent(this.EVENTNAME, EventType.Exception, "GetUserPermissions", "Auth Response failed:" + error);
          this.setState({
            IsDataLoading: true,
            IsConnectionFailed: true
          }, function () {
            this.toggleLoadingModal();
          });
        });
    }
    catch (ex) {
      this.HandleException("GetReportCategories", ex);
    }
  }


  public LoadIndexesByReportType = async (selectedReportType) => {

    var distinctIndexValues: string = "";
    var distinctIndexes = [];
    let indexes: any[] = await sp.web.lists.getByTitle("ReportTypeMaster").items.select("Title", "SearchIndexes").filter("FileNameStarts eq '" + selectedReportType.value + "' and IsActive eq 'Yes'").get();
    var rptIndex: any[] = [{ value: '', label: 'Choose an Index' }];
    var column: any = [];
    for (let i = 0; i < indexes.length; i++) {

      distinctIndexValues = (indexes[i].SearchIndexes);
    }

    distinctIndexes = (distinctIndexValues).split(";");

    for (let i = 0; i < distinctIndexes.length; i++) {
      var indexValue = distinctIndexes[i].split(":");
      rptIndex = rptIndex.concat({ value: indexValue[1], label: indexValue[0] });
    }

    if (this.state.ReportCategory.label == "Billing-Invoice") {
      column = (selectedReportType.label == "Standard Invoice") ? this.state.invoiceSTDColumns : this.state.invoiceCONColumns;
    } else {
      column = this.state.columns;
    }

    this.setState({
      ReportType: selectedReportType,
      ReportIndexes: rptIndex,
      SelectedReportLevel: { value: "All", label: "All" },
      SelectedReportLevelValue: { value: "All", label: "All" },
      FilteredReportData: [],
      ConditionCount: 0,
      ParsedIndexQuery: "",
      IndexQuery: "",
      IndexValue: "",
      ReportLevelValues: [],
      ShowParentOnly: false,
      columns: column,
      searchInput: "",
      totalRecordCountNumber: 0
    });
    this.indexQueryArray = [];
    this.parsedIndexQuery = [];
  }




  private HandleReportTypeChanged = (selectedOption: any, actionObj: any) => {
    if (actionObj.action !== "clear") {
      var _selectedOption = selectedOption;
      this.LoadIndexesByReportType(selectedOption);
      this.setState({
        ReportType: selectedOption,
        ReportIndexes: [
          { value: '', label: 'Choose an Index' }],
        CurrentReportIndex: {
          value: '', label: 'Choose an Index'
        },
        SelectedReportLevel: { value: "All", label: "All" },
        SelectedReportLevelValue: { value: "All", label: "All" },
        FilteredReportData: [],
        ConditionCount: 0,
        ParsedIndexQuery: "",
        IndexQuery: "",
        IndexValue: "",
        ReportLevelValues: [],
        ShowParentOnly: false,
        isDeptChanged: true,
        searchInput: "",
        totalRecordCountNumber: 0
      });
      this.indexQueryArray = [];
      this.parsedIndexQuery = [];
    }
    else {
      this.setState({
        ReportType: { value: 'All', label: 'All' },
        ReportIndexes: [
          { value: '', label: 'Choose an Index' }],
        CurrentReportIndex: {
          value: '', label: 'Choose an Index'
        },
        SelectedReportLevel: { value: "All", label: "All" },
        SelectedReportLevelValue: { value: "All", label: "All" },
        FilteredReportData: [],
        ConditionCount: 0,
        ParsedIndexQuery: "",
        IndexQuery: "",
        IndexValue: "",
        ReportLevelValues: [],
        ShowParentOnly: false,
        isDeptChanged: true,
        totalRecordCountNumber: 0
      });
      this.indexQueryArray = [];
      this.parsedIndexQuery = [];
    }
  }

  private HandleReportCategoryChanged = (selectedOption: any, actionObj: any) => {
    if (actionObj.action !== "clear") {
      var column: any = [];
      if (selectedOption.label == "GL Reports") {
        column = this.state.GLReportColumns;

      } else if (selectedOption.label == "Billing-Invoice") {
        column = this.state.invoiceCONColumns;
      }
      else if (selectedOption.label == "HCM Reports") {
        column = this.state.HCMColumns;
      }
      this.setState({
        ReportCategory: selectedOption,
        ReportType: { value: "All", label: "All" },
        ReportTypes: [],
        ReportIndexes: [
          { value: '', label: 'Choose an Index' }],
        CurrentReportIndex: {
          value: '', label: 'Choose an Index'
        },
        ReportLevelIndexName: [],
        SelectedReportLevelIndexName: { value: "-1", label: "Search with 3 char or more" },
        indexDateValue: false,
        indexNameValue: false,
        IndexValue: "",
        FilteredReportData: [],
        ParsedIndexQuery: "",
        IndexQuery: "",
        ConditionCount: 0,
        ShowParentOnly: false,
        isDeptChanged: true,
        columns: column,
        searchInput: "",
        totalRecordCountNumber: 0
      });
      this.indexQueryArray = [];
      this.parsedIndexQuery = [];
      if (selectedOption.value != "-1") {
        this.setState({
          ReportCategory: selectedOption,
        }, function () {
          this.LoadReportTypeDDLByCategory(selectedOption);
          this.LoadDropdownByCountry();
          this.LoadIndexesByReportType({ value: 'All', label: 'All' });
        });
      }
      else {
        this.setState({
          ReportCategory: selectedOption,
          ReportType: { value: "All", label: "All" },
          ReportTypes: [],
          ReportIndexes: [
            { value: '', label: 'Choose an Index' }],
          CurrentReportIndex: {
            value: '', label: 'Choose an Index'
          },
          ReportLevelIndexName: [],
          SelectedReportLevelIndexName: { value: "-1", label: "Search with 3 char or more" },
          indexDateValue: false,
          indexNameValue: false,
          IndexValue: "",
          FilteredReportData: [],
          ParsedIndexQuery: "",
          IndexQuery: "",
          ConditionCount: 0,
          ShowParentOnly: false,
          isDeptChanged: true,
          searchInput: "",
          totalRecordCountNumber: 0
        });
        this.indexQueryArray = [];
        this.parsedIndexQuery = [];
      }
    }
    else {
      this.setState({
        ReportCategory: selectedOption,
        ReportType: { value: "All", label: "All" },
        ReportTypes: [],
        ReportIndexes: [
          { value: '', label: 'Choose an Index' }],
        CurrentReportIndex: {
          value: '', label: 'Choose an Index'
        },
        ReportLevelIndexName: [],
        SelectedReportLevelIndexName: { value: "-1", label: "Search with 3 char or more" },
        indexDateValue: false,
        indexNameValue: false,
        IndexValue: "",
        FilteredReportData: [],
        ParsedIndexQuery: "",
        IndexQuery: "",
        totalRecordCountNumber: 0
      });
    }
  }

  private HandleCountryChanged = (selectedOption: any, actionObj: any) => {
    if (actionObj.action !== "clear") {
      this.setState({
        SelectedCountry: selectedOption,
        ShowParentOnly: false,
        SelectedReportLevel: { value: "All", label: "All" },
        ReportLevels: [],
        SelectedReportLevelValue: { value: "All", label: "All" },
        ReportLevelValues: [],
        ReportType: { value: 'All', label: 'All' },
        FilteredReportData: [],
        ParsedIndexQuery: "",
        IndexQuery: "",
        ReportIndexes: [
          { value: '', label: 'Choose an Index' },
        ],
        CurrentReportIndex: {
          value: '', label: 'Choose an Index',
        },
        IndexValue: "",
        isDeptChanged: true,
        columns: [],
        searchInput: "",
        totalRecordCountNumber: 0,
        ReportTypes: []

      }, function () {
        this.GetReportCategories();

      });
      this.indexQueryArray = [];
      this.parsedIndexQuery = [];

    }
    else {
      this.setState({
        SelectedCountry: selectedOption,
        ReportCategory: { value: "-1", label: "Select" },
        ReportCategories: [],
        FilteredReportData: [],
        totalRecordCountNumber: 0
      });
    }
  }

  private HandleReportLevelChanged = (selectedOption: any, actionObj: any) => {
    if (selectedOption.value != "All") {
      this.setState({
        SelectedReportLevel: selectedOption,
        SelectedReportLevelValue: { value: "All", label: "All" },
        FilteredReportData: [],
        isDeptChanged: true,
        searchInput: "",
        totalRecordCountNumber: 0
      }, function () {
        this.LoadReportLevelValueDDL();
      });
    }
    else {
      this.setState({
        SelectedReportLevel: selectedOption,
        SelectedReportLevelValue: { value: "All", label: "All" },
        ReportLevelValues: [],
        FilteredReportData: [],
        ShowParentOnly: false,
        isDeptChanged: true,
        totalRecordCountNumber: 0
      });
    }
  }

  public HandleReportLevelValueChanged = (selectedOptions: any, actionObj: any) => {
    var selectedDept: any[];
    if (actionObj.action == "clear" || !selectedOptions || selectedOptions.length == 0) {
      this.setState({
        SelectedReportLevelValue: { value: "All", label: "All" },
        FilteredReportData: [],
        selectedDept: [],
        isDeptChanged: true,
        searchInput: "",
        totalRecordCountNumber: 0
      });
    }
    else {
      if (actionObj.action == "select-option" && actionObj.option.value == "All") {
        this.setState({
          SelectedReportLevelValue: { value: "All", label: "All" },
          FilteredReportData: [],
          selectedDept: [],
          isDeptChanged: true,
          searchInput: "",
          totalRecordCountNumber: 0
        });
      }
      else {
        if (selectedOptions.length > 1 && selectedOptions.filter(x => x.value == "All").length > 0) {
          selectedOptions = selectedOptions.filter(x => x.value != "All");
        }
        this.setState({
          SelectedReportLevelValue: selectedOptions,
          FilteredReportData: [],
          selectedDept: selectedDept,
          isDeptChanged: true,
          searchInput: "",
          totalRecordCountNumber: 0
        });
      }
    }

  }


  public getInvoiceValuesAPI = (selectedOptions) => {
    try {

      var searchReportRequestModel: any = {};
      var selectedReportLevel = this.state.SelectedReportLevel.value;
      var selectedReportLevelValues = this.state.SelectedReportLevelValue;
      var selectedFolderPathString: string = "";
      searchReportRequestModel.SearchFor = this.state.CurrentReportIndex.label.trim(" ")[0];
      searchReportRequestModel.SearchString = selectedOptions;

      if (selectedReportLevel == "Dept") {
        if (Array.isArray(selectedReportLevelValues) && selectedReportLevelValues.filter(x => x.value != "All").length > 0) {
          var selectedFolderPaths = selectedReportLevelValues.map((item) => {
            return item.value;
          });
          selectedFolderPathString = selectedFolderPaths.join(",");
        }
      }
      searchReportRequestModel.Dept = selectedFolderPathString;

      searchReportRequestModel.NetworkId = this.state.networkID;

      return new Promise((resolve, reject) => {
        return this.webAPI.PostWebAPI(APISource.WebAPI, "Engagement", JSON.stringify(searchReportRequestModel))
          .then((response: any) => {
            Logging.AppInsightsTrackEvent(this.EVENTNAME, EventType.Information, "getInvoiceValuesAPI", "Engagement Response count:" + response.length);
            var result = response;
            console.log(result);
            if (result && result.length > 0) {
              resolve(result);

            }

          })
          .catch((error: any) => {
            Logging.AppInsightsTrackEvent(this.EVENTNAME, EventType.Exception, "GetUserPermissions", "Auth Response failed:" + error);
            this.setState({
              IsDataLoading: true,
              IsConnectionFailed: true
            }, () => reject(error));
          });
      });
    }
    catch (ex) {
      this.HandleException("GetReportCategories", ex);
    }
  }


  public onChangeIndexValueDropDown = (selectedOptions: any, actionObj: any) => {
    if (actionObj.action == "clear" || !selectedOptions || selectedOptions.length == 0) {
      this.setState({
        SelectedReportLevelIndexName: { value: "-1", label: "Search with 3 char or more" }
      });
    }
    else {
      if (selectedOptions.length > 1 && selectedOptions.filter(x => x.value == "-1").length > 0) {
        selectedOptions = selectedOptions.filter(x => x.value != "-1");
      }
      this.setState({
        SelectedReportLevelIndexName: selectedOptions
      });
    }

  }


  public filterInvoiceValuesDropdown = (inputValue, indexValues) => {
    return indexValues.filter(i =>
      i.label.toLowerCase().includes(inputValue.toLowerCase())
    );
  }


  public delayToFetchInvoices = (inputValue, callback) => {
    setTimeout(() => {
      this.fetchInvoiceValuesDropdown(inputValue, callback);
    }, 2000);

  }

  public fetchInvoiceValuesDropdown = (inputValue, callback) => {
    const selectedOptions = inputValue;
    const NoOfIndexValueChar = this.state.NoOfIndexValueChar;
    var ReportIndexValues: any = [];
    var indexValues: any = [];
    var isDeptChanged = this.state.isDeptChanged;
    var lastSearchIndexName = this.state.lastSearchIndexName;

    if ((lastSearchIndexName == selectedOptions && isDeptChanged == false) || (selectedOptions.length > NoOfIndexValueChar && selectedOptions.trim() !== "" && isDeptChanged == false && selectedOptions.includes(lastSearchIndexName))) {
      if (this.state.fetchedindexValuesArray) {
        setTimeout(() => {
          callback(this.filterInvoiceValuesDropdown(inputValue, this.state.fetchedindexValuesArray));
        }, 100);
      }
    }
    else if ((selectedOptions.length == NoOfIndexValueChar && selectedOptions.trim() !== "" && isDeptChanged) || (selectedOptions.length == NoOfIndexValueChar && selectedOptions.trim() !== "" && lastSearchIndexName != selectedOptions)) {
      this.setState({ isGetInvoiceFetchRecord: false });
      this.getInvoiceValuesAPI(selectedOptions).then((response) => {
        ReportIndexValues = response;
        indexValues = [{ value: "-1", label: "Search with 3 char or more" }].concat(ReportIndexValues.map((levelValue) => { return { value: levelValue.Number + "-" + levelValue.DeptNum, label: levelValue.Number + "-" + levelValue.Name }; }));

        this.setState({
          fetchedindexValuesArray: indexValues,
          ReportLevelIndexName: indexValues,
          lastSearchIndexName: inputValue,
          ReportIndexValues: ReportIndexValues,
          isDeptChanged: false,
          isGetInvoiceFetchRecord: true
        }, () => {
          setTimeout(() => {
            callback(this.filterInvoiceValuesDropdown(inputValue, indexValues));
          }, 100);

        });


      })
        .catch((err) => {
          console.log(err);
        });

    }
  }


  public compareArray = (arr1, arr2) => {
    let finalArray: any = [];
    arr1.forEach(e1 => arr2.forEach(e2 => {
      if (e1.name === e2.name) {
        finalArray.push(e1);
      }
    }));
    return finalArray;
  }


  public HandleTextChanged = (changedValue, newValue) => {
    var filteredRptData = this.state.ReportData;
    var filteredData: any = [];
    if (changedValue) {

      var splittedText: any = (changedValue.trim()).split(" ");
      var selector = this.compareArray(this.state.Selector, this.state.columns);
      for (let text of splittedText) {
        selector.forEach(sel => {
          filteredRptData.forEach(r => {

            var item = sel.value;
            if (item.includes("-")) {
              var startPath = item.split("-")[0];
              var endPath = item.split("-")[1];
              var code = r[startPath] ? r[startPath].toLowerCase() : "";
              var desc = r[endPath] ? r[endPath].toLowerCase() : "";
              if (code && (code + "-" + desc).includes(text.toLowerCase())) {
                filteredData.push(r);
              }
            } else {
              item = r[item] ? (r[item]).toLowerCase() : "";
              if (item && item.includes(text.toLowerCase())) {
                filteredData.push(r);
              }
            }
          });
        });

        filteredRptData = filteredData;
        filteredData = [];
      }
      filteredRptData = filteredRptData.filter((val, id, array) => array.indexOf(val) == id);
    }
    this.setState({
      FilteredReportData: filteredRptData,
      searchInput: changedValue,
      totalRecordCountNumber: 0
    });
  }
  public HandleReportIndexChanged = (selectedOption: any, actionObj: any) => {
    var today = new Date();
    var year = today.getMonth() == 0 ? (today.getFullYear() - 1).toString() : today.getFullYear().toString();
    var month = today.getMonth() == 0 ? "12" : (today.getMonth()).toString();
    var indexValue = "";
    var indexCalenderDateField: boolean = false;
    var indexNameField: boolean = false;

    var selectedOptionValue = ((selectedOption.label).toLowerCase()).trim();

    switch (selectedOptionValue) {
      case 'year':
        indexValue = year.toString();
        break;

      case 'accounting period':
        indexValue = padZeroesLeft(month, 2) + "-" + year;
        break;

      default:

        indexCalenderDateField = selectedOptionValue.indexOf('date') >= 0 ? true : false;
        indexNameField = selectedOptionValue.indexOf('name') >= 0 ? true : false;
    }

    this.setState({
      CurrentReportIndex: selectedOption,
      IndexValue: indexValue,
      indexDateValue: indexCalenderDateField,
      startDateValue: null,
      endDateValue: null,
      indexNameValue: indexNameField,
      ReportLevelIndexName: [],
      SelectedReportLevelIndexName: { value: "-1", label: "Search with 3 char or more" },

    });

  }


  public onSelectDate = (date: Date): void => {

    this.setState({
      startDateValue: date,
      endDateValue: date
    });

  }

  public onSelectEndDate = (date: Date): void => {

    this.setState({
      endDateValue: date

    });

  }
  public onFormatDate = (date: Date): string => {
    if (this.state.indexQueryDateFormat == "MM/DD/YYYY") {
      return (date.getMonth() + 1) + "/" + date.getDate() + "/" + date.getFullYear();
    }
    else {
      return date.getDate() + "/" + (date.getMonth() + 1) + "/" + date.getFullYear();
    }
  }


  public HandleAddBracket = (bracket) => {
    var query = this.state.IndexQuery.trim();
    var parsedQuery = this.state.ParsedIndexQuery.trim();
    var allowBracket = false;

    if (query) {
      if (this.state.allowAddIndex == false) {
        if (bracket == ")" && (this.indexQueryArray.lastIndexOf("(") > this.indexQueryArray.lastIndexOf(")"))) {

          query += " " + bracket;
          parsedQuery += " " + bracket;
          allowBracket = true;
        }
      }
      else {
        if (bracket == "(" && this.indexQueryArray.slice(-1) != "(" && (this.indexQueryArray.filter(e => e == "(").length - this.indexQueryArray.filter(e => e == ")").length) == 0) {
          query += " " + bracket;
          parsedQuery += " " + bracket;
          allowBracket = true;
        }
      }
    }
    else {
      if (bracket == "(") {
        query += bracket;
        parsedQuery += bracket;
        allowBracket = true;
      }
    }
    if (allowBracket) {
      this.setState({
        IndexQuery: query,
        ParsedIndexQuery: parsedQuery
      }, () => {
        this.indexQueryArray.push(bracket);
        this.parsedIndexQuery.push(bracket);
      });
    }
  }

  public removeLastAddedQuery = () => {

    var query = "";
    var parsedQuery = "";
    var removeLastAddedQuery = "";
    var allowAddIndex = this.state.allowAddIndex;
    var conditionCount = this.state.ConditionCount;
    if (this.indexQueryArray.length > 0 && this.parsedIndexQuery.length > 0) {
      removeLastAddedQuery = this.indexQueryArray.pop();
      this.parsedIndexQuery.pop();
      this.indexQueryArray.map((i, key) => {
        if (key == this.indexQueryArray.length - 1) {
          query += (i);
        } else {
          query += (i + " ");
        }
      });

      this.parsedIndexQuery.map((i, key) => {
        if (key == this.parsedIndexQuery.length - 1) {
          parsedQuery += (i);
        } else {
          parsedQuery += (i + " ");
        }
      });

      if ((removeLastAddedQuery != "and") && (removeLastAddedQuery != "or") && (removeLastAddedQuery != ")") && (removeLastAddedQuery != "(")) {
        if (conditionCount == 0) {
          conditionCount = 0;
        } else {
          conditionCount--;
        }
        allowAddIndex = true;
      } else {
        if (removeLastAddedQuery != "(") {
          allowAddIndex = false;
        }
      }

      this.setState({
        IndexQuery: query.trim(),
        ParsedIndexQuery: parsedQuery.trim(),
        ConditionCount: conditionCount,
        Condition: this.indexQueryArray.length == 0 ? "" : this.state.Condition,
        allowAddIndex: allowAddIndex
      });
    }

  }

  public addCondition = (condition) => {

    var query = this.state.IndexQuery.trim();
    var parseQuery = this.state.ParsedIndexQuery.trim();
    var peakArrayElement = this.indexQueryArray.slice(-1);
    if (this.state.ConditionCount < 3) {
      if (query && peakArrayElement != "(") {
        if (peakArrayElement == "and" || peakArrayElement == "or") {
          this.indexQueryArray.pop();
          this.indexQueryArray.push(condition);
          this.parsedIndexQuery.pop();
          this.parsedIndexQuery.push(condition);
          query = (query.substring(0, query.lastIndexOf(" ")) + " " + condition);
          parseQuery = (parseQuery.substring(0, parseQuery.lastIndexOf(" ")) + " " + condition);
        } else {
          this.indexQueryArray.push(condition);
          this.parsedIndexQuery.push(condition);
          query += " " + condition;
          parseQuery += " " + condition;
        }
      }
    }

    setTimeout(() => {
      this.setState({
        IndexQuery: query,
        ParsedIndexQuery: parseQuery,
        Condition: "",
        addedCondition: condition,
        allowAddIndex: true
      });
    }, 500);


  }

  public ClearIndexes = () => {
    this.setState({
      IndexQuery: "",
      ParsedIndexQuery: "",
      ConditionCount: 0,
      Condition: "",
      IndexValue: "",
      CurrentReportIndex: { value: '', label: "All" },
      startDateValue: null,
      endDateValue: null,
      ReportLevelIndexName: [],
      SelectedReportLevelIndexName: { value: "-1", label: "Search with 3 char or more" },
      indexDateValue: false,
      indexNameValue: false,
      allowAddIndex: true,
    });

    this.indexQueryArray = [];
    this.parsedIndexQuery = [];
  }


  public HandleAddIndex = () => {
    var showBubble = false;
    var bubbleTarget = "";
    var bubbleMessage = "";
    var regEx = new RegExp(/^\d{2}-\d{4}$/);
    var conditionCount = this.state.ConditionCount;
    var indexQuery = this.state.IndexQuery;
    var currentReportIndex = (this.state.CurrentReportIndex.label).toLowerCase();
    var startDateValue = moment(new Date(this.state.startDateValue)).format(this.state.indexQueryDateFormat);
    var endDateValue = moment(new Date(this.state.endDateValue)).format(this.state.indexQueryDateFormat);
    var query: string = "";
    var SelectedReportLevelIndexName = this.state.SelectedReportLevelIndexName;
    var selectedIndexNames: string = "";
    var indexValue: string = "";
    indexValue = currentReportIndex.indexOf('number') >= 0 ? "number" : currentReportIndex.indexOf('date') >= 0 ? "date" : currentReportIndex.indexOf('name') >= 0 ? "name" : "other";

    if (this.state.CurrentReportIndex.value == "") {
      showBubble = true;
      bubbleTarget = "#ddlIndex";
      bubbleMessage = "Please select the index";
    }
    if (this.state.ConditionCount >= 3) {
      showBubble = true;
      bubbleTarget = "#btnAddIndex";
      bubbleMessage = "Maximum of 3 conditions can be added";
    }
    if (this.state.CurrentReportIndex.value != "") {


      switch (indexValue) {
        case 'date': if (!this.state.startDateValue) {
          showBubble = true;
          bubbleTarget = "#dPStartDateValue";
          bubbleMessage = "Please select the Start Date";
        }
          break;

        case 'name':
          if (this.state.SelectedReportLevelIndexName.value == "-1") {
            showBubble = true;
            bubbleTarget = "#indexLevelValue";
            bubbleMessage = "Please select the Index Value";
          }
          break;

        case 'number':
          var indexNumber = this.state.IndexValue.split(",").length;
          var invoiceNumberLimitation = this.state.invoiceNumberLimitation;
          if (this.state.IndexValue == "") {
            showBubble = true;
            bubbleTarget = "#btnAddIndex";
            bubbleMessage = "Please enter the condition text";
          }
          if (indexNumber > invoiceNumberLimitation) {
            var maxInvoiceNumber = this.state.IndexValue.split(",")[invoiceNumberLimitation - 1];
            var indexValueNumber = this.state.IndexValue.substring(0, this.state.IndexValue.indexOf(maxInvoiceNumber) + maxInvoiceNumber.length);
            this.setState({
              IndexValue: indexValueNumber
            });
            showBubble = true;
            bubbleTarget = "#txtIndexValue";
            bubbleMessage = "Only 5 Number acceptable others will be truncate. Click on Add Index again";
          }
          break;

        case 'other':
          if (this.state.IndexValue == "") {
            showBubble = true;
            bubbleTarget = "#txtIndexValue";
            bubbleMessage = "Please enter the condition text";
          }
          if (this.state.CurrentReportIndex.value == "Period") {
            if (!regEx.test(this.state.IndexValue)) {
              showBubble = true;
              bubbleTarget = "#txtIndexValue";
              bubbleMessage = "Please enter valid accounting period in MM-YYYY format";
            }
          }
          if (this.state.CurrentReportIndex.value == "Year") {
            regEx = new RegExp(/^\d{4}$/);
            if (!regEx.test(this.state.IndexValue)) {
              showBubble = true;
              bubbleTarget = "#txtIndexValue";
              bubbleMessage = "Please enter valid Year in YYYY format";
            }
          }
          break;
      }
    }
    if (this.state.IndexQuery.trim() != "") {
      if (this.state.ConditionCount > 0 && this.state.ConditionCount < 3) {
        if (this.state.allowAddIndex == false) {
          showBubble = true;
          bubbleTarget = "#btnAddIndex";
          bubbleMessage = "Please choose 'and/or' condition and click 'Add Index' button";
        }
      }
    }
    if (!showBubble) {
      if (this.state.CurrentReportIndex.value != "" && (this.state.IndexValue != "" || this.state.startDateValue && this.state.endDateValue || this.state.SelectedReportLevelIndexName.value != -1)) {
        var parsedQuery = this.ConstructParsedQuery();

        if (parsedQuery) {
          switch (indexValue) {
            case 'date':
              indexQuery += " " + this.state.CurrentReportIndex.label + " between ( " + startDateValue + " and " + endDateValue + " )";
              query = " " + this.state.CurrentReportIndex.label + " between ( " + startDateValue + " and " + endDateValue + " )";
              break;

            case 'name':
              if (Array.isArray(SelectedReportLevelIndexName) && SelectedReportLevelIndexName.filter(x => x.value != "-1").length > 0) {
                var selectedFolderPaths = SelectedReportLevelIndexName.map((item) => {
                  return "'" + item.label + "'";
                });
                selectedIndexNames = selectedFolderPaths.join(",");
              }
              indexQuery += " " + this.state.CurrentReportIndex.label + " In(" + selectedIndexNames + ") ";
              query = " " + this.state.CurrentReportIndex.label + " In(" + selectedIndexNames + ") ";
              break;

            case 'number':
              var InviceNumber = this.state.IndexValue.includes(",") ? this.state.IndexValue.split(",").join(",") : this.state.IndexValue;
              indexQuery += " " + this.state.CurrentReportIndex.label + " In(" + InviceNumber + ") ";
              query = " " + this.state.CurrentReportIndex.label + " In(" + InviceNumber + ") ";
              break;

            case 'other':
              indexQuery += " " + this.state.CurrentReportIndex.label + "=" + this.state.IndexValue;
              query = " " + this.state.CurrentReportIndex.label + "=" + this.state.IndexValue;
              break;

          }

          conditionCount += 1;
        }
      }
      this.indexQueryArray.push(query);
      this.setState({
        ConditionCount: conditionCount,
        IndexQuery: indexQuery,
        ParsedIndexQuery: parsedQuery,
        CurrentReportIndex: { value: '', label: 'All' },
        IndexValue: "",
        startDateValue: null,
        endDateValue: null,
        ReportLevelIndexName: [],
        SelectedReportLevelIndexName: { value: "-1", label: "Search with 3 char or more" },
        indexDateValue: false,
        indexNameValue: false,
        allowAddIndex: false,
      });
    }
    else {
      this.setState({
        ShowBubble: showBubble,
        BubbleTarget: bubbleTarget,
        BubbleMessage: bubbleMessage
      });
    }
  }

  public toggleBubble = () => {
    this.setState({
      ShowBubble: !this.state.ShowBubble
    });
  }
  public ConstructParsedQuery = () => {
    try {
      var parsedQuery = this.state.ParsedIndexQuery;
      var startDateValue = moment(new Date(this.state.startDateValue)).format(this.state.parseQueryDateFormat);
      var endDateValue = moment(new Date(this.state.endDateValue)).format(this.state.parseQueryDateFormat);
      var currentReportIndex = (this.state.CurrentReportIndex.label).toLowerCase();
      var query = "";
      var SelectedReportLevelIndexName = this.state.SelectedReportLevelIndexName;
      var selectedIndexNames: string = "";
      var indexValue: string = "";
      indexValue = currentReportIndex.indexOf('period') >= 0 ? 'period' : currentReportIndex.indexOf('number') >= 0 ? "number" : currentReportIndex.indexOf('date') >= 0 ? "date" : currentReportIndex.indexOf('name') >= 0 ? "name" : "other";

      switch (indexValue) {
        case 'period':
          parsedQuery += " (Month eq '" + this.state.IndexValue.split('-')[0] + "' and Year eq '" + this.state.IndexValue.split('-')[1] + "') ";
          query = " (Month eq '" + this.state.IndexValue.split('-')[0] + "' and Year eq '" + this.state.IndexValue.split('-')[1] + "') ";
          break;

        case 'date':
          parsedQuery += " " + this.state.CurrentReportIndex.value + " between '" + startDateValue + "' and '" + endDateValue + "' ";
          query = " " + this.state.CurrentReportIndex.value + " between '" + startDateValue + "' and '" + endDateValue + "' ";
          break;

        case 'name':
          if (Array.isArray(SelectedReportLevelIndexName) && SelectedReportLevelIndexName.filter(x => x.value != "-1").length > 0) {
            var selectedFolderPaths = SelectedReportLevelIndexName.map((item) => {
              return "'" + item.value.split("-")[0].trim() + "'";
            });

            selectedFolderPaths = selectedFolderPaths.filter((val, id, array) => array.indexOf(val) == id);
            selectedIndexNames = selectedFolderPaths.join(",");
          }
          parsedQuery += " " + this.state.CurrentReportIndex.value + " In(" + selectedIndexNames + ") ";
          query = " " + this.state.CurrentReportIndex.value + " In(" + selectedIndexNames + ") ";
          break;

        case 'number':
          if (this.state.IndexValue.includes(",")) {
            var InvoiceNumber: any = this.state.IndexValue.includes(",") ? this.state.IndexValue.split(",") : this.state.IndexValue;
            InvoiceNumber = InvoiceNumber.map((item) => { return "'" + item + "'"; });
            if (this.state.CurrentReportIndex.value == "StdInvoiceId") {
              parsedQuery += " InvoiceId In(select CONSOLIDATED_INV_NUM from w_ar_xact_f where standard_inv_num in (" + InvoiceNumber + ")) ";
              query = " InvoiceId In(select CONSOLIDATED_INV_NUM from w_ar_xact_f where standard_inv_num in (" + InvoiceNumber + ")) ";
            } else {
              parsedQuery += " " + this.state.CurrentReportIndex.value + " In(" + InvoiceNumber + ") ";
              query = " " + this.state.CurrentReportIndex.value + " In(" + InvoiceNumber + ") ";
            }
          }
          else {
            if (this.state.CurrentReportIndex.value == "StdInvoiceId") {
              parsedQuery += " InvoiceId In(select CONSOLIDATED_INV_NUM from w_ar_xact_f where standard_inv_num in ('" + this.state.IndexValue + "')) ";
              query = " InvoiceId In(select CONSOLIDATED_INV_NUM from w_ar_xact_f where standard_inv_num in ('" + this.state.IndexValue + "')) ";
            } else {
              parsedQuery += " " + this.state.CurrentReportIndex.value + " In('" + this.state.IndexValue + "') ";
              query = " " + this.state.CurrentReportIndex.value + " In('" + this.state.IndexValue + "') ";
            }
          }
          break;

        case 'other':
          parsedQuery += " " + this.state.CurrentReportIndex.value + " eq'" + this.state.IndexValue + "' ";
          query = " " + this.state.CurrentReportIndex.value + " eq'" + this.state.IndexValue + "' ";
          break;
      }

      this.parsedIndexQuery.push(query);
    }
    catch (ex) {
      this.HandleException("ConstructParsedQuery", ex);
    }
    return parsedQuery;
  }

  private getFileBase64(file, cb) {
    let reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => {
      cb(reader.result);
    };

    reader.onerror = (error) => {
      throw error;
    };
  }

  public successOpenReportViewModal() {
    this.toggleLoadingModal();
    this.toggleReportViewModal();
  }

  public OpenReportViewModal = (reportObj: any) => {
    try {
      this.toggleLoadingModal();
      let returnBloburl: string = "true";
      let contentType: string = "text/xml";
      let isInline: string = "false";
      if (reportObj.FilePath.includes(".pdf")) {
        returnBloburl = "false";
        contentType = "application/pdf";
        isInline = "true";
      }
      this.webAPI.GetWebAPI(APISource.WebAPI, "File?relativefileurl=" + reportObj.FilePath + "&returnbloburl=" + returnBloburl + "&isInline=" + isInline, contentType)
        .then((response: any): void => {
          Logging.AppInsightsTrackEvent(this.EVENTNAME, EventType.Information, "OpenReportViewModal", "File previewed :" + reportObj.FilePath.split('/')[reportObj.FilePath.split('/').length - 1]);
          if (returnBloburl == "false") {
            this.getFileBase64(response, (result) => {
              this.setState({
                CurrentReportFile: result,
                CurrentReportEmbedURL: reportObj.FilePath
              }, function () {
                this.successOpenReportViewModal();
              });
            });
          } else {
            this.setState({
              CurrentReportFile: response,
              CurrentReportEmbedURL: reportObj.FilePath
            }, function () {
              this.successOpenReportViewModal();
            });
          }
        })
        .catch((error: any) => {
          Logging.AppInsightsTrackEvent(this.EVENTNAME, EventType.Exception, "OpenReportViewModal", "File preview failed:" + reportObj.FilePath.split('/')[reportObj.FilePath.split('/').length - 1]);
          this.toggleLoadingModal();
        });
    }
    catch (ex) {
      this.HandleException("OpenReportViewModal", ex);
    }
  }

  public toggleReportViewModal = () => {
    this.setState({
      manageReportViewModal: !this.state.manageReportViewModal
    });
  }
  public toggleMessageViewModal = () => {
    this.setState({
      managMessageModal: !this.state.managMessageModal
    });
  }



  public BindPermissionsData = () => {
    try {
      //Get Unique countries
      var userPermissions = this.state.UserPermissions;
      var uniqueCountries = userPermissions.map(item => item.Country)
        .filter((value, index, self) => self.indexOf(value) === index);
      var selectedCountry = { value: "-1", label: "Select" };
      if (uniqueCountries.length == 1) {
        selectedCountry = { label: uniqueCountries[0], value: uniqueCountries[0] };
      }
      this.setState({
        Countries: uniqueCountries,
        SelectedCountry: selectedCountry
      }, function () {
        if (uniqueCountries.length == 1) {
          this.LoadDropdownByCountry();
        }
      });
    }
    catch (ex) {
      this.HandleException("BindPermissionsData", ex);
    }
  }

  public LoadDropdownByCountry = () => {
    var countryPermissions = this.state.UserPermissions.filter(x => x.Country == this.state.SelectedCountry.value);
    var reportLevels = countryPermissions.map(item => item.Type)
      .filter((value, index, self) => self.indexOf(value) === index);
    var rptLevels = [];
    reportLevels.forEach((item, index) => {
      if (item == "Region") {
        rptLevels.splice(0, 0, { value: item, label: item });
      }
      else if (item == "Area") {
        rptLevels.splice(1, 0, { value: item, label: item });
      }
      else if (item == "Branch") {
        rptLevels.splice(2, 0, { value: item, label: item });
      }
      else if (item == "Dept") {
        rptLevels.splice(3, 0, { value: item, label: item });
      }
    });
    if (rptLevels.length > 0) {
      rptLevels = [{ value: "All", label: "All" }].concat(rptLevels);
    }
    this.setState({
      ReportLevels: rptLevels,
      //  ReportLevelValues:[],
      SelectedReportLevel: rptLevels[0],
      SelectedReportLevelValue: { value: "All", label: "All" },
      ScreenLoadingMsg: "Processing your request. Please wait...",
    }, function () {
      this.LoadReportLevelValueDDL();
    });
  }

  public LoadReportLevelValueDDL = () => {
    var reportLevelValues = this.state.SelectedReportLevel.value != "All" ? this.state.UserPermissions.filter(x => x.Country == this.state.SelectedCountry.value && x.Type == this.state.SelectedReportLevel.value) : [];
    this.setState({
      ReportLevelValues: reportLevelValues
    });

  }

  public LoadReportTypeDDLByCategory = async (selectedCategory) => {
    try {
      var reportTypes: any[] = this.state.ReportTypesValues;
      var rpt: any[] = [{ value: 'All', label: 'All' }];

      reportTypes = reportTypes.filter((FileReports) => FileReports.Country.includes(this.state.SelectedCountry.value) && FileReports.ReportCategory == selectedCategory.value);
      reportTypes = rpt.concat(reportTypes.map((rptCategory) => {
        return { value: rptCategory.FileNameStarts, label: rptCategory.Title };
      }));

      this.setState({ ReportTypes: reportTypes });
    } catch (ex) {
      this.HandleException("HandleSearchReports", ex);
    }
  }

  public HandleSearchReports = async () => {
    try {
      if (this.ValidateSearchParameters()) {

        this.setState({
          ReportData: [],
          FilteredReportData: [],
          totalRecordCountNumber: 0
        }, function () {
          this.toggleLoadingModal();
        });
        var searchReportRequestModel: any = {};
        var reportTypes: any[];
        searchReportRequestModel.Country = this.state.SelectedCountry.value;
        searchReportRequestModel.ReportCategory = this.state.ReportCategory.value == "-1" ? "" : this.state.ReportCategory.value;
        searchReportRequestModel.ReportType = this.state.ReportType.value == "All" ? searchReportRequestModel.ReportCategory : this.state.ReportType.value;
        searchReportRequestModel.Indexes = this.state.ParsedIndexQuery ? this.state.ParsedIndexQuery.replace(/eq/g, "=") : "";
        var selectedReportLevel = this.state.SelectedReportLevel.value;
        var selectedReportLevelValues = this.state.SelectedReportLevelValue;

        var prefixCode = "";
        var reportLevel = "";
        var lastIndex: number = (this.state.ReportLevels.length - 1);
        switch (selectedReportLevel) {
          case "Dept":
            prefixCode = "D_";
            reportLevel = "D";
            break;
          case "Branch":
            prefixCode = "B_";
            reportLevel = "B";
            break;
          case "Area":
            prefixCode = "A_";
            reportLevel = "A";
            break;
          case "Region":
            prefixCode = "R_";
            reportLevel = "R";
            break;
          case "All":
            if (this.state.ReportType.value == "All") {
              reportLevel = (this.state.ReportLevels[lastIndex].value == "Region" ? "R" : this.state.ReportLevels[lastIndex].value == "Area" ? "A" : this.state.ReportLevels[lastIndex].value == "Branch" ? "B" : this.state.ReportLevels[lastIndex].value == "Dept" ? "D" : "R");
            } else {
              reportTypes = await sp.web.lists.getByTitle("ReportTypeMaster").items.select("ReportLevel").filter("Country eq '" + this.state.SelectedCountry.value + "' and ReportCategory eq '" + this.state.ReportCategory.value + "' and IsActive eq 'Yes' and FileNameStarts eq '" + this.state.ReportType.value + "'").get();
              reportLevel = reportTypes && reportTypes[0] ? reportTypes[0].ReportLevel[0] : "";
            }
            break;
        }

        if (Array.isArray(selectedReportLevelValues) && selectedReportLevelValues.filter(x => x.value != "All").length > 0) {
          var selectedFolderPaths = selectedReportLevelValues.map((item) => {
            return item.value;
          });
          var selectedFolderPathString = selectedFolderPaths.join(",");
          searchReportRequestModel.FolderPathArr = selectedFolderPathString;
        }
        else {
          var reportLevelValues = this.state.ReportLevelValues;
          if (selectedReportLevel == "All") {
            if (this.state.ReportType.value == "All") {
              reportLevelValues = this.state.UserPermissions.filter(x => x.Country == this.state.SelectedCountry.value && x.Type == this.state.ReportLevels[lastIndex].value);
            } else {
              reportLevelValues = this.state.UserPermissions.filter(x => x.Country == this.state.SelectedCountry.value && x.Type == reportTypes[0].ReportLevel);
            }
          }
          var folderPaths = reportLevelValues.map((item) => {
            // return prefixCode + item["Code"];
            return item["Code"];
          });
          var folderPathString = folderPaths.join(",");
          searchReportRequestModel.FolderPathArr = folderPathString;
        }
        searchReportRequestModel.ShowOnly = this.state.ShowParentOnly ? prefixCode.replace(/_/g, "") : "";
        searchReportRequestModel.ShowSelectedLevelOnly = this.state.ShowParentOnly;
        searchReportRequestModel.ReportLevel = reportLevel;
        searchReportRequestModel.NetworkId = this.state.networkID;
        this.setState({
          IsReportLoading: true,
        }, function () {
          console.log("Search Started at" + new Date());
          Logging.AppInsightsTrackEvent(this.EVENTNAME, EventType.Information, "HandleSearchReports", "File search requested with parameters" + JSON.stringify(searchReportRequestModel));
          this.webAPI.PostWebAPI(APISource.WebAPI, "SPOFileSearch", JSON.stringify(searchReportRequestModel))
            .then((response: any): void => {
              console.log("Search Ended at" + new Date());
              this.setState({
                IsReportLoading: false
              });
              Logging.AppInsightsTrackEvent(this.EVENTNAME, EventType.Information, "HandleSearchReports", "File search reponse count " + response.length);
              var result = response; // JSON.parse(response);
              if (result && result.length > 0) {
                this.setState({
                  ReportData: result,
                  FilteredReportData: result,
                  totalRecordCount: result.length,
                  totalRecordCountNumber: result.length
                }, function () {
                  this.toggleLoadingModal();
                });
              }
              else {
                this.setState({
                  ReportData: [],
                  FilteredReportData: [],
                  totalRecordCountNumber: 0
                }, function () {
                  this.toggleLoadingModal();
                });
              }
            })
            .catch((error: any) => {
              Logging.AppInsightsTrackEvent(this.EVENTNAME, EventType.Exception, "HandleSearchReports", "File search failed" + error);
              // console.error(error);
              this.setState({
                IsReportLoading: false,
                ReportData: [],
                FilteredReportData: [],
                totalRecordCountNumber: 0
              }, function () {
                this.toggleLoadingModal();
                this.toggleMessageViewModal();
              });
            });
        });

      }
    }
    catch (ex) {
      this.HandleException("HandleSearchReports", ex);
    }
  }
  public ValidateSearchParameters = () => {
    let isValid: boolean = true;
    var selectedCountry = this.state.SelectedCountry;
    var selectedReportCategory = this.state.ReportCategory;
    var selectedReportType = this.state.ReportType;
    var indexValue = this.state.IndexValue;
    var startDateValue = this.state.startDateValue;
    var endDateValue = this.state.endDateValue;
    var SelectedReportLevelIndexName = this.state.SelectedReportLevelIndexName;
    var lastQueryindex = this.indexQueryArray.slice(-1);
    var showBubble = false;
    var bubbleTarget = "";
    var bubbleMessage = "";
    if (!selectedCountry || !selectedCountry.value || selectedCountry.value == "-1") {
      showBubble = true;
      bubbleTarget = "#country";
      bubbleMessage = "Please select country";
      isValid = false;
    }
    else if (!selectedReportCategory || !selectedReportCategory.value || selectedReportCategory.value == "-1") {
      showBubble = true;
      bubbleTarget = "#reportCategory";
      bubbleMessage = "Please select Report Category";
      isValid = false;
    }
    else if (!selectedReportType || !selectedReportType.value || selectedReportType.value == "") {
      showBubble = true;
      bubbleTarget = "#reportType";
      bubbleMessage = "Please select Report type";
      isValid = false;
    }
    else if (indexValue) {
      showBubble = true;
      bubbleTarget = "#txtIndexValue";
      bubbleMessage = "Kindly click 'Add Index' button to include indexes in your search filter or clear text in 'Index Value' field to continue.";
      isValid = false;
    }
    else if (startDateValue && endDateValue) {
      showBubble = true;
      bubbleTarget = "#dPStartDateValue";
      bubbleMessage = "Kindly click 'Add Index' button to include indexes in your search filter or clear text in 'Index Value' field to continue.";
      isValid = false;
    }
    else if (SelectedReportLevelIndexName.value != "-1") {
      showBubble = true;
      bubbleTarget = "#indexLevelValue";
      bubbleMessage = "Kindly click 'Add Index' button to include indexes in your search filter or clear text in 'Index Value' field to continue.";
      isValid = false;
    }
    else if (lastQueryindex == "and" || lastQueryindex == "or") {
      showBubble = true;
      bubbleTarget = "#btnSearchIndex";
      bubbleMessage = "Kindly add indexes or remove and/or condition";
      isValid = false;
    }
    else if (((this.indexQueryArray.filter(e => e == "(").length - this.indexQueryArray.filter(e => e == ")").length) != 0)) {
      showBubble = true;
      bubbleTarget = "#btnSearchIndex";
      bubbleMessage = "Kindly add closing bracket ')' or remove index ";
      isValid = false;
    }

    if (showBubble) {
      this.setState({
        ShowBubble: showBubble,
        BubbleTarget: bubbleTarget,
        BubbleMessage: bubbleMessage
      });
    }
    return isValid;
  }

  public downloadInvoiceFile = (filePath: string, fileName: string) => {
    var link = document.createElement('a');
    link.href = filePath;
    link.download = fileName;
    link.click();
  }

  public DownloadFile = (reportObj: any) => {
    try {

      let contentType: string = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      if (reportObj.FilePath.includes(".pdf")) {
        contentType = "application/pdf";
      }
      this.webAPI.GetWebAPI(APISource.WebAPI, "File?relativefileurl=" + reportObj.FilePath, contentType)
        .then((response: any): void => {
          let fileName: string = reportObj.FilePath.split('/')[reportObj.FilePath.split('/').length - 1];
          Logging.AppInsightsTrackEvent(this.EVENTNAME, EventType.Information, "DownloadFile", "File downloaded :" + fileName);
          var objectURL = URL.createObjectURL(response);
          this.downloadInvoiceFile(objectURL, fileName);
          URL.revokeObjectURL(objectURL);
        })
        .catch((error: any) => {
          // console.error(error);
          Logging.AppInsightsTrackEvent(this.EVENTNAME, EventType.Exception, "DownloadFile", "File download failed" + error);
        });
    }
    catch (ex) {
      this.HandleException("DownloadFile", ex);
    }
  }

  public handleSort = (column: any, direction: string) => {
    if (column == "Year") {
      var reportData = this.state.FilteredReportData;
      reportData = reportData.sort(SortArrayByPeriod(direction));
      this.setState({
        //ReportData: reportData,
        FilteredReportData: reportData
      });
    }
  }

  public HandleReportSelection = ({ allSelected, selectedCount, selectedRows }) => {
    var selectedReports = [];
    selectedReports = selectedRows;
    var itemSelect = (allSelected == true && selectedRows != 0) ? false : (selectedRows != 0 ? true : false);
    this.setState({
      SelectedReports: selectedReports,
      isAllRowSelect: allSelected,
      isItemSelect: itemSelect
    });
  }

  public PaginationChange = (currentRowsPerPage, currentPage) => {
    this.setState({
      isAllRowSelect: true,
      isItemSelect: false,

    });
  }

  public HandleDownloadAll = (event: any) => {
    try {
      if (!this.state.FilteredReportData || this.state.FilteredReportData.length == 0) {
        return;
      }
      event.preventDefault();

      (this.setState({
        IsDownloadAll: true
      }, () => {
        if (this.state.FilteredReportData.length >= 100) {
          this.toggleReportDownloadConfirmModal();
        } else {
          this.DownloadReports(false);
        }
      }));

    }
    catch (ex) {
      this.HandleException("HandleDownloadAll", ex);
    }
  }
  public HandleDownloadSelected = (event: any) => {
    try {
      if (!this.state.SelectedReports || this.state.SelectedReports.length == 0)
        return;
      event.preventDefault();
      this.setState({
        IsDownloadAll: false
      }, () => {
        if (this.state.SelectedReports.length >= 100) {
          this.toggleReportDownloadConfirmModal();
        } else {
          this.DownloadReports(false);
        }
      });
    }
    catch (ex) {
      this.HandleException("HandleDownloadSelected", ex);
    }
  }

  public toggleReportDownloadConfirmModal = () => {
    this.setState({
      confirmReportDownloadModal: !this.state.confirmReportDownloadModal
    });
  }
  public DownloadZipFiles = async (urls: string[], baseUrl: string) => {
    try {
      const chunkSize = this.state.downloadBatchSize;
      const currentUrlSplit = baseUrl.split('/');
      const groups = urls.map((e, i) => {
        return i % chunkSize === 0 ? urls.slice(i, i + chunkSize) : null;
      }).filter(e => { return e; });

      var AllFileZipRequestModel: any = {};
      AllFileZipRequestModel.SiteName = currentUrlSplit[2];
      AllFileZipRequestModel.DocLibName = currentUrlSplit[3];
      AllFileZipRequestModel.Country = currentUrlSplit[4];
      AllFileZipRequestModel.DownloadIndex = 0;
      var itemsProcessed: number = 0;
      groups.forEach((batchUrl, index) => {
        AllFileZipRequestModel.DownloadIndex++;
        AllFileZipRequestModel.FilePathArr = batchUrl;
        this.webAPI.PostWebAPI(APISource.WebAPI, "AllFileZip", JSON.stringify(AllFileZipRequestModel), "application/json", "application/zip")
          .then((response: any): void => {
            Logging.AppInsightsTrackEvent(this.EVENTNAME, EventType.Information, "DownloadZipFiles", "File downloadeda :" + JSON.stringify(batchUrl));
            saveAs(response, "SRPReportDownload_Part" + index + ".zip");
            itemsProcessed++;
            if (groups.length == itemsProcessed) {
              this.toggleLoadingModal();
            }
          }).catch((error: any) => {
            Logging.AppInsightsTrackException(this.EVENTNAME, "DownloadZipFiles", error);
          });
      });

    }
    catch (ex) {
      this.HandleException("ZipFiles", ex);
    }
  }

  public DownloadReports = (toggleOff: boolean = true) => {
    try {
      if (toggleOff) {
        this.toggleReportDownloadConfirmModal();
      }
      this.toggleLoadingModal();
      let reportList: any;
      if (this.state.IsDownloadAll) {
        reportList = this.state.FilteredReportData;
      }
      else {
        reportList = this.state.SelectedReports;
      }

      var filesToBeZipped: any = [];
      const baseUrl = reportList[0].FilePath.split('/').splice(0, 5).join('/');
      filesToBeZipped = (reportList.slice(0, this.state.downloadFilesCount)).map((reportObj) => {
        return reportObj.FilePath.replace(baseUrl, "");
      });
      this.DownloadZipFiles(filesToBeZipped, baseUrl);
    }
    catch (ex) {
      this.HandleException("DownloadReports", ex);
    }
  }
  public ResetFilters = () => {
    this.indexQueryArray = [];
    this.parsedIndexQuery = [];
    this.setState({
      ParsedIndexQuery: "",
      IndexQuery: "",
      ReportCategory: { value: '-1', label: 'Select' },
      ReportType: { value: 'All', label: 'All' },
      ReportIndexes: [
        { value: '', label: 'Choose an Index' },
      ],
      IndexValue: "",
      CurrentReportIndex: {
        value: '', label: 'Choose an Index',
      },
      ReportData: [],
      FilteredReportData: [],
      ConditionCount: 0,
      ShowBubble: false,
      BubbleTarget: "",
      BubbleMessage: "",
      Condition: "",
      ReportLevels: [],
      SelectedCountry: { value: "-1", label: "Select" },
      SelectedReportLevel: { value: "All", label: "All" },
      SelectedReportLevelValue: { value: "All", label: "All" },
      ReportCategories: [],
      ReportTypes: [],
      ReportLevelValues: [],
      SelectedReports: [],
      confirmReportDownloadModal: false,
      IsDownloadAll: false,
      startDateValue: null,
      endDateValue: null,
      indexDateValue: false,
      ReportLevelIndexName: [],
      SelectedReportLevelIndexName: { value: "-1", label: "Search with 3 char or more" },
      indexNameValue: false,
      allowAddIndex: true,
      isDeptChanged: true,
      columns: [],
      searchInput: ""

    }, function () {
      this.BindPermissionsData();
    });

  }

  public OpenHelpDocumentViewModal = () => {
    try {
      this.webAPI.GetWebAPI(APISource.WebAPI, "File?relativefileurl=" + Environment.HelpManualPath + "&isInline=true", "application/pdf")
        .then((response: any): void => {
          this.getFileBase64(response, (result) => {
            this.setState({
              CurrentHelpDocumentBlob: result
            }, function () {
              this.toggleHelpDocumentViewViewModal();
            });
          });
        });

    }
    catch (ex) {
      this.HandleException("OpenHelpDocumentViewModal", ex);
    }
  }

  public toggleHelpDocumentViewViewModal = () => {
    this.setState({
      manageHelpDocumentViewModal: !this.state.manageHelpDocumentViewModal
    });
  }

  public toggleLoadingModal = () => {
    this.setState({
      manageLoadingModal: !this.state.manageLoadingModal
    });
  }
  public HandleShowParentChanged = (ev: React.MouseEvent<HTMLElement>, checked: boolean) => {
    this.setState({
      ShowParentOnly: checked
    });
  }

  public handleWindowResize = () => {
    var width = (window.innerWidth > 0) ? window.innerWidth : screen.width;
    if (width < 993) {
      for (var i = 0; i < 3; i++) {
        if (document.getElementById("spaceDiv" + i) != null) {
          document.getElementById("spaceDiv" + i).style.display = "none";
        }
      }
      if (document.getElementById("showSelectedDiv") != null) {
        document.getElementById("showSelectedDiv").classList.remove("pl-0");
      }
    }
    else {
      for (var j = 0; j < 3; j++) {
        if (document.getElementById("spaceDiv" + j) != null) {
          document.getElementById("spaceDiv" + j).style.display = "inline-block";
        }
      }
      if (document.getElementById("showSelectedDiv") != null) {
        document.getElementById("showSelectedDiv").classList.add("pl-0");
      }
    }

    if (width == 1024) {
      if (document.getElementById("spLeftNav") != null) {
        document.getElementById("spLeftNav").setAttribute("style", "width:150px !important");
      }
    }
    else {
      if (document.getElementById("spLeftNav") != null) {
        document.getElementById("spLeftNav").setAttribute("style", "width:227px !important");
      }
    }
  }

  public HandleException = (methodName: string, exceptionObj: any) => {
    Logging.AppInsightsTrackException(this.EVENTNAME, methodName, exceptionObj, exceptionObj.message.toString());
  }

  public render(): React.ReactElement<IReportingPortalProps> {
    return (
      <Container fluid={true} className="float-left" >
        <Row>
          <Col lg="12" xl="12" md="12" sm="12" xs="12" className={"pl-0 text-right " + styles.helpIconDiv}>
            {/* <Label id="userNameText" className={"float-left " + styles.userNameText}></Label> */}
            <img src={require('../../../../img/help.png')} alt="Help" className={styles.cursorPointer + " " + styles.helpImage} onClick={(e) => { e.preventDefault(); this.OpenHelpDocumentViewModal(); }} /><br />
            <span className={styles.cursorPointer + " text-info " + styles.helpImage} onClick={(e) => { e.preventDefault(); this.OpenHelpDocumentViewModal(); }}>Help</span>
          </Col>
        </Row>
        {!this.state.IsDataLoading ?
          <React.Fragment>
            {this.state.IsAccessAvailable ?
              <React.Fragment>
                <Row>
                  <Col lg="2" xl="2" md="6" sm="12" xs="12">
                    <Label required className={styles.controlLabel}>Country</Label>
                    <Select
                      id="country"
                      styles={ddlStyle}
                      className="react-select"
                      value={this.state.SelectedCountry}
                      options={this.state.Countries.map((country) => { return { value: country, label: country }; })}
                      onChange={this.HandleCountryChanged}
                      isSearchable={true}
                    />
                  </Col>
                  <Col lg="2" xl="2" md="6" sm="12" xs="12" >
                    <Label required className={styles.controlLabel}>Report Category</Label>
                    <Select
                      id="reportCategory"
                      styles={ddlStyle}
                      value={this.state.ReportCategory}
                      options={this.state.ReportCategories}
                      onChange={this.HandleReportCategoryChanged}
                      isSearchable={true}
                    />
                  </Col>
                  <Col lg="3" xl="3" md="6" sm="12" xs="12" >
                    <Label required className={styles.controlLabel}>Report Type</Label>
                    <Select
                      id="reportType"
                      styles={ddlStyle}
                      value={this.state.ReportType}
                      options={this.state.ReportTypes}
                      onChange={this.HandleReportTypeChanged}
                      isSearchable={true}
                    />
                  </Col>
                  <Col lg="2" xl="2" md="6" sm="12" xs="12">
                    <Label required className={styles.controlLabel}>Report Level</Label>
                    <Select
                      id="reportLevel"
                      styles={ddlStyle}
                      value={this.state.SelectedReportLevel}
                      options={this.state.ReportLevels}
                      onChange={this.HandleReportLevelChanged}
                      isSearchable={true}
                    />
                  </Col>
                  <Col lg="3" xl="3" md="6" sm="12" xs="12" className={""}>
                    <Label required className={styles.controlLabel}>Report Level Value</Label>
                    <Select
                      id="reportLevelValue"
                      styles={reportLevelValueDDlStyle}
                      value={this.state.SelectedReportLevelValue}
                      options={
                        [{ value: "All", label: "All" }].concat(this.state.ReportLevelValues.map((levelValue) => { return { value: levelValue.Code, label: levelValue.Code + " - " + levelValue.Description }; }))}
                      onChange={this.HandleReportLevelValueChanged}
                      isSearchable={true}
                      isMulti={true}
                    />
                  </Col>
                </Row>
                <Row>
                  <Col lg="2" xl="2" md="6" sm="12" xs="12">
                    <Label className={styles.controlLabel}>Indexes</Label>
                    <Select
                      id="ddlIndex"
                      styles={ddlStyle}
                      value={this.state.CurrentReportIndex}
                      options={this.state.ReportIndexes}
                      onChange={this.HandleReportIndexChanged}
                      isSearchable={true}
                    />
                  </Col>
                  {(!this.state.indexDateValue && !this.state.indexNameValue) ?
                    <Col lg="2" xl="2" md="6" sm="12" xs="12">
                      <Label className={styles.controlLabel}>Index Value</Label>
                      <Form.Control id="txtIndexValue" className={styles.indexValue} type="text" value={this.state.IndexValue} onChange={(e) => {
                        this.setState({
                          IndexValue: e.target.value
                        });
                      }} />
                    </Col> :
                    (this.state.indexDateValue) ?
                      <Col lg="2" xl="2" md="6" sm="12" xs="12">
                        <Label required className={styles.controlLabel}>Start Date</Label>
                        <DatePicker
                          id="dPStartDateValue"
                          placeholder={this.state.indexQueryDateFormat}
                          className={styles.datePickerControl}
                          value={this.state.startDateValue}
                          formatDate={this.onFormatDate}
                          onSelectDate={this.onSelectDate}
                        />
                      </Col> :
                      <Col lg="3" xl="3" md="6" sm="12" xs="12">
                        <Label className={styles.controlLabel}>Index Value</Label>
                        <AsyncSelect
                          id="indexLevelValue"
                          styles={indexValueDDlStyle}
                          value={this.state.SelectedReportLevelIndexName}
                          cacheOptions={false}
                          loadOptions={this.state.isGetInvoiceFetchRecord == true ? this.fetchInvoiceValuesDropdown : this.delayToFetchInvoices}
                          defaultOptions
                          onChange={this.onChangeIndexValueDropDown}
                          isMulti={true}
                          cache={false}
                        />
                      </Col>
                  }

                  {this.state.indexDateValue ?
                    <Col lg="2" xl="2" md="6" sm="12" xs="12">

                      <Label className={styles.controlLabel}>End Date</Label>
                      <DatePicker
                        id="dPEndDateValue"
                        placeholder={this.state.indexQueryDateFormat}
                        className={styles.datePickerControl}
                        style={!this.state.startDateValue ? { backgroundColor: '#f3f2f1' } : {}}
                        value={this.state.endDateValue}
                        formatDate={this.onFormatDate}
                        onSelectDate={this.onSelectEndDate}
                        minDate={new Date(this.state.startDateValue)}
                        disabled={!this.state.startDateValue ? true : false}
                      />

                    </Col> :
                    this.state.indexNameValue ?
                      <Col lg="2" xl="1" md="6" sm="12" xs="12">
                        <span className="d-block"></span>
                      </Col> :
                      <Col lg="2" xl="2" md="6" sm="12" xs="12">
                        <span className="d-block"></span>
                      </Col>
                  }

                  <Col lg="1" xl="1" md="6" sm="12" xs="12">
                    <span className="d-block"></span>
                  </Col>
                  <Col lg="3" xl="3" md="6" sm="12" xs="12">
                    <Label className={styles.conditionLabel} ></Label>
                    <Button variant="outline-secondary" type="button" size="sm" className={styles.indexButton} onClick={(e) => {
                      e.preventDefault();
                      this.HandleAddBracket("(");
                    }} >(</Button>
                    <Form.Check
                      // custom
                      inline
                      type={"radio"}
                      name="condition"
                      label={"and"}
                      value="and"
                      checked={this.state.Condition === "and" ? true : false}
                      className={styles.radioBox}
                      onChange={(e) =>
                        this.setState({
                          Condition: e.target.value
                        }, () => {
                          this.addCondition('and');
                        })
                      }
                    />
                    <Form.Check
                      //custom
                      inline
                      name="condition"
                      type={"radio"}
                      label={"or"}
                      value="or"
                      className={styles.radioBox}
                      checked={this.state.Condition === "or" ? true : false}
                      onChange={(e) => this.setState({
                        Condition: e.target.value
                      }, () => {
                        this.addCondition('or');
                      })
                      }
                    />
                    <Button variant="outline-secondary" type="button" size="sm" className={styles.indexButton} onClick={(e) => {
                      e.preventDefault();
                      this.HandleAddBracket(")");
                    }}>)</Button>

                    <div className={styles.spaceDiv} id="spaceDiv0"></div>
                    <div className={styles.spaceDiv} id="spaceDiv0"></div>
                    <Button variant="outline-secondary" type="button" size="sm" className={styles.indexButton} onClick={(e) => {
                      e.preventDefault();
                      this.removeLastAddedQuery();
                    }}>X</Button>

                  </Col>

                  <Col lg="2" xl="2" md="6" sm="12" xs="12" className="text-right pt-2 pl-0" id="showSelectedDiv" >
                    {this.state.SelectedReportLevel && this.state.SelectedReportLevel.value != "All" && this.state.ReportCategory.label == "GL Reports" ?
                      <React.Fragment>
                        <Label className={styles.controlLabel + " " + styles.showSelectedLevelLabel}>Show {this.state.SelectedReportLevel && this.state.ReportLevels.length > 0 ? this.state.SelectedReportLevel.value != "All" ? this.state.SelectedReportLevel.value : this.state.ReportLevels[0].value : ""} Only</Label>
                        <Toggle className={"mt-1 " + styles.toggleButton} defaultChecked={this.state.ShowParentOnly} onText="Yes" offText="No" onChange={this.HandleShowParentChanged} />
                      </React.Fragment>
                      : ""}
                  </Col>
                </Row>
                <Row className="mt-0">
                  <Col xs="12" sm="12" md="12" lg="7" xl="7">
                    <span className="d-block text-info">{this.state.IndexQuery}</span>
                    {this.state.ShowBubble && (
                      <TeachingBubble
                        styles={
                          {
                            content: {
                              background: 'rgb(128,0,0,0.6)',
                              color: 'white'
                            },
                          }
                        }
                        target={this.state.BubbleTarget}
                        onDismiss={this.toggleBubble}
                        closeButtonAriaLabel="Close"
                      >
                        {this.state.BubbleMessage}
                      </TeachingBubble>
                    )}
                  </Col>
                  <Col xs="12" sm="12" md="12" lg={{ span: 4 }} xl={{ span: 5 }} className="" >

                    <Label className={styles.indexLabel}></Label>
                    <Button color="primary" type="button" size="sm" className={styles.portalButton + " ml-0 col-md-12 " + styles.btnIndex} id="btnAddIndex" onClick={(e) => {
                      e.preventDefault();
                      this.HandleAddIndex();
                    }}>Add Index </Button>
                    <div className={styles.spaceDiv} id="spaceDiv0"></div>
                    &nbsp;

                    <Button color="primary" type="button" size="sm" className={styles.portalButton + "  ml-0 col-md-12 " + styles.btnIndex} onClick={(e) => {
                      e.preventDefault();
                      this.ClearIndexes();
                    }}>Clear Index</Button>
                    <div className={styles.spaceDiv} id="spaceDiv0"></div>
                    &nbsp;
                    <Button type="button" color="secondary" size="sm" className={styles.portalButton + "   ml-0 col-md-12 " + styles.functionButton} onClick={(e) => {
                      e.preventDefault();
                      this.ResetFilters();
                    }} >Reset</Button>
                    <div className={styles.spaceDiv} id="spaceDiv1"></div>
                     &nbsp;
                    <Button type="button" color="primary" size="sm" className={styles.portalButton + "  ml-0 col-md-12 " + styles.functionButton} id="btnSearchIndex" onClick={(e) => {
                      e.preventDefault();
                      this.HandleSearchReports();
                    }}  >Search</Button>
                  </Col>
                </Row>
                <hr className="mt-1 mb-0" />
                <Row>
                  <DataTable
                    customStyles={tableStyle}
                    columns={this.state.columns}
                    noHeader
                    striped={true}
                    responsive={true}
                    data={this.state.FilteredReportData}
                    subHeader
                    subHeaderComponent={<ITableHeader totalCountDisplayMsg={this.state.totalRecordCountNumber >= 5000 ? this.state.totalCountDisplayMsg.replace('<ActualCount>', this.state.totalRecordCount) : null} SearchTextChanged={this.HandleTextChanged} SelectedRecords={this.state.SelectedReports} SearchBoxInputText={this.state.searchInput} DownloadSelected={this.HandleDownloadSelected} DownloadAll={this.HandleDownloadAll} />}
                    selectableRows={this.state.columns.length > 0 ? true : false}
                    onSelectedRowsChange={this.HandleReportSelection}
                    highlightOnHover={true}
                    pagination
                    paginationPerPage={10}
                    paginationRowsPerPageOptions={[10, 25, 50, 100]}
                    paginationTotalRows={this.state.FilteredReportData.length}
                    paginationComponentOptions={{ rowsPerPageText: 'Show entries', rangeSeparatorText: 'of', noRowsPerPage: false, selectAllRowsItem: false, selectAllRowsItemText: 'All' }}
                    persistTableHead
                    progressPending={this.state.IsReportLoading}
                    defaultSortField={"Year"}
                    defaultSortAsc={false}
                    clearSelectedRows={this.state.FilteredReportData.length > 0 ? false : true}
                    selectableRowsVisibleOnly={this.state.isAllRowSelect ? true : (this.state.isItemSelect ? false : true)}
                    onChangeRowsPerPage={this.PaginationChange}
                  />
                </Row>
              </React.Fragment>
              : <Row>
                <Col lg="12" xl="12" md="12" sm="12" xs="12">
                  <Label className="d-block text-center text-danger font-weight-bold">You are not an authorized user, Please contact helpdesk to gain proper access.</Label>
                </Col></Row>}
          </React.Fragment>
          : this.state.IsConnectionFailed ?
            <Row>
              <Col lg="12" xl="12" md="12" sm="12" xs="12">
                <Label className="d-block text-center text-danger font-weight-bold">Oops. Something went wrong. Connection to the Web Service might be failed. Please try again later.</Label>
              </Col></Row>
            :
            <Row>
              <Col lg="12" xl="12" md="12" sm="12" xs="12">
                <Spinner label="Please wait..." ariaLive="assertive" labelPosition="top" />
              </Col></Row>
        }
        < Modal isOpen={this.state.manageReportViewModal} toggle={this.toggleReportViewModal}
          className={'modal-xl'} >
          <ModalHeader tag={"span"}> File Viewer</ModalHeader>
          <ModalBody>
            {this.state.CurrentReportEmbedURL.indexOf('pdf') > 1 ?
              <iframe title="File Viewer" src={this.state.CurrentReportFile} className="col-lg-12 min-vh-100">  </iframe> :
              <iframe title="File Viewer" src={"https://view.officeapps.live.com/op/embed.aspx?src=" + this.state.CurrentReportFile} className="col-lg-12 min-vh-100">  </iframe>
            }
          </ModalBody>
          <ModalFooter>
            <Button color="secondary" onClick={this.toggleReportViewModal}>Close</Button>
          </ModalFooter>
        </Modal >

        <Modal backdrop={'static'} keyboard={false} isOpen={this.state.confirmReportDownloadModal} toggle={this.toggleReportDownloadConfirmModal}
          className={'modal-lg'}  >
          <ModalHeader toggle={this.toggleReportDownloadConfirmModal} tag={"span"}><b> Download Confirmation </b></ModalHeader>
          <ModalBody>
            The number of files that needs to be downloaded will affect the performance. Do you want to proceed to download as batches({this.state.downloadBatchSize} files per batch)? <br /><b>Note : The Maximum {this.state.downloadFilesCount} Files will be Downloaded. </b>
          </ModalBody>
          <ModalFooter>
            <React.Fragment>
              <Button color="secondary" onClick={(e) => { this.DownloadReports(); }}>Ok</Button>
              <Button color="default" onClick={this.toggleReportDownloadConfirmModal}>Cancel</Button>
            </React.Fragment>
          </ModalFooter>
        </Modal>

        <Modal isOpen={this.state.manageHelpDocumentViewModal} toggle={this.toggleHelpDocumentViewViewModal}
          className={'modal-xl'} >
          <ModalHeader toggle={this.toggleHelpDocumentViewViewModal} tag={"span"}> SharePoint dashboard help document</ModalHeader>
          <ModalBody>
            <iframe title="File Viewer" src={this.state.CurrentHelpDocumentBlob} className="col-lg-12 min-vh-100">  </iframe> :
          </ModalBody>
          <ModalFooter>
            <Button color="secondary" onClick={this.toggleHelpDocumentViewViewModal}>Close</Button>
          </ModalFooter>
        </Modal >
        <FabricModal
          isOpen={this.state.manageLoadingModal}
          isBlocking={true}
          className={styles.loadingModal}
        >
          <Spinner label={this.state.ScreenLoadingMsg} ariaLive="assertive" labelPosition="top" className={styles.loadingModalSpinner} />

        </FabricModal>

        <Modal isOpen={this.state.managMessageModal} toggle={this.toggleMessageViewModal}
          className={'modal-md'} >
          <ModalHeader tag={"span"}> Information</ModalHeader>
          <ModalBody>
            Something went wrong please try again.
          </ModalBody>
          <ModalFooter>
            <Button color="secondary" onClick={this.toggleMessageViewModal}>Ok</Button>
          </ModalFooter>
        </Modal >

      </Container >
    );
  }
}


