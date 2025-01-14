/* eslint-disable prefer-const */
/* eslint-disable no-unused-expressions */
/* eslint-disable no-constant-condition */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable max-lines */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable no-void */
import * as React from "react";
import { IViewFormProps } from "../IViewFormProps";
import { IDropdownOption, Modal, Stack } from "office-ui-fabric-react";
import {
  IconButton,
  Text,
  PrimaryButton,
  DefaultButton,
  IColumn,
  DetailsList,
  SelectionMode,
  Dialog,
  DialogFooter,
  mergeStyleSets,
} from "@fluentui/react";
import styles from "../Form.module.scss";
import ApproverAndReviewerTableInViewForm from "./simpleTable/reviewerAndApproverTableInViewForm";
import CommentsLogTable from "./simpleTable/commentsTable";
import WorkFlowLogsTable from "./simpleTable/workFlowLogsTable";
import FileAttatchmentTable from "./simpleTable/fileAttatchmentsTable";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { DialogBlockingExample } from "./dialogFluentUi/dialogFluentUi";
import { format } from "date-fns";
import GeneralCommentsFluentUIGrid from "./simpleTable/generalComment";
import UploadFileComponent from "./uploadFile";
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { v4 } from "uuid";
import { ATRAssignee } from "./ATR/atr";
import SuccessDialog from "./dialogFluentUi/endDialog";
import ReferBackCommentDialog from "./dialogFluentUi/referBackCommentDialog";
import RejectBtnCommentCheckDialog from "./dialogFluentUi/rejectCommentsCheckDialog";
import ReturnBtnCommentCheckDialog from "./dialogFluentUi/returnCommentsCheck";
import PDFViewer from "./pdfviewPdfDist/pdfDist";
import PasscodeModal from "./passCode/passCode";
import GistDocsConfirmation from "./dialogFluentUi/gistDocsConfirmationDialog";

import { MarkInfo } from "./markInfo/markInfo";

import "@pnp/sp/profiles";
import GistDocSubmitted from "./dialogFluentUi/gistDocs";
import GistDocEmptyModal from "./dialogFluentUi/gistDocEmptyModal";
import AutoSaveFailedDialog from "./dialogFluentUi/autoSaveFailedDialog";
import NotedCommentDialog from "./dialogFluentUi/notedCommentsDialog";
import SupportingDocumentsUploadFileComponent from "./supportingDocuments";
import CummulativeErrorDialog from "./dialogFluentUi/cummulativeDialog";
import ReferCommentsMandatoryDialog from "./dialogFluentUi/referCommentsMandiatory";
import ChangeApproverMandatoryDialog from "./dialogFluentUi/changeApproverMandiatory";

export interface IFileDetails {
  name?: string;
  content?: File;
  index?: number;
  fileUrl?: string;
  ServerRelativeUrl?: string;
  isExists?: boolean;
  Modified?: string;
  isSelected?: boolean;
}
const Cutsomstyles = mergeStyleSets({
  modal: {
    padding: "10px",
    minWidth: "300px",
    maxWidth: "80vw",
    width: "100%",
    "@media (min-width: 768px)": {
      maxWidth: "580px",
    },
    "@media (max-width: 767px)": {
      maxWidth: "290px",
    },
    margin: "auto",
    backgroundColor: "white",
    borderRadius: "4px",
    boxShadow: "0 2px 8px rgba(0, 0, 0, 0.26)",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",

    borderBottom: "1px solid #ddd",
    height: "50px",
  },
  headerTitle: {
    margin: "5px",
    marginLeft: "0px",
    fontSize: "16px",
    fontWeight: "400",
  },
  body: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    textAlign: "center",
    padding: "20px 0",
    height: "100%",
    "@media (min-width: 768px)": {
      marginLeft: "20px",
      marginRight: "20px",
    },
    "@media (max-width: 767px)": {
      marginLeft: "20px",
      marginRight: "20px",
    },
  },
  footer: {
    display: "flex",
    alignItem: "center",
    justifyContent: "flex-end",

    borderTop: "1px solid #ddd",
    paddingTop: "12px",
    height: "50px",
  },
  button: {
    maxHeight: "32px",
  },
});

const GeneralSectionInViewForm = (props: any): any => {
  return (
    <div className={styles.sectionContainer}>
      <button
        className={styles.header}
        onClick={() => props._onToggleSection(`generalSection`)}
      >
        <Text className={styles.sectionText}>General Section</Text>
        <IconButton
          iconProps={{
            iconName: props.expandSections.generalSection
              ? "ChevronUp"
              : "ChevronDown",
          }}
          title="Expand/Collapse"
          ariaLabel="Expand/Collapse"
          className={styles.chevronIcon}
        />
      </button>
      {props.expandSections.generalSection && (
        <div className={`${styles.expansionPanelInside}`}>
          <div style={{ padding: "15px", paddingTop: "4px" }}>
            {props._renderTable(props.state.eCommitteData[0].tableData)}
          </div>
        </div>
      )}
    </div>
  );
};

const DraftResolutionInViewForm = (props: any): any => {
  return (
    props.formType === "BoardNoteView" && (
      <div className={styles.sectionContainer}>
        <button
          className={styles.header}
          onClick={() => props._onToggleSection(`draftResolution`)}
        >
          <Text className={styles.sectionText}>Draft Resolution Section</Text>
          <IconButton
            iconProps={{
              iconName: props.expandSections.draftResolution
                ? "ChevronUp"
                : "ChevronDown",
            }}
            title="Expand/Collapse"
            ariaLabel="Expand/Collapse"
            className={styles.chevronIcon}
          />
        </button>
        {props.expandSections.draftResolution && (
          <div className={`${styles.expansionPanelInside}`}>
            <div style={{ padding: "15px", paddingTop: "4px" }}>
              <RichText
                value={props.state.draftResolutionFieldValue}
                isEditMode={false}
              />
            </div>
          </div>
        )}
      </div>
    )
  );
};

const ReviewerOrApproverSectionInViewForm = (props: any): any => {
  return (
    <div className={styles.sectionContainer}>
      <button
        className={styles.header}
        onClick={() => props._onToggleSection(props.toggleParameter)}
      >
        <Text className={styles.sectionText}>{props.sectionName}</Text>
        <IconButton
          iconProps={{
            iconName: props.expandSections[props.toggleParameter]
              ? "ChevronUp"
              : "ChevronDown",
          }}
          title="Expand/Collapse"
          ariaLabel="Expand/Collapse"
          className={styles.chevronIcon}
        />
      </button>
      {props.expandSections[props.toggleParameter] && (
        <div className={`${styles.expansionPanelInside}`}>
          <div style={{ padding: "15px", paddingTop: "4px" }}>
            <ApproverAndReviewerTableInViewForm
              data={props.reviewerORApproverData}
              reOrderData={props.reOrderData}
            
              type={props.type}
            />
          </div>
        </div>
      )}
    </div>
  );
};

type FieldValueType = string | number | readonly string[];

export interface IViewFormState {
  title: string;
  expandSections: { [key: string]: boolean };
  pdfLink: string;
  isLoading: boolean;
  isDataLoading: boolean;
  department: string;
  departmentAlias: string;
  noteTypeValue?: IDropdownOption;
  isNoteType: boolean;
  new: string;
  itemsFromSpList: any[];
  getAllDropDownOptions: any;
  natureOfNote: string[];
  natureOfApprovalSancation: string[];
  committename: string[];
  typeOfFinancialNote: string[];
  noteType: string[];
  isPuroposeVisable: boolean;
  isAmountVisable: boolean;
  isTypeOfFinacialNote: boolean;
  isNatureOfApprovalOrSanction: boolean;

  committeeNameFeildValue: string;
  subjectFeildValue: string;

  natureOfNoteFeildValue: string;
  noteTypeFeildValue: string;
  natureOfApprovalOrSanctionFeildValue: string;
  typeOfFinancialNoteFeildValue: string;
  searchTextFeildValue: FieldValueType;
  amountFeildValue: FieldValueType;
  puroposeFeildValue: FieldValueType;
  othersFieldValue: any;
  // eslint-disable-next-line @rushstack/no-new-null
  notePdfFile: File | null;
  // eslint-disable-next-line @rushstack/no-new-null
  supportingFile: File | null;
  isWarning: boolean;
  isWarningCommittteeName: boolean;
  isWarningSubject: boolean;
  isWarningNatureOfNote: boolean;
  isWarningNatureOfApporvalOrSanction: boolean;
  isWarningNoteType: boolean;
  isWarningTypeOfFinancialNote: boolean;

  isWarningSearchText: boolean;

  isWarningAmountField: boolean;
  isWarningPurposeField: boolean;
  eCommitteData: any;
  noteTofiles: any[];
  isWarningNoteToFiles: boolean;

  wordDocumentfiles: any[];
  isWarningWordDocumentFiles: boolean;

  supportingDocumentfiles: any[];
  isWarningSupportingDocumentFiles: boolean;

  supportingFilesInViewForm: any[];

  errorOfDocuments: any;
  errorFilesList: any;
  errorForCummulative: any;
  dialogboxForCummulativeError: any;

  isWarningPeoplePicker: boolean;
  isDialogHidden: boolean;
  isApproverOrReviewerDialogHandel: boolean;

  peoplePickerData: any;
  peoplePickerApproverData: any;
  approverInfo: any;
  reviewerInfo: any;

  status: string;
  statusNumber: any;
  auditTrail: any;
  filesClear: any;
  createdByEmail: any;
  createdByID: any;
  createdByEmailName: any;
  ApproverDetails: any;
  ApproverOrder: any;
  ApproverType: any;

  dialogFluent: any;
  dialogDetails: any;

  commentsData: any;
  generalComments: any;
  commentsLog: any;
  referComment: any;

  currentApprover: any;
  pastApprover: any;
  referredFromDetails: any;
  refferredToDetails: any;
  noteReferrerDTO: any;

  noteSecretaryDetails: any;
  secretaryGistDocs: any[];
  secretaryGistDocsList: any[];

  atrCreatorsList: any;
  atrGridData: any;
  noteATRAssigneeDetails: any;
  noteATRAssigneeDetailsAllUser: any;
  atrJoinedComments: any;
  atrType: any;

  isDialogVisible: any;
  dialogContent: any;

  isVisibleAlter: boolean;
  successStatus: any;
  isGistSuccessVisibleAlter: boolean;

  isReferDataAndCommentsNeeded: boolean;

  isChangeApproverNeeded: boolean;

  noteReferrerCommentsDTO: any;
  isReferBackAlterDialog: boolean;

  isRejectCommentsCheckAlterDialog: boolean;

  isReturnCommentsCheckAlterDialog: boolean;

  isNotedCommentsManidatoryAlterDialog: boolean;

  draftResolutionFieldValue: any;

  isPasscodeModalOpen: boolean;
  isPasscodeValidated: boolean;

  passCodeValidationFrom: any;

  isGistDocCnrf: boolean;
  isGistDocEmpty: boolean;

  noteMarkedInfoDTOState: any;

  isAutoSaveFailedDialog: any;

  isUserExistsModalVisible: any;

  approverIdsHavingSecretary: any;
  hideParellelActionAlertDialog: boolean;
  parellelActionAlertMsg: string;

  peoplePickerSelectedDataWhileReferOrChangeApprover: any;
}

const getIdFromUrl = (): any => {
  const params = new URLSearchParams(window.location.search);
  const Id = params.get("itemId");

  return Id;
};

export default class ViewForm extends React.Component<
  IViewFormProps,
  IViewFormState
> {
  private _itemId: number = Number(getIdFromUrl());
  private _currentUserEmail = this.props.context.pageContext.user.email;

  private _absUrl: any = this.props.context.pageContext.web.serverRelativeUrl;
  private _folderName: any = "";
  private _committeeType: any =
    this.props.formType === "BoardNoteView" ? "Board" : "eCommittee";

  private _committeeTypeForATR: any =
    this.props.formType === "BoardNoteView" ? "boardnote" : "committeenote";

  private _listname: any;
  private _libraryName: any;

  private _folderNameAfterApproved: any;

  constructor(props: IViewFormProps) {
    super(props);
    this.state = {
      title: "",
      isLoading: false,
      isDataLoading: true,
      department: "",
      departmentAlias: "",
      isNoteType: false,
      noteTypeValue: undefined,
      new: "",
      itemsFromSpList: [],
      getAllDropDownOptions: {},
      natureOfNote: [],
      committename: [],
      natureOfApprovalSancation: [],
      typeOfFinancialNote: [],
      noteType: [],
      isPuroposeVisable: false,
      isAmountVisable: false,
      isTypeOfFinacialNote: false,
      isNatureOfApprovalOrSanction: false,

      committeeNameFeildValue: "",
      subjectFeildValue: "",
      natureOfNoteFeildValue: "",
      noteTypeFeildValue: "",
      natureOfApprovalOrSanctionFeildValue: "",
      typeOfFinancialNoteFeildValue: "",
      searchTextFeildValue: "",
      amountFeildValue: 0,
      puroposeFeildValue: "",
      othersFieldValue: "",
      notePdfFile: null,
      supportingFile: null,
      isWarning: false,
      isWarningCommittteeName: false,
      isWarningSubject: false,
      isWarningNatureOfNote: false,
      isWarningNatureOfApporvalOrSanction: false,
      isWarningNoteType: false,
      isWarningTypeOfFinancialNote: false,
      isWarningSearchText: false,
      isWarningAmountField: false,
      isWarningPurposeField: false,
      isWarningPeoplePicker: false,
      eCommitteData: [],
      noteTofiles: [],
      isWarningNoteToFiles: false,

      wordDocumentfiles: [],
      isWarningWordDocumentFiles: false,

      supportingDocumentfiles: [],
      isWarningSupportingDocumentFiles: false,

      supportingFilesInViewForm: [],

      errorOfDocuments: false,
      errorFilesList: {
        wordDocument: [],
        notePdF: [],
        supportingDocument: [],
        gistDocument: [],
        cummlativeError: [],
      },

      errorForCummulative: false,
      dialogboxForCummulativeError: false,

      isDialogHidden: true,
      isApproverOrReviewerDialogHandel: true,
      peoplePickerData: [],
      peoplePickerApproverData: [],
      ApproverDetails: [],
      approverInfo: [],
      ApproverType: "",
      reviewerInfo: [],
      status: "",
      statusNumber: null,
      auditTrail: [],
      filesClear: [],
      expandSections: { generalSection: true },
      pdfLink: "",

      createdByEmail: "",
      createdByID: "",
      createdByEmailName: "",
      ApproverOrder: "",
      dialogFluent: true,
      dialogDetails: {},
      commentsData: [],
      generalComments: [],
      commentsLog: [],
      referComment: [],

      currentApprover: [],
      pastApprover: [],
      referredFromDetails: [],
      refferredToDetails: [],
      noteReferrerDTO: [],

      noteSecretaryDetails: [],
      secretaryGistDocs: [],
      secretaryGistDocsList: [],

      atrCreatorsList: [],
      atrGridData: [],
      noteATRAssigneeDetails: [],
      noteATRAssigneeDetailsAllUser: [],
      atrJoinedComments: [],
      atrType: "Default",

      isDialogVisible: false,
      dialogContent: {},

      isVisibleAlter: false,
      isGistSuccessVisibleAlter: false,
      successStatus: "",

      isReferDataAndCommentsNeeded: false,

      isChangeApproverNeeded: false,

      noteReferrerCommentsDTO: [],
      isReferBackAlterDialog: false,

      isRejectCommentsCheckAlterDialog: false,

      isReturnCommentsCheckAlterDialog: false,

      isNotedCommentsManidatoryAlterDialog: false,

      draftResolutionFieldValue: "",

      isPasscodeModalOpen: false,
      isPasscodeValidated: false,
      passCodeValidationFrom: "",

      isGistDocCnrf: false,
      isGistDocEmpty: false,

      noteMarkedInfoDTOState: [],

      isAutoSaveFailedDialog: false,

      isUserExistsModalVisible: false,
      approverIdsHavingSecretary: [],
      hideParellelActionAlertDialog: false,
      parellelActionAlertMsg: "",

      peoplePickerSelectedDataWhileReferOrChangeApprover: [],
    };

    const listTitle = this.props.listId;

    this._listname = listTitle?.title;

    const libraryTilte = this.props.libraryId;
    this._libraryName = libraryTilte?.title;
    this._fetchApproverDetails();
    this._fetchATRCreatorDetails();
    this._getItemData(this._itemId, this._folderName).then(async () => {
      this._folderName = await `${this._absUrl}/${
        this._libraryName
      }/${this._folderNameGenerate(this._itemId)}`;

      await this._getItemDocumentsData();
    });
  }

  private getMessageBasedOnStatusNumber(statusNumber: string): string {
    const statusMessages: { [key: string]: string } = {
      "2000": "This request has already been submitted",
      "3000": "This request has already been submitted",
      "300": "This request has already been canceled.",
      "4000": "This request has already been referred",
      "9000": "This request has already been approved",
      "8000": "This request has already been rejected",
      "4900": "This request has already been refereed back",
      "5000": "This request has been already returned.",
    };

    return statusMessages[statusNumber] || "Unknown status.";
  }

  private _getUserProperties = async (loginName: any): Promise<any> => {
    let designation;
    let email;

    const profile = await this.props.sp.profiles.getPropertiesFor(loginName);

    designation = profile.Title;
    email = profile.Email;

    const props: any = {};
    profile.UserProfileProperties.forEach(
      (prop: { Key: string | number; Value: any }) => {
        props[prop.Key] = prop.Value;
      }
    );

    profile.userProperties = props;

    return [designation, email];
  };

  private _fetchApproverDetails = async (): Promise<void> => {
    try {
      (
        await this.props.sp.web.lists
          .getByTitle("ApproverMatrix")
          .items.select(
            "*",
            "Approver/Title",
            "Approver/EMail",
            "Secretary/Title",
            "Secretary/EMail"
          )
          .expand("Approver", "Secretary")()
      ).map(async (each: any) => {
        const user = await this.props.sp.web.siteUsers.getById(
          each.ApproverId
        )();

        const dataRec = await this._getUserProperties(user.LoginName);

        if (each.ApproverType === "Approver") {
          const newObj = {
            text: each.Approver.Title,
            email: each.Approver.EMail,
            ApproversId: each.ApproverId,
            approverType: each.ApproverType,

            Title: each.Title,
            id: each.ApproverId,
            secretary: each.Secretary.Title,
            srNo: each.Approver.EMail.split("@")[0],
            optionalText: dataRec[0],
            approverTypeNum: 2,
          };

          const secretaryObj = {
            noteSecretarieId: each.SecretaryId,
            noteApproverId: each.ApproverId,
            noteId: "",
            secretaryEmail: each.Secretary.EMail,
            approverEmail: each.Approver.EMail,
            approverEmailName: each.Approver.Title,
            secretaryEmailName: each.Secretary.Title,
            createdBy: "",
            modifiedDate: "",
            modifiedBy: "",
          };
          this.setState((prev) => {
            this.setState({
              approverIdsHavingSecretary: [
                ...prev.approverIdsHavingSecretary,
                {
                  ApproverId: each.ApproverId,
                  SecretaryId: each.SecretaryId,
                  ...secretaryObj,
                },
              ],
            });
          });
          if (each.ApproverType === "Approver" && !this._itemId) {
            this.setState({ peoplePickerApproverData: [newObj] });
          }
        } else {
          const user = await this.props.sp.web.siteUsers.getById(
            each.ApproverId
          )();

          const dataRec = await this._getUserProperties(user.LoginName);

          const newObj = {
            text: each.Approver.Title,
            email: each.Approver.EMail,
            ApproversId: each.ApproverId,
            approverType: each.ApproverType,

            Title: each.Title,
            id: each.ApproverId,
            secretary: each.Secretary.Title,
            optionalText: dataRec[0],
            srNo: each.Approver.EMail.split("@")[0],

            approverTypeNum: 1,
          };

          if (!this._itemId) {
            this.setState({ peoplePickerData: [newObj] });
          }
        }
      });
    } catch (error) {
      return error;
    }
  };

  private _fetchATRCreatorDetails = async (): Promise<void> => {
    try {
      (
        await this.props.sp.web.lists
          .getByTitle("ATRCreators")
          .items.select(
            "*",
            "Author/Title",
            "Author/EMail",
            "Editor/Title",
            "Editor/EMail",
            "ATRCreators/Title",
            "ATRCreators/EMail"
          )
          .expand("Author", "ATRCreators", "Editor")()
      ).map((each: any) => {
        this.setState((prev) => {
          return {
            atrCreatorsList: [
              ...prev.atrCreatorsList,
              {
                atrCreatorId: each.ATRCreatorsId,
                atrCreatorEmail: each.ATRCreators.EMail,
                atrCreatorEmailName: each.ATRCreators.Title,
                createdDate: each.Created,
                createdBy: each.Author.EMail,
                modifiedDate: each.Modified,
                modifiedBy: each.Author.EMail,
                statusMessage: null,
              },
            ],
          };
        });
        return each;
      });
    } catch (error) {
      return error;
    }
  };

  public _folderNameGenerate(id: any): any {
    const currentyear = new Date().getFullYear();
    const nextYear = (currentyear + 1).toString().slice(-2);

    const requesterNo =
      this.props.formType === "BoardNoteView"
        ? `${this.state.title.split("/")[0]}/${currentyear}-${nextYear}/B${id}`
        : `${this.state.title.split("/")[0]}/${currentyear}-${nextYear}/C${id}`;

    this._folderNameAfterApproved = requesterNo.replace(/\//g, "-");

    const folderName = requesterNo.replace(/\//g, "-");
    return folderName;
  }

  private _getJsonifyReviewer = (item: any, type: string): any[] => {
    const parseItem = JSON.parse(item);
    const approverfilterData = parseItem.filter((each: any) => {
      if (each.approverType === "Reviewer") {
        return each;
      }
    });

    return approverfilterData;
  };

  private _getJsonifyApprover = (item: any, type: string): any[] => {
    const parseItem = JSON.parse(item);
    const approverfilterData = parseItem.filter((each: any) => {
      if (each.approverType === "Approver") {
        return each;
      }
    });

    return approverfilterData;
  };

  private _extractValueFromHtml = (htmlString: string): string => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(htmlString, "text/html");
    const extractedValue = doc.querySelector("div")?.textContent || "";

    return extractedValue;
  };

  private _getdataofMarkedInfo = async (
    data: any,
    idData: any
  ): Promise<any> => {
    if (!Array.isArray(data)) {
      return [];
    }

    const ids = await Promise.all(
      data.map(async (each: any) => {
        const userInfo = { text: each.Title, email: each.EMail };
        try {
          const users = await this.props.sp.web.siteUsers.getByEmail(
            userInfo.email
          )();

          const id = users.Id;

          return { ...userInfo, id };
        } catch (error) {
          return { ...userInfo, id: null };
        }
      })
    );

    return ids;
  };

  private _getCommentsData = (data: any) => {
    const uniqueIds = new Set<string>();

    const filterdata = data
      .filter((each: any) => each !== null)
      .filter((each: any) => {
        if (!uniqueIds.has(each.id)) {
          uniqueIds.add(each.id);
          return true;
        }
        return false;
      });

    return filterdata;
  };

  private _getATRGridData = (data: any) => {
    const newATRGridData = JSON.parse(data)

      .map((each: any) => {
      
        if (each.atrCreatorEmail === this._currentUserEmail) {
          this.setState({ atrType: each.atrType });
          return {
            comments: each.noteApproverComments,
            assignedTo: each.atrAssigneeEmailName,
            status: "Submitted",
          };
        }
      })
      .filter((each: any) => each !== undefined);

    return newATRGridData;
  };

  private _getItemDataSpList = async (id: any) => {
    const item: any = await this.props.sp.web.lists
      .getByTitle(this._listname)
      .items.getById(id)
      .select(
        "*",
        "Author/Title",
        "Author/EMail",
        "Approvers",
        "Approvers/Title",
        "Reviewers/Title",
        "Approvers/EMail",
        "Reviewers/EMail",
        "NoteMarkedInfoDTO/Title",
        "NoteMarkedInfoDTO/EMail",
        "CurrentApprover/Title",
        "CurrentApprover/EMail"
      )
      .expand(
        "Author",
        "Approvers",
        "Reviewers",
        "CurrentApprover",
        "NoteMarkedInfoDTO"
      )();

    console.log(item);

    return item;
  };

  private _generateTableData = (item: any, purposeData: any[]) => {
    return [
      item.CommitteeName !== null && {
        column1: "Note Number",
        column2: `${item.Title}`,
      },
      item.CommitteeName !== null && {
        column1: "Requester",
        column2: `${item.Author.Title}`,
      },
      item.Created !== null && {
        column1: "Request Date",
        column2: `${this._formatDateTime(item.Created)}`,
      },
      item.Status !== null && {
        column1: "Status",
        column2: `${item.Status}`,
      },
      {
        column1: "Current Approver",
        column2: item?.CurrentApprover?.Title,
      },
      item.Department !== null && {
        column1: "Department",
        column2: `${item.Department}`,
      },

      item.CommitteeName !== null && {
        column1: "CommitteeName",
        column2: `${item.CommitteeName}`,
      },
      item.Subject !== null && {
        column1: "Subject",
        column2: `${item.Subject}`,
      },
      item.NatureOfNote !== null && {
        column1: "NatureOfNote",
        column2: `${item.NatureOfNote}`,
      },
      item.NoteType !== null && {
        column1: "NoteType",
        column2: `${item.NoteType}`,
      },
      item.NatureOfApprovalOrSanction !== null && {
        column1: "NatuerOfApprovalSanction",
        column2: `${item.NatureOfApprovalOrSanction}`,
      },

      item.FinancialType !== null && {
        column1: "TypeOfFinancialNote",
        column2: `${item.FinancialType}`,
      },
      item.Amount !== null && {
        column1: "Amount",
        column2: `₹ ${item.Amount}`,
      },
      item.SearchKeyword !== null && {
        column1: "Search Keyword",
        column2: item.SearchKeyword,
      },

      purposeData[0] !== "" && {
        column1: "Purpose",
        column2: `${purposeData[0]}`,
      },
      purposeData[1] !== undefined && {
        column1: "Others",
        column2: `${purposeData[1]}`,
      },
    ];
  };

  private _getAdditionalData = async (item: any) => {
    return {
      committeeNameFeildValue:
        item.CommitteeName !== null ? item.CommitteeName : "",
      subjectFeildValue: item.Subject !== null ? item.Subject : "",
      natureOfNoteFeildValue:
        item.NatureOfNote !== null ? item.NatureOfNote : "",
      noteTypeFeildValue: item.NoteType !== null ? item.NoteType : "",
      natureOfApprovalOrSanctionFeildValue:
        item.NatureOfApprovalOrSanction !== null
          ? item.NatureOfApprovalOrSanction
          : "",
      typeOfFinancialNoteFeildValue:
        item.FinancialType !== null ? item.FinancialType : "",
      searchTextFeildValue:
        item.SearchKeyword !== null
          ? this._extractValueFromHtml(item.SearchKeyword)
          : "",
      amountFeildValue: item.Amount !== null ? item.Amount : null,
      puroposeFeildValue:
        item.Purpose !== null ? item.Purpose.split(",")[0] : "",
      othersFieldValue: item.Purpose !== null ? item.Purpose.split(",")[1] : "",

      peoplePickerData: this._getJsonifyReviewer(
        item.NoteApproversDTO,
        "Reviewer"
      ),
      peoplePickerApproverData: this._getJsonifyApprover(
        item.NoteApproversDTO,
        "Approver"
      ),
      auditTrail: JSON.parse(item.AuditTrail),
      isDataLoading: false,
      createdByEmail: item.Author.EMail,
      createdByEmailName: item.Author.Title,
      createdByID: item.AuthorId,
    };
  };

  private _getItemData = async (id: any, folderPath?: any) => {
    const item: any = await this._getItemDataSpList(id);

    console.log(item, `Item .........${id}`);

    const purposeData = item.Purpose !== null ? item.Purpose.split(",") : "";

    console.log(purposeData, "Purpose data");

    const tableData = this._generateTableData(item, purposeData);

    const additionalData = await this._getAdditionalData(item);
    this.setState({
      eCommitteData: [
        {
          tableData,
        },
      ],
    });
    this.setState({
      ...additionalData,

      status:
        item.Status === "Submitted"
          ? this._getStatus(item.NoteApproversDTO)
          : item.Status,
      statusNumber: item.StatusNumber,
      ApproverDetails: JSON.parse(item.NoteApproversDTO),
      currentApprover: [
        { ...item.CurrentApprover, id: item.CurrentApproverId },
      ],
      ApproverOrder:
        item.CurrentApprover &&
        item.StatusNumber !== "4000" &&
        this._getCurrentApproverDetails(
          item.CurrentApprover,
          item.NoteApproversDTO,
          item.StatusNumber,
          item.CurrentApproverId
        )[0]?.approverOrder,
      ApproverType:
        item.CurrentApprover &&
        item.StatusNumber !== "4000" &&
        this._getCurrentApproverDetails(
          item.CurrentApprover,
          item.NoteApproversDTO,
          item.StatusNumber,
          item.CurrentApproverId
        )[0]?.approverType,
      department: item.Department,

      title: item.Title,
      commentsLog:
        item.NoteApproverCommentsDTO !== null
          ? this._getCommentsData(JSON.parse(item.NoteApproverCommentsDTO))
          : [],
      referredFromDetails:
        item.NoteReferrerDTO !== null
          ? this._getReferedFromAndToDetails(item.NoteReferrerDTO, "from")
          : [],
      refferredToDetails:
        item.NoteReferrerDTO !== null
          ? this._getReferedFromAndToDetails(item.NoteReferrerDTO, "to")
          : [],
      draftResolutionFieldValue: item.DraftResolution,
      noteSecretaryDetails:
        item.NoteSecretaryDTO !== null ? JSON.parse(item.NoteSecretaryDTO) : [],
      noteReferrerDTO:
        item.NoteReferrerDTO !== null ? JSON.parse(item.NoteReferrerDTO) : [],
      noteReferrerCommentsDTO:
        item.NoteReferrerCommentsDTO !== null
          ? JSON.parse(item.NoteReferrerCommentsDTO)
          : [],
      noteATRAssigneeDetails:
        item.NoteATRAssigneeDTO !== null
          ? JSON.parse(item.NoteATRAssigneeDTO)
          : [],
      atrGridData:
        item.NoteATRAssigneeDTO !== null
          ? this._getATRGridData(item.NoteATRAssigneeDTO)
          : [],
      noteATRAssigneeDetailsAllUser:
        item.NoteATRAssigneeDTO !== null
          ? JSON.parse(item.NoteATRAssigneeDTO)
          : [],
      noteMarkedInfoDTOState:
        item.NoteMarkedInfoDTO !== null
          ? await this._getdataofMarkedInfo(
              item.NoteMarkedInfoDTO,
              item.NoteMarkedInfoDTOStringId
            )
          : [],
    });
    return item;
  };

  private _getStatus = (e: any): any => {
    e = JSON.parse(e);
    return e[0].mainStatus;
  };

  private _getReferedFromAndToDetails = (
    commentsData: any,
    typeOfReferee: any
  ): any => {
    commentsData = JSON.parse(commentsData);

    const lenOfCommentData = commentsData.length;
    if (typeOfReferee === "to") {
      return commentsData[lenOfCommentData - 1].referredTo;
    }
    return commentsData[lenOfCommentData - 1].referredFrom;
  };

  private _getCurrentApproverDetails = (
    currentApproverData: any,
    ApproverDetails: any,
    statusNumber: any,
    id: any
  ): any => {
    ApproverDetails = JSON.parse(ApproverDetails);

    if (statusNumber === "4000") {
      return [
        {
          email: currentApproverData.EMail,
          text: currentApproverData.Title,
          id: id,
        },
      ];
    }

    if (currentApproverData) {
      const filterApproverData = ApproverDetails.filter((each: any) => {
        if ((each.email || each.approverEmail) === currentApproverData.EMail) {
          return { ...each, ...currentApproverData };
        }
      });

  
      return filterApproverData;
    }

    return null;
  };

  private _formatDateTime = (date: string | number | Date) => {
    const formattedDate = format(new Date(date), "dd-MMM-yyyy");
    const formattedTime = format(new Date(date), "hh:mm a");
    return `${formattedDate} ${formattedTime}`;
  };

  private _checkRefereeAvailable = (): any => {
    if (this.state.noteReferrerDTO.length > 0) {
      const currrentReferee =
        this.state.noteReferrerDTO[this.state.noteReferrerDTO.length - 1];

      return (
        currrentReferee.referrerEmail === this._currentUserEmail &&
        this.state.statusNumber !== "4900"
      );
    } else {
      return undefined;
    }
  };

  private _checkCurrentUserIs_Approved_Refered_Reject_TheCurrentRequest = ():
    | boolean
    | null => {
    let result: boolean | null = null;
   

    this.state.ApproverDetails.forEach((each: any) => {
      if (
        (each.approverEmail || each.approverEmailName || each.email) ===
          this._currentUserEmail &&
        each.approverOrder === this.state.ApproverOrder
      ) {
        
        switch (this.state.statusNumber) {
          case "9000":
            result = false;
            break;
          case "1000":
          case "2000":
          case "3000":
          case "6000":
          case "4900":
            result = true;
            break;
          case "4000":
          case "5000":
          case "8000":
            result = false;
            break;
          default:
            result = false;
            break;
        }
      }
    });
    
    return result;
  };

  private _getFileObj = (data: any): any => {
    const tenantUrl = window.location.protocol + "//" + window.location.host;

    const formatDateTime = (date: string | number | Date) => {
      const formattedDate = format(new Date(date), "dd-MMM-yyyy");
      const formattedTime = format(new Date(), "hh:mm a");
      return `${formattedDate} ${formattedTime}`;
    };

    const result = formatDateTime(data.TimeCreated);

    const filesObj = {
      name: data.Name,
      content: data,
      index: 0,
      LinkingUri: data.LinkingUri || data.LinkingUrl,
      fileUrl: tenantUrl + data.ServerRelativeUrl,
      ServerRelativeUrl: "",
      isExists: true,
      Modified: "",
      isSelected: false,
      size: parseInt(data.Length),
      type: `application/${data.Name.split(".")[1]}`,
      modifiedBy: data.Author.Title,
      createData: result,
    };

    return filesObj;
  };

  private _getItemDocumentsData = async () => {
    try {
      const folderItemsPdf = await this.props.sp.web
        .getFolderByServerRelativePath(`${this._folderName}/Pdf`)
        .files.select("*")
        .expand("Author", "Editor")()
        .then((res) => res);

      

      if (this.state.statusNumber === "9000") {
      
        const filteredFolderItemsPdf = folderItemsPdf.filter((file) => {
         
          return file.Name.toLowerCase().includes(
            this._folderNameAfterApproved.toLowerCase()
          );
        });
        if (filteredFolderItemsPdf.length > 0) {
          const tempFilesPdf: IFileDetails[] = [];
          filteredFolderItemsPdf.forEach((values) => {
            const fileObj = this._getFileObj(values);
            tempFilesPdf.push(fileObj);

        
            if (!this.state.pdfLink) {
              this.setState({ pdfLink: fileObj.fileUrl });
            }
          });

          
          this.setState({ noteTofiles: tempFilesPdf });
        } else {
          const filteredFolderItemsPdf = folderItemsPdf.filter((file) => {
           
            return !file.Name.toLowerCase().includes(
              this._folderNameAfterApproved.toLowerCase()
            );
          });

          
          const tempFilesPdf: IFileDetails[] = [];
          filteredFolderItemsPdf.forEach((values) => {
            const fileObj = this._getFileObj(values);
            tempFilesPdf.push(fileObj);

           
            if (!this.state.pdfLink) {
              this.setState({ pdfLink: fileObj.fileUrl });
            }
          });

         
          this.setState({ noteTofiles: tempFilesPdf });
        }

       
      } else {
       

        const filteredFolderItemsPdf = folderItemsPdf.filter((file) => {
         
          return !file.Name.toLowerCase().includes(
            this._folderNameAfterApproved.toLowerCase()
          );
        });

       
        const tempFilesPdf: IFileDetails[] = [];
        filteredFolderItemsPdf.forEach((values) => {
          const fileObj = this._getFileObj(values);
          tempFilesPdf.push(fileObj);

          
          if (!this.state.pdfLink) {
            this.setState({ pdfLink: fileObj.fileUrl });
          }
        });

       
        this.setState({ noteTofiles: tempFilesPdf });
      }

      const folderItemsWordDocument = await this.props.sp.web
        .getFolderByServerRelativePath(`${this._folderName}/WordDocument`)
        .files.select("*")
        .expand("Author", "Editor")()
        .then((res) => res);

      const tempFilesWordDocument: IFileDetails[] = [];
      folderItemsWordDocument.forEach((values) => {
        tempFilesWordDocument.push(this._getFileObj(values));
      });

      this.setState({ wordDocumentfiles: tempFilesWordDocument });

      const SupportingDocument = await this.props.sp.web
        .getFolderByServerRelativePath(`${this._folderName}/SupportingDocument`)
        .files.select("*")
        .expand("Author", "Editor")()
        .then((res) => res);

      const tempFilesSupportingDocument: IFileDetails[] = [];
      SupportingDocument.forEach((values) => {
        tempFilesSupportingDocument.push(this._getFileObj(values));
      });

      this.setState({ supportingDocumentfiles: tempFilesSupportingDocument });

      const GistDocument = await this.props.sp.web
        .getFolderByServerRelativePath(`${this._folderName}/GistDocuments`)
        .files.select("*")
        .expand("Author", "Editor")()
        .then((res) => res);

      const tempFilesGistDocument: IFileDetails[] = [];
      GistDocument.forEach((values) => {
        tempFilesGistDocument.push(this._getFileObj(values));
      });

      this.setState({
        secretaryGistDocsList: tempFilesGistDocument,
        secretaryGistDocs: tempFilesGistDocument,
      });
    } catch (e) {
      return e;
    }
  };

  private _onToggleSection = (section: string): void => {
    this.setState((prevState) => ({
      expandSections: {
        [section]: !prevState.expandSections[section],
        ...Object.keys(prevState.expandSections)
          .filter((key) => key !== section)
          .reduce((acc, key) => ({ ...acc, [key]: false }), {}),
      },
    }));
  };

  private _renderTable = (tableData: any[]): JSX.Element => {
    const columns: IColumn[] = [
      {
        key: "column1",
        name: "Column 1",
        fieldName: "column1",
        minWidth: 120,
        maxWidth: 200,
        onRender: (item: any) => <strong>{item.column1}</strong>,
      },
      {
        key: "column2",
        name: "Column 2",
        fieldName: "column2",
        minWidth: 120,
        maxWidth: 200,
        onRender: (item: any) => <span>{item.column2}</span>,
      },
    ];

    return (
      <div>
        <DetailsList
          items={tableData.filter((row) => row.column2 !== undefined)}
          columns={columns}
          setKey="set"
          selectionMode={SelectionMode.none}
          layoutMode={0}
          onRenderDetailsHeader={() => null}
          styles={{
            root: { width: "100%", paddingTop: "4px" },
          }}
        />
      </div>
    );
  };

  private _renderPDFView = (): JSX.Element => {
    return (
      <div style={{ width: "100%" }}>
        <PDFViewer pdfPath={this.state.pdfLink} noteNumber={this.state.title} />
      </div>
    );
  };

  public reOrderData = (reOrderData: any[]): void => {
    this.setState({ peoplePickerData: reOrderData });
  };

  

  private _getAuditTrail = async (status: any) => {
    const item = await this._getItemDataSpList(this._itemId);
    
    const auditTrail = JSON.parse(item.AuditTrail);
    if (status === "gistDocuments") {
      const auditLog = [
        {
          actionBy: this.props.context.pageContext.user.displayName,
          action:
            this.props.formType === "View"
              ? `Gist Documents are updated for Ecommittee Note`
              : `Gist Documents are updated for Board Note`,

          createdDate: this._formatDateTime(new Date()),
        },
      ];

      return JSON.stringify([...auditTrail, ...auditLog]);
    } else {
      const auditLog = [
        {
          actionBy: this.props.context.pageContext.user.displayName,

          action:
            this.props.formType === "View"
              ? `ECommittee Note ${status}`
              : `Board Note ${status}`,

          createdDate: this._formatDateTime(new Date()),
        },
      ];

      return JSON.stringify([...auditTrail, ...auditLog]);
    }
  };

  public async clearFolder(
    libraryName: any,
    folderRelativeUrl: string
  ): Promise<void> {
    try {
      const folder = await this.props.sp.web.getFolderByServerRelativePath(
        folderRelativeUrl
      );

      const items = await folder.files();

      for (const item of items) {
        await this.props.sp.web
          .getFileByServerRelativePath(item.ServerRelativeUrl)
          .recycle();
      }
    } catch (error) {
      return error;
    }
  }

  private async updateGistDocumentFolderItems(
    libraryName: any[],
    folderPath: string,
    type: string
  ) {
    await this.clearFolder(libraryName, folderPath);

    try {
      for (const file of libraryName) {
        const arrayBuffer = await this.getFileArrayBuffer(file);

        await this.props.sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, arrayBuffer, {
            Overwrite: true,
          });
      }
    } catch (error) {
      return error;
    }
  }

  private getFileArrayBuffer = async (file: any): Promise<ArrayBuffer> => {
    if (file.arrayBuffer) {
      return await file.arrayBuffer();
    } else {
      let blob: Blob;
      if (file instanceof Blob) {
        blob = file;
      } else {
        blob = new Blob([file]);
      }

      return new Promise<ArrayBuffer>((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => {
          if (reader.result) {
            resolve(reader.result as ArrayBuffer);
          } else {
            reject(new Error("Failed to read file as ArrayBuffer"));
          }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(blob);
      });
    }
  };

  private async updateSupportingDocumentFolderItems(
    libraryName: any[],
    folderPath: string,
    type: string
  ) {
    try {
      for (const file of libraryName) {
        const arrayBuffer = await this.getFileArrayBuffer(file);

        await this.props.sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, arrayBuffer, {
            Overwrite: true,
          });
      }
    } catch (error) {
      return error;
    }
  }

  private _getCurrentApproverDetailsFromApproverDTO = (): any => {
    const currentApproverDetails = this.state.ApproverDetails.filter(
      (each: any) => each.userId === this.state.currentApprover[0].id
    );

   
    return currentApproverDetails;
  };

  private _updateDefaultNoteATRAssigneeDetails = async (): Promise<any> => {
    const currentAtrCreator = this.state.atrCreatorsList.filter(
      (each: any) =>
        each.atrCreatorEmail === this.props.context.pageContext.user.email
    );

    this._atrJoinedCommentsToDTO();

    const currentApproverDetailsFromApproverDTO =
      this._getCurrentApproverDetailsFromApproverDTO();

    
    

    const defaultNoteATRAssigneeDetails =(stateValue:any)=> [
      {
        atrType: "Default",
        atrAssigneeId: stateValue.createdByID,
        atrCreatorId: currentAtrCreator[0].atrCreatorId,
        atrCreatorEmail: currentAtrCreator[0].atrCreatorEmail,

        atrAssigneeEmailName: stateValue.createdByEmailName,
        atrAssigneeEmail: stateValue.createdByEmail,
        approverEmailName:
          currentApproverDetailsFromApproverDTO.approverEmailName,
        atrCreatorEmailName: currentAtrCreator[0].atrCreatorEmailName,

        createdDate: this._formatDateTime(new Date()),
        createdBy: this.props.context.pageContext.user.email,
        modifiedDate: this._formatDateTime(new Date()),
        modifiedBy: this.props.context.pageContext.user.email,
        statusMessage: null,
        atrId: "",
        noteApproverId: currentApproverDetailsFromApproverDTO.userId,
        approverType: currentApproverDetailsFromApproverDTO.approverType,
        approverOrder: currentApproverDetailsFromApproverDTO.approverOrder,
        approverStatus: 1,
        approverEmail: currentApproverDetailsFromApproverDTO.approverEmail,
        noteApproverComments: this._atrJoinedCommentsToDTO(),
        strATRStatus: "Submitted",
        atrStatus: 1,
        noteId: this._itemId,
      },
    ];
    this.setState((prevState)=>{

     
      
      return {
      noteATRAssigneeDetails: defaultNoteATRAssigneeDetails(prevState),
    }});

    return [
      ...this.state.noteATRAssigneeDetailsAllUser,
      ...defaultNoteATRAssigneeDetails(this.state),
    ];
  };

  private _updateATRRequest = async (currentApproverId: any): Promise<void> => {
    this.state.noteATRAssigneeDetails.map(async (each: any) => {
      
      try {
        const auditLog = [
          {
            actionBy: this.props.context.pageContext.user.displayName,

            action: `ATR Created`,
            createdDate: this._formatDateTime(new Date()),
          },
        ];

        const joinedCommentsData = this.state.generalComments
          .filter((each: any) => !!each)
          .map(
            (each: any) => `${each?.pageNum} ${each?.page} ${each?.comment}`
          );

        const atrObj = {
          Title: this.state.title,
          NoteTo: "",
          Status: "Submitted",
          ATRNoteID: this.state.title,
          Department: this.state.department,
          Subject: this.state.subjectFeildValue,
          AssignedById: each.atrCreatorId,
          Remarks: joinedCommentsData.join(", "),

          AuditTrail: JSON.stringify(auditLog),
          AssigneeId: each.atrAssigneeId,
          StatusNumber: "1000",
          NoteID: `${this._itemId}`,
          CurrentApproverId: each.atrAssigneeId,
          NoteType: this._committeeTypeForATR,
          CommitteeName: this.state.committeeNameFeildValue,
          NoteApproversDTO: JSON.stringify(this.state.ApproverDetails),
          startProcessing: true,
          ATRType: this.state.atrType,
        };

       
        await this.props.sp.web.lists
          .getByTitle("ATRRequests")
          .items.add(atrObj);
      } catch (error) {
        return error;
      }
    });
  };

  private _defaultUserAsATR = async (currentApproverId: any): Promise<any> => {
    let defaultAtrObj = {};

    try {
      const joinedCommentsData = this.state.generalComments
        .filter((each: any) => !!each)
        .map((each: any) => `${each?.pageNum} ${each?.page} ${each?.comment}`);

      const auditLog = [
        {
          actionBy: this.props.context.pageContext.user.displayName,

          action: `ATR Submitted`,
          createdDate: this._formatDateTime(new Date()),
        },
      ];

      defaultAtrObj = {
        Title: this.state.title,
        NoteTo: "",
        Status: "Submitted",
        ATRNoteID: this.state.title,
        Department: this.state.department,
        Subject: this.state.subjectFeildValue,
        AssignedById: [(await this.props.sp?.web.currentUser())?.Id][0],

        Remarks: joinedCommentsData.join(", "),
        AuditTrail: JSON.stringify(auditLog),
        AssigneeId: this.state.createdByID,
        StatusNumber: "1000",
        NoteID: `${this._itemId}`,
        CurrentApproverId: this.state.createdByID,
        NoteType: this._committeeTypeForATR,

        CommitteeName: this.state.committeeNameFeildValue,
        NoteApproversDTO: JSON.stringify(this.state.ApproverDetails),
        startProcessing: true,
        ATRType: "Default",
      };

    

      await this.props.sp.web.lists
        .getByTitle("ATRRequests")
        .items.add(defaultAtrObj);
    } catch (error) {
      return error;
    }

    return defaultAtrObj;
  };

  private _getCurrentApproverDetailsInHandleApprover = (
    modifyApproveDetails: any
  ): any => {
    const currentApproverdata = modifyApproveDetails.filter((each: any) => {
      if (each.status === "Pending") {
        return each;
      }
    });


    return currentApproverdata[0];
  };

  private _modifiedApproverInHandleApprover = (
    _ApproverDTO: any,
    statusFromEvent: any
  ): any => {
    let previousApprover: any;
    const modifyApproveDetails = _ApproverDTO.map(
      (each: any, index: number) => {
        if (
          each.approverEmail === this._currentUserEmail ||
          each.email === this._currentUserEmail
        ) {
          previousApprover = [
            {
              ...each,
              status: statusFromEvent,
              actionDate: this._formatDateTime(new Date()),
              mainStatus: "Approved",
              statusNumber: "9000",
            },
          ];

          return {
            ...each,
            status: statusFromEvent,
            actionDate: this._formatDateTime(new Date()),
            mainStatus: "Approved",
            statusNumber: "9000",
          };
        }

        if (each.approverOrder === this.state.ApproverOrder + 1) {
          return {
            ...each,
            status: "Pending",
            mainStatus:
              each.approverType === "Approver"
                ? "Pending with approver"
                : "Pending with reviewer",
            statusNumber: each.approverType === "Approver" ? "3000" : "2000",
          };
        }
        return each;
      }
    );

    return [previousApprover, modifyApproveDetails];
  };

  private _checkCurrentUserInApproverDto = (
    _ApproverInfoDTOId: any,
    currentUserId: any
  ) => {
    if (!_ApproverInfoDTOId.includes(currentUserId)) {
      this.setState({
        isLoading: false,
        hideParellelActionAlertDialog: true,
        parellelActionAlertMsg:
          "This request has been taken action by approver.",
      });

      return null;
    }
  };

  private _updateItemInHandleApproverBtn = async (
    modifyApproveDetails: any,
    currentApproverDetail: any,
    previousApprover: any,
    updateNoteATRAssigneeDTO: any,
    updateAuditTrial: any,
    _CommentsLog: any
  ): Promise<any> => {
    let noteATRAssigneeDTO;

    
    if (this._checkCurrentUserIsAATRAssignee()) {
      if (this.state.atrGridData.length > 0) {
        noteATRAssigneeDTO = JSON.stringify([
          ...this.state.noteATRAssigneeDetailsAllUser,
          ...updateNoteATRAssigneeDTO,
        ]);
      } else {
        const defaultDetails =
          await this._updateDefaultNoteATRAssigneeDetails();
        noteATRAssigneeDTO = JSON.stringify(defaultDetails);
      }
    } else {
      noteATRAssigneeDTO = JSON.stringify([
        ...this.state.noteATRAssigneeDetailsAllUser,
        ...updateNoteATRAssigneeDTO,
      ]);
    }

    return {
      NoteApproversDTO: JSON.stringify(modifyApproveDetails),
      Status: currentApproverDetail?.mainStatus,
      StatusNumber: currentApproverDetail?.statusNumber,
      AuditTrail: updateAuditTrial,
      NoteApproverCommentsDTO: JSON.stringify([
        ..._CommentsLog,
        ...this.state.generalComments,
      ]),
      CurrentApproverId:
        this.state.ApproverOrder === modifyApproveDetails.length
          ? null
          : currentApproverDetail.userId,
      PreviousApproverId: previousApprover[0].userId,

      NoteATRAssigneeDTO: noteATRAssigneeDTO,
      PreviousActionerId: [(await this.props.sp?.web.currentUser())?.Id],
      startProcessing: true,
    };
  };

  private _EndStatusNumberUpdate = async (
    statusFromEvent: any,
    statusNumber: any
  ): Promise<any> => {
    if (this.state.ApproverDetails.length === this.state.ApproverOrder) {
      this.setState({ status: statusFromEvent });
      await this.props.sp.web.lists
        .getByTitle(this._listname)
        .items.getById(this._itemId)
        .update({
          Status: statusFromEvent,
          StatusNumber: statusNumber,
        });
    }
  };

  private _handleApproverButton = async (
    statusFromEvent: string,
    statusNumber: string
  ) => {
    const item = await this._getItemDataSpList(this._itemId);
 
    const StatusNumber = item?.StatusNumber;
    this._closeDialog();

    this.setState({ isLoading: true });

    const checkCurrentApproverIsCurrentUser = item?.CurrentApproverId;

    const currentUserId = (await this.props.sp?.web.currentUser())?.Id;
    const _ApproverDTO = JSON.parse(item?.NoteApproversDTO);
    const _ApproverInfoDTOId = JSON.parse(item?.NoteApproversDTO).map(
      (each: any) => each.userId
    );
    const _CommentsLog =
      item.NoteApproverCommentsDTO !== null
        ? this._getCommentsData(JSON.parse(item.NoteApproverCommentsDTO))
        : [];
    if (
      StatusNumber !== "200" &&
      currentUserId === checkCurrentApproverIsCurrentUser
    ) {
      this._checkCurrentUserInApproverDto(_ApproverInfoDTOId, currentUserId);

      const previousApprover = this._modifiedApproverInHandleApprover(
        _ApproverDTO,
        statusFromEvent
      )[0];
      const modifyApproveDetails = this._modifiedApproverInHandleApprover(
        _ApproverDTO,
        statusFromEvent
      )[1];

      const currentApproverDetail =
        this._getCurrentApproverDetailsInHandleApprover(modifyApproveDetails);
      const currentApproverId =
        this.state.ApproverOrder === modifyApproveDetails.length
          ? null
          : currentApproverDetail.id;
      const updateNoteATRAssigneeDTO = this.state.noteATRAssigneeDetails.map(
        (each: any) => {
          return {
            ...each,
            noteApproverComments: this._atrJoinedCommentsToDTO(),
          };
        }
      );
      try {
        const updateAuditTrial = await this._getAuditTrail(
          this._checkCurrentUserIsAATRAssignee() ? "Noted" : "Approved"
        );

        const updateItems = await this._updateItemInHandleApproverBtn(
          modifyApproveDetails,
          currentApproverDetail,
          previousApprover,
          updateNoteATRAssigneeDTO,
          updateAuditTrial,
          _CommentsLog
        );
        console.log(updateItems);

        await this.props.sp.web.lists
          .getByTitle(this._listname)
          .items.getById(this._itemId)
          .update(updateItems);

        this._checkCurrentUserIsAATRAssignee() &&
          (this.state.atrGridData.length > 0
            ? await this._updateATRRequest(currentApproverId)
            : await this._defaultUserAsATR(currentApproverId));

        await this.updateSupportingDocumentFolderItems(
          this.state.supportingFilesInViewForm,
          `${this._folderName}/SupportingDocument`,
          "Supporting documents"
        );

        this._EndStatusNumberUpdate(statusFromEvent, statusNumber);

        this.setState({ isLoading: false, isVisibleAlter: true });
      } catch (error) {
        return error;
      }
    } else {
      this.setState({
        isLoading: false,
        hideParellelActionAlertDialog: true,
        parellelActionAlertMsg:
          currentUserId !== checkCurrentApproverIsCurrentUser
            ? "This request has been taken action by approver."
            : "This request has been call back",
      });
    }
  };

  private _checkingCurrentUserInSecretaryDTO = (): any => {
    const checkCurrentUserIsAnApprover = this.state.ApproverDetails.filter(
      (each: any) => each.secretaryEmail && each.approverType === "Approver"
    );
    const currentUserIsFromSecDTOAndHeIsSECOrApp =
      this.state.noteSecretaryDetails.some((each: any) => {
        if (
          each.secretaryEmail === this._currentUserEmail ||
          each.approverEmail === this._currentUserEmail
        ) {
          return true;
        }
      });

    return (
      checkCurrentUserIsAnApprover && currentUserIsFromSecDTOAndHeIsSECOrApp
    );
  };

  private _checkingCurrentUserInSecretaryDTOAfterApproved = (): any => {
    const currentUserIsFromSecDTOAndHeIsSECOrApp =
      this.state.noteSecretaryDetails.some((each: any) => {
        if (each.secretaryEmail === this._currentUserEmail) {
          return true;
        }
      });

    return currentUserIsFromSecDTOAndHeIsSECOrApp;
  };

  private _checkingCurrentUserAsApproverDTOInSecretaryDTO = (): any => {
    const checkCurrentUserIsAnApprover = this.state.ApproverDetails.filter(
      (each: any) => each.secretaryEmail && each.approverType === "Approver"
    );

    const currentUserIsFromSecDTOAndHeIsSECOrApp =
      this.state.noteSecretaryDetails.some((each: any) => {
        if (each.approverEmail === this._currentUserEmail) {
          return true;
        }
      });

    return (
      checkCurrentUserIsAnApprover && currentUserIsFromSecDTOAndHeIsSECOrApp
    );
  };

  private _checkingCurrentUserIsSecretaryDTO = (): any => {
    const currentUserHavingSecretaryisApproved =
      this.state.ApproverDetails.filter((each: any) => {
        if (
          each.secretary === this.props.context.pageContext.user.displayName &&
          each.statusNumber !== "9000" &&
          each.approverType === "Approver"
        ) {
          return each;
        }
      });

    const filterAllApproverMailHavingSec =
      currentUserHavingSecretaryisApproved.map(
        (each: any) => each.approverEmail
      );

    const checkCurrentUserISanApprover =
      this.state.currentApprover?.length > 0 &&
      filterAllApproverMailHavingSec.includes(
        this.state.currentApprover[0]?.EMail
      );

    const userIsSec = this.state.noteSecretaryDetails.some((each: any) => {
      if (each.secretaryEmail === this._currentUserEmail) {
        return true;
      }
    });

    return (
      userIsSec &&
      currentUserHavingSecretaryisApproved.length > 0 &&
      checkCurrentUserISanApprover
    );
  };

  private handleReject = async (
    statusFromEvent: string,
    statusNumber: string
  ) => {
    this._closeDialog();
    this.setState({ isLoading: true });
    const item = await this._getItemDataSpList(this._itemId);
    const StatusNumber = item?.StatusNumber;

    const checkCurrentApproverIsCurrentUser = item?.CurrentApproverId;
   

    const currentUserId = (await this.props.sp?.web.currentUser())?.Id;

   
    const _ApproverDTO = JSON.parse(item?.NoteApproversDTO);
    const _CommentsLog =
      item.NoteApproverCommentsDTO !== null
        ? this._getCommentsData(JSON.parse(item.NoteApproverCommentsDTO))
        : [];
    const _ApproverInfoDTOId = JSON.parse(item?.NoteApproversDTO).map(
      (each: any) => each.userId
    );

  

    if (StatusNumber === "8000") {
      this.setState({
        isLoading: false,
        hideParellelActionAlertDialog: true,
        parellelActionAlertMsg:
          "This request has been taken action by approver.",
      });

      return;
    }

    if (
      StatusNumber !== "200" &&
      currentUserId === checkCurrentApproverIsCurrentUser
    ) {
      if (!_ApproverInfoDTOId.includes(currentUserId)) {
        this.setState({
          isLoading: false,
          hideParellelActionAlertDialog: true,
          parellelActionAlertMsg:
            "This request has been taken action by approver.",
        });

        return;
      }
      const modifyApproveDetails = _ApproverDTO.map(
        (each: any, index: number) => {
          if (each.approverEmail === this._currentUserEmail) {
            return {
              ...each,
              status: statusFromEvent,
              actionDate: this._formatDateTime(new Date()),
              mainStatus: statusFromEvent,
              statusNumber: statusNumber,
            };
          }

          return each;
        }
      );

      const updateAuditTrial = await this._getAuditTrail(statusFromEvent);

      await this.props.sp.web.lists
        .getByTitle(this._listname)
        .items.getById(this._itemId)
        .update({
          NoteApproversDTO: JSON.stringify(modifyApproveDetails),
          Status: statusFromEvent,
          StatusNumber: statusNumber,
          AuditTrail: updateAuditTrial,
          NoteApproverCommentsDTO: JSON.stringify([
            ..._CommentsLog,
            ...this.state.generalComments,
          ]),

          PreviousActionerId: [(await this.props.sp?.web.currentUser())?.Id],
          startProcessing: true,
        });

      await this.updateSupportingDocumentFolderItems(
        this.state.supportingFilesInViewForm,
        `${this._folderName}/SupportingDocument`,
        "Supporting documents"
      );

      if (this.state.ApproverDetails.length === this.state.ApproverOrder) {
        this.setState({ status: statusFromEvent });
        await this.props.sp.web.lists
          .getByTitle(this._listname)
          .items.getById(this._itemId)
          .update({
            Status: statusFromEvent,
            StatusNumber: statusNumber,
          });

        await this.updateSupportingDocumentFolderItems(
          this.state.supportingFilesInViewForm,
          `${this._folderName}/SupportingDocument`,
          "Supporting documents"
        );
      }

      this.setState({ isVisibleAlter: true, isLoading: false });
    } else {
      this.setState({
        isLoading: false,
        hideParellelActionAlertDialog: true,
        parellelActionAlertMsg:
          currentUserId !== checkCurrentApproverIsCurrentUser
            ? "This request has been taken action by approver."
            : "This request has been call back",
      });
    }
  };

  private referPassCodeTrigger = (): any => {
    if (!this.state.isPasscodeValidated) {
      this.setState({
        isPasscodeModalOpen: true,
        passCodeValidationFrom: "4000",
        dialogFluent: true,
      });
      return null;
    }
  };

  private changeApproverPassCodeTrigger = (): any => {
    if (!this.state.isPasscodeValidated) {
      this.setState({
        isPasscodeModalOpen: true,
        passCodeValidationFrom: "7500",
        dialogFluent: true,
      });
      return null;
    }
  };

  private _referCommentsAndDataMandatory = (): any => {
    this.setState({ dialogFluent: true, isReferDataAndCommentsNeeded: true });
  };

  private handleRefer = async (
    statusFromEvent: string,
    statusNumber: string,
    commentsObj: any
  ) => {
    console.log(commentsObj);
    console.log(this.state);
    this._closeDialog();
    this.setState({ isLoading: true });
    const item = await this._getItemDataSpList(this._itemId);
    const StatusNumber = item?.StatusNumber;
    const _ApproverDTO = JSON.parse(item?.NoteApproversDTO);
    const _ReferDTO =
      item.NoteReferrerDTO !== null ? JSON.parse(item.NoteReferrerDTO) : [];

    const _CommentsLog =
      item.NoteApproverCommentsDTO !== null
        ? this._getCommentsData(JSON.parse(item.NoteApproverCommentsDTO))
        : [];

    const checkCurrentApproverIsCurrentUser = item?.CurrentApproverId;
   

    const currentUserId = (await this.props.sp?.web.currentUser())?.Id;
    const _ApproverInfoDTOId = JSON.parse(item?.NoteApproversDTO).map(
      (each: any) => each.userId
    );

   

    if (
      StatusNumber !== "200" &&
      currentUserId === checkCurrentApproverIsCurrentUser
    ) {
      if (!_ApproverInfoDTOId.includes(currentUserId)) {
        this.setState({
          isLoading: false,
          hideParellelActionAlertDialog: true,
          parellelActionAlertMsg:
            "This request has been taken action by approver.",
        });

        return;
      }

      const modifyApproveDetails = _ApproverDTO.map(
        (each: any, index: number) => {
          if (
            (each.approverEmail || each.approverEmailName) ===
            this._currentUserEmail
          ) {
            return {
              ...each,
              status: statusFromEvent,
              statusNumber,
              actionDate: this._formatDateTime(new Date()),
            };
          }

          return each;
        }
      );

      const updateAuditTrial = await this._getAuditTrail(statusFromEvent);
      const referedId = v4();

      const obj = {
        NoteApproversDTO: JSON.stringify(modifyApproveDetails),
        Status: statusFromEvent,
        StatusNumber: statusNumber,
        AuditTrail: updateAuditTrial,
        NoteApproverCommentsDTO: JSON.stringify([
          ..._CommentsLog,
          ...this.state.generalComments,
        ]),

        CurrentApproverId: this.state.refferredToDetails[0].id,
        PreviousActionerId: [(await this.props.sp?.web.currentUser())?.Id],

        startProcessing: true,
        NoteReferrerDTO: JSON.stringify([
          ..._ReferDTO,
          {
            approverEmail:
              this.state.referredFromDetails[0].EMail ||
              this.state.referredFromDetails[0].approverEmail,
            approverEmailName:
              this.state.referredFromDetails[0].Text ||
              this.state.referredFromDetails[0].Title ||
              this.state.referredFromDetails[0].approverEmailName,
            approverType: this.state.referredFromDetails[0].approverType,
            createdBy:
              this.state.referredFromDetails[0].EMail ||
              this.state.referredFromDetails[0].approverEmail,
            createdDate: this._formatDateTime(new Date()),
            modifiedBy:
              this.state.referredFromDetails[0].EMail ||
              this.state.referredFromDetails[0].approverEmail,
            modifiedDate: this._formatDateTime(new Date()),
            noteApproverId: this.state.referredFromDetails[0].id,
            noteId: this._itemId,
            noteReferrerId: referedId,
            referrerId: this.state.refferredToDetails[0].id,
            noteSupportingDocumentsDTO: null,
            referrerEmail:
              this.state.refferredToDetails[0].email ||
              this.state.refferredToDetails[0].approverEmail,
            referrerEmailName:
              this.state.refferredToDetails[0].text ||
              this.state.refferredToDetails[0].approverEmailName,
            referrerStatus: 1,
            referrerStatusType: this.state.refferredToDetails[0].status,
            comments: this.state.referComment.comment,
          },
        ]),
      };
     

      await this.props.sp.web.lists
        .getByTitle(this._listname)
        .items.getById(this._itemId)
        .update(obj);

      await this.updateSupportingDocumentFolderItems(
        this.state.supportingFilesInViewForm,
        `${this._folderName}/SupportingDocument`,
        "Supporting documents"
      );

      if (this.state.ApproverDetails.length === this.state.ApproverOrder) {
        this.setState({ status: statusFromEvent });
        await this.props.sp.web.lists
          .getByTitle(this._listname)
          .items.getById(this._itemId)
          .update({
            Status: statusFromEvent,
            StatusNumber: statusNumber,
          });
      }
      this.setState({ isVisibleAlter: true, isLoading: false });
    } else {
      this.setState({
        isLoading: false,
        hideParellelActionAlertDialog: true,
        parellelActionAlertMsg:
          currentUserId !== checkCurrentApproverIsCurrentUser
            ? "This request has been taken action by approver."
            : "This request has been call back",
      });
    }
  };

  private handleReferBack = async (
    statusFromEvent: string,
    statusNumber: string,
    commentsObj: any
  ) => {
    this._closeDialog();
    this.setState({ isLoading: true });
    const item = await this._getItemDataSpList(this._itemId);
    const StatusNumber = item?.StatusNumber;

    const _ApproverDTO = JSON.parse(item?.NoteApproversDTO);
    const _CommentsLog =
      item.NoteApproverCommentsDTO !== null
        ? this._getCommentsData(JSON.parse(item.NoteApproverCommentsDTO))
        : [];

    const checkCurrentApproverIsCurrentUser = item?.CurrentApproverId;
   

    const currentUserId = (await this.props.sp?.web.currentUser())?.Id;

   
    if (currentUserId !== checkCurrentApproverIsCurrentUser) {
      this.setState({
        hideParellelActionAlertDialog: true,
        parellelActionAlertMsg:
          "This request has been taken action by approver",
      });

      return;
    }

    const _ApproverInfoDTOId = JSON.parse(item?.NoteApproversDTO).map(
      (each: any) => each.userId
    );

   

    if (_ApproverInfoDTOId.includes(currentUserId)) {
      this.setState({
        isLoading: false,
        hideParellelActionAlertDialog: true,
        parellelActionAlertMsg:
          "This request has been taken action by approver.",
      });

      return;
    }

    if (StatusNumber !== "200") {
      this.setState({ isLoading: true });
      let currentApproverId = null;

      const modifyApproveDetails = _ApproverDTO.map(
        (each: any, index: number) => {
          if (each.statusNumber === "4000") {
            if (each.approverType === "Reviewer") {
              currentApproverId = each.userId;

              return {
                ...each,
                status: "Pending",
                statusNumber: "2000",
                actionDate: this._formatDateTime(new Date()),
              };
            } else {
              currentApproverId = each.userId;

              return {
                ...each,
                status: "Pending",
                statusNumber: "3000",
                actionDate: this._formatDateTime(new Date()),
              };
            }
          }

          return each;
        }
      );

      const updateAuditTrial = await this._getAuditTrail(statusFromEvent);

      const obj = {
        NoteApproversDTO: JSON.stringify(modifyApproveDetails),
        Status: statusFromEvent,
        StatusNumber: statusNumber,
        CurrentApproverId: currentApproverId,
        AuditTrail: updateAuditTrial,
        NoteApproverCommentsDTO: JSON.stringify([
          ..._CommentsLog,
          ...this.state.generalComments,
        ]),
        NoteReferrerCommentsDTO: JSON.stringify(
          this.state.noteReferrerCommentsDTO
        ),

        startProcessing: true,
        PreviousActionerId: [(await this.props.sp?.web.currentUser())?.Id],
      };

      await this.props.sp.web.lists
        .getByTitle(this._listname)
        .items.getById(this._itemId)
        .update(obj);

      await this.updateSupportingDocumentFolderItems(
        this.state.supportingFilesInViewForm,
        `${this._folderName}/SupportingDocument`,
        "Supporting documents"
      );

      this.setState({ isVisibleAlter: true, isLoading: false });
    } else {
      this.setState({
        isLoading: false,
        hideParellelActionAlertDialog: true,
        parellelActionAlertMsg: "This request has been call back",
      });
    }
  };

  private handleReturn = async (
    statusFromEvent: string,
    statusNumber: string
  ) => {
    this._closeDialog();
    this.setState({ isLoading: true });

    const item = await this._getItemDataSpList(this._itemId);
    const StatusNumber = item?.StatusNumber;

    const _ApproverDTO = JSON.parse(item?.NoteApproversDTO);
    const _CommentsLog =
      item.NoteApproverCommentsDTO !== null
        ? this._getCommentsData(JSON.parse(item.NoteApproverCommentsDTO))
        : [];

    const checkCurrentApproverIsCurrentUser = item?.CurrentApproverId;
   

    const currentUserId = (await this.props.sp?.web.currentUser())?.Id;

    

    const _ApproverInfoDTOId = JSON.parse(item?.NoteApproversDTO).map(
      (each: any) => each.userId
    );

    

    if (
      StatusNumber !== "200" &&
      currentUserId === checkCurrentApproverIsCurrentUser
    ) {
      if (!_ApproverInfoDTOId.includes(currentUserId)) {
        this.setState({
          isLoading: false,
          hideParellelActionAlertDialog: true,
          parellelActionAlertMsg:
            "This request has been taken action by approver.",
        });

        return;
      }

      const modifyApproveDetails = _ApproverDTO.map(
        (each: any, index: number) => {
          if (each.approverEmail === this._currentUserEmail) {
            return {
              ...each,
              status: statusFromEvent,
              statusNumber: "5000",
              actionDate: this._formatDateTime(new Date()),
            };
          }

          if (each.approverOrder === this.state.ApproverOrder + 1) {
            return { ...each, status: "Pending" };
          }
          return each;
        }
      );

      const updateAuditTrial = await this._getAuditTrail(statusFromEvent);

      await this.props.sp.web.lists
        .getByTitle(this._listname)
        .items.getById(this._itemId)
        .update({
          NoteApproversDTO: JSON.stringify(modifyApproveDetails),
          CurrentApproverId: this.state.createdByID,

          NoteATRAssigneeDTO: JSON.stringify([]),
          Status: statusFromEvent,
          StatusNumber: statusNumber,
          AuditTrail: updateAuditTrial,
          NoteApproverCommentsDTO: JSON.stringify([
            ..._CommentsLog,
            ...this.state.generalComments,
          ]),

          startProcessing: true,
          PreviousActionerId: [(await this.props.sp?.web.currentUser())?.Id],
        });

      await this.updateSupportingDocumentFolderItems(
        this.state.supportingFilesInViewForm,
        `${this._folderName}/SupportingDocument`,
        "Supporting documents"
      );

      if (this.state.ApproverDetails.length === this.state.ApproverOrder) {
        this.setState({ status: statusFromEvent });
        await this.props.sp.web.lists
          .getByTitle(this._listname)
          .items.getById(this._itemId)
          .update({
            Status: statusFromEvent,
            StatusNumber: statusNumber,
          });
      }
      this.setState({ isVisibleAlter: true, isLoading: false });
    } else {
      this.setState({
        isLoading: false,
        hideParellelActionAlertDialog: true,
        parellelActionAlertMsg:
          currentUserId !== checkCurrentApproverIsCurrentUser
            ? "This request has been taken action by approver."
            : "This request has been call back",
      });
    }
  };

  private handleCallBack = async (
    statusFromEvent: string,
    statusNumber: string
  ) => {
    this._closeDialog();
    this.setState({ isLoading: true });

    const item = await this._getItemDataSpList(this._itemId);
    const StatusNumber = item?.StatusNumber;

    const checkCurrentApproverIsCurrentUser = item?.CurrentApproverId;
 

    const _ApproverInfoDTOId = JSON.parse(item?.NoteApproversDTO).map(
      (each: any) => each.userId
    );

    

    const _NoteReferrerDTO =
      item?.NoteReferrerDTO !== null ? JSON.parse(item?.NoteReferrerDTO) : [];
    const _refereeId =
      _NoteReferrerDTO.length > 0
        ? _NoteReferrerDTO[_NoteReferrerDTO.length - 1].referrerId
        : null;
   
    const _actionersIDs = [..._ApproverInfoDTOId, _refereeId];
   
    const actionTakenOcurredOrNot = JSON.parse(item?.NoteApproversDTO)[0]
      .actionDate;

    if (actionTakenOcurredOrNot !== "") {
      if (_actionersIDs.includes(checkCurrentApproverIsCurrentUser)) {
        this.setState({
          isLoading: false,
          hideParellelActionAlertDialog: true,
          parellelActionAlertMsg:
            "This request has been taken action by approver.",
        });

        return;
      }
    }

    if (StatusNumber === "200") {
      this.setState({
        isLoading: false,
        hideParellelActionAlertDialog: true,
        parellelActionAlertMsg: "This request has been called back.",
      });

      return;
    }

    const _ApproverInfoDTO = JSON.parse(item?.NoteApproversDTO);
    if (
      _ApproverInfoDTO?.every(
        (obj: any) => obj.status === "Pending" || obj.status === "Waiting"
      )
    ) {
      this.setState({ isLoading: true });
      const updateAuditTrial = await this._getAuditTrail(statusFromEvent);

      await this.props.sp.web.lists
        .getByTitle(this._listname)
        .items.getById(this._itemId)
        .update({
          startProcessing: true,
          CurrentApproverId: this.state.createdByID,
          Status: statusFromEvent,
          StatusNumber: statusNumber,
          AuditTrail: updateAuditTrial,
          PreviousActionerId: [(await this.props.sp?.web.currentUser())?.Id],
        });
      this.setState({ isVisibleAlter: true, isLoading: false });
    } else {
      this.setState({
        hideParellelActionAlertDialog: true,
        parellelActionAlertMsg:
          "This request has been taken action by approver",
      });
    }
  };

  private _getNoteMarkedId = (): any => {
    const ids = this.state.noteMarkedInfoDTOState.map((each: any) => {
      return each.id;
    });

    return ids;
  };

  private _handleMarkInfoSubmit = async (): Promise<any> => {
    this.setState({ isLoading: true });
    const updateAuditTrial = await this._getAuditTrail("Mark Info Added");
    await this.props.sp.web.lists
      .getByTitle(this._listname)
      .items.getById(this._itemId)
      .update({
        NoteMarkedInfoDTOId: this._getNoteMarkedId(),
        AuditTrail: updateAuditTrial,
        PreviousActionerId: [(await this.props.sp?.web.currentUser())?.Id],
      });

    this.setState({ isLoading: false });
  };

  private _changeApproverDataMandatory = (): any => {
    if (
      this.state.peoplePickerSelectedDataWhileReferOrChangeApprover.length === 0
    ) {
      this.setState({ dialogFluent: true, isChangeApproverNeeded: true });
    }
  };

  private handleChangeApprover = async (
    statusFromEvent: string,
    statusNumber: string,
    data: any
  ) => {
    this._closeDialog();
    this.setState({ isLoading: true });
    const item = await this._getItemDataSpList(this._itemId);
    const _ApproverDTO = JSON.parse(item?.NoteApproversDTO);
   
    const _ReferDTO =
      item.NoteReferrerDTO !== null ? JSON.parse(item.NoteReferrerDTO) : [];

    const StatusNumber = item?.StatusNumber;

    if (
      item?.StatusNumber !== "9000" &&
      item?.StatusNumber !== "5000" &&
      item?.StatusNumber !== "8000" &&
      item?.StatusNumber !== "200"
    ) {
      if (this.state.statusNumber === "4000") {
        const updateAuditTrial = await this._getAuditTrail(statusFromEvent);

        const updateLastNoteReferDTO = {
          ..._ReferDTO[_ReferDTO.length - 1],
          referrerEmail:
            this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0]
              .email,
          referrerEmailName:
            this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0]
              .text,
          referrerId:
            this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0].id,
        };

        const updateNoteReferDTO = _ReferDTO.map((each: any, index: any) => {
          if (each.noteReferrerId === updateLastNoteReferDTO.noteReferrerId) {
            return {
              ...each,
              referrerEmail:
                this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0]
                  .email,
              referrerEmailName:
                this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0]
                  .text,
              referrerId:
                this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0]
                  .id,
            };
          }
          return each;
        });

      

        await this.props.sp.web.lists
          .getByTitle(this._listname)
          .items.getById(this._itemId)
          .update({
            startProcessing: true,
            CurrentApproverId:
              this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0]
                .id,
            AuditTrail: updateAuditTrial,
            NoteReferrerDTO: JSON.stringify(updateNoteReferDTO),
            PreviousActionerId: [(await this.props.sp?.web.currentUser())?.Id],
          });

        this.setState({ isVisibleAlter: true, isLoading: false });

        return;
      }

      const checkSelectedApproverHasSecretary =
        this.state.approverIdsHavingSecretary.filter(
          (each: any) =>
            each.ApproverId ===
            this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0].id
        );

      const secretaryObj = {
        noteSecretarieId:
          checkSelectedApproverHasSecretary[0]?.noteSecretarieId,
        noteApproverId: checkSelectedApproverHasSecretary[0]?.noteApproverId,
        noteId: this._itemId,
        secretaryEmail: checkSelectedApproverHasSecretary[0]?.secretaryEmail,
        approverEmail: checkSelectedApproverHasSecretary[0]?.approverEmail,
        approverEmailName:
          checkSelectedApproverHasSecretary[0]?.approverEmailName,
        secretaryEmailName:
          checkSelectedApproverHasSecretary[0]?.secretaryEmailName,
        createdBy: "",
        modifiedDate: "",
        modifiedBy: "",
      };

      const updateCurrentApprover = (): any => {
        const upatedCurrentApprover = _ApproverDTO.filter((each: any) => {
         

          if (each.status === "Pending") {
            return {
              ...this.state.peoplePickerSelectedDataWhileReferOrChangeApprover,
              status: "Pending",
              actionDate: this._formatDateTime(new Date()),
              mainStatus: each.mainStatus,
              secretary:
                checkSelectedApproverHasSecretary.length > 0
                  ? checkSelectedApproverHasSecretary[0].secretaryEmailName
                  : "",
              secretaryEmail:
                checkSelectedApproverHasSecretary.length > 0
                  ? checkSelectedApproverHasSecretary[0].secretaryEmail
                  : "",
            };
          }
        });

        return [
          {
            approverType: upatedCurrentApprover[0].approverType,
            approverEmail:
              this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0]
                .email ||
              this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0]
                .secondaryText,
            approverOrder: upatedCurrentApprover[0].approverOrder,
            approverStatus: upatedCurrentApprover[0].approverStatus,

            srNo: this.state
              .peoplePickerSelectedDataWhileReferOrChangeApprover[0].srNo,
            designation:
              this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0]
                .optionalText,
            approverEmailName:
              this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0]
                .text,
            userId:
              this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0]
                .id,
            status: "Pending",
            statusNumber: upatedCurrentApprover[0].statusNumber,

            mainStatus: upatedCurrentApprover[0].mainStatus,
            actionDate: upatedCurrentApprover[0].actionDate,

            secretary:
              checkSelectedApproverHasSecretary.length > 0
                ? checkSelectedApproverHasSecretary[0].secretaryEmailName
                : "",
            secretaryEmail:
              checkSelectedApproverHasSecretary.length > 0
                ? checkSelectedApproverHasSecretary[0].secretaryEmail
                : "",
          },
        ];
      };

      const modifyApproverDetails = _ApproverDTO.map((each: any) => {
        if (each.status === "Pending") {
          return { ...updateCurrentApprover()[0] };
        } else {
          return each;
        }
      });
     

      const reviewerIds = modifyApproverDetails
        .filter((each: any) => each.approverType === "Reviewer")
        .map((each: any) => each.userId);
      const approverId = modifyApproverDetails
        .filter((each: any) => each.approverType === "Approver")
        .map((each: any) => each.userId);
      const currentApproverId =
        this.state.peoplePickerSelectedDataWhileReferOrChangeApprover[0].id;

      const updateAuditTrial = await this._getAuditTrail(statusFromEvent);

      await this.props.sp.web.lists
        .getByTitle(this._listname)
        .items.getById(this._itemId)
        .update({
          startProcessing: true,
          CurrentApproverId: currentApproverId,
          AuditTrail: updateAuditTrial,
          NoteApproversDTO: JSON.stringify(modifyApproverDetails),
          PreviousActionerId: [(await this.props.sp?.web.currentUser())?.Id],
          FinalApproverId:
            modifyApproverDetails[modifyApproverDetails.length - 1].id,
          NoteSecretaryDTO:
            checkSelectedApproverHasSecretary.length > 0
              ? JSON.stringify([
                  ...this.state.noteSecretaryDetails,
                  secretaryObj,
                ])
              : JSON.stringify([...this.state.noteSecretaryDetails]),
          ReviewersId: reviewerIds,
          ApproversId: approverId,
        });

      this.setState({ isVisibleAlter: true, isLoading: false });

      checkSelectedApproverHasSecretary.length > 0 &&
        this.setState((prevState) => ({
          noteSecretaryDetails: [
            ...prevState.noteSecretaryDetails,
            secretaryObj,
          ],
        }));
    } else {
      this.setState({
        isLoading: false,
        hideParellelActionAlertDialog: true,
        parellelActionAlertMsg:
          this.getMessageBasedOnStatusNumber(StatusNumber),
      });
    }
  };

  private _checkApproveredStatusIsFound = (): any => {
    const checkActionDateIsUpdated = this.state.ApproverDetails.some(
      (each: any) => each.actionDate !== ""
    );

   

    return checkActionDateIsUpdated;
  };

  private _getApproverAndReviewerStageButton = (): any => {
    return (
      <div className={styles.approveEtcBtns}>
        <PrimaryButton
          className={`${styles.responsiveButton}`}
          iconProps={{ iconName: "EditNote" }}
          styles={{
            root: {
              border: "none",
            },
            rootHovered: {
              border: "none",
            },
            rootPressed: {
              border: "none",
            },
          }}
          onClick={
            this._checkCurrentUserIsAATRAssignee() &&
            this._checkCurrentUserIsApproverType()
              ? (e) => {
                  this.setState({ successStatus: "noted" });

                  if (this.state.errorForCummulative) {
                    this.setState({ dialogboxForCummulativeError: true });
                    return;
                  }

                  if (this.state.errorOfDocuments) {
                    this.setState({ isAutoSaveFailedDialog: true });
                  } else if (this.state.generalComments.length === 0) {
                    this.setState({
                      isNotedCommentsManidatoryAlterDialog: true,
                    });
                  } else {
                    this.setState({
                      isPasscodeModalOpen: true,
                      passCodeValidationFrom: "9000",
                    });
                  }
                }
              : (e) => {
                  if (this.state.errorForCummulative) {
                    this.setState({ dialogboxForCummulativeError: true });
                    return;
                  }
                  if (this.state.errorOfDocuments) {
                    this.setState({ isAutoSaveFailedDialog: true });
                  } else {
                    this.setState({ successStatus: "approved" });

                    this.setState({
                      isPasscodeModalOpen: true,
                      passCodeValidationFrom: "9000",
                    });
                  }
                }
          }
        >
          {this._checkCurrentUserIsAATRAssignee() &&
          this._checkCurrentUserIsApproverType()
            ? "Noted"
            : "Approve"}
        </PrimaryButton>

        <PrimaryButton
          className={`${styles.responsiveButton}`}
          iconProps={{ iconName: "PageRemove" }}
          styles={{
            root: {
              border: "none",
            },
            rootHovered: {
              border: "none",
            },
            rootPressed: {
              border: "none",
            },
          }}
          onClick={(e) => {
            if (this.state.errorForCummulative) {
              this.setState({ dialogboxForCummulativeError: true });
              return;
            }

            if (this.state.errorOfDocuments) {
              this.setState({ isAutoSaveFailedDialog: true });
            } else if (this.state.generalComments.length === 0) {
              this.setState({ isRejectCommentsCheckAlterDialog: true });
            } else {
              this.setState({ successStatus: "rejected" });

              this.setState({
                isPasscodeModalOpen: true,
                passCodeValidationFrom: "8000",
              });
            }
          }}
        >
          Reject
        </PrimaryButton>

        <PrimaryButton
          className={`${styles.responsiveButton}`}
          iconProps={{ iconName: "Share" }}
          onClick={(e) => {
            if (this.state.errorForCummulative) {
              this.setState({ dialogboxForCummulativeError: true });
              return;
            }

            this.setState({ successStatus: "referred" });
            if (this.state.errorOfDocuments) {
              this.setState({ isAutoSaveFailedDialog: true });
              return;
            }

            this._hanldeFluentDialog(
              "Refer",
              "Refered",
              "4000",
              ["Add Referee", "Comments"],
              "",
              this._closeDialog,
              this.referPassCodeTrigger
            );
          }}
        >
          Refer
        </PrimaryButton>

        <PrimaryButton
          className={`${styles.responsiveButton}`}
          iconProps={{ iconName: "Undo" }}
          onClick={(e) => {
            if (this.state.errorForCummulative) {
              this.setState({ dialogboxForCummulativeError: true });
              return;
            }

            if (this.state.errorOfDocuments) {
              this.setState({ isAutoSaveFailedDialog: true });
              return;
            }
            if (this.state.generalComments.length === 0) {
              this.setState({ isReturnCommentsCheckAlterDialog: true });
            } else {
              this.setState({ successStatus: "returned" });

              this.setState({
                isPasscodeModalOpen: true,
                passCodeValidationFrom: "5000",
              });
            }
          }}
        >
          Return
        </PrimaryButton>
      </div>
    );
  };

  private _closeDialog = () => {
    this.setState({ dialogFluent: true });
  };

  private _hanldeFluentDialog = (
    btnType: string,
    currentStatus: string,
    currentStatusNumber: string,
    message: any,
    functionType: any,
    closeFunction: any,
    referPassFuntion: any
  ) => {
    this.setState({
      dialogFluent: false,
      dialogDetails: {
        type: btnType,
        status: currentStatus,
        statusNumber: currentStatusNumber,
        subText: `Are you sure you want to ${btnType} this request?`,
        message: message,
        functionType: functionType,
        closeFunction: closeFunction,
        referPassFuntion: referPassFuntion,
      },
    });
  };

  public _getCommentData = (
    commentsData: any,
    type: string = "",
    id: string = ""
  ) => {
    console.log(commentsData);
    if (this.state.statusNumber === "4000") {
      this.setState((prevState) => ({
        noteReferrerCommentsDTO: [
          ...prevState.noteReferrerCommentsDTO,
          {
            ...commentsData,
            approverEmailName: prevState.currentApprover[0].Title,
          },
        ],
      }));
    }

    if (type === "add") {
      this.setState((prev) => {
        return {
          commentsLog: [...prev.commentsLog, commentsData],
          commentsData: [...prev.commentsData, commentsData],
          generalComments: [...prev.generalComments, commentsData],
        };
      });
    } else if (type === "delete") {
      const filteredComments = this.state.generalComments.filter(
        (comment: any) => comment !== null
      );

      const updatingCommentData = filteredComments.filter((each: any) => {
        return each.id !== id;
      });

      const filterCommentLogOFNotCurrentUser = this.state.commentsLog.filter(
        (each: any) => each.commentedByEmail !== this._currentUserEmail
      );
      this.setState({
        commentsData: updatingCommentData,
        generalComments: updatingCommentData,
        commentsLog: [
          ...filterCommentLogOFNotCurrentUser,
          ...updatingCommentData,
        ],
      });
    } else {
      const filterNullData = this.state.commentsLog.filter(
        (each: any) => each !== null
      );

      const filterIdforUpdateState = filterNullData.filter(
        (each: any) => each?.id === id
      )[0];

      const returnValue = (rowData: any): any => {
        const result = rowData
          .filter((each: any) => each !== null)
          .map((item: any) => {
            if (item.id === filterIdforUpdateState.id) {
              return commentsData;
            }
            return item;
          });

        return result;
      };

      const filterNullGeneral = this.state.generalComments.filter(
        (each: any) => each !== null
      );

      const filterIdforUpdateStateGen = filterNullGeneral.filter(
        (each: any) => each.id === id
      )[0];

      const returnValueGen = (rowData: any): any => {
        const result = rowData
          .filter((each: any) => each !== null)
          .map((item: any) => {
            if (item.id === filterIdforUpdateStateGen.id) {
              return commentsData;
            }
            return item;
          });

        return result;
      };

      this.setState((prevState) => ({
        commentsData: returnValue(prevState.commentsData),
        commentsLog: returnValue(prevState.commentsLog),
        generalComments: returnValueGen(prevState.generalComments),
      }));
      
    }
  };

  public _atrJoinedCommentsToDTO = (): void => {
    const joinedCommentsData = this.state.generalComments
      .filter((each: any) => !!each)
      .map((each: any) => `${each?.pageNum} ${each?.page} ${each?.comment}`)
      .join(", ");

    return joinedCommentsData;
  };

  private handleSupportingFileChangeInViewForm = (
    files: File[],
    typeOfDoc: string
  ) => {
    if (files) {
      const filesArray = Array.from(files);

      if (files.length > 0) {
        this.setState({
          supportingFilesInViewForm: [...filesArray],
        });
      } else {
        this.setState({
          supportingFilesInViewForm: filesArray,
        });
      }
    }
  };

  private handleGistDocuments = (files: File[], typeOfDoc: string) => {
    if (files) {
      const filesArray = Array.from(files);

      this.setState({
        secretaryGistDocs: filesArray,
        secretaryGistDocsList: filesArray,
      });
    }
  };

  public _checkCurrentRequestIsReturnedOrRejected = (): boolean => {
    switch (this.state.statusNumber) {
      case "8000":
      case "5000":
      case "200":
      case "9000":
      case "300":
        return false;
      default:
        return true;
    }
  };

  private _checkCurrentUserIsAATRAssignee = (): any => {
    const checkingATRAvailable = this.state.atrCreatorsList.some(
      (each: any) => {
        if (each.atrCreatorEmail === this._currentUserEmail) {
          return true;
        }
      }
    );

    return checkingATRAvailable;
  };

  private _checkCurrentUserIsApproverType = (): any => {
    const checkingATRAvailable = this.state.ApproverDetails.some(
      (each: any) => {
        if (
          each.approverEmail === this._currentUserEmail &&
          each.approverType === "Approver"
        ) {
          return true;
        }
      }
    );

    return checkingATRAvailable;
  };

  private _checkingCurrentATRCreatorisCurrentApproverOrNot = (): any => {
    const checkingCurrentATRCreatorisCurrentApproverOrNot =
      this.state.currentApprover?.length > 0 &&
      this.state.currentApprover[0]?.EMail === this._currentUserEmail;

    return (
      checkingCurrentATRCreatorisCurrentApproverOrNot &&
      this.state.statusNumber !== "4000" &&
      this.state.statusNumber !== "5000" &&
      this.state.statusNumber !== "8000"
    );
  };

  public _closeDialogAlter = (type: string) => {
    if (type === "success") {
      const pageURL: string = this.props.homePageUrl;

      window.location.href = `${pageURL}`;
    } else if (type === "commentsNeeded") {
      this.setState({
        expandSections: { generalComments: true, generalSection: false },
      });
    }

    this.setState({
      isVisibleAlter: false,
      isGistSuccessVisibleAlter: false,
      isReferBackAlterDialog: false,
      isRejectCommentsCheckAlterDialog: false,
      isReturnCommentsCheckAlterDialog: false,
      isNotedCommentsManidatoryAlterDialog: false,
    });
  };

  public handlePasscodeSuccess = () => {
    this.setState(
      { isPasscodeValidated: true, isPasscodeModalOpen: false },
      () => {
        switch (this.state.passCodeValidationFrom) {
          case "9000":
            this._hanldeFluentDialog(
              this.state.successStatus === "approved" ? "approve" : "note",
              this.state.successStatus === "approved" ? "Noted" : "Approved",
              "9000",
              this.state.successStatus === "approved"
                ? "Please check the details filled along with attachment and click on Confirm button to approve the request."
                : "Please check the details filled along with attachment and click on Confirm button to note the request.",
              this._handleApproverButton,
              this._closeDialog,
              ""
            );
            break;
          case "1000":
          case "2000":
          case "3000":
          case "6000":
          case "4900":
            this._hanldeFluentDialog(
              "refer back",
              "Refered Back",
              "4900",
              "Please check the details filled along with attachment and click on Confirm button to refer back the request.",
              this.handleReferBack,
              this._closeDialog,
              ""
            );
            break;
          case "4000":
            this._hanldeFluentDialog(
              "refer",
              "Refered",
              "4000",
              "Please check the details filled along with attachment and click on Confirm button to refer the request.",
              this.handleRefer,
              this._closeDialog,
              ""
            );
            break;
          case "5000":
            this._hanldeFluentDialog(
              "return",
              "Returned",
              "5000",
              "Please check the details filled along with attachment and click on Confirm button to return the request.",
              this.handleReturn,
              this._closeDialog,
              ""
            );
            break;
          case "8000":
            this._hanldeFluentDialog(
              "reject",
              "Rejected",
              "8000",
              "Please check the details filled along with attachment and click on Confirm button to reject the request.",
              this.handleReject,
              this._closeDialog,
              ""
            );

            break;
          case "200":
            this.handleCallBack("Call Back", "200");
            break;
          case "7500":
            this.setState({ dialogFluent: true });
            this._hanldeFluentDialog(
              "change approver",
              "Approver Changed",
              "7500",
              "Please click on Confirm button to change approver.",
              this.handleChangeApprover,
              this._closeDialog,
              ""
            );
            break;

          default:
            break;
        }
      }
    );
  };

  private _randomFileIcon = (docType: string): any => {
    const docExtension = docType.split(".");
    const fileExtession = docExtension[docExtension.length - 1];

    let doctype;

    switch (fileExtession.toLocaleLowerCase()) {
      case "docx":
      case "doc":
        doctype = "docx";
        break;

      case "pdf":
        doctype = "pdf";
        break;

      case "xlsx":
        doctype = "xlsx";
        break;

      default:
        doctype = "txt";
    }

    const url = `https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/assets/item-types/16/${doctype}.svg`;
    return url;
  };

  private _getCummulativeError = (data: any): any => {
    data !== null
      ? this.setState({
          errorForCummulative: true,
        })
      : this.setState({
          errorForCummulative: false,
        });
  };

  private _getFileWithError = (data: any): any => {
    const newObj = this.state.errorFilesList;
    newObj[data[1]] = data[0];

    this.setState((prevState) => {
      const newObj = { ...prevState.errorFilesList }; 
      newObj[data[1]] = data[0]; 
    
      return { errorFilesList: newObj }; 
    });

    if (
      newObj.wordDocument.length > 0 ||
      newObj.notePdF.length > 0 ||
      newObj.supportingDocument.length > 0 ||
      newObj.gistDocument.length > 0
    ) {
      this.setState({
        errorOfDocuments: true,
      });
    } else {
      this.setState({
        errorOfDocuments: false,
      });
    }
  };

  private _getAtrCommentsGrid = (data: any): any => {
    console.log(data)
    if (
      this.state.currentApprover !== null &&
      this.state.currentApprover[0]?.approverEmail ||this.state.currentApprover[0]?.EMail === this._currentUserEmail
    ) {
      const joinedCommentsData = this.state.generalComments .filter((each: any) => !!each)
        .map((each: any) => `${each?.pageNumber} ${each?.docReference} ${each?.comments}`);

      console.log(joinedCommentsData)

      return data.map((each: any) => {
        return { ...each, comments: joinedCommentsData.join(", ") };
      });
    } else {
      return this.state.atrGridData;
    }
  };

  private closeUserExistsModal = () => {
    this.setState({ isUserExistsModalVisible: false });
  };

  private getUserExistsModalJSX = (): any => {
    return (
      <Modal
        isOpen={this.state.isUserExistsModalVisible}
        onDismiss={this.closeUserExistsModal}
        isBlocking={true}
        styles={{
          main: {
            width: "100%",
            maxWidth: "290px",
            "@media (min-width: 768px)": {
              maxWidth: "580px",
            },
          },
        }}
      >
        <div
          style={{
            display: "flex",
            flexDirection: "row",
            justifyContent: "space-between",
            alignItems: "center",
            padding: "8px 12px",
            borderBottom: "1px solid #ddd",
          }}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: "8px",
            }}
          >
            <IconButton iconProps={{ iconName: "Info" }} />

            <h4 className={styles.headerTitle}>Alert</h4>
          </div>

          <IconButton
            iconProps={{ iconName: "Cancel" }}
            ariaLabel="Close modal"
            onClick={this.closeUserExistsModal}
          />
        </div>

        <Stack
          tokens={{ padding: "16px" }}
          horizontalAlign="center"
          verticalAlign="center"
        >
          <Text
            style={{ margin: "16px 0", fontSize: "14px", textAlign: "center" }}
          >
            The selected approver cannont be same as existing
            Reviewers/Requester/referee/CurrentActioner
          </Text>
        </Stack>

        <div
          style={{
            display: "flex",
            justifyContent: "flex-end",
            padding: "12px 16px",
            borderTop: "1px solid #ddd",
          }}
        >
          <PrimaryButton
            iconProps={{ iconName: "ReplyMirrored" }}
            text="ok"
            onClick={this.closeUserExistsModal}
            ariaLabel="Close modal"
          />
        </div>
      </Modal>
    );
  };

  private _getCallBackAndChangeApproverBtn = () => {
    return this._checkApproveredStatusIsFound() ? (
      <PrimaryButton
        className={`${styles.responsiveButton}`}
        iconProps={{ iconName: "Contact" }}
        onClick={(e) => {
          this.setState({ successStatus: "approver changed" });
          this._hanldeFluentDialog(
            "Change Approver",
            "changeApprover",
            "7500",
            "Change Approver",
            "",
            this._closeDialog,
            this.changeApproverPassCodeTrigger
          );
        }}
      >
        Change Approver
      </PrimaryButton>
    ) : (
      this.state.statusNumber !== "100" && (
        <PrimaryButton
          className={`${styles.responsiveButton}`}
          iconProps={{ iconName: "Previous" }}
          onClick={(e) => {
            this.setState({ successStatus: "call back" });

            if (!this.state.isPasscodeValidated) {
              this.setState({
                isPasscodeModalOpen: true,
                passCodeValidationFrom: "200",
              });
              return null;
            }
          }}
        >
          Call Back
        </PrimaryButton>
      )
    );
  };

  private _getReferBackAndApproverStageButtons = () => {
   
    return this.state.noteReferrerDTO.length > 0 &&
      this.state.noteReferrerDTO[this.state.noteReferrerDTO.length - 1]
        ?.referrerEmail === this._currentUserEmail &&
      this.state.statusNumber === "4000" ? (
      <PrimaryButton
        className={`${styles.responsiveButton}`}
        iconProps={{ iconName: "Reply" }}
        styles={{
          root: {
            border: "none",
          },
          rootHovered: {
            border: "none",
          },
          rootPressed: {
            border: "none",
          },
        }}
        onClick={(e) => {
          this.setState({ successStatus: "refered back" });

          if (this.state.errorForCummulative) {
            this.setState({ dialogboxForCummulativeError: true });
            return;
          }

          if (this.state.errorOfDocuments) {
            this.setState({ isAutoSaveFailedDialog: true });
            return;
          }
          if (this.state.generalComments.length === 0) {
            this.setState({ isReferBackAlterDialog: true });
          } else if (!this.state.isPasscodeValidated) {
            this.setState({
              isPasscodeModalOpen: true,
              passCodeValidationFrom: "4900",
            });
            return;
          }
        }}
      >
        Refer Back
      </PrimaryButton>
    ) : (
      this._checkCurrentUserIs_Approved_Refered_Reject_TheCurrentRequest() &&
        this._getApproverAndReviewerStageButton()
    );
  };

  private getMainStatus = (): any => {
    const approver = this.state.ApproverDetails.find(
      (detail: any) =>
        (detail.approverEmail || detail.email || detail.secondaryText) ===
        (this.state.currentApprover[0].approverEmail ||
          this.state.currentApprover[0].EMail ||
          this.state.currentApprover[0].secondaryText)
    );

    return approver ? approver.mainStatus : undefined;
  };

  private _DialogBlockingExample = (): any => {
    return (
      <DialogBlockingExample
        changeApproverDataMandatory={this._changeApproverDataMandatory}
        referCommentsAndDataMandatory={this._referCommentsAndDataMandatory}
        statusNumberForChangeApprover={this.state.statusNumber}
        referDto={
          this.state.noteReferrerDTO[this.state.noteReferrerDTO.length - 1]
        }
        requesterEmail={this.state.createdByEmail}
        isUserExistingDialog={() =>
          this.setState({ isUserExistsModalVisible: true })
        }
        dialogUserCheck={{
          peoplePickerApproverData: this.state.peoplePickerApproverData,
          peoplePickerData: this.state.peoplePickerData,
        }}
        hiddenProp={this.state.dialogFluent}
        dialogDetails={this.state.dialogDetails}
        sp={this.props.sp}
        context={this.props.context}
        fetchReferData={(data: any) => {
         
          this.setState((prevState) => ({
            commentsData: [...prevState.commentsData, data],
            generalComments: [...prevState.commentsData, data],
            commentsLog: [...prevState.commentsLog, data],
            referComment: data,
          }));
        }}
        fetchAnydata={(
          data: any,
          typeOfBtnTriggered: any,
          status: any,
          commentData: any
        ) => {
          this.setState({
            peoplePickerSelectedDataWhileReferOrChangeApprover: data,
          });
          if (typeOfBtnTriggered === "Refer") {
            this.setState((prevState) => ({
              refferredToDetails: [{ ...data[0], status: status }],
              referredFromDetails: [...prevState.currentApprover],
            }));
          }
        }}
      />
    );
  };

  private _RenderMainViewForm = (): any => {
    const { expandSections } = this.state;

    const formTitle =
      this.props.formType === "BoardNoteView"
        ? `Board Note - ${this.state.title}`
        : `eCommittee Note - ${this.state.title}`;
    return (
      <div className={styles.viewFormMainContainer}>
        <form>
          <PasscodeModal
            createPasscodeUrl={this.props.passCodeUrl}
            isOpen={this.state.isPasscodeModalOpen}
            onClose={() =>
              this.setState({
                isPasscodeModalOpen: false,
                isPasscodeValidated: false,
              })
            }
            onSuccess={this.handlePasscodeSuccess}
            sp={this.props.sp}
            user={this.props.context.pageContext.user}
          />
        </form>

        {this.getUserExistsModalJSX()}

        <SuccessDialog
          existUrl={this.props.existPageUrl}
          statusOfReq={this.state.successStatus}
          isVisibleAlter={this.state.isVisibleAlter}
          onCloseAlter={() => {
            this._closeDialogAlter("success");
          }}
          typeOfNote={this._committeeType}
        />
        <Modal
          isOpen={this.state.hideParellelActionAlertDialog}
          onDismiss={() => {
          
            this.setState((prevState) => ({
              hideParellelActionAlertDialog: !prevState.hideParellelActionAlertDialog,
            }));
            
          }}
          isBlocking={true}
          containerClassName={Cutsomstyles.modal}
        >
          <div className={Cutsomstyles.header}>
            <div style={{ display: "flex", alignItems: "center" }}>
              <IconButton iconProps={{ iconName: "Info" }} />
              <h4 className={Cutsomstyles.headerTitle}>Alert</h4>
            </div>
            <IconButton
              iconProps={{ iconName: "Cancel" }}
              onClick={() => {
               
                window.location.reload();
                this.setState({ hideParellelActionAlertDialog: false });
              }}
            />
          </div>
          <div className={Cutsomstyles.body}>
            <p>{this.state.parellelActionAlertMsg}</p>
          </div>
          <div className={Cutsomstyles.footer}>
            <PrimaryButton
              className={Cutsomstyles.button}
              iconProps={{ iconName: "ReplyMirrored" }}
              onClick={() => {
                this.setState({ hideParellelActionAlertDialog: false });
                window.location.reload();
              }}
              text="OK"
            />
          </div>
        </Modal>

        <ChangeApproverMandatoryDialog
          isVisibleAlter={this.state.isChangeApproverNeeded}
          onCloseAlter={() => {
            this.setState({ isChangeApproverNeeded: false });
          }}
        />

        {this.state.isLoading && (
          <div>
            <Modal
              isOpen={this.state.isLoading}
              containerClassName={styles.spinnerModalTranparency}
              styles={{
                main: {
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  background: "transparent",
                  boxShadow: "none",
                },
              }}
            >
              <div className="spinner">
                <Spinner
                  label="still loading..."
                  ariaLive="assertive"
                  size={SpinnerSize.large}
                />
              </div>
            </Modal>
          </div>
        )}

        <ReferCommentsMandatoryDialog
          isVisibleAlter={this.state.isReferDataAndCommentsNeeded}
          onCloseAlter={() => {
            this.setState({ isReferDataAndCommentsNeeded: false });
          }}
          statusOfReq={
            this.state.peoplePickerSelectedDataWhileReferOrChangeApprover
          }
        />

        <ReferBackCommentDialog
          statusOfReq={this.state.status}
          isVisibleAlter={this.state.isReferBackAlterDialog}
          onCloseAlter={() => {
            this._closeDialogAlter("commentsNeeded");
          }}
        />

        <NotedCommentDialog
          statusOfReq={this.state.status}
          isVisibleAlter={this.state.isNotedCommentsManidatoryAlterDialog}
          onCloseAlter={() => {
            this._closeDialogAlter("commentsNeeded");
          }}
        />

        <GistDocEmptyModal
          isVisibleAlter={this.state.isGistDocEmpty}
          onCloseAlter={() => {
            this.setState({ isGistDocEmpty: false });
          }}
          statusOfReq={undefined}
        />

        <GistDocsConfirmation
          isVisibleAlter={this.state.isGistDocCnrf}
          onCloseAlter={() => {
            this.setState({ isGistDocCnrf: false });
          }}
          handleConfirmatBtn={async () => {
            this.setState({ isGistDocCnrf: false, isLoading: true });

            try {
              await this.updateGistDocumentFolderItems(
                this.state.secretaryGistDocs,
                `${this._folderName}/GistDocuments`,
                "gistDocument"
              ).then(async () => {
                const updateAuditTrial = await this._getAuditTrail(
                  "gistDocuments"
                );
                await this.props.sp.web.lists
                  .getByTitle(this._listname)
                  .items.getById(this._itemId)
                  .update({
                    AuditTrail: updateAuditTrial,
                  });
              });

              this.setState({
                isLoading: false,
                isGistSuccessVisibleAlter: true,
              });
            } catch (e) {
              return e;
            }
          }}
          statusOfReq={undefined}
        />

        <GistDocSubmitted
          existUrl={this.props.existPageUrl}
          isVisibleAlter={this.state.isGistSuccessVisibleAlter}
          onCloseAlter={() => {
            this._closeDialogAlter("success");
          }}
          statusOfReq={undefined}
        />

        <RejectBtnCommentCheckDialog
          statusOfReq={this.state.status}
          isVisibleAlter={this.state.isRejectCommentsCheckAlterDialog}
          onCloseAlter={() => {
            this._closeDialogAlter("commentsNeeded");
          }}
        />

        <ReturnBtnCommentCheckDialog
          statusOfReq={this.state.status}
          isVisibleAlter={this.state.isReturnCommentsCheckAlterDialog}
          onCloseAlter={() => {
            this._closeDialogAlter("commentsNeeded");
          }}
        />

        <Dialog
          hidden={!this.state.isDialogVisible}
          onDismiss={() => this.setState({ isDialogVisible: false })}
          dialogContentProps={{
            title: this.state.dialogContent.title,
          }}
        >
          <div>{this.state.dialogContent.message}</div>{" "}
          <DialogFooter>
            <PrimaryButton
              onClick={() => this.setState({ isDialogVisible: false })}
              text={this.state.dialogContent.buttonText}
            />
          </DialogFooter>
        </Dialog>

        <CummulativeErrorDialog
          isVisibleAlter={this.state.dialogboxForCummulativeError}
          onCloseAlter={() => {
            this.setState({ dialogboxForCummulativeError: false });
          }}
          statusOfReq={undefined}
        />

        {this.state.isAutoSaveFailedDialog && (
          <AutoSaveFailedDialog
            statusOfReq={this.state.successStatus}
            isVisibleAlter={this.state.isAutoSaveFailedDialog}
            onCloseAlter={() => {
              this.setState({ isAutoSaveFailedDialog: false });
            }}
          />
        )}

        <div
          className={`${styles.generalSectionMainContainer} ${styles.viewFormHeaderSection}`}
          style={{ paddingLeft: "10px", paddingRight: "10px" }}
        >
          <h1
            className={`${styles.generalHeader} ${styles.viewFormHeaderSectionContainer}`}
          >
            pending with:{" "}
           
            {this.state.currentApprover[0]?.Title}
          </h1>

          <h1
            className={`${styles.generalHeader} ${styles.viewFormHeaderSectionContainer} `}
          >
            {formTitle}
          </h1>

          <h1
            className={`${styles.generalHeader} ${styles.viewFormHeaderSectionContainer}`}
          >
            Status:{" "}
            {this.state.statusNumber === "4900"
              ? this.getMainStatus()
              : this.state.status}
          </h1>
        </div>

        <div className={`${styles.viewFormContentContainer}`}>
          <div className={styles.expansionAndPdfContainer}>
            <div className={styles.expandingContainer}>
              <GeneralSectionInViewForm
                _onToggleSection={this._onToggleSection}
                expandSections={expandSections}
                state={this.state}
                _renderTable={this._renderTable}
              />

              <DraftResolutionInViewForm
                _onToggleSection={this._onToggleSection}
                expandSections={expandSections}
                state={this.state}
                formType={this.props.formType}
              />

              <ReviewerOrApproverSectionInViewForm
                sectionName="Reviewers Section"
                _onToggleSection={this._onToggleSection}
                toggleParameter="reviewersSection"
                expandSections={expandSections}
                state={this.state}
                reviewerORApproverData={this.state.peoplePickerData}
                reOrderData={this.reOrderData}
              
                type="Reviewer"
              />
              <ReviewerOrApproverSectionInViewForm
                sectionName="Approvers Section"
                _onToggleSection={this._onToggleSection}
                toggleParameter="approversSection"
                expandSections={expandSections}
                state={this.state}
                reviewerORApproverData={this.state.peoplePickerApproverData}
                reOrderData={this.reOrderData}
                
                type="Approver"
              />

              {(this._checkCurrentUserIs_Approved_Refered_Reject_TheCurrentRequest() &&
                this._currentUserEmail !== this.state.createdByEmail) ||
              this._checkRefereeAvailable() ? (
                <div className={styles.sectionContainer}>
                  <button
                    className={styles.header}
                    onClick={() => this._onToggleSection(`generalComments`)}
                  >
                    <Text className={styles.sectionText}>General Comments</Text>
                    <IconButton
                      iconProps={{
                        iconName: expandSections.generalComments
                          ? "ChevronUp"
                          : "ChevronDown",
                      }}
                      title="Expand/Collapse"
                      ariaLabel="Expand/Collapse"
                      className={styles.chevronIcon}
                    />
                  </button>

                  {expandSections.generalComments && (
                    <div className={`${styles.expansionPanelInside}`}>
                      <div style={{ padding: "15px", paddingTop: "4px" }}>
                        <GeneralCommentsFluentUIGrid
                          handleCommentDataFuntion={this._getCommentData}
                          data={this.state.generalComments}
                          currentUserDetails={
                            this.props.context.pageContext.user
                          }
                        />
                      </div>
                    </div>
                  )}
                </div>
              ) : (
                ""
              )}

              {this._checkCurrentUserIsAATRAssignee() &&
                this._checkCurrentUserIsApproverType() && (
                  <div className={styles.sectionContainer}>
                    <button
                      className={styles.header}
                      onClick={() => this._onToggleSection(`atrAssignees`)}
                    >
                      <Text className={styles.sectionText}>ATR Assignees</Text>
                      <IconButton
                        iconProps={{
                          iconName: expandSections.atrAssignees
                            ? "ChevronUp"
                            : "ChevronDown",
                        }}
                        title="Expand/Collapse"
                        ariaLabel="Expand/Collapse"
                        className={styles.chevronIcon}
                      />
                    </button>
                    {expandSections.atrAssignees && (
                      <div
                        className={`${styles.expansionPanelInside}`}
                        style={{ overflowX: "scroll" }}
                      >
                        <div style={{ padding: "15px" }}>
                          <ATRAssignee
                            atrType={this.state.atrType}
                            getATRTypeOnChange={(type: any) => {
                              this.setState({ atrType: type });
                            }}
                            clearAtrGridDataOnSelectionOFATRType={() => {
                              this.setState({
                                atrGridData: [],
                                noteATRAssigneeDetails: [],
                              });
                            }}
                            checkingCurrentATRCreatorisCurrentApproverOrNot={this._checkingCurrentATRCreatorisCurrentApproverOrNot()}
                            getATRJoinedComments={(data: any) => {
                              this.setState({ atrJoinedComments: data });
                            }}
                            approverDetails={this.state.ApproverDetails}
                            currentATRCreatorDetails={this._currentUserEmail}
                            sp={this.props.sp}
                            context={this.props.context}
                            commentsData={this.state.generalComments}
                            artCommnetsGridData={this._getAtrCommentsGrid(
                              this.state.atrGridData
                            )}
                            deletedGridData={(data: any) => {
                              this.setState({ atrGridData: data });
                            }}
                            updategirdData={(data: any): void => {
                              this.setState({ atrType: data.atrType });

                              const currentAtrCreator =
                                this.state.atrCreatorsList.filter(
                                  (each: any) =>
                                    each.atrCreatorEmail ===
                                    this.props.context.pageContext.user.email
                                );

                              const { assigneeDetails } = data;

                              const currentApproverDetailsFromApproverDTO =
                                this._getCurrentApproverDetailsFromApproverDTO();

                            

                              this.setState((prevState) => ({
                                atrGridData: data.comments,

                                noteATRAssigneeDetails: [
                                  ...prevState.noteATRAssigneeDetails,
                                  {
                                    atrType: data.atrType,
                                    atrAssigneeId: assigneeDetails.id,
                                    atrCreatorId:
                                      currentAtrCreator[0].atrCreatorId,
                                    atrCreatorEmail:
                                      currentAtrCreator[0].atrCreatorEmail,

                                    atrAssigneeEmailName: assigneeDetails.text,
                                    atrAssigneeEmail: assigneeDetails.email,
                                    approverEmailName:
                                      currentApproverDetailsFromApproverDTO[0]
                                        .approverEmailName,

                                    atrCreatorEmailName:
                                      currentAtrCreator[0].atrCreatorEmailName,

                                    createdDate: this._formatDateTime(
                                      new Date()
                                    ),
                                    createdBy:
                                      this.props.context.pageContext.user.email,
                                    modifiedDate: this._formatDateTime(
                                      new Date()
                                    ),
                                    modifiedBy:
                                      this.props.context.pageContext.user.email,
                                    statusMessage: null,
                                    atrId: "",
                                    noteApproverId:
                                      currentApproverDetailsFromApproverDTO[0]
                                        .userId,
                                    approverType:
                                      currentApproverDetailsFromApproverDTO[0]
                                        .approverType,
                                    approverOrder:
                                      currentApproverDetailsFromApproverDTO[0]
                                        .approverOrder,
                                    approverStatus: 1,
                                    approverEmail:
                                      currentApproverDetailsFromApproverDTO[0]
                                        .approverEmail,
                                    noteApproverComments: "",
                                    strATRStatus: "Pending",
                                    atrStatus: 1,
                                    noteId: this._itemId,
                                  },
                                ],
                              }));
                            }}
                          />
                        </div>
                      </div>
                    )}
                  </div>
                )}

              <div className={styles.sectionContainer}>
                <button
                  className={styles.header}
                  onClick={() => this._onToggleSection(`commentsLog`)}
                >
                  <Text className={styles.sectionText}>Comments Log</Text>
                  <IconButton
                    iconProps={{
                      iconName: expandSections.commentsLog
                        ? "ChevronUp"
                        : "ChevronDown",
                    }}
                    title="Expand/Collapse"
                    ariaLabel="Expand/Collapse"
                    className={styles.chevronIcon}
                  />
                </button>
                {expandSections.commentsLog && (
                  <div className={`${styles.expansionPanelInside}`}>
                    <div style={{ padding: "15px", paddingTop: "4px" }}>
                      <CommentsLogTable
                        data={this.state.commentsLog}
                        type="commentsLog"
                        formType="view"
                      />
                    </div>
                  </div>
                )}
              </div>

              {(this.state.currentApprover?.[0]?.approverEmail ||
                this.state.currentApprover?.[0]?.EMail) ===
              this._currentUserEmail ? (
                <div className={styles.sectionContainer}>
                  <button
                    className={styles.header}
                    onClick={() =>
                      this._onToggleSection(`attachSupportingDocuments`)
                    }
                  >
                    <Text className={styles.sectionText}>
                      Attach Supporting Documents
                    </Text>
                    <IconButton
                      iconProps={{
                        iconName: expandSections.attachSupportingDocuments
                          ? "ChevronUp"
                          : "ChevronDown",
                      }}
                      title="Expand/Collapse"
                      ariaLabel="Expand/Collapse"
                      className={styles.chevronIcon}
                    />
                  </button>
                  {expandSections.attachSupportingDocuments && (
                    <div
                      className={`${styles.expansionPanelInside}`}
                      style={{ width: "100%", margin: "0px" }}
                    >
                      <div style={{ padding: "15px", paddingTop: "4px" }}>
                        <SupportingDocumentsUploadFileComponent
                          errorData={this._getFileWithError}
                          typeOfDoc="supportingDocument"
                          onChange={this.handleSupportingFileChangeInViewForm}
                          accept=".xlsx,.pdf,.doc,.docx"
                          multiple={true}
                          maxFileSizeMB={25}
                          data={this.state.supportingFilesInViewForm}
                          addtionalData={this.state.supportingDocumentfiles}
                          cummulativeError={this._getCummulativeError}
                        />
                        <p
                          className={styles.message}
                          style={{ margin: "0px", textAlign: "right" }}
                        >
                          Allowed Formats (pdf,doc,docx,xlsx only) Upto 25MB
                          max.
                        </p>
                      </div>
                    </div>
                  )}
                </div>
              ) : (
                ""
              )}

              {this._checkingCurrentUserInSecretaryDTO() &&
              this.state.statusNumber !== "5000" &&
              this.state.statusNumber !== "8000" &&
              this.state.statusNumber !== "4000" ? (
                <div className={styles.sectionContainer}>
                  <button
                    className={styles.header}
                    onClick={() => this._onToggleSection(`gistDocuments`)}
                  >
                    <Text className={styles.sectionText}>Gist Document</Text>
                    <IconButton
                      iconProps={{
                        iconName: expandSections.gistDocuments
                          ? "ChevronUp"
                          : "ChevronDown",
                      }}
                      title="Expand/Collapse"
                      ariaLabel="Expand/Collapse"
                      className={styles.chevronIcon}
                    />
                  </button>
                  {expandSections.gistDocuments && (
                    <div
                      className={`${styles.expansionPanelInside}`}
                      style={{ width: "100%", margin: "0px" }}
                    >
                      <div style={{ padding: "6px", paddingTop: "4px" }}>
                        <div
                          style={{
                            display: "flex",
                            flexDirection: "column",
                            alignItems: "flex-start",
                            padding: "15px",
                            paddingTop: "4px",
                          }}
                        >
                          {this._checkingCurrentUserIsSecretaryDTO() ? (
                            <UploadFileComponent
                              errorData={this._getFileWithError}
                              typeOfDoc="gistDocument"
                              onChange={this.handleGistDocuments}
                              accept=".pdf,.doc,.docx "
                              multiple={false}
                              maxFileSizeMB={5}
                              data={this.state.secretaryGistDocs}
                              addtionalData={this.state.secretaryGistDocsList}
                            />
                          ) : (
                            this._checkingCurrentUserInSecretaryDTOAfterApproved() && (
                              <div
                                style={{
                                  padding: "6px",
                                  border: "1px solid rgb(211, 211, 211)",
                                  width: "100%",
                                }}
                              >
                                <p>Gist Document</p>
                                {this._checkingCurrentUserInSecretaryDTO() &&
                                this.state.secretaryGistDocsList.length > 0 ? (
                                  this.state.secretaryGistDocsList.map(
                                    (file, index) => {
                                      if (!file || !file.name) {
                                        return null;
                                      }

                                      return (
                                        <li
                                          key={v4()}
                                          style={{
                                            width: "100%",
                                            marginTop: "5px",
                                          }}
                                          className={`${styles.basicLi} ${styles.attachementli}`}
                                        >
                                          <div
                                            className={`${styles.fileIconAndNameWithErrorContainer}`}
                                          >
                                            <img
                                              alt="typeOfIconInGist1"
                                              src={this._randomFileIcon(
                                                file.name
                                              )}
                                              width={32}
                                              height={32}
                                            />

                                            <a
                                              data-interception="off"
                                              className={styles.notePdfCustom}
                                              href={
                                                file.name
                                                  .toLowerCase()
                                                  .endsWith(".pdf")
                                                  ? file.fileUrl
                                                  : file.LinkingUri
                                              }
                                              target="_blank"
                                              rel="noopener noreferrer"
                                              style={{
                                                marginTop: "9px",
                                                paddingLeft: "4px",
                                                textDecoration: "none",
                                              }}
                                            >
                                              <span
                                                style={{
                                                  paddingBottom: "0px",
                                                  marginBottom: "0px",
                                                  paddingLeft: "4px",
                                                }}
                                              >
                                                {file.name.length > 30
                                                  ? `${file.name.slice(
                                                      0,
                                                      20
                                                    )}...`
                                                  : file.name}
                                              </span>
                                            </a>
                                          </div>
                                        </li>
                                      );
                                    }
                                  )
                                ) : (
                                  <h4>No File Found</h4>
                                )}
                              </div>
                            )
                          )}
                          {this._checkingCurrentUserIsSecretaryDTO() && (
                            <p
                              className={styles.message}
                              style={{ margin: "0px", textAlign: "right" }}
                            >
                              Allowed Formats (pdf,doc,docx,only) Upto 5MB max.
                            </p>
                          )}
                          {this._checkingCurrentUserAsApproverDTOInSecretaryDTO() && (
                            <div
                              style={{
                                padding: "6px",
                                border: "1px solid rgb(211, 211, 211)",
                                width: "100%",
                              }}
                            >
                              <p>Gist Document</p>
                              {this.state.secretaryGistDocsList.length > 0 ? (
                                this.state.secretaryGistDocsList.map(
                                  (file, index) => {
                                    if (!file || !file.name) {
                                      return null;
                                    }

                                    return (
                                      <li
                                        key={v4()}
                                        style={{
                                          width: "100%",
                                          marginTop: "5px",
                                        }}
                                        className={`${styles.basicLi} ${styles.attachementli}`}
                                      >
                                        <div
                                          className={`${styles.fileIconAndNameWithErrorContainer}`}
                                        >
                                          <img
                                            alt="typeOfIconInGist2"
                                            src={this._randomFileIcon(
                                              file.name
                                            )}
                                            width={32}
                                            height={32}
                                          />

                                          <a
                                            data-interception="off"
                                            className={styles.notePdfCustom}
                                            href={
                                              file.name
                                                .toLowerCase()
                                                .endsWith(".pdf")
                                                ? file.fileUrl
                                                : file.LinkingUri
                                            }
                                            target="_blank"
                                            rel="noopener noreferrer"
                                            style={{
                                              marginTop: "9px",
                                              paddingLeft: "4px",
                                              textDecoration: "none",
                                            }}
                                          >
                                            <span
                                              style={{
                                                paddingBottom: "0px",
                                                marginBottom: "0px",
                                                paddingLeft: "4px",
                                              }}
                                            >
                                              {file.name.length > 30
                                                ? `${file.name.slice(0, 20)}...`
                                                : file.name}
                                            </span>
                                          </a>
                                        </div>
                                      </li>
                                    );
                                  }
                                )
                              ) : (
                                <h4>No File Found</h4>
                              )}
                            </div>
                          )}
                        </div>
                      </div>
                      {""}
                      <div />
                    </div>
                  )}
                </div>
              ) : (
                ""
              )}

              <div className={styles.sectionContainer}>
                <button
                  className={styles.header}
                  onClick={() => this._onToggleSection(`workflowLog`)}
                >
                  <Text className={styles.sectionText}>Workflow Log</Text>
                  <IconButton
                    iconProps={{
                      iconName: expandSections.workflowLog
                        ? "ChevronUp"
                        : "ChevronDown",
                    }}
                    title="Expand/Collapse"
                    ariaLabel="Expand/Collapse"
                    className={styles.chevronIcon}
                  />
                </button>
                {expandSections.workflowLog && (
                  <div className={`${styles.expansionPanelInside}`}>
                    <div style={{ padding: "15px", paddingTop: "4px" }}>
                      <WorkFlowLogsTable
                        data={this.state.auditTrail}
                        type="Approver"
                      />
                    </div>
                  </div>
                )}
              </div>

              <div className={styles.sectionContainer}>
                <button
                  className={styles.header}
                  onClick={() => this._onToggleSection(`fileAttachments`)}
                >
                  <Text className={styles.sectionText}>File Attachments</Text>
                  <IconButton
                    iconProps={{
                      iconName: expandSections.fileAttachments
                        ? "ChevronUp"
                        : "ChevronDown",
                    }}
                    title="Expand/Collapse"
                    ariaLabel="Expand/Collapse"
                    className={styles.chevronIcon}
                  />
                </button>
                {expandSections.fileAttachments && (
                  <div
                    className={`${styles.expansionPanelInside} ${styles.responsiveContainerheaderForFileAttachment}`}
                  >
                    <div
                      style={{
                        padding: "15px",
                        paddingTop: "4px",
                        width: "100%",
                      }}
                    >
                      <p className={styles.responsiveHeading}>
                        Main Note Link :<a
                          href={this.state.noteTofiles[0]?.fileUrl}
                          target="_blank"
                          rel="noopener noreferrer"
                          data-interception="off"
                          className={styles.notePdfCustom}
                        >
                          {this.state.noteTofiles[0]?.name}
                        </a>
                      </p>
                      {this._checkingCurrentUserInSecretaryDTO() &&
                        this.state.wordDocumentfiles.length > 0 && (
                          <p
                            className={styles.responsiveHeading}
                            style={{ minWidth: "150px" }}
                          >
                            Word Documents :<a
                              href={this.state.wordDocumentfiles[0]?.LinkingUri}
                              target="_blank"
                              rel="noopener noreferrer"
                              data-interception="off"
                              className={styles.notePdfCustom}
                            >
                              {" "}
                              {this.state.wordDocumentfiles[0]?.name}
                            </a>
                          </p>
                        )}

                      <div style={{ width: "100%", overflow: "auto" }}>
                        <p
                          className={styles.responsiveHeading}
                          style={{ marginTop: "5px", marginBottom: "5px" }}
                        >
                          Support Documents:
                        </p>
                        <FileAttatchmentTable
                          data={this.state.supportingDocumentfiles}
                        />
                      </div>
                    </div>
                  </div>
                )}
              </div>

              {this.state.statusNumber === "9000" &&
                this.state.createdByEmail ===
                  this.props.context.pageContext.user.email && (
                  <div className={styles.sectionContainer}>
                    <button
                      className={styles.header}
                      onClick={() => this._onToggleSection(`markInfo`)}
                    >
                      <Text className={styles.sectionText}>
                        Mark for Information Section
                      </Text>
                      <IconButton
                        iconProps={{
                          iconName: expandSections.markInfo
                            ? "ChevronUp"
                            : "ChevronDown",
                        }}
                        title="Expand/Collapse"
                        ariaLabel="Expand/Collapse"
                        className={styles.chevronIcon}
                      />
                    </button>
                    {expandSections.markInfo && (
                      <div
                        className={`${styles.expansionPanelInside}`}
                        style={{ overflowX: "scroll" }}
                      >
                        <div style={{ padding: "15px" }}>
                          <MarkInfo
                            homePageUrl={this.props.homePageUrl}
                            sp={this.props.sp}
                            context={this.props.context}
                            submitFunctionForMarkInfo={
                              this._handleMarkInfoSubmit
                            }
                            artCommnetsGridData={
                              this.state.noteMarkedInfoDTOState
                            }
                            deletedGridData={(data: any) => {
                              this.setState({
                                noteMarkedInfoDTOState: data,
                              });
                            }}
                            updategirdData={(data: any): void => {
                              const { markInfoassigneeDetails } = data;
                              this.setState((prevState) => ({
                                noteMarkedInfoDTOState: [
                                  ...prevState.noteMarkedInfoDTOState,
                                  markInfoassigneeDetails,
                                ],
                              }));
                            }}
                          />
                        </div>
                      </div>
                    )}
                  </div>
                )}
            </div>

            <div className={styles.pdfContainer}>
              {this.state.pdfLink && this._renderPDFView()}
            </div>
          </div>

          <div className={styles.btnsContainer}>
            {this._checkCurrentRequestIsReturnedOrRejected() &&
              (this._currentUserEmail === this.state.createdByEmail
                ? this._getCallBackAndChangeApproverBtn()
                : this._getReferBackAndApproverStageButtons())}

            {this._checkingCurrentUserIsSecretaryDTO() &&
              this.state.statusNumber !== "5000" &&
              this.state.statusNumber !== "8000" &&
              this.state.statusNumber !== "9000" &&
              this.state.statusNumber !== "4000" && (
                <PrimaryButton
                  iconProps={{ iconName: "Send" }}
                  style={{
                    alignSelf: "flex-end",
                    marginRight: "8px",
                    marginLeft: "8px",
                  }}
                  onClick={async () => {
                    const item = await this._getItemDataSpList(this._itemId);
                    const checkCurrentApproverIsCurrentUser =
                      item?.CurrentApproverId;

                    const _ApproverDTO = JSON.parse(item?.NoteApproversDTO);

                    const secWithApprover = _ApproverDTO.filter(
                      (each: any) =>
                        each.userId === checkCurrentApproverIsCurrentUser &&
                        each.secretaryEmail !== ""
                    );

                    console.log(secWithApprover);

                    if (
                      secWithApprover.length === 0 ||
                      (secWithApprover[0]?.userId !==
                        checkCurrentApproverIsCurrentUser &&
                        secWithApprover[0]?.status !== "Pending")
                    ) {
                      this.setState({
                        hideParellelActionAlertDialog: true,
                        parellelActionAlertMsg:
                          "This request has been taken action by approver.",
                      });

                      return;
                    }

                    if (this.state.errorOfDocuments) {
                      this.setState({ isAutoSaveFailedDialog: true });
                    } else {
                      this.state.secretaryGistDocs.length === 0
                        ? this.setState({ isGistDocEmpty: true })
                        : this.setState({ isGistDocCnrf: true });
                    }
                  }}
                >
                  Submit
                </PrimaryButton>
              )}

            <DefaultButton
              onClick={() => {
                const pageURL: string = this.props.existPageUrl;
                window.location.href = `${pageURL}`;
              }}
              className={`${styles.responsiveButton} `}
              style={{ marginLeft: "10px" }}
              iconProps={{ iconName: "Cancel" }}
            >
              Exit
            </DefaultButton>
          </div>
        </div>
      </div>
    );
  };

  public render(): React.ReactElement<IViewFormProps> {
    console.log(this.state);

 

    return (
      <div className={styles.viewForm}>
        {this.state.isDataLoading ? (
          <div>
            <Modal
              isOpen={this.state.isDataLoading}
              containerClassName={styles.spinnerModalTranparency}
              styles={{
                main: {
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  background: "transparent",
                  boxShadow: "none",
                },
              }}
            >
              <div className="spinner">
                <Spinner
                  label="still loading..."
                  ariaLive="assertive"
                  size={SpinnerSize.large}
                />
              </div>
            </Modal>
          </div>
        ) : (
          this._RenderMainViewForm()
        )}
        {!this.state.dialogFluent && this._DialogBlockingExample()}
      </div>
    );
  }
}
