/* eslint-disable no-unused-expressions */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  PrimaryButton,
  TextField,
  SelectionMode,
  
  DefaultButton,
  Modal,
  mergeStyleSets,
  IconButton,

} from "@fluentui/react";

import * as React from "react";
import { v4 } from "uuid";
import CommentsMandatoryDialog from "../dialogFluentUi/generalCommentsMandiatoryDialog";
import SpanComponent from "../spanComponent/spanComponent";




interface IGridRow {
  id: string;
  pageNumber:any;
  docReference: any;
  comments: any;
  approverEmailName: string;

}

interface IGridProps {
  
  data: any;
  currentUserDetails: any;
  
  handleCommentDataFuntion: (data: any, action: any, id?: any) => void; 
}

interface IGridState {
  isVisibleAlter: any;
  pageNumValue: string;
  pageValue: string;
  commentValue: string;
  rowsData: IGridRow[];
  editRowId: string;
  isDialogOpen: boolean;
  isEditMode: boolean;
}

const styles = mergeStyleSets({
  modal: {
    minWidth: "300px",
    maxWidth: "80vw",
    width: '100%',
    "@media (min-width: 768px)": {
      maxWidth: "580px",
    },
    "@media (max-width: 767px)": {
      maxWidth: "290px", 
    },
    margin: "auto",
    padding: "10px",
    backgroundColor: "white",
    borderRadius: "4px",
   

    boxShadow: "0 2px 8px rgba(0, 0, 0, 0.26)",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    borderBottom: "1px solid #ddd",
    minHeigth: "50px",
    padding: "5px",
  },
  headerTitle: {
    margin: "5px",
    marginLeft: "5px",
    fontSize: "16px",
    fontWeight: "400",
  },
  headerIcon: {
    paddingRight: "0px", 
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
    justifyContent: "space-between", 

    borderTop: "1px solid #ddd",
    paddingTop: "10px",
    minHeight: "50px",
  },
  button: {
    maxHeight: "32px",
    flex: "1 1 50%", 
    margin: "0 5px", 
  },
  buttonContent: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  buttonIcon: {
    marginRight: "4px", 
  },

  removeTopMargin: {
    marginTop: "4px",
    marginBottom: "14px",
    fontWeight: "400",
  },
});

export default class GeneralCommentsFluentUIGrid extends React.Component<
  IGridProps,
  IGridState
> {
  constructor(props: IGridProps) {
    super(props);
    this.state = {
      isVisibleAlter: false,
      pageNumValue: "",
      pageValue: "",
      commentValue: "",
      rowsData: this._getCurentUserComment(),
      editRowId: "",
      isDialogOpen: false,
      isEditMode: false,
    };
  }

  private _getCurentUserComment = (): IGridRow[] => {
  
    if (this.props.data.length > 0) {
      return this.props.data?.filter(
        (each: any) =>
          each?.approverEmailName === this.props.currentUserDetails.displayName 
      );
    } else {
      return [];
    }
  };

  private handleInputChange = (event: any, field: string) => {
    this.setState({ [field]: event.target.value } as Pick<
      IGridState,
      keyof IGridState
    >);
  };


  private handleAddBtn = () => {
  
    this.setState({
      pageNumValue: "",
      pageValue: "",
      commentValue: "",
      isDialogOpen: true,
      isEditMode: false, 
    });
  };


  private handleSave = () => {
    if (this.state.isEditMode) {
      this.handleSaveBtn();
    } else {
      this.state.commentValue !== ""
        ? this.handleAddNewComment()
        : this.setState({ isVisibleAlter: true });
    }
  };


  private handleAddNewComment = () => {
    const { pageNumValue, pageValue, commentValue } = this.state;

    

    

    const commentsObj: IGridRow = {
      id: v4(),
      pageNumber: pageNumValue || 'N/A',
      docReference: pageValue || 'N/A',
      comments: commentValue,
      approverEmailName: this.props.currentUserDetails.displayName,
      
     
     
      

    };

    this.setState((prevState) => ({
      rowsData: [...prevState.rowsData, commentsObj],
      pageNumValue: "",
      pageValue: "",
      commentValue: "",
      isDialogOpen: false,
    }));

  
    this.props.handleCommentDataFuntion(commentsObj, "add");
  };

  
  private handleEdit = (id: string) => {
    const row = this.state.rowsData.find((each) => each.id === id);
    if (row) {

     
      this.setState({
        pageNumValue: row.pageNumber,
        pageValue: row.docReference,
        commentValue: row.comments,
        editRowId: id,
        isDialogOpen: true,
        isEditMode: true, 
      });
    }
  };

 
  private handleSaveBtn = () => {
    const { editRowId, pageNumValue, pageValue, commentValue } = this.state;

    const updatedRows = this.state.rowsData.map((row) =>
      row.id === editRowId
        ? {
            ...row,
            pageNum: pageNumValue,
            page: pageValue,
            comment: commentValue,
          }
        : row
    ); 

   

    this.setState((prev)=>{

      
    const updatedRows = prev.rowsData.map((row) =>
      row.id === editRowId
        ? {
            ...row,
            pageNum: prev.pageNumValue,
            page: prev.pageValue,
            comment: prev.commentValue,
          }
        : row
    );

      
      
      
      return {
      rowsData: updatedRows,
      editRowId: "",
      pageNumValue: "",
      pageValue: "",
      commentValue: "",
      isDialogOpen: false,
    }})

   
    const updatedRow = updatedRows.find((row) => row.id === editRowId);
    this.props.handleCommentDataFuntion(updatedRow, "edit", editRowId);
  };


  private handleDelete = (id: string) => {
    this.setState((prevState) => {
      const filteredRows = prevState.rowsData.filter((row) => row.id !== id);
      return { rowsData: filteredRows };
    });

  
    this.props.handleCommentDataFuntion(
      this.state.rowsData.filter((item: { id: any }) => {
      
        return item.id === id;
      }),
      "delete",
      id
    );
  };

  private closeDialog = () => {
    this.setState({ isDialogOpen: false });
  };

  public render(): React.ReactElement<any> {
    console.log(this.props)
    const columns: IColumn[] = [
      {
        key: "pageNumber",
        name: "Page#",
        fieldName: "pageNumber",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
      },
      {
        key: "docReference",
        name: "Doc Reference",
        fieldName: "docReference",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
      },
      {
        key: "comments",
        name: "Comment",
        fieldName: "comments",
        minWidth: 200,
        maxWidth: 300,
        isResizable: true,
      },
      {
        key: "actions",
        name: "Actions",
        fieldName: "actions",
        minWidth: 250,

        onRender: (item: IGridRow) => (
          <>
            <PrimaryButton
              text="Edit"
              onClick={() => this.handleEdit(item.id)}
              iconProps={{ iconName: "Edit" }}
            />
            <PrimaryButton
              text="Delete"
              onClick={() => this.handleDelete(item.id)}
              style={{ marginLeft: 8 }}
              iconProps={{ iconName: "Delete" }}
            />
          </>
        ),
      },
    ];
    
    return (
      <div style={{ display: "flex", flexDirection: "column" }}>
      
        <PrimaryButton
          style={{ alignSelf: "flex-end" }}
          text="Add Comment"
          onClick={this.handleAddBtn}
          iconProps={{ iconName: "Comment" }}
        />
        <CommentsMandatoryDialog
          isVisibleAlter={this.state.isVisibleAlter}
          onCloseAlter={() => {
            this.setState({ isVisibleAlter: false });
          }}
          statusOfReq={"undefined"}
        />
       

        <Modal
          isOpen={this.state.isDialogOpen}
          onDismiss={this.closeDialog}
          isBlocking={true}
          containerClassName={styles.modal}
        >
          <div className={styles.header}>
            <div style={{ display: "flex", alignItems: "center" }}>
           
            <IconButton
              iconProps={{ iconName: "CommentAdd" }}
              
            />
            
              <h4 className={styles.headerTitle}>Add Comment</h4>
            </div>
            <IconButton
              iconProps={{ iconName: "Cancel" }}
              onClick={this.closeDialog}
            />
          </div>
          <div className={styles.body}>
            <div
              style={{ width: "90%" }}
            > 
               <label htmlFor="pageNumValue" style={{textAlign:'left', display: 'block', marginBottom: '4px',fontWeight:'400',fontSize:'16px' }}>Page#</label>
              <TextField
              id="pageNumValue"
                
                value={this.state.pageNumValue}
                onChange={(e) => this.handleInputChange(e, "pageNumValue")}
                
              />
               <label htmlFor="pageValue" style={{textAlign:'left', display: 'block', marginBottom: '4px',fontWeight:'400',fontSize:'16px' }}>Doc Reference</label>
              <TextField
                  id="pageValue"
                value={this.state.pageValue}
                onChange={(e) => this.handleInputChange(e, "pageValue")}
              />
               <label style={{ textAlign:'left',display: 'block', marginBottom: '4px',fontWeight:'400',fontSize:'16px' }}>Comment <SpanComponent/></label>
              <TextField
                
                value={this.state.commentValue}
                multiline
                rows={4}
                onChange={(e) => this.handleInputChange(e, "commentValue")}
              />
            </div>
          </div>
          <div className={styles.footer}>
            <PrimaryButton
              text={this.state.isEditMode ? "Save" : "Add"}
              onClick={this.handleSave}
              iconProps={{ iconName: this.state.isEditMode ? "Save" : "Add" }}
              styles={{ root: styles.buttonContent }}
              className={styles.button}
            />
            <DefaultButton
              text="Cancel"
              onClick={this.closeDialog}
              iconProps={{ iconName: "Cancel" }}
              styles={{ root: styles.buttonContent }}
              className={styles.button}
            />
          </div>
        </Modal>

      
        <DetailsList
          items={this.state.rowsData}
          columns={columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.fixedColumns}
          selectionMode={SelectionMode.none}
        />
      </div>
    );
  }
}
