/* eslint-disable no-useless-escape */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @rushstack/no-new-null */
import * as React from "react";
import { IconButton } from "@fluentui/react";


import styles from "../Form.module.scss";
import { v4 } from "uuid";

export interface IUploadFileProps {
  cummulativeError: any;
  typeOfDoc: string;
  onChange: (files: any[] | null, typeOfDoc: string) => void;
  accept?: string;
  maxFileSizeMB: number;
  multiple: boolean;
  
  data: File[];
  errorData: any;
  addtionalData: any[];
}

interface IFileWithError {
  id: string;
  file: File;
  error: string | null;
  cumulativeError: any;
}

interface IUploadFileState {
  selectedFiles: any[];
  cummError: string | null;
  errorOfFile: any;
}

const _randomFileIcon = (docType: string): any => {
 
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


export default class SupportingDocumentsUploadFileComponent extends React.Component<
  IUploadFileProps,
  IUploadFileState
> {
  private fileInputRef: React.RefObject<HTMLInputElement>;

  public constructor(props: IUploadFileProps) {
    super(props);
    this.state = {
      selectedFiles: [],
      cummError: null,
      errorOfFile: null,
    };
    this.fileInputRef = React.createRef<HTMLInputElement>();
  }

  public componentDidMount(): void {
    this.validateFiles(this.props.data);
  }

  public componentDidUpdate(prevProps: IUploadFileProps): void {
    if (prevProps.data !== this.props.data) {
      this.validateFiles(this.props.data);
    }
  }

  private isFileNameValid(name: string): boolean {
    const regex = /^[a-zA-Z0-9._ -]+$/;
    return regex.test(name);
  }

  private validateFiles(files: any[]): void {
    const { maxFileSizeMB } = this.props;
    const maxFileSizeBytes = maxFileSizeMB * 1024 * 1024;

    let validFiles: IFileWithError[] = [];
    let currentTotalSize = 0;
    let cumulativeError = null;

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      let error: string | null = null;

      const allowedFileTypes = [".pdf", ".doc", ".docx", ".xlsx"];
      if (
        !allowedFileTypes.includes(
          file.name.substring(file.name.lastIndexOf("."))
        )
      ) {
        error = "File type is not allowed";
      } else if ((file.size ||file.fileSize) > maxFileSizeBytes) {
        error = `File size should not exceed more ${maxFileSizeMB}MB`;
      } else if (!this.isFileNameValid(file.name)) {
        error = "File name should not contain special characters";
      }

    

      if (
       
        currentTotalSize + (file.size ||file.fileSize) >
        maxFileSizeBytes
      ) {
        cumulativeError =
          "Cumulative size of all the supporting documents should not exceed 25 MB.";
      } else {
        cumulativeError = null;
      }

      currentTotalSize += (file.size ||file.fileSize);
      validFiles.push({
        id: `${file.name}-${i}`,
        file,
        error,
        cumulativeError,
      });
     
      const filterNullerrorInvalidFiles = validFiles.filter((each: any) => {
        return each.error !== null;
      });
     
      this.props.errorData([filterNullerrorInvalidFiles, this.props.typeOfDoc]);
      this.props.cummulativeError(cumulativeError);
      this.setState({ errorOfFile: error, cummError: cumulativeError });
    }

    this.setState({ selectedFiles: validFiles, cummError: cumulativeError });
  }


  

  private handleFileChange =async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    if (e.target.files) {
      const files = Array.from(e.target.files);
     

     
      const hasAdditionalArray =
        this.props.addtionalData && this.props.addtionalData.length > 0;

      const newFiles = files.filter((file) => {
        const isDuplicateInSelectedFiles = this.state.selectedFiles.some(
          (selectedFile) => selectedFile.file.name === file.name
        );

        const isDuplicateInAnotherArray = hasAdditionalArray
          ? this.props.addtionalData.some(
              (anotherFile) => anotherFile.name === file.name
            )
          : false;

        
        return !isDuplicateInSelectedFiles && !isDuplicateInAnotherArray;
      });
      const filePromises = newFiles.map((file) =>
        this.convertToFileArrayBuffer(file)
      );

      


      Promise.all(filePromises)
        .then((fileBuffers) => {
          console.log(fileBuffers)
          const filesWithBuffers = fileBuffers.map((result, index) => ({
           
            
            file: {...result.fileInfo},
            buffer: result.fileInfo.content, // Use the content from fileInfo
            error: null,
            cumulativeError: null,
            ...result.fileInfo,
          }));
          console.log(filesWithBuffers)

          const updatedFiles = this.props.multiple
            ? [...this.state.selectedFiles, ...filesWithBuffers]
            : filesWithBuffers;

            this.setState((prevState) => {
              const updatedFiles = this.props.multiple
                ? [...prevState.selectedFiles, ...filesWithBuffers]
                : filesWithBuffers;
              
              return { selectedFiles: updatedFiles };
            }, () => {
              this.validateFiles(this.state.selectedFiles.map((f) => f.file));
            });
            console.log(updatedFiles)

          

          this.props.onChange(
            updatedFiles.map((f) => f.file),
            this.props.typeOfDoc
          );

          if (this.fileInputRef.current) {
            this.fileInputRef.current.value = "";
          }
        })
        .catch((error) => {
          console.error("Error converting files to ArrayBuffer", error);
        });
    }
  };

  

  private validateFile(file: File): { isValid: boolean; errorMsg: string } {
    const maxFileSizeBytes = 25 * 1024 * 1024; // 25 MB
    const allowedExtensions = [".pdf", ".doc", ".docx", ".xlsx"];
    const validNamePattern = /^[a-zA-Z0-9._ -]+$/;
  
    const fileExt = `.${file.name.split(".").pop()?.toLowerCase()}`;
  
    if (!allowedExtensions.includes(fileExt)) {
      return { isValid: false, errorMsg: "File type is not allowed" };
    }
  
    if (file.size > maxFileSizeBytes) {
      return { isValid: false, errorMsg: `File size should not exceed more than ${this.props.maxFileSizeMB}MB` };
    }
  
    if (!validNamePattern.test(file.name)) {
      return { isValid: false, errorMsg: "File name should not contain special characters" };
    }
  
    return { isValid: true, errorMsg: "" };
  }

  private convertToFileArrayBuffer(file: File): Promise<{
    fileInfo: {
      name: string;
      content: ArrayBuffer | null;
      id: any;
      fileUrl: string;
      ServerRelativeUrl: string;
      isExists: boolean;
      Modified: string;
      isSelected: boolean;
      fileSize: number;
      fileValidation: boolean;
      errormsg: string;
    };
  }> {
    return new Promise((resolve, reject) => {
    
  
      const filesId = v4();
     

      const { isValid, errorMsg } = this.validateFile(file);
  
      // Initial Validation
      let fileValidation = true;
    
      // Read file content if valid
      const reader = new FileReader();
      reader.onload = () => {
        const content = reader.result instanceof ArrayBuffer ? reader.result : null;
  
        const fileInfo = {
          name: file.name,
          content: content,
          id: filesId,
          fileUrl: "",
          ServerRelativeUrl: "",
          isExists: false,
          Modified: new Date().toISOString(),
          isSelected: false,
          fileSize: file.size,
          fileValidation: fileValidation,
          errormsg: isValid ? "" : errorMsg,
        };
  
        resolve({ fileInfo });
      };
  
      reader.onerror = (error) => {
        const fileInfo = {
          name: file.name,
          content: null,
          id: filesId,
          fileUrl: "",
          ServerRelativeUrl: "",
          isExists: false,
          Modified: new Date().toISOString(),
          isSelected: false,
          fileSize: file.size,
          fileValidation: false,
          errormsg: "Error reading file content",
        };
        reject({ fileInfo, error });
      };
  
      reader.readAsArrayBuffer(file);
    });
  }
  

  

  private handleDeleteFile = (fileId: string): void => {
    this.setState((prevState) => {
      const updatedFiles = prevState.selectedFiles.filter(
        (fileWithError) => fileWithError.id !== fileId
      );
      
      const { maxFileSizeMB } = this.props;
      const maxFileSizeBytes = maxFileSizeMB * 1024 * 1024;
      let cumulativeError = null;
      let currentTotalSize = 0;
  
     
      for (let file of updatedFiles) {
        currentTotalSize += file.file.size;
        if (currentTotalSize > maxFileSizeBytes) {
          cumulativeError = "Cumulative size of all the supporting documents should not exceed 25 MB.";
          break; 
        }
      }
  
     
      this.props.errorData([updatedFiles, this.props.typeOfDoc]);
      this.props.cummulativeError(cumulativeError);
  
      return { selectedFiles: updatedFiles };
    }, () => {
      
      this.validateFiles(this.state.selectedFiles.map((f) => f.file));
  
      this.props.errorData([this.state.selectedFiles, this.props.typeOfDoc]);
      this.props.onChange(
        this.state.selectedFiles.map((f) => f.file),
        this.props.typeOfDoc
      );
    });
  };
  

  public render(): React.ReactElement<IUploadFileProps> {
    const { accept, typeOfDoc, multiple } = this.props;
    const { selectedFiles, cummError } = this.state;


    return (
      <ul className={`${styles.fileAttachementsUl}`}>
        <li className={`${styles.basicLi} ${styles.inputField}`}>
          <div style={{ padding: "8px" }}>
            <div>
              <button
                type="button"
                onClick={() => {
                  if (this.fileInputRef.current) {
                    this.fileInputRef.current.click();
                  }
                }}
              >
                Choose File
              </button>

              <input
                type="file"
                ref={this.fileInputRef}
                onChange={this.handleFileChange}
                accept={accept}
                multiple={multiple}
                style={{ display: "none" }}
              />
            </div>

            {typeOfDoc === "supportingDocument" &&
              cummError &&
              cummError.trim() !== "" && (
                <span
                  style={{
                    color: "red",
                    fontSize: "10px",
                   
                    margin: "0px",
                  }}
                >
                  {cummError}
                </span>
              )}
          </div>
        </li>

        {selectedFiles.length > 0 &&
          selectedFiles.map(({ id, file, error }) => {
         
            return (
              <li
                key={id}
                className={`${styles.basicLi} ${styles.attachementli}`}
              >
                <div className={`${styles.fileIconAndNameWithErrorContainer}`}>
                  <img
                  alt="typeOfIcon"
                  
                    src={_randomFileIcon(file.name)}
                    width={32}
                    height={32}
                  />
                 
                  <span className={`${styles.fileNameAndErrorContainer} `}>
                    <span
                      style={{
                        paddingBottom: "0px",
                        marginBottom: "0px",
                        paddingLeft: "4px",
                        whiteSpace: "nowrap",
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        display: "inline-block",
                       
                      }}
                    >
                    
                      {(file.name)}
                    </span>
                    {error && (
                      <span
                        style={{
                          color: "red",
                          fontSize: "10px",
                          paddingLeft: "4px",
                          margin: "0px",
                        }}
                      >
                        {error}
                      </span>
                    )}
                  </span>
                </div>

                <IconButton
                  iconProps={{ iconName: "Cancel" }}
                  title="Delete File"
                  ariaLabel="Delete File"
                  onClick={() => this.handleDeleteFile(id)}
                />
              </li>
            );
          })}
      </ul>
    );
  }
}