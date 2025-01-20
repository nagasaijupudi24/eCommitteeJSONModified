/* eslint-disable @typescript-eslint/no-unused-vars */

import * as React from "react";
import {
  Modal,
  PrimaryButton,
  IconButton,

  mergeStyleSets,
  
} from "@fluentui/react";


interface MyModalProps {
  hidden: boolean;
  handleDialogBox: () => void;
}




const ApproverOrReviewerModal: React.FC<MyModalProps> = ({
  hidden,
  handleDialogBox,
}) => {
 
  const styles = mergeStyleSets({
        modal: {
          padding: '10px',
          minWidth: '300px',
          maxWidth: '80vw',
          width: '100%',
          '@media (min-width: 768px)': {
            maxWidth: '580px', 
          },
          '@media (max-width: 767px)': {
            maxWidth: '290px', 
          },
          margin: 'auto',
          backgroundColor: 'white',
          borderRadius: '4px',
          boxShadow: '0 2px 8px rgba(0, 0, 0, 0.26)',
        },
        header: {
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
         
          
          borderBottom: '1px solid #ddd',
          minHeight: "50px",
        },
        headerTitle: {
          margin:'5px',
          marginLeft:'5px',
          fontSize:'16px',
          fontWeight:'400'
         },
        body: {
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          textAlign: 'center',
          padding: '20px 0',
        },
        footer: {
          display: 'flex',
          justifyContent: 'flex-end',
          marginTop: '20px',
          borderTop: '1px solid #ddd', 
          paddingTop: '10px',
        },
      });
    
      return (
        <Modal
          isOpen={!hidden}
          onDismiss={handleDialogBox}
          isBlocking={true}
          containerClassName={styles.modal}
        >
          <div className={styles.header}>
            <div style={{ display: 'flex', alignItems: 'center' }}>
              <IconButton iconProps={{ iconName: 'Info' }} />
              <h4 className={styles.headerTitle}>Alert</h4>
            </div>
            <IconButton iconProps={{ iconName: 'Cancel' }} onClick={handleDialogBox} />
          </div>
          <div className={styles.body}>
            <p>The selected reviewer cannot be the same as existing Reviewers/ Approver/ Requester/ Current Actioner.</p>
          </div>
          <div className={styles.footer}>
            <PrimaryButton iconProps={{ iconName: 'ReplyMirrored' }} onClick={()=>{
              handleDialogBox()
              
            }} text="OK" />
          </div>
        </Modal>
      );
  
};

export default ApproverOrReviewerModal;
