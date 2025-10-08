import * as React from 'react';
import styles from './MyTasks.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/files/web";
import CustomPopup from './CustomPopup';


interface IMyTaskProps {
  context: WebPartContext;
  description: string;
    siteUrl: string;
    userDisplayName: string;
    isDarkTheme: boolean;
    environmentMessage: string;
    hasTeamsContext: boolean;
}

interface Task {
  projectId: any;
  sno: number;
  projectName: string;
  projectType: string;
  deliverable: string;
  deliverableId: number;
  area: string;
  docType: string;
  docNumber: string;
  revisionnumber: string;
  assignedTo: string;
  org: string;
  status: "Pending" | "Approved" | "In-Progress";
}

interface AuditHistoryItem {
  sno: number;
  approvalLevel: string;
  assignedTo: string;
  assignedToRole: string;
  requestorName: string;
  requestedDate: string;
  actionTakenBy: string;
  actionTakenOn: string;
  remark: string;
  status: string;
}

const MyTask: React.FC<IMyTaskProps> = ({ context }) => {
  const [tasks, setTasks] = React.useState<Task[]>([]);
  const [currentFilter, setCurrentFilter] = React.useState<string>("Pending"); // Set default to "Pending"
  const [filteredTasks, setFilteredTasks] = React.useState<Task[]>([]);
  const [selectedTask, setSelectedTask] = React.useState<Task | null>(null);
  const [showForm, setShowForm] = React.useState<boolean>(false);
  const [auditHistory, setAuditHistory] = React.useState<AuditHistoryItem[]>([]);
  const [showNoAuditHistory, setShowNoAuditHistory] = React.useState<boolean>(false);
  const [selectedFile, setSelectedFile] = React.useState<File | null>(null);
  const [comment, setComment] = React.useState<string>("");
  const [documentControllerId, setDocumentControllerId] = React.useState<number | null>(null);
  const [uploadedFileName, setUploadedFileName] = React.useState<string>("");
  const [uploadedFileUrl, setUploadedFileUrl] = React.useState<string>("");
  // Popup state
const [popup, setPopup] = React.useState<{
  isOpen: boolean;
  type: 'confirmation' | 'validation' | 'success' | 'error';
  title: string;
  message: string;
  onConfirm?: () => void;
}>({
  isOpen: false,
  type: 'success',
  title: '',
  message: ''
});



  // Initialize SP
  const sp = spfi().using(SPFx(context));

  // Function to get Document Controller from ProjectConfiguration list
  const getDocumentController = async (): Promise<number | null> => {
    try {
      const items = await sp.web.lists.getByTitle("ProjectConfiguration").items
        .select(
          "*",
          "DocumentController/Title",
          "DocumentController/EMail",
          "DocumentController/ID"
        )
        .expand("DocumentController")
        .orderBy("Created", false)() // Order by Created desc

      if (items.length > 0) {
        const documentControllerId = items[0].DocumentControllerId;
        console.log("Document Controller ID:", documentControllerId);
        return documentControllerId;
      }
      return null;
    } catch (error) {
      console.error("Error fetching Document Controller:", error);
      return null;
    }
  };

  React.useEffect(() => {
    const initializeData = async () => {
      // Get Document Controller first
      const docControllerId = await getDocumentController();
      setDocumentControllerId(docControllerId);
      
      // Then fetch tasks
      await fetchTasks();
    };

    initializeData();
  }, []);

  const fetchTasks = async () => {
    const currentUserId = (await sp.web.currentUser()).Id;
    try {
      // Fetch main deliverables items
      const items = await sp.web.lists.getByTitle("DeliverablesDetails").items.select(
        "*",
        "ProjectCreationListID/ID",
        "ProjectCreationListID/ProjectName",
        "AssignedTo/Id",
        "AssignedTo/Title"
      ).expand(
        "ProjectCreationListID",
        "AssignedTo"
      ).filter(`AssignedTo/Id eq ${currentUserId}`)();

      // Transform items with additional data from ProjectCreationList
      const transformedTasks: Task[] = await Promise.all(
        items.map(async (item: any, index: number) => {
          try {
            // Fetch additional project creation data with ProjectType lookup expanded
            const creationitem = await sp.web.lists.getByTitle("ProjectCreationList").items.getById(item.ProjectCreationListID.ID)
              .select(
                "*",
                "ProjectType/ProjectType",
                "ProjectType/Id"
              ).expand("ProjectType")();
              console.log('Creation Item:', creationitem);

            return {
              sno: index + 1,
              projectName: item.ProjectCreationListID?.ProjectName || "",
              projectType: creationitem.ProjectType?.ProjectType || "",
              deliverable: item.Deliverables || "",
              deliverableId: item.Id || 0,
              area: item.Area || "",
              docType: item.DocumentType || "",
              docNumber: item.DocNumber || "",
              assignedTo: item.AssignedTo?.Title || "",
              org: item.Organization || "",
              status: item.Status,
              projectId: item.ProjectCreationListID?.ID,
              revisionnumber: item.RevisionNumber || "0",
              creationDate: creationitem.Created || "",
            };
          } catch (error) {
            console.error(`Error fetching creation item for project ${item.ProjectCreationListID?.ID}:`, error);
            // Return a fallback task if creationitem fetch fails
            return {
              sno: index + 1,
              projectName: item.ProjectCreationListID?.ProjectName || "",
              projectType: "Type A",
              deliverable: item.Deliverables || "",
              deliverableId: item.Id || 0,
              area: item.Area || "",
              docType: item.DocumentType || "",
              docNumber: item.DocNumber || "",
              assignedTo: item.AssignedTo?.Title || "",
              org: item.Organization || "",
              status: item.Status,
              projectId: item.ProjectCreationListID?.ID,
              revisionnumber: item.RevisionNumber || "0",
              creationDate: "",
            };
          }
        })
      );

      setTasks(transformedTasks);
      // Set filtered tasks to show only pending tasks by default
      const pendingTasks = transformedTasks.filter(task => task.status === "Pending");
      setFilteredTasks(pendingTasks);
    } catch (error) {
      console.error("Error loading tasks:", error);
    }
  };

  // Count
  const pendingCount = tasks.filter(t => t.status === "Pending").length;
  const completedCount = tasks.filter(t => t.status === "Approved" || t.status === "In-Progress").length;

  const getStatusClass = (status: string) => {
    switch (status) {
      case "Pending": return styles.pendingStatus;
      case "In-Progress": return styles.inProgressStatus;
      case "Approved": return styles.completedStatus;
      default: return styles.pendingStatus;
    }
  };

  const getDeliverablesDetails = (status: string) => {
    setCurrentFilter(status);
    if (status === "All") {
      setFilteredTasks(tasks);
    } else if (status === "Completed" || status === "Approved") {
      setFilteredTasks(tasks.filter(task => task.status === "Approved" || task.status === "In-Progress"));
    } else {
      setFilteredTasks(tasks.filter(task => task.status === status));
    }
  };

  const getAuditHistoryDeliverables = async (item: any, context: WebPartContext): Promise<AuditHistoryItem[]> => {
    const auditHistoryArr: AuditHistoryItem[] = [];

    try {
      const items = await sp.web.lists.getByTitle("ProjectApprovals").items
        .select(
          "*",
          "AssignedTo/Id",
          "AssignedTo/Title",
          "AssignedTo/EMail",
          "DeliverablesDetailsId/Id",
          "DeliverablesDetailsId/Deliverables",
          "Author/Id",
          "Author/Title",
          "Author/EMail"
        )
        .expand(
          "AssignedTo",
          "DeliverablesDetailsId",
          "Author"
        )
        .filter(`DeliverablesDetailsId/ID eq ${item.deliverableId}`)();

      if (items.length > 0) {
        items.forEach((auditHistoryItem: any, index: number) => {
          // Normalize status
          let status = auditHistoryItem.Status;
          if (status && status.toLowerCase() === 'pending') {
            status = 'Pending';
          }

          auditHistoryArr.push({
            sno: index + 1,
            approvalLevel: auditHistoryItem.ApprovalLevel || `Level ${index}`,
            assignedTo: auditHistoryItem.AssignedTo?.Title || 'N/A',
            assignedToRole: auditHistoryItem.ApproverRole || '',
            requestorName: auditHistoryItem.Author?.Title || auditHistoryItem.RequestorName || 'N/A',
            requestedDate: auditHistoryItem.Created ? new Date(auditHistoryItem.Created).toLocaleString('en-GB') : 'N/A',
            actionTakenBy: auditHistoryItem.ActionTakenBy?.Title || auditHistoryItem.ModifiedBy?.Title || 'N/A',
            actionTakenOn: auditHistoryItem.Modified ? new Date(auditHistoryItem.Modified).toLocaleString('en-GB') : '',
            remark: auditHistoryItem.Remarks || '',
            status: status || 'Pending'
          });
        });
      }
    } catch (error) {
      console.error("Error fetching audit history:", error);
    }

    return auditHistoryArr;
  };

  // const handleViewClick = async (task: Task) => {
  //   setSelectedTask(task);
  //   setShowForm(true);

  //   try {
  //     const history = await getAuditHistoryDeliverables(task, context);
  //     setAuditHistory(history);
  //     setShowNoAuditHistory(history.length === 0);
  //   } catch (error) {
  //     console.error("Error loading audit history:", error);
  //     setAuditHistory([]);
  //     setShowNoAuditHistory(true);
  //   }
  // };
const handleViewClick = async (task: Task) => {
  setSelectedTask(task);
  setShowForm(true);
  setUploadedFileName("");
  setUploadedFileUrl("");
  setComment("");

  try {
    // 1️⃣ Load Audit History
    const history = await getAuditHistoryDeliverables(task, context);
    setAuditHistory(history);
    setShowNoAuditHistory(history.length === 0);

    // 2️⃣ Load file and comment if applicable
    if (task.status === "Approved" || task.status === "In-Progress") {
      const deliverablesList = sp.web.lists.getByTitle("DeliverablesDetails");
      const deliverableItem = await deliverablesList.items
        .getById(task.deliverableId)
        .select("DeliverablesDocumentID/ID", "DocumentComments")
        .expand("DeliverablesDocumentID")();

      if (deliverableItem.DocumentComments) {
        setComment(deliverableItem.DocumentComments);
      }

      const deliverablesDocumentId = deliverableItem.DeliverablesDocumentID?.ID;
      if (deliverablesDocumentId) {
        // Get the file directly from the document library
        const fileItem = await sp.web.lists.getByTitle("DeliverablesDocument").items
          .getById(deliverablesDocumentId)
          .select("ID", "FileLeafRef", "File/ServerRelativeUrl")
          .expand("File")();

        if (fileItem.File) {
          setUploadedFileName(fileItem.FileLeafRef);
          
          // Construct the proper file URL
          const serverRelativeUrl = fileItem.File.ServerRelativeUrl;
          const fileUrl = `${context.pageContext.web.absoluteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(serverRelativeUrl)}`;
          
          console.log("File URL:", fileUrl);
          setUploadedFileUrl(fileUrl);
        }
      }
    } else {
      setComment("");
      setUploadedFileName("");
      setUploadedFileUrl("");
    }
  } catch (error) {
    console.error("Error loading task details:", error);
    setAuditHistory([]);
    setShowNoAuditHistory(true);
  }
};



  const handleBackClick = () => {
    setSelectedTask(null);
    setShowForm(false);
    setAuditHistory([]);
    setShowNoAuditHistory(false);
  };


  const handleSubmitClick = async () => {
    if (!selectedTask || !selectedFile) {
      setPopup({
  isOpen: true,
  type: 'validation',
  title: 'Validation',
  message: 'Please select a file before submitting.'
});
return;

    }

    // Check if Document Controller ID is available
    if (!documentControllerId) {
      setPopup({
  isOpen: true,
  type: 'error',
  title: 'Configuration Missing',
  message: 'Document Controller not configured. Please contact administrator.'
});
return;

    }

    try {
      // STEP 1: Upload file to DeliverablesDocument library
      const folder = sp.web.getFolderByServerRelativePath("DeliverablesDocument");
      const uploadResult = await folder.files.addUsingPath(selectedFile.name, selectedFile, { Overwrite: true });

      // Get list item associated with uploaded file
      const fileItem = await uploadResult.file.getItem();
      const uploadedFileItemId = (fileItem as any).Id;
      console.log("Uploaded File Item ID:", uploadedFileItemId);

      // STEP 2: Update DeliverablesDetails item
      await sp.web.lists.getByTitle("DeliverablesDetails").items
        .getById(selectedTask.deliverableId)
        .update({
          DocumentComments: comment,
          RevisionNumber: "0",
          Status: "In-Progress",
          DeliverablesDocumentIDId: uploadedFileItemId
        });

      console.log("DeliverablesDetails updated successfully");

      // STEP 3: Mark Vendor's approval as Completed
      const projectId = selectedTask?.projectId;
      const deliverableId = selectedTask?.deliverableId;
      const approverType = "Vendor";
      const currentUser = await sp.web.currentUser();

      try {
        const items = await sp.web.lists
          .getByTitle("ProjectApprovals")
          .items.select(
            "*",
            "DeliverablesDetailsId/ID",
            "ProjectCreationListID/ID",
            "AssignedTo/ID",
            "AssignedTo/Title",
            "AssignedTo/EMail"
          )
          .expand("DeliverablesDetailsId", "ProjectCreationListID", "AssignedTo")
          .orderBy("SerialNumber", true) // true = ascending order
          .filter(
            `ProjectCreationListID/ID eq ${projectId} and ` +
            `DeliverablesDetailsId/ID eq ${deliverableId} and ` +
            `AssignedTo/ID eq ${currentUser.Id} and ` +
            `ApproverRole eq '${approverType}' and ` +
            `Status eq 'Pending'`
          )();

        if (items.length > 0) {
          const vendorTaskId = items[0].Id;
          await sp.web.lists.getByTitle("ProjectApprovals").items.getById(vendorTaskId).update({
            Status: "Completed",
            Remarks: comment, // your user input
            ApprovalDate: new Date()
          });
          console.log("Vendor task marked as completed");
        } else {
          console.log("No matching ProjectApprovals found for current user");
        }
      } catch (error) {
        console.error("Error updating ProjectApprovals:", error);
      }

      // STEP 4: Create new ProjectApproval item for Document Controller
      await sp.web.lists.getByTitle("ProjectApprovals").items.add({
        DeliverablesDetailsIdId: selectedTask.deliverableId,
        ProjectCreationListIDId: selectedTask.projectId,
        AssignedToId: documentControllerId, // Use the Document Controller ID from ProjectConfiguration
        DocumentType: selectedTask.docType,
        ApproverRole: "Document Controller",
        ProjectType: selectedTask.projectType,
        Level: "Level 1",
        SerialNumber: 0,
        ApprovalCriteria: "Anyone",
        RequestedById: currentUser.Id,
        RequestedDate: new Date(),
        IncomingDate: new Date(),
        RequestedRole: "Vendor",
        RevisionNumber: "0",
        Status: "Pending"
      });

     setPopup({
  isOpen: true,
  type: 'success',
  title: 'Success',
  message: 'Task submitted successfully.',
  onConfirm: () => {
    setPopup(prev => ({ ...prev, isOpen: false }));
    handleBackClick();
    window.location.reload();
  }
});


    } catch (error) {
      console.error("Error in submission:", error);
      setPopup({
  isOpen: true,
  type: 'error',
  title: 'Error',
  message: 'An error occurred during submission. Please check console for details.'
});

    }
  }



  return (
    <div className={styles.myTask}>
      {!showForm ? (
        <>
          <h2>My Tasks</h2>

          {/* Tiles */}
          <div className={styles.tilesContainer}>
            <div
              className={`${styles.tileCard} ${currentFilter === "Pending" ? styles.activeTile : ''}`}
              onClick={() => getDeliverablesDetails("Pending")}
            >
              <div className={styles.tileBody}>
                <div className={styles.tileRow}>
                  <div className={styles.tileIcon}>
                    <div className={`${styles.avatar} ${styles.primary}`}>
                      <span className={styles.icon}>⏱</span>
                    </div>
                  </div>
                  <div className={styles.tileContent}>
                    <h3 className={styles.tileCount}>{pendingCount}</h3>
                    <p className={styles.tileLabel}>Pending</p>
                  </div>
                </div>
              </div>
            </div>

            <div
              className={`${styles.tileCard} ${currentFilter === "Approved" ? styles.activeTile : ''}`}
              onClick={() => getDeliverablesDetails("Approved")}
            >
              <div className={styles.tileBody}>
                <div className={styles.tileRow}>
                  <div className={styles.tileIcon}>
                    <div className={`${styles.avatar} ${styles.success}`}>
                      <span className={styles.icon}>✔️</span>
                    </div>
                  </div>
                  <div className={styles.tileContent}>
                    <h3 className={styles.tileCount}>{completedCount}</h3>
                    <p className={styles.tileLabel}>Completed</p>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* Table */}
          <div className={styles.mainTableContainer}>
            <div className={styles.tableCard}>
              <div className={styles.tableWrapper}>
                <table className={styles.taskTable}>
                  <thead>
                    <tr>
                      <th>S.No</th>
                      <th>Project Name</th>
                      <th>Project Type</th>
                      <th>Deliverable</th>
                      <th>Document Type</th>
                      <th>Assigned To</th>
                      <th>Organization</th>
                      <th>Status</th>
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredTasks.map((task, index) => (
                      <tr key={task.sno}>
                        <td>{index + 1}</td>
                        <td>{task.projectName}</td>
                        <td>{task.projectType}</td>
                        <td>{task.deliverable}</td>
                        <td>{task.docType}</td>
                        <td>{task.assignedTo}</td>
                        <td>{task.org}</td>
                        <td>
                          <span className={`${styles.statusBadge} ${getStatusClass(task.status)}`}>
                            {task.status}
                          </span>
                        </td>
                        <td>
                          <span
                            className={styles.actionIcon}
                            onClick={() => handleViewClick(task)}
                            title={`View ${task.deliverable}`}
                          >
                            👁️
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            <div className={styles.scrollHint}><em>Scroll horizontally → to view all columns</em></div>
          </div>
        </>
      ) : (
        <>
          {/* Form View */}
          <div className={styles.formContainer}>
            <div className={styles.formHeader}>
              <h3>
                My Task &gt;&gt; {selectedTask?.docNumber}
              </h3>
              <button className={styles.backButton} onClick={handleBackClick}>Back</button>
            </div>
            <div className={styles.formView}>
              <div className={styles.formInner}>
                <div className={styles.formBody}>
                  <div className={styles.formGrid}>
                    <div>
                      <label>Project Name</label>
                      <input type="text" value={selectedTask?.projectName || ''} disabled />
                    </div>
                    <div>
                      <label>Project Type</label>
                      <input type="text" value={selectedTask?.projectType || ''} disabled />
                    </div>
                    <div>
                      <label>Client Name</label>
                      <input type="text" value="Client 1" disabled />
                    </div>
                    <div>
                      <label>Project Date</label>
                      <input type="text" value="25/09/2025" disabled />
                    </div>
                    <div>
                      <label>Prepared By</label>
                      <input type="text" value={selectedTask?.assignedTo || ''} disabled />
                    </div>
                    <div>
                      <label>Deliverable</label>
                      <input type="text" value={selectedTask?.deliverable || ''} disabled />
                    </div>
                    <div>
                      <label>Document Type</label>
                      <input type="text" value={selectedTask?.docType || ''} disabled />
                    </div>
                    <div>
                      <label>Document Number</label>
                      <input
                        type="text"
                        value={selectedTask?.docNumber || ''} 
                        disabled
                      />
                    </div>
                    <div>
                      <label>Area</label>
                      <input type="text" value={selectedTask?.area || ''} disabled />
                    </div>
                    <div>
                      <label>Organization</label>
                      <input type="text" value={selectedTask?.org || ''} disabled />
                    </div>
                    <div>
                      <label>Revision Number</label>
                      <input type="text" value="0" disabled />
                    </div>
                  </div>

                  {/* Upload & Comment Section */}
                  <div className={styles.formActions}>
                    {uploadedFileName ? (
                <>
                  <label>Uploaded Document:</label>
                  <p>
                    <a href={uploadedFileUrl} target="_blank" rel="noopener noreferrer">
                      {uploadedFileName}
                    </a>
                  </p>
                </>
              ) : (
                <>
                  <label>Upload Document*</label>
                  <input type="file" onChange={(e) => setSelectedFile(e.target.files?.[0] || null)} />
                </>
              )}
                    <label>Comment*</label>
                    <textarea
                      placeholder="Enter your comment"
                      value={comment}
                      onChange={(e) => setComment(e.target.value)}
                      disabled={selectedTask?.status === "Approved" || selectedTask?.status === "In-Progress"}
                    />
                  </div>

                  {/* Audit History */}
                  {/* <div className={styles.auditHistory}>
                    <h4>Audit History</h4>
                    <table>
                      <thead>
                        <tr>
                          <th>SNo</th>
                          <th>Approval Level</th>
                          <th>Assigned To</th>
                          <th>Assigned To Role</th>
                          <th>Requestor Name</th>
                          <th>Requested Date</th>
                          <th>Action Taken By</th>
                          <th>Action Taken On</th>
                          <th>Remark</th>
                          <th>Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td>1</td>
                          <td>Level 0</td>
                          <td>{selectedTask?.assignedTo}</td>
                          <td>Vendor</td>
                          <td>{selectedTask?.assignedTo}</td>
                          <td>25.09.2025 12:26:53</td>
                          <td>{selectedTask?.assignedTo}</td>
                          <td></td>
                          <td></td>
                          <td>
                            <span className={`${styles.statusBadge} ${getStatusClass(selectedTask?.status || '')}`}>
                              {selectedTask?.status}
                            </span>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div> */}
                  {/* Audit History */}
                  <div className={styles.auditHistory}>
                    <h4>Audit History</h4>
                    {showNoAuditHistory && auditHistory.length === 0 ? (
                      <p>No audit history available</p>
                    ) : (
                      <table>
                        <thead>
                          <tr>
                            <th>SNo</th>
                            <th>Approval Level</th>
                            <th>Assigned To</th>
                            <th>Assigned To Role</th>
                            <th>Requestor Name</th>
                            <th>Requested Date</th>
                            <th>Action Taken By</th>
                            <th>Action Taken On</th>
                            <th>Remark</th>
                            <th>Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          {auditHistory.map((historyItem) => (
                            <tr key={historyItem.sno}>
                              <td>{historyItem.sno}</td>
                              <td>{historyItem.approvalLevel}</td>
                              <td>{historyItem.assignedTo}</td>
                              <td>{historyItem.assignedToRole}</td>
                              <td>{historyItem.requestorName}</td>
                              <td>{historyItem.requestedDate}</td>
                              <td>{historyItem.assignedTo}</td>
                              <td>{historyItem.actionTakenOn}</td>
                              <td>{historyItem.remark}</td>
                              <td>
                                <span className={`${styles.statusBadge} ${getStatusClass(historyItem.status)}`}>
                                  {historyItem.status}
                                </span>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>

                  <div className={styles.formButtons}>
                    {selectedTask?.status === "Pending" && (
                      <>
                        <button className={styles.submitButton} onClick={handleSubmitClick}>Submit</button>
                        <button className={styles.cancelButton} onClick={handleBackClick}>Cancel</button>
                      </>
                    )}
                    {selectedTask?.status === "Approved" && (
                      <button className={styles.cancelButton} onClick={handleBackClick}>Back</button>
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </>
      )}
      <CustomPopup
  isOpen={popup.isOpen}
  type={popup.type}
  title={popup.title}
  message={popup.message}
  onConfirm={popup.onConfirm}
  onCancel={() => setPopup(prev => ({ ...prev, isOpen: false }))}
  onClose={() => setPopup(prev => ({ ...prev, isOpen: false }))}
  onSuccessOk={() => {
    setPopup(prev => ({ ...prev, isOpen: false }));
    if (popup.onConfirm) popup.onConfirm();
  }}
/>

    </div>
  );
};

export default MyTask;