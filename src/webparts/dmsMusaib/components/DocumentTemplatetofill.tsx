
let submitnew = require('../assets/submit-new.png')
let cancelnew = require('../assets/cancelnew.png')
let fullw = require('../assets/fulls.png')
let fullw1 = require('../assets/exitf.png')
// import React, { useState, useEffect } from "react";
// import Swal from "sweetalert2";
// import { getSP } from "../loc/pnpjsConfig";
// import Select from "react-select";
// import { SPFI } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items/get-all";
// import "@pnp/sp/items";
// import "@pnp/sp/folders";
// import "@pnp/sp/files";
// import "@pnp/sp/sites";
// import "@pnp/sp/presets/all";
// import "@pnp/sp/site-users/web";
// import "./DocumnetTemplate";

// const DocumentTemplatetofill = (props: any) => {
//       const sp: SPFI = getSP();
//   const [editUrl, setEditUrl] = useState<string | null>(null);
// const [isFullScreen, setIsFullScreen] = useState(false);

//   useEffect(() => {
//     if (props.fileinedit) {
//       // props.fileinedit should already be built as an "edit" URL
//       setEditUrl(props.fileinedit);
//     }
//   }, [props.fileinedit]);

//   const handleSubmit = async () => {
//     try {
//       // This is the source file (template file in Document Template library)
//       const sourceUrl = "/sites/Intranetdemos/Document Template/Transmittal.docx"; 
//       const targetLibrary = "/sites/Intranetdemos/Document Template Destination";
//       const newFileName = `Edited_${Date.now()}.docx`;

  

//       // Get the latest version of the file
//       const file = await sp.web.getFileByServerRelativePath(sourceUrl).getBuffer();
//       const fileBlob = new Blob([file], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
   
//       // Save into target library
//       await sp.web.getFolderByServerRelativePath(targetLibrary)
//         .files.addChunked(newFileName , fileBlob);

//       Swal.fire("Success", "Edited document saved to FinalDocs!", "success");
//     } catch (err: any) {
//       Swal.fire("Error", err.message, "error");
//     }
//   };

//   return (
//     <div>
//       {editUrl && (
//         <div>
//             <div className="mt-3">
//         <button className="btn btn-success" onClick={handleSubmit}>
//           Submit Edited File
//         </button>
//       </div>
//         <iframe
//           src={editUrl}
//           style={{ width: "100%", height: "600px", border: "none" }}
//           title="Editable Document"
//         ></iframe>
//         </div>
      
//       )}

    
//     </div>
//   );
// };

// export default DocumentTemplatetofill;


// 2nd approach
// import React, { useState, useEffect } from "react";
// import Swal from "sweetalert2";
// import { getSP } from "../loc/pnpjsConfig";
// import { SPFI } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items/get-all";
// import "@pnp/sp/items";
// import "@pnp/sp/folders";
// import "@pnp/sp/files";
// import "@pnp/sp/sites";
// import "@pnp/sp/presets/all";
// import "@pnp/sp/site-users/web";
// import "./DocumnetTemplate";

// const DocumentTemplatetofill = (props: any) => {
//   const sp: SPFI = getSP();
//   const [editUrl, setEditUrl] = useState<string | null>(null);
//   const [isFullScreen, setIsFullScreen] = useState(false);

//   useEffect(() => {
//     if (props.fileinedit) {
//       setEditUrl(props.fileinedit);
//     }
//   }, [props.fileinedit]);

//   const handleSubmit = async () => {
//     try {
//       const sourceUrl = "/sites/Intranetdemos/Document Template/Transmittal.docx"; 
//       const targetLibrary = "/sites/Intranetdemos/Document Template Destination";
//       const newFileName = `Edited_${Date.now()}.docx`;

//       const file = await sp.web.getFileByServerRelativePath(sourceUrl).getBuffer();
//       const fileBlob = new Blob([file], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
   
//       await sp.web.getFolderByServerRelativePath(targetLibrary)
//         .files.addChunked(newFileName , fileBlob);

//       Swal.fire("Success", "Edited document saved to FinalDocs!", "success");
//     } catch (err: any) {
//       Swal.fire("Error", err.message, "error");
//     }
//   };

//   return (
//     <div>
//       {editUrl && (
//         <div>
//           <div className="mt-3 d-flex gap-2">
//             <button className="btn btn-success" onClick={handleSubmit}>
//               Submit Edited File
//             </button>
//             <button 
//               className="btn btn-primary"
//               onClick={() => setIsFullScreen(!isFullScreen)}
//             >
//               {isFullScreen ? "Exit Full Screen" : "Full Screen Edit"}
//             </button>
//           </div>

//           <iframe
//             src={editUrl}
//             style={
//               isFullScreen
//                 ? {
//                     position: "fixed",
//                     top: 0,
//                     left: 0,
//                     width: "100%",
//                     height: "100%",
//                     border: "none",
//                     zIndex: 9999,
//                     background: "white",
//                   }
//                 : { width: "100%", height: "600px", border: "none" }
//             }
//             title="Editable Document"
//           ></iframe>
//         </div>
//       )}
//     </div>
//   );
// };

// export default DocumentTemplatetofill;



// import React, { useState, useEffect } from "react";
// import Swal from "sweetalert2";
// import { getSP } from "../loc/pnpjsConfig";
// import { SPFI } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items/get-all";
// import "@pnp/sp/items";
// import "@pnp/sp/folders";
// import "@pnp/sp/files";
// import "@pnp/sp/sites";
// import "@pnp/sp/presets/all";
// import "@pnp/sp/site-users/web";
// import "./DocumnetTemplate";

// const DocumentTemplatetofill = (props: any) => {
//   const sp: SPFI = getSP();
//   const [editUrl, setEditUrl] = useState<string | null>(null);
//   const [isFullScreen, setIsFullScreen] = useState(false);
//  const [approvalhierachy, setapprovalhierachy] = useState<any[]>([]);

//   useEffect(() => {
//     getapprovals();
//   },[])
//   const getapprovals = async () => {
//       const approvalhierachy = await sp.web.lists
//       .getByTitle("DMSFolderPermissionMaster").items.filter(`SiteName eq 'Location' and DocumentLibraryName eq 'Approval section'`).getAll();
//       console.log("approvalhierachy",approvalhierachy);
//       setapprovalhierachy(approvalhierachy);
//   }

//   useEffect(() => {
//     if (props.fileinedit) {
//       setEditUrl(props.fileinedit);
//     }
//   }, [props.fileinedit]);
//   const getUniqueRequestNo = async () => {
//   const counterItem = await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1)();
//   console.log("Counter Item 0", counterItem);
//   console.log("Counter Item 1", counterItem.FileCount);
//   let fileCounter = counterItem.FileCount;

//   // Increment the counter
//   fileCounter++;

//   // Generate the new RequestNo
//   const newRequestNo = `File${String(fileCounter).padStart(2, '0')}`;

//   // Update the counter in the CounterList
//   await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1).update({
//     FileCount: fileCounter
//   });

//   return newRequestNo;
// };
// const handleSubmit = async () => {
//   try {
//     const sourceUrl = "/sites/Intranetdemos/Document Template/Transmittal.docx";
//     const targetLibrary = "/sites/Intranetdemos/Location/Section";
//     const newFileName = `Edited_${Date.now()}.docx`;

//     // Step 1: Download template file
//     const fileBuffer = await sp.web
//       .getFileByServerRelativePath(sourceUrl)
//       .getBuffer();

//     // Step 2: Upload as new file into target library
//    const fileBlob = new Blob([fileBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
// // const uploadedFile = await sp.web
// //   .getFolderByServerRelativePath(targetLibrary)
// //   .files.addChunked(newFileName, fileBlob);
//   let siteID = "9f942ad9-f2b6-4d4a-b99c-80f1171b57e3";
      
//           const testidsub = await sp.site.openWebById(siteID);

//            let uploadedFile2 = await testidsub.web
//     .getFolderByServerRelativePath(
//       "/sites/Intranetdemos/Location/Section"
//     )
//     .files.addUsingPath(newFileName, fileBlob, { Overwrite: true });
//      const item = await uploadedFile2.file.getItem();
//      const timestampSuffix = Date.now().toString().slice(-4);
//      const documentNo = `Transmittal${timestampSuffix}`;
//       await item.update({
//     // ...formValues, // your dynamic fields
//     Status: "Pending", // fixed value
//     // DocumentCategory: props.selectedCategory, // from props
//     // Template: selectedSubCategory.value, // from dropdown
//     DocumentNo:documentNo 
//   });
//   debugger
//       const encodeSharePointURL = (url: string) => {
//         return encodeURIComponent(url)
//             .replace(/'/g, "%27")
//             .replace(/-/g, "%2D")
//             .replace(/_/g, "%5F")
//             .replace(/\./g, "%2E")
//             .replace(/!/g, "%21")
//             .replace(/\*/g, "%2A")
//             .replace(/\(/g, "%28")
//             .replace(/\)/g, "%29")
//             .replace(/~/g, "%7E")
//             .replace(/@/g, "%40")
//             .replace(/\$/g, "%24")
//             .replace(/,/g, "%2C")
//             .replace(/;/g, "%3B")
//             .replace(/:/g, "%3A")
//             .replace(/\+/g, "%2B")
//             .replace(/=/g, "%3D")
//             .replace(/\?/g, "%3F")
//             .replace(/\//g, "%2F")
//             .replace(/#/g, "%23")
//             .replace(/&/g, "%26");
//     };
//      const parentFolder = uploadedFile2.data.ServerRelativeUrl.substring(0, uploadedFile2.data.ServerRelativeUrl.lastIndexOf('/'));
//      const encodedFilePath = encodeSharePointURL(uploadedFile2.data.ServerRelativeUrl);
//     // const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeSharePointURL(parentFolder)}`;
//     const previewUrl = `https://officeindia.sharepoint.com/sites/Intranetdemos/Location/Section/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeSharePointURL(parentFolder)}`;
//      debugger
//       const newRequestNo = await getUniqueRequestNo();
//       debugger
//       try{
    //   const newItem = await sp.web.lists.getByTitle(`DMSLocationFileMaster`).items.add({
    //       FileName: String(uploadedFile2.data.Name),
    //       FileSize: String(uploadedFile2.data.Length),
    //       FileVersion: String(uploadedFile2.data.MajorVersion),
    //       CurrentFolderPath: String("/sites/Intranetdemos/Location/Section"),
    //       FileUID: String(uploadedFile2.data.UniqueId),
    //       CurrentUser: String(props.currentuseremail),
    //       SiteID: String(siteID),
    //       Status: String('Pending'),
    //       FilePreviewURL: String(previewUrl),
    //       DocumentLibraryName: String('Section'),
    //       SiteName: String('Location'),
    //       MyRequest: true,
    //       Processname: 'New File Request',
    //       RequestNo: newRequestNo,
       
    //   });
    //      const uploadfileapproval=   await sp.web.lists.getByTitle('DMSFileApprovalList').items.add({
    //   SiteName: 'Location',
    //   DocumentLibraryName: 'Section',
    //   RequestedBy: props.currentuseremail,
    //   FileName: String(uploadedFile2.data.Name),
    //   FileUID: String(uploadedFile2.data.UniqueId),
    //   FilePreviewUrl: String(previewUrl), // Set this to the correct preview URL if available
    //   Status: String('Pending'),
    //   FolderPath:  "/sites/Intranetdemos/Location/Section", // Set this to the correct folder path if available
    //   ApproveAction: String('Submitted'),
    //   ApprovedLevel: 1,
    //   RequestNo: newRequestNo,
    //   Processname: 'New File Request',
    //   CurrentLevel: 1
    // });
//       debugger
//       }catch (error) {
//         console.log(error , "error adding data into DMSfileMaster ");
//       }
 
    
//   // const uploadedFile = await sp.site
//   // .openWebById("9f942ad9-f2b6-4d4a-b99c-80f1171b57e3")
//   // .then(async ({web}) => {
//   //   const uploadedFile = await web
//   //     .getFolderByServerRelativePath("Section")
//   //     .files.addChunked(newFileName, fileBlob);

//   //   console.log("Uploaded file:", uploadedFile);
//   // });
//     // Step 3: Build Office Online edit URL
//     const absoluteUrl = `${window.location.origin}${uploadedFile2.data.ServerRelativeUrl}`;
//     const editUrl = `https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?src=${encodeURIComponent(absoluteUrl)}`;

//     setEditUrl(editUrl);

//     Swal.fire("Success", "Template copied. You can now edit the document!", "success");
//   } catch (err: any) {
//     Swal.fire("Error", err.message, "error");
//   }
// };


//  const getApprovalTypeText = (type: number) => {
//     return type === 0 ? "Approved by One Only" : "Approved by All";
//   };
//   return (
//     <div>
//       {editUrl && (
//         <div>
//           {/* <div className="mt-3 flex gap-2">
//             <button className="btn btn-success" onClick={handleSubmit}>
//               Submit Edited File
//             </button>

//           <button
//   className="btn btn-primary"
//   onClick={() => setIsFullScreen(!isFullScreen)}
//   style={{
//     position: isFullScreen ? "fixed" : "static",
//     top: isFullScreen ? "10px" : "auto",
//     right: isFullScreen ? "10px" : "auto",
//     zIndex: 10000, // ensure it's above iframe
//   }}
// >
//   {isFullScreen ? "Exit Full Screen" : "Full Screen"}
// </button>
//           </div> */}
// <div className="mt-3 flex gap-2">
//   <button className="btn btn-success" onClick={handleSubmit}>
//     Submit Edited File
//   </button>

//   <button
//     className="btn btn-primary"
//     onClick={() => setIsFullScreen(!isFullScreen)}
//     style={{
//       position: isFullScreen ? "fixed" : "static",
//       top: isFullScreen ? "10px" : "auto",
//       right: isFullScreen ? "10px" : "auto",
//       zIndex: 10000,
//     }}
//   >
//     {isFullScreen ? "Exit Full Screen" : "Full Screen"}
//   </button>
// </div>
//           <iframe
//             src={editUrl}
//             style={{
//               width: isFullScreen ? "100vw" : "100%",
//               height: isFullScreen ? "100vh" : "600px",
//               border: "none",
//               position: isFullScreen ? "fixed" : "relative",
//               top: isFullScreen ? 0 : "auto",
//               left: isFullScreen ? 0 : "auto",
//               zIndex: isFullScreen ? 9999 : "auto",
//               background: "#fff",
//             }}
//             title="Editable Document"
//           ></iframe>
//           <div>
//             <table className="table table-bordered">
//         <thead>
//           <tr>
//             <th>User</th>
//             <th>Level</th>
//             <th>Approval Type</th>
//           </tr>
//         </thead>
//        <tbody>
//   {approvalhierachy && approvalhierachy.length > 0 ? (
//     approvalhierachy.map((item: any) => (
//       <tr key={item.Id}>
//         <td>{item.CurrentUser}</td>
//         <td>{item.Level}</td>
//         <td>{getApprovalTypeText(item.ApprovalType)}</td>
//       </tr>
//     ))
//   ) : (
//     <tr>
//       <td colSpan={3} style={{ textAlign: "center" }}>
//         No approval hierarchy found
//       </td>
//     </tr>
//   )}
// </tbody>
//       </table>
//           </div>
//         </div>
//       )}
//     </div>
//   );
// };


// export default DocumentTemplatetofill;



// just oening preview url in ifrmae but copy file
// import React, { useState, useEffect } from "react";


// import Swal from "sweetalert2";
// import { getSP } from "../loc/pnpjsConfig";
// import { SPFI } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items/get-all";
// import "@pnp/sp/items";
// import "@pnp/sp/folders";
// import "@pnp/sp/files";
// import "@pnp/sp/sites";
// import "@pnp/sp/presets/all";
// import "@pnp/sp/site-users/web";
// import "./DocumnetTemplate";

// const DocumentTemplatetofill = () => {
//     const sp: SPFI = getSP();
//     const [editUrl, setEditUrl] = useState<string | null>(null);
//     const [isFullScreen, setIsFullScreen] = useState(false);
//     const [approvalHierachy, setApprovalHierachy] = useState<any[]>([]);

//     const getUniqueRequestNo = async () => {
//         try {
//             const counterItem = await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1)();
//             const fileCounter = counterItem.FileCount + 1;

//             await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1).update({
//                 FileCount: fileCounter
//             });
//             return `File${String(fileCounter).padStart(2, '0')}`;
//         } catch (error) {
//             console.error("Error getting unique request number:", error);
//             throw new Error("Could not generate a unique request number.");
//         }
//     };

//     const copyFile = async () => {
//         try {
//             const sourceFileUrl = "/sites/Intranetdemos/Document Template/Transmittal.docx";
//             const destinationFolderUrl = "/sites/Intranetdemos/Location/Section";
//             const siteID = "9f942ad9-f2b6-4d4a-b99c-80f1171b57e3";

//             Swal.fire({
//                 title: "Preparing Document...",
//                 text: "Copying template for editing. Please wait.",
//                 allowOutsideClick: false,
//                 didOpen: () => Swal.showLoading(),
//             });

//             // 1. Get the content of the template file.
//             const fileBuffer = await sp.web.getFileByServerRelativePath(sourceFileUrl).getBuffer();
//             const fileBlob = new Blob([fileBuffer]);

//             // 2. Get the web context for the destination subsite.
//             const subsiteSp = await sp.site.openWebById(siteID);

//             // 3. Create a NEW unique file name.
//             const newFileName = `Transmittal_${Date.now()}.docx`;

//             // 4. Upload the file buffer to the destination folder.
//             const uploadedFileResponse = await subsiteSp.web
//                 .getFolderByServerRelativePath(destinationFolderUrl)
//                 .files.addUsingPath(newFileName, fileBlob, { Overwrite: true });

//             // 5. Update metadata for the new file.
//             const newItem = await uploadedFileResponse.file.getItem();
//             await newItem.update({
//                 Status: "Pending",
//                 DocumentNo: `Transmittal${Date.now().toString().slice(-4)}`,
//             });
            
//             // 6. Generate the preview URL in the format you requested
//             const encodeSharePointURL = (url: string) => {
//                 return encodeURIComponent(url)
//                     .replace(/'/g, "%27")
//                     .replace(/-/g, "%2D")
//                     .replace(/_/g, "%5F")
//                     .replace(/\./g, "%2E")
//                     .replace(/!/g, "%21")
//                     .replace(/\*/g, "%2A")
//                     .replace(/\(/g, "%28")
//                     .replace(/\)/g, "%29")
//                     .replace(/~/g, "%7E")
//                     .replace(/@/g, "%40")
//                     .replace(/\$/g, "%24")
//                     .replace(/,/g, "%2C")
//                     .replace(/;/g, "%3B")
//                     .replace(/:/g, "%3A")
//                     .replace(/\+/g, "%2B")
//                     .replace(/=/g, "%3D")
//                     .replace(/\?/g, "%3F")
//                     .replace(/\//g, "%2F")
//                     .replace(/#/g, "%23")
//                     .replace(/&/g, "%26");
//             };
            
//             const fileServerRelativeUrl = uploadedFileResponse.data.ServerRelativeUrl;
//             const parentFolder = fileServerRelativeUrl.substring(0, fileServerRelativeUrl.lastIndexOf('/'));
            
//             const encodedFilePath = encodeSharePointURL(fileServerRelativeUrl);
//             const encodedParentFolder = encodeSharePointURL(parentFolder);
            
//             const previewUrl = `https://officeindia.sharepoint.com/sites/Intranetdemos/Location/Section/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodedParentFolder}`;
//             console.log("Preview URL:", previewUrl);
//             debugger
//             // 7. Set the preview URL for the iframe
//             setEditUrl(previewUrl);
            
//             Swal.close();
//             Swal.fire("Success", "Document copy created successfully!", "success");

//         } catch (err: any) {
//             Swal.close();
//             console.error("Error in copyFile:", err);
//             Swal.fire("Error", `An error occurred while copying the document. ${err.message}`, "error");
//         }
//     };

//     const getApprovals = async () => {
//         try {
//             const hierarchy = await sp.web.lists
//                 .getByTitle("DMSFolderPermissionMaster")
//                 .items.filter(`SiteName eq 'Location' and DocumentLibraryName eq 'Approval section'`)
//                 .getAll();
//             setApprovalHierachy(hierarchy);
//         } catch (error) {
//             console.error("Error fetching approvals:", error);
//         }
//     };

//     useEffect(() => {
//         // Call copyFile when component mounts
//         copyFile();
//         getApprovals();
//     }, []);

//    const handleSubmit = async () => {
//     try {
//         // Get the current editable file information
//         const siteID = "9f942ad9-f2b6-4d4a-b99c-80f1171b57e3";
//         const targetLibrary = "/sites/Intranetdemos/Location/Section";
        
//         // Extract filename from the editUrl
//         const filenameMatch = editUrl?.match(/file=([^&]+)/);
//         const filename = filenameMatch ? decodeURIComponent(filenameMatch[1]) : `Transmittal_${Date.now()}.docx`;
        
//         // Get the file details from the destination library
//         const targetSite = await sp.site.openWebById(siteID);
//         const fileDetails = await targetSite.web
//             .getFileByServerRelativePath(`${targetLibrary}/${filename}`)();
        
//         // Get the list item associated with the file to update its Status
//         const listItem = await targetSite.web
//             .getFileByServerRelativePath(`${targetLibrary}/${filename}`)
//             .getItem();
        
//         // Update the document's Status column to "Pending"
//         await listItem.update({
//             Status: "Pending"
//         });
        
//         // Generate unique request number
//         const newRequestNo = await getUniqueRequestNo();
        
//         // Generate preview URL for the file
//         const encodeSharePointURL = (url: string) => {
//             return encodeURIComponent(url)
//                 .replace(/'/g, "%27")
//                 .replace(/-/g, "%2D")
//                 .replace(/_/g, "%5F")
//                 .replace(/\./g, "%2E")
//                 .replace(/!/g, "%21")
//                 .replace(/\*/g, "%2A")
//                 .replace(/\(/g, "%28")
//                 .replace(/\)/g, "%29")
//                 .replace(/~/g, "%7E")
//                 .replace(/@/g, "%40")
//                 .replace(/\$/g, "%24")
//                 .replace(/,/g, "%2C")
//                 .replace(/;/g, "%3B")
//                 .replace(/:/g, "%3A")
//                 .replace(/\+/g, "%2B")
//                 .replace(/=/g, "%3D")
//                 .replace(/\?/g, "%3F")
//                 .replace(/\//g, "%2F")
//                 .replace(/#/g, "%23")
//                 .replace(/&/g, "%26");
//         };
        
//         const fileServerRelativeUrl = fileDetails.ServerRelativeUrl;
//         const parentFolder = fileServerRelativeUrl.substring(0, fileServerRelativeUrl.lastIndexOf('/'));
//         const encodedFilePath = encodeSharePointURL(fileServerRelativeUrl);
//         const encodedParentFolder = encodeSharePointURL(parentFolder);
//         const previewUrl = `https://officeindia.sharepoint.com/sites/Intranetdemos/Location/Section/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodedParentFolder}`;
         
//         const getcurrentuseremail = async () => {
//             const currentUser = await sp.web.currentUser();
//             return currentUser.Email;
//         }
//         const currentUserEmail = await getcurrentuseremail();

//         // Add to DMSLocationFileMaster list
//         const newItem = await sp.web.lists.getByTitle(`DMSLocationFileMaster`).items.add({
//             FileName: String(fileDetails.Name),
//             FileSize: String(fileDetails.Length),
//             FileVersion: String(fileDetails.MajorVersion),
//             CurrentFolderPath: String(targetLibrary),
//             FileUID: String(fileDetails.UniqueId),
//             CurrentUser: String(currentUserEmail),
//             SiteID: String(siteID),
//             Status: 'Pending',
//             FilePreviewURL: String(previewUrl),
//             DocumentLibraryName: 'Section',
//             SiteName: 'Location',
//             MyRequest: true,
//             Processname: 'New File Request',
//             RequestNo: newRequestNo,
//         });
        
//         // Add to DMSFileApprovalList list
//         const uploadfileapproval = await sp.web.lists.getByTitle('DMSFileApprovalList').items.add({
//             SiteName: 'Location',
//             DocumentLibraryName: 'Section',
//             RequestedBy: String(currentUserEmail),
//             FileName: String(fileDetails.Name),
//             FileUID: String(fileDetails.UniqueId),
//             FilePreviewUrl: String(previewUrl),
//             Status: 'Pending',
//             FolderPath: targetLibrary,
//             ApproveAction: 'Submitted',
//             ApprovedLevel: 1,
//             RequestNo: newRequestNo,
//             Processname: 'New File Request',
//             CurrentLevel: 1
//         });
        
//         Swal.fire("Success", "Document submitted for approval!", "success");
        
//     } catch (err: any) {
//         console.error("Error in handleSubmit:", err);
//         Swal.fire("Error", `An error occurred while submitting the document. ${err.message}`, "error");
//     }
// };

//     const getApprovalTypeText = (type: number) => {
//         return type === 0 ? "Approved by One Only" : "Approved by All";
//     };

//     return (
//         <div>
//             {editUrl ? (
//                 <div>
//                     <div className="mt-3 flex gap-2">
//                         <button className="btn btn-primary" onClick={handleSubmit}>
//                             Submit Edited File
//                         </button>
//                         <button
//                             className="btn btn-primary"
//                             onClick={() => setIsFullScreen(!isFullScreen)}
//                             style={{
//                                 position: isFullScreen ? "fixed" : "static",
//                                 top: isFullScreen ? "10px" : "auto",
//                                 right: isFullScreen ? "10px" : "auto",
//                                 zIndex: 10000,
//                             }}
//                         >
//                             {isFullScreen ? "Exit Full Screen" : "Full Screen"}
//                         </button>
//                     </div>
//                     <iframe
//                         src={editUrl}
//                         style={{
//                             width: isFullScreen ? "100vw" : "100%",
//                             height: isFullScreen ? "100vh" : "600px",
//                             border: "none",
//                             position: isFullScreen ? "fixed" : "relative",
//                             top: isFullScreen ? 0 : "auto",
//                             left: isFullScreen ? 0 : "auto",
//                             zIndex: isFullScreen ? 9999 : "auto",
//                             background: "#fff",
//                         }}
//                         title="Editable Document"
//                     ></iframe>
//                     <div>
//                         <table className="table table-bordered">
//                             <thead>
//                                 <tr>
//                                     <th>User</th>
//                                     <th>Level</th>
//                                     <th>Approval Type</th>
//                                 </tr>
//                             </thead>
//                             <tbody>
//                                 {approvalHierachy && approvalHierachy.length > 0 ? (
//                                     approvalHierachy.map((item: any) => (
//                                         <tr key={item.Id}>
//                                             <td>{item.CurrentUser}</td>
//                                             <td>{item.Level}</td>
//                                             <td>{getApprovalTypeText(item.ApprovalType)}</td>
//                                         </tr>
//                                     ))
//                                 ) : (
//                                     <tr>
//                                         <td colSpan={3} style={{ textAlign: "center" }}>
//                                             No approval hierarchy found
//                                         </td>
//                                     </tr>
//                                 )}
//                             </tbody>
//                         </table>
//                     </div>
//                 </div>
//             ) : (
//                 <div className="mt-3">
//                     <p>Loading document from template...</p>
//                 </div>
//             )}
//         </div>
//     );
// };

// export default DocumentTemplatetofill;


// her it was updating file but after submit lock file error 
// import React, { useState, useEffect } from "react";
// import Swal from "sweetalert2";
// import { getSP } from "../loc/pnpjsConfig";
// import { SPFI } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items/get-all";
// import "@pnp/sp/items";
// import "@pnp/sp/folders";
// import "@pnp/sp/files";
// import "@pnp/sp/sites";
// import "./DocumnetTemplate.css";

// const DocumentTemplatetofill = () => {
//     const sp: SPFI = getSP();
//     const [editUrl, setEditUrl] = useState<string | null>(null);
//     const [isFullScreen, setIsFullScreen] = useState(false);
//     const [approvalHierachy, setApprovalHierachy] = useState<any[]>([]);
//     const [newFileName, setNewFileName] = useState<string>("");
//     const [newFileItemId, setNewFileItemId] = useState<number>(0);

//     const copyFile = async () => {
//         try {
//             const sourceFileUrl = "/sites/Intranetdemos/Document Template/Transmittal.docx";
//             const destinationFolderUrl = "/sites/Intranetdemos/Location/Section";
//             const siteID = "9f942ad9-f2b6-4d4a-b99c-80f1171b57e3";

//             Swal.fire({
//                 title: "Preparing Document...",
//                 text: "Copying template for editing. Please wait.",
//                 allowOutsideClick: false,
//                 didOpen: () => Swal.showLoading(),
//             });

//             // 1. Get the content of the template file.
//             const fileBuffer = await sp.web.getFileByServerRelativePath(sourceFileUrl).getBuffer();
//             const fileBlob = new Blob([fileBuffer]);

//             // 2. Get the web context for the destination subsite.
//             const subsiteSp = await sp.site.openWebById(siteID);

//             // 3. Create a NEW unique file name.
//             const uniqueFileName = `Transmittal_${Date.now()}.docx`;
//             setNewFileName(uniqueFileName);

//             // 4. Upload the file buffer to the destination folder.
//             const uploadedFileResponse = await subsiteSp.web
//                 .getFolderByServerRelativePath(destinationFolderUrl)
//                 .files.addUsingPath(uniqueFileName, fileBlob, { Overwrite: true });

//             // 5. Update metadata for the new file.
//             const newItem = await uploadedFileResponse.file.getItem();
//             const documentNo = `Transmittal${Date.now().toString().slice(-4)}`;
            
//             await newItem.update({
//                 Status: "Pending",
//                 DocumentNo: documentNo,
//             });
            
//             // Get the updated item to retrieve the ID
//        const updatedItem: any = await uploadedFileResponse.file.getItem();
//        console.log("Updated item:", updatedItem);
//        debugger
//       setNewFileItemId(updatedItem.Id);
            
//             // 6. Get the file's unique ID (GUID)
//             const fileProperties = await uploadedFileResponse.file.getItem()
//             const fileUniqueId = uploadedFileResponse.data.UniqueId;
            
//             // 7. Generate the correct EDIT URL with GUID format
//             const fileServerRelativeUrl = uploadedFileResponse.data.ServerRelativeUrl;
//             const listId = "Section"; // This should be the actual list ID if available
            
//             // Format the URL as you specified
//             const editFileUrl = `https://officeindia.sharepoint.com/:w:/r/sites/Intranetdemos${encodeURIComponent(fileServerRelativeUrl.split('/sites/Intranetdemos')[1])}?web=1`;
            
//             console.log("Edit URL:", editFileUrl);
//             debugger
//             // Alternative approach if the above doesn't work
//             // const editFileUrl = `https://officeindia.sharepoint.com/:w:/r/sites/Intranetdemos/Location/_layouts/15/Doc.aspx?sourcedoc=%7B${fileUniqueId}%7D&action=edit&uid=%7B${fileUniqueId}%7D&ListItemId=${updatedItem.Id}&ListId=%2Fsites%2FIntranetdemos%2FLocation%2FSection&odsp=1&env=prod`;
            
//             // 8. Set the edit URL for the iframe
//             setEditUrl(editFileUrl);
            
//             Swal.close();
//             Swal.fire("Success", "Document copy created successfully!", "success");

//         } catch (err: any) {
//             Swal.close();
//             console.error("Error in copyFile:", err);
//             Swal.fire("Error", `An error occurred while copying the document. ${err.message}`, "error");
//         }
//     };

// const handleSubmit = async () => {
//     try {
//         Swal.fire({
//             title: "Submitting Document",
//             text: "Please wait while we submit your document for approval.",
//             allowOutsideClick: false,
//             didOpen: () => Swal.showLoading(),
//         });
        
//         // Update the file status to submitted
//         const siteID = "9f942ad9-f2b6-4d4a-b99c-80f1171b57e3";
//         const subsiteSp = await sp.site.openWebById(siteID);
        
//         // First, ensure the file is checked in (to resolve the lock issue)
//         try {
//             const file = await subsiteSp.web.getFileByServerRelativePath(`/sites/Intranetdemos/Location/Section/${newFileName}`);
//             await file.checkin("Document submitted for approval", 1); // 1 = Major check-in
//             console.log("File checked in successfully");
//         } catch (checkinError) {
//             console.log("File may not have been checked out or already checked in:", checkinError);
//             // Continue with submission even if checkin fails
//         }
        
//         // Now update the item metadata
//         await subsiteSp.web.lists.getByTitle("Section").items.getById(newFileItemId).update({
//             Status: "Submitted",
//         });
        
//         // Here you would typically also create approval tasks, send notifications, etc.
        
//         Swal.close();
//         Swal.fire("Success", "Document submitted for approval!", "success");
        
//         // Redirect or show a success message
//         setTimeout(() => {
//             // You might want to redirect to a different page or reset the component
//             setEditUrl(null);
//         }, 2000);
        
//     } catch (error: any) {
//         Swal.close();
//         console.error("Error submitting document:", error);
//         Swal.fire("Error", `An error occurred while submitting the document: ${error.message}`, "error");
//     }
// };

//     const getApprovals = async () => {
//         try {
//             const hierarchy = await sp.web.lists
//                 .getByTitle("DMSFolderPermissionMaster")
//                 .items.filter(`SiteName eq 'Location' and DocumentLibraryName eq 'Approval section'`)
//                 .getAll();
//             setApprovalHierachy(hierarchy);
//         } catch (error) {
//             console.error("Error fetching approvals:", error);
//         }
//     };

//     useEffect(() => {
//         // Call copyFile when component mounts
//         copyFile();
//         getApprovals();
//     }, []);

//     const getApprovalTypeText = (type: number) => {
//         return type === 0 ? "Approved by One Only" : "Approved by All";
//     };

//     return (
//         <div className="document-editor-container">
//             {editUrl ? (
//                 <div>
//                     <div className="editor-controls mt-3 flex gap-2">
//                         <button className="btn btn-primary" onClick={handleSubmit}>
//                             Submit Edited File
//                         </button>
//                         <button
//                             className="btn btn-secondary"
//                             onClick={() => setIsFullScreen(!isFullScreen)}
//                             style={{
//                                 position: isFullScreen ? "fixed" : "static",
//                                 top: isFullScreen ? "10px" : "auto",
//                                 right: isFullScreen ? "10px" : "auto",
//                                 zIndex: 10000,
//                             }}
//                         >
//                             {isFullScreen ? "Exit Full Screen" : "Full Screen"}
//                         </button>
//                     </div>
                    
//                     <div className="editor-frame-container">
//                         <iframe
//                             src={editUrl}
//                             style={{
//                                 width: isFullScreen ? "100vw" : "100%",
//                                 height: isFullScreen ? "100vh" : "70vh",
//                                 border: "none",
//                                 position: isFullScreen ? "fixed" : "relative",
//                                 top: isFullScreen ? 0 : "auto",
//                                 left: isFullScreen ? 0 : "auto",
//                                 zIndex: isFullScreen ? 9999 : "auto",
//                                 background: "#fff",
//                             }}
//                             title="Editable Document"
//                             allowFullScreen
//                         ></iframe>
//                     </div>
                    
//                     <div className="approval-section mt-4">
//                         <h3>Approval Hierarchy</h3>
//                         <table className="table table-bordered">
//                             <thead>
//                                 <tr>
//                                     <th>User</th>
//                                     <th>Level</th>
//                                     <th>Approval Type</th>
//                                 </tr>
//                             </thead>
//                             <tbody>
//                                 {approvalHierachy && approvalHierachy.length > 0 ? (
//                                     approvalHierachy.map((item: any) => (
//                                         <tr key={item.Id}>
//                                             <td>{item.CurrentUser}</td>
//                                             <td>{item.Level}</td>
//                                             <td>{getApprovalTypeText(item.ApprovalType)}</td>
//                                         </tr>
//                                     ))
//                                 ) : (
//                                     <tr>
//                                         <td colSpan={3} style={{ textAlign: "center" }}>
//                                             No approval hierarchy found
//                                         </td>
//                                     </tr>
//                                 )}
//                             </tbody>
//                         </table>
//                     </div>
//                 </div>
//             ) : (
//                 <div className="loading-container mt-3">
//                     <p>Loading document from template...</p>
//                     <div className="spinner-border text-primary" role="status">
//                         <span className="visually-hidden">Loading...</span>
//                     </div>
//                 </div>
//             )}
//         </div>
//     );
// };

// export default DocumentTemplatetofill;


import React, { useState, useEffect , useRef} from "react";
import Select from "react-select";

import Swal from "sweetalert2";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items/get-all";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/sites";
import "./DocumnetTemplate.css";
import JSZip from "jszip";
import { set } from "@microsoft/sp-lodash-subset";
const DocumentTemplatetofill = () => {
    const sp: SPFI = getSP();
    const [editUrl, setEditUrl] = useState<string | null>(null);
    const [isFullScreen, setIsFullScreen] = useState(false);
    const [approvalHierachy, setApprovalHierachy] = useState<any[]>([]);
    const [newFileName, setNewFileName] = useState<string>("");
    const [newFileItemId, setNewFileItemId] = useState<number>(0);
    const [fileServerRelativeUrl, setFileServerRelativeUrl] = useState<string>("");
    const [selectedfile, setselectedFile] = useState<any>(null);
    const [uploadedFileResponse, setUploadedFileResponse] = useState<any>(null);

     const currentUserEmailRef = useRef('');
     const currentUserIDref = useRef<number>(0);
     const currentUserTitleRef = useRef('');
const [companyMasterList, setCompanyMasterList] = useState<any[]>([]);
const [classificationList, setClassificationList] = useState<any[]>([]);
const [selectedCompany, setSelectedCompany] = useState<any>(null);
const [selectedClassification, setSelectedClassification] = useState<any>(null);
const [fileCounter, setFileCounter] = useState<number>(0);

    // Add these state variables after your existing useState declarations
    const [allUsers, setAllUsers] = useState<any[]>([]);
    const [selectedFromUser, setSelectedFromUser] = useState<any>(null);
    const [selectedDate, setSelectedDate] = useState<string>("");
const [reference, setReference] = useState<string>("");

    const [selectedToUser, setSelectedToUser] = useState<any>(null);
    
const [newRequestNo, setNewRequestNo] = useState<string>("");
//     const copyFile = async () => {
//         try {
//             const sourceFileUrl = "/sites/Intranetdemos/Document Template/Transmittal.docx";
//             const destinationFolderUrl = "/sites/Intranetdemos/Location/Section";
//             const siteID = "9f942ad9-f2b6-4d4a-b99c-80f1171b57e3";

//             Swal.fire({
//                 title: "Preparing Document...",
//                 text: "Please wait.",
//                 allowOutsideClick: false,
//                 didOpen: () => Swal.showLoading(),
//             });

//             // 1. Get the content of the template file.
//             const fileBuffer = await sp.web.getFileByServerRelativePath(sourceFileUrl).getBuffer();
//             const fileBlob = new Blob([fileBuffer]);

//             // 2. Get the web context for the destination subsite.
//             const subsiteWeb = await sp.site.openWebById(siteID);

//             // 3. Create a NEW unique file name.
//             const uniqueFileName = `Transmittal_${Date.now()}.docx`;
//             setNewFileName(uniqueFileName);

//             // 4. Upload the file buffer to the destination folder.
//             let uploadedFileResponse = await subsiteWeb.web
//                 .getFolderByServerRelativePath(destinationFolderUrl)
//                 .files.addUsingPath(uniqueFileName, fileBlob, { Overwrite: true });

//             // Store the server relative URL for later use
//             setFileServerRelativeUrl(uploadedFileResponse.data.ServerRelativeUrl);

//             // 5. Get the item ID from the file
//             const fileItem = await uploadedFileResponse.file.getItem();
//             // const itemId = fileItem.ID; // Use ID instead of Id
       
//             // 6. Update metadata for the new file.
//             const documentNo = `Transmittal${Date.now().toString().slice(-4)}`;
//               const listItem = await uploadedFileResponse.file.getItem();
//                await listItem.update({ Status: "Pending" });
                  

//             // await subsiteWeb.web.lists.getByTitle("Section").items.getById(itemId).update({
//             //     Status: "Pending",
//             //     DocumentNo: documentNo,
//             // });
            
//             // 7. Generate the EDIT URL
//             const editFileUrl = `https://officeindia.sharepoint.com/:w:/r/sites/Intranetdemos${encodeURIComponent(uploadedFileResponse.data.ServerRelativeUrl.split('/sites/Intranetdemos')[1])}?web=1`;
            
//             // 8. Set the edit URL for the iframe
//             setEditUrl(editFileUrl);
            
//             Swal.close();
           

//         } catch (err: any) {
//             Swal.close();
//             console.error("Error in copyFile:", err);
//             Swal.fire("Error", `An error occurred while copying the document. ${err.message}`, "error");
//         }
//     };

//     const handleSubmit = async () => {
//         try {
//             Swal.fire({
//                 title: "Submitting Document",
//                 text: "Please wait while we submit your document for approval.",
//                 allowOutsideClick: false,
//                 didOpen: () => Swal.showLoading(),
//             });
            
//             const siteID = "9f942ad9-f2b6-4d4a-b99c-80f1171b57e3";
//             const subsiteWeb = await sp.site.openWebById(siteID);

           

//    const getUniqueRequestNo = async () => {
//         try {
//             const counterItem = await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1)();
//             const fileCounter = counterItem.FileCount + 1;

//             await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1).update({
//                 FileCount: fileCounter
//             });
//             return `File${String(fileCounter).padStart(2, '0')}`;
//         } catch (error) {
//             console.error("Error getting unique request number:", error);
//             throw new Error("Could not generate a unique request number.");
//         }
//     };
//  const newRequestNo = await getUniqueRequestNo();

//         const encodeSharePointURL = (url: string) => {
//             return encodeURIComponent(url)
//                 .replace(/'/g, "%27")
//                 .replace(/-/g, "%2D")
//                 .replace(/_/g, "%5F")
//                 .replace(/\./g, "%2E")
//                 .replace(/!/g, "%21")
//                 .replace(/\*/g, "%2A")
//                 .replace(/\(/g, "%28")
//                 .replace(/\)/g, "%29")
//                 .replace(/~/g, "%7E")
//                 .replace(/@/g, "%40")
//                 .replace(/\$/g, "%24")
//                 .replace(/,/g, "%2C")
//                 .replace(/;/g, "%3B")
//                 .replace(/:/g, "%3A")
//                 .replace(/\+/g, "%2B")
//                 .replace(/=/g, "%3D")
//                 .replace(/\?/g, "%3F")
//                 .replace(/\//g, "%2F")
//                 .replace(/#/g, "%23")
//                 .replace(/&/g, "%26");
//         };
//         //  Extract filename from the editUrl
//         const filenameMatch = editUrl?.match(/file=([^&]+)/);
//         console.log("filenameMatch:", filenameMatch);
//         debugger
//         const filename = filenameMatch ? decodeURIComponent(filenameMatch[1]) : `Transmittal_${Date.now()}.docx`;
//         debugger
//         // Get the file details from the destination library
//         const targetSite = await sp.site.openWebById(siteID);
//          const targetLibrary = "/sites/Intranetdemos/Location/Section";
//         const fileDetails = await targetSite.web
//             .getFileByServerRelativePath(`${targetLibrary}/${filename}`)();
//         const fileServerRelativeUrl = fileDetails.ServerRelativeUrl;
//         const parentFolder = fileServerRelativeUrl.substring(0, fileServerRelativeUrl.lastIndexOf('/'));
//         const encodedFilePath = encodeSharePointURL(fileServerRelativeUrl);
//         const encodedParentFolder = encodeSharePointURL(parentFolder);
//         const previewUrl = `https://officeindia.sharepoint.com/sites/Intranetdemos/Location/Section/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodedParentFolder}`;
//         console.log(previewUrl , "previewurl")
//         debugger
//               const newItem = await sp.web.lists.getByTitle(`DMSLocationFileMaster`).items.add({
//           FileName: String(uploadedFileResponse.data.Name),
//           FileSize: String(uploadedFileResponse.data.Length),
//           FileVersion: String(uploadedFileResponse.data.MajorVersion),
//           CurrentFolderPath: String("/sites/Intranetdemos/Location/Section"),
//           FileUID: String(uploadedFileResponse.data.UniqueId),
//           CurrentUser: String(props.currentuseremail),
//           SiteID: String(siteID),
//           Status: String('Pending'),
//           FilePreviewURL: String(previewUrl),
//           DocumentLibraryName: String('Section'),
//           SiteName: String('Location'),
//           MyRequest: true,
//           Processname: 'New File Request',
//           RequestNo: newRequestNo,
       
//       });

//         const uploadfileapproval=   await sp.web.lists.getByTitle('DMSFileApprovalList').items.add({
//       SiteName: 'Location',
//       DocumentLibraryName: 'Section',
//       RequestedBy: props.currentuseremail,
//       FileName: String(uploadedFileResponse.data.Name),
//       FileUID: String(uploadedFileResponse.data.UniqueId),
//       FilePreviewUrl: String(previewUrl), // Set this to the correct preview URL if available
//       Status: String('Pending'),
//       FolderPath:  "/sites/Intranetdemos/Location/Section", // Set this to the correct folder path if available
//       ApproveAction: String('Submitted'),
//       ApprovedLevel: 1,
//       RequestNo: newRequestNo,
//       Processname: 'New File Request',
//       CurrentLevel: 1
//     });
//             // First, try to release the file lock using a direct API call
//             try {
//                 await releaseFileLock(subsiteWeb, fileServerRelativeUrl);
//             } catch (lockError) {
//                 console.log("File lock release failed, trying to update metadata anyway:", lockError);
//             }
            
//             // Now update the item metadata

            
//             // await subsiteWeb.web.lists.getByTitle("Section").items.getById(newFileItemId).update({
//             //     Status: "Submitted",
//             // });
            
//             // Here you would typically also create approval tasks, send notifications, etc.
            
//             Swal.close();
//             Swal.fire({
//   icon: 'success',
//   title: 'Document submitted ',
//   showConfirmButton: true,
//   confirmButtonText: 'OK'
// }).then((result) => {
//   if (result.isConfirmed) {
//     window.location.reload();
//   }
// });
            
//             // Redirect or show a success message
//             setTimeout(() => {
//                 // You might want to redirect to a different page or reset the component
//                 setEditUrl(null);
//             }, 2000);
            
//         } catch (error: any) {
//             Swal.close();
//             console.error("Error submitting document:", error);
            
//             if (error.message.includes("locked for shared use") || error.message.includes("423")) {
//                 Swal.fire({
//                     title: "File Still in Use",
//                     html: `The document appears to be still open.<br><br>
//                            <strong>Please make sure:</strong>
//                            <ol>
//                              <li>You've saved all changes in Word Online</li>
//                              <li>You've closed the document in Word Online</li>
//                              <li>You're not viewing the document in another browser tab</li>
//                            </ol>
//                            Then try submitting again.`,
//                     icon: "warning",
//                     confirmButtonText: "Try Again",
//                     showCancelButton: true,
//                     cancelButtonText: "Cancel"
//                 }).then((result) => {
//                     if (result.isConfirmed) {
//                         handleSubmit();
//                     }
//                 });
//             } else {
//                 Swal.fire("Error", `An error occurred while submitting the document: ${error.message}`, "error");
//             }
//         }
//     };

    // Special function to release file lock using direct API call
    // Generate request number when company or classification changes
    // Fetch data in useEffect
useEffect(() => {
    const fetchData = async () => {
        try {
            // Fetch company master and classification
            const companymaster = await sp.web.lists
                .getByTitle("ESSAcompanymaster")
                .items.select("*")
                .getAll();
            
            const companyclassification = await sp.web.lists
                .getByTitle("ESSAcompanyclassification")
                .items.select("*")
                .getAll();
            
            console.log(companymaster, "companymaster");
            console.log(companyclassification, "companyclassification");
            
            // Transform to dropdown format
            const companyOptions = companymaster.map((item) => ({
                value: item.CompanyNameShort,
                label: item.CompanyDescription,
                id: item.ID
            }));
            
            const classificationOptions = companyclassification.map((item) => ({
                value: item.classificationshort,
                label: item.classification,
                id: item.ID
            }));
            
            setCompanyMasterList(companyOptions);
            setClassificationList(classificationOptions);
            
            // Fetch counter ONCE
            const counterItem = await sp.web.lists
                .getByTitle("DMSFileCounterList")
                .items.getById(1)();
            
            const counter = counterItem.FileCount + 1;
            setFileCounter(counter);
            
            // Update counter in database
            await sp.web.lists
                .getByTitle("DMSFileCounterList")
                .items.getById(1)
                .update({ FileCount: counter });
            
            console.log("File Counter fetched:", counter);
            
        } catch (error) {
            console.error("Error fetching data:", error);
        }
    };
    
    fetchData();
}, []);


useEffect(() => {
    if (selectedCompany && selectedClassification && fileCounter) {
        const year = new Date().getFullYear();
        const increment = String(fileCounter).padStart(4, "0");
        
        // Format: 0091/TR/PAU-EXT/2026
        // where PAU = CompanyNameShort, EXT = classificationshort
        const requestNo = `${increment}/TR/${selectedCompany.value}-${selectedClassification.value}/${year}`;
        
        console.log("Generated Request No:", requestNo);
        setNewRequestNo(requestNo);
    } else {
        setNewRequestNo("");
    }
}, [selectedCompany, selectedClassification, fileCounter]);
    // --- Copy File Function ---
    const fetchAllUsers = async () => {
        // const companymaster = await sp.web.lists.getByTitle("ESSAcompanymaster").items.select("*").getAll();
        // const companyclassification = await sp.web.lists.getByTitle("ESSAcompanyclassification").items.select("*").getAll();
        // console.log(companymaster , "companymaster")
        // console.log(companyclassification , "companyclassification")

//         const counterItem = await sp.web.lists
//   .getByTitle("DMSFileCounterList")
//   .items.getById(1)();

// const fileCounter = counterItem.FileCount + 1;

// await sp.web.lists
//   .getByTitle("DMSFileCounterList")
//   .items.getById(1)
//   .update({ FileCount: fileCounter });

// const year = new Date().getFullYear();
// const increment = String(fileCounter).padStart(4, "0");

// const requestNo = `${increment}/TR/PAU-EXT/${year}`;

// console.log("Generated Request No:", requestNo);
// setNewRequestNo(requestNo);
        try {
          
  const userProfile = await sp.profiles.myProperties();
  console.log(userProfile , "userProfile")
  console.log(userProfile.Title , "userProfile userProfile.Title")
  const userdata = await sp.web.currentUser();

  console.log(userdata , "user data edc")
  console.log(userdata.Id , "user data edc")
  currentUserIDref.current = userdata.Id;
  currentUserEmailRef.current = userdata.Email;
  currentUserTitleRef.current = userdata.Title;
       if(userdata){
        console.log(userdata.Email , "current user email")
          setSelectedFromUser({
      label: userdata.Title,
      value: userdata.Email
    });

    // default today date (YYYY-MM-DD for input type="date")
    const today = new Date().toISOString().split("T")[0];
    setSelectedDate(today);
       }
            const siteID = "931e8b14-6d73-412e-a411-720ab18bd124";
            const subsiteWeb = await sp.site.openWebById(siteID);
            
             // Get the root site collection
        const rootSite = await sp.site.rootWeb();
        
        // Get all users from the site collection
        const users = await sp.web.siteUsers()          
            // Format users for react-select
            const formattedUsers = users
                .filter((user: any) => !user.IsHiddenInUI && user.Email) // Filter out system accounts
                .map((user: any) => ({
                    value: user.Title, // Display name
                    label: user.Title,
                    email: user.Email,
                    id: user.Id
                }));
            
            setAllUsers(formattedUsers);
            console.log("Fetched users:", formattedUsers);
        } catch (error) {
            console.error("Error fetching users:", error);
            Swal.fire("Error", "Failed to load users", "error");
        }
    };
const copyFile = async () => {
    try {
        const sourceFileUrl = "/sites/multiverseintranetportal/Document Template/Transmittal.docx";
        const destinationFolderUrl = "/sites/multiverseintranetportal/Location/TRANSMITTAL";
        // const siteID = "9f942ad9-f2b6-4d4a-b99c-80f1171b57e3";
        const siteID = "931e8b14-6d73-412e-a411-720ab18bd124";

        Swal.fire({
            title: "Preparing Document...",
            text: "Please wait.",
            allowOutsideClick: false,
            didOpen: () => Swal.showLoading(),
        });

        const fileBuffer = await sp.web.getFileByServerRelativePath(sourceFileUrl).getBuffer();
        const fileBlob = new Blob([fileBuffer]);

        const subsiteWeb = await sp.site.openWebById(siteID);
        const uniqueFileName = `Transmittal_${Date.now()}.docx`;

        // Upload
        const response = await subsiteWeb.web
            .getFolderByServerRelativePath(destinationFolderUrl)
            .files.addUsingPath(uniqueFileName, fileBlob, { Overwrite: true });
        setselectedFile(response);
        setUploadedFileResponse(response);
        setFileServerRelativeUrl(response.data.ServerRelativeUrl);

        // Update initial metadata
        // const listItem = await response.file.getItem();
        // await listItem.update({ Status: "Pending" });

        // Generate a request number from counter list
        // const counterItem = await sp.web.lists.getByTitle("DMSFileCounterList").items.getById(1)();
        // const fileCounter = counterItem.FileCount + 1;
        // await sp.web.lists.getByTitle("DMSFileCounterList").items.getById(1).update({
        //     FileCount: fileCounter,
        // });
        // const requestNo = `File${String(fileCounter).padStart(2, "0")}`;
        // setNewRequestNo(requestNo);

        //previous working reqno 
//         const counterItem = await sp.web.lists
//   .getByTitle("DMSFileCounterList")
//   .items.getById(1)();

// const fileCounter = counterItem.FileCount + 1;

// await sp.web.lists
//   .getByTitle("DMSFileCounterList")
//   .items.getById(1)
//   .update({ FileCount: fileCounter });

// const now = new Date();
// const month = String(now.getMonth() + 1).padStart(2, "0"); // 01–12
// const year = now.getFullYear();
// const increment = String(fileCounter).padStart(4, "0");

// const requestNo = `TR_${month}_${year}_${increment}`;
// console.log("Generated Request No:", requestNo);
// setNewRequestNo(requestNo);


        // Build edit URL
        const editFileUrl = `https://multiverse.sharepoint.com/sites/multiverseintranetportal${encodeURIComponent(
            response.data.ServerRelativeUrl.split("/sites/multiverseintranetportal")[1]
        )}?web=1`;

        setEditUrl(editFileUrl);

        Swal.close();
    } catch (err: any) {
        Swal.close();
        console.error("Error in copyFile:", err);
        Swal.fire("Error", `An error occurred while copying the document. ${err.message}`, "error");
    }
};
// const copyFile = async () => {
//     try {
//         const sourceFileUrl = "/sites/Intranetdemos/Document Template/Transmittal.docx";
//         const destinationFolderUrl = "/sites/Intranetdemos/Location/Section";
//         const siteID = "9f942ad9-f2b6-4d4a-b99c-80f1171b57e3";
//         Swal.fire({
//             title: "Preparing Document...",
//             text: "Please wait.",
//             allowOutsideClick: false,
//             didOpen: () => Swal.showLoading(),
//         });
//         const fileBuffer = await sp.web.getFileByServerRelativePath(sourceFileUrl).getBuffer();
//         const fileBlob = new Blob([fileBuffer]);
//         const subsiteWeb = await sp.site.openWebById(siteID);
//         const uniqueFileName = `Transmittal_${Date.now()}.docx`;
//         // Upload
//         const response = await subsiteWeb.web
//             .getFolderByServerRelativePath(destinationFolderUrl)
//             .files.addUsingPath(uniqueFileName, fileBlob, { Overwrite: true });
//         setUploadedFileResponse(response);
//         setFileServerRelativeUrl(response.data.ServerRelativeUrl);
//         // Update initial metadata
//         const listItem = await response.file.getItem();
//         await listItem.update({ Status: "Pending" });
//         // Generate a request number from counter list
//         const counterItem = await sp.web.lists.getByTitle("DMSFileCounterList").items.getById(1)();
//         const fileCounter = counterItem.FileCount + 1;
//         await sp.web.lists.getByTitle("DMSFileCounterList").items.getById(1).update({
//             FileCount: fileCounter,
//         });
//         const requestNo = `File${String(fileCounter).padStart(2, "0")}`;
//         setNewRequestNo(requestNo);

//         // Update document content using SharePoint's Office Online integration
//         try {
//             const docNumber = "123456"; // Your document number - you can use fileCounter or requestNo here
            
//             // Option A: Use SharePoint's content replacement via custom properties
//             // Update the existing listItem with the document number
//             await listItem.update({ 
//                 DOCNumber: docNumber,
//                 // Add other properties that your Word template can reference
//             });
            
//             console.log("Document metadata updated with DOC Number:", docNumber);
            
//             // Option B: If you have JSZip available, uncomment and use this approach:
//             /*
//             // First, make sure to import JSZip: import JSZip from 'jszip';
//             const uploadedFileBuffer = await response.file.getBuffer();
//             const arrayBuffer = uploadedFileBuffer.slice();
            
//             const zip = new JSZip();
//             const docxZip = await zip.loadAsync(arrayBuffer);
//             const documentXml = await docxZip.file("word/document.xml").async("text");
//             const updatedXml = documentXml.replace(/DOC Number: \[Fill here\]/g, `DOC Number: ${docNumber}`);
            
//             docxZip.file("word/document.xml", updatedXml);
//             const updatedFileBuffer = await docxZip.generateAsync({type: "arraybuffer"});
//             const updatedBlob = new Blob([updatedFileBuffer], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
            
//             await subsiteWeb.web
//                 .getFolderByServerRelativePath(destinationFolderUrl)
//                 .files.addUsingPath(uniqueFileName, updatedBlob, { Overwrite: true });
//             */
            
//         } catch (updateError) {
//             console.error("Error updating document content:", updateError);
//             // Continue with the rest of the process even if document update fails
//         }

//         // Build edit URL
//         const editFileUrl = `https://officeindia.sharepoint.com/:w:/r/sites/Intranetdemos${encodeURIComponent(
//             response.data.ServerRelativeUrl.split("/sites/Intranetdemos")[1]
//         )}?web=1`;
//         setEditUrl(editFileUrl);
//         Swal.close();
//     } catch (err: any) {
//         Swal.close();
//         console.error("Error in copyFile:", err);
//         Swal.fire("Error", `An error occurred while copying the document. ${err.message}`, "error");
//     }
// };
// --- Submit Function ---



///previous working code 
// const handleSubmit = async () => {
//     try {
//         if (!uploadedFileResponse) {
//             Swal.fire("Error", "No file has been prepared. Please run Copy first.", "error");
//             return;
//         }

//         Swal.fire({
//             title: "Submitting Document",
//             text: "Please wait while we submit your document for approval.",
//             allowOutsideClick: false,
//             didOpen: () => Swal.showLoading(),
//         });

//         const siteID = "9f942ad9-f2b6-4d4a-b99c-80f1171b57e3";
//         const subsiteWeb = await sp.site.openWebById(siteID);
        
//         //u[date file status to pending
       
//         // Build preview URL
//         const encodeSharePointURL = (url: string) => encodeURIComponent(url);
//         const parentFolder = fileServerRelativeUrl.substring(0, fileServerRelativeUrl.lastIndexOf("/"));
//         const previewUrl = `https://officeindia.sharepoint.com/sites/Intranetdemos/Location/Section/Forms/AllItems.aspx?id=${encodeSharePointURL(
//             fileServerRelativeUrl
//         )}&parent=${encodeSharePointURL(parentFolder)}`;
//            const getcurrentuseremail = async () => {
//             const currentUser = await sp.web.currentUser();
//             return currentUser.Email;
//         }
//         const currentUserEmail = await getcurrentuseremail();
//         // Save in DMSLocationFileMaster
//         await sp.web.lists.getByTitle("DMSLocationFileMaster").items.add({
//             FileName: String(uploadedFileResponse.data.Name),
//             FileSize: String(uploadedFileResponse.data.Length),
//             FileVersion: String(uploadedFileResponse.data.MajorVersion),
//             CurrentFolderPath: "/sites/Intranetdemos/Location/Section",
//             FileUID: String(uploadedFileResponse.data.UniqueId),
//             CurrentUser: String(currentUserEmail),
//             SiteID: siteID,
//             Status: "Pending",
//             FilePreviewURL: previewUrl,
//             DocumentLibraryName: "Section",
//             SiteName: "Location",
//             MyRequest: true,
//             Processname: "New File Request",
//             RequestNo: newRequestNo,
//         });

//         // Save in approval list
//         await sp.web.lists.getByTitle("DMSFileApprovalList").items.add({
//             SiteName: "Location",
//             DocumentLibraryName: "Section",
//             RequestedBy: String(currentUserEmail),
//             FileName: String(uploadedFileResponse.data.Name),
//             FileUID: String(uploadedFileResponse.data.UniqueId),
//             FilePreviewUrl: previewUrl,
//             Status: "Pending",
//             FolderPath: "/sites/Intranetdemos/Location/Section",
//             ApproveAction: "Submitted",
//             ApprovedLevel: 1,
//             RequestNo: newRequestNo,
//             Processname: "New File Request",
//             CurrentLevel: 1,
//         });

         
//         // Release lock if required
//         try {
//             await releaseFileLock(subsiteWeb, fileServerRelativeUrl);
//         } catch (lockError) {
//             console.log("File lock release failed:", lockError);
//         }
//          console.log("File to update:", selectedfile);
//          const listItem = await selectedfile.file.getItem();
//          debugger
//          const updatestatus=   await listItem.update({ Status: "Pending" });
//          console.log("File status updated to Pending:", updatestatus);
//         Swal.close();
//         Swal.fire({
//             icon: "success",
//             title: "Document submitted",
//             text: `Your Documnet Name is ${uploadedFileResponse.data.Name} for your future reference`,
//             confirmButtonText: "OK",
//         }).then((result) => {
//             if (result.isConfirmed) {
//                 window.location.reload();
//             }
//         });

//         setTimeout(() => {
//             setEditUrl(null);
//         }, 2000);
//     } catch (error: any) {
//         Swal.close();
//         console.error("Error submitting document:", error);

//         if (error.message.includes("locked for shared use") || error.message.includes("423")) {
//             Swal.fire({
//                 title: "File Still in Use",
//                 html: `Please close the file in Word Online and try again.`,
//                 icon: "warning",
//                 confirmButtonText: "Try Again",
//                 showCancelButton: true,
//             }).then((result) => {
//                 if (result.isConfirmed) {
//                     handleSubmit();
//                 }
//             });
//         } else {
//             Swal.fire("Error", `An error occurred while submitting the document: ${error.message}`, "error");
//         }
//     }
// };
    const releaseFileLock = async (subsiteWeb: any, fileUrl: string) => {
        try {
            console.log("Attempting to release file lock for:", fileUrl);
            
            // Use the _api web endpoint to release the lock
            const result = await subsiteWeb.web.getFileByServerRelativePath(fileUrl).select("CheckInComment")();
            
            // If we get here, the file is accessible, try to check it in
            try {
                await subsiteWeb.web.getFileByServerRelativePath(fileUrl).checkin("Document submitted for approval", 1);
                console.log("File checked in successfully");
                return true;
            } catch (checkinError) {
                console.log("Standard checkin failed, trying undo checkout:", checkinError);
                
                // Try to undo checkout first
                try {
                    await subsiteWeb.web.getFileByServerRelativePath(fileUrl).undoCheckout();
                    console.log("Undo checkout successful");
                    return true;
                } catch (undoError) {
                    console.log("Undo checkout also failed:", undoError);
                    throw undoError;
                }
            }
        } catch (error) {
            console.error("Error in releaseFileLock:", error);
            
            // If all else fails, try a direct API call using fetch
            try {
                console.log("Trying direct API call to release lock");
                const siteUrl = subsiteWeb.web.toUrl();
                const endpoint = `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${fileUrl}')/UndoCheckout`;
                
                const response = await fetch(endpoint, {
                    method: 'POST',
                    headers: {
                        'Accept': 'application/json;odata=verbose',
                        'Content-Type': 'application/json;odata=verbose',
                        'X-RequestDigest': (document.querySelector('#__REQUESTDIGEST') as HTMLInputElement)?.value || ''
                    },
                    credentials: 'include'
                });
                
                if (response.ok) {
                    console.log("Direct API call to undo checkout succeeded");
                    return true;
                } else {
                    console.error("Direct API call failed:", response.status, response.statusText);
                    throw new Error(`API call failed: ${response.status}`);
                }
            } catch (apiError) {
                console.error("Direct API call also failed:", apiError);
                throw apiError;
            }
        }
    };

///previous working code 
// Helper function to check if file is locked
const isFileLocked = async (file: any): Promise<boolean> => {
    try {
        await file.getItem();
        return false; // If we can get the item, it's not locked
    } catch (error: any) {
        if (error.message.includes("locked") || error.message.includes("423")) {
            return true;
        }
        throw error; // Re-throw other errors
    }
};

// Helper function to wait for file unlock
const waitForFileUnlock = async (file: any, maxWaitTime: number = 30000): Promise<boolean> => {
    const startTime = Date.now();
    const checkInterval = 2000;
    
    while (Date.now() - startTime < maxWaitTime) {
        const locked = await isFileLocked(file);
        if (!locked) {
            return true;
        }
        await new Promise(resolve => setTimeout(resolve, checkInterval));
    }
    return false;
};

// const handleSubmit = async () => {
//     try {
//         if (!uploadedFileResponse) {
//             Swal.fire("Error", "No file has been prepared. Please run Copy first.", "error");
//             return;
//         }

//         Swal.fire({
//             title: "Submitting Document",
//             text: "Please wait while we submit your document for approval.",
//             allowOutsideClick: false,
//             didOpen: () => Swal.showLoading(),
//         });

//         // const siteID = "9f942ad9-f2b6-4d4a-b99c-80f1171b57e3";
//           const siteID = "931e8b14-6d73-412e-a411-720ab18bd124";
//         const subsiteWeb = await sp.site.openWebById(siteID);
        
//         // Build preview URL
//         const encodeSharePointURL = (url: string) => encodeURIComponent(url);
//         const parentFolder = fileServerRelativeUrl.substring(0, fileServerRelativeUrl.lastIndexOf("/"));
//         const previewUrl = `https://multiverse.sharepoint.com/sites/multiverseintranetportal/Location/TRANSMITTAL/Forms/AllItems.aspx?id=${encodeSharePointURL(
//             fileServerRelativeUrl
//         )}&parent=${encodeSharePointURL(parentFolder)}`;
           
//         const getcurrentuseremail = async () => {
//             const currentUser = await sp.web.currentUser();
//             return currentUser.Email;
//         }
//         const currentUserEmail = await getcurrentuseremail();

//         // Save in DMSLocationFileMaster
//         await sp.web.lists.getByTitle("DMSLocationFileMaster").items.add({
//             FileName: String(uploadedFileResponse.data.Name),
//             FileSize: String(uploadedFileResponse.data.Length),
//             FileVersion: String(uploadedFileResponse.data.MajorVersion),
//             CurrentFolderPath: "/sites/multiverseintranetportal/Location/TRANSMITTAL",
//             FileUID: String(uploadedFileResponse.data.UniqueId),
//             CurrentUser: String(currentUserEmail),
//             SiteID: siteID,
//             Status: "Pending",
//             FilePreviewURL: previewUrl,
//             DocumentLibraryName: "TRANSMITTAL",
//             SiteName: "Location",
//             MyRequest: true,
//             Processname: "New File Request",
//             RequestNo: newRequestNo,
//         });

//         // Save in approval list
//         await sp.web.lists.getByTitle("DMSFileApprovalList").items.add({
//             SiteName: "Location",
//             DocumentLibraryName: "TRANSMITTAL",
//             RequestedBy: String(currentUserEmail),
//             FileName: String(uploadedFileResponse.data.Name),
//             FileUID: String(uploadedFileResponse.data.UniqueId),
//             FilePreviewUrl: previewUrl,
//             Status: "Pending",
//             FolderPath: "/sites/multiverseintranetportal/Location/TRANSMITTAL",
//             ApproveAction: "Submitted",
//             ApprovedLevel: 1,
//             RequestNo: newRequestNo,
//             Processname: "New File Request",
//             CurrentLevel: 1,
//         });
//          //  wordEditorFrame
//           const iframe = document.getElementById("wordEditorFrame") as HTMLIFrameElement;
//           if (iframe) {
//               iframe.remove();

//           }
//         // Release lock first
//       try {
//     await releaseFileLock(subsiteWeb, fileServerRelativeUrl);
    
//     // Wait for file to be unlocked
//     const isUnlocked = await waitForFileUnlock(selectedfile.file, 30000);
    
//     if (isUnlocked) {
//         await new Promise(resolve => setTimeout(resolve, 8000)); // 3 second delay
//         const listItem = await selectedfile.file.getItem();
//         const updatestatus = await listItem.update({ Status: "Pending" });
//         console.log("File status updated to Pending:", updatestatus);
//     } else {
//         console.warn("File remained locked after waiting period, status update skipped");
//     }
    
// } catch (lockError) {
//     console.log("File lock handling failed:", lockError);
// }

//         Swal.close();
//         Swal.fire({
//             icon: "success",
//             title: "Document submitted",
//             text: `Your Document Name is ${uploadedFileResponse.data.Name} for your future reference`,
//             confirmButtonText: "OK",
//         }).then((result) => {
//             if (result.isConfirmed) {
//                 window.location.reload();
//             }
//         });

//         setTimeout(() => {
//             setEditUrl(null);
//         }, 2000);
//     } catch (error: any) {
//         Swal.close();
//         console.error("Error submitting document:", error);

//         if (error.message.includes("locked for shared use") || error.message.includes("423")) {
//             Swal.fire({
//                 title: "File Still in Use",
//                 html: `Please close the file in Word Online and try again.`,
//                 icon: "warning",
//                 confirmButtonText: "Try Again",
//                 showCancelButton: true,
//             }).then((result) => {
//                 if (result.isConfirmed) {
//                     handleSubmit();
//                 }
//             });
//         } else {
//             Swal.fire("Error", `An error occurred while submitting the document: ${error.message}`, "error");
//         }
//     }
// };


// working code previous updating from and to columns
// const handleSubmit = async () => {
//     try {
//         // Validations
//         if (!selectedFromUser) {
//             Swal.fire("Validation Error", "Please select 'From' user", "warning");
//             return;
//         }
        
//         if (!selectedToUser) {
//             Swal.fire("Validation Error", "Please select 'To' user", "warning");
//             return;
//         }
        
//         if (!uploadedFileResponse) {
//             Swal.fire("Error", "No file has been prepared. Please run Copy first.", "error");
//             return;
//         }

//         Swal.fire({
//             title: "Submitting Document",
//             text: "Please wait while we submit your document for approval.",
//             allowOutsideClick: false,
//             didOpen: () => Swal.showLoading(),
//         });

//         const siteID = "931e8b14-6d73-412e-a411-720ab18bd124";
//         const subsiteWeb = await sp.site.openWebById(siteID);
        
//         // 🔥 STEP 1: Remove iframe FIRST to close the editing session
//         const iframe = document.getElementById("wordEditorFrame") as HTMLIFrameElement;
//         if (iframe) {
//             iframe.remove();
//         }
        
//         // 🔥 STEP 2: Wait a moment for Word Online to release the file
//         await new Promise(resolve => setTimeout(resolve, 3000));
        
//         // 🔥 STEP 3: Try to force release the lock
//         try {
//             // Get the file object
//             const file = subsiteWeb.web.getFileByServerRelativePath(fileServerRelativeUrl);
            
//             // Check if file is checked out
//             const fileInfo = await file.select("CheckOutType", "CheckedOutByUser")();
//             console.log("File checkout status:", fileInfo);
            
//             if (fileInfo.CheckOutType !== 2) { // 2 = None (not checked out)
//                 try {
//                     // Try to discard checkout
//                     await file.undoCheckout();
//                     console.log("Successfully undid checkout");
//                     await new Promise(resolve => setTimeout(resolve, 2000));
//                 } catch (undoError) {
//                     console.log("Undo checkout failed, file may not be checked out:", undoError);
//                 }
//             }
//         } catch (lockError) {
//             console.log("Lock check/release failed:", lockError);
//         }
        
//         // 🔥 STEP 4: Wait for file to be completely unlocked
//         const isUnlocked = await waitForFileUnlock(selectedfile.file, 45000); // Increased timeout
        
//         if (!isUnlocked) {
//             Swal.close();
//             Swal.fire({
//                 title: "File Still Locked",
//                 html: "The file is still being edited in Word Online. Please:<br/>1. Close the Word editor<br/>2. Wait a few seconds<br/>3. Try submitting again",
//                 icon: "warning",
//                 confirmButtonText: "Try Again",
//                 showCancelButton: true,
//                 cancelButtonText: "Cancel"
//             }).then((result) => {
//                 if (result.isConfirmed) {
//                     handleSubmit();
//                 }
//             });
//             return;
//         }
        
//         // 🔥 STEP 5: Now update the list item columns (From, To, Status)
//         await new Promise(resolve => setTimeout(resolve, 2000)); // Additional safety delay
        
//         const listItem = await selectedfile.file.getItem();
//         await listItem.update({ 
//             Status: "Pending",
//             From: selectedFromUser.value,  
//             To: selectedToUser.value       
//         });
        
//         console.log("File status and From/To fields updated successfully");
        
//         // Build preview URL
//         const encodeSharePointURL = (url: string) => encodeURIComponent(url);
//         const parentFolder = fileServerRelativeUrl.substring(0, fileServerRelativeUrl.lastIndexOf("/"));
//         const previewUrl = `https://multiverse.sharepoint.com/sites/multiverseintranetportal/Location/TRANSMITTAL/Forms/AllItems.aspx?id=${encodeSharePointURL(
//             fileServerRelativeUrl
//         )}&parent=${encodeSharePointURL(parentFolder)}`;
           
//         const currentUserEmail = (await sp.web.currentUser()).Email;

//         // Save in DMSLocationFileMaster
//         await sp.web.lists.getByTitle("DMSLocationFileMaster").items.add({
//             FileName: String(uploadedFileResponse.data.Name),
//             FileSize: String(uploadedFileResponse.data.Length),
//             FileVersion: String(uploadedFileResponse.data.MajorVersion),
//             CurrentFolderPath: "/sites/multiverseintranetportal/Location/TRANSMITTAL",
//             FileUID: String(uploadedFileResponse.data.UniqueId),
//             CurrentUser: String(currentUserEmail),
//             SiteID: siteID,
//             Status: "Pending",
//             FilePreviewURL: previewUrl,
//             DocumentLibraryName: "TRANSMITTAL",
//             SiteName: "Location",
//             MyRequest: true,
//             Processname: "New File Request",
//             RequestNo: newRequestNo,
//         });

//         // Save in approval list
//         await sp.web.lists.getByTitle("DMSFileApprovalList").items.add({
//             SiteName: "Location",
//             DocumentLibraryName: "TRANSMITTAL",
//             RequestedBy: String(currentUserEmail),
//             FileName: String(uploadedFileResponse.data.Name),
//             FileUID: String(uploadedFileResponse.data.UniqueId),
//             FilePreviewUrl: previewUrl,
//             Status: "Pending",
//             FolderPath: "/sites/multiverseintranetportal/Location/TRANSMITTAL",
//             ApproveAction: "Submitted",
//             ApprovedLevel: 1,
//             RequestNo: newRequestNo,
//             Processname: "New File Request",
//             CurrentLevel: 1,
//         });

//         Swal.close();
//         Swal.fire({
//             icon: "success",
//             title: "Document submitted",
//             text: `Your Document Name is ${uploadedFileResponse.data.Name} for your future reference`,
//             confirmButtonText: "OK",
//         }).then((result) => {
//             if (result.isConfirmed) {
//                 window.location.reload();
//             }
//         });

//         setTimeout(() => {
//             setEditUrl(null);
//         }, 2000);
        
//     } catch (error: any) {
//         Swal.close();
//         console.error("Error submitting document:", error);

//         if (error.message.includes("locked") || error.message.includes("423") || error.message.includes("checked out")) {
//             Swal.fire({
//                 title: "File Still in Use",
//                 html: `The file is still being edited. Please:<br/>
//                        1. Make sure Word editor is closed<br/>
//                        2. Wait a few seconds<br/>
//                        3. Try again`,
//                 icon: "warning",
//                 confirmButtonText: "Try Again",
//                 showCancelButton: true,
//             }).then((result) => {
//                 if (result.isConfirmed) {
//                     handleSubmit();
//                 }
//             });
//         } else {
//             Swal.fire("Error", `An error occurred: ${error.message}`, "error");
//         }
//     }
// };
// const handleSubmit = async () => {
//     try {
//         // Validations
//         if (!selectedFromUser) {
//             Swal.fire("Validation Error", "Please select 'From' user", "warning");
//             return;
//         }
        
//         if (!selectedToUser) {
//             Swal.fire("Validation Error", "Please select 'To' user", "warning");
//             return;
//         }
        
//         if (!uploadedFileResponse) {
//             Swal.fire("Error", "No file has been prepared. Please run Copy first.", "error");
//             return;
//         }

//         Swal.fire({
//             title: "Updating Document",
//             text: "Please wait while we update the document fields...",
//             allowOutsideClick: false,
//             didOpen: () => Swal.showLoading(),
//         });

//         // 🔥 STEP 1: Update document content BEFORE closing iframe
//         try {
//             const fieldMappings = {
//                 "Fill Here": newRequestNo || "N/A",  // Replace "No: Fill Here" with request number
//                 // Add more mappings as needed
//                 // "To \\(Kepada\\): Fill Here": selectedToUser.label,
//                 // "From : Fill Here": selectedFromUser.label,
//                 // "Date \\(Tanggal\\): Fill Here": new Date().toLocaleDateString(),
//             };
//             // Add this function before handleSubmit
// const updateDocumentContent = async (
//     iframeId: string, 
//     fieldMappings: { [key: string]: string }
// ) => {
//     return new Promise<void>((resolve, reject) => {
//         try {
//             const iframe = document.getElementById(iframeId) as HTMLIFrameElement;
            
//             if (!iframe || !iframe.contentWindow) {
//                 reject(new Error("Iframe not found or not accessible"));
//                 return;
//             }

//             // Wait for Office.js to be ready in the iframe
//             const checkOfficeReady = setInterval(() => {
//                 try {
//                     const iframeWindow = iframe.contentWindow as any;
                    
//                     if (iframeWindow.Office && iframeWindow.Office.context) {
//                         clearInterval(checkOfficeReady);
                        
//                         // Use Office.js API to update content
//                         iframeWindow.Office.context.document.body.getAsync(
//                             iframeWindow.Office.CoercionType.Text,
//                             (result: any) => {
//                                 if (result.status === iframeWindow.Office.AsyncResultStatus.Succeeded) {
//                                     let content = result.value;
                                    
//                                     // Replace placeholders with actual values
//                                     Object.keys(fieldMappings).forEach(placeholder => {
//                                         const value = fieldMappings[placeholder];
//                                         content = content.replace(
//                                             new RegExp(placeholder, 'g'), 
//                                             value
//                                         );
//                                     });
                                    
//                                     // Set updated content back
//                                     iframeWindow.Office.context.document.setSelectedDataAsync(
//                                         content,
//                                         { coercionType: iframeWindow.Office.CoercionType.Text },
//                                         (setResult: any) => {
//                                             if (setResult.status === iframeWindow.Office.AsyncResultStatus.Succeeded) {
//                                                 console.log("Document content updated successfully");
//                                                 resolve();
//                                             } else {
//                                                 reject(new Error("Failed to update document content"));
//                                             }
//                                         }
//                                     );
//                                 } else {
//                                     reject(new Error("Failed to read document content"));
//                                 }
//                             }
//                         );
//                     }
//                 } catch (error) {
//                     clearInterval(checkOfficeReady);
//                     reject(error);
//                 }
//             }, 500);
            
//             // Timeout after 10 seconds
//             setTimeout(() => {
//                 clearInterval(checkOfficeReady);
//                 reject(new Error("Timeout waiting for Office.js"));
//             }, 10000);
            
//         } catch (error) {
//             reject(error);
//         }
//     });
// };
//             await updateDocumentContent("wordEditorFrame", fieldMappings);
//             console.log("Document content updated");
            
//             // Wait for changes to be saved by Word Online
//             await new Promise(resolve => setTimeout(resolve, 3000));
            
//         } catch (updateError) {
//             console.warn("Could not auto-update document content:", updateError);
//             // Continue anyway - user may have filled manually
//         }

//         Swal.fire({
//             title: "Submitting Document",
//             text: "Please wait while we submit your document for approval.",
//             allowOutsideClick: false,
//             didOpen: () => Swal.showLoading(),
//         });

//         const siteID = "931e8b14-6d73-412e-a411-720ab18bd124";
//         const subsiteWeb = await sp.site.openWebById(siteID);
        
//         // 🔥 STEP 2: Remove iframe to close the editing session
//         const iframe = document.getElementById("wordEditorFrame") as HTMLIFrameElement;
//         if (iframe) {
//             iframe.remove();
//         }
        
//         // 🔥 STEP 3: Wait for Word Online to release the file
//         await new Promise(resolve => setTimeout(resolve, 3000));
        
//         // 🔥 STEP 4: Try to force release the lock
//         try {
//             const file = subsiteWeb.web.getFileByServerRelativePath(fileServerRelativeUrl);
//             const fileInfo = await file.select("CheckOutType", "CheckedOutByUser")();
//             console.log("File checkout status:", fileInfo);
            
//             if (fileInfo.CheckOutType !== 2) {
//                 try {
//                     await file.undoCheckout();
//                     console.log("Successfully undid checkout");
//                     await new Promise(resolve => setTimeout(resolve, 2000));
//                 } catch (undoError) {
//                     console.log("Undo checkout failed:", undoError);
//                 }
//             }
//         } catch (lockError) {
//             console.log("Lock check/release failed:", lockError);
//         }
        
//         // 🔥 STEP 5: Wait for file to be completely unlocked
//         const isUnlocked = await waitForFileUnlock(selectedfile.file, 45000);
        
//         if (!isUnlocked) {
//             Swal.close();
//             Swal.fire({
//                 title: "File Still Locked",
//                 html: "The file is still being edited. Please close Word editor and try again.",
//                 icon: "warning",
//                 confirmButtonText: "Try Again",
//                 showCancelButton: true,
//             }).then((result) => {
//                 if (result.isConfirmed) {
//                     handleSubmit();
//                 }
//             });
//             return;
//         }
        
//         // 🔥 STEP 6: Update list item columns
//         await new Promise(resolve => setTimeout(resolve, 2000));
        
//         const listItem = await selectedfile.file.getItem();
//         await listItem.update({ 
//             Status: "Pending",
//             From: selectedFromUser.value,  
//             To: selectedToUser.value,
//             DocumentNumber: newRequestNo,
//             MetaDataUpdated: "No",
//             Date: selectedDate.toString(),
//             Ref: reference       
//         });
//             const listItemData = await listItem.select("ID")();
//         const itemID = listItemData.ID;
//         console.log("List Item ID:", itemID);
//         console.log("File metadata updated successfully");
        
//         // Rest of your code (preview URL, DMSLocationFileMaster, approval list, etc.)
//         const encodeSharePointURL = (url: string) => encodeURIComponent(url);
//         const parentFolder = fileServerRelativeUrl.substring(0, fileServerRelativeUrl.lastIndexOf("/"));
//         const previewUrl = `https://multiverse.sharepoint.com/sites/multiverseintranetportal/Location/TRANSMITTAL/Forms/AllItems.aspx?id=${encodeSharePointURL(
//             fileServerRelativeUrl
//         )}&parent=${encodeSharePointURL(parentFolder)}`;
           
//         const currentUserEmail = (await sp.web.currentUser()).Email;

//         await sp.web.lists.getByTitle("DMSLocationFileMaster").items.add({
//             FileName: String(uploadedFileResponse.data.Name),
//             FileSize: String(uploadedFileResponse.data.Length),
//             FileVersion: String(uploadedFileResponse.data.MajorVersion),
//             CurrentFolderPath: "/sites/multiverseintranetportal/Location/TRANSMITTAL",
//             FileUID: String(uploadedFileResponse.data.UniqueId),
//             CurrentUser: String(currentUserEmail),
//             SiteID: siteID,
//             Status: "Pending",
//             FilePreviewURL: previewUrl,
//             DocumentLibraryName: "TRANSMITTAL",
//             SiteName: "Location",
//             MyRequest: true,
//             Processname: "New File Request",
//             RequestNo: newRequestNo,
//         });

//         await sp.web.lists.getByTitle("DMSFileApprovalList").items.add({
//             SiteName: "Location",
//             DocumentLibraryName: "TRANSMITTAL",
//             RequestedBy: String(currentUserEmail),
//             FileName: String(uploadedFileResponse.data.Name),
//             FileUID: String(uploadedFileResponse.data.UniqueId),
//             FilePreviewUrl: previewUrl,
//             Status: "Pending",
//             FolderPath: "/sites/multiverseintranetportal/Location/TRANSMITTAL",
//             ApproveAction: "Submitted",
//             ApprovedLevel: 1,
//             RequestNo: newRequestNo,
//             Processname: "New File Request",
//             CurrentLevel: 1,
//         });

//         Swal.close();
//         Swal.fire({
//             icon: "success",
//             title: "Document submitted",
//             text: `Your Document Name is ${uploadedFileResponse.data.Name}`,
//             confirmButtonText: "OK",
//         }).then((result) => {
//             if (result.isConfirmed) {
//                 window.location.reload();
//             }
//         });

//         setTimeout(() => {
//             setEditUrl(null);
//         }, 2000);
//          try {
//             const apiResponse = await fetch(
//                 "http://23.100.43.23:8081/MultiverseIntranetPortalAPI/api/UpdateMetaDataInTemplateFile",
//                 {
//                     method: "POST",
//                     headers: {
//                         "Content-Type": "application/json",
//                     },
//                     body: JSON.stringify({
//                         SiteUrl: "https://multiverse.sharepoint.com/sites/multiverseintranetportal/Location",
//                         ClientID: "6a92b992-0f2a-4e00-a8dd-1927138068b6",
//                         ClientSecret: "SlFrOFF+TktkdXNhZ3JGRWY1Y2NNQW9tZDRRVVlaeEVhQ0h+Q2FzRQ==",
//                         ItemID: itemID.toString()
//                     })
//                 }
//             );

//             if (!apiResponse.ok) {
//                 throw new Error(`API returned status ${apiResponse.status}`);
//             }

//             const apiResult = await apiResponse.json();
//             console.log("API Response:", apiResult);
            
//         } catch (apiError: any) {
//             console.error("Error calling UpdateMetaDataInTemplateFile API:", apiError);
//             // Continue with success message even if API fails
//             Swal.fire({
//                 icon: "warning",
//                 title: "Partial Success",
//                 text: "Document submitted but metadata update failed. Please contact administrator.",
//                 confirmButtonText: "OK",
//             });
//         }

//     } catch (error: any) {
//         Swal.close();
//         console.error("Error submitting document:", error);
//         Swal.fire("Error", `An error occurred: ${error.message}`, "error");
//     }
// };
const handleSubmit = async () => {
    try {
      // Validations
      if (!selectedFromUser) {
        Swal.fire("Validation Error", "Please select 'From' user", "warning");
        return;
      }
      if (!selectedToUser) {
        Swal.fire("Validation Error", "Please select 'To' user", "warning");
        return;
      }
      if (!uploadedFileResponse) {
        Swal.fire("Error", "No file has been prepared. Please run Copy first.", "error");
        return;
      }
  
      // 🔥 Show single loading alert that we'll update
      Swal.fire({
        title: "Processing Document",
        text: "Updating document fields...",
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading(),
      });
  
      // 🔥 STEP 1: Update document content BEFORE closing iframe
      try {
        const fieldMappings = {
          "Fill Here": newRequestNo || "N/A",
        };
  
        const updateDocumentContent = async (
          iframeId: string,
          fieldMappings: { [key: string]: string }
        ) => {
          return new Promise((resolve, reject) => {
            try {
              const iframe = document.getElementById(iframeId) as HTMLIFrameElement;
              if (!iframe || !iframe.contentWindow) {
                reject(new Error("Iframe not found or not accessible"));
                return;
              }
  
              const checkOfficeReady = setInterval(() => {
                try {
                  const iframeWindow = iframe.contentWindow as any;
                  if (iframeWindow.Office && iframeWindow.Office.context) {
                    clearInterval(checkOfficeReady);
  
                    iframeWindow.Office.context.document.body.getAsync(
                      iframeWindow.Office.CoercionType.Text,
                      (result: any) => {
                        if (result.status === iframeWindow.Office.AsyncResultStatus.Succeeded) {
                          let content = result.value;
  
                          Object.keys(fieldMappings).forEach(placeholder => {
                            const value = fieldMappings[placeholder];
                            content = content.replace(
                              new RegExp(placeholder, 'g'),
                              value
                            );
                          });
  
                          iframeWindow.Office.context.document.setSelectedDataAsync(
                            content,
                            { coercionType: iframeWindow.Office.CoercionType.Text },
                            (setResult: any) => {
                              if (setResult.status === iframeWindow.Office.AsyncResultStatus.Succeeded) {
                                console.log("Document content updated successfully");
                                resolve("Document content updated successfully");
                              } else {
                                reject(new Error("Failed to update document content"));
                              }
                            }
                          );
                        } else {
                          reject(new Error("Failed to read document content"));
                        }
                      }
                    );
                  }
                } catch (error) {
                  clearInterval(checkOfficeReady);
                  reject(error);
                }
              }, 500);
  
              setTimeout(() => {
                clearInterval(checkOfficeReady);
                reject(new Error("Timeout waiting for Office.js"));
              }, 10000);
            } catch (error) {
              reject(error);
            }
          });
        };
  
        await updateDocumentContent("wordEditorFrame", fieldMappings);
        console.log("Document content updated");
        await new Promise(resolve => setTimeout(resolve, 3000));
      } catch (updateError) {
        console.warn("Could not auto-update document content:", updateError);
      }
  
      // 🔥 Update alert text
      Swal.update({
        title: "Processing Document",
        text: "Submitting document for approval...",
        showConfirmButton: false
      });
  
      const siteID = "931e8b14-6d73-412e-a411-720ab18bd124";
      const subsiteWeb = await sp.site.openWebById(siteID);
  
      // 🔥 STEP 2: Remove iframe to close the editing session
      const iframe = document.getElementById("wordEditorFrame") as HTMLIFrameElement;
      if (iframe) {
        iframe.remove();
      }
  
      // 🔥 STEP 3: Wait for Word Online to release the file
      await new Promise(resolve => setTimeout(resolve, 3000));
  
      // 🔥 STEP 4: Try to force release the lock
      try {
        const file = subsiteWeb.web.getFileByServerRelativePath(fileServerRelativeUrl);
        const fileInfo = await file.select("CheckOutType", "CheckedOutByUser")();
        console.log("File checkout status:", fileInfo);
  
        if (fileInfo.CheckOutType !== 2) {
          try {
            await file.undoCheckout();
            console.log("Successfully undid checkout");
            await new Promise(resolve => setTimeout(resolve, 2000));
          } catch (undoError) {
            console.log("Undo checkout failed:", undoError);
          }
        }
      } catch (lockError) {
        console.log("Lock check/release failed:", lockError);
      }
  
      // 🔥 STEP 5: Wait for file to be completely unlocked
      const isUnlocked = await waitForFileUnlock(selectedfile.file, 45000);
      if (!isUnlocked) {
        Swal.close();
        Swal.fire({
          title: "File Still Locked",
          html: "The file is still being edited. Please close Word editor and try again.",
          icon: "warning",
          confirmButtonText: "Try Again",
          showCancelButton: true,
        }).then((result) => {
          if (result.isConfirmed) {
            handleSubmit();
          }
        });
        return;
      }
  
      // 🔥 STEP 6: Update list item columns
      await new Promise(resolve => setTimeout(resolve, 2000));
  
      const listItem = await selectedfile.file.getItem();
      await listItem.update({
        Status: "Pending",
        From: selectedFromUser.value,
        To: selectedToUser.value,
        TransmittalNumber: newRequestNo,
        MetaDataUpdated: "No",
        Date: selectedDate.toString(),
        Ref: reference
      });
  
      const listItemData = await listItem.select("ID")();
      const itemID = listItemData.ID;
      console.log("List Item ID:", itemID);
      console.log("File metadata updated successfully");
  
      const encodeSharePointURL = (url: string) => encodeURIComponent(url);
      const parentFolder = fileServerRelativeUrl.substring(0, fileServerRelativeUrl.lastIndexOf("/"));
      const previewUrl = `https://multiverse.sharepoint.com/sites/multiverseintranetportal/Location/TRANSMITTAL/Forms/AllItems.aspx?id=${encodeSharePointURL(
        fileServerRelativeUrl
      )}&parent=${encodeSharePointURL(parentFolder)}`;
  
      const currentUserEmail = (await sp.web.currentUser()).Email;
  
      await sp.web.lists.getByTitle("DMSLocationFileMaster").items.add({
        FileName: String(uploadedFileResponse.data.Name),
        FileSize: String(uploadedFileResponse.data.Length),
        FileVersion: String(uploadedFileResponse.data.MajorVersion),
        CurrentFolderPath: "/sites/multiverseintranetportal/Location/TRANSMITTAL",
        FileUID: String(uploadedFileResponse.data.UniqueId),
        CurrentUser: String(currentUserEmail),
        SiteID: siteID,
        Status: "Pending",
        FilePreviewURL: previewUrl,
        DocumentLibraryName: "TRANSMITTAL",
        SiteName: "Location",
        MyRequest: true,
        Processname: "New File Request",
        RequestNo: newRequestNo,
      });
  
      await sp.web.lists.getByTitle("DMSFileApprovalList").items.add({
        SiteName: "Location",
        DocumentLibraryName: "TRANSMITTAL",
        RequestedBy: String(currentUserEmail),
        FileName: String(uploadedFileResponse.data.Name),
        FileUID: String(uploadedFileResponse.data.UniqueId),
        FilePreviewUrl: previewUrl,
        Status: "Pending",
        FolderPath: "/sites/multiverseintranetportal/Location/TRANSMITTAL",
        ApproveAction: "Submitted",
        ApprovedLevel: 1,
        RequestNo: newRequestNo,
        Processname: "New File Request",
        CurrentLevel: 1,
      });
  
      // 🔥 Show success message
      Swal.close();
      Swal.fire({
        icon: "success",
        title: "Document Submitted",
        text: `Your Document Name is ${uploadedFileResponse.data.Name}`,
        confirmButtonText: "OK",
      }).then((result) => {
        if (result.isConfirmed) {
          window.location.reload();
        }
      });
  
      setTimeout(() => {
        setEditUrl(null);
      }, 2000);
  
    } catch (error: any) {
      Swal.close();
      console.error("Error submitting document:", error);
      Swal.fire("Error", `An error occurred: ${error.message}`, "error");
    }
  };

    const getApprovals = async () => {
        try {
           const hierarchy = await sp.web.lists
  .getByTitle("DMSFolderPermissionMaster")
  .items.select(
    "SiteName",
    "DocumentLibraryName",
    "ApprovalType",
    "Level",
    "ApprovalUser/Title",
    "ApprovalUser/EMail"
  )
  .expand("ApprovalUser")
  .filter(`SiteName eq 'Location' and DocumentLibraryName eq 'TRANSMITTAL'`)
  .orderBy("Level", true)
  .getAll();

// ✅ Use `hierarchy` (not `approvalHierachy`) for grouping
const groupedHierarchy = hierarchy.reduce((acc: any, item: any) => {
  if (!acc[item.Level]) {
    acc[item.Level] = {
      Level: item.Level,
      ApprovalType: item.ApprovalType,
      Emails: [],
    };
  }
  acc[item.Level].Emails.push(item.ApprovalUser.EMail);
  return acc;
}, {});

// ✅ Update state with grouped data
setApprovalHierachy(Object.values(groupedHierarchy));

console.log("Grouped Hierarchy:", Object.values(groupedHierarchy));

            console.log("Approval Hierarchy:", hierarchy);
        } catch (error) {
            console.error("Error fetching approvals:", error);
        }
    };

    useEffect(() => {
        // alert("this file")
        // Call copyFile when component mounts
        fetchAllUsers();
        copyFile();
        getApprovals();
    }, []);

    const getApprovalTypeText = (type: number) => {
        return type === 0 ? "Approved by One Only" : "Approved by All";
    };

    const updatedoc = async () => {
        await Word.run(async (context) => {
        const body = context.document.body;
         body.insertText("New Value beside table cell", Word.InsertLocation.end);
       await context.sync();
       }); 
    }

    return (
        <div  style={{ fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif" }}>
            {editUrl ? (
                <div>
                    <div className="Newfilealign">

                    <div style={{ marginBottom: "15px", display: "flex", justifyContent:"end", gap: "10px", flexWrap: "wrap" }}>
                      <h6 className="mb-1 fw-bold text-dark header-title"
                            style={{
                              color: "black",
                              marginBottom: "0px",
                              fontSize: "16px",
                            }}
                          >
                            Transmittal Template
                          </h6>
                        
                      {/* <button
                            style={{ 
                                padding: "10px 15px", 
                                border: "none", 
                                borderRadius: "4px", 
                                cursor: "pointer", 
                                fontWeight: "600", 
                                backgroundColor: "#6c757d", 
                                color: "white",
                                position: isFullScreen ? "fixed" : "static",
                                top: isFullScreen ? "10px" : "auto",
                                right: isFullScreen ? "10px" : "auto",
                                zIndex: 10000,
                                transition: "background-color 0.3s ease"
                            }}
                            onMouseOver={(e) => {
                                const target = e.target as HTMLButtonElement;
                                target.style.backgroundColor = "#5a6268";
                            }}
                            onMouseOut={(e) => {
                                const target = e.target as HTMLButtonElement;
                                target.style.backgroundColor = "#6c757d";
                            }}
                            onClick={() => setIsFullScreen(!isFullScreen)}
                        >
                            {isFullScreen ? "Exit Full Screen" : "Full Screen"}
                        </button> */}
                        <div
     // ⬅ prevents form submission / page reload
    style={{ 
       
        border: "none", 
       
        cursor: "pointer",  marginTop:"0px",
        fontWeight: "600", 
        
        position: isFullScreen ? "fixed" : "static",
        top: isFullScreen ? "6px" : "auto",
        right: isFullScreen ? "96px" : "auto",
        zIndex: 10000,  padding:'6px',
        transition: "background-color 0.3s ease"
    }}
    // onMouseOver={(e) => {
    //     const target = e.target as HTMLButtonElement;
    //     target.style.backgroundColor = "#5a6268";
    // }}
    // onMouseOut={(e) => {
    //     const target = e.target as HTMLButtonElement;
    //     target.style.backgroundColor = "#6c757d";
    // }}
    onClick={() => setIsFullScreen(!isFullScreen)}
>
{
  isFullScreen ? (
    <span className="mb-1 mt-2" data-tooltip="Exit Full Screen">
      <img style={{width:'13px'}} src={fullw1} alt="Exit Full Screen" />
    </span>
  ) : (
    <span className="mb-1 mt-2" data-tooltip="Full Screen">
      <img src={fullw} alt="Full Screen" />
    </span>
  )
}

    {/* <span className="mb-1 mt-2" data-tooltip="Full Screen">
                              <img  src={fullw}></img></span> */}
</div>

                    </div>
                   {/* Add this section just before the iframe div */}
<div className="card mb-3">
    <div className="card-body">
        <div className="row">
            <div className="col-md-6 mb-3">
                <label style={{ fontWeight: 600,     marginBottom: "8px", }}>
                Transmittal number
</label>

<input
  type="text"
  value={newRequestNo}
  readOnly
  style={{
    width: "100%",
    padding: "6px",
    border: "1px solid #ddd",
    backgroundColor: "#f3f2f1",
    cursor: "copy"
  }}
  onClick={(e:any) => e.target.select()}
/>
 {/* Company Master Dropdown */}
               

               

            </div>
            
            <div className="col-md-6 mb-3">
                <label htmlFor="toUser" style={{ 
                    fontWeight: "600", 
                    marginBottom: "8px",
                    display: "block",
                    color: "#323130"
                }}>
                    To <span style={{ color: "red" }}>*</span>
                </label>
                <Select
                    id="toUser"
                    options={allUsers}
                    value={selectedToUser}
                    onChange={(selected:any) => setSelectedToUser(selected)}
                    placeholder="Select To User"
                    isSearchable={true}
                    styles={{
                        control: (base:any) => ({
                            ...base,
                            minHeight: "38px",
                            borderColor: "#ddd",
                        })
                    }}
                />
            </div>

            <div className="col-md-6 mb-3">
            <label htmlFor="companyMaster" style={{ 
                    fontWeight: "600", 
                    marginBottom: "8px",
                    display: "block",
                    color: "#323130"
                }}>
                    Company <span style={{ color: "red" }}>*</span>
                </label>
                <Select
                    id="companyMaster"
                    options={companyMasterList}
                    value={selectedCompany}
                    onChange={(selected: any) => setSelectedCompany(selected)}
                    placeholder="Select Company"
                    isSearchable={true}
                    styles={{
                        control: (base: any) => ({
                            ...base,
                            minHeight: "38px",
                            borderColor: "#ddd",
                        })
                    }}
                />
                </div>

                <div className="col-md-6 mb-3">
                     {/* Classification Dropdown */}
                <label htmlFor="classification" style={{ 
                    fontWeight: "600", 
                    marginBottom: "8px",
                  
                    display: "block",
                    color: "#323130"
                }}>
                    Classification <span style={{ color: "red" }}>*</span>
                </label>
                <Select
                    id="classification"
                    options={classificationList}
                    value={selectedClassification}
                    onChange={(selected: any) => setSelectedClassification(selected)}
                    placeholder="Select Classification"
                    isSearchable={true}
                    styles={{
                        control: (base: any) => ({
                            ...base,
                            minHeight: "38px",
                            borderColor: "#ddd",
                        })
                    }}
                />

              

  

                    </div>
         
<div className="col-md-6 mb-3">
<label htmlFor="fromUser" style={{ 
                    fontWeight: "600", 
                    marginBottom: "8px",
                    display: "block",
                    color: "#323130"
                }}>
                    From <span style={{ color: "red" }}>*</span>
                </label>
                <Select
  id="fromUser"
  options={allUsers}
  value={selectedFromUser}   // ✅ full object
  isDisabled={true}          // ✅ correct prop
  placeholder="From User"
  styles={{
    control: (base:any) => ({
      ...base,
      minHeight: "38px",
      borderColor: "#ddd",
      backgroundColor: "#f3f2f1" // optional disabled look
    })
  }}
/>
</div>
<div className="col-md-6 mb-3">
<label style={{marginBottom: "8px",}}>Date <span style={{color:"red"}}>*</span></label>
<input
  type="date"
  className="form-control"
  value={selectedDate}
  onChange={(e) => setSelectedDate(e.target.value)}
/>
</div>

<div className="col-md-6 mb-3">
            
<label>Reference</label>
<input style={{height:'40px'}}
  type="text"
  className="form-control"
  value={reference}
  onChange={(e) => setReference(e.target.value)}
  placeholder="Enter reference"
/>
    </div>

        </div>
    </div>
</div>
                    
                    <div style={{ 
                        border: "1px solid #ddd", 
                        borderRadius: "4px", 
                        overflow: "hidden", 
                        boxShadow: "0 2px 10px rgba(0, 0, 0, 0.1)",
                        marginBottom: "20px"
                    }}>
                        <iframe
                            id="wordEditorFrame"
                            src={editUrl}
                            style={{
                                width: isFullScreen ? "100vw" : "100%",
                                height: isFullScreen ? "100vh" : "70vh",
                                border: "none",
                                position: isFullScreen ? "fixed" : "relative",
                                top: isFullScreen ? 0 : "auto",
                                left: isFullScreen ? 0 : "auto",
                                zIndex: isFullScreen ? 9999 : "auto",
                                background: "#fff",
                            }}
                            title="Editable Document"
                            allowFullScreen
                        ></iframe>
                    </div> </div>
                    
                    <div className="card mar-90">
                    <div className="card-body">
                        <h3 className="mb-1 fw-bold text-dark header-title"
                        >
                            Approval Hierarchy
                        </h3>
                        <table style={{ 
                            width: "100%", 
                            borderCollapse: "collapse",
                            marginTop: "10px"
                        }}>
                            <thead>
                                <tr>
                                    <th style={{ 
                                        backgroundColor: "#f6f9fc", 
                                        color: "6f6f6f", 
                                        textAlign: "left", 
                                        padding: "12px",
                                        fontSize: "0.9rem"
                                    }}>
                                        User
                                    </th>
                                    <th style={{ 
                                         backgroundColor: "#f6f9fc", 
                                         color: "6f6f6f",
                                        textAlign: "left", 
                                        padding: "12px",
                                        fontSize: "0.9rem"
                                    }}>
                                        Level
                                    </th>
                                    <th style={{ 
                                          backgroundColor: "#f6f9fc", 
                                          color: "6f6f6f",
                                        textAlign: "left", 
                                        padding: "12px",
                                        fontSize: "0.9rem"
                                    }}>
                                        Approval Type
                                    </th>
                                </tr>
                            </thead>
                            <tbody>
                                {/* {
                                approvalHierachy && approvalHierachy.length > 0 ? (
                                    approvalHierachy.map((item: any) => (
                                        <tr key={item.Id} style={{ borderBottom: "1px solid #ddd" }}>
                                            <td style={{ padding: "12px", fontSize: "0.9rem" }}>{item.ApprovalUser.EMail}</td>
                                            <td style={{ padding: "12px", fontSize: "0.9rem" }}>{item.Level}</td>
                                            <td style={{ padding: "12px", fontSize: "0.9rem" }}>{getApprovalTypeText(item.ApprovalType)}</td>
                                        </tr>
                                    ))
                                ) 
                                : (
                                    <tr>
                                        <td colSpan={3} style={{ textAlign: "center", padding: "12px", fontSize: "0.9rem" }}>
                                            No approval hierarchy found
                                        </td>
                                    </tr>
                                )} */}
                              {
  approvalHierachy && approvalHierachy.length > 0 ? (
    approvalHierachy.map((item: any) => (
      <tr key={item.Level} style={{ borderBottom: "1px solid #ddd" }}>
        <td style={{ padding: "12px", fontSize: "0.9rem" }}>
          {item.Emails.join(", ")}
        </td>
        <td style={{ padding: "12px", fontSize: "0.9rem" }}>
         <span className="circlelevel">{item.Level}</span> 
        </td>
        <td style={{ padding: "12px", fontSize: "0.9rem" }}>
        <span className="appbg">  {getApprovalTypeText(item.ApprovalType)} </span>
        </td>
      </tr>
    ))
  ) : (
    <tr>
      <td colSpan={3} style={{ textAlign: "center", padding: "12px", fontSize: "0.9rem" }}>
        No approval hierarchy found
      </td>
    </tr>
  )
}


                            </tbody>
                        </table>
                    </div>
                   
                </div>
                <div className="d-flex justify-content-end gap-3 mt-3 mar-90">
                    <button className="btncolorCreate1 mt-0"
                            type="button"
                            // onMouseOver={(e) => {
                            //     const target = e.target as HTMLButtonElement;
                            //     target.style.backgroundColor = "#106ebe";
                            // }}
                            // onMouseOut={(e) => {
                            //     const target = e.target as HTMLButtonElement;
                            //     target.style.backgroundColor = "#008751";
                            // }}
                            onClick={handleSubmit}
                        > <span className="mb-1 mt-2" data-tooltip="Submited Edited File ">
                              <img  src={submitnew}></img></span>
                           
                        </button>
                          <button className=" btncolorCreate1 mt-0 alitool" onClick={()=>window.location.reload()} 
                            > <span className="mb-1 mt-2" data-tooltip="Cancel">
                             <img  src={cancelnew}></img></span> </button>
          </div></div>
            ) :
            
             (
                <div style={{ textAlign: "center", padding: "40px" }}>
                    {/* <p style={{ marginBottom: "15px", color: "#323130", fontSize: "18px" }}>Loading document from template...</p>
                    <div style={{
                        width: "3rem",
                        height: "3rem",
                        border: "5px solid #f3f3f3",
                        borderTop: "5px solid #008751",
                        borderRadius: "50%",
                        animation: "spin 1s linear infinite",
                        margin: "0 auto"
                    }}></div>
                    <style>
                        {`
                        @keyframes spin {
                            0% { transform: rotate(0deg); }
                            100% { transform: rotate(360deg); }
                        }
                        `}
                    </style> */}
                </div>
            )}
        </div>
    );
};

export default DocumentTemplatetofill;