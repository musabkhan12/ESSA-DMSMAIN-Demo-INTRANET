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



import React, { useState, useEffect } from "react";
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
import "@pnp/sp/presets/all";
import "@pnp/sp/site-users/web";
import "./DocumnetTemplate";

const DocumentTemplatetofill = (props: any) => {
  const sp: SPFI = getSP();
  const [editUrl, setEditUrl] = useState<string | null>(null);
  const [isFullScreen, setIsFullScreen] = useState(false);
 const [approvalhierachy, setapprovalhierachy] = useState<any[]>([]);

  useEffect(() => {
    getapprovals();
  },[])
  const getapprovals = async () => {
      const approvalhierachy = await sp.web.lists
      .getByTitle("DMSFolderPermissionMaster").items.filter(`SiteName eq 'Location' and DocumentLibraryName eq 'Approval section'`).getAll();
      console.log("approvalhierachy",approvalhierachy);
      setapprovalhierachy(approvalhierachy);
  }

  useEffect(() => {
    if (props.fileinedit) {
      setEditUrl(props.fileinedit);
    }
  }, [props.fileinedit]);

  const handleSubmit = async () => {
    try {
      const sourceUrl =
        "/sites/Intranetdemos/Document Template/Transmittal.docx";
      const targetLibrary = "/sites/Intranetdemos/Document Template Destination";
      const newFileName = `Edited_${Date.now()}.docx`;

      const file = await sp.web
        .getFileByServerRelativePath(sourceUrl)
        .getBuffer();
      const fileBlob = new Blob([file], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      await sp.web
        .getFolderByServerRelativePath(targetLibrary)
        .files.addChunked(newFileName, fileBlob);

      Swal.fire("Success", "Edited document saved to FinalDocs!", "success");
    } catch (err: any) {
      Swal.fire("Error", err.message, "error");
    }
  };
 const getApprovalTypeText = (type: number) => {
    return type === 0 ? "Approved by One Only" : "Approved by All";
  };
  return (
    <div>
      {editUrl && (
        <div>
          {/* <div className="mt-3 flex gap-2">
            <button className="btn btn-success" onClick={handleSubmit}>
              Submit Edited File
            </button>

          <button
  className="btn btn-primary"
  onClick={() => setIsFullScreen(!isFullScreen)}
  style={{
    position: isFullScreen ? "fixed" : "static",
    top: isFullScreen ? "10px" : "auto",
    right: isFullScreen ? "10px" : "auto",
    zIndex: 10000, // ensure it's above iframe
  }}
>
  {isFullScreen ? "Exit Full Screen" : "Full Screen"}
</button>
          </div> */}
<div className="mt-0 mb-2 d-flex justify-content-end  gap-2">
  <button className="btn btn-success mt-0" onClick={handleSubmit}>
    Submit Edited File
  </button>

  <button
    className="btn btn-primary mt-0"
    onClick={() => setIsFullScreen(!isFullScreen)}
    style={{
      position: isFullScreen ? "fixed" : "static",
      top: isFullScreen ? "10px" : "auto",
      right: isFullScreen ? "10px" : "auto",
      zIndex: 10000,
    }}
  >
    {isFullScreen ? "Exit Full Screen" : "Full Screen"}
  </button>
</div>
          <iframe
            src={editUrl}
            style={{
              width: isFullScreen ? "100vw" : "100%",
              height: isFullScreen ? "100vh" : "600px",
              border: "none",
              position: isFullScreen ? "fixed" : "relative",
              top: isFullScreen ? 0 : "auto",
              left: isFullScreen ? 0 : "auto",
              zIndex: isFullScreen ? 9999 : "auto",
              background: "#fff",
            }}
            title="Editable Document"
          ></iframe>
          <div>
            <table className="mtbalenew">
        <thead>
          <tr>
            <th>User</th>
            <th style={{minWidth:"50px",maxWidth:"50px"}}>Level</th>
            <th>Approval Type</th>
          </tr>
        </thead>
       <tbody>
  {approvalhierachy && approvalhierachy.length > 0 ? (
    approvalhierachy.map((item: any) => (
      <tr key={item.Id}>
        <td>{item.CurrentUser}</td>
        <td style={{minWidth:"50px",maxWidth:"50px",textAlign:"center"}}>{item.Level}</td>
        <td>{getApprovalTypeText(item.ApprovalType)}</td>
      </tr>
    ))
  ) : (
    <tr>
      <td colSpan={3} style={{ textAlign: "center" }}>
        No approval hierarchy found
      </td>
    </tr>
  )}
</tbody>
      </table>
          </div>
        </div>
      )}
    </div>
  );
};


export default DocumentTemplatetofill;
