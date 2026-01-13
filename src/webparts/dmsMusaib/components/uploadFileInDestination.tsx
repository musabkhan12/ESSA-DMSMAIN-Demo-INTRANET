import * as React from "react";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp/presets/all";
import Swal from "sweetalert2";
import { Modal } from 'react-bootstrap';
let folderdes = require('../assets/folderdes.png');
import '../components/uploadfilecss.css';
// srs new
import * as XLSX from 'xlsx';
const COLUMNS = [
    "S.No",
    "Location",
    "Department",
    "DocumentType",
    "Discipline",
    "Template",
    "FileName",
    "External Party",
    "From",
    "Issued Date",
    "Year",
    "Subject",
    "Project",
    "Tag No.",
    "Area"
  ];
  
interface FolderNode {
    name: string;
    children?: FolderNode[];
}
const uploadFileInDestination = () => {
    const sp: SPFI = getSP();
    const [selectedFile, setSelectedFile] = React.useState<File | null>(null);
    const [sourceItems, setSourceItems] = React.useState<any[]>([]);
    const [destItems, setDestItems] = React.useState<any[]>([]);
    const [loading, setLoading] = React.useState<boolean>(false);

    const [showModal, setShowModal] = React.useState(false);



    const [selectedSource, setSelectedSource] = React.useState<string>("");
    const [selectedDest, setSelectedDest] = React.useState<string>("");


    const [breadcrumbs, setBreadcrumbs] = React.useState<string[]>([]);
// srs 8/1/26
    const [previewUrl, setPreviewUrl] = React.useState<string>("");
const [isPreviewLoading, setIsPreviewLoading] = React.useState<boolean>(false);

// srs new 
const [excelData, setExcelData] = React.useState<any[]>([]);
// Add this with your other useState hooks
const [localFiles, setLocalFiles] = React.useState<File[]>([]);
const [uploadProgress, setUploadProgress] = React.useState({ current: 0, total: 0 });

    const folderData: FolderNode[] = [
        {
            name: "EHO",
            children: [
                 {
                    name: "HRD",
                    children: [
                        { name: "Document Type Docs" },
                        // { name: "DOCUMENT REPORTS" }
                    ]
                },
                 {
                    name: "ICT",
                    children: [
                        // { name: "DOCUMENT DRAWINGS" },
                        // { name: "DOCUMENT REPORTS" }
                    ]
                },
                 {
                    name: "ADM",
                    children: [
                        // { name: "DOCUMENT DRAWINGS" },
                        // { name: "DOCUMENT REPORTS" }
                    ]
                },
                 {
                    name: "CIVIL",
                    children: [
                        // { name: "DOCUMENT DRAWINGS" },
                        // { name: "DOCUMENT REPORTS" }
                    ]
                },
                {
                    name: "Document Type",
                    children: [
                        { name: "DOCUMENT DRAWINGS" },
                        { name: "DOCUMENT REPORTS" }
                    ]
                },
               
            ]
        },
        {
            name: "BAP",
            children: [
                // {
                //     name: "Document Type",
                //     children: [
                //         // { name: "DOCUMENT DRAWINGS" },
                //         // { name: "DOCUMENT REPORTS" }
                //     ]
                // },
                // {
                //     name: "Policies",
                //     children: [
                //         // { name: "HR" },
                //         // { name: "Finance" }
                //     ]
                // }
            ]
        },
        {
            name: "AKR",
            children: [
                {
                    name: "Test",
                    children: [
                        // { name: "DOCUMENT DRAWINGS" },
                        // { name: "DOCUMENT REPORTS" }
                    ]
                },
                // {
                //     name: "Policies",
                //     children: [
                //         // { name: "HR" },
                //         // { name: "Finance" }
                //     ]
                // }
            ]
        },{
            name: "BSP",
            children: [
                {
                    name: "HR",
                    children: [
                        // { name: "DOCUMENT DRAWINGS" },
                        // { name: "DOCUMENT REPORTS" }
                    ]
                },
                // {
                //     name: "Policies",
                //     children: [
                //         // { name: "HR" },
                //         // { name: "Finance" }
                //     ]
                // }
            ]
        },
        {
            name: "PLP",
            children: [
                {
                    name: "ADM",
                    children: [
                        // { name: "DOCUMENT DRAWINGS" },
                        // { name: "DOCUMENT REPORTS" }
                    ]
                },
                {
                    name: "BOD",
                    children: [
                        // { name: "HR" },
                        // { name: "Finance" }
                    ]
                },
                {
                    name: "BDP",
                    children: [
                        // { name: "HR" },
                        // { name: "Finance" }
                    ]
                },
                {
                    name: "CEO",
                    children: [
                        // { name: "HR" },
                        // { name: "Finance" }
                    ]
                }
            ]
        }
    ];
    React.useEffect(() => {
        fetchData();
    }, []);

    const fetchData = async () => {
        // Example: Fetch items from a SharePoint list
        const items: any[] = await sp.web.lists.getByTitle("DestinationFolder").items.select("*").getAll();
        // console.log(items);
        // setsouceItems(items);

        if (items && items.length > 0) {
            const sourceItems = items
                .map(i => i.SourceFolder)
                .filter(v => v !== null && v !== undefined && v !== ""); // remove blanks

            const destItems = items
                .map(i => i.DestinationFolder)
                .filter(v => v !== null && v !== undefined && v !== ""); // remove blanks

            setSourceItems(sourceItems);
            setDestItems(destItems);
        }
    }

    const handleFile = async (file: File) => {
        // 🔹 Show loader
        setLoading(true);

        // 🔹 Wait for 2 seconds
        setTimeout(() => {
            setLoading(false); // hide loader (optional, since page reloads)
            Swal.fire('Success', 'The documents request have been submitted successfully. They will be available in the respective folders shortly.', 'success').then(async (result) => {
                if (result.isConfirmed) {
                    window.location.reload();
                }
            });
            // window.location.reload(); // reload the page
        }, 2500);
    };


    const SEP = "||"; // path separator (use something unlikely to appear in names)

    // const FolderTree: React.FC = () => {
    // expanded keys stored as Set of path strings like "EHO||Document Type"
    const [expanded, setExpanded] = React.useState<Set<string>>(new Set());
    // selected breadcrumb path as array of names
    const [selectedPath, setSelectedPath] = React.useState<string[]>([]);

    const getKeyFromPath = (path: string[]) => path.join(SEP);

    // generate all ancestor keys for a pathKey (including itself)
    const getAncestorKeys = (pathKey: string) => {
        const parts = pathKey.split(SEP);
        const keys: string[] = [];
        for (let i = 1; i <= parts.length; i++) {
            keys.push(parts.slice(0, i).join(SEP));
        }
        return keys;
    };

    // toggle expand/collapse; ensure ancestors are expanded when opening;
    // when collapsing, remove descendants too
    const toggleExpand = (pathKey: string) => {
        setExpanded(prev => {
            const s = new Set(prev);
            if (s.has(pathKey)) {
                // collapse: remove this and any descendants
                s.delete(pathKey);
                for (const k of Array.from(s)) {
                    if (k.startsWith(pathKey + SEP)) s.delete(k);
                }
            } else {
                // expand: add all ancestors so the path is visible
                const ancestors = getAncestorKeys(pathKey);
                ancestors.forEach(a => s.add(a));
            }
            return s;
        });
    };

    // when a folder is clicked: update breadcrumbs + toggle expansion
    const onFolderClick = (path: string[]) => {
        setSelectedPath(path);
        const key = getKeyFromPath(path);
        toggleExpand(key);
    };

    // clicking breadcrumb: set it as selected and ensure its ancestors expanded
    const onBreadcrumbClick = (index: number) => {
        const newPath = selectedPath.slice(0, index + 1);
        const key = getKeyFromPath(newPath);
        setSelectedPath(newPath);
        setExpanded(prev => {
            const s = new Set(prev);
            getAncestorKeys(key).forEach(a => s.add(a));
            return s;
        });
    };

    // recursive renderer
    const renderNode = (node: FolderNode, parentPath: string[]) => {
        const currPath = [...parentPath, node.name];
        const key = getKeyFromPath(currPath);
        const isExpanded = expanded.has(key);
        const hasChildren = !!(node.children && node.children.length);

        return (
            <li key={key} style={{ margin: "4px 0", listStyle:"none" }}>
                <button type="button"
                    onClick={() => onFolderClick(currPath)}
                    aria-expanded={isExpanded}
                    style={{
                        cursor: "pointer",
                        border: "none",
                        background: "transparent",
                        padding: 0,
                        font: "inherit",
                        display: "inline-flex",
                        alignItems: "center",
                        gap: 4,
                        // background: "none",
                        // border: "none",
                        color: "#000",
                        fontWeight: "400",
                        marginTop:'0px'
                        // cursor: "pointer",
                        // padding: "4px 8px"
                    }}
                >
                    {/* <span style={{ width: 18, textAlign: "center" }}>
                        {hasChildren ? (isExpanded ? "📂 -" : "📁 +") : "📁"}
                    </span> */}
                    <span style={{ width: 18, textAlign: "center",position:'relative', display: "inline-block" }}>
                        {hasChildren ? (
                            <>
                                <span style={{ marginRight: "4px" }}>{isExpanded ? "📂" : "📁"}</span>

                                <span style={{ cursor: "pointer",top:"1px", position:'absolute', left:'6px' }}>{isExpanded ? "–" : "+"}</span>
                            </>
                        ) : (
                            <span>📁</span>
                        )}
                    </span>
                    <span>{node.name}</span>
                </button>

                {hasChildren && isExpanded && (
                    <ul style={{ marginLeft: 20 }}>
                        {node.children!.map(child => renderNode(child, currPath))}
                    </ul>
                )}
            </li>
        );
    };

// srs new
// const handleFileSelection = async (e: React.ChangeEvent<HTMLInputElement>) => {
//     const file = e.target.files && e.target.files[0];
//     if (!file) {
//         setSelectedFile(null);
//         setPreviewUrl(""); 
//         return;
//     }
//     // if (!file) return;

//     setSelectedFile(file);
//     setIsPreviewLoading(true);
//     setPreviewUrl(""); // Clear previous preview

//     try {
//         // --- 1. UPLOAD TO TEMP LIB FOR PREVIEW ---
//         // We must upload it so SharePoint's iframe can "see" it
//         const folder = sp.web.getFolderByServerRelativePath("TepmBulkUploadLib");
//         const uploadResult = await folder.files.addChunked(file.name, file);
        
//         // Construct the Preview URL
//         const siteUrl = window.location.origin;
//         const serverRelUrl = uploadResult.data.ServerRelativeUrl;
//         const encodedPath = encodeURIComponent(serverRelUrl);
        
//         // Extract /sites/sitename logic
//         const locationPath = window.location.pathname.match(/\/sites\/[^\/]+/)[0];
        
//         // This specific URL format is required for the iframe to render Excel online
//         const finalUrl = `${siteUrl}${locationPath}/TepmBulkUploadLib/Forms/AllItems.aspx?id=${encodedPath}&parent=${encodeURIComponent(serverRelUrl.substring(0, serverRelUrl.lastIndexOf('/')))}`;
        
//         setPreviewUrl(finalUrl);

//         // --- 2. READ EXCEL DATA ---
//         const reader = new FileReader();
//         reader.onload = (evt) => {
//             const bstr = evt.target?.result;
//             const wb = XLSX.read(bstr, { type: 'binary' });
//             const wsname = wb.SheetNames[0];
//             const data = XLSX.utils.sheet_to_json(wb.Sheets[wsname]);
//             setExcelData(data); 
//             console.log("Excel Mapping Loaded:", data);
//         };
//         reader.readAsBinaryString(file);

//     } catch (error) {
//         console.error("Preview failed:", error);
//         Swal.fire("Error", "Could not generate preview. Check if 'TepmBulkUploadLib' exists.", "error");
//     } finally {
//         setIsPreviewLoading(false);
//     }
// };
const handleFileSelection = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files && e.target.files[0];
    if (!file) {
        setSelectedFile(null);
        setExcelData([]);
        return;
    }

    setSelectedFile(file);
    setIsPreviewLoading(true);

    try {
        // --- 1. UPLOAD TO TEMP LIB FOR BACKEND LOGIC ---
        const folder = sp.web.getFolderByServerRelativePath("TepmBulkUploadLib");
        await folder.files.addChunked(file.name, file);
        console.log("File uploaded to TepmBulkUploadLib for backend processing.");

        // --- 2. READ EXCEL DATA LOCALLY FOR HTML TABLE PREVIEW ---
        const reader = new FileReader();
        reader.onload = (evt) => {
            const bstr = evt.target?.result;
            const wb = XLSX.read(bstr, { type: 'binary' });
            const wsname = wb.SheetNames[0];
            const data: any[] = XLSX.utils.sheet_to_json(wb.Sheets[wsname]);
            setExcelData(data); 
            setIsPreviewLoading(false);
        };
        reader.readAsBinaryString(file);

    } catch (error) {
        console.error("Operation failed:", error);
        Swal.fire("Error", "Check if 'TepmBulkUploadLib' exists or permissions are correct.", "error");
        setIsPreviewLoading(false);
    }
};
// srs new
// Final Submit Logic
// const handleFinalSubmit = async () => {
//     // Check if Excel data and Local Files are loaded
//     if (!selectedFile || excelData.length === 0 || localFiles.length === 0) {
//         Swal.fire("Error", "Please select the Excel mapping and the local source folder.", "error");
//         return;
//     }
    
//     setLoading(true);
//     // Initialize progress
//     setUploadProgress({ current: 0, total: excelData.length });
//     let successCount = 0;

//     try {
//         // 1. Get the subsite connection
//         const ehoResult = await sp.site.openWebById("7cef7a05-41ed-4ef8-b247-8e9f1d2b9962");
        
//         // 2. Access the .web property for folder operations
//         const ehoWeb = ehoResult.web; 

//         // for (const row of excelData) {
//         //     const fileNameFromExcel = row["FileName"];
//         //     const docType = row["DocumentType"]; 

//         //     // Explicitly type 'f' as File to resolve the 'any' error
//         //     const matchedFile = localFiles.find((f: File) => f.name === fileNameFromExcel);

//         //     if (matchedFile) {
//         //         // Ensure the path starts from the library level. 
//         //         // If 'Civil' is a folder inside 'Shared Documents', use 'Shared Documents/Civil/...'
//         //         const targetFolderPath = `Civil/${docType}`;
                
//         //         console.log(`Uploading ${matchedFile.name} to ${targetFolderPath}`);

//         //         await ehoWeb.getFolderByServerRelativePath(targetFolderPath)
//         //             .files.addChunked(matchedFile.name, matchedFile);
                
//         //         successCount++;
//         //     }
//         // }
//         for (let i = 0; i < excelData.length; i++) {
//             const row = excelData[i];
//             const fileNameFromExcel = row["FileName"];
//             const docType = row["DocumentType"]; 

//             // Update progressive count for the UI
//             setUploadProgress(prev => ({ ...prev, current: i + 1 }));

//             const matchedFile = localFiles.find((f: File) => f.name === fileNameFromExcel);

//             if (matchedFile) {
//                 const targetFolderPath = `Civil/${docType}`;
//                 await ehoWeb.getFolderByServerRelativePath(targetFolderPath)
//                     .files.addChunked(matchedFile.name, matchedFile);
                
//                 successCount++;
//             }
//         }
//         setLoading(false);
//         Swal.fire("Success", `Processed ${successCount} files successfully.`, "success").then(() => {
//             window.location.reload();
//         });
//     } catch (error) {
//         console.error("Upload error:", error);
//         Swal.fire("Error", "Upload failed. Verify that folders 'Civil/CORRESPONDENCE' and 'Civil/DOCUMENT DRAWINGS' exist in the EHO site.", "error");
//     } 
// };
const handleFinalSubmit = async () => {
    // 1. Basic Validation
    if (!selectedFile || excelData.length === 0 || localFiles.length === 0) {
        Swal.fire("Error", "Please select the Excel mapping and the local source files.", "error");
        return;
    }
    
    setLoading(true);

    // 2. FILTER the data so the progress bar is accurate
    // Only count rows where the FileName in Excel exists in the files you manually picked
    const filesToUpload = excelData.filter(row => 
        localFiles.some((f: File) => f.name === row["FileName"])
    );

    // 3. Set the progress bar total to the actual match count (e.g., 12)
    setUploadProgress({ current: 0, total: filesToUpload.length });
    
    let successCount = 0;

    try {
        const ehoResult = await sp.site.openWebById("7cef7a05-41ed-4ef8-b247-8e9f1d2b9962");
        const ehoWeb = ehoResult.web; 

        // 4. Loop through the filtered list only
        for (let i = 0; i < filesToUpload.length; i++) {
            const row = filesToUpload[i];
            const fileNameFromExcel = row["FileName"];
            const docType = row["DocumentType"]; 

            // Update progress count: 1 of 12, 2 of 12, etc.
            setUploadProgress(prev => ({ ...prev, current: i + 1 }));

            const matchedFile = localFiles.find((f: File) => f.name === fileNameFromExcel);

            if (matchedFile) {
                const targetFolderPath = `Civil/${docType}`;
                await ehoWeb.getFolderByServerRelativePath(targetFolderPath)
                    .files.addChunked(matchedFile.name, matchedFile);
                
                successCount++;
            }
        }

        setLoading(false);
        Swal.fire("Success", `Processed ${successCount} files successfully.`, "success").then(() => {
            window.location.reload();
        });

    } catch (error) {
        setLoading(false);
        console.error("Upload error:", error);
        Swal.fire("Error", "Upload failed. Verify folder paths in SharePoint.", "error");
    } 
};



    return (
        <>
        <div>
           



            {loading && (
                <div className='loaderOverlay'>
                    <div className='loader'>
                        <img style={{ width: '116px', margin: '31px' }} src={require("../assets/ESSAROLLER.gif")} alt="Loading..." />
                        {/* srs progressive count */}
            <div style={{ color: '#000', fontWeight: 'bold', fontSize: '18px', marginTop: '-20px', paddingBottom: '20px' }}>
                Uploading {uploadProgress.current} of {uploadProgress.total}...
            </div>
                    </div>
                </div>
            )}
            <div className="card mt-0">
                
                <div className="card-body">
                     <h3 className="mb-3 fw-bold text-dark header-title">Bulk Upload</h3>

                    <div className="row">

                    <div className="col-sm-12">
                           


                        </div>

                        {/* <div className="col-sm-6"> */}
{/* srs 8/1/26 */}
                            {/* <div className="borderprev" style={{ border: "1px solid #ccc", height: "400px", position: "relative" }}>
    {isPreviewLoading && (
        <div style={{ position: "absolute", top: "50%", left: "50%", transform: "translate(-50%, -50%)" }}>
            <img src={require("../assets/ESSAROLLER.gif")} style={{ width: "50px" }} />
            <p>Loading Preview...</p>
        </div>
    )}
    {previewUrl ? (
        <iframe src={previewUrl} width="100%" height="100%" style={{ border: "none" }} />
    ) : (
        !isPreviewLoading && <div style={{ textAlign: "center", paddingTop: "180px" }}>Select a file to preview</div>
    )}
</div> */}
{/* srs 8/1/26 */}
                            {/* 🔹 Source Folder Dropdown */}
                            {/* <label>Insert source folder path : </label>
                            <input style={{height:'60px'}} type="text" className="form-control" value={selectedSource} onChange={(e) => setSelectedSource(e.target.value)} /> */}
{/* srs new */}
                            {/* <label>Select Source Folder containing the actual files:</label>
<input 
    type="file" 
    multiple 
    // @ts-ignore
    webkitdirectory="true" 
    onChange={(e) => {
        const files = Array.from(e.target.files);
        setLocalFiles(files);
    }} 
/> */}
                        
                        {/* </div> */}

     
      
                        <div  className="col-sm-6 mt-3">
                        <div style={{  marginBottom: "20px" }}>
                                <div>
                                    <label htmlFor="file-upload">Select file:</label>

                                    {/* <input
                                        type="file"
                                        id="file-upload"
                                        accept=".xlsx, .xls,.csv"
                                        onChange={e => {
                                            const file = e.target.files && e.target.files[0];
                                            setSelectedFile(file || null);
                                        }}
                                    /> */}
                                    {/* srs 8/1/26 */}
                                    <input   style={{height:'40px', padding:'5px 10px'}}  className="form-control"
    type="file"
    id="file-upload"
    multiple
                    // @ts-ignore
                    // webkitdirectory="true"
                    accept=".pdf, .doc, .docx, .xlsx, .xls, .csv, .txt"
    onChange={handleFileSelection}
/>
                                </div>

                            </div>
                            {/* 🔹 Destination Folder Dropdown */}
                            <button className="newselc" type="button" onClick={() => setShowModal(true)}>  <img className="sidebariconssmall" src={folderdes}></img> Select Destination Folder</button>
                            {/* Breadcrumbs */}
                            <div style={{ marginBottom: 10  ,marginTop: 10  , display: "flex", justifyContent:'start'}}>
                                {selectedPath.length === 0 ? (
                                    <span style={{ color: "#666" }}>No folder selected</span>
                                ) : (
                                    selectedPath.map((p, i) => (
                                        <span style={{display:"flex", alignItems:"center",marginTop:"10px"}} key={i}>
                                            <button type="button" className="newhover"
                                                onClick={() => onBreadcrumbClick(i)}
                                                style={{
                                                    background: "transparent",
                                                    border: "none",
                                                    color: i === selectedPath.length - 1 ? "#000" : "#0078d4",
                                                    cursor: "pointer", fontSize:"12px",
                                                    padding: 0, marginTop:"0px",
                                                    marginRight: 6,
                                                    textDecoration: i === selectedPath.length - 1 ? "none" : "underline"
                                                }}
                                            >
                                                {p}
                                            </button>
                                            {i < selectedPath.length - 1 && <span style={{ marginRight: 6 }}>{">"}</span>}
                                        </span>
                                    ))
                                )}
                            </div>
                            {/* <select className="form-select"
                                value={selectedDest}
                                onChange={(e) => setSelectedDest(e.target.value)}
                            >
                                <option value="">-- Select Destination --</option>
                                {destItems.map((item, idx) => (
                                    <option key={idx} value={item}>
                                        {item}
                                    </option>
                                ))}
                            </select> */}

                        </div>

                        <div className="col-sm-6 mt-3">
                        <div className="" style={{ 
                border: " 0px solid #ccc", 
               
                overflow: "auto", 
                background: "#fff",
                borderRadius: "0px"
            }}>
    
      <div>
           
               
            
            <div className="p-2 pt-0">
                <label className="mb-1">Select Source Folder containing the actual files:</label>
                <input  style={{height:'40px', padding:'5px 10px'}}
                    className="form-control"
                    type="file" 
                    multiple 
                  
                     // @ts-ignore
                    // webkitdirectory="true"
                    accept=".pdf, .doc, .docx, .xlsx, .xls, .csv, .txt"
                   
                    onChange={(e) => {
                        const files = Array.from(e.target.files);
                        setLocalFiles(files);
                    }} 
                />
                {localFiles.length > 0 && (
                    <small className="text-success fw-bold">
                        {localFiles.length} files detected in local folder.
                    </small>
                )}
            </div>
            </div>
        

     </div>
    
</div>

                        <div className="col-sm-12">
                        {/* {isPreviewLoading ? (
                    <div style={{ textAlign: "center", paddingTop: "150px" }}>
                        <img src={require("../assets/ESSAROLLER.gif")} style={{ width: "50px" }} />
                        <p className="mt-2">Uploading & Processing...</p>
                    </div>
                ) : (
                    

<div style={{ width: "100%", overflowX: "auto", maxHeight: "500px", overflowY: "auto" }}>
  <table
    className="mtable table-sm table-hover bloptable"
    style={{
      fontSize: "13px",
      width: "100%",
      borderCollapse: "collapse"
    }}
  >
    
    <thead
      className="table-light"
      style={{ position: "sticky", top: 0, zIndex: 1 }}
    >
      <tr>
        {COLUMNS.map((col:any, i:any) => (
          <th
            key={i}
            style={{
              whiteSpace: "nowrap",
              padding: "8px 12px",
              borderBottom: "1px solid #dee2e6",
              fontWeight: 600
            }}
          >
            {col}
          </th>
        ))}
      </tr>
    </thead>

   
    <tbody>
      {excelData.length > 0 ? (
        excelData.map((row, rIndex) => (
          <tr key={rIndex}>
            {COLUMNS.map((col:any, cIndex:any) => (
              <td
                key={cIndex}
                style={{
                  padding: "8px 12px",
                  borderBottom: "1px solid #dee2e6",
                  whiteSpace: "nowrap"
                }}
              >
                {row[col] ?? "\u00A0"}
              </td>
            ))}
          </tr>
        ))
      ) : (
        <tr>
          <td colSpan={COLUMNS.length} style={{ textAlign: "center", padding: 20 }}>
            No data found in sheet.
          </td>
        </tr>
      )}
    </tbody>
  </table>
</div>

                )} */}
                {isPreviewLoading && (
  <div style={{ textAlign: "center", paddingTop: "150px" }}>
    <img src={require("../assets/ESSAROLLER.gif")} style={{ width: "50px" }} />
    <p className="mt-2">Uploading & Processing...</p>
  </div>
)}

{/* ✅ SHOW TABLE ONLY WHEN FILE IS SELECTED */}
{!isPreviewLoading && selectedFile && (
  <div style={{ width: "100%",  overflowY: "auto", display:'grid' }}>
    <table
      className="mtable table-sm table-hover bloptable"
      style={{
        fontSize: "13px",
        width: "100%", borderRadius:'0px',
        borderCollapse: "collapse",overflow: "auto", maxHeight: "500px",
      }}
    >
      <thead
        className="table-light"
        style={{ position: "sticky", top: 0, zIndex: 1 }}
      >
        <tr>
          {COLUMNS.map((col, i) => (
            <th
              key={i}
              style={{
                whiteSpace: "nowrap",
                padding: "8px 12px",
                borderBottom: "1px solid #dee2e6",
                fontWeight: 600
              }}
            >
              {col}
            </th>
          ))}
        </tr>
      </thead>

      <tbody style={{overflow:'visible',maxHeight:'50000px'}}>
        {excelData.length > 0 ? (
          excelData.map((row, rIndex) => (
            <tr key={rIndex}>
              {COLUMNS.map((col, cIndex) => (
                <td
                  key={cIndex}
                  style={{
                    padding: "8px 12px",
                    borderBottom: "1px solid #dee2e6",
                    whiteSpace: "nowrap"
                  }}
                >
                  {row[col] ?? "\u00A0"}
                </td>
              ))}
            </tr>
          ))
        ) : (
          <tr>
            <td colSpan={COLUMNS.length} style={{ textAlign: "center", padding: 20 }}>
              No data found in sheet.
            </td>
          </tr>
        )}
      </tbody>
    </table>
  </div>
)}

                        </div>

                        



                       
                    </div>
                </div>
                
                
                </div>

                <div style={{ flex: 1, display: "flex", gap: '10px', justifyContent: "end", alignItems: 'center' }} className="newrequ">
                            <button className="btncolorCreate1"
                                type="button"
                                // onClick={() => {
                                //     handleFile((document.getElementById("file-upload") as HTMLInputElement).files![0]);
                                // }}
                                // srs new 
                                onClick={handleFinalSubmit}
                                style={{ height: "36px" }}
                                // disabled={!selectedFile || !selectedSource.trim() || !selectedPath.length}
                            >
                                
 <span className="mb-1 mt-2" data-tooltip="Submit">
                                      <img  src={require("../assets/submit-new.png")}/></span>
                            </button>

                            <button className="btncolorCreate1 alitool"
                                type="button"
                                onClick={() => {
                                    window.location.reload();
                                }}
                                style={{ height: "36px" }}

                            >
                                <span className="mb-1 mt-2" data-tooltip="Cancel">
                                <img  src={require("../assets/cancelnew.png")}/></span>
                            </button>

                            {/* <button
                                    type="button"
                                    style={{ height: "36px" }}
                                    disabled={selectedRow === null}
                                    onClick={() => {
                                        if (selectedRow !== null) {
                                            setEditingRow(selectedRow);
                                            setEditData(listItems[selectedRow]);
                                        }
                                    }}
                                >
                                    Edit Selected Row
                                </button> */}

                        </div>
            <Modal show={showModal} onHide={() => setShowModal(false)} className='newmobmodal' id="popfolder">
                <Modal.Header closeButton>
                    <Modal.Title > <h4 className='font-16 text-dark fw-bold mb-1'>Select Destination Folder</h4>
                        {/* <p className='text-muted font-14 mb-0 fw-400'>Below are the attachment details for IMS Audit Plan
                        </p> */}

                    </Modal.Title>


                </Modal.Header>

                <Modal.Body className="" id="style-5">
                    <div className="treeview">



                        {/* Tree */}
                        <ul style={{ paddingLeft: 0 }}>
                            {folderData.map(root => renderNode(root, []))}
                        </ul>
                    </div>
                </Modal.Body>
            </Modal>

        </div>
      </>
    )
}

export default uploadFileInDestination
