import * as React from "react";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp/presets/all";
import Swal from "sweetalert2";
import { Modal } from 'react-bootstrap';
let folderdes = require('../assets/folderdes.png');
import '../components/uploadfilecss.css';

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
                    name: "CIV",
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


    return (
        <div>
           



            {loading && (
                <div className='loaderOverlay'>
                    <div className='loader'>
                        <img style={{ width: '116px', margin: '31px' }} src={require("../assets/ESSAROLLER.gif")} alt="Loading..." />
                    </div>
                </div>
            )}
            <div className="card mt-0">
                
                <div className="card-body">
                     <h3 className="mb-3 fw-bold text-dark header-title">Bulk Upload</h3>

                    <div className="row">

                    <div className="col-sm-12">
                            <div style={{  marginBottom: "20px" }}>
                                <div>
                                    <label htmlFor="file-upload">Select file:</label>

                                    <input
                                        type="file"
                                        id="file-upload"
                                        accept=".xlsx, .xls,.csv"
                                        onChange={e => {
                                            const file = e.target.files && e.target.files[0];
                                            setSelectedFile(file || null);
                                        }}
                                    />
                                </div>

                            </div>


                        </div>

                        <div className="col-sm-6">
                            {/* 🔹 Source Folder Dropdown */}
                            <label>Insert source folder path : </label>
                            <input style={{height:'60px'}} type="text" className="form-control" value={selectedSource} onChange={(e) => setSelectedSource(e.target.value)} />
                            {/* <select className="form-select"
                                value={selectedSource}
                                onChange={(e) => setSelectedSource(e.target.value)}
                            >
                                <option value="">-- Select Source --</option>
                                {sourceItems.map((item, idx) => (
                                    <option key={idx} value={item}>
                                        {item}
                                    </option>
                                ))}
                            </select> */}
                        </div>
                        <div style={{textAlign:'right'}} className="col-sm-6 mt-3">

                            {/* 🔹 Destination Folder Dropdown */}
                            <button className="newselc" type="button" onClick={() => setShowModal(true)}>  <img className="sidebariconssmall" src={folderdes}></img> Select Destination Folder</button>
                            {/* Breadcrumbs */}
                            <div style={{ marginBottom: 10  ,marginTop: 10  , display: "flex", justifyContent:'end'}}>
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

                        



                       
                    </div>
                </div>
                
                
                </div>

                <div style={{ flex: 1, display: "flex", gap: '10px', justifyContent: "end", alignItems: 'center' }} className="newrequ">
                            <button className="btncolorCreate1"
                                type="button"
                                onClick={() => {
                                    handleFile((document.getElementById("file-upload") as HTMLInputElement).files![0]);
                                }}
                                style={{ height: "36px" }}
                                disabled={!selectedFile || !selectedSource.trim() || !selectedPath.length}
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

    )
}

export default uploadFileInDestination
