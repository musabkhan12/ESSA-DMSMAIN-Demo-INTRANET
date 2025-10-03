import * as React from "react";
import * as XLSX from 'xlsx';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { forEach } from "jszip";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp/presets/all";
import Swal from "sweetalert2";

const bulkUpload = () => {
    const sp: SPFI = getSP();

    const [listItems, setItems] = React.useState<any[]>([]);

    const [selectedFile, setSelectedFile] = React.useState<File | null>(null);
    const [selectedRow, setSelectedRow] = React.useState<number | null>(null);
    const [selectedRowChecked, setSelectedRowChecked] = React.useState<number | null>(null);
    const [editingRow, setEditingRow] = React.useState<number | null>(null);
    const [editData, setEditData] = React.useState<any>({});

    React.useEffect(() => {
        // Any side effects or data fetching can go here
        fetchData();
    }, []);

    const fetchData = async () => {
        // Example: Fetch items from a SharePoint list
        const items: any[] = await sp.web.lists.getByTitle("DocumentBulkUpload").items.select("*").getAll();
        // console.log(items);
        setItems(items);
    }

    const handleFile = async (file: File) => {
        const reader = new FileReader();

        reader.onload = (e: any) => {

            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            console.log('Excel Data:', jsonData);
            // Now you can map jsonData to SP list columns

            // Extract headers
            const headers = jsonData[0] as string[];

            // Find index of DocumentType column
            const documentTypeIndex = headers.findIndex(h => h === 'DocumentType');

            if (documentTypeIndex === -1) {
                console.error('DocumentType column not found');
                Swal.fire('Error', 'Invalid file. DocumentType column is missing.', 'error');
                return;
            }
            let bulkArr = [];

            // Loop through all data rows except header
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i] as any[]; // Type assertion for each row

                const documentTypeValue = row[documentTypeIndex];

                // Convert Excel date number to JS date string (yyyy-mm-dd) for IssuedDate
                // Handle Issued Date
                let issuedDateValue = row[headers.findIndex(h => h === 'Issued Date')];
                let issuedDateFormatted = null;
                if (issuedDateValue && !isNaN(issuedDateValue)) {
                    // Excel date numbers: days since 1899-12-31
                    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
                    const jsDate = new Date(excelEpoch.getTime() + (issuedDateValue * 24 * 60 * 60 * 1000));
                    // Format as yyyy-mm-dd for SharePoint
                    // issuedDateFormatted = jsDate.toISOString().split('T')[0];
                    // issuedDateFormatted = jsDate.toISOString();
                    issuedDateFormatted = jsDate
                }

                // Handle From (if it's a date)
                let fromValue = row[headers.findIndex(h => h === 'From')];
                let fromFormatted = null;
                if (fromValue && !isNaN(fromValue)) {
                    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
                    const jsDate = new Date(excelEpoch.getTime() + (fromValue * 24 * 60 * 60 * 1000));
                    // fromFormatted = jsDate.toISOString();
                    fromFormatted = jsDate
                } else {
                    fromFormatted = fromValue || null;
                }

                const yearValue = row[headers.findIndex(h => h === 'Year')];
                const item = {
                    DocumentType: row[headers.findIndex(h => h === 'DocumentType')] || '',
                    Template: row[headers.findIndex(h => h === 'Template')] || '',
                    ExternalParty: row[headers.findIndex(h => h === 'External Party')] || '',
                    From: fromFormatted,
                    IssuedDate: issuedDateFormatted,
                    Year: yearValue !== undefined && yearValue !== null ? String(yearValue) : '',
                    Subject: row[headers.findIndex(h => h === 'Subject')] || '',
                    Project: row[headers.findIndex(h => h === 'Project')] || '',
                    TagNo: row[headers.findIndex(h => h === 'Tag No.')] || '',
                    Area: row[headers.findIndex(h => h === 'Area')] || ''
                };
                // Clear the file input after processing the file
                const fileInput = document.getElementById("file-upload") as HTMLInputElement;
                if (fileInput) {
                    fileInput.value = "";
                }
                bulkArr.push(item);

                // console.log('Item to save:', item);


            }
            saveItem(bulkArr);

        };

        reader.readAsArrayBuffer(file);
    };

    const saveItem = async (bulkArr: any) => {
        bulkArr.forEach(async (item: any) => {
            await sp.web.lists.getByTitle("DocumentBulkUpload").items.add(item);
        }


        );

        Swal.fire({
            title: "Success!",
            text: "Data added successfully.",
            icon: "success"
        });

        setTimeout(() => {
            fetchData();
        }, 1000);

        // const addData = await sp.web.lists.getByTitle(currentList).items.add(item);
    }
    // Function to handle saving the edited row
    const handleSaveEdit = async (rowIndex: number) => {
        if (editingRow === null) return;
        const itemToEdit = listItems[editingRow];
        try {
            await sp.web.lists.getByTitle("DocumentBulkUpload").items.getById(itemToEdit.Id).update(editData);
            Swal.fire({
                title: "Success!",
                text: "Row updated successfully.",
                icon: "success"
            });
            setEditingRow(null);
            setEditData({});
            setSelectedRow(null); // Uncheck all checkboxes after save
            fetchData();
        } catch (error) {
            Swal.fire("Error", "Failed to update row.", "error");
        }
    };

    return (
        <div>

            <div>
                <div style={{ display: "flex", alignItems: "flex-end", gap: "16px", marginBottom: "20px" }}>
                    <div style={{ width: "25%" }}>
                        <label htmlFor="file-upload">Attach Document:</label>
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
                    <button
                        type="button"
                        onClick={() => {
                            handleFile((document.getElementById("file-upload") as HTMLInputElement).files![0]);
                        }}
                        style={{ height: "36px" }}
                        disabled={!selectedFile}
                    >
                        Submit
                    </button>
                    <div style={{ flex: 1, display: "flex", justifyContent: "center" }}>
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
                        <button
                            type="button"
                            style={{ height: "36px" }}
                            onClick={async () => {
                                // Download the first file from BulkUploadTemplate library
                                try {
                                    const files = await sp.web.lists.getByTitle("BulkUploadTemplate").items.select("FileRef", "FileLeafRef").top(1).getAll();
                                    if (files.length > 0) {
                                        const fileUrl = files[0].FileRef;
                                        const fileName = files[0].FileLeafRef;
                                        // Create a temporary link to download
                                        const link = document.createElement("a");
                                        link.href = fileUrl;
                                        link.download = fileName;
                                        document.body.appendChild(link);
                                        link.click();
                                        document.body.removeChild(link);
                                    } else {
                                        Swal.fire("Error", "No template file found in BulkUploadTemplate library.", "error");
                                    }
                                } catch (err) {
                                    Swal.fire("Error", "Failed to download template file.", "error");
                                }
                            }}
                        >
                            Download Template
                        </button>
                    </div>
                </div>


                <style>
                    {`
                        .bulk-upload-table {
                            width: 100%;
                            margin-top: 20px;
                            border-collapse: collapse;
                        }
                        .bulk-upload-table th, .bulk-upload-table td {
                            border: 1px solid #ccc;
                            padding: 8px;
                            text-align: left;
                        }
                        .bulk-upload-table th {
                            background-color: #f2f2f2;
                        }
                        .bulk-upload-table tbody tr:nth-child(even) {
                            background-color: #fafafa;
                        }
                        .bulk-upload-table tbody tr:nth-child(odd) {
                            background-color: #fff;
                        }
                    `}
                </style>
                <div style={{ overflowX: "auto" }}>
                    <table
                        className="bulk-upload-table"
                        style={{
                            minWidth: "1200px",
                            border: "none",
                            borderCollapse: "separate",
                            borderSpacing: "0",
                            // boxShadow: "0 2px 8px rgba(0,0,0,0.08)",
                            // background: "#fff",
                        }}
                    >
                        <thead>
                            <tr>
                                <th style={{
                                    minWidth: 140,
                                    background: "#43a047", // Green color
                                    color: "#fff",
                                    fontWeight: 600,
                                    borderBottom: "2px solid #e0e0e0"
                                }}>DocumentType</th>
                                <th style={{
                                    minWidth: 140,
                                    background: "#43a047",
                                    color: "#fff",
                                    fontWeight: 600,
                                    borderBottom: "2px solid #e0e0e0"
                                }}>Template</th>
                                <th style={{
                                    minWidth: 140,
                                    background: "#43a047",
                                    color: "#fff",
                                    fontWeight: 600,
                                    borderBottom: "2px solid #e0e0e0"
                                }}>External Party</th>
                                <th style={{
                                    minWidth: 140,
                                    background: "#43a047",
                                    color: "#fff",
                                    fontWeight: 600,
                                    borderBottom: "2px solid #e0e0e0"
                                }}>From</th>
                                <th style={{
                                    minWidth: 140,
                                    background: "#43a047",
                                    color: "#fff",
                                    fontWeight: 600,
                                    borderBottom: "2px solid #e0e0e0"
                                }}>Issued Date</th>
                                <th style={{
                                    minWidth: 140,
                                    background: "#43a047",
                                    color: "#fff",
                                    fontWeight: 600,
                                    borderBottom: "2px solid #e0e0e0"
                                }}>Year</th>
                                <th style={{
                                    minWidth: 140,
                                    background: "#43a047",
                                    color: "#fff",
                                    fontWeight: 600,
                                    borderBottom: "2px solid #e0e0e0"
                                }}>Subject</th>
                                <th style={{
                                    minWidth: 140,
                                    background: "#43a047",
                                    color: "#fff",
                                    fontWeight: 600,
                                    borderBottom: "2px solid #e0e0e0"
                                }}>Project</th>
                                <th style={{
                                    minWidth: 140,
                                    background: "#43a047",
                                    color: "#fff",
                                    fontWeight: 600,
                                    borderBottom: "2px solid #e0e0e0"
                                }}>TagNumber</th>
                                <th style={{
                                    minWidth: 140,
                                    background: "#43a047",
                                    color: "#fff",
                                    fontWeight: 600,
                                    borderBottom: "2px solid #e0e0e0"
                                }}>Area</th>
                                {/* {selectedRow !== null && (
                                    <th
                                        style={{
                                            minWidth: 120,
                                            background: "#43a047",
                                            color: "#fff",
                                            borderBottom: "2px solid #e0e0e0"
                                        }}
                                    ></th>
                                )} */}
                            </tr>
                        </thead>
                        <tbody>
                            {listItems.map((item, index) => (
                                <tr key={index} style={{ borderBottom: "1px solid #e0e0e0", transition: "background 0.2s" }}>
                                    {editingRow === index ? (
                                        <>
                                            <td style={{ minWidth: 140 }}>
                                                <input
                                                    type="text"
                                                    value={editData.DocumentType}
                                                    onChange={e => setEditData({ ...editData, DocumentType: e.target.value })}
                                                    style={{ width: "90%", padding: "6px", borderRadius: 4, border: "1px solid #d0d0d0" }}
                                                />
                                            </td>
                                            <td style={{ minWidth: 140 }}>
                                                <input
                                                    type="text"
                                                    value={editData.Template}
                                                    onChange={e => setEditData({ ...editData, Template: e.target.value })}
                                                    style={{ width: "90%", padding: "6px", borderRadius: 4, border: "1px solid #d0d0d0" }}
                                                />
                                            </td>
                                            <td style={{ minWidth: 140 }}>
                                                <input
                                                    type="text"
                                                    value={editData.ExternalParty}
                                                    onChange={e => setEditData({ ...editData, ExternalParty: e.target.value })}
                                                    style={{ width: "90%", padding: "6px", borderRadius: 4, border: "1px solid #d0d0d0" }}
                                                />
                                            </td>
                                            <td style={{ minWidth: 140 }}>
                                                <input
                                                    type="date"
                                                    value={editData.From ? new Date(editData.From).toISOString().split('T')[0] : ''}
                                                    onChange={e => setEditData({ ...editData, From: e.target.value })}
                                                    style={{ width: "90%", padding: "6px", borderRadius: 4, border: "1px solid #d0d0d0" }}
                                                />
                                            </td>
                                            <td style={{ minWidth: 140 }}>
                                                <input
                                                    type="date"
                                                    value={editData.IssuedDate ? new Date(editData.IssuedDate).toISOString().split('T')[0] : ''}
                                                    onChange={e => setEditData({ ...editData, IssuedDate: e.target.value })}
                                                    style={{ width: "90%", padding: "6px", borderRadius: 4, border: "1px solid #d0d0d0" }}
                                                />
                                            </td>
                                            <td style={{ minWidth: 140 }}>
                                                <input
                                                    type="text"
                                                    value={editData.Year}
                                                    onChange={e => setEditData({ ...editData, Year: e.target.value })}
                                                    style={{ width: "90%", padding: "6px", borderRadius: 4, border: "1px solid #d0d0d0" }}
                                                />
                                            </td>
                                            <td style={{ minWidth: 140 }}>
                                                <input
                                                    type="text"
                                                    value={editData.Subject}
                                                    onChange={e => setEditData({ ...editData, Subject: e.target.value })}
                                                    style={{ width: "90%", padding: "6px", borderRadius: 4, border: "1px solid #d0d0d0" }}
                                                />
                                            </td>
                                            <td style={{ minWidth: 140 }}>
                                                <input
                                                    type="text"
                                                    value={editData.Project}
                                                    onChange={e => setEditData({ ...editData, Project: e.target.value })}
                                                    style={{ width: "90%", padding: "6px", borderRadius: 4, border: "1px solid #d0d0d0" }}
                                                />
                                            </td>
                                            <td style={{ minWidth: 140 }}>
                                                <input
                                                    type="text"
                                                    value={editData.TagNo}
                                                    onChange={e => setEditData({ ...editData, TagNo: e.target.value })}
                                                    style={{ width: "90%", padding: "6px", borderRadius: 4, border: "1px solid #d0d0d0" }}
                                                />
                                            </td>
                                            <td style={{ minWidth: 140 }}>
                                                <input
                                                    type="text"
                                                    value={editData.Area}
                                                    onChange={e => setEditData({ ...editData, Area: e.target.value })}
                                                    style={{ width: "90%", padding: "6px", borderRadius: 4, border: "1px solid #d0d0d0" }}
                                                />
                                            </td>
                                            <td style={{ minWidth: 120, visibility: selectedRow === index ? "visible" : "hidden" }}>
                                                {selectedRow === index && (
                                                    <>
                                                        <button
                                                            type="button"
                                                            onClick={() => handleSaveEdit(index)}
                                                            style={{
                                                                background: "#1976d2",
                                                                color: "#fff",
                                                                border: "none",
                                                                borderRadius: 4,
                                                                padding: "6px 12px",
                                                                marginRight: 8,
                                                                cursor: "pointer",
                                                            }}
                                                        >
                                                            Save
                                                        </button>
                                                        <button
                                                            type="button"
                                                            onClick={() => setEditingRow(null)}
                                                            style={{
                                                                background: "#e0e0e0",
                                                                color: "#333",
                                                                border: "none",
                                                                borderRadius: 4,
                                                                padding: "6px 12px",
                                                                cursor: "pointer",
                                                            }}
                                                        >
                                                            Cancel
                                                        </button>
                                                    </>
                                                )}
                                            </td>
                                        </>
                                    ) : (
                                        <>
                                            <td
                                                style={{
                                                    minWidth: 140,
                                                    background: selectedRow === index ? "#e3f2fd" : undefined,
                                                    padding: "8px",
                                                }}
                                                className="bulk-upload-checkbox-cell"
                                                onMouseEnter={() => setSelectedRowChecked(index)}
                                                onMouseLeave={() => setSelectedRowChecked(null)}
                                            >
                                                <div style={{
                                                    display: "flex",
                                                    alignItems: "center"
                                                }}>
                                                    <input
                                                        type="checkbox"
                                                        style={{
                                                            accentColor: "#1976d2",
                                                            borderRadius: "50%",
                                                            marginRight: 8,
                                                            opacity: selectedRowChecked === index || selectedRow === index ? 1 : 0,
                                                            transition: "opacity 0.2s ease"
                                                        }}
                                                        checked={selectedRow === index}
                                                        onChange={e => {
                                                            setSelectedRow(e.target.checked ? index : null);
                                                        }}
                                                        className="bulk-upload-checkbox"
                                                    />
                                                    <span>{item.DocumentType}</span>
                                                </div>
                                            </td>



                                            <td style={{ minWidth: 140 }}>{item.Template}</td>
                                            <td style={{ minWidth: 140 }}>{item.ExternalParty}</td>
                                            <td style={{ minWidth: 140 }}>{item.From ? new Date(item.From).toLocaleDateString('en-GB') : ''}</td>
                                            <td style={{ minWidth: 140 }}>{item.IssuedDate ? new Date(item.IssuedDate).toLocaleDateString('en-GB') : ''}</td>
                                            <td style={{ minWidth: 140 }}>{item.Year}</td>
                                            <td style={{ minWidth: 140 }}>{item.Subject}</td>
                                            <td style={{ minWidth: 140 }}>{item.Project}</td>
                                            <td style={{ minWidth: 140 }}>{item.TagNo}</td>
                                            <td style={{ minWidth: 140 }}>{item.Area}</td>
                                            <td style={{ minWidth: 120, visibility: selectedRow === index ? "visible" : "hidden" }}>
                                                {selectedRow === index && (
                                                    <button
                                                        type="button"
                                                        onClick={() => {
                                                            setEditingRow(index);
                                                            setEditData(item);
                                                        }}
                                                        style={{
                                                            background: "#1976d2",
                                                            color: "#fff",
                                                            border: "none",
                                                            borderRadius: 4,
                                                            padding: "6px 12px",
                                                            cursor: "pointer",
                                                        }}
                                                    >
                                                        Edit Row
                                                    </button>
                                                )}
                                            </td>
                                        </>
                                    )}
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>

        </div>
    )
}

export default bulkUpload
