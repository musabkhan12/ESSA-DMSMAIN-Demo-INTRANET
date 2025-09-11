import * as React from "react";
import * as XLSX from 'xlsx';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { forEach } from "jszip";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp/presets/all";

const bulkUpload = () => {
    const sp: SPFI = getSP();

    const [listItems, setItems] = React.useState<any[]>([]);

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
                return;
            }
            let bulkArr =[];

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

    const saveItem = async (bulkArr:any) => {
        bulkArr.forEach(async (item:any) => {
            await sp.web.lists.getByTitle("DocumentBulkUpload").items.add(item);
        }

       
        );

        setTimeout(() => {
            fetchData();
        }, 1000);
    
        // const addData = await sp.web.lists.getByTitle(currentList).items.add(item);
    }
    return (
        <div>

            <div>
                <div style={{ width: "25%" }}>
                    <label htmlFor="file-upload">Attach Document:</label>
                    <input type="file" id="file-upload" onChange={(e) => handleFile(e.target.files[0])} />
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
                <table className="bulk-upload-table" style={{ border: "1px solid #888", borderCollapse: "collapse" }}>
                    <thead>
                        <tr>
                            <th>DocumentType</th>
                            <th>Template</th>
                            <th>External Party</th>
                            <th>From</th>
                            <th>Issued Date</th>
                            <th>Year</th>
                            <th>Subject</th>
                            <th>Project</th>
                            <th>TagNumber</th>
                            <th>Area</th>
                        </tr>
                    </thead>
                    <tbody>
                        {listItems.map((item, index) => (
                            <tr key={index}>
                                <td>{item.DocumentType}</td>
                                <td>{item.Template}</td>
                                <td>{item.ExternalParty}</td>
                                <td>{item.From ? new Date(item.From).toLocaleDateString('en-GB') : ''}</td>
                                <td>{item.IssuedDate ? new Date(item.IssuedDate).toLocaleDateString('en-GB') : ''}</td>
                                <td>{item.Year}</td>
                                <td>{item.Subject}</td>
                                <td>{item.Project}</td>
                                <td>{item.TagNo}</td>
                                <td>{item.Area}</td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>

        </div>
    )
}

export default bulkUpload
