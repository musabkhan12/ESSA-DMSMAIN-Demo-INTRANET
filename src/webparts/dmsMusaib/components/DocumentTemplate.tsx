import React, { useEffect, useState } from "react";
import { getSP } from "../loc/pnpjsConfig";
import Select from "react-select";
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
import Swal from "sweetalert2";

let siteID = "9f942ad9-f2b6-4d4a-b99c-80f1171b57e3";
let isApproval :any
const DocumentTemplate = (props: any) => {
  const sp: SPFI = getSP();
  let locationPath=window.location.pathname.match(/\/sites\/[^\/]+/)[0];
  const [subCategories, setSubCategories] = useState<any[]>([]);
  const [selectedSubCategory, setSelectedSubCategory] = useState<any | null>(
    null
  );
  const [fields, setFields] = useState<any[]>([]);
  const [formValues, setFormValues] = useState<any>({});
  const [files, setFiles] = useState<File[]>([]);
  const [checkapprovaltype, setCheckapprovaltype] = useState<boolean>(false);
  const [users, setUsers] = React.useState<any[]>([]);
  const [selectedUsers, setSelectedUsers] = useState<any[]>([]);
    const [errorsForUserSelection,setErrorsForUserSelection]=useState<{ [key: number]: { userSelect?: string} }>({});
    const [rows, setRows] = React.useState<
      { id: number; selectionType: "All" | "One"; approvedUserList: string[] }[]
    >([{ id: 0, selectionType: "One", approvedUserList: [] }]);
  useEffect(() => {
    if (props.selectedCategory) {
      fetchData(props.selectedCategory);
    }
  }, [props.selectedCategory]);

  // fetch subcategories
  const fetchData = async (selectedCategory: string) => {
    try {
      
      const items = await sp.web.lists
        .getByTitle("TemplateDocumentCategorySubCategoryMapping")
        .items.select(
          "DocumentCategory/DocumentCategory",
          "DocumentSubCategory/DocumentSubCategory"
        )
        .expand("DocumentCategory", "DocumentSubCategory")
        .filter(`DocumentCategory/DocumentCategory eq '${selectedCategory}'`)
        .getAll();

      const formatted = items.map((item: any) => ({
        label: item.DocumentSubCategory.DocumentSubCategory,
        value: item.DocumentSubCategory.DocumentSubCategory,
      }));

      setSubCategories(formatted);
      setSelectedSubCategory(null);
    } catch (error) {
      console.error("Error fetching subcategories:", error);
    }
  };

  // fetch fields mapping
  const handleSubCategorySelect = async (selectedOption: any) => {
    setSelectedSubCategory(selectedOption);
    console.log("Selected SubCategory:", selectedSubCategory);
    if (!selectedOption) return;

    const items = await sp.web.lists
      .getByTitle("TemplateDocumentCategoryFieldsMapping")
      .items.filter(`DocumentSubCategory eq '${selectedOption.value}'`)
      .getAll();

    const formatted = items.map((item: any) => ({
      label: item.FieldName,
      value: item.FieldType,
      require: item.SaveMandantory,
      internalName: item.InternalName || item.FieldName, // must store real internal name
    }));

    setFields(formatted);
    setFormValues({});
     
    const getfolderhaveapprovalornot = await sp.web.lists
      .getByTitle("DMSPreviewFormMaster").items.filter(`SiteName eq 'Location' and DocumentLibraryName eq 'Approval section'`).getAll();
     getfolderhaveapprovalornot.map((item:any)=>{
      if(item.IsApproval === true){
        isApproval=true;
        // alert("isApproval" + isApproval)
      }
     })
 
   const checkapprovaltype = await sp.web.lists
      .getByTitle("TemplateDocumentSubCategory").items.filter(`DocumentSubCategory eq '${selectedOption.value}'`).getAll();
      console.log("checkapprovaltype",checkapprovaltype);
// 
      const approvalhierachy = await sp.web.lists
      .getByTitle("DMSFolderPermissionMaster").items.filter(`SiteName eq 'Location' and DocumentLibraryName eq 'Approval section'`).getAll();
      console.log("approvalhierachy",approvalhierachy);
      

          const testidsub = await sp.site.openWebById(siteID);
              // const user0 = await siteContext.web.siteUsers();
        const approvalGroupUsers= await testidsub.web.siteGroups.getByName(`Location_Approval`).users();
        // console.log("approvalGroupUsers",approvalGroupUsers);
        const finalArray=approvalGroupUsers.map((user)=>(
                {
                    userId:user.Id,
                    value: user.Title,
                    label: user.Title,
                    email: user.Email,
                }
        ));
        console.log("finalArray",finalArray);
        setUsers(finalArray);
      checkapprovaltype.map((item:any)=>{
        if(item.WorkflowRouting === "User can modify"){
                 setCheckapprovaltype(true);
              
                  // alert("User can modify Approval")

                 
        }else{
         setCheckapprovaltype(false);
        }
      })
  };

  // handle input change for dynamic fields
  const handleFieldChange = (field: any, value: any) => {
    setFormValues((prev: any) => ({
      ...prev,
      [field.internalName]: value, // store values by internal name
    }));
  };

  // handle file selection
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const newFiles = Array.from(e.target.files);
      setFiles((prev) => [...prev, ...newFiles]);
    }
  };

  // remove file
  const removeFile = (index: number) => {
    setFiles((prev) => prev.filter((_, i) => i !== index));
  };

  // render dynamic field input
  const renderInput = (field: any) => {
    switch (field.value) {
      case "Single Line Text":
        return (
          <input
            type="text"
            className="edc-input"
            required={field.require}
            onChange={(e) => handleFieldChange(field, e.target.value)}
          />
        );
      case "Multiple Line Text":
        return (
          <textarea
            className="edc-input"
            required={field.require}
            onChange={(e) => handleFieldChange(field, e.target.value)}
          ></textarea>
        );
      case "Number":
        return (
          <input
            type="number"
            className="edc-input"
            required={field.require}
            onChange={(e) => handleFieldChange(field, e.target.value)}
          />
        );
      case "Date":
        return (
          <input
            type="date"
            className="edc-input"
            required={field.require}
            onChange={(e) => handleFieldChange(field, e.target.value)}
          />
        );
      case "YesNo":
        return (
          <select
            className="edc-input"
            required={field.require}
            onChange={(e) => handleFieldChange(field, e.target.value)}
          >
            <option value="">Select</option>
            <option value="Yes">Yes</option>
            <option value="No">No</option>
          </select>
        );
      default:
        return (
          <input
            type="text"
            className="edc-input"
            required={field.require}
            onChange={(e) => handleFieldChange(field, e.target.value)}
          />
        );
    }
  };

  const getUniqueRequestNo = async () => {
  const counterItem = await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1)();
  console.log("Counter Item 0", counterItem);
  console.log("Counter Item 1", counterItem.FileCount);
  let fileCounter = counterItem.FileCount;

  // Increment the counter
  fileCounter++;

  // Generate the new RequestNo
  const newRequestNo = `File${String(fileCounter).padStart(2, '0')}`;

  // Update the counter in the CounterList
  await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1).update({
    FileCount: fileCounter
  });

  return newRequestNo;
};
  // submit form → upload file(s) + save JSON metadata
  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const testidsub = await sp.site.openWebById(siteID);

    try {
for (const file of files) {
  // 1. upload file to Project Document library
  let uploadedFile = await testidsub.web
    .getFolderByServerRelativePath(
      "/sites/Intranetdemos/Location/Section"
    )
    .files.addUsingPath(file.name, file, { Overwrite: true });

  console.log("Uploaded File:", uploadedFile);

  // 2. get the list item behind the file
  const item = await uploadedFile.file.getItem();
  //  2.2 get file preview url
  const listItem = await uploadedFile.file.getItem();
      console.log("ListItems ", listItem);

      const parentFolder = uploadedFile.data.ServerRelativeUrl.substring(0, uploadedFile.data.ServerRelativeUrl.lastIndexOf('/'));
      const siteUrl = window.location.origin;
      const encodeSharePointURL = (url: string) => {
        return encodeURIComponent(url)
            .replace(/'/g, "%27")
            .replace(/-/g, "%2D")
            .replace(/_/g, "%5F")
            .replace(/\./g, "%2E")
            .replace(/!/g, "%21")
            .replace(/\*/g, "%2A")
            .replace(/\(/g, "%28")
            .replace(/\)/g, "%29")
            .replace(/~/g, "%7E")
            .replace(/@/g, "%40")
            .replace(/\$/g, "%24")
            .replace(/,/g, "%2C")
            .replace(/;/g, "%3B")
            .replace(/:/g, "%3A")
            .replace(/\+/g, "%2B")
            .replace(/=/g, "%3D")
            .replace(/\?/g, "%3F")
            .replace(/\//g, "%2F")
            .replace(/#/g, "%23")
            .replace(/&/g, "%26");
    };
    
    // Usage
    const encodedFilePath = encodeSharePointURL(uploadedFile.data.ServerRelativeUrl);
    // const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeSharePointURL(parentFolder)}`;
    const previewUrl = `https://officeindia.sharepoint.com/sites/Intranetdemos/Location/Section/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeSharePointURL(parentFolder)}`;
    

    
    console.log("Preview URL:", previewUrl);
  // 3. update list item metadata directly with formValues
  // Generate unique DocumentNo
const categoryPrefix = (props.selectedCategory || "").substring(0, 2).toUpperCase();
const templatePrefix = (selectedSubCategory?.value || "").substring(0, 2).toUpperCase();

// take last 4 digits of timestamp for uniqueness
const timestampSuffix = Date.now().toString().slice(-4);

const documentNo = `${categoryPrefix}${templatePrefix}${timestampSuffix}`;

  await item.update({
    ...formValues, // your dynamic fields
    Status: "Pending", // fixed value
    DocumentCategory: props.selectedCategory, // from props
    Template: selectedSubCategory.value, // from dropdown
    DocumentNo:documentNo 
  });
      const newRequestNo = await getUniqueRequestNo();
   const newItem = await sp.web.lists.getByTitle(`DMSLocationFileMaster`).items.add({
          FileName: String(uploadedFile.data.Name),
          FileSize: String(uploadedFile.data.Length),
          FileVersion: String(uploadedFile.data.MajorVersion),
          CurrentFolderPath: String("/sites/Intranetdemos/Location/Section"),
          FileUID: String(uploadedFile.data.UniqueId),
          CurrentUser: String(props.currentuseremail),
          SiteID: String(siteID),
          Status: String('Pending'),
          FilePreviewURL: String(previewUrl),
          DocumentLibraryName: String('Section'),
          SiteName: String('Location'),
          MyRequest: true,
          Processname: 'New File Request',
          RequestNo: newRequestNo,
       
      });

  // Move DMSFileApprovalList logic inside the loop so uploadedFile is in scope
  if (isApproval === true) {
     const newRequestNo = await getUniqueRequestNo();
     const uploadfileapproval=   await sp.web.lists.getByTitle('DMSFileApprovalList').items.add({
      SiteName: 'Location',
      DocumentLibraryName: 'Section',
      RequestedBy: props.currentuseremail,
      FileName: String(uploadedFile.data.Name),
      FileUID: String(uploadedFile.data.UniqueId),
      FilePreviewUrl: String(previewUrl), // Set this to the correct preview URL if available
      Status: String('Pending'),
      FolderPath:  "/sites/Intranetdemos/Location/Section", // Set this to the correct folder path if available
      ApproveAction: String('Submitted'),
      ApprovedLevel: 1,
      RequestNo: newRequestNo,
      Processname: 'New File Request',
      CurrentLevel: 1
    });
      if (uploadfileapproval) {
        //  alert("File approval uploaded and saved successfully!");
      }
  }
}

      Swal.fire({
        icon: "success",
        title: "File(s) uploaded and saved successfully!",
        showConfirmButton: false,
        timer: 1500
      })
      setFiles([]);
      setFormValues({});
    } catch (err) {
      console.error("Error uploading:", err);
      alert("Error submitting form. Check console.");
    }

    location.reload();
              props.onReturnToMain();
  };
  const handleSelectionModeChange = (id: number, type: "All" | "One") => {
    const newRows = rows.map((row) =>
      row.id === id ? { ...row, selectionType: type } : row
    );
    setRows(newRows);
  };

  //   remove new row
    const handleRemoveRow = (
      id: number,
      event: React.MouseEvent<HTMLAnchorElement, MouseEvent>
    ) => {
      event.preventDefault();
      setRows(rows.filter((row) => row.id !== id));
    };

      const handleAddRow = (
        event: React.MouseEvent<HTMLAnchorElement, MouseEvent>
      ) => {
        event.preventDefault();
        const newId = rows.length ? rows[rows.length - 1].id + 1 : 0;
        // setRows([...rows, { id: newId, approvedUser: "", searchTerm: "", filteredUsers: [] }]);
    
        // start
        setRows([
          ...rows,
          { id: newId, selectionType: "One", approvedUserList: [] },
        ]);
        //end
      };

        const handleUserSelect = (selected: any, id: any) => {
    console.log(selectedUsers, "selectedUsers");
    console.log(selectedUsers, "selectedUsers");
    setSelectedUsers(selected || []);
    console.log(selected, "selected ");
    const newRows = rows.map((row) =>
      row.id === id ? { ...row, approvedUserList: selected } : row
    );
    console.log("Selected items", selected, id);
    // console.log(rows.length);
    setRows(newRows);
  };
    
  return (
    <div>
    <div className="card">

      <div className="card-body">

      <form onSubmit={handleSubmit}>
        {/* Category Select */}
        <div className="row">
          <div className="col-sm-4">
             <label className="maincategory">
                  Document Category 
                  <span className="required">*</span>
                </label>
            <Select
          options={[
            { label: props.selectedCategory, value: props.selectedCategory },
          ]}
          value={{
            label: props.selectedCategory,
            value: props.selectedCategory,
          }}
          isDisabled
          placeholder="Selected Category"
        />

          </div>
      
          <div className="col-sm-4">
            <label className="subcategory">
                  Document Sub-Category 
                  <span className="required">*</span>
                </label>
        {/* SubCategory Select */}
        <Select
          options={subCategories}
          placeholder="Select Document Sub Category"
          isClearable
          value={selectedSubCategory}
          onChange={handleSubCategorySelect}
        />

          </div>
           <div className="col-sm-4">
            {/* File Upload */}
        <div className="file-upload">
          <label>Upload Files</label>
          <input type="file" multiple onChange={handleFileChange} />
          <div className="file-preview">
            {files.map((file, idx) => (
              <div key={idx} className="file-item">
                <span
                  className="file-name"
                  onClick={() => window.open(URL.createObjectURL(file))}
                >
                  {file.name}
                </span>
                <button type="button" onClick={() => removeFile(idx)}>
                  ❌
                </button>
              </div>
            ))}
          </div>
        </div>

          </div>
</div>
          
             {/* Dynamic Fields */}
        {fields.length > 0 && (
          <div className="edc-grid row mt-2">
            {fields.map((field, idx) => (
              <div  key={idx} className="edc-field mb-3 col-sm-4">
                <label className="edc-label">
                  {field.label}{" "}
                  {field.require && <span className="required">*</span>}
                </label>
                {renderInput(field)}
              </div>
            ))}
          </div>
        )}

          
         
         
        


       

        

{/* approval hierarchy section */}
           {/* {checkapprovaltype ? (
        <div className="card cardborder p-31 mt-3">
          <div className="" style={{
            
        }}>
          
            <div className="row">
              <div className="col-sm-10 w90">
              <h3 className="header-title text-dark font-16 mb-1">Approval Hierarchy</h3>
     
     <p className="subheader font-14 mb-3">
       Define approval hierarchy for the documents submitted by Team
       members in this folder.
     </p>

              </div>

              <div className="col-sm-2">
              <div style={{height:'0px', position:'relative'}} className="mb-0">
              <div className="col-12 d-flex justify-content-end">
                <a onClick={handleAddRow}>
                  <img
                    className="bi bi-plus"
                    src={require("../assets/plus.png")}
                    alt="add"
                    style={{ width: "50px", top:'0px', position:'absolute', right:'0px', left:'auto', height: "50px" }}
                  />
                </a>
              </div>
            </div>

              </div>
            </div>
            

           
            <div style={{clear:'both'}} className="row mb-2 approvalheirarcystyle">
            <table className="mtbalenew mtbalenewn createc">
    <thead>
      <tr>
        <th> Level</th>
        <th> Approver</th>
        <th style={{textAlign:'center', minWidth:'70px', maxWidth:'70px'}}> Select</th>
        <th style={{minWidth:'40px',maxWidth:'40px', textAlign:'center'}}> Action</th>
      </tr>
    </thead>
    <tbody>


             
               {rows.map((row) => (
              <tr className="approvalheirarchyfield" key={row.id}>
                <td>
                  <input style={{height:'36px'}}
                    type="text"
                    className="form-control"
                    id={`level-${row.id}`}
                    value={`Level ${row.id + 1}`}
                    disabled
                  />
                </td>
                <td>
             
                  <Select
                    isMulti
                    options={users}
                  
                    onChange={(selected: any) =>
                      handleUserSelect(selected, row.id)
                    }
                    placeholder="Enter names or email addresses..."
                    noOptionsMessage={() => "No User Found..."}
                  />
                  
                  {errorsForUserSelection[row.id]?.userSelect && (
                      <span className="text-danger">{errorsForUserSelection[row.id].userSelect}</span>
                  )}
                 
                </td>
          
                <td style={{textAlign:'center', minWidth:'70px', maxWidth:'70px'}}>
                <div style={{gap:'10px', justifyContent:'center'}} className="d-flex">
                  <div className="form-check">
                    <input
                      className="form-check-input"
                      type="radio"
                      name={`selection-${row.id}`}
                      id={`all-${row.id}`}
                      value="all"
                      checked={row.selectionType === "All"}
                      onChange={() => handleSelectionModeChange(row.id, "All")}
                    />
                    <label
                      className="form-check-label"
                      htmlFor={`all-${row.id}`}
                    >
                      All
                    </label>
                  </div>
                  <div className="form-check">
                    <input
                      className="form-check-input"
                      type="radio"
                      name={`selection-${row.id}`}
                      id={`one-${row.id}`}
                      value="one"
                      checked={row.selectionType === "One"}
                      onChange={() => handleSelectionModeChange(row.id, "One")}
                    />
                    <label
                      className="form-check-label"
                      htmlFor={`one-${row.id}`}
                    >
                      One
                    </label>
                  </div>
                </div>
                </td>
         
                <td style={{minWidth:'40px',maxWidth:'40px', textAlign:'center'}}>
                {row.id === 0 ? (
                 null
                ) : (
                  <div style={{justifyContent:'center'}} className="d-flex align-items-end">
                    <a
                      onClick={(e) => handleRemoveRow(row.id, e)}
                      style={{
                        width: "50px",
                       
                        cursor: "pointer",
                      }}
                    >
                      <img
                        className="fas fa-trash"
                        src={require("../assets/del.png")}
                        alt="delete"
                      />
                    </a>
                  </div>
                )}
                </td>

               
              </tr>
            ))}
            </tbody>
            </table>
            </div>
           

          </div>
        </div>
      ) : null
      } */}

      
        <button type="submit" className="submit-btn">
          Submit
        </button>
      </form>
              
      </div>
    </div>
    </div>
  );
};

export default DocumentTemplate;
