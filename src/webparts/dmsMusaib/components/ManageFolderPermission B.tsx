import React, { useEffect } from 'react'
import { useRef, useState } from "react";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import Select from "react-select";

interface ManageFolderPermissionProps {
    OthProps: { [key: string]: string };
    onReturnToMain: () => void;
  }

let IsUpdate:boolean;
let IsPrivateColumnId:number;

const ManageFolderPermission : React.FC<ManageFolderPermissionProps> = ({
    OthProps,
    onReturnToMain,
  }) =>{

    console.log(OthProps, "oth props");
    const sp: SPFI = getSP();

    const currentUserEmailRef = useRef('');
    const getcurrentuseremail = async()=>{
        const userdata = await sp.web.currentUser();
        currentUserEmailRef.current = userdata.Email;
        defaultValue();
    }

    // New Code
    const [rowsForPermission, setRowsForPermission] = React.useState<
    { id: number; selectedUserForPermission: string[]; selectedPermission:"" }[]
    >([{ id: 0, selectedUserForPermission: [],selectedPermission:"" }]);
    console.log("rowsForPermission",rowsForPermission);
    // End

    const [users, setUsers] = React.useState<any[]>([]);
    console.log("Users Array", users);

    const [toggelPermission,setTogglePermission]=React.useState<string>();
    console.log("toggelPermission",toggelPermission);

    const handlesetTogglePermission=()=>{
        setTogglePermission("Yes");
    }

    // const [selectedPermission,setSelectedPermission]=useState([]);
    // const [defaultUser,setDefaultUser]=useState<{
    //     userId:number,
    //     value: String,
    //     label: String,
    //   }[]>([]);

    // console.log("Default user array",defaultUser)
  
    const permissionArray:any[]=[
      {value:"Admin",label:"Admin"},
      {value:"Contribute",label:"Contribute"},
      {value:"Edit",label:"Edit"},
      {value:"Read",label:"Read"},
      {value:"View",label:"View"}
    ];
  
    const handleUserSelectForPermission=(selectedUser:any,Id:any)=>{
        console.log("Selected user for permission",selectedUser);
        const newRows = rowsForPermission.map((row) =>
          row.id === Id ? { ...row, selectedUserForPermission: selectedUser } : row
        );
        setRowsForPermission(newRows);
    }
  
    const handlePermissionSelect=(selectedPermission:any,Id:any)=>{
        console.log("Selected Permission",selectedPermission)
        const newRows = rowsForPermission.map((row) =>
          row.id === Id ? { ...row, selectedPermission: selectedPermission } : row
        );
        setRowsForPermission(newRows);

    }

    const handleAddRow = (
      event: React.MouseEvent<HTMLAnchorElement, MouseEvent>
    ) => {
      event.preventDefault();
      const newId = rowsForPermission.length ? rowsForPermission[rowsForPermission.length - 1].id + 1 : 0;
      setRowsForPermission([
        ...rowsForPermission,
        { id: newId, selectedUserForPermission: [], selectedPermission:"" },
      ]);
    };

    const handleRemoveRow = (
      id: number,
      event: React.MouseEvent<HTMLAnchorElement, MouseEvent>
    ) => {
      event.preventDefault();
      setRowsForPermission(rowsForPermission.filter((row) => row.id !== id));
  };

    // Fetch users from SharePoint
  React.useEffect(() => {
    getcurrentuseremail();
    console.log(currentUserEmailRef.current ,"my current id")
    const fetchUsers = async () => {
      try {
        // start
        const siteContext = await sp.site.openWebById(OthProps.SiteID);
        const user0 = await siteContext.web.siteUsers();

        const combineUsersArray=user0.map((user)=>(
              {
              userId:user.Id,
              value: user.Title,
              label: user.Title,
              email: user.Email,
          }
        ))
        setUsers(combineUsersArray);
        console.log("Sub site users",combineUsersArray);
        // const user0 = await sp.web.siteUsers();
        // const [
        //   users,
        //   users1,
        //   users2,
        //   users3,
        //   users4,
        // ] = await Promise.all([
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Read`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Initiator`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Contribute`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Admin`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_View`).users(),
        // ]);
        // console.log(users, "users ", users1,users2,users3,users4);
        // const combineArray = [
        //   ...(users || []),
        //   ...(users1 || []),
        //   ...(users2 || []),
        //   ...(users3 || []),
        //   ...(users4 || []),
        // ];
        // setUsers(
        //   combineArray.map((user) => ( 
        //   {
        //     userId:user.Id,
        //     value: user.Title,
        //     label: user.Title,
        //     email: user.Email,
        //   }
        // ))
        // );
        // console.log("combineArray", combineArray);
        // end
      } catch (error) {
        console.error("Error fetching site users:", error);
      }
    };

    fetchUsers();
  }, []);

  // interface User {
  //   userId: number;
  //   value: string;
  //   label: string;
  //   Permission: string;
  // }
  
  // interface GroupedPermissions {
  //   id: number;
  //   selectedUserForPermission: User[];  
  //   selectedPermission: string;
  // }
// fetch the initial data from the  DMSFolderPrivacy list
    const defaultValue=async()=>{
        try {

            // fetch the IsPrivate from DMSPreviewFormMaster to chcek the folder is private or not
            const permissionDetails=await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("IsPrivate","Id").filter(`SiteName eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName}' and IsDocumentLibrary eq 1`)();
            console.log("approverDetails",permissionDetails);
            IsUpdate=permissionDetails[0].IsPrivate;
            IsPrivateColumnId=permissionDetails[0].Id;

            
            if(!IsUpdate){
              setTogglePermission("No");
            }else{
              setTogglePermission("Yes");
            }
            const fetchData=await sp.web.lists.getByTitle("DMSFolderPrivacy").items.select("User","UserID","UserPermission").filter(`SiteName eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName}' and CurrentUser eq '${currentUserEmailRef.current}'`)();
            console.log("Fetch data",fetchData);    
            

            // Initialize array to store the default users  
            const arrayToStoreDefaultUser = fetchData.map((user) => ({
                userId: user.UserID,
                value: user.User,
                label: user.User,
                Permission:user.UserPermission
            }));

            // filter the data.
            const filteredData = arrayToStoreDefaultUser.filter(item => item.userId !== null && item.value !== null && item.label !== null);

            // const permission=filteredData[0].Permission;
            // const objectPermission={value:permission,label:permission};
            // console.log("Permission Object",objectPermission);
            // setSelectedPermission([objectPermission]);
            console.log("filteredData",filteredData);
            // New Code
            // Helper function to generate a random ID
            const generateRandomId = (): number => Math.floor(Math.random() * 100000);
            const grouped: { [key: string]: any } = {};
            filteredData.forEach((user)=>{
                  const { Permission } = user;
                  // If the permission group doesn't exist, create it with a random id
                  if (!grouped[Permission]) {
                      grouped[Permission] = {
                      id: generateRandomId(),
                      selectedUserForPermission: [],
                      selectedPermission: {value:Permission,label:Permission}
                  };
                }

                // Add the user to the correct permission group
                grouped[Permission].selectedUserForPermission.push(user);
            })
            console.log("grouped",grouped)
            const permissionsArray = Object.keys(grouped).map(key => grouped[key]);
            console.log("permissionsArray",permissionsArray);
            setRowsForPermission(permissionsArray)
            // End
            // setDefaultUser(filteredData);
            console.log("Fetch data from DMSFolderPrivacy",fetchData);

        } catch (error) {
            console.log("Erroe fetching data from DMSFolderPrivacy",error);
        }
    }
    

  const handleCreate=async()=>{
    console.log("create called");
    console.log("rowsForPermission",rowsForPermission);
    // console.log("selected User Array",defaultUser);
    // console.log("selected permission",selectedPermission);

    try {
          // update the column IsPrivate if its false and add the data in folder prvivay without the user.
          if(!IsUpdate){
            // Update list item by ID
            try {
                await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.getById(IsPrivateColumnId).update({
                    IsPrivate:true
                });
                console.log("Item Updated in DMSPreviewFormMaster");
    
                const payloadForDMSFolderPrivacyWithoutUser={
                    SiteName:OthProps.SiteTitle,
                    DocumentLibraryName:OthProps.DocumentLibraryName,
                    CurrentUser:currentUserEmailRef.current,
                }
                const addedItem = await sp.web.lists.getByTitle("DMSFolderPrivacy").items.add(payloadForDMSFolderPrivacyWithoutUser);
                console.log("Item added to the list DMSFolderPrivacy after selected value without users ",addedItem);
            } catch (error) {
                console.log("error in updating data in DMSPreviewFormMaster or creating item in DMSFolderPrivacy without user",error);
            }     
        }

        const fetchData=await sp.web.lists.getByTitle("DMSFolderPrivacy").items.select("User","UserID","UserPermission","ID").filter(`SiteName eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName}' and CurrentUser eq '${currentUserEmailRef.current}'`)();

        console.log("fetched user for delete and create new user",fetchData);
        fetchData.forEach(async(user)=>{
            try {
                console.log("User->",user.User)
                if(user.User !== null){
                    // console.log("Skip");
                    await sp.web.lists.getByTitle("DMSFolderPrivacy").items.getById(user.ID).delete();
                    console.log(`Item with ID: ${user.ID} has been deleted successfully.`);
                }
                
            } catch (error) {
                console.log("Error deleting item from DMSFolderPrivacy",error);
            }
        })


        const payloadForDMSFolderPrivacy={
            SiteName:OthProps.SiteTitle,
            DocumentLibraryName:OthProps.DocumentLibraryName,
            CurrentUser:currentUserEmailRef.current,
            IsModified:true,
            // UserPermission:selectedPermission[0].value
        }
        rowsForPermission.forEach((row:any)=>{
            try {
                
                // (payloadForDMSFolderPrivacy as any).User=user.value;
                // (payloadForDMSFolderPrivacy as any).UserID=user.userId;
                (payloadForDMSFolderPrivacy as any).UserPermission=row.selectedPermission.value;
                row.selectedUserForPermission.forEach(async(user:any)=>{
                  (payloadForDMSFolderPrivacy as any).User=user.value;
                  (payloadForDMSFolderPrivacy as any).UserID=user.userId;
                  console.log("Payload for DMSFolderPrivacy",payloadForDMSFolderPrivacy);
                  const addedItem = await sp.web.lists.getByTitle("DMSFolderPrivacy").items.add(payloadForDMSFolderPrivacy);
                  console.log("Item added to the list DMSFolderPrivacy after selected value ",addedItem);
                })
                
            } catch (error) {
                console.log("error adding data to the DMSFolderPrivacy",error);
            }
            
        })
     

    } catch (error) {
        console.log("error creating data insode the handle create function",error);
    }

  }


  return (
    <div className="container mt-4 second">
            <div className="modal show d-block" tabIndex={-1}>
                    <div className="modal-dialog">
                        <div className="modal-content" style={{
                                    width:"140%",
                                    padding:"0px"
                                    
                        }}>


                    {toggelPermission === "Yes" ? 
                    (
                      <div>           
                            <div className="card cardborder marginleftcard" style={{
                              height: "auto",
                              width:"auto",
                          }}
                          >
                          <h5 className="mb-3 " style={{
                              display:"block",
                              width:"200px"
                          }}>
                              <strong>Manage Permission</strong>
                          </h5>
                          <div className="mb-3">
                            <div className="col-12 d-flex justify-content-end">
                              <a onClick={handleAddRow}>
                                <img className="bi bi-plus" src={require("../assets/plus.png")} alt="add" style={{ width: "50px", height: "50px" }} />
                              </a>
                            </div>
                          </div>
                          {rowsForPermission.map((row)=>(
                          <div className="row mb-3 approvalheirarcystyle" key={row.id}>
                                  <div className="col-12 col-md-6">
                                      <Select
                                          value={row.selectedUserForPermission}
                                          isMulti
                                          options={users}
                                          onChange={(selected: any) =>
                                          handleUserSelectForPermission(selected,row.id)
                                          }
                                          placeholder="Enter names or email addresses..."
                                          noOptionsMessage={() => "No User Found..."}
                                      />
                                  </div>
                                  <div className="col-12 col-md-6" style={{
                                  width:"auto"
                                  }}>
                                      <Select
                                          value={row.selectedPermission || null}
                                          options={permissionArray}
                                          onChange={(selected: any) =>
                                          handlePermissionSelect(selected,row.id)
                                          }
                                          placeholder="Select Permission..."
                                          noOptionsMessage={() => "No Such Permission Find"}
                                      />
                                  </div>
                                  {/* {row.id === 0 ? null : ( */}
                                    <div className="col-12 col-md-2 d-flex align-items-end">
                                      <a onClick={(e) => handleRemoveRow(row.id, e)} style={{ width: "50px",    height: "50px", cursor: "pointer" }}>
                                        <img className="fas" src={require("../assets/delete.png")} alt="delete" />
                                      </a>
                                    </div>
                                  {/* )} */}
                          </div>
                          ))}
                          </div>
                        
                        {/* <div className="card cardborder marginleftcard" style={{
                                height: "auto",
                                width:"auto",
                            }}>
                            <h5 className="mb-3 " style={{
                                display:"block",
                                width:"200px"
                            }}>
                                <strong>Manage Permission</strong>
                            </h5>
                            <div className="row mb-3 approvalheirarcystyle">
                                    <div className="col-12 col-md-6">
                                        <Select
                                            value={defaultUser}
                                            isMulti
                                            options={users}
                                            onChange={(selected: any) =>
                                            handleUserSelectForPermission(selected)
                                            }
                                            placeholder="Enter names or email addresses..."
                                            noOptionsMessage={() => "No User Found..."}
                                        />
                                    </div>
                                    <div className="col-12 col-md-6" style={{
                                    width:"auto"
                                    }}>
                                        <Select
                                            value={selectedPermission}
                                            options={permissionArray}
                                            onChange={(selected: any) =>
                                            handlePermissionSelect(selected)
                                            }
                                            placeholder="Select Permission"
                                            noOptionsMessage={() => "No Such Permission Find"}
                                        />
                                    </div>

                            </div>
                            </div> */}
                            <div className="modal-footer">
                                    <button type="button" className="btn btn-primary" 
                                    onClick={handleCreate}
                                    >
                               <span className="mb-1 mt-2" data-tooltip="Submit">
                               <img  src={require("../assets/submit-new.png")}/></span>
                                    </button>
                                    <button type="button" className="btn btn-secondary" 
                                    //   onClick={toggleModal}
                                    onClick={onReturnToMain}
                                    >
                                        <span className="mb-1 mt-2" data-tooltip="Cancel">
                                        <img  src={require("../assets/cancelnew.png")}/></span>
                                    </button>
                            </div>
                      </div>
                    )
                    :null
                    }             
                    {toggelPermission === "No"  ?
                    (
                    <div>
                       <h6>"This document library is public. Would you like to Add More Permission?"</h6>
                        <div className='setApprover' style={{
                          display:"flex",
                          marginLeft:"280px",
                          gap:"10px"
                        }}>
                        <button type="button" className="btn btn-primary" onClick={handlesetTogglePermission}>Set Permission</button>
                        <button type="button" className="btn btn-secondary" 
                          onClick={onReturnToMain}
                        >
                          Cancel{" "}
                        </button>
                      </div>
                  </div>
                )
                : null
                }
                        </div>
                    </div>
            </div>
</div>
  )
}

export default ManageFolderPermission