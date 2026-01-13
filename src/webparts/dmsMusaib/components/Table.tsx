// // @ts-ignore
// import * as React from "react";
// import { getSP } from "../loc/pnpjsConfig";
// import { SPFI } from "@pnp/sp";
// // import { faEdit, faTrashAlt } from '@fortawesome/free-regular-svg-icons';
// // import { IMediaMasterProps } from './IMediaMasterProps';
// import "bootstrap/dist/css/bootstrap.min.css";
// // import "bootstrap//dist/"
// import "../../../CustomCss/mainCustom.scss";
// // import "../../verticalSideBar/components/VerticalSidebar.scss";
// // import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
// import UserContext from "../../../GlobalContext/context";
// import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
// import {  faEllipsisV, faFileExport, faSort , faExclamation , faListSquares
// } from '@fortawesome/free-solid-svg-icons';
// // import { useState , useEffect } from "react";
// // import Provider from "../../../GlobalContext/provider";

// import { useMediaQuery } from "react-responsive";
// import "@pnp/sp/webs";
// import "@pnp/sp/folders";
// import "@pnp/sp/files";
// import "@pnp/sp/sites"
// import "@pnp/sp/presets/all"
// import "bootstrap/dist/css/bootstrap.min.css";
// import "../../../CustomCss/mainCustom.scss";

// // import "../../verticalSideBar/components/VerticalSidebar.scss";
// import "./dmscss";
// import { useState , useRef , useEffect} from "react";
// import "./MediaMaster.module.scss"
// import "./mediamaster.scss"
// // import "./CustomTable.scss"

// import Swal from 'sweetalert2';
// import * as XLSX from 'xlsx';
// import { escape } from "querystring";
// // import moment from 'moment';
// // import { title } from "process";
// let currentsiteID = ""
// interface CreateFolderProps {
//   Currentbuttonclick: { 
//     buttonclickis: string;
//     documentLibraryData?: any[]; // ✅ Add this optional property
//   };
//   onReturnToMain: () => void;
// }

// const Table: React.FC<CreateFolderProps> = ({Currentbuttonclick , onReturnToMain }) => {
//   console.log(Currentbuttonclick , "Currentbuttonclick")
//   const sp: SPFI = getSP();
//   console.log(sp, "sp");

//   const { useHide }: any = React.useContext(UserContext);
//   const elementRef = React.useRef<HTMLDivElement>(null);
//   const [isSidebarOpen, setIsSidebarOpen] = React.useState(true);

//   interface IListItem {
//     ID: number;
//     Title: string;
//     // Add other properties as needed from your list
//   }
//    const [mediaData, setmediaData] = useState<IListItem[]>([]); 
  
//   // const [mediaData, setmediaData] = React.useState([]);
//   const isMobile = useMediaQuery({ query: '(max-width: 768px)' });
//   const [filters, setFilters] = React.useState({
//     SNo: '',
//     Title : '',
//     FileName: '',
//     CurrentUser: '',
//     Modified: '',
//     Status: '',

//     SubmittedDate: ''
//   });
//   const [sortConfig, setSortConfig] = React.useState({ key: '', direction: 'ascending' });

//   React.useEffect(() => {
//     console.log("This function is called only once", useHide);

//     const showNavbar = (
//       toggleId: string,
//       navId: string,
//       bodyId: string,
//       headerId: string
//     ) => {
//       const toggle = document.getElementById(toggleId);
//       const nav = document.getElementById(navId);
//       const bodypd = document.getElementById(bodyId);
//       const headerpd = document.getElementById(headerId);

//       if (toggle && nav && bodypd && headerpd) {
//         toggle.addEventListener("click", () => {
//           nav.classList.toggle("show");
//           toggle.classList.toggle("bx-x");
//           bodypd.classList.toggle("body-pd");
//           headerpd.classList.toggle("body-pd");
//         });
//       }
//     };

//     showNavbar("header-toggle", "nav-bar", "body-pd", "header");

//     const linkColor = document.querySelectorAll(".nav_link");

//     function colorLink(this: HTMLElement) {
//       if (linkColor) {
//         linkColor.forEach((l) => l.classList.remove("active"));
//         this.classList.add("active");
//       }
//     }

//     linkColor.forEach((l) => l.addEventListener("click", colorLink));
//   }, [useHide]);
//   // Media query to check if the screen width is less than 768px


//   React.useEffect(() => {
//     console.log("This function is called only once", useHide);

//     const showNavbar = (
//       toggleId: string,
//       navId: string,
//       bodyId: string,
//       headerId: string
//     ) => {
//       const toggle = document.getElementById(toggleId);
//       const nav = document.getElementById(navId);
//       const bodypd = document.getElementById(bodyId);
//       const headerpd = document.getElementById(headerId);

//       if (toggle && nav && bodypd && headerpd) {
//         toggle.addEventListener("click", () => {
//           nav.classList.toggle("show");
//           toggle.classList.toggle("bx-x");
//           bodypd.classList.toggle("body-pd");
//           headerpd.classList.toggle("body-pd");
//         });
//       }
//     };

//     showNavbar("header-toggle", "nav-bar", "body-pd", "header");

//     const linkColor = document.querySelectorAll(".nav_link");

//     function colorLink(this: HTMLElement) {
//       if (linkColor) {
//         linkColor.forEach((l) => l.classList.remove("active"));
//         this.classList.add("active");
//       }
//     }

//     linkColor.forEach((l) => l.addEventListener("click", colorLink));
//   }, [useHide]);
//   React.useEffect(() => {
//     const handleEscape = (e: KeyboardEvent) => {
//       if (e.key === "Escape") {
//         if (document.fullscreenElement) {
//           document.exitFullscreen();
//         }
//       }
//     };

//     window.addEventListener("keydown", handleEscape);
//     return () => window.removeEventListener("keydown", handleEscape);
//   }, []);
// /////////////////// DMS Code start / ////////////////////////////////////
  
//   console.log("This is current side ID",currentsiteID)
//   const currentUserEmailRef = useRef('');
//   useEffect(() => {
//      getcurrentuseremail()
//      ApiCall()
// }, []);
//  const getcurrentuseremail = async()=>{
//   const userdata = await sp.web.currentUser();
//   currentUserEmailRef.current = userdata.Email;
//   console.log(currentUserEmailRef.current, "currentuser")
//  }

//  const ApiCall = async () => {
//   if (Currentbuttonclick.buttonclickis === 'Myrequest') {
//     console.log(Currentbuttonclick , "Currentbuttonclick")
//     console.log(typeof Currentbuttonclick , "typeof Currentbuttonclick")

//     // do something
//         // Fetch the list of active lists
//         const FilesItems = await sp.web.lists
//         .getByTitle("MasterSiteURL")
//         .items.select("Title", "SiteID", "FileMasterList", "Active")
//         .filter(`Active eq 'Yes'`)();
    
//       console.log("Active Sites List Names", FilesItems);
    
//       FilesItems.forEach(async (fileItem) => {
  
//           const filesData = await sp.web.lists
//             .getByTitle(`${fileItem.FileMasterList}`)
//             .items.select("ID" , "FileName", "FileUID", "FileSize", "FileVersion" ,"Status" , "SiteID" , "CurrentUser"
//               , "Modified"
//             )
//             .filter(
//               `CurrentUser eq '${currentUserEmailRef.current}' and MyRequest eq 1`
//             )();
    
//           console.log(`Files of Current user ${fileItem.FileMasterList}`, filesData);
//           setmediaData((prevMediaData) => [...prevMediaData, ...filesData]); 
//           // setmediaData(filesData)
//           console.log(filesData , "filesData")
//   });
//   } else if (Currentbuttonclick.buttonclickis === 'Favourite') {
//     console.log(Currentbuttonclick , "Currentbuttonclick")
//     console.log(typeof Currentbuttonclick , "typeof Currentbuttonclick")

//     // do something else
//         // Fetch the list of active lists
//         const FilesItems = await sp.web.lists
//         .getByTitle("MasterSiteURL")
//         .items.select("Title", "SiteID", "FileMasterList", "Active")
//         .filter(`Active eq 'Yes'`)();
    
//       console.log("Active Sites List Names", FilesItems);
    
//       FilesItems.forEach(async (fileItem) => {
  
//           const filesData = await sp.web.lists
//             .getByTitle(`${fileItem.FileMasterList}`)
//             .items.select("ID" , "FileName", "FileUID", "FileSize", "FileVersion" ,"Status" , "SiteID" , "CurrentUser"
//               , "Modified" , "IsFavourite"
//             )
//             .filter(
//               `IsFavourite eq 1 and CurrentUser eq '${currentUserEmailRef.current}'`
//             )();
    
//           console.log(`Files of Current user ${fileItem.FileMasterList}`, filesData);
//           setmediaData((prevMediaData) => [...prevMediaData, ...filesData]); 
//           // setmediaData(filesData)
//           console.log(filesData , "filesData")
//   });
//   }  else if (Currentbuttonclick.buttonclickis === 'DocumentLibrary') {
//     // ✅ Use the passed document library data
//     alert('Document Library Data Loaded');
//     const docLibData = Currentbuttonclick.documentLibraryData || [];
//     setmediaData(docLibData as any);
//   }


// console.log(mediaData , "statelistData")
//  }
// const headers = [
//   { label: 'S.No.', key: 'ID', style: { width: '5%' } },
//   { label: 'Title', key: 'Title', style: { width: '20%' } },
//   { label: 'Image', key: 'mediaandNewsBannerImage', type: 'image', style: { width: '10%' } },
//   { label: 'Description', key: 'Description', style: { width: '50%' } },
//   { label: 'Date', key: 'SubmittedDate', style: { width: '15%' } },
//   { label: 'Action', key: 'Action', style: { width: '15%' } },

// ];
// const handleFilterChange = (e: React.ChangeEvent<HTMLInputElement>, field: string) => {
//   setFilters({
//     ...filters,
//     [field]: e.target.value,
//   });
//   console.log(filters , "filters filters")
// };
// console.log(filters , "filters filters")
// const handleSortChange = (key: string) => {
//   let direction = 'ascending';
//   if (sortConfig.key === key && sortConfig.direction === 'ascending') {
//     direction = 'descending';
//   }
//   setSortConfig({ key, direction });
// };
// const applyFiltersAndSorting = (data: any[]) => {
//   debugger;
//   // Filter data
//   const filteredData = data.filter((item, index) => {
//     return (
//       (filters.Title === '' || item.FileName.toLowerCase().indexOf(filters.Title.toLowerCase()) !== -1) &&
//        (filters.Title === '' || item.FileName.toLowerCase().includes(filters.Title.toLowerCase())) &&
//        (filters.CurrentUser === '' || item.CurrentUser.toLowerCase().includes(filters.CurrentUser.toLowerCase())) &&
//       (filters.Status === '' || item.Modified.toLowerCase().includes(filters?.Status?.toLowerCase())) &&
//       (filters.SubmittedDate === '' || item.Status.toLowerCase().includes(filters.SubmittedDate.toLowerCase()))
//     );
//   });

//   // Natural sort function for alphanumeric values
//   const naturalSort = (a:any, b:any) => {
//     return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
//   };

//   // Sort data
//   const sortedData = filteredData.sort((a, b) => {
//     if (sortConfig.key === 'SNo') {
//       // Sort by index
//       const aIndex = data.indexOf(a);
//       const bIndex = data.indexOf(b);
//       return sortConfig.direction === 'ascending' ? aIndex - bIndex : bIndex - aIndex;
//     } else if (sortConfig.key) {
//       // Sort by other keys
//       const aValue = a[sortConfig.key] ? a[sortConfig.key].toLowerCase() : '';
//       const bValue = b[sortConfig.key] ? b[sortConfig.key].toLowerCase() : '';

//       return sortConfig.direction === 'ascending' ? naturalSort(aValue, bValue) : naturalSort(bValue, aValue);
//     }
//     return 0;
//   });

//   return sortedData;
// };

// const filteredAnnouncementData = applyFiltersAndSorting(mediaData);

// const [currentPage, setCurrentPage] = React.useState(1);
// const itemsPerPage = 10;
// const totalPages = Math.ceil(filteredAnnouncementData.length / itemsPerPage);

// const handlePageChange = (pageNumber: any) => {
//   if (pageNumber > 0 && pageNumber <= totalPages) {
//     setCurrentPage(pageNumber);
//   }
// };

// const startIndex = (currentPage - 1) * itemsPerPage;
// const endIndex = startIndex + itemsPerPage;
// const currentData = filteredAnnouncementData.slice(startIndex, endIndex);
// console.log(currentData , "currentData")

// //#region Download exl file 
// const handleExportClick = () => {
//   console.log(currentData,'currentData');
  
//   const exportData = currentData.map((item, index) => ({

//     'S.No.': startIndex + index + 1,
//     'FileName': item.FileName,
//     'SubmittedBy': item.CurrentUser,
//     'Modified': item.Modified,
//     'Status': item.Status,

//   }));

//   exportToExcel(exportData, 'MeadiaGallery');
// };
// const exportToExcel = (data: any[], fileName: string) => {
//   const workbook = XLSX.utils.book_new();
//   const worksheet = XLSX.utils.json_to_sheet(data);
//   XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//   XLSX.writeFile(workbook, `${fileName}.xlsx`);
// };

// const [isOpen, setIsOpen] = React.useState(false);
// const toggleDropdown = () => {
//   setIsOpen(!isOpen);
// };
// const Editmedia = (id: any) => {
//   debugger
//   //  setUseId(id)

//   // window.location.href = `${siteUrl}/SitePages/MediaGalleryForm.aspx`;
// }
// //#endregion

// //#region 
// const Deletemedia = (id: any) => {
//   Swal.fire({
//     title: "Are you sure?",
//     text: "You won't be able to revert this!",
//     icon: "warning",
//     showCancelButton: true,
//     confirmButtonColor: "#3085d6",
//     cancelButtonColor: "#d33",
//     confirmButtonText: "Yes, delete it!"
//   }).then((result) => {
//     if (result.isConfirmed) {
//       // const DeleteRes = DeletemediaAPI(sp, id)
//       ApiCall()
//       Swal.fire({
//         title: "Deleted!",
//         text: "Item has been deleted.",
//         icon: "success"
//       });

//     }
//   })
// }
// interface PaginationProps{
//   currentPage: number;
//   totalPages: any;
//   handlePageChange: any;
// }
// const Pagination = ( { currentPage, totalPages, handlePageChange }: PaginationProps) => {
//   const pageLimit = 5; // Number of visible page items

//   // Determine the start and end page based on the current page and total pages
//   const startPage = Math.max(1, currentPage - Math.floor(pageLimit / 2));
//   const endPage = Math.min(totalPages, startPage + pageLimit - 1);

//   // Adjust start page if it's too close to the end
//   const adjustedStartPage = Math.max(1, Math.min(startPage, totalPages - pageLimit + 1));

//   // Create an array for the visible page numbers
//   const visiblePages = [];
//   const limit = Math.min(pageLimit, totalPages);
//   for (let i = 0; i < limit; i++) {
//     visiblePages.push(adjustedStartPage + i);
//   }


//   return (
//     <nav className="pagination-container">
//       <ul className="pagination">
//         {/* Previous Button */}
//         <li className={`page-item ${currentPage === 1 ? 'disabled' : ''}`}>
//           <a
//             className="page-link PreviousPage"
//             onClick={() => handlePageChange(currentPage - 1)}
//             aria-label="Previous"
//           >
//             «
//           </a>
//         </li>

//         {/* Render visible page numbers */}
//      {visiblePages.map((pageNumber:any) => (
//           <li
//             key={pageNumber}
//             className={`page-item ${currentPage === pageNumber ? 'active' : ''}`}
//           >
//             <a className="page-link" onClick={() => handlePageChange(pageNumber)}>
//               {pageNumber}
//             </a>
//           </li>
//         ))} 

//         {/* Next Button */}
//         <li className={`page-item ${currentPage === totalPages ? 'disabled' : ''}`}>
//           <a
//             className="page-link NextPage"
//             onClick={() => handlePageChange(currentPage + 1)}
//             aria-label="Next"
//           >
//             »
//           </a>
//         </li>
//       </ul>
//     </nav>
//   );
// };



//   return (
// <>
// {/* <button onClick={onReturnToMain}>Back To Main Component</button> */}
// {/* table for my request and my favoruite */}
// {/* <table className="mtbalenew">
//                         <thead>
//                           <tr>
//                             <th  style={{ minWidth: '50px', maxWidth: '50px'}}>
 
//                               <div
//                                 >
//                                 <span>S.No.</span>
                               
//                               </div>
                          
//                             </th>
//                             <th ><div className=" ">
//                               <div className="" >
//                                 <span >Document Name</span>  
                               
//                                 </div>
                             
//                             </div></th>
//                             <th >
//                               <div >
//                                 <div  >
//                                   <span >	Submiited By</span>  
                                 
//                                 </div>
//                               </div>
//                             </th>
//                             <th  >
//                               <div>
//                                 <div >
//                                   <span >Modified date</span> 
                                  
//                                    </div>
                              
//                               </div>
//                             </th>
//                             <th   style={{ minWidth: '80px', maxWidth: '80px' }}>
//                               <div >
//                               <div >
//                                 <span >Status</span> 
                                
//                                 </div>
                             
//                             </div></th>
//                             <th  style={{ minWidth: '80px', maxWidth: '80px' }} >
//                               <div>
//                                 <div  >  <span >Size</span>  &nbsp;&nbsp;
//                                <div className="dropdown">
//                                   <FontAwesomeIcon  icon={faEllipsisV} onClick={toggleDropdown} size='sm'/>
//                                 </div> 
//                                 </div>
//                                 <div className=" bd-highlight">   <div style={{zIndex:'9'}} id="myDropdown" className={`dropdown-content ${isOpen ? 'show' : ''}`}>
//                                   <div onClick={handleExportClick} className="" >
//                                     <FontAwesomeIcon icon={faFileExport} />  Export
//                                   </div>
//                                 </div></div>
 
//                               </div>
                         
//                             </th>
//                           </tr>
//                         </thead>
//                         <tbody style={{ maxHeight: '5000px' }}>
//                           {currentData.length > 0 ? currentData.map((item, index) => {
//                             console.log(item , "item >>>>>>>>")
                           
//                             return (
 
//                               <tr key={index}>
//                                 <td style={{ minWidth: '50px', maxWidth: '50px' }}>
//                                   <span className="indexdesign"> {index + 1}</span>
//                                  </td>
//                                 <td>{item.FileName}</td>
//                                 <td>{item.CurrentUser}</td>
                               
//                                 <td>

//                                    {new Date(item?.Modified).toLocaleString('en-GB', {
//                                           day: '2-digit',
//                                           month: '2-digit',
//                                           year: 'numeric',
//                                           hour: '2-digit',
//                                           minute: '2-digit',
                                         
//                                           hour12: true
//                                         })}
//                                 </td>
//                                 <td style={{ minWidth: '80px', maxWidth: '80px' }}>
                                 
//                                   <div style={{background:'#d5efd6', color:'#5fc360', width:'auto', float:'none'}} className="filestatus">
//                                   {item.Status} </div></td>
//                                  <td style={{ minWidth: '80px', maxWidth: '80px' }} className="ng-binding">
                                  
//                                    {((item.FileSize as unknown as number) / (1024 * 1024)).toFixed(2)} MB
//                                 </td> 
//                               </tr>
//                             )
//                           }) : ""
//                           }
//                         </tbody>
//                       </table> */}
//                       {/* Table for My Request and My Favourite */}
//     {(Currentbuttonclick.buttonclickis === 'Myrequest' || Currentbuttonclick.buttonclickis === 'Favourite') && (
//       <table className="mtbalenew">
//         <thead>
//           <tr>
//             <th style={{ minWidth: '50px', maxWidth: '50px'}}>
//               <div><span>S.No.</span></div>
//             </th>
//             <th>
//               <div><span>Document Name</span></div>
//             </th>
//             <th>
//               <div><span>Submitted By</span></div>
//             </th>
//             <th>
//               <div><span>Modified date</span></div>
//             </th>
//             <th style={{ minWidth: '80px', maxWidth: '80px' }}>
//               <div><span>Status</span></div>
//             </th>
//             <th style={{ minWidth: '80px', maxWidth: '80px' }}>
//               <div>
//                 <span>Size</span> &nbsp;&nbsp;
//                 <div className="dropdown">
//                   <FontAwesomeIcon icon={faEllipsisV} onClick={toggleDropdown} size='sm'/>
//                 </div>
//               </div>
//               <div className="bd-highlight">
//                 <div style={{zIndex:'9'}} id="myDropdown" className={`dropdown-content ${isOpen ? 'show' : ''}`}>
//                   <div onClick={handleExportClick} className="">
//                     <FontAwesomeIcon icon={faFileExport} /> Export
//                   </div>
//                 </div>
//               </div>
//             </th>
//           </tr>
//         </thead>
//         <tbody style={{ maxHeight: '5000px' }}>
//           {currentData.length > 0 ? currentData.map((item, index) => (
//             <tr key={index}>
//               <td style={{ minWidth: '50px', maxWidth: '50px' }}>
//                 <span className="indexdesign">{startIndex + index + 1}</span>
//               </td>
//               <td>{item.FileName}</td>
//               <td>{item.CurrentUser}</td>
//               <td>
//                 {new Date(item?.Modified).toLocaleString('en-GB', {
//                   day: '2-digit',
//                   month: '2-digit',
//                   year: 'numeric',
//                   hour: '2-digit',
//                   minute: '2-digit',
//                   hour12: true
//                 })}
//               </td>
//               <td style={{ minWidth: '80px', maxWidth: '80px' }}>
//                 <div style={{background:'#d5efd6', color:'#5fc360', width:'auto', float:'none'}} className="filestatus">
//                   {item.Status}
//                 </div>
//               </td>
//               <td style={{ minWidth: '80px', maxWidth: '80px' }} className="ng-binding">
//                 {((item.FileSize as unknown as number) / (1024 * 1024)).toFixed(2)} MB
//               </td>
//             </tr>
//           )) : (
//             <tr>
//               <td colSpan={6} className="text-center">No results found</td>
//             </tr>
//           )}
//         </tbody>
//       </table>
//     )}

//     {/* Table for Document Library - Dynamic Columns */}
//     {Currentbuttonclick.buttonclickis === 'DocumentLibrary' && (
//       <table className="mtbalenew">
//         <thead>
//           <tr>
//             <th style={{ minWidth: '50px', maxWidth: '50px'}}>
//               <div><span>S.No.</span></div>
//             </th>
            
//             {/* ✅ Dynamic headers for Document Library */}
//             {mediaData.length > 0 && Object.keys(mediaData[0])
//               .filter(key => key !== 'ID')
//               .map((key) => (
//                 <th key={key} style={{ minWidth: '120px' }}>
//                   <div><span>{key}</span></div>
//                 </th>
//               ))}
//           </tr>
//         </thead>
//         <tbody>
//           {currentData.length > 0 ? currentData.map((item, index) => (
//             <tr key={index}>
//               <td style={{ minWidth: '50px', maxWidth: '50px' }}>
//                 <span className="indexdesign">{startIndex + index + 1}</span>
//               </td>
              
//               {/* ✅ Dynamic columns for Document Library */}
//               {Object.entries(item)
//                 .filter(([key]) => key !== 'ID')
//                 .map(([key, value], i) => (
//                   <td key={i} style={{ minWidth: '120px' }}>
//                     {value as string}
//                   </td>
//                 ))}
//             </tr>
//           )) : (
//             <tr>
//               <td colSpan={20} className="text-center">No results found</td>
//             </tr>
//           )}
//         </tbody>
//       </table>
//     )}
    
  
//                   <Pagination
//                   currentPage={currentPage}
//                   totalPages={totalPages}
//                   handlePageChange={handlePageChange}
  
//                   />
// </>
         
            
//   );
// };



// // const Table = () => {
// //   return (
// //     <Provider>
// //       <ArgPoc />
// //     </Provider>
// //   );
// // };

// export default Table;
// @ts-ignore
import * as React from "react";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import UserContext from "../../../GlobalContext/context";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faEllipsisV, faFileExport, faSort, faExclamation, faListSquares } from '@fortawesome/free-solid-svg-icons';
import { useMediaQuery } from "react-responsive";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/sites"
import "@pnp/sp/presets/all"
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "./dmscss";
import { useState, useRef, useEffect } from "react";
import "./MediaMaster.module.scss"
import "./mediamaster.scss"
import Swal from 'sweetalert2';
import * as XLSX from 'xlsx';

let currentsiteID = ""

interface CreateFolderProps {
  Currentbuttonclick: {
    buttonclickis: string;
    documentLibraryData?: any[];
  };
  onReturnToMain: () => void;
}

const Table: React.FC<CreateFolderProps> = ({ Currentbuttonclick, onReturnToMain }) => {
  console.log(Currentbuttonclick, "Currentbuttonclick")
  const sp: SPFI = getSP();
  console.log(sp, "sp");
  const { useHide }: any = React.useContext(UserContext);
  const elementRef = React.useRef(null);
  const [isSidebarOpen, setIsSidebarOpen] = React.useState(true);

  interface IListItem {
    ID: number;
    Title: string;
  }

  const [mediaData, setmediaData] = useState([]);
  const isMobile = useMediaQuery({ query: '(max-width: 768px)' });
  
  // Updated filters state - will be dynamic for Document Library
  const [filters, setFilters] = React.useState<{ [key: string]: string }>({
    SNo: '',
    FileName: '',
    CurrentUser: '',
    Modified: '',
    Status: '',
    FileSize: ''
  });

  const [sortConfig, setSortConfig] = React.useState({
    key: '',
    direction: 'ascending'
  });

  React.useEffect(() => {
    console.log("This function is called only once", useHide);
    const showNavbar = (
      toggleId: string,
      navId: string,
      bodyId: string,
      headerId: string
    ) => {
      const toggle = document.getElementById(toggleId);
      const nav = document.getElementById(navId);
      const bodypd = document.getElementById(bodyId);
      const headerpd = document.getElementById(headerId);
      if (toggle && nav && bodypd && headerpd) {
        toggle.addEventListener("click", () => {
          nav.classList.toggle("show");
          toggle.classList.toggle("bx-x");
          bodypd.classList.toggle("body-pd");
          headerpd.classList.toggle("body-pd");
        });
      }
    };

    showNavbar("header-toggle", "nav-bar", "body-pd", "header");
    const linkColor = document.querySelectorAll(".nav_link");
    function colorLink(this: HTMLElement) {
      if (linkColor) {
        linkColor.forEach((l) => l.classList.remove("active"));
        this.classList.add("active");
      }
    }
    linkColor.forEach((l) => l.addEventListener("click", colorLink));
  }, [useHide]);

  React.useEffect(() => {
    const handleEscape = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        if (document.fullscreenElement) {
          document.exitFullscreen();
        }
      }
    };
    window.addEventListener("keydown", handleEscape);
    return () => window.removeEventListener("keydown", handleEscape);
  }, []);

  const currentUserEmailRef = useRef('');

  useEffect(() => {
    getcurrentuseremail()
    ApiCall()
  }, []);

  const getcurrentuseremail = async () => {
    const userdata = await sp.web.currentUser();
    currentUserEmailRef.current = userdata.Email;
    console.log(currentUserEmailRef.current, "currentuser")
  }

  const ApiCall = async () => {
    if (Currentbuttonclick.buttonclickis === 'Myrequest') {
      const FilesItems = await sp.web.lists
        .getByTitle("MasterSiteURL")
        .items.select("Title", "SiteID", "FileMasterList", "Active")
        .filter(`Active eq 'Yes'`)();
      console.log("Active Sites List Names", FilesItems);
      
      FilesItems.forEach(async (fileItem) => {
        const filesData = await sp.web.lists
          .getByTitle(`${fileItem.FileMasterList}`)
          .items.select("ID", "FileName", "FileUID", "FileSize", "FileVersion", "Status", "SiteID", "CurrentUser", "Modified")
          .filter(`CurrentUser eq '${currentUserEmailRef.current}' and MyRequest eq 1`)();
        console.log(`Files of Current user ${fileItem.FileMasterList}`, filesData);
        setmediaData((prevMediaData) => [...prevMediaData, ...filesData]);
      });
    } else if (Currentbuttonclick.buttonclickis === 'Favourite') {
      const FilesItems = await sp.web.lists
        .getByTitle("MasterSiteURL")
        .items.select("Title", "SiteID", "FileMasterList", "Active")
        .filter(`Active eq 'Yes'`)();
      console.log("Active Sites List Names", FilesItems);
      
      FilesItems.forEach(async (fileItem) => {
        const filesData = await sp.web.lists
          .getByTitle(`${fileItem.FileMasterList}`)
          .items.select("ID", "FileName", "FileUID", "FileSize", "FileVersion", "Status", "SiteID", "CurrentUser", "Modified", "IsFavourite")
          .filter(`IsFavourite eq 1 and CurrentUser eq '${currentUserEmailRef.current}'`)();
        console.log(`Files of Current user ${fileItem.FileMasterList}`, filesData);
        setmediaData((prevMediaData) => [...prevMediaData, ...filesData]);
      });
    } else if (Currentbuttonclick.buttonclickis === 'DocumentLibrary') {
      const docLibData = Currentbuttonclick.documentLibraryData || [];
      setmediaData(docLibData as any);
      
      // Initialize filters for Document Library columns
      if (docLibData.length > 0) {
        const dynamicFilters: { [key: string]: string } = { SNo: '' };
        Object.keys(docLibData[0]).forEach(key => {
          if (key !== 'ID') {
            dynamicFilters[key] = '';
          }
        });
        setFilters(dynamicFilters);
      }
    }
  }

  const handleFilterChange = (e: React.ChangeEvent<HTMLInputElement>, field: string) => {
    setFilters({
      ...filters,
      [field]: e.target.value,
    });
  };

  const handleSortChange = (key: string) => {
    let direction = 'ascending';
    if (sortConfig.key === key && sortConfig.direction === 'ascending') {
      direction = 'descending';
    }
    setSortConfig({ key, direction });
  };

  const applyFiltersAndSorting = (data: any[]) => {
    // Filter data
    const filteredData = data.filter((item) => {
      return Object.keys(filters).every(key => {
        if (key === 'SNo' || filters[key] === '') return true;
        
        const itemValue = item[key];
        const filterValue = filters[key].toLowerCase();
        
        if (itemValue === null || itemValue === undefined) return false;
        
        // Handle different data types
        if (typeof itemValue === 'number') {
          return itemValue.toString().includes(filterValue);
        }
        
        return itemValue.toString().toLowerCase().includes(filterValue);
      });
    });

    // Natural sort function
    const naturalSort = (a: any, b: any) => {
      return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
    };

    // Sort data
    const sortedData = filteredData.sort((a, b) => {
      if (sortConfig.key === 'SNo') {
        const aIndex = data.indexOf(a);
        const bIndex = data.indexOf(b);
        return sortConfig.direction === 'ascending' ? aIndex - bIndex : bIndex - aIndex;
      } else if (sortConfig.key) {
        const aValue = a[sortConfig.key] ? a[sortConfig.key].toString().toLowerCase() : '';
        const bValue = b[sortConfig.key] ? b[sortConfig.key].toString().toLowerCase() : '';
        return sortConfig.direction === 'ascending' ? naturalSort(aValue, bValue) : naturalSort(bValue, aValue);
      }
      return 0;
    });

    return sortedData;
  };

  const filteredAnnouncementData = applyFiltersAndSorting(mediaData);
  const [currentPage, setCurrentPage] = React.useState(1);
  const itemsPerPage = 10;
  const totalPages = Math.ceil(filteredAnnouncementData.length / itemsPerPage);

  const handlePageChange = (pageNumber: any) => {
    if (pageNumber > 0 && pageNumber <= totalPages) {
      setCurrentPage(pageNumber);
    }
  };

  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const currentData = filteredAnnouncementData.slice(startIndex, endIndex);

  const handleExportClick = () => {
    const exportData = currentData.map((item, index) => {
      if (Currentbuttonclick.buttonclickis === 'DocumentLibrary') {
        const row: any = { 'S.No.': startIndex + index + 1 };
        Object.entries(item).forEach(([key, value]) => {
          if (key !== 'ID') {
            row[key] = value;
          }
        });
        return row;
      } else {
        return {
          'S.No.': startIndex + index + 1,
          'FileName': item.FileName,
          'SubmittedBy': item.CurrentUser,
          'Modified': item.Modified,
          'Status': item.Status,
        };
      }
    });
    exportToExcel(exportData, 'MediaGallery');
  };

  const exportToExcel = (data: any[], fileName: string) => {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, `${fileName}.xlsx`);
  };

  const [isOpen, setIsOpen] = React.useState(false);
  const toggleDropdown = () => {
    setIsOpen(!isOpen);
  };

  const Deletemedia = (id: any) => {
    Swal.fire({
      title: "Are you sure?",
      text: "You won't be able to revert this!",
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#3085d6",
      cancelButtonColor: "#d33",
      confirmButtonText: "Yes, delete it!"
    }).then((result) => {
      if (result.isConfirmed) {
        ApiCall()
        Swal.fire({
          title: "Deleted!",
          text: "Item has been deleted.",
          icon: "success"
        });
      }
    })
  }

  interface PaginationProps {
    currentPage: number;
    totalPages: any;
    handlePageChange: any;
  }

  const Pagination = ({ currentPage, totalPages, handlePageChange }: PaginationProps) => {
    const pageLimit = 5;
    const startPage = Math.max(1, currentPage - Math.floor(pageLimit / 2));
    const endPage = Math.min(totalPages, startPage + pageLimit - 1);
    const adjustedStartPage = Math.max(1, Math.min(startPage, totalPages - pageLimit + 1));
    const visiblePages = [];
    const limit = Math.min(pageLimit, totalPages);
    
    for (let i = 0; i < limit; i++) {
      visiblePages.push(adjustedStartPage + i);
    }

    return (
      <nav>
        <ul className="pagination justify-content-center">
          <li className={`page-item ${currentPage === 1 ? 'disabled' : ''}`}>
            <a className="page-link" onClick={() => handlePageChange(currentPage - 1)} aria-label="Previous">
              <span aria-hidden="true">«</span>
            </a>
          </li>
          {visiblePages.map((pageNumber: any) => (
            <li key={pageNumber} className={`page-item ${currentPage === pageNumber ? 'active' : ''}`}>
              <a className="page-link" onClick={() => handlePageChange(pageNumber)}>
                {pageNumber}
              </a>
            </li>
          ))}
          <li className={`page-item ${currentPage === totalPages ? 'disabled' : ''}`}>
            <a className="page-link" onClick={() => handlePageChange(currentPage + 1)} aria-label="Next">
              <span aria-hidden="true">»</span>
            </a>
          </li>
        </ul>
      </nav>
    );
  };

  return (
    <>
      <div style={{background: '#fbfbfb'}} className="container-fluid">
        <div className="row">
          <div className="col-12">
            {/* <button className="btn btn-secondary mb-3" onClick={onReturnToMain}>
              ← Back
            </button> */}
            
            {/* <div className="d-flex justify-content-end mb-0 mt-0">
            
            </div> */}

            {/* Table for My Request and My Favourite */}
            {(Currentbuttonclick.buttonclickis === 'Myrequest' || Currentbuttonclick.buttonclickis === 'Favourite') && (
              <div className="d-grid">
                <table className="mtable  table-hover newtablescrollty">
                  <thead className="thead-dark">
                    <tr>
                      <th style={{minWidth:'50px',maxWidth:'50px'}} >S.No.</th>
                      <th >Document Name</th>
                      <th >Submitted By</th>
                      <th >Modified date</th>
                      <th >Status</th>
                      <th style={{minWidth:'70px',maxWidth:'70px'}}>Size

                      <button className="btn btn-success mt-0 newexport"  title="Export" onClick={handleExportClick}>
                <FontAwesomeIcon icon={faFileExport} />  
              </button>
                      </th>
                    </tr>
                    <tr>
                      <th style={{minWidth:'50px',maxWidth:'50px',paddingTop:'0px'}}>
                        <input
                          type="text"
                          className="form-control form-control-sm"
                          placeholder="Search..."
                          value={filters.SNo}
                          onChange={(e) => handleFilterChange(e, 'SNo')}
                          disabled
                        />
                      </th>
                      <th style={{paddingTop:'0px'}}>
                        <input
                          type="text"
                          className="form-control form-control-sm"
                          placeholder="Search..."
                          value={filters.FileName}
                          onChange={(e) => handleFilterChange(e, 'FileName')}
                        />
                      </th>
                      <th style={{paddingTop:'0px'}}>
                        <input
                          type="text"
                          className="form-control form-control-sm"
                          placeholder="Search..."
                          value={filters.CurrentUser}
                          onChange={(e) => handleFilterChange(e, 'CurrentUser')}
                        />
                      </th>
                      <th style={{paddingTop:'0px'}}>
                        <input
                          type="text"
                          className="form-control form-control-sm"
                          placeholder="Search..."
                          value={filters.Modified}
                          onChange={(e) => handleFilterChange(e, 'Modified')}
                        />
                      </th>
                      <th style={{paddingTop:'0px'}}>
                        <input
                          type="text"
                          className="form-control form-control-sm"
                          placeholder="Search..."
                          value={filters.Status}
                          onChange={(e) => handleFilterChange(e, 'Status')}
                        />
                      </th>
                      <th style={{minWidth:'70px',maxWidth:'70px',paddingTop:'0px'}}>
                        <input
                          type="text"
                          className="form-control form-control-sm"
                          placeholder="Search..."
                          value={filters.FileSize}
                          onChange={(e) => handleFilterChange(e, 'FileSize')}
                        />
                      </th>
                    </tr>
                  </thead>
                  <tbody style={{maxHeight:'5000px'}}>
                    {currentData.length > 0 ? (
                      currentData.map((item, index) => (
                        <tr key={item.ID}>
                          <td  style={{minWidth:'50px',maxWidth:'50px', fontSize:'14px', borderBottom:'1px solid #efefef'}}>{startIndex + index + 1}</td>
                          <td style={{borderBottom:'1px solid #efefef', fontSize:'14px',}}>{item.FileName}</td>
                          <td style={{borderBottom:'1px solid #efefef', fontSize:'14px',}}>{item.CurrentUser}</td>
                          <td style={{borderBottom:'1px solid #efefef', fontSize:'14px',}}>
                            {new Date(item?.Modified).toLocaleString('en-GB', {
                              day: '2-digit',
                              month: '2-digit',
                              year: 'numeric',
                              hour: '2-digit',
                              minute: '2-digit',
                              hour12: true
                            })}
                          </td>
                          <td style={{borderBottom:'1px solid #efefef', fontSize:'14px',}}>{item.Status}</td>
                          <td  style={{minWidth:'70px',maxWidth:'70px', fontSize:'14px',borderBottom:'1px solid #efefef'}}>{((item.FileSize as unknown as number) / (1024 * 1024)).toFixed(2)} MB</td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={6} className="text-center">No results found</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            )}

            {/* Table for Document Library - Dynamic Columns */}
            {Currentbuttonclick.buttonclickis === 'DocumentLibrary' && (
                            <div className="d-grid mb-2">
                <table className="mtable  table-hover newtablescrollty">
                  <thead className="thead-dark">
                    <tr>
                      <th style={{ minWidth: '50px',maxWidth:'50px' }}>S.No.</th>
                      {mediaData.length > 0 &&
                        Object.keys(mediaData[0])
                          .filter(key => key !== 'ID')
                          .map((key) => (
                            <th key={key}>{key}</th>
                          ))}
                    </tr>
                    <tr>
                      <th style={{ minWidth: '50px',maxWidth:'50px' }}>
                        <input
                          type="text"
                          className="form-control form-control-sm"
                          placeholder="Search..."
                          disabled
                        />
                      </th>
                      {mediaData.length > 0 &&
                        Object.keys(mediaData[0])
                          .filter(key => key !== 'ID')
                          .map((key) => (
                            <th style={{paddingTop:'0px'}} key={key}>
                              <input
                                type="text"
                                className="form-control form-control-sm"
                                placeholder="Search..."
                                value={filters[key] || ''}
                                onChange={(e) => handleFilterChange(e, key)}
                              />
                            </th>
                          ))}
                    </tr>
                  </thead>
                  <tbody>
                    {currentData.length > 0 ? (
                      currentData.map((item, index) => (
                        <tr key={index}>
                          <td style={{ minWidth: '50px',maxWidth:'50px', fontSize:'14px', borderBottom:'1px solid #efefef' }}>{startIndex + index + 1}</td>
                          {Object.entries(item)
                            .filter(([key]) => key !== 'ID')
                            .map(([key, value], i) => (
                              <td  style={{ borderBottom:'1px solid #efefef', fontSize:'14px', }} key={i}>{value as string}</td>
                            ))}
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={mediaData.length > 0 ? Object.keys(mediaData[0]).length : 1} className="text-center">
                          No results found
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            )}

            <Pagination 
              currentPage={currentPage}
              totalPages={totalPages}
              handlePageChange={handlePageChange}
            />
          </div>
        </div>
      </div>
    </>
  );
};

export default Table;