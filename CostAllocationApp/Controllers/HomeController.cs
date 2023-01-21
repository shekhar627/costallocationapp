using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Web;
using System.Web.Mvc;
using CostAllocationApp.Dtos;
using CostAllocationApp.BLL;
using CostAllocationApp.ViewModels;
using CostAllocationApp.Models;

namespace CostAllocationApp.Controllers
{
    public class HomeController : Controller
    {
        EmployeeAssignmentBLL employeeAssignmentBLL = new EmployeeAssignmentBLL();
        UploadExcelBLL _uploadExcelBll = new UploadExcelBLL();
        UploadExcel _uploadExcel;

        // GET: Home
        public ActionResult Index()
        {
            DataTable dt = new DataTable();

            try
            {
                dt = (DataTable)Session["tmpdata"];
            }
            catch (Exception ex)
            {

            }

            return View(dt);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Index(HttpPostedFileBase upload)
        {
            Dictionary<int, int> check = new Dictionary<int, int>();
            if (ModelState.IsValid)
            {

                if (upload != null && upload.ContentLength > 0)
                {
                    // ExcelDataReader works with the binary Excel file, so it needs a FileStream
                    // to get started. This is how we avoid dependencies on ACE or Interop:
                    Stream stream = upload.InputStream;

                    IExcelDataReader reader = null;


                    if (upload.FileName.EndsWith(".xls"))
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (upload.FileName.EndsWith(".xlsx"))
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else
                    {
                        ModelState.AddModelError("File", "This file format is not supported");
                        return View();
                    }
                    //int fieldcount = reader.FieldCount;
                    //int rowcount = reader.RowCount;
                    //DataTable dt = new DataTable();
                    DataRow row;
                    DataTable dt_ = new DataTable();
                    try
                    {
                        int tempAssignmentId = 0;
                        string tempRow = "";
                        int tempYear = 2023;
                        dt_ = reader.AsDataSet().Tables[0];
                        int rowcount = dt_.Rows.Count;

                        for (int i = 1; i < rowcount; i++)
                        {
                            _uploadExcel = new UploadExcel();                            
                            _uploadExcel.SectionId = _uploadExcelBll.GetSectionIdByName(dt_.Rows[i][0].ToString().Trim());
                            _uploadExcel.DepartmentId = _uploadExcelBll.GetDepartmentIdByName(dt_.Rows[i][1].ToString().Trim());
                            _uploadExcel.InchargeId = _uploadExcelBll.GetInchargeIdByName(dt_.Rows[i][2].ToString().Trim());
                            _uploadExcel.RoleId = _uploadExcelBll.GetRoleIdByName(dt_.Rows[i][3].ToString().Trim());
                            _uploadExcel.ExplanationId = _uploadExcelBll.GetExplanationIdByName(dt_.Rows[i][4].ToString().Trim());
                            _uploadExcel.CompanyId = _uploadExcelBll.GetCompanyIdByName(dt_.Rows[i][6].ToString().Trim());
                            _uploadExcel.GradeId = _uploadExcelBll.GetGradeIdByUnitPrice(Convert.ToInt32(dt_.Rows[i][8].ToString().Trim()));
                            //IsValidMasterSetting(int masterSettingId);
                            if (string.IsNullOrEmpty(_uploadExcel.SectionId.ToString()))
                            {
                                continue;
                            }
                            if (string.IsNullOrEmpty(_uploadExcel.DepartmentId.ToString()))
                            {
                                continue;
                            }
                            if (string.IsNullOrEmpty(_uploadExcel.InchargeId.ToString()))
                            {
                                continue;
                            }
                            if (string.IsNullOrEmpty(_uploadExcel.RoleId.ToString()))
                            {
                                continue;
                            }
                            if (string.IsNullOrEmpty(_uploadExcel.CompanyId.ToString()))
                            {
                                continue;
                            }
                            if (string.IsNullOrEmpty(_uploadExcel.GradeId.ToString()))
                            {
                                continue;
                            }

                            //if (Convert.ToInt32(dt_.Rows[i][7])==0)
                            //{
                            //    continue;
                            //}
                            //if (Convert.ToDecimal(dt_.Rows[i][8]) <= 0)
                            //{
                            //    continue;
                            //}
                            //int roleId = 0;
                            //Int32.TryParse(dt_.Rows[i][3].ToString(), out roleId);
                            //if (roleId<=0)
                            //{
                            //    continue;
                            //}

                            //if (String.IsNullOrEmpty(dt_.Rows[i][5].ToString()))
                            //{
                            //    continue;
                            //}

                            var assignmentViewModels = employeeAssignmentBLL.GetEmployeesByName(dt_.Rows[i][5].ToString().Trim());
                            _uploadExcel.EmployeeName = dt_.Rows[i][5].ToString().Trim();
                            _uploadExcel.UnitPrice = Convert.ToInt32(dt_.Rows[i][8].ToString().Trim());

                            if (assignmentViewModels.Count > 0)
                            {
                                //CreateAssignmentForExcelUpload(dt_, i, assignmentViewModels.Count);
                                CreateAssignmentForExcelUpload(_uploadExcel, i, assignmentViewModels.Count);
                                tempAssignmentId = employeeAssignmentBLL.GetLastId();
                            }
                            else
                            {
                                //CreateAssignmentForExcelUpload(dt_, i);
                                CreateAssignmentForExcelUpload(_uploadExcel, i);
                                tempAssignmentId = employeeAssignmentBLL.GetLastId();
                            }
                            check.Add(i, tempAssignmentId);


                            //tempAssignmentId = EmployeeAssignment();
                            //EmployeeAssignmentViewModel employeeAssignment = employeeAssignmentBLL.GetAssignmentById(tempAssignmentId);
                            //tempAssignmentId = Convert.ToInt32(dt_.Rows[i][0]);
                            //tempMonthWiseUnitPrice
                            tempRow = $@"10_{dt_.Rows[i][9].ToString()}_{Convert.ToDecimal(dt_.Rows[i][9]) * Convert.ToDecimal(dt_.Rows[i][8])},11_{dt_.Rows[i][10].ToString()}_{Convert.ToDecimal(dt_.Rows[i][10]) * Convert.ToDecimal(dt_.Rows[i][8])},12_{dt_.Rows[i][11].ToString()}_{Convert.ToDecimal(dt_.Rows[i][11]) * Convert.ToDecimal(dt_.Rows[i][8])},1_{dt_.Rows[i][12].ToString()}_{Convert.ToDecimal(dt_.Rows[i][12]) * Convert.ToDecimal(dt_.Rows[i][8])},2_{dt_.Rows[i][13].ToString()}_{Convert.ToDecimal(dt_.Rows[i][13]) * Convert.ToDecimal(dt_.Rows[i][8])},3_{dt_.Rows[i][14].ToString()}_{Convert.ToDecimal(dt_.Rows[i][14]) * Convert.ToDecimal(dt_.Rows[i][8])},4_{dt_.Rows[i][15].ToString()}_{Convert.ToDecimal(dt_.Rows[i][15]) * Convert.ToDecimal(dt_.Rows[i][8])},5_{dt_.Rows[i][16].ToString()}_{Convert.ToDecimal(dt_.Rows[i][16]) * Convert.ToDecimal(dt_.Rows[i][8])},6_{dt_.Rows[i][17].ToString()}_{Convert.ToDecimal(dt_.Rows[i][17]) * Convert.ToDecimal(dt_.Rows[i][8])},7_{dt_.Rows[i][18].ToString()}_{Convert.ToDecimal(dt_.Rows[i][18]) * Convert.ToDecimal(dt_.Rows[i][8])},8_{dt_.Rows[i][19].ToString()}_{Convert.ToDecimal(dt_.Rows[i][19]) * Convert.ToDecimal(dt_.Rows[i][8])},9_{dt_.Rows[i][20].ToString()}_{Convert.ToDecimal(dt_.Rows[i][20]) * Convert.ToDecimal(dt_.Rows[i][8])}";

                            SendToApi(tempAssignmentId, tempRow, tempYear);
                        }


                    }
                    catch (Exception ex)
                    {
                        ModelState.AddModelError("File", ex);
                        return View();
                    }

                    //DataSet result = new DataSet();
                    //result.Tables.Add(dt);
                    reader.Close();
                    reader.Dispose();
                    //DataTable tmp = result.Tables[0];
                    //Session["tmpdata"] = tmp;  //store datatable into session
                    return RedirectToAction("Index");
                }
                else
                {
                    ModelState.AddModelError("File", "Please Upload Your file");
                }
            }
            return View();
        }

        public int EmployeeAssignment()
        {
            return 0;
        }
        public void SendToApi(int assignmentId, string row, int year)
        {

            SendToForecaseApiDto sendToForecaseApiDto = new SendToForecaseApiDto();
            sendToForecaseApiDto.Data = row;
            sendToForecaseApiDto.Year = year;
            sendToForecaseApiDto.AssignmentId = assignmentId;

            using (var client = new HttpClient())
            {
                //client.BaseAddress = new Uri("http://198.38.92.119:8081/api/Forecasts?data=" + row + "&year=" + year + "&assignmentId=" + assignmentId);
                client.BaseAddress = new Uri("http://localhost:59198/api/Forecasts?data=" + row + "&year=" + year + "&assignmentId=" + assignmentId);

                //HTTP POST
                var postTask = client.GetAsync("");
                postTask.Wait();

                var result = postTask.Result;
                if (result.IsSuccessStatusCode)
                {
                    //return RedirectToAction("Index");
                }
            }

        }

        //public int CreateAssignmentForExcelUpload(DataTable dt_, int i, int subCodeCount = 0)
        public int CreateAssignmentForExcelUpload(UploadExcel dt_, int i, int subCodeCount = 0)
        {
            EmployeeAssignmentDTO employeeAssignmentDTO = new EmployeeAssignmentDTO();
            EmployeeAssignment employeeAssignment = new EmployeeAssignment();

            employeeAssignmentDTO = new EmployeeAssignmentDTO();
            employeeAssignment.EmployeeName = dt_.EmployeeName;
            employeeAssignment.SectionId = Convert.ToInt32(dt_.SectionId.ToString().Trim());
            employeeAssignment.InchargeId = Convert.ToInt32(dt_.InchargeId.ToString().Trim());
            employeeAssignment.DepartmentId = Convert.ToInt32(dt_.DepartmentId.ToString().Trim());
            employeeAssignment.RoleId = Convert.ToInt32(dt_.RoleId.ToString().Trim());
            employeeAssignment.CompanyId = Convert.ToInt32(dt_.CompanyId.ToString().Trim());
            employeeAssignment.ExplanationId = String.IsNullOrEmpty(dt_.ExplanationId.ToString()) ? null : dt_.ExplanationId.ToString().Trim();
            employeeAssignment.UnitPrice = Convert.ToInt32(dt_.UnitPrice.ToString().Trim());
            employeeAssignment.GradeId = Convert.ToInt32(dt_.GradeId.ToString().Trim());
            employeeAssignment.SubCode = subCodeCount + 1;
            //employeeAssignment.EmployeeName = dt_.Rows[i][5].ToString().Trim();
            //employeeAssignment.SectionId = Convert.ToInt32(dt_.Rows[i][0].ToString().Trim());
            //employeeAssignment.InchargeId = Convert.ToInt32(dt_.Rows[i][2].ToString().Trim());
            //employeeAssignment.DepartmentId = Convert.ToInt32(dt_.Rows[i][1].ToString().Trim());
            //employeeAssignment.RoleId = Convert.ToInt32(dt_.Rows[i][3].ToString().Trim());
            //employeeAssignment.CompanyId = Convert.ToInt32(dt_.Rows[i][6].ToString().Trim());
            //employeeAssignment.ExplanationId = String.IsNullOrEmpty(dt_.Rows[i][4].ToString()) ? null : dt_.Rows[i][4].ToString().Trim();
            //employeeAssignment.UnitPrice = Convert.ToInt32(dt_.Rows[i][8].ToString().Trim());
            //employeeAssignment.GradeId = Convert.ToInt32(dt_.Rows[i][7].ToString().Trim());
            //employeeAssignment.SubCode = subCodeCount + 1;
            //var remarks = $('#memo_new').val();

            int tempValue = 0;
            decimal tempUnitPrice = 0;

            #region validation of inputs
            //if (!String.IsNullOrEmpty(employeeAssignmentDTO.EmployeeName))
            //{
            //    var checkResult = employeeAssignmentBLL.CheckEmployeeName(employeeAssignmentDTO.EmployeeName.Trim());
            //    if (checkResult && employeeAssignmentDTO.SubCode == 1)
            //    {
            //        return 0;
            //    }
            //    else
            //    {
            //        employeeAssignment.EmployeeName = employeeAssignmentDTO.EmployeeName.Trim();
            //    }

            //}
            //else
            //{
            //    return 0;
            //}
            //if (int.TryParse(employeeAssignmentDTO.SectionId, out tempValue))
            //{
            //    if (tempValue <= 0)
            //    {
            //        return 0;

            //    }
            //    employeeAssignment.SectionId = tempValue;
            //}
            //else
            //{
            //    return 0;
            //}
            //if (int.TryParse(employeeAssignmentDTO.DepartmentId, out tempValue))
            //{
            //    if (tempValue <= 0)
            //    {

            //        return 0;
            //    }
            //    employeeAssignment.DepartmentId = tempValue;
            //}
            //else
            //{
            //    return 0;
            //}
            //if (int.TryParse(employeeAssignmentDTO.InchargeId, out tempValue))
            //{
            //    if (tempValue <= 0)
            //    {
            //        return 0;

            //    }
            //    employeeAssignment.InchargeId = tempValue;
            //}
            //else
            //{
            //    return 0;
            //}
            //if (String.IsNullOrEmpty(employeeAssignmentDTO.RoleId))
            //{
            //    employeeAssignment.ExplanationId = null;
            //}
            //else
            //{
            //    if (int.TryParse(employeeAssignmentDTO.RoleId, out tempValue))
            //    {
            //        if (tempValue <= 0)
            //        {
            //            return 0;

            //        }
            //        employeeAssignment.RoleId = tempValue;
            //    }
            //    else
            //    {
            //        return 0;
            //    }
            //}
            //if (String.IsNullOrEmpty(employeeAssignmentDTO.ExplanationId))
            //{
            //    employeeAssignment.ExplanationId = null;
            //}
            //else
            //{
            //    if (int.TryParse(employeeAssignmentDTO.ExplanationId, out tempValue))
            //    {
            //        if (tempValue <= 0)
            //        {

            //            return 0;
            //        }
            //        employeeAssignment.ExplanationId = tempValue.ToString();
            //    }
            //    else
            //    {
            //        return 0;
            //    }
            //}


            //if (int.TryParse(employeeAssignmentDTO.CompanyId, out tempValue))
            //{
            //    if (tempValue <= 0)
            //    {
            //        return 0;

            //    }
            //    employeeAssignment.CompanyId = tempValue;
            //}
            //else
            //{
            //    return 0;
            //}
            //if (decimal.TryParse(employeeAssignmentDTO.UnitPrice, out tempUnitPrice))
            //{
            //    if (tempValue < 0)
            //    {
            //        return 0;

            //    }
            //    employeeAssignment.UnitPrice = tempUnitPrice;
            //}
            //else
            //{
            //    return 0;
            //}
            //if (int.TryParse(employeeAssignmentDTO.GradeId, out tempValue))
            //{
            //    if (tempValue <= 0)
            //    {
            //        return 0;

            //    }
            //    employeeAssignment.GradeId = tempValue;
            //}
            //else
            //{
            //    return 0;
            //}

            //if (String.IsNullOrEmpty(employeeAssignmentDTO.Remarks))
            //{
            //    employeeAssignmentDTO.Remarks = "";
            //}
            #endregion

            //employeeAssignment.ExplanationId = employeeAssignmentDTO.ExplanationId;
            employeeAssignment.CreatedBy = "";
            employeeAssignment.CreatedDate = DateTime.Now;
            employeeAssignment.IsActive = "1";
            employeeAssignment.Remarks = "";
            //employeeAssignment.SubCode = employeeAssignmentDTO.SubCode;


            int result = employeeAssignmentBLL.CreateAssignment(employeeAssignment);
            if (result == 0)
            {
                throw new Exception();
            }
            return result;
            //if (result > 0)
            //{
            //    return true;
            //}
            //else
            //{
            //    return false;
            //}
        }

    }
}