using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using ExcelDataReader;
using System.Web.Mvc;
using System.Net.Http;
using CostAllocationApp.Dtos;
using CostAllocationApp.BLL;
using CostAllocationApp.ViewModels;
using CostAllocationApp.Models;
using CostAllocationApp.DAL;
using System.Data;

namespace CostAllocationApp.Controllers
{
    public class ForecastsController : Controller
    {

        EmployeeAssignmentBLL employeeAssignmentBLL = new EmployeeAssignmentBLL();
        UploadExcelBLL _uploadExcelBll = new UploadExcelBLL();
        UploadExcel _uploadExcel;
        char[] trimElements = { '\r', '\n', ' ' };
        SectionBLL sectionBLL = new SectionBLL();
        private DepartmentBLL _departmentBLL = new DepartmentBLL();
        // GET: Forecasts
        public ActionResult CreateForecast(string departmentId = "")
        {
            //string departmentId = Request.QueryString["departmentId"];

            int tempDepartmentId = 0;
            Int32.TryParse(departmentId, out tempDepartmentId);
            if (tempDepartmentId == 0)
            {
                return RedirectToAction("Index", "Dashboard");
            }
            var singleDepartment = _departmentBLL.GetDepartmentByDepartemntId(tempDepartmentId);
            if (singleDepartment == null)
            {
                return RedirectToAction("Index", "Dashboard");
            }

            if (TempData["seccess"] != null)
            {
                ViewBag.Success = TempData["seccess"];
            }
            else
            {
                ViewBag.Success = null;
            }
            ForecastViewModal forecastViewModal = new ForecastViewModal
            {
                Departments = _departmentBLL.GetAllDepartments(),
                Department = singleDepartment,
            };
            ViewBag.ErrorCount = 0;
            return View(forecastViewModal);
        }

        //[NonAction]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Index(HttpPostedFileBase uploaded_file, int upload_year,int department_id)
        {
            ForecastViewModal forecastViewModal = new ForecastViewModal
            {
                Departments = _departmentBLL.GetAllDepartments(),
                Department = _departmentBLL.GetDepartmentByDepartemntId(department_id)
            };

            TempData["seccess"] = null;
            Dictionary<int, int> check = new Dictionary<int, int>();
            if (ModelState.IsValid)
            {

                if ((uploaded_file != null && uploaded_file.ContentLength > 0) && upload_year > 0)
                {
                    // ExcelDataReader works with the binary Excel file, so it needs a FileStream
                    // to get started. This is how we avoid dependencies on ACE or Interop:
                    Stream stream = uploaded_file.InputStream;

                    IExcelDataReader reader = null;


                    //if (uploaded_file.FileName.EndsWith(".xls"))
                    //{
                    //    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    //}
                    if (uploaded_file.FileName.EndsWith(".xlsx"))
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else
                    {
                        ModelState.AddModelError("File", "This file format is not supported");
                        ViewBag.ErrorCount = 1;
                        return View("CreateForecast", forecastViewModal);
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
                        int tempYear = upload_year;
                        tempYear = 2022;
                        dt_ = reader.AsDataSet().Tables[0];
                        int rowcount = dt_.Rows.Count;

                        for (int i = 2; i < rowcount; i++)
                        {                            
                            _uploadExcel = new UploadExcel();
                            if (string.IsNullOrEmpty(dt_.Rows[i][0].ToString().Trim(trimElements)))
                            {
                                continue;
                            }
                            else
                            {
                                _uploadExcel.EmployeeName = dt_.Rows[i][0].ToString().Trim(trimElements);
                            }                            

                            int? sectionId = _uploadExcelBll.GetSectionIdByName(dt_.Rows[i][1].ToString().Trim(trimElements));
                            _uploadExcel.SectionId = sectionId == 0 ? null : sectionId;

                            int? companyId = _uploadExcelBll.GetCompanyIdByName(dt_.Rows[i][2].ToString().Trim(trimElements));
                            _uploadExcel.CompanyId = companyId == 0 ? null : companyId;

                            int? departmentId = _uploadExcelBll.GetDepartmentIdByName(dt_.Rows[i][3].ToString().Trim(trimElements));
                            _uploadExcel.DepartmentId = departmentId == 0 ? null : departmentId;

                            //allocation
                            int? explanationId = _uploadExcelBll.GetExplanationIdByName(dt_.Rows[i][6].ToString().Trim(trimElements));
                            _uploadExcel.ExplanationId = explanationId == 0 ? null : explanationId;
                            
                            decimal tempUnitPrice = 0;
                            decimal.TryParse(dt_.Rows[i][5].ToString(), out tempUnitPrice);

                            string tempUnitPrice2 = dt_.Rows[i][5].ToString();
                            string tempGradePoint = dt_.Rows[i][4].ToString();
                            UploadExcel _salary = new UploadExcel();
                            if (tempUnitPrice > 0)
                            {
                                //int? gradeId = _uploadExcelBll.GetGradeIdByUnitPrice(tempUnitPrice.ToString());
                                //_uploadExcel.GradeId = gradeId == 0 ? null : gradeId;
                                _uploadExcel.GradeId = null;
                                _uploadExcel.UnitPrice = tempUnitPrice;
                            }
                            else if (!string.IsNullOrEmpty(dt_.Rows[i][4].ToString()))
                            {
                                _salary = _uploadExcelBll.GetGradeIdByGradePoints(dt_.Rows[i][4].ToString());
                                _uploadExcel.GradeId = _salary.GradeId;

                                GradeSalaryType _salaryType = new GradeSalaryType();
                                if (!string.IsNullOrEmpty(_uploadExcel.GradeId.ToString()) && !string.IsNullOrEmpty(_uploadExcel.DepartmentId.ToString()))
                                {
                                    _salaryType = _uploadExcelBll.GetGradeSalaryTypeIdByGradeId(_uploadExcel.GradeId, _uploadExcel.DepartmentId,2022,2);
                                    _uploadExcel.GradeId = _salaryType.Id;
                                    _uploadExcel.UnitPrice = Convert.ToDecimal(_salaryType.GradeLowPoints) ;
                                }
                                else
                                {
                                    _uploadExcel.GradeId = null;
                                    _uploadExcel.UnitPrice = 0;
                                }
                                //if(!string.IsNullOrEmpty(_salary.GradeHighPoints)){
                                //    _uploadExcel.UnitPrice = Convert.ToDecimal(_salary.GradeHighPoints);
                                //}else{
                                //    _uploadExcel.UnitPrice = 0;
                                //}                                
                            }
                            else
                            {
                                _uploadExcel.GradeId = null;
                                _uploadExcel.UnitPrice = 0;
                            }                           
                            
                            var assignmentViewModels = employeeAssignmentBLL.GetEmployeesByName(_uploadExcel.EmployeeName);
                            if (assignmentViewModels.Count > 0)
                            {
                                CreateAssignmentForExcelUpload(_uploadExcel, i, assignmentViewModels.Count);
                                tempAssignmentId = employeeAssignmentBLL.GetLastId();
                            }
                            else
                            {
                                //CreateAssignmentForExcelUpload(dt_, i);
                                CreateAssignmentForExcelUpload(_uploadExcel, i);
                                tempAssignmentId = employeeAssignmentBLL.GetLastId();
                            }

                            decimal octInput = 0;
                            decimal.TryParse(dt_.Rows[i][7].ToString(), out octInput);

                            decimal novInput = 0;
                            decimal.TryParse(dt_.Rows[i][8].ToString(), out novInput);

                            decimal decInput = 0;
                            decimal.TryParse(dt_.Rows[i][9].ToString(), out decInput);

                            decimal janInput = 0;
                            decimal.TryParse(dt_.Rows[i][10].ToString(), out janInput);

                            decimal febInput = 0;
                            decimal.TryParse(dt_.Rows[i][11].ToString(), out febInput);

                            decimal marInput = 0;
                            decimal.TryParse(dt_.Rows[i][12].ToString(), out marInput);

                            decimal aprInput = 0;
                            decimal.TryParse(dt_.Rows[i][13].ToString(), out aprInput);

                            decimal mayInput = 0;
                            decimal.TryParse(dt_.Rows[i][14].ToString(), out mayInput);

                            decimal junInput = 0;
                            decimal.TryParse(dt_.Rows[i][15].ToString(), out junInput);

                            decimal julInput = 0;
                            decimal.TryParse(dt_.Rows[i][16].ToString(), out julInput);

                            decimal augInput = 0;
                            decimal.TryParse(dt_.Rows[i][17].ToString(), out augInput);

                            decimal septInput = 0;
                            decimal.TryParse(dt_.Rows[i][18].ToString(), out septInput);

                            tempRow = $@"10_{octInput}_{octInput * _uploadExcel.UnitPrice},11_{novInput}_{novInput * _uploadExcel.UnitPrice},12_{decInput}_{decInput * _uploadExcel.UnitPrice},1_{janInput}_{janInput * _uploadExcel.UnitPrice},2_{febInput}_{febInput * _uploadExcel.UnitPrice},3_{marInput}_{marInput * _uploadExcel.UnitPrice},4_{aprInput}_{aprInput * _uploadExcel.UnitPrice},5_{mayInput}_{mayInput * _uploadExcel.UnitPrice},6_{junInput}_{junInput * _uploadExcel.UnitPrice},7_{julInput}_{julInput * _uploadExcel.UnitPrice},8_{augInput}_{augInput * _uploadExcel.UnitPrice},9_{septInput}_{septInput * _uploadExcel.UnitPrice}";

                            SendToApi(tempAssignmentId, tempRow, tempYear, _uploadExcel.ExplanationId);
                        }


                    }
                    catch (Exception ex)
                    {
                        ModelState.AddModelError("File", ex);
                        ViewBag.ErrorCount = 1;
                        return View("CreateForecast", forecastViewModal);
                    }

                    //DataSet result = new DataSet();
                    //result.Tables.Add(dt);
                    reader.Close();
                    reader.Dispose();
                    //DataTable tmp = result.Tables[0];
                    //Session["tmpdata"] = tmp;  //store datatable into session
                    TempData["seccess"] = "Data imported successfully";
                    return RedirectToAction("CreateForecast",new { departmentId =department_id});
                }
                else
                {
                    ViewBag.ErrorCount = 1;
                    ModelState.AddModelError("File", "invalid File or Year");
                }
            }
            return View("CreateForecast", forecastViewModal);
        }

        //[NonAction]
        public void SendToApi(int assignmentId, string row, int year, int? allocationId)
        {

            SendToForecaseApiDto sendToForecaseApiDto = new SendToForecaseApiDto();
            sendToForecaseApiDto.Data = row;
            sendToForecaseApiDto.Year = year;
            sendToForecaseApiDto.AssignmentId = assignmentId;
            sendToForecaseApiDto.AllocationId = allocationId;


            using (var client = new HttpClient())
            {
                //client.BaseAddress = new Uri("http://198.38.92.119:8081/api/Forecasts?data=" + row + "&year=" + year + "&assignmentId=" + assignmentId + "&allocationId=" + allocationId);
                client.BaseAddress = new Uri("http://localhost:59198/api/Forecasts?data=" + row + "&year=" + year + "&assignmentId=" + assignmentId+ "&allocationId=" + allocationId);

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
        //[NonAction]
        public int CreateAssignmentForExcelUpload(UploadExcel dt_, int i, int subCodeCount = 0)
        {
            EmployeeAssignmentDTO employeeAssignmentDTO = new EmployeeAssignmentDTO();
            EmployeeAssignment employeeAssignment = new EmployeeAssignment();

            employeeAssignmentDTO = new EmployeeAssignmentDTO();
            employeeAssignment.EmployeeName = dt_.EmployeeName;
            employeeAssignment.SectionId = dt_.SectionId;
            employeeAssignment.DepartmentId = dt_.DepartmentId;
            employeeAssignment.CompanyId = dt_.CompanyId;
            employeeAssignment.ExplanationId = dt_.ExplanationId == null ? null : dt_.ExplanationId.ToString();
            employeeAssignment.UnitPrice = dt_.UnitPrice;
            employeeAssignment.GradeId = dt_.GradeId;
            employeeAssignment.SubCode = subCodeCount + 1;

            employeeAssignment.CreatedBy = "";
            employeeAssignment.CreatedDate = DateTime.Now;
            employeeAssignment.IsActive = "1";
            employeeAssignment.Remarks = "";


            int result = employeeAssignmentBLL.CreateAssignment(employeeAssignment);
            if (result == 0)
            {
                throw new Exception();
            }
            return result;
        }
    }
}