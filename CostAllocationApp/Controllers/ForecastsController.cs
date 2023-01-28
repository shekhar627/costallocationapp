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
        // GET: Forecasts
        public ActionResult CreateForecast(string sectionId="")
        {
            int tempSectionId = 0;
            Int32.TryParse(sectionId, out tempSectionId);
            if(tempSectionId==0)
            {
                return RedirectToAction("Index","Dashboard");
            }
            var singleSection = sectionBLL.GetSectionBySectionId(tempSectionId);
            if (singleSection == null)
            {
                return RedirectToAction("Index", "Dashboard");  
            }
            
            if (TempData["seccess"]!=null)
            {
                ViewBag.Success = TempData["seccess"];
            }
            else
            {
                ViewBag.Success = null;
            }
            ForecastViewModal forecastViewModal = new ForecastViewModal
            {
                _sections = new List<Section> { singleSection },
                Sections = sectionBLL.GetAllSections(),
                SectionId = singleSection.Id
            };            
            ViewBag.ErrorCount = 0;
            return View(forecastViewModal);
        }

        //[NonAction]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Index(HttpPostedFileBase uploaded_file, int upload_year)
        {
            ForecastViewModal forecastViewModal = new ForecastViewModal
            {
                _sections = sectionBLL.GetAllSections()
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
                        return View("CreateForecast");
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
                        dt_ = reader.AsDataSet().Tables[0];
                        int rowcount = dt_.Rows.Count;

                        for (int i = 1; i < rowcount; i++)
                        {
                            _uploadExcel = new UploadExcel();
                            if (i == 56)
                            {

                            }

                            int? sectionId = _uploadExcelBll.GetSectionIdByName(dt_.Rows[i][0].ToString().Trim(trimElements));
                            _uploadExcel.SectionId = sectionId == 0 ? null : sectionId;
                            int? departmentId = _uploadExcelBll.GetDepartmentIdByName(dt_.Rows[i][1].ToString().Trim(trimElements));
                            _uploadExcel.DepartmentId = departmentId == 0 ? null : departmentId;
                            int? explanationId = _uploadExcelBll.GetExplanationIdByName(dt_.Rows[i][4].ToString().Trim(trimElements));
                            _uploadExcel.ExplanationId = explanationId == 0 ? null : explanationId;
                            int? companyId = _uploadExcelBll.GetCompanyIdByName(dt_.Rows[i][6].ToString().Trim(trimElements));
                            _uploadExcel.CompanyId = companyId == 0 ? null : companyId;

                            //if (string.IsNullOrEmpty(dt_.Rows[i][0].ToString()))
                            //{
                            //    continue;
                            //}
                            //else
                            //{
                            //    _uploadExcel.SectionId = _uploadExcelBll.GetSectionIdByName(dt_.Rows[i][0].ToString().Trim(trimElements));
                            //    if (_uploadExcel.SectionId == 0)
                            //    {
                            //        continue;
                            //    }
                            //}

                            //if (string.IsNullOrEmpty(dt_.Rows[i][1].ToString()))
                            //{
                            //    continue;
                            //}
                            //else
                            //{
                            //    _uploadExcel.DepartmentId = _uploadExcelBll.GetDepartmentIdByName(dt_.Rows[i][1].ToString().Trim(trimElements));
                            //    if (_uploadExcel.DepartmentId == 0)
                            //    {
                            //        continue;
                            //    }
                            //}

                            //if (string.IsNullOrEmpty(dt_.Rows[i][2].ToString()))
                            //{
                            //    continue;
                            //}
                            //else
                            //{
                            //    _uploadExcel.InchargeId = _uploadExcelBll.GetInchargeIdByName(dt_.Rows[i][2].ToString().Trim(trimElements));
                            //    if (_uploadExcel.InchargeId == 0)
                            //    {
                            //        continue;
                            //    }
                            //}

                            //if (string.IsNullOrEmpty(dt_.Rows[i][3].ToString()))
                            //{
                            //    continue;
                            //}
                            //else
                            //{
                            //    _uploadExcel.RoleId = _uploadExcelBll.GetRoleIdByName(dt_.Rows[i][3].ToString().Trim(trimElements));
                            //    if (_uploadExcel.RoleId == 0)
                            //    {
                            //        continue;
                            //    }
                            //}

                            //if (!string.IsNullOrEmpty(dt_.Rows[i][4].ToString()))
                            //{
                            //    _uploadExcel.ExplanationId = _uploadExcelBll.GetExplanationIdByName(dt_.Rows[i][4].ToString().Trim(trimElements));
                            //}
                            //else
                            //{
                            //    _uploadExcel.ExplanationId = null;
                            //}

                            //if (string.IsNullOrEmpty(dt_.Rows[i][6].ToString()))
                            //{
                            //    continue;
                            //}
                            //else
                            //{
                            //    _uploadExcel.CompanyId = _uploadExcelBll.GetCompanyIdByName(dt_.Rows[i][6].ToString().Trim(trimElements));
                            //    if (_uploadExcel.CompanyId == 0)
                            //    {
                            //        continue;
                            //    }
                            //}


                            if (string.IsNullOrEmpty(dt_.Rows[i][8].ToString()))
                            {
                                continue;
                            }
                            else
                            {

                                decimal tempUnitPrice = 0;
                                decimal.TryParse(dt_.Rows[i][8].ToString(), out tempUnitPrice);
                                int? gradeId = _uploadExcelBll.GetGradeIdByUnitPrice(dt_.Rows[i][8].ToString().Trim(trimElements));
                                _uploadExcel.GradeId = gradeId == 0 ? null : gradeId;
                                _uploadExcel.UnitPrice = Convert.ToDecimal(dt_.Rows[i][8].ToString().Trim(trimElements));

                                
                            }


                            if (string.IsNullOrEmpty(dt_.Rows[i][5].ToString().Trim(trimElements)))
                            {
                                continue;
                            }
                            else
                            {
                                _uploadExcel.EmployeeName = dt_.Rows[i][5].ToString().Trim(trimElements);
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
                            decimal.TryParse(dt_.Rows[i][9].ToString(), out octInput);

                            decimal novInput = 0;
                            decimal.TryParse(dt_.Rows[i][10].ToString(), out novInput);

                            decimal decInput = 0;
                            decimal.TryParse(dt_.Rows[i][11].ToString(), out decInput);

                            decimal janInput = 0;
                            decimal.TryParse(dt_.Rows[i][12].ToString(), out janInput);

                            decimal febInput = 0;
                            decimal.TryParse(dt_.Rows[i][13].ToString(), out febInput);

                            decimal marInput = 0;
                            decimal.TryParse(dt_.Rows[i][14].ToString(), out marInput);

                            decimal aprInput = 0;
                            decimal.TryParse(dt_.Rows[i][15].ToString(), out aprInput);

                            decimal mayInput = 0;
                            decimal.TryParse(dt_.Rows[i][16].ToString(), out mayInput);

                            decimal junInput = 0;
                            decimal.TryParse(dt_.Rows[i][17].ToString(), out junInput);

                            decimal julInput = 0;
                            decimal.TryParse(dt_.Rows[i][18].ToString(), out julInput);

                            decimal augInput = 0;
                            decimal.TryParse(dt_.Rows[i][19].ToString(), out augInput);

                            decimal septInput = 0;
                            decimal.TryParse(dt_.Rows[i][20].ToString(), out septInput);

                            tempRow = $@"10_{octInput}_{octInput * _uploadExcel.UnitPrice},11_{novInput}_{novInput * _uploadExcel.UnitPrice},12_{decInput}_{decInput * _uploadExcel.UnitPrice},1_{janInput}_{janInput * _uploadExcel.UnitPrice},2_{febInput}_{febInput * _uploadExcel.UnitPrice},3_{marInput}_{marInput * _uploadExcel.UnitPrice},4_{aprInput}_{aprInput * _uploadExcel.UnitPrice},5_{mayInput}_{mayInput * _uploadExcel.UnitPrice},6_{junInput}_{junInput * _uploadExcel.UnitPrice},7_{julInput}_{julInput * _uploadExcel.UnitPrice},8_{augInput}_{augInput * _uploadExcel.UnitPrice},9_{septInput}_{septInput * _uploadExcel.UnitPrice}";

                            SendToApi(tempAssignmentId, tempRow, tempYear);
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
                    return RedirectToAction("CreateForecast");
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
        public void SendToApi(int assignmentId, string row, int year)
        {

            SendToForecaseApiDto sendToForecaseApiDto = new SendToForecaseApiDto();
            sendToForecaseApiDto.Data = row;
            sendToForecaseApiDto.Year = year;
            sendToForecaseApiDto.AssignmentId = assignmentId;

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri("http://198.38.92.119:8081/api/Forecasts?data=" + row + "&year=" + year + "&assignmentId=" + assignmentId);
                //client.BaseAddress = new Uri("http://localhost:59198/api/Forecasts?data=" + row + "&year=" + year + "&assignmentId=" + assignmentId);

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
            //employeeAssignment.InchargeId = Convert.ToInt32(dt_.InchargeId.ToString().Trim(trimElements));
            employeeAssignment.DepartmentId = dt_.DepartmentId;
            //employeeAssignment.RoleId = Convert.ToInt32(dt_.RoleId.ToString().Trim(trimElements));
            employeeAssignment.CompanyId = dt_.CompanyId;
            employeeAssignment.ExplanationId = dt_.ExplanationId==null ? null : dt_.ExplanationId.ToString();
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