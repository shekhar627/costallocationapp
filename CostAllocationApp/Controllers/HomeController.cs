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
                    int fieldcount = reader.FieldCount;
                    int rowcount = reader.RowCount;
                    //DataTable dt = new DataTable();
                    DataRow row;
                    DataTable dt_ = new DataTable();
                    try
                    {
                        int tempAssignmentId = 0;
                        string tempRow = "";
                        int tempYear = 2023;
                        dt_ = reader.AsDataSet().Tables[0];         
                        for (int i = 1; i < rowcount; i++)
                        {
                            tempAssignmentId = CreateAssignmentForExcelUpload(dt_, i);

                            //tempAssignmentId = EmployeeAssignment();
                            //EmployeeAssignmentViewModel employeeAssignment = employeeAssignmentBLL.GetAssignmentById(tempAssignmentId);
                            //tempAssignmentId = Convert.ToInt32(dt_.Rows[i][0]);
                            tempRow += $@"10_{dt_.Rows[i][9].ToString()}_{dt_.Rows[i][21].ToString()},11_{dt_.Rows[i][10].ToString()}_{dt_.Rows[i][22].ToString()},12_{dt_.Rows[i][11].ToString()}_{dt_.Rows[i][23].ToString()},1_{dt_.Rows[i][12].ToString()}_{dt_.Rows[i][24].ToString()},2_{dt_.Rows[i][13].ToString()}_{dt_.Rows[i][25].ToString()},3_{dt_.Rows[i][14].ToString()}_{dt_.Rows[i][26].ToString()},4_{dt_.Rows[i][15].ToString()}_{dt_.Rows[i][27].ToString()},5_{dt_.Rows[i][16].ToString()}_{dt_.Rows[i][28].ToString()},6_{dt_.Rows[i][17].ToString()}_{dt_.Rows[i][29].ToString()},7_{dt_.Rows[i][18].ToString()}_{dt_.Rows[i][30].ToString()},8_{dt_.Rows[i][19].ToString()}_{dt_.Rows[i][31].ToString()},9_{dt_.Rows[i][20].ToString()}_{dt_.Rows[i][31].ToString()}";

                            SendToApi(tempAssignmentId, tempRow, tempYear);
                        }


                    }
                    catch (Exception ex)
                    {
                        ModelState.AddModelError("File", "Unable to Upload file!");
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
                client.BaseAddress = new Uri("http://localhost:59198/api/Forecasts?data="+row+"&year="+year+ "&assignmentId="+assignmentId);

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
        public int CreateAssignmentForExcelUpload(DataTable dt_, int i)
        {
            EmployeeAssignmentDTO employeeAssignmentDTO = new EmployeeAssignmentDTO();
            EmployeeAssignment employeeAssignment = new EmployeeAssignment();

            employeeAssignmentDTO = new EmployeeAssignmentDTO();
            employeeAssignmentDTO.EmployeeName = dt_.Rows[i][5].ToString();
            employeeAssignmentDTO.SectionId = dt_.Rows[i][0].ToString();
            employeeAssignmentDTO.InchargeId = dt_.Rows[i][2].ToString();
            employeeAssignmentDTO.DepartmentId = dt_.Rows[i][1].ToString();
            employeeAssignmentDTO.RoleId = dt_.Rows[i][3].ToString();
            employeeAssignmentDTO.CompanyId = dt_.Rows[i][6].ToString();
            employeeAssignmentDTO.ExplanationId = dt_.Rows[i][4].ToString();
            employeeAssignmentDTO.UnitPrice = dt_.Rows[i][8].ToString();
            employeeAssignmentDTO.GradeId = dt_.Rows[i][7].ToString();
            employeeAssignmentDTO.SubCode = 1;
            //var remarks = $('#memo_new').val();

            int tempValue = 0;
            decimal tempUnitPrice = 0;

            #region validation of inputs
            if (!String.IsNullOrEmpty(employeeAssignmentDTO.EmployeeName))
            {
                var checkResult = employeeAssignmentBLL.CheckEmployeeName(employeeAssignmentDTO.EmployeeName.Trim());
                if (checkResult && employeeAssignmentDTO.SubCode == 1)
                {
                    return 0;
                }
                else
                {
                    employeeAssignment.EmployeeName = employeeAssignmentDTO.EmployeeName.Trim();
                }

            }
            else
            {
                return 0;
            }
            if (int.TryParse(employeeAssignmentDTO.SectionId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return 0;

                }
                employeeAssignment.SectionId = tempValue;
            }
            else
            {
                return 0;
            }
            if (int.TryParse(employeeAssignmentDTO.DepartmentId, out tempValue))
            {
                if (tempValue <= 0)
                {

                    return 0;
                }
                employeeAssignment.DepartmentId = tempValue;
            }
            else
            {
                return 0;
            }
            if (int.TryParse(employeeAssignmentDTO.InchargeId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return 0;

                }
                employeeAssignment.InchargeId = tempValue;
            }
            else
            {
                return 0;
            }
            if (int.TryParse(employeeAssignmentDTO.RoleId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return 0;

                }
                employeeAssignment.RoleId = tempValue;
            }
            else
            {
                return 0;
            }
            if (String.IsNullOrEmpty(employeeAssignmentDTO.ExplanationId))
            {
                employeeAssignment.ExplanationId = null;
            }
            else
            {
                if (int.TryParse(employeeAssignmentDTO.ExplanationId, out tempValue))
                {
                    if (tempValue <= 0)
                    {

                        return 0;
                    }
                    employeeAssignment.ExplanationId = tempValue.ToString();
                }
                else
                {
                    return 0;
                }
            }


            if (int.TryParse(employeeAssignmentDTO.CompanyId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return 0;

                }
                employeeAssignment.CompanyId = tempValue;
            }
            else
            {
                return 0;
            }
            if (decimal.TryParse(employeeAssignmentDTO.UnitPrice, out tempUnitPrice))
            {
                if (tempValue < 0)
                {
                    return 0;

                }
                employeeAssignment.UnitPrice = tempUnitPrice;
            }
            else
            {
                return 0;
            }
            if (int.TryParse(employeeAssignmentDTO.GradeId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return 0;

                }
                employeeAssignment.GradeId = tempValue;
            }
            else
            {
                return 0;
            }

            if (String.IsNullOrEmpty(employeeAssignmentDTO.Remarks))
            {
                employeeAssignmentDTO.Remarks = "";
            }
            #endregion

            //employeeAssignment.ExplanationId = employeeAssignmentDTO.ExplanationId;
            employeeAssignment.CreatedBy = "";
            employeeAssignment.CreatedDate = DateTime.Now;
            employeeAssignment.IsActive = "1";
            employeeAssignment.Remarks = employeeAssignmentDTO.Remarks.Trim();
            employeeAssignment.SubCode = employeeAssignmentDTO.SubCode;


            int result = employeeAssignmentBLL.CreateAssignment(employeeAssignment);
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