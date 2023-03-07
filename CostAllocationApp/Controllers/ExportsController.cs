using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Web;
using System.Web.Mvc;
using CostAllocationApp.BLL;
using CostAllocationApp.ViewModels;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using CostAllocationApp.Dtos;
using System.Data;
using CostAllocationApp.Models;

namespace CostAllocationApp.Controllers
{
    public class ExportsController : Controller
    {
        private DepartmentBLL _departmentBLL = null;
        private ExportBLL _exportBLL = null;
        private SalaryBLL _salaryBLL = null;
        private ExplanationsBLL _explanationsBLL = null;
        private GradeBLL _gradeBll = null;
        private CommonMasterBLL _commonMasterBLL = null;
        private UnitPriceTypeBLL _unitPriceTypeBLL = null;
        private ExportDepartmentAllocationViewModel _exportDepartmentAllocationViewModel = null;
        private Utility _utility = new Utility();

        public ExportsController()
        {
            _departmentBLL = new DepartmentBLL();
            _exportBLL = new ExportBLL();
            _salaryBLL = new SalaryBLL();
            _explanationsBLL = new ExplanationsBLL();
            _gradeBll = new GradeBLL();
            _commonMasterBLL = new CommonMasterBLL();
            _unitPriceTypeBLL = new UnitPriceTypeBLL();
            _exportDepartmentAllocationViewModel = new ExportDepartmentAllocationViewModel();
        }
        // GET: Exports
        public FileResult ExportBySection(int sectionId = 0, string sectionName = "")
        {
            List<ForecastAssignmentViewModel> forecastAssignmentViewModels = new List<ForecastAssignmentViewModel>();
            using (var client = new HttpClient())
            {
                string uri = ""+_utility.Address+"/api/utilities/SearchForecastEmployee?employeeName=&sectionId=" + sectionId + "&departmentId=&inchargeId=&roleId=&explanationId=&companyId=&status=&year=";
                client.BaseAddress = new Uri(uri);
                //HTTP GET
                var responseTask = client.GetAsync("");
                responseTask.Wait();

                var result = responseTask.Result;
                if (result.IsSuccessStatusCode)
                {
                    var readTask = result.Content.ReadAsAsync<List<ForecastAssignmentViewModel>>();
                    readTask.Wait();

                    forecastAssignmentViewModels = readTask.Result;
                }
                else //web api sent error response 
                {
                    //log response status here..

                    forecastAssignmentViewModels = new List<ForecastAssignmentViewModel>();

                    // ModelState.AddModelError(string.Empty, "Server error. Please contact administrator.");
                }
            }
            string htmlTable = "";
            if (forecastAssignmentViewModels.Count > 0)
            {
                int i = 1;
                htmlTable = @"<!DOCTYPE html>
                <html>
                    <head>
                        <style>
                        #dev_placement {
                          border-collapse: collapse;
                          width: 50%;
                        }

                        #dev_placement td, #dev_placement th {
                          border: 1px solid #ddd;
                        }
                        </style>
                    </head>
                <body>

                <table id='dev_placement'>
                <thead>
                  <tr>
                    <th>No</th>
                    <th>FullName</th>
                    <th>Section</th>
                    <th>Department</th>
                    <th>Grade</th>
                    <th>UnitPrice</th>
                    <th>10</th>
                    <th>11</th>
                    <th>12</th>
                    <th>1</th>
                    <th>2</th>
                    <th>3</th>
                    <th>4</th>
                    <th>5</th>
                    <th>6</th>
                    <th>7</th>
                    <th>8</th>
                    <th>9</th>
                    <th>10</th>
                    <th>11</th>
                    <th>12</th>
                    <th>1</th>
                    <th>2</th>
                    <th>3</th>
                    <th>4</th>
                    <th>5</th>
                    <th>6</th>
                    <th>7</th>
                    <th>8</th>
                    <th>9</th>
                  </tr>
                </thead>
                <tbody>";
                foreach (var item in forecastAssignmentViewModels)
                {
                    string td = "";
                    string tdPoints = "";
                    if (item.forecasts.Count > 0)
                    {
                        foreach (var forecastItem in item.forecasts)
                        {
                            if (forecastItem.Points == 0)
                            {
                                td += $@"<td style='background-color:#808080;padding-right:20px;'> </td>";
                            }
                            else
                            {
                                td += $@"<td style='padding-right:20px;'>{forecastItem.Points}</td>";
                            }

                        }
                        foreach (var forecastItem in item.forecasts)
                        {
                            if (forecastItem.Points == 0)
                            {
                                tdPoints += $@"<td style='background-color:#808080;padding-right:20px;'> </td>";
                            }
                            else
                            {
                                //tdPoints += $@"<td style='padding-right:20px;text-align: right;'>{forecastItem.Total}</td>";
                                tdPoints += $@"<td style='text-align: right;'>{forecastItem.Total}</td>";
                            }

                        }
                    }
                    htmlTable += $@"
                          <tr>
                            <td>{i}</td>
                            <td>{item.EmployeeName}</td>
                            <td>{item.SectionName}</td>
                            <td>{item.DepartmentName}</td>
                            <td>{item.GradePoint}</td>
                            <td>{item.UnitPrice}</td>
                            {td}
                            {tdPoints}
                          </tr>
                    ";
                    i++;
                }
                htmlTable += @"
                 </tbody>
               </table>
                </body>
                </html> ";
            }
            UTF8Encoding uTF8Encoding = new UTF8Encoding();
            return File(uTF8Encoding.GetBytes(htmlTable), "application/vnd.ms-excel", sectionName + ".xls");
        }


        public ActionResult DataExportByDepartment()
        {
            return View(new ExportViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }

        [HttpPost]
        public ActionResult DataExportByDepartment(int departmentId = 0)
        {
            List<ForecastAssignmentViewModel> forecastAssignmentViewModels = new List<ForecastAssignmentViewModel>();
            var _department = _departmentBLL.GetDepartmentByDepartemntId(departmentId);

            using (var client = new HttpClient())
            {
                string uri = "" + _utility.Address + "/api/utilities/SearchForecastEmployee?employeeName=&sectionId=&departmentId=" + departmentId + "&inchargeId=&roleId=&explanationId=&companyId=&status=&year=";
                client.BaseAddress = new Uri(uri);
                //HTTP GET
                var responseTask = client.GetAsync("");
                responseTask.Wait();

                var result = responseTask.Result;
                if (result.IsSuccessStatusCode)
                {
                    var readTask = result.Content.ReadAsAsync<List<ForecastAssignmentViewModel>>();
                    readTask.Wait();

                    forecastAssignmentViewModels = readTask.Result;
                }
                else //web api sent error response 
                {
                    //log response status here..

                    forecastAssignmentViewModels = new List<ForecastAssignmentViewModel>();

                    // ModelState.AddModelError(string.Empty, "Server error. Please contact administrator.");
                }
            }

            foreach (var forecastItem in forecastAssignmentViewModels)
            {
                if (forecastItem.CompanyName.ToLower() != "mw" || forecastItem.SectionName.ToLower() != "mw")
                {
                    forecastItem.GradePoint = "";
                }

            }


            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = "Name";
                sheet.Cells["A1"].Style.Font.Bold = true;
                sheet.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["B1"].Value = "Section";
                sheet.Cells["B1"].Style.Font.Bold = true; ;
                sheet.Cells["B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["B1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["C1"].Value = "Company";
                sheet.Cells["C1"].Style.Font.Bold = true;
                sheet.Cells["C1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["C1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


                sheet.Cells["D1"].Value = "Department";
                sheet.Cells["D1"].Style.Font.Bold = true;
                sheet.Cells["D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["D1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


                sheet.Cells["E1"].Value = "Grade";
                sheet.Cells["E1"].Style.Font.Bold = true;
                sheet.Cells["E1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["E1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


                sheet.Cells["F1"].Value = "Unit Price";
                sheet.Cells["F1"].Style.Font.Bold = true;
                sheet.Cells["F1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["F1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["G1"].Value = "10";
                sheet.Cells["G1"].Style.Font.Bold = true;
                sheet.Cells["G1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["G1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["H1"].Value = "11";
                sheet.Cells["H1"].Style.Font.Bold = true;
                sheet.Cells["H1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["H1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["I1"].Value = "12";
                sheet.Cells["I1"].Style.Font.Bold = true;
                sheet.Cells["I1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["I1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["J1"].Value = "1";
                sheet.Cells["J1"].Style.Font.Bold = true;
                sheet.Cells["J1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["J1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["K1"].Value = "2";
                sheet.Cells["K1"].Style.Font.Bold = true;
                sheet.Cells["K1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["K1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["L1"].Value = "3";
                sheet.Cells["L1"].Style.Font.Bold = true;
                sheet.Cells["L1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["L1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["M1"].Value = "4";
                sheet.Cells["M1"].Style.Font.Bold = true;
                sheet.Cells["M1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["M1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["N1"].Value = "5";
                sheet.Cells["N1"].Style.Font.Bold = true;
                sheet.Cells["N1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["N1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["O1"].Value = "6";
                sheet.Cells["O1"].Style.Font.Bold = true;
                sheet.Cells["O1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["O1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["P1"].Value = "7";
                sheet.Cells["P1"].Style.Font.Bold = true;
                sheet.Cells["P1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["P1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["Q1"].Value = "8";
                sheet.Cells["Q1"].Style.Font.Bold = true;
                sheet.Cells["Q1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["Q1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["R1"].Value = "9";
                sheet.Cells["R1"].Style.Font.Bold = true;
                sheet.Cells["R1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["R1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


                // OUTPUT...
                sheet.Cells["S1"].Value = "10";
                sheet.Cells["S1"].Style.Font.Bold = true;
                sheet.Cells["S1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["S1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["T1"].Value = "11";
                sheet.Cells["T1"].Style.Font.Bold = true;
                sheet.Cells["T1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["T1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["U1"].Value = "12";
                sheet.Cells["U1"].Style.Font.Bold = true;
                sheet.Cells["U1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["U1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["V1"].Value = "1";
                sheet.Cells["V1"].Style.Font.Bold = true;
                sheet.Cells["V1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["V1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["W1"].Value = "2";
                sheet.Cells["W1"].Style.Font.Bold = true;
                sheet.Cells["W1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["W1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["X1"].Value = "3";
                sheet.Cells["X1"].Style.Font.Bold = true;
                sheet.Cells["X1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["X1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["Y1"].Value = "4";
                sheet.Cells["Y1"].Style.Font.Bold = true;
                sheet.Cells["Y1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["Y1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["Z1"].Value = "5";
                sheet.Cells["Z1"].Style.Font.Bold = true;
                sheet.Cells["Z1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["Z1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["AA1"].Value = "6";
                sheet.Cells["AA1"].Style.Font.Bold = true;
                sheet.Cells["AA1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["AA1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


                sheet.Cells["AB1"].Value = "7";
                sheet.Cells["AB1"].Style.Font.Bold = true;
                sheet.Cells["AB1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["AB1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["AC1"].Value = "8";
                sheet.Cells["AC1"].Style.Font.Bold = true;
                sheet.Cells["AC1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["AC1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells["AD1"].Value = "9";
                sheet.Cells["AD1"].Style.Font.Bold = true;
                sheet.Cells["AD1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["AD1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                if (forecastAssignmentViewModels.Count > 0)
                {
                    int count = 2;
                    foreach (var item in forecastAssignmentViewModels)
                    {
                        sheet.Cells["A" + count].Value = item.EmployeeName;
                        sheet.Cells["B" + count].Value = item.SectionName;
                        sheet.Cells["C" + count].Value = item.CompanyName;
                        sheet.Cells["D" + count].Value = item.DepartmentName;
                        sheet.Cells["E" + count].Value = item.GradePoint;
                        sheet.Cells["F" + count].Value = Convert.ToDecimal(item.UnitPrice).ToString("N0");
                        sheet.Cells["F" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        //string[] cellNumber = new string[12]{"G","H","I","J","K","L","M","N","O","P","Q","R"};

                        if (item.forecasts.Count > 0)
                        {
                            for (int i = 0; i < item.forecasts.Count; i++)
                            {
                                sheet.Cells["G" + count].Value = item.forecasts[i].Points;
                                sheet.Cells["H" + count].Value = item.forecasts[i].Points;
                                sheet.Cells["I" + count].Value = item.forecasts[i].Points;
                                sheet.Cells["J" + count].Value = item.forecasts[i].Points;
                                sheet.Cells["K" + count].Value = item.forecasts[i].Points;
                                sheet.Cells["L" + count].Value = item.forecasts[i].Points;
                                sheet.Cells["M" + count].Value = item.forecasts[i].Points;
                                sheet.Cells["N" + count].Value = item.forecasts[i].Points;
                                sheet.Cells["O" + count].Value = item.forecasts[i].Points;
                                sheet.Cells["P" + count].Value = item.forecasts[i].Points;
                                sheet.Cells["Q" + count].Value = item.forecasts[i].Points;
                                sheet.Cells["R" + count].Value = item.forecasts[i].Points;
                            }

                            for (int i = 0; i < item.forecasts.Count; i++)
                            {
                                sheet.Cells["S" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                                sheet.Cells["S" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                sheet.Cells["T" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                                sheet.Cells["T" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                sheet.Cells["U" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                                sheet.Cells["U" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                sheet.Cells["V" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                                sheet.Cells["V" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                sheet.Cells["W" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                                sheet.Cells["W" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                sheet.Cells["X" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                                sheet.Cells["X" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                sheet.Cells["Y" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                                sheet.Cells["Y" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                sheet.Cells["Z" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                                sheet.Cells["Z" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                sheet.Cells["AA" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                                sheet.Cells["AA" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                sheet.Cells["AB" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                                sheet.Cells["AB" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                sheet.Cells["AC" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                                sheet.Cells["AC" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                sheet.Cells["AD" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                                sheet.Cells["AD" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            }
                        }
                        count++;
                    }
                }


                var excelData = package.GetAsByteArray();
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                var fileName = _department.DepartmentName + ".xlsx";
                return File(excelData, contentType, fileName);
            }
        }

        public ActionResult DataExportByAllocation()
        {
            return View(new ExportViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }

        public ActionResult ExportAllData()
        {
            var _departments = _departmentBLL.GetAllDepartments();
            using (var package = new ExcelPackage())
            {
                if (_departments.Count > 0)
                {
                    foreach (var item in _departments)
                    {
                        GetDepartmentWiseData(package, item);
                    }
                }
                if (_departments.Count > 0)
                {
                    foreach (var department in _departments)
                    {
                        var explanationList = _explanationsBLL.GetAllExplanationsByDepartmentId(department.Id);
                        if (explanationList.Count > 0)
                        {
                            foreach (var explanation in explanationList)
                            {
                                GetAllocationWiseData(package, department, explanation);
                            }
                        }
                    }
                }

                // get salary master
                GetSalaryMasterByYear(package, 2022);
                // get common master
                GetCommonMaster(package);

                var excelData = package.GetAsByteArray();
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                var fileName = "all_data.xlsx";
                return File(excelData, contentType, fileName);
            }

        }

        private void GetDepartmentWiseData(ExcelPackage excelPackage, Department department)
        {
            List<ForecastAssignmentViewModel> forecastAssignmentViewModels = new List<ForecastAssignmentViewModel>();
            var _department = _departmentBLL.GetDepartmentByDepartemntId(department.Id);

            // get data from API
            using (var client = new HttpClient())
            {
                string uri = "" + _utility.Address + "/api/utilities/SearchForecastEmployee?employeeName=&sectionId=&departmentId=" + department.Id + "&inchargeId=&roleId=&explanationId=&companyId=&status=&year=";
                client.BaseAddress = new Uri(uri);
                //HTTP GET
                var responseTask = client.GetAsync("");
                responseTask.Wait();

                var result = responseTask.Result;
                if (result.IsSuccessStatusCode)
                {
                    var readTask = result.Content.ReadAsAsync<List<ForecastAssignmentViewModel>>();
                    readTask.Wait();

                    forecastAssignmentViewModels = readTask.Result;
                }
                else //web api sent error response 
                {
                    //log response status here..

                    forecastAssignmentViewModels = new List<ForecastAssignmentViewModel>();

                    // ModelState.AddModelError(string.Empty, "Server error. Please contact administrator.");
                }
            }

            // check mw
            foreach (var forecastItem in forecastAssignmentViewModels)
            {
                if (forecastItem.CompanyName.ToLower() != "mw" || forecastItem.SectionName.ToLower() != "mw")
                {
                    forecastItem.GradePoint = "";
                }

            }

            // fill the sheet
            var sheet = excelPackage.Workbook.Worksheets.Add(department.DepartmentName);
            sheet.TabColor = Color.Red;

            sheet.Cells["A1"].Value = "Name";
            sheet.Cells["A1"].Style.Font.Bold = true;
            sheet.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["B1"].Value = "Section";
            sheet.Cells["B1"].Style.Font.Bold = true; ;
            sheet.Cells["B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["B1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["C1"].Value = "Company";
            sheet.Cells["C1"].Style.Font.Bold = true;
            sheet.Cells["C1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["C1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


            sheet.Cells["D1"].Value = "Department";
            sheet.Cells["D1"].Style.Font.Bold = true;
            sheet.Cells["D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["D1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


            sheet.Cells["E1"].Value = "Grade";
            sheet.Cells["E1"].Style.Font.Bold = true;
            sheet.Cells["E1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["E1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


            sheet.Cells["F1"].Value = "Unit Price";
            sheet.Cells["F1"].Style.Font.Bold = true;
            sheet.Cells["F1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["F1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["G1"].Value = "10";
            sheet.Cells["G1"].Style.Font.Bold = true;
            sheet.Cells["G1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["G1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["H1"].Value = "11";
            sheet.Cells["H1"].Style.Font.Bold = true;
            sheet.Cells["H1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["H1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["I1"].Value = "12";
            sheet.Cells["I1"].Style.Font.Bold = true;
            sheet.Cells["I1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["I1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["J1"].Value = "1";
            sheet.Cells["J1"].Style.Font.Bold = true;
            sheet.Cells["J1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["J1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["K1"].Value = "2";
            sheet.Cells["K1"].Style.Font.Bold = true;
            sheet.Cells["K1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["K1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["L1"].Value = "3";
            sheet.Cells["L1"].Style.Font.Bold = true;
            sheet.Cells["L1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["L1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["M1"].Value = "4";
            sheet.Cells["M1"].Style.Font.Bold = true;
            sheet.Cells["M1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["M1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["N1"].Value = "5";
            sheet.Cells["N1"].Style.Font.Bold = true;
            sheet.Cells["N1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["N1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["O1"].Value = "6";
            sheet.Cells["O1"].Style.Font.Bold = true;
            sheet.Cells["O1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["O1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["P1"].Value = "7";
            sheet.Cells["P1"].Style.Font.Bold = true;
            sheet.Cells["P1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["P1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["Q1"].Value = "8";
            sheet.Cells["Q1"].Style.Font.Bold = true;
            sheet.Cells["Q1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["Q1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["R1"].Value = "9";
            sheet.Cells["R1"].Style.Font.Bold = true;
            sheet.Cells["R1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["R1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


            // OUTPUT...
            sheet.Cells["S1"].Value = "10";
            sheet.Cells["S1"].Style.Font.Bold = true;
            sheet.Cells["S1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["S1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["T1"].Value = "11";
            sheet.Cells["T1"].Style.Font.Bold = true;
            sheet.Cells["T1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["T1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["U1"].Value = "12";
            sheet.Cells["U1"].Style.Font.Bold = true;
            sheet.Cells["U1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["U1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["V1"].Value = "1";
            sheet.Cells["V1"].Style.Font.Bold = true;
            sheet.Cells["V1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["V1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["W1"].Value = "2";
            sheet.Cells["W1"].Style.Font.Bold = true;
            sheet.Cells["W1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["W1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["X1"].Value = "3";
            sheet.Cells["X1"].Style.Font.Bold = true;
            sheet.Cells["X1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["X1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["Y1"].Value = "4";
            sheet.Cells["Y1"].Style.Font.Bold = true;
            sheet.Cells["Y1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["Y1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["Z1"].Value = "5";
            sheet.Cells["Z1"].Style.Font.Bold = true;
            sheet.Cells["Z1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["Z1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["AA1"].Value = "6";
            sheet.Cells["AA1"].Style.Font.Bold = true;
            sheet.Cells["AA1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["AA1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


            sheet.Cells["AB1"].Value = "7";
            sheet.Cells["AB1"].Style.Font.Bold = true;
            sheet.Cells["AB1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["AB1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["AC1"].Value = "8";
            sheet.Cells["AC1"].Style.Font.Bold = true;
            sheet.Cells["AC1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["AC1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            sheet.Cells["AD1"].Value = "9";
            sheet.Cells["AD1"].Style.Font.Bold = true;
            sheet.Cells["AD1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["AD1"].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

            if (forecastAssignmentViewModels.Count > 0)
            {
                int count = 2;
                foreach (var item in forecastAssignmentViewModels)
                {
                    sheet.Cells["A" + count].Value = item.EmployeeName;
                    sheet.Cells["B" + count].Value = item.SectionName;
                    sheet.Cells["C" + count].Value = item.CompanyName;
                    sheet.Cells["D" + count].Value = item.DepartmentName;
                    sheet.Cells["E" + count].Value = item.GradePoint;
                    sheet.Cells["F" + count].Value = String.IsNullOrEmpty(item.UnitPrice) == true ? "" : Convert.ToDecimal(item.UnitPrice).ToString("N0");
                    sheet.Cells["F" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    //string[] cellNumber = new string[12]{"G","H","I","J","K","L","M","N","O","P","Q","R"};

                    if (item.forecasts.Count > 0)
                    {
                        for (int i = 0; i < item.forecasts.Count; i++)
                        {
                            sheet.Cells["G" + count].Value = item.forecasts[i].Points;
                            sheet.Cells["H" + count].Value = item.forecasts[i].Points;
                            sheet.Cells["I" + count].Value = item.forecasts[i].Points;
                            sheet.Cells["J" + count].Value = item.forecasts[i].Points;
                            sheet.Cells["K" + count].Value = item.forecasts[i].Points;
                            sheet.Cells["L" + count].Value = item.forecasts[i].Points;
                            sheet.Cells["M" + count].Value = item.forecasts[i].Points;
                            sheet.Cells["N" + count].Value = item.forecasts[i].Points;
                            sheet.Cells["O" + count].Value = item.forecasts[i].Points;
                            sheet.Cells["P" + count].Value = item.forecasts[i].Points;
                            sheet.Cells["Q" + count].Value = item.forecasts[i].Points;
                            sheet.Cells["R" + count].Value = item.forecasts[i].Points;
                        }

                        for (int i = 0; i < item.forecasts.Count; i++)
                        {
                            sheet.Cells["S" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                            sheet.Cells["S" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells["T" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                            sheet.Cells["T" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells["U" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                            sheet.Cells["U" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells["V" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                            sheet.Cells["V" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells["W" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                            sheet.Cells["W" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells["X" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                            sheet.Cells["X" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells["Y" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                            sheet.Cells["Y" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells["Z" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                            sheet.Cells["Z" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells["AA" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                            sheet.Cells["AA" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells["AB" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                            sheet.Cells["AB" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells["AC" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                            sheet.Cells["AC" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells["AD" + count].Value = Convert.ToDecimal(item.forecasts[i].Total).ToString("N0");
                            sheet.Cells["AD" + count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        }
                    }
                    count++;
                }
            }


        }

        private void GetAllocationWiseData(ExcelPackage excelPackage, Department department, Explanation explanation)
        {
            List<Grade> grades = _gradeBll.GetAllGrade();

            List<SalaryAssignmentDto> salaryAssignmentDtos = new List<SalaryAssignmentDto>();
            //var department = _departmentBLL.GetDepartmentByDepartemntId(departmentId);
            //var explanation = _explanationsBLL.GetExplanationByExplanationId(explanationId);
            var salaries = _salaryBLL.GetAllSalaryPoints();

            // get all data with department and allocation
            List<ForecastAssignmentViewModel> assignmentsWithForecast = _exportBLL.AssignmentsByAllocation(department.Id, explanation.Id);
            //filtered by grade id
            List<ForecastAssignmentViewModel> assignmentsWithGrade = assignmentsWithForecast.Where(a => a.GradeId != null).ToList();
            //filtered by section and company
            List<ForecastAssignmentViewModel> assignmentsWithSectionAndCompany = assignmentsWithForecast.Where(a => a.SectionId != null && a.CompanyId != null).ToList();


            foreach (var item in grades)
            {
                List<ForecastAssignmentViewModel> forecastAssignmentViews = new List<ForecastAssignmentViewModel>();

                SalaryAssignmentDto salaryAssignmentDto = new SalaryAssignmentDto();
                salaryAssignmentDto.Grade = item;

                var gradeSalaryTypes = _exportBLL.GetGradeUnitPriceTypes(item.Id, department.Id, 2022, 1);

                foreach (var gradeSalary in gradeSalaryTypes)
                {
                    List<ForecastAssignmentViewModel> filteredAssignmentsByGradeSalaryTypeId = assignmentsWithGrade.Where(a => a.GradeId == gradeSalary.GradeId.ToString()).ToList();
                    forecastAssignmentViews.AddRange(filteredAssignmentsByGradeSalaryTypeId);
                }

                salaryAssignmentDto.ForecastAssignmentViewModels = forecastAssignmentViews;
                salaryAssignmentDtos.Add(salaryAssignmentDto);

            }

            var sheet = excelPackage.Workbook.Worksheets.Add("【" + department.DepartmentName + "】" + explanation.ExplanationName);
            if (department.DepartmentName == "開発")
            {
                sheet.TabColor = Color.LightGreen;
            }
            if (department.DepartmentName == "企画")
            {
                sheet.TabColor = Color.Yellow;
            }

            #region header column
            //row-1
            sheet.Cells[1, 1].Value = "導入";
            sheet.Cells[1, 1].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

            sheet.Cells[1, 2].Value = department.DepartmentName;
            sheet.Cells[1, 2].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells[1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

            //row-2
            sheet.Cells[2, 1].Value = "FY2022";

            //row-3
            sheet.Cells[3, 1].Value = "";
            sheet.Cells[3, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 1].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[3, 2].Value = "";
            sheet.Cells[3, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 2].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[3, 3].Value = "FY2022";
            sheet.Cells[3, 3].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[3, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 3].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[3, 4].Value = "FY2022";
            sheet.Cells[3, 4].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[3, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 4].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[3, 5].Value = "FY2022";
            sheet.Cells[3, 5].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[3, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 5].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[3, 6].Value = "FY2022";
            sheet.Cells[3, 6].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[3, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 6].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[3, 7].Value = "FY2022";
            sheet.Cells[3, 7].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[3, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 7].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[3, 8].Value = "FY2022";
            sheet.Cells[3, 8].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[3, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 8].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[3, 9].Value = "FY2022";
            sheet.Cells[3, 9].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[3, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 9].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[3, 10].Value = "FY2022";
            sheet.Cells[3, 10].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[3, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 10].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[3, 11].Value = "FY2022";
            sheet.Cells[3, 11].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[3, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 11].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[3, 12].Value = "FY2022";
            sheet.Cells[3, 12].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[3, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 12].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[3, 13].Value = "FY2022";
            sheet.Cells[3, 13].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[3, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 13].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[3, 14].Value = "FY2022";
            sheet.Cells[3, 14].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[3, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[3, 14].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            // row -4

            sheet.Cells[4, 1].Value = "";
            sheet.Cells[4, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 1].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[4, 2].Value = "";
            sheet.Cells[4, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 2].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[4, 3].Value = "10月";
            sheet.Cells[4, 3].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[4, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 3].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[4, 4].Value = "11月";
            sheet.Cells[4, 4].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[4, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 4].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[4, 5].Value = "12月";
            sheet.Cells[4, 5].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[4, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 5].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[4, 6].Value = "1月";
            sheet.Cells[4, 6].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[4, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 6].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[4, 7].Value = "2月";
            sheet.Cells[4, 7].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[4, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 7].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[4, 8].Value = "3月";
            sheet.Cells[4, 8].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[4, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 8].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[4, 9].Value = "4月";
            sheet.Cells[4, 9].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[4, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 9].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[4, 10].Value = "5月";
            sheet.Cells[4, 10].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[4, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 10].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[4, 11].Value = "6月";
            sheet.Cells[4, 11].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[4, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 11].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[4, 12].Value = "7月";
            sheet.Cells[4, 12].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[4, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 12].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[4, 13].Value = "8月";
            sheet.Cells[4, 13].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[4, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 13].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheet.Cells[4, 14].Value = "9月";
            sheet.Cells[4, 14].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[4, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[4, 14].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            #endregion

            int rowCount = 5;

            sheet.Cells[rowCount, 1].Value = "人数";

            decimal oct = 0;
            decimal nov = 0;
            decimal dec = 0;
            decimal jan = 0;
            decimal feb = 0;
            decimal mar = 0;
            decimal apr = 0;
            decimal may = 0;
            decimal jun = 0;
            decimal jul = 0;
            decimal aug = 0;
            decimal sep = 0;


            decimal octOverTime = 0;
            decimal novOverTime = 0;
            decimal decOverTime = 0;
            decimal janOverTime = 0;
            decimal febOverTime = 0;
            decimal marOverTime = 0;
            decimal aprOverTime = 0;
            decimal mayOverTime = 0;
            decimal junOverTime = 0;
            decimal julOverTime = 0;
            decimal augOverTime = 0;
            decimal sepOverTime = 0;

            decimal octTotal = 0;
            decimal novTotal = 0;
            decimal decTotal = 0;
            decimal janTotal = 0;
            decimal febTotal = 0;
            decimal marTotal = 0;
            decimal aprTotal = 0;
            decimal mayTotal = 0;
            decimal junTotal = 0;
            decimal julTotal = 0;
            decimal augTotal = 0;
            decimal sepTotal = 0;

            //man month point
            foreach (var item in salaryAssignmentDtos)
            {
                sheet.Cells[rowCount, 2].Value = item.Grade.GradeName;



                foreach (var singleAssignment in item.ForecastAssignmentViewModels)
                {
                    oct += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 10).Points);
                    nov += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 11).Points);
                    dec += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 12).Points);
                    jan += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 1).Points);
                    feb += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 2).Points);
                    mar += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 3).Points);
                    apr += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 4).Points);
                    may += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 5).Points);
                    jun += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 6).Points);
                    jul += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 7).Points);
                    aug += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 8).Points);
                    sep += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 9).Points);
                }
                sheet.Cells[rowCount, 3].Value = oct;
                octTotal += oct;
                sheet.Cells[rowCount, 4].Value = nov;
                novTotal += nov;
                sheet.Cells[rowCount, 5].Value = dec;
                decTotal += dec;
                sheet.Cells[rowCount, 6].Value = jan;
                janTotal += jan;
                sheet.Cells[rowCount, 7].Value = feb;
                febTotal += feb;
                sheet.Cells[rowCount, 8].Value = mar;
                marTotal += mar;
                sheet.Cells[rowCount, 9].Value = apr;
                aprTotal += apr;
                sheet.Cells[rowCount, 10].Value = may;
                mayTotal += may;
                sheet.Cells[rowCount, 11].Value = jun;
                junTotal += jun;
                sheet.Cells[rowCount, 12].Value = jul;
                julTotal += jul;
                sheet.Cells[rowCount, 13].Value = aug;
                augTotal += aug;
                sheet.Cells[rowCount, 14].Value = sep;
                sepTotal += sep;


                //valuesWithGrades.Add(new ValuesWithGrade { GradeId = item.Grade.Id,GradeName=item.Grade.GradeName,Point= octTotal,MonthId=10 });
                oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                rowCount++;
            }
            sheet.Cells[rowCount, 2].Value = "合計";
            sheet.Cells[rowCount, 3].Value = octTotal;
            sheet.Cells[rowCount, 4].Value = novTotal;
            sheet.Cells[rowCount, 5].Value = decTotal;
            sheet.Cells[rowCount, 6].Value = janTotal;
            sheet.Cells[rowCount, 7].Value = febTotal;
            sheet.Cells[rowCount, 8].Value = marTotal;
            sheet.Cells[rowCount, 9].Value = aprTotal;
            sheet.Cells[rowCount, 10].Value = mayTotal;
            sheet.Cells[rowCount, 11].Value = junTotal;
            sheet.Cells[rowCount, 12].Value = julTotal;
            sheet.Cells[rowCount, 13].Value = augTotal;
            sheet.Cells[rowCount, 14].Value = sepTotal;
            rowCount++;

            #region over time
            sheet.Cells[rowCount, 1].Value = "1人あたりの時間外勤務見込 \n (みなし時間（固定時間）\n を含む残業時間 \n を入力してください";

            //over time
            foreach (var item in salaryAssignmentDtos)
            {
                sheet.Cells[rowCount, 2].Value = item.Grade.GradeName;



                foreach (var singleAssignment in item.ForecastAssignmentViewModels)
                {
                    octOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 10).OverTime);
                    novOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 11).OverTime);
                    decOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 12).OverTime);
                    janOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 1).OverTime);
                    febOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 2).OverTime);
                    marOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 3).OverTime);
                    aprOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 4).OverTime);
                    mayOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 5).OverTime);
                    junOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 6).OverTime);
                    julOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 7).OverTime);
                    augOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 8).OverTime);
                    sepOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 9).OverTime);
                }
                sheet.Cells[rowCount, 3].Value = octOverTime;
                
                sheet.Cells[rowCount, 4].Value = novOverTime;
              
                sheet.Cells[rowCount, 5].Value = decOverTime;
              
                sheet.Cells[rowCount, 6].Value = janOverTime;
              
                sheet.Cells[rowCount, 7].Value = febOverTime;
              
                sheet.Cells[rowCount, 8].Value = marOverTime;
              
                sheet.Cells[rowCount, 9].Value = aprOverTime;
              
                sheet.Cells[rowCount, 10].Value = mayOverTime;
              
                sheet.Cells[rowCount, 11].Value = junOverTime;
              
                sheet.Cells[rowCount, 12].Value = julOverTime;
              
                sheet.Cells[rowCount, 13].Value = augOverTime;
              
                sheet.Cells[rowCount, 14].Value = sepOverTime;
              


                //valuesWithGrades.Add(new ValuesWithGrade { GradeId = item.Grade.Id,GradeName=item.Grade.GradeName,Point= octTotal,MonthId=10 });
                octOverTime = 0; novOverTime = 0; decOverTime = 0; janOverTime = 0; febOverTime = 0; marOverTime = 0; aprOverTime = 0; mayOverTime = 0; junOverTime = 0; julOverTime = 0; augOverTime = 0; sepOverTime = 0;
                rowCount++;
            }

            #endregion

            rowCount++;

            int rowCountForFuture = rowCount;
            rowCount += 12;

            #region variables for salary types.

            // 2. regular total
            double regularSalaryTotalOct = 0, regularSalaryTotalNov = 0, regularSalaryTotalDec = 0, regularSalaryTotalJan = 0, regularSalaryTotalFeb = 0, regularSalaryTotalMar = 0, regularSalaryTotalApr = 0, regularSalaryTotalMay = 0, regularSalaryTotalJun = 0, regularSalaryTotalJul = 0, regularSalaryTotalAug = 0, regularSalaryTotalSep = 0;
            // 3. fixed total
            double fixedSalaryTotalOct = 0, fixedSalaryTotalNov = 0, fixedSalaryTotalDec = 0, fixedSalaryTotalJan = 0, fixedSalaryTotalFeb = 0, fixedSalaryTotalMar = 0, fixedSalaryTotalApr = 0, fixedSalaryTotalMay = 0, fixedSalaryTotalJun = 0, fixedSalaryTotalJul = 0, fixedSalaryTotalAug = 0, fixedSalaryTotalSep = 0;
            // 4. over time
            double overTimeSalaryTotalOct = 0, overTimeSalaryTotalNov = 0, overTimeSalaryTotalDec = 0, overTimeSalaryTotalJan = 0, overTimeSalaryTotalFeb = 0, overTimeSalaryTotalMar = 0, overTimeSalaryTotalApr = 0, overTimeSalaryTotalMay = 0, overTimeSalaryTotalJun = 0, overTimeSalaryTotalJul = 0, overTimeSalaryTotalAug = 0, overTimeSalaryTotalSep = 0;
            // 5. total Salary
            double salaryTotalOct = 0, salaryTotalNov = 0, salaryTotalDec = 0, salaryTotalJan = 0, salaryTotalFeb = 0, salaryTotalMar = 0, salaryTotalApr = 0, salaryTotalMay = 0, salaryTotalJun = 0, salaryTotalJul = 0, salaryTotalAug = 0, salaryTotalSep = 0;
            // 6. Miscellaneous Wages
            double mwTotalOct = 0, mwTotalNov = 0, mwTotalDec = 0, mwTotalJan = 0, mwTotalFeb = 0, mwTotalMar = 0, mwTotalApr = 0, mwTotalMay = 0, mwTotalJun = 0, mwTotalJul = 0, mwTotalAug = 0, mwTotalSep = 0;
            // 7. Dispatch Fee
            double dispatchFeeTotalOct = 0, dispatchFeeTotalNov = 0, dispatchFeeTotalDec = 0, dispatchFeeTotalJan = 0, dispatchFeeTotalFeb = 0, dispatchFeeTotalMar = 0, dispatchFeeTotalApr = 0, dispatchFeeTotalMay = 0, dispatchFeeTotalJun = 0, dispatchFeeTotalJul = 0, dispatchFeeTotalAug = 0, dispatchFeeTotalSep = 0;
            // 8. Provision for Employee Bonus
            double employeeBonusTotalOct = 0, employeeBonusTotalNov = 0, employeeBonusTotalDec = 0, employeeBonusTotalJan = 0, employeeBonusTotalFeb = 0, employeeBonusTotalMar = 0, employeeBonusTotalApr = 0, employeeBonusTotalMay = 0, employeeBonusTotalJun = 0, employeeBonusTotalJul = 0, employeeBonusTotalAug = 0, employeeBonusTotalSep = 0;
            // 9. Commuting Expenses
            double commutingExpensesTotalOct = 0, commutingExpensesTotalNov = 0, commutingExpensesTotalDec = 0, commutingExpensesTotalJan = 0, commutingExpensesTotalFeb = 0, commutingExpensesTotalMar = 0, commutingExpensesTotalApr = 0, commutingExpensesTotalMay = 0, commutingExpensesTotalJun = 0, commutingExpensesTotalJul = 0, commutingExpensesTotalAug = 0, commutingExpensesTotalSep = 0;
            // 10. Salary Statutory Welfare Expenses
            double welfareExpensesTotalOct = 0, welfareExpensesTotalNov = 0, welfareExpensesTotalDec = 0, welfareExpensesTotalJan = 0, welfareExpensesTotalFeb = 0, welfareExpensesTotalMar = 0, welfareExpensesTotalApr = 0, welfareExpensesTotalMay = 0, welfareExpensesTotalJun = 0, welfareExpensesTotalJul = 0, welfareExpensesTotalAug = 0, welfareExpensesTotalSep = 0;
            // 11. Provision for Statotory Welfare Expenses for Bonus
            double welfareExpensesBonusTotalOct = 0, welfareExpensesBonusTotalNov = 0, welfareExpensesBonusTotalDec = 0, welfareExpensesBonusTotalJan = 0, welfareExpensesBonusTotalFeb = 0, welfareExpensesBonusTotalMar = 0, welfareExpensesBonusTotalApr = 0, welfareExpensesBonusTotalMay = 0, welfareExpensesBonusTotalJun = 0, welfareExpensesBonusTotalJul = 0, welfareExpensesBonusTotalAug = 0, welfareExpensesBonusTotalSep = 0;
            // 12. Total Statutory Benefites
            double statutoryTotalOct = 0, statutoryTotalNov = 0, statutoryTotalDec = 0, statutoryTotalJan = 0, statutoryTotalFeb = 0, statutoryTotalMar = 0, statutoryTotalApr = 0, statutoryTotalMay = 0, statutoryTotalJun = 0, statutoryTotalJul = 0, statutoryTotalAug = 0, statutoryTotalSep = 0;
            // 13. Total Expenses
            double expensesTotalOct = 0, expensesTotalNov = 0, expensesTotalDec = 0, expensesTotalJan = 0, expensesTotalFeb = 0, expensesTotalMar = 0, expensesTotalApr = 0, expensesTotalMay = 0, expensesTotalJun = 0, expensesTotalJul = 0, expensesTotalAug = 0, expensesTotalSep = 0;

            #endregion




            #region other grade with entities
            rowCount++;
            int count = 1;
            foreach (var item in salaryAssignmentDtos)
            {
                sheet.Cells[rowCount, 1].Value = item.Grade.GradeName;

                if (count % 2 == 0)
                {
                    // even
                    sheet.Cells[rowCount, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 1].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                }
                else
                {
                    // odd
                    sheet.Cells[rowCount, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 1].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                }

                //salary allow regular
                double totalRegularOct = 0, totalRegularNov = 0, totalRegularDec = 0, totalRegularJan = 0, totalRegularFeb = 0, totalRegularMar = 0, totalRegularApr = 0, totalRegularMay = 0, totalRegularJun = 0, totalRegularJul = 0, totalRegularAug = 0, totalRegularSep = 0;
                // salary allow fixed
                double totalFixedOct = 0, totalFixedNov = 0, totalFixedDec = 0, totalFixedJan = 0, totalFixedFeb = 0, totalFixedMar = 0, totalFixedApr = 0, totalFixedMay = 0, totalFixedJun = 0, totalFixedJul = 0, totalFixedAug = 0, totalFixedSep = 0;
                // salary allow overtime
                double totalOverOct = 0, totalOverNov = 0, totalOverDec = 0, totalOverJan = 0, totalOverFeb = 0, totalOverMar = 0, totalOverApr = 0, totalOverMay = 0, totalOverJun = 0, totalOverJul = 0, totalOverAug = 0, totalOverSep = 0;
                // total salary
                double totalSalaryOct = 0, totalSalaryNov = 0, totalSalaryDec = 0, totalSalaryJan = 0, totalSalaryFeb = 0, totalSalaryMar = 0, totalSalaryApr = 0, totalSalaryMay = 0, totalSalaryJun = 0, totalSalaryJul = 0, totalSalaryAug = 0, totalSalarySep = 0;
                // Miscellaneous Wages
                double mWagesOct = 0, mWagesNov = 0, mWagesDec = 0, mWagesJan = 0, mWagesFeb = 0, mWagesMar = 0, mWagesApr = 0, mWagesMay = 0, mWagesJun = 0, mWagesJul = 0, mWagesAug = 0, mWagesSep = 0;
                //dispatch fee
                double dispatchFeeOct = 0, dispatchFeeNov = 0, dispatchFeeDec = 0, dispatchFeeJan = 0, dispatchFeeFeb = 0, dispatchFeeMar = 0, dispatchFeeApr = 0, dispatchFeeMay = 0, dispatchFeeJun = 0, dispatchFeeJul = 0, dispatchFeeAug = 0, dispatchFeeSep = 0;
                // employee bonus
                double employeeBonusOct = 0, employeeBonusNov = 0, employeeBonusDec = 0, employeeBonusJan = 0, employeeBonusFeb = 0, employeeBonusMar = 0, employeeBonusApr = 0, employeeBonusMay = 0, employeeBonusJun = 0, employeeBonusJul = 0, employeeBonusAug = 0, employeeBonusSep = 0;
                //commuting expenses
                double commutingExpensesOct = 0, commutingExpensesNov = 0, commutingExpensesDec = 0, commutingExpensesJan = 0, commutingExpensesFeb = 0, commutingExpensesMar = 0, commutingExpensesApr = 0, commutingExpensesMay = 0, commutingExpensesJun = 0, commutingExpensesJul = 0, commutingExpensesAug = 0, commutingExpensesSep = 0;
                // welfare expenses
                double wExpensesOct = 0, wExpensesNov = 0, wExpensesDec = 0, wExpensesJan = 0, wExpensesFeb = 0, wExpensesMar = 0, wExpensesApr = 0, wExpensesMay = 0, wExpensesJun = 0, wExpensesJul = 0, wExpensesAug = 0, wExpensesSep = 0;
                // welfare expenses bonuses
                double wExpBonusOct = 0, wExpBonusNov = 0, wExpBonusDec = 0, wExpBonusJan = 0, wExpBonusFeb = 0, wExpBonusMar = 0, wExpBonusApr = 0, wExpBonusMay = 0, wExpBonusJun = 0, wExpBonusJul = 0, wExpBonusAug = 0, wExpBonusSep = 0;
                // total statutory
                double totalStatutoryOct = 0, totalStatutoryNov = 0, totalStatutoryDec = 0, totalStatutoryJan = 0, totalStatutoryFeb = 0, totalStatutoryMar = 0, totalStatutoryApr = 0, totalStatutoryMay = 0, totalStatutoryJun = 0, totalStatutoryJul = 0, totalStatutoryAug = 0, totalStatutorySep = 0;

                int innerCount = 1;
                int salaryTypeCount = 1;
                foreach (var salaryType in _unitPriceTypeBLL.GetAllSalaryTypes())
                {
                    if (innerCount > 1)
                    {
                        sheet.Cells[rowCount, 1].Value = "";

                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 1].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 1].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                    }

                    sheet.Cells[rowCount, 2].Value = salaryType.SalaryTypeName;
                    if (count % 2 == 0)
                    {
                        // even
                        sheet.Cells[rowCount, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCount, 2].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                    }
                    else
                    {
                        // odd
                        sheet.Cells[rowCount, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCount, 2].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                    }


                    // alligned with the serial of salary type

                    // executives
                    if (salaryTypeCount == 1)
                    {
                        sheet.Cells[rowCount, 3].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 4].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 5].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 6].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 7].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 8].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 9].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 10].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 11].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 12].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 13].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 14].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                    }
                    // salary allowance (regular)
                    if (salaryTypeCount == 2)
                    {


                            double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                            double beginningValue = _salaryBLL.GetGradeUnitPriceType(department.Id,1, 2022, item.Grade.Id).GradeLowPoints;
                            double commonValue = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).SalaryIncreaseRate);

                        sheet.Cells[rowCount, 3].Value = (manpointOct * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalRegularOct += manpointOct * beginningValue * commonValue;
                        regularSalaryTotalOct += totalRegularOct;

                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                        sheet.Cells[rowCount, 4].Value = (manpointNov * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalRegularNov += manpointNov * beginningValue * commonValue;
                        regularSalaryTotalNov += totalRegularNov;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                        sheet.Cells[rowCount, 5].Value = (manpointDec * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalRegularDec += manpointDec * beginningValue * commonValue;
                        regularSalaryTotalDec += totalRegularDec;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                        sheet.Cells[rowCount, 6].Value = (manpointJan * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalRegularJan += manpointJan * beginningValue * commonValue;
                        regularSalaryTotalJan += totalRegularJan;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                        sheet.Cells[rowCount, 7].Value = (manpointFeb * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalRegularFeb += manpointFeb * beginningValue * commonValue;
                        regularSalaryTotalFeb += totalRegularFeb;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                        sheet.Cells[rowCount, 8].Value = (manpointMar * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalRegularMar += manpointMar * beginningValue * commonValue;
                        regularSalaryTotalMar += totalRegularMar;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                        sheet.Cells[rowCount, 9].Value = (manpointApr * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalRegularApr += manpointApr * beginningValue * commonValue;
                        regularSalaryTotalApr += totalRegularApr;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                        sheet.Cells[rowCount, 10].Value = (manpointMay * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalRegularMay += manpointMay * beginningValue * commonValue;
                        regularSalaryTotalMay += totalRegularMay;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                        sheet.Cells[rowCount, 11].Value = (manpointJun * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalRegularJun += manpointJun * beginningValue * commonValue;
                        regularSalaryTotalJun += totalRegularJun;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }



                        double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                        sheet.Cells[rowCount, 12].Value = (manpointJul * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalRegularJul += manpointJul * beginningValue * commonValue;
                        regularSalaryTotalJul += totalRegularJul;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                        sheet.Cells[rowCount, 13].Value = (manpointAug * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalRegularAug += manpointAug * beginningValue * commonValue;
                        regularSalaryTotalAug += totalRegularAug;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                        sheet.Cells[rowCount, 14].Value = (manpointSep * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalRegularSep += manpointSep * beginningValue * commonValue;
                        regularSalaryTotalSep += totalRegularSep;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                    }
                    // salary allowance (fixed)
                    if (salaryTypeCount == 3)
                    {


                            double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                            double beginningValue = _salaryBLL.GetGradeUnitPriceType(department.Id, 2, 2022, item.Grade.Id).GradeLowPoints;
                            double commonValue = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).SalaryIncreaseRate);

                        sheet.Cells[rowCount, 3].Value = (manpointOct * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalFixedOct += manpointOct * beginningValue * commonValue;
                        fixedSalaryTotalOct += totalFixedOct;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                        sheet.Cells[rowCount, 4].Value = (manpointNov * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalFixedNov += manpointNov * beginningValue * commonValue;
                        fixedSalaryTotalNov += totalFixedNov;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                        sheet.Cells[rowCount, 5].Value = (manpointDec * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalFixedDec += manpointDec * beginningValue * commonValue;
                        fixedSalaryTotalDec += totalFixedDec;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                        sheet.Cells[rowCount, 6].Value = (manpointJan * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalFixedJan += manpointJan * beginningValue * commonValue;
                        fixedSalaryTotalJan += totalFixedJan;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                        sheet.Cells[rowCount, 7].Value = (manpointFeb * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalFixedFeb += manpointFeb * beginningValue * commonValue;
                        fixedSalaryTotalFeb += totalFixedFeb;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                        sheet.Cells[rowCount, 8].Value = (manpointMar * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalFixedMar += manpointMar * beginningValue * commonValue;
                        fixedSalaryTotalMar += totalFixedMar;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                        sheet.Cells[rowCount, 9].Value = (manpointApr * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalFixedApr += manpointApr * beginningValue * commonValue;
                        fixedSalaryTotalApr += totalFixedApr;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                        sheet.Cells[rowCount, 10].Value = (manpointMay * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalFixedMay += manpointMay * beginningValue * commonValue;
                        fixedSalaryTotalMay += totalFixedMay;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                        sheet.Cells[rowCount, 11].Value = (manpointJun * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalFixedJun += manpointJun * beginningValue * commonValue;
                        fixedSalaryTotalJun += totalFixedJun;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }



                        double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                        sheet.Cells[rowCount, 12].Value = (manpointJul * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalFixedJul += manpointJul * beginningValue * commonValue;
                        fixedSalaryTotalJul += totalFixedJul;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                        sheet.Cells[rowCount, 13].Value = (manpointAug * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalFixedAug += manpointAug * beginningValue * commonValue;
                        fixedSalaryTotalAug += totalFixedAug;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                        sheet.Cells[rowCount, 14].Value = (manpointSep * beginningValue * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalFixedSep += manpointSep * beginningValue * commonValue;
                        fixedSalaryTotalSep += totalFixedSep;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                    }
                    // salary allowance (overtime)
                    if (salaryTypeCount == 4)
                    {
                        CommonMaster commonMaster = _commonMasterBLL.GetCommonMasters().Where(cm => cm.GradeId == item.Grade.Id).SingleOrDefault();
                        double salaryMasterOverTimeLowPoints = _exportBLL.GetGradeUnitPriceTypes(item.Grade.Id, department.Id, 2022, 3).SingleOrDefault().GradeLowPoints;

                        double calculatedOverTimeOct = GetTotalOverTime(salaryAssignmentDtos, 10, item.Grade.Id);
                        double manPointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                        //double overTimeOct = 0;
                        if (calculatedOverTimeOct> Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverOct = manPointOct*(calculatedOverTimeOct- Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints* Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalOct += totalOverOct;

                        sheet.Cells[rowCount, 3].Value = totalOverOct.ToString("N0");
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double calculatedOverTimeNov = GetTotalOverTime(salaryAssignmentDtos, 11, item.Grade.Id);
                        double manPointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                        
                        if (calculatedOverTimeNov > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverNov = manPointNov * (calculatedOverTimeNov - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalNov += totalOverNov;

                        sheet.Cells[rowCount, 4].Value = totalOverNov.ToString("N0");
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double calculatedOverTimeDec = GetTotalOverTime(salaryAssignmentDtos, 12, item.Grade.Id);
                        double manPointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                        
                        if (calculatedOverTimeDec > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverDec = manPointDec * (calculatedOverTimeDec - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalDec += totalOverDec;

                        sheet.Cells[rowCount, 5].Value = totalOverDec.ToString("N0");
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double calculatedOverTimeJan = GetTotalOverTime(salaryAssignmentDtos, 1, item.Grade.Id);
                        double manPointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                        if (calculatedOverTimeJan > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverJan = manPointJan * (calculatedOverTimeJan - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalJan += totalOverJan;
                        sheet.Cells[rowCount, 6].Value = totalOverJan.ToString("N0");
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double calculatedOverTimeFeb = GetTotalOverTime(salaryAssignmentDtos, 2, item.Grade.Id);
                        double manPointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                        
                        if (calculatedOverTimeFeb > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverFeb = manPointFeb * (calculatedOverTimeFeb - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalFeb += totalOverFeb;
                        sheet.Cells[rowCount, 7].Value = totalOverFeb.ToString("N0");
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double calculatedOverTimeMar = GetTotalOverTime(salaryAssignmentDtos, 3, item.Grade.Id);
                        double manPointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                        
                        if (calculatedOverTimeMar > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverMar = manPointMar * (calculatedOverTimeMar - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalMar += totalOverMar;
                        sheet.Cells[rowCount, 8].Value = totalOverMar.ToString("N0");
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double calculatedOverTimeApr = GetTotalOverTime(salaryAssignmentDtos, 4, item.Grade.Id);
                        double manPointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                        
                        if (calculatedOverTimeApr > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverApr = manPointApr * (calculatedOverTimeApr - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalApr += totalOverApr;
                        sheet.Cells[rowCount, 9].Value = totalOverApr.ToString("N0");
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double calculatedOverTimeMay = GetTotalOverTime(salaryAssignmentDtos, 5, item.Grade.Id);
                        double manPointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                        
                        if (calculatedOverTimeMay > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverMay = manPointMay * (calculatedOverTimeMay - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalMay += totalOverMay;
                        sheet.Cells[rowCount, 10].Value = totalOverMay.ToString("N0");
                        sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double calculatedOverTimeJun = GetTotalOverTime(salaryAssignmentDtos, 6, item.Grade.Id);
                        double manPointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                        
                        if (calculatedOverTimeJun > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverJun = manPointJun * (calculatedOverTimeJun - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalJun += totalOverJun;
                        sheet.Cells[rowCount, 11].Value = totalOverJun.ToString("N0");
                        sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double calculatedOverTimeJul = GetTotalOverTime(salaryAssignmentDtos, 7, item.Grade.Id);
                        double manPointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                        if (calculatedOverTimeJul > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverJul = manPointJul * (calculatedOverTimeJul - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalJul += totalOverJul;
                        sheet.Cells[rowCount, 12].Value = totalOverJul.ToString("N0");
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double calculatedOverTimeAug = GetTotalOverTime(salaryAssignmentDtos, 8, item.Grade.Id);
                        double manPointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                       
                        if (calculatedOverTimeAug > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverAug = manPointAug * (calculatedOverTimeAug - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalAug += totalOverAug;
                        sheet.Cells[rowCount, 13].Value = totalOverAug.ToString("N0");
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        double calculatedOverTimeSept = GetTotalOverTime(salaryAssignmentDtos, 9, item.Grade.Id);
                        double manPointSept = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                        
                        if (calculatedOverTimeSept > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverSep = manPointSept * (calculatedOverTimeSept - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalSep += totalOverSep;

                        sheet.Cells[rowCount, 14].Value = totalOverSep.ToString("N0");
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                    }
                    // salary allowance (total)
                    if (salaryTypeCount == 5)
                    {
                        sheet.Cells[rowCount, 3].Value = (totalRegularOct + totalFixedOct + totalOverOct).ToString("N0");
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalSalaryOct += totalRegularOct + totalFixedOct + totalOverOct;
                        salaryTotalOct += totalSalaryOct;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 4].Value = (totalRegularNov + totalFixedNov + totalOverNov).ToString("N0");
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalSalaryNov += totalRegularNov + totalFixedNov + totalOverNov;
                        salaryTotalNov += totalSalaryNov;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 5].Value = (totalRegularDec + totalFixedDec + totalOverDec).ToString("N0");
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalSalaryDec += totalRegularDec + totalFixedDec + totalOverDec;
                        salaryTotalDec += totalSalaryDec;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 6].Value = (totalRegularJan + totalFixedJan + totalOverJan).ToString("N0");
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalSalaryJan += totalRegularJan + totalFixedJan + totalOverJan;
                        salaryTotalJan += totalSalaryJan;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 7].Value = (totalRegularFeb + totalFixedFeb + totalOverFeb).ToString("N0");
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalSalaryFeb += totalRegularFeb + totalFixedFeb + totalOverFeb;
                        salaryTotalFeb += totalSalaryFeb;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 8].Value = (totalRegularMar + totalFixedMar + totalOverMar).ToString("N0");
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalSalaryMar += totalRegularMar + totalFixedMar + totalOverMar;
                        salaryTotalMar += totalSalaryMar;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 9].Value = (totalRegularApr + totalFixedApr + totalOverApr).ToString("N0");
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalSalaryApr += totalRegularApr + totalFixedApr + totalOverApr;
                        salaryTotalApr += totalSalaryApr;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 10].Value = (totalRegularMay + totalFixedMay + totalOverMay).ToString("N0");
                        sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalSalaryMay += totalRegularMay + totalFixedMay + totalOverMay;
                        salaryTotalMay += totalSalaryMay;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 11].Value = (totalRegularJun + totalFixedJun + totalOverJun).ToString("N0");
                        sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalSalaryJun += totalRegularJun + totalFixedJun + totalOverJun;
                        salaryTotalJun += totalSalaryJun;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 12].Value = (totalRegularJul + totalFixedJul + totalOverJul).ToString("N0");
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalSalaryJul += totalRegularJul + totalFixedJul + totalOverJul;
                        salaryTotalJul += totalSalaryJul;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 13].Value = (totalRegularAug + totalFixedAug + totalOverAug).ToString("N0");
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalSalaryAug += totalRegularAug + totalFixedAug + totalOverAug;
                        salaryTotalAug += totalSalaryAug;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 14].Value = (totalRegularSep + totalFixedSep + totalOverSep).ToString("N0");
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalSalarySep += totalRegularSep + totalFixedSep + totalOverSep;
                        salaryTotalSep += totalSalarySep;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                    }
                    // Miscellneous Wages
                    if (salaryTypeCount == 6)
                    {


                            double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                            double beginningValue = _salaryBLL.GetGradeUnitPriceType(department.Id, salaryType.Id, 2022, item.Grade.Id).GradeLowPoints;
                            //double commonValue = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).SalaryIncreaseRate);

                        sheet.Cells[rowCount, 3].Value = (manpointOct * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        mWagesOct += manpointOct * beginningValue;
                        mwTotalOct += mWagesOct;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                        sheet.Cells[rowCount, 4].Value = (manpointNov * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        mWagesNov += manpointNov * beginningValue;
                        mwTotalNov += mWagesNov;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                        sheet.Cells[rowCount, 5].Value = (manpointDec * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        mWagesDec += manpointDec * beginningValue;
                        mwTotalDec += mWagesDec;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                        sheet.Cells[rowCount, 6].Value = (manpointJan * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        mWagesJan += manpointJan * beginningValue;
                        mwTotalJan += mWagesJan;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                        sheet.Cells[rowCount, 7].Value = (manpointFeb * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        mWagesFeb += manpointFeb * beginningValue;
                        mwTotalFeb += mWagesFeb;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                        sheet.Cells[rowCount, 8].Value = (manpointMar * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        mWagesMar += manpointMar * beginningValue;
                        mwTotalMar += mWagesMar;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                        sheet.Cells[rowCount, 9].Value = (manpointApr * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        mWagesApr += manpointApr * beginningValue;
                        mwTotalApr += mWagesApr;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                            double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                            sheet.Cells[rowCount, 10].Value = (manpointMay * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesMay += manpointMay * beginningValue;
                            mwTotalMay += mWagesMay;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                            sheet.Cells[rowCount, 11].Value = (manpointJun * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesJun += manpointJun * beginningValue;
                            mwTotalJun += mWagesJun;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }



                        double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                        sheet.Cells[rowCount, 12].Value = (manpointJul * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        mWagesJul += manpointJul * beginningValue;
                        mwTotalJul += mWagesJul;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                        sheet.Cells[rowCount, 13].Value = (manpointAug * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        mWagesAug += manpointAug * beginningValue;
                        mwTotalAug += mWagesAug;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                        sheet.Cells[rowCount, 14].Value = (manpointSep * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        mWagesSep += manpointSep * beginningValue;
                        mwTotalSep += mWagesSep;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                    }
                    // Dispatch Fee
                    if (salaryTypeCount == 7)
                    {


                            double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                            double beginningValue = _salaryBLL.GetGradeUnitPriceType(department.Id, salaryType.Id, 2022, item.Grade.Id).GradeLowPoints;


                        sheet.Cells[rowCount, 3].Value = (manpointOct * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        dispatchFeeOct += manpointOct * beginningValue;
                        dispatchFeeTotalOct += dispatchFeeOct;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                        sheet.Cells[rowCount, 4].Value = (manpointNov * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        dispatchFeeNov += manpointNov * beginningValue;
                        dispatchFeeTotalNov += dispatchFeeNov;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                        sheet.Cells[rowCount, 5].Value = (manpointDec * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        dispatchFeeDec += manpointDec * beginningValue;
                        dispatchFeeTotalDec += dispatchFeeDec;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                        sheet.Cells[rowCount, 6].Value = (manpointJan * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        dispatchFeeJan += manpointJan * beginningValue;
                        dispatchFeeTotalJan += dispatchFeeJan;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                        sheet.Cells[rowCount, 7].Value = (manpointFeb * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        dispatchFeeFeb += manpointFeb * beginningValue;
                        dispatchFeeTotalFeb += dispatchFeeFeb;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                        sheet.Cells[rowCount, 8].Value = (manpointMar * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        dispatchFeeMar += manpointMar * beginningValue;
                        dispatchFeeTotalMar += dispatchFeeMar;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                        sheet.Cells[rowCount, 9].Value = (manpointApr * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        dispatchFeeApr += manpointApr * beginningValue;
                        dispatchFeeTotalApr += dispatchFeeApr;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                        sheet.Cells[rowCount, 10].Value = (manpointMay * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        dispatchFeeMay += manpointMay * beginningValue;
                        dispatchFeeTotalMay += dispatchFeeMay;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                        sheet.Cells[rowCount, 11].Value = (manpointJun * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        dispatchFeeJun += manpointJun * beginningValue;
                        dispatchFeeTotalJun += dispatchFeeJun;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }



                        double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                        sheet.Cells[rowCount, 12].Value = (manpointJul * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        dispatchFeeJul += manpointJul * beginningValue;
                        dispatchFeeTotalJul += dispatchFeeJul;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                        sheet.Cells[rowCount, 13].Value = (manpointAug * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        dispatchFeeAug += manpointAug * beginningValue;
                        dispatchFeeTotalAug += dispatchFeeAug;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                        sheet.Cells[rowCount, 14].Value = (manpointSep * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        dispatchFeeSep += manpointSep * beginningValue;
                        dispatchFeeTotalSep += dispatchFeeSep;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                    }
                    // Provision for Employee Bonus
                    if (salaryTypeCount == 8)
                    {
                        double overWorkFixedTime = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).OverWorkFixedTime);
                        double bonusReserveConstant = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).BonusReserveConstant);

                        double valueOct = (totalRegularOct * overWorkFixedTime) / bonusReserveConstant;
                        valueOct = Double.IsNaN(valueOct) ? 0 : valueOct;
                        valueOct = Double.IsInfinity(valueOct) ? 0 : valueOct;
                        employeeBonusOct += valueOct;
                        employeeBonusTotalOct += employeeBonusOct;
                        sheet.Cells[rowCount, 3].Value = valueOct.ToString("N0");
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        double valueNov = (totalRegularNov * overWorkFixedTime) / bonusReserveConstant;
                        valueNov = Double.IsNaN(valueNov) ? 0 : valueNov;
                        valueNov = Double.IsInfinity(valueNov) ? 0 : valueNov;
                        employeeBonusNov += valueNov;
                        employeeBonusTotalNov += employeeBonusNov;
                        sheet.Cells[rowCount, 4].Value = valueNov.ToString("N0");
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        double valueDec = (totalRegularDec * overWorkFixedTime) / bonusReserveConstant;
                        valueDec = Double.IsNaN(valueDec) ? 0 : valueDec;
                        valueDec = Double.IsInfinity(valueDec) ? 0 : valueDec;
                        employeeBonusDec += valueDec;
                        employeeBonusTotalDec += employeeBonusDec;
                        sheet.Cells[rowCount, 5].Value = valueDec.ToString("N0");
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        double valueJan = (totalRegularJan * overWorkFixedTime) / bonusReserveConstant;
                        valueJan = Double.IsNaN(valueJan) ? 0 : valueJan;
                        valueJan = Double.IsInfinity(valueJan) ? 0 : valueJan;
                        employeeBonusJan += valueJan;
                        employeeBonusTotalJan += employeeBonusJan;
                        sheet.Cells[rowCount, 6].Value = valueJan.ToString("N0");
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double valueFeb = (totalRegularFeb * overWorkFixedTime) / bonusReserveConstant;
                        valueFeb = Double.IsNaN(valueFeb) ? 0 : valueFeb;
                        valueFeb = Double.IsInfinity(valueFeb) ? 0 : valueFeb;
                        employeeBonusFeb += valueFeb;
                        employeeBonusTotalFeb += employeeBonusFeb;
                        sheet.Cells[rowCount, 7].Value = valueFeb.ToString("N0");
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        double valueMar = (totalRegularMar * overWorkFixedTime) / bonusReserveConstant;
                        valueMar = Double.IsNaN(valueMar) ? 0 : valueMar;
                        valueMar = Double.IsInfinity(valueMar) ? 0 : valueMar;
                        employeeBonusMar += valueMar;
                        employeeBonusTotalMar += employeeBonusMar;
                        sheet.Cells[rowCount, 8].Value = valueMar.ToString("N0");
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        double valueApr = (totalRegularApr * overWorkFixedTime) / bonusReserveConstant;
                        valueApr = Double.IsNaN(valueApr) ? 0 : valueApr;
                        valueApr = Double.IsInfinity(valueApr) ? 0 : valueApr;
                        employeeBonusApr += valueApr;
                        employeeBonusTotalApr += employeeBonusApr;
                        sheet.Cells[rowCount, 9].Value = valueApr.ToString("N0");
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        double valueMay = (totalRegularMay * overWorkFixedTime) / bonusReserveConstant;
                        valueMay = Double.IsNaN(valueMay) ? 0 : valueMay;
                        valueMay = Double.IsInfinity(valueMay) ? 0 : valueMay;
                        employeeBonusMay += valueMay;
                        employeeBonusTotalMay += employeeBonusMay;
                        sheet.Cells[rowCount, 10].Value = valueMay.ToString("N0");
                        sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        double valueJun = (totalRegularJun * overWorkFixedTime) / bonusReserveConstant;
                        valueJun = Double.IsNaN(valueJun) ? 0 : valueJun;
                        valueJun = Double.IsInfinity(valueJun) ? 0 : valueJun;
                        employeeBonusJun += valueJun;
                        employeeBonusTotalJun += employeeBonusJun;
                        sheet.Cells[rowCount, 11].Value = valueJun.ToString("N0");
                        sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double valueJul = (totalRegularJul * overWorkFixedTime) / bonusReserveConstant;
                        valueJul = Double.IsNaN(valueJul) ? 0 : valueJul;
                        valueJul = Double.IsInfinity(valueJul) ? 0 : valueJul;
                        employeeBonusJul += valueJul;
                        employeeBonusTotalJul += employeeBonusJul;
                        sheet.Cells[rowCount, 12].Value = valueJul.ToString("N0");
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        double valueAug = (totalRegularAug * overWorkFixedTime) / bonusReserveConstant;
                        valueAug = Double.IsNaN(valueAug) ? 0 : valueAug;
                        valueAug = Double.IsInfinity(valueAug) ? 0 : valueAug;
                        employeeBonusAug += valueAug;
                        employeeBonusTotalAug += employeeBonusAug;
                        sheet.Cells[rowCount, 13].Value = valueAug.ToString("N0");
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double valueSep = (totalRegularSep * overWorkFixedTime) / bonusReserveConstant;
                        valueSep = Double.IsNaN(valueSep) ? 0 : valueSep;
                        valueSep = Double.IsInfinity(valueSep) ? 0 : valueSep;
                        employeeBonusSep += valueSep;
                        employeeBonusTotalSep += employeeBonusSep;
                        sheet.Cells[rowCount, 14].Value = valueSep.ToString("N0");
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                    }
                    //commuting expenses
                    if (salaryTypeCount == 9)
                    {


                            double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                            double beginningValue = _salaryBLL.GetGradeUnitPriceType(department.Id, 4, 2022, item.Grade.Id).GradeLowPoints;
                            //double commonValue = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).SalaryIncreaseRate);

                        sheet.Cells[rowCount, 3].Value = (manpointOct * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        commutingExpensesOct += manpointOct * beginningValue;
                        commutingExpensesTotalOct += commutingExpensesOct;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                        sheet.Cells[rowCount, 4].Value = (manpointNov * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        commutingExpensesNov += manpointNov * beginningValue;
                        commutingExpensesTotalNov += commutingExpensesNov;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                        sheet.Cells[rowCount, 5].Value = (manpointDec * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        commutingExpensesDec += manpointDec * beginningValue;
                        commutingExpensesTotalDec += commutingExpensesDec;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                        sheet.Cells[rowCount, 6].Value = (manpointJan * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        commutingExpensesJan += manpointJan * beginningValue;
                        commutingExpensesTotalJan += commutingExpensesJan;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                        sheet.Cells[rowCount, 7].Value = (manpointFeb * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        commutingExpensesFeb += manpointFeb * beginningValue;
                        commutingExpensesTotalFeb += commutingExpensesFeb;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                        sheet.Cells[rowCount, 8].Value = (manpointMar * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        commutingExpensesMar += manpointMar * beginningValue;
                        commutingExpensesTotalMar += commutingExpensesMar;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                        sheet.Cells[rowCount, 9].Value = (manpointApr * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        commutingExpensesApr += manpointApr * beginningValue;
                        commutingExpensesTotalApr += commutingExpensesApr;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                        sheet.Cells[rowCount, 10].Value = (manpointMay * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        commutingExpensesMay += manpointMay * beginningValue;
                        commutingExpensesTotalMay += commutingExpensesMay;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                        sheet.Cells[rowCount, 11].Value = (manpointJun * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        commutingExpensesJun += manpointJun * beginningValue;
                        commutingExpensesTotalJun += commutingExpensesJun;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }



                        double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                        sheet.Cells[rowCount, 12].Value = (manpointJul * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        commutingExpensesJul += manpointJul * beginningValue;
                        commutingExpensesTotalJul += commutingExpensesJul;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                        sheet.Cells[rowCount, 13].Value = (manpointAug * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        commutingExpensesAug += manpointAug * beginningValue;
                        commutingExpensesTotalAug += commutingExpensesAug;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                        sheet.Cells[rowCount, 14].Value = (manpointSep * beginningValue).ToString("N0");
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        commutingExpensesSep += manpointSep * beginningValue;
                        commutingExpensesTotalSep += commutingExpensesSep;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                    }
                    // welfare expenses
                    if (salaryTypeCount == 10)
                    {
                        double commonValue = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).WelfareCostRatio);
                        sheet.Cells[rowCount, 3].Value = ((totalSalaryOct + mWagesOct + commutingExpensesOct) * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        wExpensesOct += (totalSalaryOct + mWagesOct + commutingExpensesOct) * commonValue;
                        welfareExpensesTotalOct += wExpensesOct;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 4].Value = ((totalSalaryNov + mWagesNov + commutingExpensesNov) * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        wExpensesNov += (totalSalaryNov + mWagesNov + commutingExpensesNov) * commonValue;
                        welfareExpensesTotalNov += wExpensesNov;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 5].Value = ((totalSalaryDec + mWagesDec + commutingExpensesDec) * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        wExpensesDec += (totalSalaryDec + mWagesDec + commutingExpensesDec) * commonValue;
                        welfareExpensesTotalDec += wExpensesDec;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 6].Value = ((totalSalaryJan + mWagesJan + commutingExpensesJan) * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        wExpensesJan += (totalSalaryJan + mWagesJan + commutingExpensesJan) * commonValue;
                        welfareExpensesTotalJan += wExpensesJan;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 7].Value = ((totalSalaryFeb + mWagesFeb + commutingExpensesFeb) * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        wExpensesFeb += (totalSalaryFeb + mWagesFeb + commutingExpensesFeb) * commonValue;
                        welfareExpensesTotalFeb += wExpensesFeb;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 8].Value = ((totalSalaryMar + mWagesMar + commutingExpensesMar) * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        wExpensesMar += (totalSalaryMar + mWagesMar + commutingExpensesMar) * commonValue;
                        welfareExpensesTotalMar += wExpensesMar;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 9].Value = ((totalSalaryApr + mWagesApr + commutingExpensesApr) * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        wExpensesApr += (totalSalaryApr + mWagesApr + commutingExpensesApr) * commonValue;
                        welfareExpensesTotalApr += wExpensesApr;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 10].Value = ((totalSalaryMay + mWagesMay + commutingExpensesMay) * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        wExpensesMay += (totalSalaryMay + mWagesMay + commutingExpensesMay) * commonValue;
                        welfareExpensesTotalMay += wExpensesMay;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 11].Value = ((totalSalaryJun + mWagesJun + commutingExpensesJun) * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        wExpensesJun += (totalSalaryJun + mWagesJun + commutingExpensesJun) * commonValue;
                        welfareExpensesTotalJun += wExpensesJun;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 12].Value = ((totalSalaryJul + mWagesJul + commutingExpensesJul) * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        wExpensesJul += (totalSalaryJul + mWagesJul + commutingExpensesJul) * commonValue;
                        welfareExpensesTotalJul += wExpensesJul;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 13].Value = ((totalSalaryAug + mWagesAug + commutingExpensesAug) * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        wExpensesAug += (totalSalaryAug + mWagesAug + commutingExpensesAug) * commonValue;
                        welfareExpensesTotalAug += wExpensesAug;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 14].Value = ((totalSalarySep + mWagesSep + commutingExpensesSep) * commonValue).ToString("N0");
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        wExpensesSep += (totalSalarySep + mWagesSep + commutingExpensesSep) * commonValue;
                        welfareExpensesTotalSep += wExpensesSep;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                    }
                    // welfare expenses bonus
                    if (salaryTypeCount == 11)
                    {
                        sheet.Cells[rowCount, 3].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 4].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 5].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 6].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 7].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 8].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 9].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 10].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 11].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 12].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 13].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 14].Value = 0.ToString("N0");
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                    }
                    // total statutory
                    if (salaryTypeCount == 12)
                    {
                        sheet.Cells[rowCount, 3].Value = (wExpensesOct + wExpBonusOct).ToString("N0");
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalStatutoryOct += wExpensesOct + wExpBonusOct;
                        statutoryTotalOct += totalStatutoryOct;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 4].Value = (wExpensesNov + wExpBonusNov).ToString("N0");
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalStatutoryNov += wExpensesNov + wExpBonusNov;
                        statutoryTotalNov += totalStatutoryNov;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 5].Value = (wExpensesDec + wExpBonusDec).ToString("N0");
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalStatutoryDec += wExpensesDec + wExpBonusDec;
                        statutoryTotalDec += totalStatutoryDec;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 6].Value = (wExpensesJan + wExpBonusJan).ToString("N0");
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalStatutoryJan += wExpensesJan + wExpBonusJan;
                        statutoryTotalJan += totalStatutoryJan;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 7].Value = (wExpensesFeb + wExpBonusFeb).ToString("N0");
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalStatutoryFeb += wExpensesFeb + wExpBonusFeb;
                        statutoryTotalFeb += totalStatutoryFeb;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 8].Value = (wExpensesMar + wExpBonusMar).ToString("N0");
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalStatutoryMar += wExpensesMar + wExpBonusMar;
                        statutoryTotalMar += totalStatutoryMar;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 9].Value = (wExpensesApr + wExpBonusApr).ToString("N0");
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalStatutoryApr += wExpensesApr + wExpBonusApr;
                        statutoryTotalApr += totalStatutoryApr;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 10].Value = (wExpensesMay + wExpBonusMay).ToString("N0");
                        sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalStatutoryMay += wExpensesMay + wExpBonusMay;
                        statutoryTotalMay += totalStatutoryMay;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 11].Value = (wExpensesJun + wExpBonusJun).ToString("N0");
                        sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalStatutoryJun += wExpensesJun + wExpBonusJun;
                        statutoryTotalJun += totalStatutoryJun;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 12].Value = (wExpensesJul + wExpBonusJul).ToString("N0");
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalStatutoryJul += wExpensesJul + wExpBonusJul;
                        statutoryTotalJul += totalStatutoryJul;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        sheet.Cells[rowCount, 13].Value = (wExpensesAug + wExpBonusAug).ToString("N0");
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalStatutoryAug += wExpensesAug + wExpBonusAug;
                        statutoryTotalAug += totalStatutoryAug;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        sheet.Cells[rowCount, 14].Value = (wExpensesSep + wExpBonusSep).ToString("N0");
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        totalStatutorySep += wExpensesSep + wExpBonusSep;
                        statutoryTotalSep += totalStatutorySep;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                    }
                    // total expenses
                    if (salaryTypeCount == 13)
                    {
                        expensesTotalOct += totalSalaryOct + mWagesOct + dispatchFeeOct + employeeBonusOct + commutingExpensesOct + totalStatutoryOct;
                        sheet.Cells[rowCount, 3].Value = (totalSalaryOct + mWagesOct + dispatchFeeOct + employeeBonusOct + commutingExpensesOct + totalStatutoryOct).ToString("N0");
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        expensesTotalNov += totalSalaryNov + mWagesNov + dispatchFeeNov + employeeBonusNov + commutingExpensesNov + totalStatutoryNov;
                        sheet.Cells[rowCount, 4].Value = (totalSalaryNov + mWagesNov + dispatchFeeNov + employeeBonusNov + commutingExpensesNov + totalStatutoryNov).ToString("N0");
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        expensesTotalDec += totalSalaryDec + mWagesDec + dispatchFeeDec + employeeBonusDec + commutingExpensesDec + totalStatutoryDec;
                        sheet.Cells[rowCount, 5].Value = (totalSalaryDec + mWagesDec + dispatchFeeDec + employeeBonusDec + commutingExpensesDec + totalStatutoryDec).ToString("N0");
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        expensesTotalJan += totalSalaryJan + mWagesJan + dispatchFeeJan + employeeBonusJan + commutingExpensesJan + totalStatutoryJan;
                        sheet.Cells[rowCount, 6].Value = (totalSalaryJan + mWagesJan + dispatchFeeJan + employeeBonusJan + commutingExpensesJan + totalStatutoryJan).ToString("N0");
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        expensesTotalFeb += totalSalaryFeb + mWagesFeb + dispatchFeeFeb + employeeBonusFeb + commutingExpensesFeb + totalStatutoryFeb;
                        sheet.Cells[rowCount, 7].Value = (totalSalaryFeb + mWagesFeb + dispatchFeeFeb + employeeBonusFeb + commutingExpensesFeb + totalStatutoryFeb).ToString("N0");
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        expensesTotalMar += totalSalaryMar + mWagesMar + dispatchFeeMar + employeeBonusMar + commutingExpensesMar + totalStatutoryMar;
                        sheet.Cells[rowCount, 8].Value = (totalSalaryMar + mWagesMar + dispatchFeeMar + employeeBonusMar + commutingExpensesMar + totalStatutoryMar).ToString("N0");
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        expensesTotalApr += totalSalaryApr + mWagesApr + dispatchFeeApr + employeeBonusApr + commutingExpensesApr + totalStatutoryApr;
                        sheet.Cells[rowCount, 9].Value = (totalSalaryApr + mWagesApr + dispatchFeeApr + employeeBonusApr + commutingExpensesApr + totalStatutoryApr).ToString("N0");
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        expensesTotalMay += totalSalaryMay + mWagesMay + dispatchFeeMay + employeeBonusMay + commutingExpensesMay + totalStatutoryMay;
                        sheet.Cells[rowCount, 10].Value = (totalSalaryMay + mWagesMay + dispatchFeeMay + employeeBonusMay + commutingExpensesMay + totalStatutoryMay).ToString("N0");
                        sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        expensesTotalJun += totalSalaryJun + mWagesJun + dispatchFeeJun + employeeBonusJun + commutingExpensesJun + totalStatutoryJun;
                        sheet.Cells[rowCount, 11].Value = (totalSalaryJun + mWagesJun + dispatchFeeJun + employeeBonusJun + commutingExpensesJun + totalStatutoryJun).ToString("N0");
                        sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        expensesTotalJul += totalSalaryJul + mWagesJul + dispatchFeeJul + employeeBonusJul + commutingExpensesJul + totalStatutoryJul;
                        sheet.Cells[rowCount, 12].Value = (totalSalaryJul + mWagesJul + dispatchFeeJul + employeeBonusJul + commutingExpensesJul + totalStatutoryJul).ToString("N0");
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }
                        expensesTotalAug += totalSalaryAug + mWagesAug + dispatchFeeAug + employeeBonusAug + commutingExpensesAug + totalStatutoryAug;
                        sheet.Cells[rowCount, 13].Value = (totalSalaryAug + mWagesAug + dispatchFeeAug + employeeBonusAug + commutingExpensesAug + totalStatutoryAug).ToString("N0");
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                        expensesTotalSep += totalSalarySep + mWagesSep + dispatchFeeSep + employeeBonusSep + commutingExpensesSep + totalStatutorySep;
                        sheet.Cells[rowCount, 14].Value = (totalSalarySep + mWagesSep + dispatchFeeSep + employeeBonusSep + commutingExpensesSep + totalStatutorySep).ToString("N0");
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }

                    }


                    rowCount++;
                    innerCount++;
                    salaryTypeCount++;
                }

                count++;

            }
            #endregion


            #region salary types total
            int salaryTypeSummeryCount = 1;
            foreach (var item in _unitPriceTypeBLL.GetAllSalaryTypes())
            {
                sheet.Cells[rowCountForFuture, 1].Value = item.SalaryTypeName;
                // executive
                if (salaryTypeSummeryCount == 1)
                {
                    sheet.Cells[rowCountForFuture, 3].Value = 0;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 4].Value = 0;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 5].Value = 0;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 6].Value = 0;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 7].Value = 0;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 8].Value = 0;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 9].Value = 0;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 10].Value = 0;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 11].Value = 0;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 12].Value = 0;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 13].Value = 0;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 14].Value = 0;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                }
                // regular
                if (salaryTypeSummeryCount == 2)
                {
                    sheet.Cells[rowCountForFuture, 3].Value = regularSalaryTotalOct.ToString("N0");
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 4].Value = regularSalaryTotalNov.ToString("N0");
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 5].Value = regularSalaryTotalDec.ToString("N0");
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 6].Value = regularSalaryTotalJan.ToString("N0");
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 7].Value = regularSalaryTotalFeb.ToString("N0");
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 8].Value = regularSalaryTotalMar.ToString("N0");
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 9].Value = regularSalaryTotalApr.ToString("N0");
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 10].Value = regularSalaryTotalMay.ToString("N0");
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 11].Value = regularSalaryTotalJun.ToString("N0");
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 12].Value = regularSalaryTotalJul.ToString("N0");
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 13].Value = regularSalaryTotalAug.ToString("N0");
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 14].Value = regularSalaryTotalSep.ToString("N0");
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                }
                // fixed
                if (salaryTypeSummeryCount == 3)
                {
                    sheet.Cells[rowCountForFuture, 3].Value = fixedSalaryTotalOct.ToString("N0");
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 4].Value = fixedSalaryTotalNov.ToString("N0");
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 5].Value = fixedSalaryTotalDec.ToString("N0");
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 6].Value = fixedSalaryTotalJan.ToString("N0");
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 7].Value = fixedSalaryTotalFeb.ToString("N0");
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 8].Value = fixedSalaryTotalMar.ToString("N0");
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 9].Value = fixedSalaryTotalApr.ToString("N0");
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 10].Value = fixedSalaryTotalMay.ToString("N0");
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 11].Value = fixedSalaryTotalJun.ToString("N0");
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 12].Value = fixedSalaryTotalJul.ToString("N0");
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 13].Value = fixedSalaryTotalAug.ToString("N0");
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 14].Value = fixedSalaryTotalSep.ToString("N0");
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                }
                // overtime
                if (salaryTypeSummeryCount == 4)
                {
                    sheet.Cells[rowCountForFuture, 3].Value = overTimeSalaryTotalOct.ToString("N0");
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 4].Value = overTimeSalaryTotalNov.ToString("N0");
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 5].Value = overTimeSalaryTotalDec.ToString("N0");
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 6].Value = overTimeSalaryTotalJan.ToString("N0");
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 7].Value = overTimeSalaryTotalFeb.ToString("N0");
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 8].Value = overTimeSalaryTotalMar.ToString("N0");
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 9].Value = overTimeSalaryTotalApr.ToString("N0");
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 10].Value = overTimeSalaryTotalMay.ToString("N0");
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 11].Value = overTimeSalaryTotalJun.ToString("N0");
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 12].Value = overTimeSalaryTotalJul.ToString("N0");
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 13].Value = overTimeSalaryTotalAug.ToString("N0");
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 14].Value = overTimeSalaryTotalSep.ToString("N0");
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                }
                // total salary
                if (salaryTypeSummeryCount == 5)
                {
                    sheet.Cells[rowCountForFuture, 3].Value = salaryTotalOct.ToString("N0");
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 4].Value = salaryTotalNov.ToString("N0");
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 5].Value = salaryTotalDec.ToString("N0");
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 6].Value = salaryTotalJan.ToString("N0");
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 7].Value = salaryTotalFeb.ToString("N0");
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 8].Value = salaryTotalMar.ToString("N0");
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 9].Value = salaryTotalApr.ToString("N0");
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 10].Value = salaryTotalMay.ToString("N0");
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 11].Value = salaryTotalJun.ToString("N0");
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 12].Value = salaryTotalJul.ToString("N0");
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 13].Value = salaryTotalAug.ToString("N0");
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 14].Value = salaryTotalSep.ToString("N0");
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                }
                // miscellaneous wages
                if (salaryTypeSummeryCount == 6)
                {
                    sheet.Cells[rowCountForFuture, 3].Value = mwTotalOct.ToString("N0");
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 4].Value = mwTotalNov.ToString("N0");
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 5].Value = mwTotalDec.ToString("N0");
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 6].Value = mwTotalJan.ToString("N0");
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 7].Value = mwTotalFeb.ToString("N0");
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 8].Value = mwTotalMar.ToString("N0");
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 9].Value = mwTotalApr.ToString("N0");
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 10].Value = mwTotalMay.ToString("N0");
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 11].Value = mwTotalJun.ToString("N0");
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 12].Value = mwTotalJul.ToString("N0");
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 13].Value = mwTotalAug.ToString("N0");
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 14].Value = mwTotalSep.ToString("N0");
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                }
                // dispatch fee
                if (salaryTypeSummeryCount == 7)
                {
                    sheet.Cells[rowCountForFuture, 3].Value = dispatchFeeTotalOct.ToString("N0");
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 4].Value = dispatchFeeTotalNov.ToString("N0");
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 5].Value = dispatchFeeTotalDec.ToString("N0");
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 6].Value = dispatchFeeTotalJan.ToString("N0");
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 7].Value = dispatchFeeTotalFeb.ToString("N0");
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 8].Value = dispatchFeeTotalMar.ToString("N0");
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 9].Value = dispatchFeeTotalApr.ToString("N0");
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 10].Value = dispatchFeeTotalMay.ToString("N0");
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 11].Value = dispatchFeeTotalJun.ToString("N0");
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 12].Value = dispatchFeeTotalJul.ToString("N0");
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 13].Value = dispatchFeeTotalAug.ToString("N0");
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 14].Value = dispatchFeeTotalSep.ToString("N0");
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                }
                // provision for employee bonus
                if (salaryTypeSummeryCount == 8)
                {
                    sheet.Cells[rowCountForFuture, 3].Value = employeeBonusTotalOct.ToString("N0");
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 4].Value = employeeBonusTotalNov.ToString("N0");
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 5].Value = employeeBonusTotalDec.ToString("N0");
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 6].Value = employeeBonusTotalJan.ToString("N0");
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 7].Value = employeeBonusTotalFeb.ToString("N0");
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 8].Value = employeeBonusTotalMar.ToString("N0");
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 9].Value = employeeBonusTotalApr.ToString("N0");
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 10].Value = employeeBonusTotalMay.ToString("N0");
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 11].Value = employeeBonusTotalJun.ToString("N0");
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 12].Value = employeeBonusTotalJul.ToString("N0");
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 13].Value = employeeBonusTotalAug.ToString("N0");
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 14].Value = employeeBonusTotalSep.ToString("N0");
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                }
                // commuting expenses
                if (salaryTypeSummeryCount == 9)
                {
                    sheet.Cells[rowCountForFuture, 3].Value = commutingExpensesTotalOct.ToString("N0");
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 4].Value = commutingExpensesTotalNov.ToString("N0");
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 5].Value = commutingExpensesTotalDec.ToString("N0");
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 6].Value = commutingExpensesTotalJan.ToString("N0");
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 7].Value = commutingExpensesTotalFeb.ToString("N0");
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 8].Value = commutingExpensesTotalMar.ToString("N0");
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 9].Value = commutingExpensesTotalApr.ToString("N0");
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 10].Value = commutingExpensesTotalMay.ToString("N0");
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 11].Value = commutingExpensesTotalJun.ToString("N0");
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 12].Value = commutingExpensesTotalJul.ToString("N0");
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 13].Value = commutingExpensesTotalAug.ToString("N0");
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 14].Value = commutingExpensesTotalSep.ToString("N0");
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                }
                // salary statutory welfare expenses
                if (salaryTypeSummeryCount == 10)
                {
                    sheet.Cells[rowCountForFuture, 3].Value = welfareExpensesTotalOct.ToString("N0");
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 4].Value = welfareExpensesTotalNov.ToString("N0");
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 5].Value = welfareExpensesTotalDec.ToString("N0");
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 6].Value = welfareExpensesTotalJan.ToString("N0");
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 7].Value = welfareExpensesTotalFeb.ToString("N0");
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 8].Value = welfareExpensesTotalMar.ToString("N0");
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 9].Value = welfareExpensesTotalApr.ToString("N0");
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 10].Value = welfareExpensesTotalMay.ToString("N0");
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 11].Value = welfareExpensesTotalJun.ToString("N0");
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 12].Value = welfareExpensesTotalJul.ToString("N0");
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 13].Value = welfareExpensesTotalAug.ToString("N0");
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 14].Value = welfareExpensesTotalSep.ToString("N0");
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                }
                // provision for statotory welfare expenses for bonus
                if (salaryTypeSummeryCount == 11)
                {
                    sheet.Cells[rowCountForFuture, 3].Value = welfareExpensesBonusTotalOct.ToString("N0");
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 4].Value = welfareExpensesBonusTotalNov.ToString("N0");
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 5].Value = welfareExpensesBonusTotalDec.ToString("N0");
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 6].Value = welfareExpensesBonusTotalJan.ToString("N0");
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 7].Value = welfareExpensesBonusTotalFeb.ToString("N0");
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 8].Value = welfareExpensesBonusTotalMar.ToString("N0");
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 9].Value = welfareExpensesBonusTotalApr.ToString("N0");
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 10].Value = welfareExpensesBonusTotalMay.ToString("N0");
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 11].Value = welfareExpensesBonusTotalJun.ToString("N0");
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 12].Value = welfareExpensesBonusTotalJul.ToString("N0");
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 13].Value = welfareExpensesBonusTotalAug.ToString("N0");
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 14].Value = welfareExpensesBonusTotalSep.ToString("N0");
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                }
                // total statutory benefites
                if (salaryTypeSummeryCount == 12)
                {
                    sheet.Cells[rowCountForFuture, 3].Value = statutoryTotalOct.ToString("N0");
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 4].Value = statutoryTotalNov.ToString("N0");
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 5].Value = statutoryTotalDec.ToString("N0");
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 6].Value = statutoryTotalJan.ToString("N0");
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 7].Value = statutoryTotalFeb.ToString("N0");
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 8].Value = statutoryTotalMar.ToString("N0");
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 9].Value = statutoryTotalApr.ToString("N0");
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 10].Value = statutoryTotalMay.ToString("N0");
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 11].Value = statutoryTotalJun.ToString("N0");
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 12].Value = statutoryTotalJul.ToString("N0");
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 13].Value = statutoryTotalAug.ToString("N0");
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 14].Value = statutoryTotalSep.ToString("N0");
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                }
                // total expenses
                if (salaryTypeSummeryCount == 13)
                {
                    sheet.Cells[rowCountForFuture, 3].Value = expensesTotalOct.ToString("N0");
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 4].Value = expensesTotalNov.ToString("N0");
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 5].Value = expensesTotalDec.ToString("N0");
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 6].Value = expensesTotalJan.ToString("N0");
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 7].Value = expensesTotalFeb.ToString("N0");
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 8].Value = expensesTotalMar.ToString("N0");
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 9].Value = expensesTotalApr.ToString("N0");
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 10].Value = expensesTotalMay.ToString("N0");
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 11].Value = expensesTotalJun.ToString("N0");
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 12].Value = expensesTotalJul.ToString("N0");
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 13].Value = expensesTotalAug.ToString("N0");
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCountForFuture, 14].Value = expensesTotalSep.ToString("N0");
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                }




                //rowCount++;
                rowCountForFuture++;
                salaryTypeSummeryCount++;
            }
            #endregion



            rowCount = rowCount + 2;

            sheet.Cells[rowCount, 1, rowCount, 14].Style.Font.Bold = true;
            sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

            sheet.Cells[rowCount, 1].Value = "Costing";
            rowCount++;

            octTotal = 0; novTotal = 0; decTotal = 0; janTotal = 0; febTotal = 0; marTotal = 0; aprTotal = 0; mayTotal = 0; junTotal = 0; julTotal = 0; augTotal = 0; sepTotal = 0;
            // costing
            foreach (var item in salaryAssignmentDtos)
            {
                sheet.Cells[rowCount, 2].Value = item.Grade.GradeName;


                foreach (var singleAssignment in item.ForecastAssignmentViewModels)
                {
                    oct += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 10).Total);
                    nov += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 11).Total);
                    dec += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 12).Total);
                    jan += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 1).Total);
                    feb += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 2).Total);
                    mar += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 3).Total);
                    apr += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 4).Total);
                    may += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 5).Total);
                    jun += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 6).Total);
                    jul += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 7).Total);
                    aug += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 8).Total);
                    sep += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 9).Total);
                }
                sheet.Cells[rowCount, 3].Value = oct.ToString("N0");
                octTotal += oct;
                sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 4].Value = nov.ToString("N0");
                novTotal += nov;
                sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 5].Value = dec.ToString("N0");
                decTotal += dec;
                sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 6].Value = jan.ToString("N0");
                janTotal += jan;
                sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 7].Value = feb.ToString("N0");
                febTotal += feb;
                sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 8].Value = mar.ToString("N0");
                marTotal += mar;
                sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 9].Value = apr.ToString("N0");
                aprTotal += apr;
                sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 10].Value = may.ToString("N0");
                mayTotal += may;
                sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 11].Value = jun.ToString("N0");
                junTotal += jun;
                sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 12].Value = jul.ToString("N0");
                julTotal += jul;
                sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 13].Value = aug.ToString("N0");
                augTotal += aug;
                sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 14].Value = sep.ToString("N0");
                sepTotal += sep;
                sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                rowCount++;
            }


            sheet.Cells[rowCount, 2].Value = "合計";
            sheet.Cells[rowCount, 3].Value = octTotal.ToString("N0");
            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            sheet.Cells[rowCount, 4].Value = novTotal.ToString("N0");
            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            sheet.Cells[rowCount, 5].Value = decTotal.ToString("N0");
            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            sheet.Cells[rowCount, 6].Value = janTotal.ToString("N0");
            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            sheet.Cells[rowCount, 7].Value = febTotal.ToString("N0");
            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            sheet.Cells[rowCount, 8].Value = marTotal.ToString("N0");
            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            sheet.Cells[rowCount, 9].Value = aprTotal.ToString("N0");
            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            sheet.Cells[rowCount, 10].Value = mayTotal.ToString("N0");
            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            sheet.Cells[rowCount, 11].Value = junTotal.ToString("N0");
            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            sheet.Cells[rowCount, 12].Value = julTotal.ToString("N0");
            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            sheet.Cells[rowCount, 13].Value = augTotal.ToString("N0");
            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            sheet.Cells[rowCount, 14].Value = sepTotal.ToString("N0");
            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            rowCount++;

            rowCount = rowCount + 2;

            sheet.Cells[rowCount, 1, rowCount, 14].Style.Font.Bold = true;
            sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

            sheet.Cells[rowCount, 2].Value = "工数";
            rowCount++;

            if (assignmentsWithSectionAndCompany.Count > 0)
            {
                List<ForecastAssignmentViewModel> customizedList = assignmentsWithSectionAndCompany.ToList();
                List<ForecastAssignmentViewModel> tempList = null;
                octTotal = 0; novTotal = 0; decTotal = 0; janTotal = 0; febTotal = 0; marTotal = 0; aprTotal = 0; mayTotal = 0; junTotal = 0; julTotal = 0; augTotal = 0; sepTotal = 0;

                foreach (var item in assignmentsWithSectionAndCompany)
                {

                    tempList = customizedList.Where(l => l.SectionId == item.SectionId && l.CompanyId == item.CompanyId).ToList();
                    if (tempList.Count > 0)
                    {
                        foreach (var removableItem in tempList)
                        {
                            customizedList.Remove(removableItem);
                        }



                        sheet.Cells[rowCount, 1].Value = tempList[0].SectionName;
                        sheet.Cells[rowCount, 2].Value = tempList[0].CompanyName;

                        foreach (var tempItem in tempList)
                        {

                            oct += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 10).Points);
                            nov += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 11).Points);
                            dec += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 12).Points);
                            jan += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 1).Points);
                            feb += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 2).Points);
                            mar += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 3).Points);
                            apr += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 4).Points);
                            may += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 5).Points);
                            jun += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 6).Points);
                            jul += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 7).Points);
                            aug += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 8).Points);
                            sep += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 9).Points);
                        }

                        sheet.Cells[rowCount, 3].Value = oct;
                        octTotal += oct;
                        sheet.Cells[rowCount, 4].Value = nov;
                        novTotal += nov;
                        sheet.Cells[rowCount, 5].Value = dec;
                        decTotal += dec;
                        sheet.Cells[rowCount, 6].Value = jan;
                        janTotal += jan;
                        sheet.Cells[rowCount, 7].Value = feb;
                        febTotal += feb;
                        sheet.Cells[rowCount, 8].Value = mar;
                        marTotal += mar;
                        sheet.Cells[rowCount, 9].Value = apr;
                        aprTotal += apr;
                        sheet.Cells[rowCount, 10].Value = may;
                        mayTotal += may;
                        sheet.Cells[rowCount, 11].Value = jun;
                        junTotal += jun;
                        sheet.Cells[rowCount, 12].Value = jul;
                        julTotal += jul;
                        sheet.Cells[rowCount, 13].Value = aug;
                        augTotal += aug;
                        sheet.Cells[rowCount, 14].Value = sep;
                        sepTotal += sep;

                        oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                        rowCount++;
                    }
                }

                sheet.Cells[rowCount, 2].Value = "合計";
                sheet.Cells[rowCount, 3].Value = octTotal;
                sheet.Cells[rowCount, 4].Value = novTotal;
                sheet.Cells[rowCount, 5].Value = decTotal;
                sheet.Cells[rowCount, 6].Value = janTotal;
                sheet.Cells[rowCount, 7].Value = febTotal;
                sheet.Cells[rowCount, 8].Value = marTotal;
                sheet.Cells[rowCount, 9].Value = aprTotal;
                sheet.Cells[rowCount, 10].Value = mayTotal;
                sheet.Cells[rowCount, 11].Value = junTotal;
                sheet.Cells[rowCount, 12].Value = julTotal;
                sheet.Cells[rowCount, 13].Value = augTotal;
                sheet.Cells[rowCount, 14].Value = sepTotal;
                rowCount++;

                rowCount = rowCount + 2;
                sheet.Cells[rowCount, 1, rowCount, 14].Style.Font.Bold = true;
                sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.BackgroundColor.SetColor(Color.LightGray);


                sheet.Cells[rowCount, 2].Value = "金額";
                rowCount++;

                customizedList = assignmentsWithSectionAndCompany.ToList();
                octTotal = 0; novTotal = 0; decTotal = 0; janTotal = 0; febTotal = 0; marTotal = 0; aprTotal = 0; mayTotal = 0; junTotal = 0; julTotal = 0; augTotal = 0; sepTotal = 0;
                foreach (var item in assignmentsWithSectionAndCompany)
                {

                    tempList = customizedList.Where(l => l.SectionId == item.SectionId && l.CompanyId == item.CompanyId).ToList();
                    if (tempList.Count > 0)
                    {
                        foreach (var removableItem in tempList)
                        {
                            customizedList.Remove(removableItem);
                        }



                        sheet.Cells[rowCount, 1].Value = tempList[0].SectionName;
                        sheet.Cells[rowCount, 2].Value = tempList[0].CompanyName;

                        foreach (var tempItem in tempList)
                        {

                            oct += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 10).Total);
                            nov += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 11).Total);
                            dec += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 12).Total);
                            jan += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 1).Total);
                            feb += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 2).Total);
                            mar += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 3).Total);
                            apr += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 4).Total);
                            may += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 5).Total);
                            jun += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 6).Total);
                            jul += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 7).Total);
                            aug += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 8).Total);
                            sep += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 9).Total);
                        }

                        sheet.Cells[rowCount, 3].Value = oct.ToString("N0");
                        octTotal += oct;
                        sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[rowCount, 4].Value = nov.ToString("N0");
                        novTotal += nov;
                        sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[rowCount, 5].Value = dec.ToString("N0");
                        decTotal += dec;
                        sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[rowCount, 6].Value = jan.ToString("N0");
                        janTotal += jan;
                        sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[rowCount, 7].Value = feb.ToString("N0");
                        febTotal += feb;
                        sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[rowCount, 8].Value = mar.ToString("N0");
                        marTotal += mar;
                        sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[rowCount, 9].Value = apr.ToString("N0");
                        aprTotal += apr;
                        sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[rowCount, 10].Value = may.ToString("N0");
                        mayTotal += may;
                        sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[rowCount, 11].Value = jun.ToString("N0");
                        junTotal += jun;
                        sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[rowCount, 12].Value = jul.ToString("N0");
                        julTotal += jul;
                        sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[rowCount, 13].Value = aug.ToString("N0");
                        augTotal += aug;
                        sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[rowCount, 14].Value = sep.ToString("N0");
                        sepTotal += sep;
                        sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                        rowCount++;
                    }
                }
                sheet.Cells[rowCount, 2].Value = "合計";
                sheet.Cells[rowCount, 3].Value = octTotal.ToString("N0");
                sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 4].Value = novTotal.ToString("N0");
                sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 5].Value = decTotal.ToString("N0");
                sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 6].Value = janTotal.ToString("N0");
                sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 7].Value = febTotal.ToString("N0");
                sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 8].Value = marTotal.ToString("N0");
                sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 9].Value = aprTotal.ToString("N0");
                sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 10].Value = mayTotal.ToString("N0");
                sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 11].Value = junTotal.ToString("N0");
                sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 12].Value = julTotal.ToString("N0");
                sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 13].Value = augTotal.ToString("N0");
                sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 14].Value = sepTotal.ToString("N0");
                sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                rowCount++;


            }




        }

        private void GetSalaryMasterByYear(ExcelPackage excelPackage, int year)
        {
            List<Grade> grades = _gradeBll.GetAllGrade();
            var sheetSalaryMaster = excelPackage.Workbook.Worksheets.Add("FY" + year + "給与マスタ");
            sheetSalaryMaster.TabColor = Color.BlueViolet;
            List<Department> departments = _departmentBLL.GetAllDepartments().OrderBy(dep => dep.Id).ToList();
            List<SalaryMasterExportDto> salaryMasterExportDtos = new List<SalaryMasterExportDto>();
            var unitPriceTypeIds = _salaryBLL.GetUnitPriceTypeIdByYear(2022);
            foreach (var unitPriceType in unitPriceTypeIds)
            {
                foreach (var grade in grades)
                {
                    salaryMasterExportDtos.Add(_salaryBLL.GetUnitPriceTypeWithGradeSalaryByYear(year, grade.Id, unitPriceType));
                }
            }

            int rowCountSalaryMaster = 6;
            // year
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Value = "FY" + year;

            rowCountSalaryMaster++;
            // heading 1
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Value = "期初目標";
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Style.Font.Color.SetColor(Color.White);
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);


            sheetSalaryMaster.Cells[rowCountSalaryMaster, 7].Value = "下修目標";
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 7].Style.Font.Color.SetColor(Color.White);
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 7].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheetSalaryMaster.Cells[rowCountSalaryMaster, 8].Value = "期初目標";
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 8].Style.Font.Color.SetColor(Color.White);
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 8].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);


            sheetSalaryMaster.Cells[rowCountSalaryMaster, 9].Value = "下修目標";
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 9].Style.Font.Color.SetColor(Color.White);
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 9].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);
            rowCountSalaryMaster++;

            // heading 2
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 4].Value = "";
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 4].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheetSalaryMaster.Cells[rowCountSalaryMaster, 5].Value = "";
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheetSalaryMaster.Cells[rowCountSalaryMaster, 5].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            int extendedRow = 5;
            foreach (var item in departments)
            {
                sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Value = item.DepartmentName;
                sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Merge = true;
                sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.Font.Bold = true;
                sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.Font.Color.SetColor(Color.White);
                sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);
                extendedRow += 2;
            }

            rowCountSalaryMaster++;
            int prevUnitPriceTypeId = 0;

            foreach (var item in salaryMasterExportDtos)
            {
                if (item.GradeSalaryTypes.Count > 1)
                {
                    if (prevUnitPriceTypeId != item.UnitPriceType.Id)
                    {
                        sheetSalaryMaster.Cells[rowCountSalaryMaster, 4].Value = item.UnitPriceType.TypeName;
                        prevUnitPriceTypeId = item.UnitPriceType.Id;
                    }

                    // for grade name
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, 5].Value = item.GradeSalaryTypes[0].GradeName;
                    // for beginning target
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Value = item.GradeSalaryTypes[0].GradeLowPoints.ToString("N0");
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    // for beginning target
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, 7].Value = item.GradeSalaryTypes[0].GradeHighPoints.ToString("N0");
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    // for downward revision target
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, 8].Value = item.GradeSalaryTypes[1].GradeLowPoints.ToString("N0");
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    // for downward revision target
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, 9].Value = item.GradeSalaryTypes[1].GradeHighPoints.ToString("N0");
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    rowCountSalaryMaster++;
                }



            }
        }

        private void GetCommonMaster(ExcelPackage excelPackage)
        {
            var sheetCommonMaster = excelPackage.Workbook.Worksheets.Add("共通マスタ");
            sheetCommonMaster.TabColor = Color.BlueViolet;

            List<CommonMaster> commonMasters = _commonMasterBLL.GetCommonMasters();

            sheetCommonMaster.Cells[6, 4].Value = "";
            sheetCommonMaster.Cells[6, 4].Style.Font.Color.SetColor(Color.White);
            sheetCommonMaster.Cells[6, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheetCommonMaster.Cells[6, 4].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheetCommonMaster.Cells[6, 5].Value = "昇給率";
            sheetCommonMaster.Cells[6, 5].Style.Font.Color.SetColor(Color.White);
            sheetCommonMaster.Cells[6, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheetCommonMaster.Cells[6, 5].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheetCommonMaster.Cells[6, 6].Value = "残業固定時間";
            sheetCommonMaster.Cells[6, 6].Style.Font.Color.SetColor(Color.White);
            sheetCommonMaster.Cells[6, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheetCommonMaster.Cells[6, 6].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);


            sheetCommonMaster.Cells[6, 7].Value = "賞与引当金比率";
            sheetCommonMaster.Cells[6, 7].Style.Font.Color.SetColor(Color.White);
            sheetCommonMaster.Cells[6, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheetCommonMaster.Cells[6, 7].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheetCommonMaster.Cells[6, 8].Value = "賞与引当金定数";
            sheetCommonMaster.Cells[6, 8].Style.Font.Color.SetColor(Color.White);
            sheetCommonMaster.Cells[6, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheetCommonMaster.Cells[6, 8].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            sheetCommonMaster.Cells[6, 9].Value = "給与法定福利費比率";
            sheetCommonMaster.Cells[6, 9].Style.Font.Color.SetColor(Color.White);
            sheetCommonMaster.Cells[6, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheetCommonMaster.Cells[6, 9].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

            int rowCountCommonMaster = 7;


            if (commonMasters.Count > 0)
            {
                foreach (var commonMaster in commonMasters)
                {
                    sheetCommonMaster.Cells[rowCountCommonMaster, 4].Value = commonMaster.GradeName;
                    sheetCommonMaster.Cells[rowCountCommonMaster, 5].Value = commonMaster.SalaryIncreaseRate;
                    sheetCommonMaster.Cells[rowCountCommonMaster, 6].Value = commonMaster.OverWorkFixedTime;
                    sheetCommonMaster.Cells[rowCountCommonMaster, 7].Value = commonMaster.BonusReserveRatio;
                    sheetCommonMaster.Cells[rowCountCommonMaster, 8].Value = commonMaster.BonusReserveConstant;
                    sheetCommonMaster.Cells[rowCountCommonMaster, 9].Value = commonMaster.WelfareCostRatio;

                    rowCountCommonMaster++;
                }
            }
        }


        [HttpPost]
        public ActionResult DataExportByAllocation(int departmentId = 0, int explanationId = 0)
        {

            List<Grade> grades = _gradeBll.GetAllGrade();

            List<SalaryAssignmentDto> salaryAssignmentDtos = new List<SalaryAssignmentDto>();
            var department = _departmentBLL.GetDepartmentByDepartemntId(departmentId);
            var explanation = _explanationsBLL.GetExplanationByExplanationId(explanationId);
            var salaries = _salaryBLL.GetAllSalaryPoints();


            // get all data with department and allocation
            List<ForecastAssignmentViewModel> assignmentsWithForecast = _exportBLL.AssignmentsByAllocation(departmentId, explanationId);
            //filtered by grade id
            List<ForecastAssignmentViewModel> assignmentsWithGrade = assignmentsWithForecast.Where(a => a.GradeId != null).ToList();
            //filtered by section and company
            List<ForecastAssignmentViewModel> assignmentsWithSectionAndCompany = assignmentsWithForecast.Where(a => a.SectionId != null && a.CompanyId != null).ToList();

            foreach (var item in grades)
            {
                List<ForecastAssignmentViewModel> forecastAssignmentViews = new List<ForecastAssignmentViewModel>();

                SalaryAssignmentDto salaryAssignmentDto = new SalaryAssignmentDto();
                salaryAssignmentDto.Grade = item;

                var gradeSalaryTypes = _exportBLL.GetGradeUnitPriceTypes(item.Id, departmentId, 2022, 1);

                foreach (var gradeSalary in gradeSalaryTypes)
                {
                    List<ForecastAssignmentViewModel> filteredAssignmentsByGradeSalaryTypeId = assignmentsWithGrade.Where(a => a.GradeId == gradeSalary.GradeId.ToString()).ToList();
                    forecastAssignmentViews.AddRange(filteredAssignmentsByGradeSalaryTypeId);
                }

                salaryAssignmentDto.ForecastAssignmentViewModels = forecastAssignmentViews;
                salaryAssignmentDtos.Add(salaryAssignmentDto);

            }

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add(explanation.ExplanationName + " (" + DateTime.Today.ToString("dd-MMM-yyyy") + ")");

                #region header column
                //row-1
                sheet.Cells[1, 1].Value = "導入";
                sheet.Cells[1, 1].Style.Font.Color.SetColor(Color.Red);
                sheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                sheet.Cells[1, 2].Value = department.DepartmentName;
                sheet.Cells[1, 2].Style.Font.Color.SetColor(Color.Red);
                sheet.Cells[1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                //row-2
                sheet.Cells[2, 1].Value = "FY2022";

                //row-3
                sheet.Cells[3, 1].Value = "";
                sheet.Cells[3, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 1].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[3, 2].Value = "";
                sheet.Cells[3, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 2].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[3, 3].Value = "FY2022";
                sheet.Cells[3, 3].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[3, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 3].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[3, 4].Value = "FY2022";
                sheet.Cells[3, 4].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[3, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 4].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[3, 5].Value = "FY2022";
                sheet.Cells[3, 5].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[3, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 5].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[3, 6].Value = "FY2022";
                sheet.Cells[3, 6].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[3, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 6].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[3, 7].Value = "FY2022";
                sheet.Cells[3, 7].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[3, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 7].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[3, 8].Value = "FY2022";
                sheet.Cells[3, 8].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[3, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 8].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[3, 9].Value = "FY2022";
                sheet.Cells[3, 9].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[3, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 9].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[3, 10].Value = "FY2022";
                sheet.Cells[3, 10].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[3, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 10].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[3, 11].Value = "FY2022";
                sheet.Cells[3, 11].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[3, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 11].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[3, 12].Value = "FY2022";
                sheet.Cells[3, 12].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[3, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 12].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[3, 13].Value = "FY2022";
                sheet.Cells[3, 13].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[3, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 13].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[3, 14].Value = "FY2022";
                sheet.Cells[3, 14].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[3, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 14].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                // row -4

                sheet.Cells[4, 1].Value = "";
                sheet.Cells[4, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 1].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[4, 2].Value = "";
                sheet.Cells[4, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 2].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[4, 3].Value = "10月";
                sheet.Cells[4, 3].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[4, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 3].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[4, 4].Value = "11月";
                sheet.Cells[4, 4].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[4, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 4].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[4, 5].Value = "12月";
                sheet.Cells[4, 5].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[4, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 5].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[4, 6].Value = "1月";
                sheet.Cells[4, 6].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[4, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 6].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[4, 7].Value = "2月";
                sheet.Cells[4, 7].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[4, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 7].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[4, 8].Value = "3月";
                sheet.Cells[4, 8].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[4, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 8].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[4, 9].Value = "4月";
                sheet.Cells[4, 9].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[4, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 9].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[4, 10].Value = "5月";
                sheet.Cells[4, 10].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[4, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 10].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[4, 11].Value = "6月";
                sheet.Cells[4, 11].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[4, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 11].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[4, 12].Value = "7月";
                sheet.Cells[4, 12].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[4, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 12].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[4, 13].Value = "8月";
                sheet.Cells[4, 13].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[4, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 13].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheet.Cells[4, 14].Value = "9月";
                sheet.Cells[4, 14].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[4, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[4, 14].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                #endregion

                int rowCount = 5;

                sheet.Cells[rowCount, 1].Value = "人数";

                decimal oct = 0;
                decimal nov = 0;
                decimal dec = 0;
                decimal jan = 0;
                decimal feb = 0;
                decimal mar = 0;
                decimal apr = 0;
                decimal may = 0;
                decimal jun = 0;
                decimal jul = 0;
                decimal aug = 0;
                decimal sep = 0;

                decimal octOverTime = 0;
                decimal novOverTime = 0;
                decimal decOverTime = 0;
                decimal janOverTime = 0;
                decimal febOverTime = 0;
                decimal marOverTime = 0;
                decimal aprOverTime = 0;
                decimal mayOverTime = 0;
                decimal junOverTime = 0;
                decimal julOverTime = 0;
                decimal augOverTime = 0;
                decimal sepOverTime = 0;


                decimal octTotal = 0;
                decimal novTotal = 0;
                decimal decTotal = 0;
                decimal janTotal = 0;
                decimal febTotal = 0;
                decimal marTotal = 0;
                decimal aprTotal = 0;
                decimal mayTotal = 0;
                decimal junTotal = 0;
                decimal julTotal = 0;
                decimal augTotal = 0;
                decimal sepTotal = 0;

                //man month point
                foreach (var item in salaryAssignmentDtos)
                {
                    sheet.Cells[rowCount, 2].Value = item.Grade.GradeName;



                    foreach (var singleAssignment in item.ForecastAssignmentViewModels)
                    {
                        oct += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 10).Points);
                        nov += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 11).Points);
                        dec += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 12).Points);
                        jan += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 1).Points);
                        feb += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 2).Points);
                        mar += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 3).Points);
                        apr += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 4).Points);
                        may += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 5).Points);
                        jun += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 6).Points);
                        jul += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 7).Points);
                        aug += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 8).Points);
                        sep += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 9).Points);
                    }
                    sheet.Cells[rowCount, 3].Value = oct;
                    octTotal += oct;
                    sheet.Cells[rowCount, 4].Value = nov;
                    novTotal += nov;
                    sheet.Cells[rowCount, 5].Value = dec;
                    decTotal += dec;
                    sheet.Cells[rowCount, 6].Value = jan;
                    janTotal += jan;
                    sheet.Cells[rowCount, 7].Value = feb;
                    febTotal += feb;
                    sheet.Cells[rowCount, 8].Value = mar;
                    marTotal += mar;
                    sheet.Cells[rowCount, 9].Value = apr;
                    aprTotal += apr;
                    sheet.Cells[rowCount, 10].Value = may;
                    mayTotal += may;
                    sheet.Cells[rowCount, 11].Value = jun;
                    junTotal += jun;
                    sheet.Cells[rowCount, 12].Value = jul;
                    julTotal += jul;
                    sheet.Cells[rowCount, 13].Value = aug;
                    augTotal += aug;
                    sheet.Cells[rowCount, 14].Value = sep;
                    sepTotal += sep;


                    //valuesWithGrades.Add(new ValuesWithGrade { GradeId = item.Grade.Id,GradeName=item.Grade.GradeName,Point= octTotal,MonthId=10 });
                    oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                    rowCount++;
                }
                sheet.Cells[rowCount, 2].Value = "合計";
                sheet.Cells[rowCount, 3].Value = octTotal;
                sheet.Cells[rowCount, 4].Value = novTotal;
                sheet.Cells[rowCount, 5].Value = decTotal;
                sheet.Cells[rowCount, 6].Value = janTotal;
                sheet.Cells[rowCount, 7].Value = febTotal;
                sheet.Cells[rowCount, 8].Value = marTotal;
                sheet.Cells[rowCount, 9].Value = aprTotal;
                sheet.Cells[rowCount, 10].Value = mayTotal;
                sheet.Cells[rowCount, 11].Value = junTotal;
                sheet.Cells[rowCount, 12].Value = julTotal;
                sheet.Cells[rowCount, 13].Value = augTotal;
                sheet.Cells[rowCount, 14].Value = sepTotal;
                rowCount++;

                #region over time
                sheet.Cells[rowCount, 1].Value = "1人あたりの時間外勤務見込 \n (みなし時間（固定時間）\n を含む残業時間 \n を入力してください";

                //over time
                foreach (var item in salaryAssignmentDtos)
                {
                    sheet.Cells[rowCount, 2].Value = item.Grade.GradeName;



                    foreach (var singleAssignment in item.ForecastAssignmentViewModels)
                    {
                        octOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 10).OverTime);
                        novOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 11).OverTime);
                        decOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 12).OverTime);
                        janOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 1).OverTime);
                        febOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 2).OverTime);
                        marOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 3).OverTime);
                        aprOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 4).OverTime);
                        mayOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 5).OverTime);
                        junOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 6).OverTime);
                        julOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 7).OverTime);
                        augOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 8).OverTime);
                        sepOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 9).OverTime);
                    }
                    sheet.Cells[rowCount, 3].Value = octOverTime;

                    sheet.Cells[rowCount, 4].Value = novOverTime;

                    sheet.Cells[rowCount, 5].Value = decOverTime;

                    sheet.Cells[rowCount, 6].Value = janOverTime;

                    sheet.Cells[rowCount, 7].Value = febOverTime;

                    sheet.Cells[rowCount, 8].Value = marOverTime;

                    sheet.Cells[rowCount, 9].Value = aprOverTime;

                    sheet.Cells[rowCount, 10].Value = mayOverTime;

                    sheet.Cells[rowCount, 11].Value = junOverTime;

                    sheet.Cells[rowCount, 12].Value = julOverTime;

                    sheet.Cells[rowCount, 13].Value = augOverTime;

                    sheet.Cells[rowCount, 14].Value = sepOverTime;



                    //valuesWithGrades.Add(new ValuesWithGrade { GradeId = item.Grade.Id,GradeName=item.Grade.GradeName,Point= octTotal,MonthId=10 });
                    octOverTime = 0; novOverTime = 0; decOverTime = 0; janOverTime = 0; febOverTime = 0; marOverTime = 0; aprOverTime = 0; mayOverTime = 0; junOverTime = 0; julOverTime = 0; augOverTime = 0; sepOverTime = 0;
                    rowCount++;
                }

                #endregion

                rowCount++;

                int rowCountForFuture = rowCount;
                rowCount += 12;

                #region variables for salary types.

                // 2. regular total
                double regularSalaryTotalOct = 0, regularSalaryTotalNov = 0, regularSalaryTotalDec = 0, regularSalaryTotalJan = 0, regularSalaryTotalFeb = 0, regularSalaryTotalMar = 0, regularSalaryTotalApr = 0, regularSalaryTotalMay = 0, regularSalaryTotalJun = 0, regularSalaryTotalJul = 0, regularSalaryTotalAug = 0, regularSalaryTotalSep = 0;
                // 3. fixed total
                double fixedSalaryTotalOct = 0, fixedSalaryTotalNov = 0, fixedSalaryTotalDec = 0, fixedSalaryTotalJan = 0, fixedSalaryTotalFeb = 0, fixedSalaryTotalMar = 0, fixedSalaryTotalApr = 0, fixedSalaryTotalMay = 0, fixedSalaryTotalJun = 0, fixedSalaryTotalJul = 0, fixedSalaryTotalAug = 0, fixedSalaryTotalSep = 0;
                // 4. over time
                double overTimeSalaryTotalOct = 0, overTimeSalaryTotalNov = 0, overTimeSalaryTotalDec = 0, overTimeSalaryTotalJan = 0, overTimeSalaryTotalFeb = 0, overTimeSalaryTotalMar = 0, overTimeSalaryTotalApr = 0, overTimeSalaryTotalMay = 0, overTimeSalaryTotalJun = 0, overTimeSalaryTotalJul = 0, overTimeSalaryTotalAug = 0, overTimeSalaryTotalSep = 0;
                // 5. total Salary
                double salaryTotalOct = 0, salaryTotalNov = 0, salaryTotalDec = 0, salaryTotalJan = 0, salaryTotalFeb = 0, salaryTotalMar = 0, salaryTotalApr = 0, salaryTotalMay = 0, salaryTotalJun = 0, salaryTotalJul = 0, salaryTotalAug = 0, salaryTotalSep = 0;
                // 6. Miscellaneous Wages
                double mwTotalOct = 0, mwTotalNov = 0, mwTotalDec = 0, mwTotalJan = 0, mwTotalFeb = 0, mwTotalMar = 0, mwTotalApr = 0, mwTotalMay = 0, mwTotalJun = 0, mwTotalJul = 0, mwTotalAug = 0, mwTotalSep = 0;
                // 7. Dispatch Fee
                double dispatchFeeTotalOct = 0, dispatchFeeTotalNov = 0, dispatchFeeTotalDec = 0, dispatchFeeTotalJan = 0, dispatchFeeTotalFeb = 0, dispatchFeeTotalMar = 0, dispatchFeeTotalApr = 0, dispatchFeeTotalMay = 0, dispatchFeeTotalJun = 0, dispatchFeeTotalJul = 0, dispatchFeeTotalAug = 0, dispatchFeeTotalSep = 0;
                // 8. Provision for Employee Bonus
                double employeeBonusTotalOct = 0, employeeBonusTotalNov = 0, employeeBonusTotalDec = 0, employeeBonusTotalJan = 0, employeeBonusTotalFeb = 0, employeeBonusTotalMar = 0, employeeBonusTotalApr = 0, employeeBonusTotalMay = 0, employeeBonusTotalJun = 0, employeeBonusTotalJul = 0, employeeBonusTotalAug = 0, employeeBonusTotalSep = 0;
                // 9. Commuting Expenses
                double commutingExpensesTotalOct = 0, commutingExpensesTotalNov = 0, commutingExpensesTotalDec = 0, commutingExpensesTotalJan = 0, commutingExpensesTotalFeb = 0, commutingExpensesTotalMar = 0, commutingExpensesTotalApr = 0, commutingExpensesTotalMay = 0, commutingExpensesTotalJun = 0, commutingExpensesTotalJul = 0, commutingExpensesTotalAug = 0, commutingExpensesTotalSep = 0;
                // 10. Salary Statutory Welfare Expenses
                double welfareExpensesTotalOct = 0, welfareExpensesTotalNov = 0, welfareExpensesTotalDec = 0, welfareExpensesTotalJan = 0, welfareExpensesTotalFeb = 0, welfareExpensesTotalMar = 0, welfareExpensesTotalApr = 0, welfareExpensesTotalMay = 0, welfareExpensesTotalJun = 0, welfareExpensesTotalJul = 0, welfareExpensesTotalAug = 0, welfareExpensesTotalSep = 0;
                // 11. Provision for Statotory Welfare Expenses for Bonus
                double welfareExpensesBonusTotalOct = 0, welfareExpensesBonusTotalNov = 0, welfareExpensesBonusTotalDec = 0, welfareExpensesBonusTotalJan = 0, welfareExpensesBonusTotalFeb = 0, welfareExpensesBonusTotalMar = 0, welfareExpensesBonusTotalApr = 0, welfareExpensesBonusTotalMay = 0, welfareExpensesBonusTotalJun = 0, welfareExpensesBonusTotalJul = 0, welfareExpensesBonusTotalAug = 0, welfareExpensesBonusTotalSep = 0;
                // 12. Total Statutory Benefites
                double statutoryTotalOct = 0, statutoryTotalNov = 0, statutoryTotalDec = 0, statutoryTotalJan = 0, statutoryTotalFeb = 0, statutoryTotalMar = 0, statutoryTotalApr = 0, statutoryTotalMay = 0, statutoryTotalJun = 0, statutoryTotalJul = 0, statutoryTotalAug = 0, statutoryTotalSep = 0;
                // 13. Total Expenses
                double expensesTotalOct = 0, expensesTotalNov = 0, expensesTotalDec = 0, expensesTotalJan = 0, expensesTotalFeb = 0, expensesTotalMar = 0, expensesTotalApr = 0, expensesTotalMay = 0, expensesTotalJun = 0, expensesTotalJul = 0, expensesTotalAug = 0, expensesTotalSep = 0;

                #endregion

                #region other grade with entities
                rowCount++;
                int count = 1;
                foreach (var item in salaryAssignmentDtos)
                {
                    sheet.Cells[rowCount, 1].Value = item.Grade.GradeName;

                    if (count % 2 == 0)
                    {
                        // even
                        sheet.Cells[rowCount, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCount, 1].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                    }
                    else
                    {
                        // odd
                        sheet.Cells[rowCount, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCount, 1].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                    }

                    //salary allow regular
                    double totalRegularOct = 0, totalRegularNov = 0, totalRegularDec = 0, totalRegularJan = 0, totalRegularFeb = 0, totalRegularMar = 0, totalRegularApr = 0, totalRegularMay = 0, totalRegularJun = 0, totalRegularJul = 0, totalRegularAug = 0, totalRegularSep = 0;
                    // salary allow fixed
                    double totalFixedOct = 0, totalFixedNov = 0, totalFixedDec = 0, totalFixedJan = 0, totalFixedFeb = 0, totalFixedMar = 0, totalFixedApr = 0, totalFixedMay = 0, totalFixedJun = 0, totalFixedJul = 0, totalFixedAug = 0, totalFixedSep = 0;
                    // salary allow oertime
                    double totalOverOct = 0, totalOverNov = 0, totalOverDec = 0, totalOverJan = 0, totalOverFeb = 0, totalOverMar = 0, totalOverApr = 0, totalOverMay = 0, totalOverJun = 0, totalOverJul = 0, totalOverAug = 0, totalOverSep = 0;
                    // total salary
                    double totalSalaryOct = 0, totalSalaryNov = 0, totalSalaryDec = 0, totalSalaryJan = 0, totalSalaryFeb = 0, totalSalaryMar = 0, totalSalaryApr = 0, totalSalaryMay = 0, totalSalaryJun = 0, totalSalaryJul = 0, totalSalaryAug = 0, totalSalarySep = 0;
                    // Miscellaneous Wages
                    double mWagesOct = 0, mWagesNov = 0, mWagesDec = 0, mWagesJan = 0, mWagesFeb = 0, mWagesMar = 0, mWagesApr = 0, mWagesMay = 0, mWagesJun = 0, mWagesJul = 0, mWagesAug = 0, mWagesSep = 0;
                    //dispatch fee
                    double dispatchFeeOct = 0, dispatchFeeNov = 0, dispatchFeeDec = 0, dispatchFeeJan = 0, dispatchFeeFeb = 0, dispatchFeeMar = 0, dispatchFeeApr = 0, dispatchFeeMay = 0, dispatchFeeJun = 0, dispatchFeeJul = 0, dispatchFeeAug = 0, dispatchFeeSep = 0;
                    // employee bonus
                    double employeeBonusOct = 0, employeeBonusNov = 0, employeeBonusDec = 0, employeeBonusJan = 0, employeeBonusFeb = 0, employeeBonusMar = 0, employeeBonusApr = 0, employeeBonusMay = 0, employeeBonusJun = 0, employeeBonusJul = 0, employeeBonusAug = 0, employeeBonusSep = 0;
                    //commuting expenses
                    double commutingExpensesOct = 0, commutingExpensesNov = 0, commutingExpensesDec = 0, commutingExpensesJan = 0, commutingExpensesFeb = 0, commutingExpensesMar = 0, commutingExpensesApr = 0, commutingExpensesMay = 0, commutingExpensesJun = 0, commutingExpensesJul = 0, commutingExpensesAug = 0, commutingExpensesSep = 0;
                    // welfare expenses
                    double wExpensesOct = 0, wExpensesNov = 0, wExpensesDec = 0, wExpensesJan = 0, wExpensesFeb = 0, wExpensesMar = 0, wExpensesApr = 0, wExpensesMay = 0, wExpensesJun = 0, wExpensesJul = 0, wExpensesAug = 0, wExpensesSep = 0;
                    // welfare expenses bonuses
                    double wExpBonusOct = 0, wExpBonusNov = 0, wExpBonusDec = 0, wExpBonusJan = 0, wExpBonusFeb = 0, wExpBonusMar = 0, wExpBonusApr = 0, wExpBonusMay = 0, wExpBonusJun = 0, wExpBonusJul = 0, wExpBonusAug = 0, wExpBonusSep = 0;
                    // total statutory
                    double totalStatutoryOct = 0, totalStatutoryNov = 0, totalStatutoryDec = 0, totalStatutoryJan = 0, totalStatutoryFeb = 0, totalStatutoryMar = 0, totalStatutoryApr = 0, totalStatutoryMay = 0, totalStatutoryJun = 0, totalStatutoryJul = 0, totalStatutoryAug = 0, totalStatutorySep = 0;

                    int innerCount = 1;
                    int salaryTypeCount = 1;
                    foreach (var salaryType in _unitPriceTypeBLL.GetAllSalaryTypes())
                    {
                        if (innerCount > 1)
                        {
                            sheet.Cells[rowCount, 1].Value = "";

                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 1].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 1].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                        }

                        sheet.Cells[rowCount, 2].Value = salaryType.SalaryTypeName;
                        if (count % 2 == 0)
                        {
                            // even
                            sheet.Cells[rowCount, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 2].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                        }
                        else
                        {
                            // odd
                            sheet.Cells[rowCount, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[rowCount, 2].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                        }


                        // alligned with the serial of salary type

                        // executives
                        if (salaryTypeCount == 1)
                        {
                            sheet.Cells[rowCount, 3].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 4].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 5].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 6].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 7].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 8].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 9].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 10].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 11].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 12].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 13].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 14].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                        }
                        // salary allowance (regular)
                        if (salaryTypeCount == 2)
                        {


                            double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                            double beginningValue = _salaryBLL.GetGradeUnitPriceType(departmentId, 1, 2022, item.Grade.Id).GradeLowPoints;
                            double commonValue = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).SalaryIncreaseRate);

                            sheet.Cells[rowCount, 3].Value = (manpointOct * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalRegularOct += manpointOct * beginningValue * commonValue;
                            regularSalaryTotalOct += totalRegularOct;

                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                            sheet.Cells[rowCount, 4].Value = (manpointNov * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalRegularNov += manpointNov * beginningValue * commonValue;
                            regularSalaryTotalNov += totalRegularNov;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                            sheet.Cells[rowCount, 5].Value = (manpointDec * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalRegularDec += manpointDec * beginningValue * commonValue;
                            regularSalaryTotalDec += totalRegularDec;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                            sheet.Cells[rowCount, 6].Value = (manpointJan * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalRegularJan += manpointJan * beginningValue * commonValue;
                            regularSalaryTotalJan += totalRegularJan;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                            sheet.Cells[rowCount, 7].Value = (manpointFeb * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalRegularFeb += manpointFeb * beginningValue * commonValue;
                            regularSalaryTotalFeb += totalRegularFeb;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                            sheet.Cells[rowCount, 8].Value = (manpointMar * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalRegularMar += manpointMar * beginningValue * commonValue;
                            regularSalaryTotalMar += totalRegularMar;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                            sheet.Cells[rowCount, 9].Value = (manpointApr * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalRegularApr += manpointApr * beginningValue * commonValue;
                            regularSalaryTotalApr += totalRegularApr;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                            sheet.Cells[rowCount, 10].Value = (manpointMay * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalRegularMay += manpointMay * beginningValue * commonValue;
                            regularSalaryTotalMay += totalRegularMay;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                            sheet.Cells[rowCount, 11].Value = (manpointJun * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalRegularJun += manpointJun * beginningValue * commonValue;
                            regularSalaryTotalJun += totalRegularJun;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }



                            double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                            sheet.Cells[rowCount, 12].Value = (manpointJul * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalRegularJul += manpointJul * beginningValue * commonValue;
                            regularSalaryTotalJul += totalRegularJul;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                            sheet.Cells[rowCount, 13].Value = (manpointAug * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalRegularAug += manpointAug * beginningValue * commonValue;
                            regularSalaryTotalAug += totalRegularAug;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                            sheet.Cells[rowCount, 14].Value = (manpointSep * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalRegularSep += manpointSep * beginningValue * commonValue;
                            regularSalaryTotalSep += totalRegularSep;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                        }
                        // salary allowance (fixed)
                        if (salaryTypeCount == 3)
                        {


                            double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                            double beginningValue = _salaryBLL.GetGradeUnitPriceType(departmentId, 2, 2022, item.Grade.Id).GradeLowPoints;
                            double commonValue = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).SalaryIncreaseRate);

                            sheet.Cells[rowCount, 3].Value = (manpointOct * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalFixedOct += manpointOct * beginningValue * commonValue;
                            fixedSalaryTotalOct += totalFixedOct;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                            sheet.Cells[rowCount, 4].Value = (manpointNov * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalFixedNov += manpointNov * beginningValue * commonValue;
                            fixedSalaryTotalNov += totalFixedNov;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                            sheet.Cells[rowCount, 5].Value = (manpointDec * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalFixedDec += manpointDec * beginningValue * commonValue;
                            fixedSalaryTotalDec += totalFixedDec;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                            sheet.Cells[rowCount, 6].Value = (manpointJan * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalFixedJan += manpointJan * beginningValue * commonValue;
                            fixedSalaryTotalJan += totalFixedJan;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                            sheet.Cells[rowCount, 7].Value = (manpointFeb * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalFixedFeb += manpointFeb * beginningValue * commonValue;
                            fixedSalaryTotalFeb += totalFixedFeb;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                            sheet.Cells[rowCount, 8].Value = (manpointMar * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalFixedMar += manpointMar * beginningValue * commonValue;
                            fixedSalaryTotalMar += totalFixedMar;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                            sheet.Cells[rowCount, 9].Value = (manpointApr * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalFixedApr += manpointApr * beginningValue * commonValue;
                            fixedSalaryTotalApr += totalFixedApr;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                            sheet.Cells[rowCount, 10].Value = (manpointMay * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalFixedMay += manpointMay * beginningValue * commonValue;
                            fixedSalaryTotalMay += totalFixedMay;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                            sheet.Cells[rowCount, 11].Value = (manpointJun * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalFixedJun += manpointJun * beginningValue * commonValue;
                            fixedSalaryTotalJun += totalFixedJun;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }



                            double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                            sheet.Cells[rowCount, 12].Value = (manpointJul * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalFixedJul += manpointJul * beginningValue * commonValue;
                            fixedSalaryTotalJul += totalFixedJul;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                            sheet.Cells[rowCount, 13].Value = (manpointAug * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalFixedAug += manpointAug * beginningValue * commonValue;
                            fixedSalaryTotalAug += totalFixedAug;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                            sheet.Cells[rowCount, 14].Value = (manpointSep * beginningValue * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalFixedSep += manpointSep * beginningValue * commonValue;
                            fixedSalaryTotalSep += totalFixedSep;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                        }
                        // salary allowance (overtime)
                        if (salaryTypeCount == 4)
                        {
                            CommonMaster commonMaster = _commonMasterBLL.GetCommonMasters().Where(cm => cm.GradeId == item.Grade.Id).SingleOrDefault();
                            double salaryMasterOverTimeLowPoints = _exportBLL.GetGradeUnitPriceTypes(item.Grade.Id, department.Id, 2022, 3).SingleOrDefault().GradeLowPoints;

                            double calculatedOverTimeOct = GetTotalOverTime(salaryAssignmentDtos, 10, item.Grade.Id);
                            double manPointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                            //double overTimeOct = 0;
                            if (calculatedOverTimeOct > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                            {
                                totalOverOct = manPointOct * (calculatedOverTimeOct - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                            }
                            overTimeSalaryTotalOct += totalOverOct;

                            sheet.Cells[rowCount, 3].Value = totalOverOct.ToString("N0");
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double calculatedOverTimeNov = GetTotalOverTime(salaryAssignmentDtos, 11, item.Grade.Id);
                            double manPointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);

                            if (calculatedOverTimeNov > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                            {
                                totalOverNov = manPointNov * (calculatedOverTimeNov - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                            }
                            overTimeSalaryTotalNov += totalOverNov;

                            sheet.Cells[rowCount, 4].Value = totalOverNov.ToString("N0");
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double calculatedOverTimeDec = GetTotalOverTime(salaryAssignmentDtos, 12, item.Grade.Id);
                            double manPointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);

                            if (calculatedOverTimeDec > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                            {
                                totalOverDec = manPointDec * (calculatedOverTimeDec - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                            }
                            overTimeSalaryTotalDec += totalOverDec;

                            sheet.Cells[rowCount, 5].Value = totalOverDec.ToString("N0");
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double calculatedOverTimeJan = GetTotalOverTime(salaryAssignmentDtos, 1, item.Grade.Id);
                            double manPointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                            if (calculatedOverTimeJan > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                            {
                                totalOverJan = manPointJan * (calculatedOverTimeJan - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                            }
                            overTimeSalaryTotalJan += totalOverJan;
                            sheet.Cells[rowCount, 6].Value = totalOverJan.ToString("N0");
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double calculatedOverTimeFeb = GetTotalOverTime(salaryAssignmentDtos, 2, item.Grade.Id);
                            double manPointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);

                            if (calculatedOverTimeFeb > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                            {
                                totalOverFeb = manPointFeb * (calculatedOverTimeFeb - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                            }
                            overTimeSalaryTotalFeb += totalOverFeb;
                            sheet.Cells[rowCount, 7].Value = totalOverFeb.ToString("N0");
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double calculatedOverTimeMar = GetTotalOverTime(salaryAssignmentDtos, 3, item.Grade.Id);
                            double manPointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);

                            if (calculatedOverTimeMar > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                            {
                                totalOverMar = manPointMar * (calculatedOverTimeMar - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                            }
                            overTimeSalaryTotalMar += totalOverMar;
                            sheet.Cells[rowCount, 8].Value = totalOverMar.ToString("N0");
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double calculatedOverTimeApr = GetTotalOverTime(salaryAssignmentDtos, 4, item.Grade.Id);
                            double manPointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);

                            if (calculatedOverTimeApr > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                            {
                                totalOverApr = manPointApr * (calculatedOverTimeApr - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                            }
                            overTimeSalaryTotalApr += totalOverApr;
                            sheet.Cells[rowCount, 9].Value = totalOverApr.ToString("N0");
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double calculatedOverTimeMay = GetTotalOverTime(salaryAssignmentDtos, 5, item.Grade.Id);
                            double manPointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);

                            if (calculatedOverTimeMay > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                            {
                                totalOverMay = manPointMay * (calculatedOverTimeMay - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                            }
                            overTimeSalaryTotalMay += totalOverMay;
                            sheet.Cells[rowCount, 10].Value = totalOverMay.ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double calculatedOverTimeJun = GetTotalOverTime(salaryAssignmentDtos, 6, item.Grade.Id);
                            double manPointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);

                            if (calculatedOverTimeJun > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                            {
                                totalOverJun = manPointJun * (calculatedOverTimeJun - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                            }
                            overTimeSalaryTotalJun += totalOverJun;
                            sheet.Cells[rowCount, 11].Value = totalOverJun.ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double calculatedOverTimeJul = GetTotalOverTime(salaryAssignmentDtos, 7, item.Grade.Id);
                            double manPointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                            if (calculatedOverTimeJul > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                            {
                                totalOverJul = manPointJul * (calculatedOverTimeJul - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                            }
                            overTimeSalaryTotalJul += totalOverJul;
                            sheet.Cells[rowCount, 12].Value = totalOverJul.ToString("N0");
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double calculatedOverTimeAug = GetTotalOverTime(salaryAssignmentDtos, 8, item.Grade.Id);
                            double manPointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);

                            if (calculatedOverTimeAug > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                            {
                                totalOverAug = manPointAug * (calculatedOverTimeAug - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                            }
                            overTimeSalaryTotalAug += totalOverAug;
                            sheet.Cells[rowCount, 13].Value = totalOverAug.ToString("N0");
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            double calculatedOverTimeSept = GetTotalOverTime(salaryAssignmentDtos, 9, item.Grade.Id);
                            double manPointSept = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);

                            if (calculatedOverTimeSept > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                            {
                                totalOverSep = manPointSept * (calculatedOverTimeSept - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                            }
                            overTimeSalaryTotalSep += totalOverSep;

                            sheet.Cells[rowCount, 14].Value = totalOverSep.ToString("N0");
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                        }
                        // salary allowance (total)
                        if (salaryTypeCount == 5)
                        {
                            sheet.Cells[rowCount, 3].Value = (totalRegularOct + totalFixedOct + totalOverOct).ToString("N0");
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalSalaryOct += totalRegularOct + totalFixedOct + totalOverOct;
                            salaryTotalOct += totalSalaryOct;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 4].Value = (totalRegularNov + totalFixedNov + totalOverNov).ToString("N0");
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalSalaryNov += totalRegularNov + totalFixedNov + totalOverNov;
                            salaryTotalNov += totalSalaryNov;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 5].Value = (totalRegularDec + totalFixedDec + totalOverDec).ToString("N0");
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalSalaryDec += totalRegularDec + totalFixedDec + totalOverDec;
                            salaryTotalDec += totalSalaryDec;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 6].Value = (totalRegularJan + totalFixedJan + totalOverJan).ToString("N0");
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalSalaryJan += totalRegularJan + totalFixedJan + totalOverJan;
                            salaryTotalJan += totalSalaryJan;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 7].Value = (totalRegularFeb + totalFixedFeb + totalOverFeb).ToString("N0");
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalSalaryFeb += totalRegularFeb + totalFixedFeb + totalOverFeb;
                            salaryTotalFeb += totalSalaryFeb;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 8].Value = (totalRegularMar + totalFixedMar + totalOverMar).ToString("N0");
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalSalaryMar += totalRegularMar + totalFixedMar + totalOverMar;
                            salaryTotalMar += totalSalaryMar;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 9].Value = (totalRegularApr + totalFixedApr + totalOverApr).ToString("N0");
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalSalaryApr += totalRegularApr + totalFixedApr + totalOverApr;
                            salaryTotalApr += totalSalaryApr;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 10].Value = (totalRegularMay + totalFixedMay + totalOverMay).ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalSalaryMay += totalRegularMay + totalFixedMay + totalOverMay;
                            salaryTotalMay += totalSalaryMay;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 11].Value = (totalRegularJun + totalFixedJun + totalOverJun).ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalSalaryJun += totalRegularJun + totalFixedJun + totalOverJun;
                            salaryTotalJun += totalSalaryJun;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 12].Value = (totalRegularJul + totalFixedJul + totalOverJul).ToString("N0");
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalSalaryJul += totalRegularJul + totalFixedJul + totalOverJul;
                            salaryTotalJul += totalSalaryJul;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 13].Value = (totalRegularAug + totalFixedAug + totalOverAug).ToString("N0");
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalSalaryAug += totalRegularAug + totalFixedAug + totalOverAug;
                            salaryTotalAug += totalSalaryAug;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 14].Value = (totalRegularSep + totalFixedSep + totalOverSep).ToString("N0");
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalSalarySep += totalRegularSep + totalFixedSep + totalOverSep;
                            salaryTotalSep += totalSalarySep;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                        }
                        // Miscellneous Wages
                        if (salaryTypeCount == 6)
                        {


                            double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                            double beginningValue = _salaryBLL.GetGradeUnitPriceType(departmentId, salaryType.Id, 2022, item.Grade.Id).GradeLowPoints;
                            //double commonValue = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).SalaryIncreaseRate);

                            sheet.Cells[rowCount, 3].Value = (manpointOct * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesOct += manpointOct * beginningValue;
                            mwTotalOct += mWagesOct;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                            sheet.Cells[rowCount, 4].Value = (manpointNov * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesNov += manpointNov * beginningValue;
                            mwTotalNov += mWagesNov;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                            sheet.Cells[rowCount, 5].Value = (manpointDec * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesDec += manpointDec * beginningValue;
                            mwTotalDec += mWagesDec;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                            sheet.Cells[rowCount, 6].Value = (manpointJan * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesJan += manpointJan * beginningValue;
                            mwTotalJan += mWagesJan;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                            sheet.Cells[rowCount, 7].Value = (manpointFeb * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesFeb += manpointFeb * beginningValue;
                            mwTotalFeb += mWagesFeb;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                            sheet.Cells[rowCount, 8].Value = (manpointMar * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesMar += manpointMar * beginningValue;
                            mwTotalMar += mWagesMar;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                            sheet.Cells[rowCount, 9].Value = (manpointApr * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesApr += manpointApr * beginningValue;
                            mwTotalApr += mWagesApr;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                            sheet.Cells[rowCount, 10].Value = (manpointMay * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesMay += manpointMay * beginningValue;
                            mwTotalMay += mWagesMay;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                            sheet.Cells[rowCount, 11].Value = (manpointJun * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesJun += manpointJun * beginningValue;
                            mwTotalJun += mWagesJun;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }



                            double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                            sheet.Cells[rowCount, 12].Value = (manpointJul * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesJul += manpointJul * beginningValue;
                            mwTotalJul += mWagesJul;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                            sheet.Cells[rowCount, 13].Value = (manpointAug * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesAug += manpointAug * beginningValue;
                            mwTotalAug += mWagesAug;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                            sheet.Cells[rowCount, 14].Value = (manpointSep * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            mWagesSep += manpointSep * beginningValue;
                            mwTotalSep += mWagesSep;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                        }
                        // Dispatch Fee
                        if (salaryTypeCount == 7)
                        {


                            double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                            double beginningValue = _salaryBLL.GetGradeUnitPriceType(departmentId, salaryType.Id, 2022, item.Grade.Id).GradeLowPoints;


                            sheet.Cells[rowCount, 3].Value = (manpointOct * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            dispatchFeeOct += manpointOct * beginningValue;
                            dispatchFeeTotalOct += dispatchFeeOct;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                            sheet.Cells[rowCount, 4].Value = (manpointNov * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            dispatchFeeNov += manpointNov * beginningValue;
                            dispatchFeeTotalNov += dispatchFeeNov;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                            sheet.Cells[rowCount, 5].Value = (manpointDec * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            dispatchFeeDec += manpointDec * beginningValue;
                            dispatchFeeTotalDec += dispatchFeeDec;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                            sheet.Cells[rowCount, 6].Value = (manpointJan * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            dispatchFeeJan += manpointJan * beginningValue;
                            dispatchFeeTotalJan += dispatchFeeJan;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                            sheet.Cells[rowCount, 7].Value = (manpointFeb * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            dispatchFeeFeb += manpointFeb * beginningValue;
                            dispatchFeeTotalFeb += dispatchFeeFeb;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                            sheet.Cells[rowCount, 8].Value = (manpointMar * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            dispatchFeeMar += manpointMar * beginningValue;
                            dispatchFeeTotalMar += dispatchFeeMar;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                            sheet.Cells[rowCount, 9].Value = (manpointApr * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            dispatchFeeApr += manpointApr * beginningValue;
                            dispatchFeeTotalApr += dispatchFeeApr;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                            sheet.Cells[rowCount, 10].Value = (manpointMay * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            dispatchFeeMay += manpointMay * beginningValue;
                            dispatchFeeTotalMay += dispatchFeeMay;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                            sheet.Cells[rowCount, 11].Value = (manpointJun * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            dispatchFeeJun += manpointJun * beginningValue;
                            dispatchFeeTotalJun += dispatchFeeJun;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }



                            double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                            sheet.Cells[rowCount, 12].Value = (manpointJul * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            dispatchFeeJul += manpointJul * beginningValue;
                            dispatchFeeTotalJul += dispatchFeeJul;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                            sheet.Cells[rowCount, 13].Value = (manpointAug * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            dispatchFeeAug += manpointAug * beginningValue;
                            dispatchFeeTotalAug += dispatchFeeAug;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                            sheet.Cells[rowCount, 14].Value = (manpointSep * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            dispatchFeeSep += manpointSep * beginningValue;
                            dispatchFeeTotalSep += dispatchFeeSep;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                        }
                        // Provision for Employee Bonus
                        if (salaryTypeCount == 8)
                        {
                            double overWorkFixedTime = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).OverWorkFixedTime);
                            double bonusReserveConstant = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).BonusReserveConstant);

                            double valueOct = (totalRegularOct * overWorkFixedTime) / bonusReserveConstant;
                            valueOct = Double.IsNaN(valueOct) ? 0 : valueOct;
                            valueOct = Double.IsInfinity(valueOct) ? 0 : valueOct;
                            employeeBonusOct += valueOct;
                            employeeBonusTotalOct += employeeBonusOct;
                            sheet.Cells[rowCount, 3].Value = valueOct.ToString("N0");
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            double valueNov = (totalRegularNov * overWorkFixedTime) / bonusReserveConstant;
                            valueNov = Double.IsNaN(valueNov) ? 0 : valueNov;
                            valueNov = Double.IsInfinity(valueNov) ? 0 : valueNov;
                            employeeBonusNov += valueNov;
                            employeeBonusTotalNov += employeeBonusNov;
                            sheet.Cells[rowCount, 4].Value = valueNov.ToString("N0");
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            double valueDec = (totalRegularDec * overWorkFixedTime) / bonusReserveConstant;
                            valueDec = Double.IsNaN(valueDec) ? 0 : valueDec;
                            valueDec = Double.IsInfinity(valueDec) ? 0 : valueDec;
                            employeeBonusDec += valueDec;
                            employeeBonusTotalDec += employeeBonusDec;
                            sheet.Cells[rowCount, 5].Value = valueDec.ToString("N0");
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            double valueJan = (totalRegularJan * overWorkFixedTime) / bonusReserveConstant;
                            valueJan = Double.IsNaN(valueJan) ? 0 : valueJan;
                            valueJan = Double.IsInfinity(valueJan) ? 0 : valueJan;
                            employeeBonusJan += valueJan;
                            employeeBonusTotalJan += employeeBonusJan;
                            sheet.Cells[rowCount, 6].Value = valueJan.ToString("N0");
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double valueFeb = (totalRegularFeb * overWorkFixedTime) / bonusReserveConstant;
                            valueFeb = Double.IsNaN(valueFeb) ? 0 : valueFeb;
                            valueFeb = Double.IsInfinity(valueFeb) ? 0 : valueFeb;
                            employeeBonusFeb += valueFeb;
                            employeeBonusTotalFeb += employeeBonusFeb;
                            sheet.Cells[rowCount, 7].Value = valueFeb.ToString("N0");
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            double valueMar = (totalRegularMar * overWorkFixedTime) / bonusReserveConstant;
                            valueMar = Double.IsNaN(valueMar) ? 0 : valueMar;
                            valueMar = Double.IsInfinity(valueMar) ? 0 : valueMar;
                            employeeBonusMar += valueMar;
                            employeeBonusTotalMar += employeeBonusMar;
                            sheet.Cells[rowCount, 8].Value = valueMar.ToString("N0");
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            double valueApr = (totalRegularApr * overWorkFixedTime) / bonusReserveConstant;
                            valueApr = Double.IsNaN(valueApr) ? 0 : valueApr;
                            valueApr = Double.IsInfinity(valueApr) ? 0 : valueApr;
                            employeeBonusApr += valueApr;
                            employeeBonusTotalApr += employeeBonusApr;
                            sheet.Cells[rowCount, 9].Value = valueApr.ToString("N0");
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            double valueMay = (totalRegularMay * overWorkFixedTime) / bonusReserveConstant;
                            valueMay = Double.IsNaN(valueMay) ? 0 : valueMay;
                            valueMay = Double.IsInfinity(valueMay) ? 0 : valueMay;
                            employeeBonusMay += valueMay;
                            employeeBonusTotalMay += employeeBonusMay;
                            sheet.Cells[rowCount, 10].Value = valueMay.ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            double valueJun = (totalRegularJun * overWorkFixedTime) / bonusReserveConstant;
                            valueJun = Double.IsNaN(valueJun) ? 0 : valueJun;
                            valueJun = Double.IsInfinity(valueJun) ? 0 : valueJun;
                            employeeBonusJun += valueJun;
                            employeeBonusTotalJun += employeeBonusJun;
                            sheet.Cells[rowCount, 11].Value = valueJun.ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double valueJul = (totalRegularJul * overWorkFixedTime) / bonusReserveConstant;
                            valueJul = Double.IsNaN(valueJul) ? 0 : valueJul;
                            valueJul = Double.IsInfinity(valueJul) ? 0 : valueJul;
                            employeeBonusJul += valueJul;
                            employeeBonusTotalJul += employeeBonusJul;
                            sheet.Cells[rowCount, 12].Value = valueJul.ToString("N0");
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            double valueAug = (totalRegularAug * overWorkFixedTime) / bonusReserveConstant;
                            valueAug = Double.IsNaN(valueAug) ? 0 : valueAug;
                            valueAug = Double.IsInfinity(valueAug) ? 0 : valueAug;
                            employeeBonusAug += valueAug;
                            employeeBonusTotalAug += employeeBonusAug;
                            sheet.Cells[rowCount, 13].Value = valueAug.ToString("N0");
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double valueSep = (totalRegularSep * overWorkFixedTime) / bonusReserveConstant;
                            valueSep = Double.IsNaN(valueSep) ? 0 : valueSep;
                            valueSep = Double.IsInfinity(valueSep) ? 0 : valueSep;
                            employeeBonusSep += valueSep;
                            employeeBonusTotalSep += employeeBonusSep;
                            sheet.Cells[rowCount, 14].Value = valueSep.ToString("N0");
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                        }
                        //commuting expenses
                        if (salaryTypeCount == 9)
                        {


                            double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                            double beginningValue = _salaryBLL.GetGradeUnitPriceType(departmentId, 4, 2022, item.Grade.Id).GradeLowPoints;
                            //double commonValue = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).SalaryIncreaseRate);

                            sheet.Cells[rowCount, 3].Value = (manpointOct * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            commutingExpensesOct += manpointOct * beginningValue;
                            commutingExpensesTotalOct += commutingExpensesOct;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                            sheet.Cells[rowCount, 4].Value = (manpointNov * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            commutingExpensesNov += manpointNov * beginningValue;
                            commutingExpensesTotalNov += commutingExpensesNov;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                            sheet.Cells[rowCount, 5].Value = (manpointDec * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            commutingExpensesDec += manpointDec * beginningValue;
                            commutingExpensesTotalDec += commutingExpensesDec;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                            sheet.Cells[rowCount, 6].Value = (manpointJan * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            commutingExpensesJan += manpointJan * beginningValue;
                            commutingExpensesTotalJan += commutingExpensesJan;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                            sheet.Cells[rowCount, 7].Value = (manpointFeb * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            commutingExpensesFeb += manpointFeb * beginningValue;
                            commutingExpensesTotalFeb += commutingExpensesFeb;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                            sheet.Cells[rowCount, 8].Value = (manpointMar * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            commutingExpensesMar += manpointMar * beginningValue;
                            commutingExpensesTotalMar += commutingExpensesMar;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                            sheet.Cells[rowCount, 9].Value = (manpointApr * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            commutingExpensesApr += manpointApr * beginningValue;
                            commutingExpensesTotalApr += commutingExpensesApr;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                            sheet.Cells[rowCount, 10].Value = (manpointMay * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            commutingExpensesMay += manpointMay * beginningValue;
                            commutingExpensesTotalMay += commutingExpensesMay;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                            sheet.Cells[rowCount, 11].Value = (manpointJun * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            commutingExpensesJun += manpointJun * beginningValue;
                            commutingExpensesTotalJun += commutingExpensesJun;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }



                            double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                            sheet.Cells[rowCount, 12].Value = (manpointJul * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            commutingExpensesJul += manpointJul * beginningValue;
                            commutingExpensesTotalJul += commutingExpensesJul;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                            sheet.Cells[rowCount, 13].Value = (manpointAug * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            commutingExpensesAug += manpointAug * beginningValue;
                            commutingExpensesTotalAug += commutingExpensesAug;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                            sheet.Cells[rowCount, 14].Value = (manpointSep * beginningValue).ToString("N0");
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            commutingExpensesSep += manpointSep * beginningValue;
                            commutingExpensesTotalSep += commutingExpensesSep;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                        }
                        // welfare expenses
                        if (salaryTypeCount == 10)
                        {
                            double commonValue = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).WelfareCostRatio);
                            sheet.Cells[rowCount, 3].Value = ((totalSalaryOct + mWagesOct + commutingExpensesOct) * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            wExpensesOct += (totalSalaryOct + mWagesOct + commutingExpensesOct) * commonValue;
                            welfareExpensesTotalOct += wExpensesOct;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 4].Value = ((totalSalaryNov + mWagesNov + commutingExpensesNov) * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            wExpensesNov += (totalSalaryNov + mWagesNov + commutingExpensesNov) * commonValue;
                            welfareExpensesTotalNov += wExpensesNov;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 5].Value = ((totalSalaryDec + mWagesDec + commutingExpensesDec) * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            wExpensesDec += (totalSalaryDec + mWagesDec + commutingExpensesDec) * commonValue;
                            welfareExpensesTotalDec += wExpensesDec;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 6].Value = ((totalSalaryJan + mWagesJan + commutingExpensesJan) * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            wExpensesJan += (totalSalaryJan + mWagesJan + commutingExpensesJan) * commonValue;
                            welfareExpensesTotalJan += wExpensesJan;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 7].Value = ((totalSalaryFeb + mWagesFeb + commutingExpensesFeb) * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            wExpensesFeb += (totalSalaryFeb + mWagesFeb + commutingExpensesFeb) * commonValue;
                            welfareExpensesTotalFeb += wExpensesFeb;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 8].Value = ((totalSalaryMar + mWagesMar + commutingExpensesMar) * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            wExpensesMar += (totalSalaryMar + mWagesMar + commutingExpensesMar) * commonValue;
                            welfareExpensesTotalMar += wExpensesMar;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 9].Value = ((totalSalaryApr + mWagesApr + commutingExpensesApr) * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            wExpensesApr += (totalSalaryApr + mWagesApr + commutingExpensesApr) * commonValue;
                            welfareExpensesTotalApr += wExpensesApr;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 10].Value = ((totalSalaryMay + mWagesMay + commutingExpensesMay) * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            wExpensesMay += (totalSalaryMay + mWagesMay + commutingExpensesMay) * commonValue;
                            welfareExpensesTotalMay += wExpensesMay;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 11].Value = ((totalSalaryJun + mWagesJun + commutingExpensesJun) * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            wExpensesJun += (totalSalaryJun + mWagesJun + commutingExpensesJun) * commonValue;
                            welfareExpensesTotalJun += wExpensesJun;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 12].Value = ((totalSalaryJul + mWagesJul + commutingExpensesJul) * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            wExpensesJul += (totalSalaryJul + mWagesJul + commutingExpensesJul) * commonValue;
                            welfareExpensesTotalJul += wExpensesJul;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 13].Value = ((totalSalaryAug + mWagesAug + commutingExpensesAug) * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            wExpensesAug += (totalSalaryAug + mWagesAug + commutingExpensesAug) * commonValue;
                            welfareExpensesTotalAug += wExpensesAug;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 14].Value = ((totalSalarySep + mWagesSep + commutingExpensesSep) * commonValue).ToString("N0");
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            wExpensesSep += (totalSalarySep + mWagesSep + commutingExpensesSep) * commonValue;
                            welfareExpensesTotalSep += wExpensesSep;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                        }
                        // welfare expenses bonus
                        if (salaryTypeCount == 11)
                        {
                            sheet.Cells[rowCount, 3].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 4].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 5].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 6].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 7].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 8].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 9].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 10].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 11].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 12].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 13].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 14].Value = 0.ToString("N0");
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                        }
                        // total statutory
                        if (salaryTypeCount == 12)
                        {
                            sheet.Cells[rowCount, 3].Value = (wExpensesOct + wExpBonusOct).ToString("N0");
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalStatutoryOct += wExpensesOct + wExpBonusOct;
                            statutoryTotalOct += totalStatutoryOct;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 4].Value = (wExpensesNov + wExpBonusNov).ToString("N0");
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalStatutoryNov += wExpensesNov + wExpBonusNov;
                            statutoryTotalNov += totalStatutoryNov;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 5].Value = (wExpensesDec + wExpBonusDec).ToString("N0");
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalStatutoryDec += wExpensesDec + wExpBonusDec;
                            statutoryTotalDec += totalStatutoryDec;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 6].Value = (wExpensesJan + wExpBonusJan).ToString("N0");
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalStatutoryJan += wExpensesJan + wExpBonusJan;
                            statutoryTotalJan += totalStatutoryJan;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 7].Value = (wExpensesFeb + wExpBonusFeb).ToString("N0");
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalStatutoryFeb += wExpensesFeb + wExpBonusFeb;
                            statutoryTotalFeb += totalStatutoryFeb;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 8].Value = (wExpensesMar + wExpBonusMar).ToString("N0");
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalStatutoryMar += wExpensesMar + wExpBonusMar;
                            statutoryTotalMar += totalStatutoryMar;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 9].Value = (wExpensesApr + wExpBonusApr).ToString("N0");
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalStatutoryApr += wExpensesApr + wExpBonusApr;
                            statutoryTotalApr += totalStatutoryApr;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 10].Value = (wExpensesMay + wExpBonusMay).ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalStatutoryMay += wExpensesMay + wExpBonusMay;
                            statutoryTotalMay += totalStatutoryMay;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 11].Value = (wExpensesJun + wExpBonusJun).ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalStatutoryJun += wExpensesJun + wExpBonusJun;
                            statutoryTotalJun += totalStatutoryJun;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 12].Value = (wExpensesJul + wExpBonusJul).ToString("N0");
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalStatutoryJul += wExpensesJul + wExpBonusJul;
                            statutoryTotalJul += totalStatutoryJul;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            sheet.Cells[rowCount, 13].Value = (wExpensesAug + wExpBonusAug).ToString("N0");
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalStatutoryAug += wExpensesAug + wExpBonusAug;
                            statutoryTotalAug += totalStatutoryAug;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }


                            sheet.Cells[rowCount, 14].Value = (wExpensesSep + wExpBonusSep).ToString("N0");
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            totalStatutorySep += wExpensesSep + wExpBonusSep;
                            statutoryTotalSep += totalStatutorySep;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                        }
                        // total expenses
                        if (salaryTypeCount == 13)
                        {
                            expensesTotalOct += totalSalaryOct + mWagesOct + dispatchFeeOct + employeeBonusOct + commutingExpensesOct + totalStatutoryOct;
                            sheet.Cells[rowCount, 3].Value = (totalSalaryOct + mWagesOct + dispatchFeeOct + employeeBonusOct + commutingExpensesOct + totalStatutoryOct).ToString("N0");
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            expensesTotalNov += totalSalaryNov + mWagesNov + dispatchFeeNov + employeeBonusNov + commutingExpensesNov + totalStatutoryNov;
                            sheet.Cells[rowCount, 4].Value = (totalSalaryNov + mWagesNov + dispatchFeeNov + employeeBonusNov + commutingExpensesNov + totalStatutoryNov).ToString("N0");
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            expensesTotalDec += totalSalaryDec + mWagesDec + dispatchFeeDec + employeeBonusDec + commutingExpensesDec + totalStatutoryDec;
                            sheet.Cells[rowCount, 5].Value = (totalSalaryDec + mWagesDec + dispatchFeeDec + employeeBonusDec + commutingExpensesDec + totalStatutoryDec).ToString("N0");
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            expensesTotalJan += totalSalaryJan + mWagesJan + dispatchFeeJan + employeeBonusJan + commutingExpensesJan + totalStatutoryJan;
                            sheet.Cells[rowCount, 6].Value = (totalSalaryJan + mWagesJan + dispatchFeeJan + employeeBonusJan + commutingExpensesJan + totalStatutoryJan).ToString("N0");
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            expensesTotalFeb += totalSalaryFeb + mWagesFeb + dispatchFeeFeb + employeeBonusFeb + commutingExpensesFeb + totalStatutoryFeb;
                            sheet.Cells[rowCount, 7].Value = (totalSalaryFeb + mWagesFeb + dispatchFeeFeb + employeeBonusFeb + commutingExpensesFeb + totalStatutoryFeb).ToString("N0");
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            expensesTotalMar += totalSalaryMar + mWagesMar + dispatchFeeMar + employeeBonusMar + commutingExpensesMar + totalStatutoryMar;
                            sheet.Cells[rowCount, 8].Value = (totalSalaryMar + mWagesMar + dispatchFeeMar + employeeBonusMar + commutingExpensesMar + totalStatutoryMar).ToString("N0");
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            expensesTotalApr += totalSalaryApr + mWagesApr + dispatchFeeApr + employeeBonusApr + commutingExpensesApr + totalStatutoryApr;
                            sheet.Cells[rowCount, 9].Value = (totalSalaryApr + mWagesApr + dispatchFeeApr + employeeBonusApr + commutingExpensesApr + totalStatutoryApr).ToString("N0");
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            expensesTotalMay += totalSalaryMay + mWagesMay + dispatchFeeMay + employeeBonusMay + commutingExpensesMay + totalStatutoryMay;
                            sheet.Cells[rowCount, 10].Value = (totalSalaryMay + mWagesMay + dispatchFeeMay + employeeBonusMay + commutingExpensesMay + totalStatutoryMay).ToString("N0");
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            expensesTotalJun += totalSalaryJun + mWagesJun + dispatchFeeJun + employeeBonusJun + commutingExpensesJun + totalStatutoryJun;
                            sheet.Cells[rowCount, 11].Value = (totalSalaryJun + mWagesJun + dispatchFeeJun + employeeBonusJun + commutingExpensesJun + totalStatutoryJun).ToString("N0");
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            expensesTotalJul += totalSalaryJul + mWagesJul + dispatchFeeJul + employeeBonusJul + commutingExpensesJul + totalStatutoryJul;
                            sheet.Cells[rowCount, 12].Value = (totalSalaryJul + mWagesJul + dispatchFeeJul + employeeBonusJul + commutingExpensesJul + totalStatutoryJul).ToString("N0");
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }
                            expensesTotalAug += totalSalaryAug + mWagesAug + dispatchFeeAug + employeeBonusAug + commutingExpensesAug + totalStatutoryAug;
                            sheet.Cells[rowCount, 13].Value = (totalSalaryAug + mWagesAug + dispatchFeeAug + employeeBonusAug + commutingExpensesAug + totalStatutoryAug).ToString("N0");
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                            expensesTotalSep += totalSalarySep + mWagesSep + dispatchFeeSep + employeeBonusSep + commutingExpensesSep + totalStatutorySep;
                            sheet.Cells[rowCount, 14].Value = (totalSalarySep + mWagesSep + dispatchFeeSep + employeeBonusSep + commutingExpensesSep + totalStatutorySep).ToString("N0");
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            if (count % 2 == 0)
                            {
                                // even
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                            }
                            else
                            {
                                // odd
                                sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 252, 213, 180);
                            }

                        }


                        rowCount++;
                        innerCount++;
                        salaryTypeCount++;
                    }

                    count++;

                }
                #endregion


                #region salary types total
                int salaryTypeSummeryCount = 1;
                foreach (var item in _unitPriceTypeBLL.GetAllSalaryTypes())
                {
                    sheet.Cells[rowCountForFuture, 1].Value = item.SalaryTypeName;
                    // executive
                    if (salaryTypeSummeryCount == 1)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = 0;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = 0;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = 0;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = 0;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = 0;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = 0;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = 0;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = 0;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = 0;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = 0;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = 0;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = 0;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // regular
                    if (salaryTypeSummeryCount == 2)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = regularSalaryTotalOct.ToString("N0");
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = regularSalaryTotalNov.ToString("N0");
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = regularSalaryTotalDec.ToString("N0");
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = regularSalaryTotalJan.ToString("N0");
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = regularSalaryTotalFeb.ToString("N0");
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = regularSalaryTotalMar.ToString("N0");
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = regularSalaryTotalApr.ToString("N0");
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = regularSalaryTotalMay.ToString("N0");
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = regularSalaryTotalJun.ToString("N0");
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = regularSalaryTotalJul.ToString("N0");
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = regularSalaryTotalAug.ToString("N0");
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = regularSalaryTotalSep.ToString("N0");
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // fixed
                    if (salaryTypeSummeryCount == 3)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = fixedSalaryTotalOct.ToString("N0");
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = fixedSalaryTotalNov.ToString("N0");
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = fixedSalaryTotalDec.ToString("N0");
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = fixedSalaryTotalJan.ToString("N0");
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = fixedSalaryTotalFeb.ToString("N0");
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = fixedSalaryTotalMar.ToString("N0");
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = fixedSalaryTotalApr.ToString("N0");
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = fixedSalaryTotalMay;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = fixedSalaryTotalJun.ToString("N0");
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = fixedSalaryTotalJul.ToString("N0");
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = fixedSalaryTotalAug.ToString("N0");
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = fixedSalaryTotalSep.ToString("N0");
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // overtime
                    if (salaryTypeSummeryCount == 4)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = overTimeSalaryTotalOct.ToString("N0");
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = overTimeSalaryTotalNov.ToString("N0");
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = overTimeSalaryTotalDec.ToString("N0");
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = overTimeSalaryTotalJan.ToString("N0");
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = overTimeSalaryTotalFeb.ToString("N0");
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = overTimeSalaryTotalMar.ToString("N0");
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = overTimeSalaryTotalApr.ToString("N0");
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = overTimeSalaryTotalMay.ToString("N0");
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = overTimeSalaryTotalJun.ToString("N0");
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = overTimeSalaryTotalJul.ToString("N0");
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = overTimeSalaryTotalAug.ToString("N0");
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = overTimeSalaryTotalSep.ToString("N0");
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // total salary
                    if (salaryTypeSummeryCount == 5)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = salaryTotalOct.ToString("N0");
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = salaryTotalNov.ToString("N0");
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = salaryTotalDec.ToString("N0");
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = salaryTotalJan.ToString("N0");
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = salaryTotalFeb.ToString("N0");
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = salaryTotalMar.ToString("N0");
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = salaryTotalApr.ToString("N0");
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = salaryTotalMay.ToString("N0");
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = salaryTotalJun.ToString("N0");
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = salaryTotalJul.ToString("N0");
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = salaryTotalAug.ToString("N0");
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = salaryTotalSep.ToString("N0");
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // miscellaneous wages
                    if (salaryTypeSummeryCount == 6)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = mwTotalOct.ToString("N0");
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = mwTotalNov.ToString("N0");
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = mwTotalDec.ToString("N0");
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = mwTotalJan.ToString("N0");
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = mwTotalFeb.ToString("N0");
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = mwTotalMar.ToString("N0");
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = mwTotalApr.ToString("N0");
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = mwTotalMay.ToString("N0");
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = mwTotalJun.ToString("N0");
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = mwTotalJul.ToString("N0");
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = mwTotalAug.ToString("N0");
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = mwTotalSep.ToString("N0");
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // dispatch fee
                    if (salaryTypeSummeryCount == 7)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = dispatchFeeTotalOct.ToString("N0");
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = dispatchFeeTotalNov.ToString("N0");
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = dispatchFeeTotalDec.ToString("N0");
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = dispatchFeeTotalJan.ToString("N0");
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = dispatchFeeTotalFeb.ToString("N0");
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = dispatchFeeTotalMar.ToString("N0");
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = dispatchFeeTotalApr.ToString("N0");
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = dispatchFeeTotalMay.ToString("N0");
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = dispatchFeeTotalJun.ToString("N0");
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = dispatchFeeTotalJul.ToString("N0");
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = dispatchFeeTotalAug.ToString("N0");
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = dispatchFeeTotalSep.ToString("N0");
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // provision for employee bonus
                    if (salaryTypeSummeryCount == 8)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = employeeBonusTotalOct.ToString("N0");
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = employeeBonusTotalNov.ToString("N0");
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = employeeBonusTotalDec.ToString("N0");
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = employeeBonusTotalJan.ToString("N0");
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = employeeBonusTotalFeb.ToString("N0");
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = employeeBonusTotalMar.ToString("N0");
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = employeeBonusTotalApr.ToString("N0");
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = employeeBonusTotalMay.ToString("N0");
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = employeeBonusTotalJun.ToString("N0");
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = employeeBonusTotalJul.ToString("N0");
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = employeeBonusTotalAug.ToString("N0");
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = employeeBonusTotalSep.ToString("N0");
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // commuting expenses
                    if (salaryTypeSummeryCount == 9)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = commutingExpensesTotalOct.ToString("N0");
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = commutingExpensesTotalNov.ToString("N0");
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = commutingExpensesTotalDec.ToString("N0");
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = commutingExpensesTotalJan.ToString("N0");
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = commutingExpensesTotalFeb.ToString("N0");
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = commutingExpensesTotalMar.ToString("N0");
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = commutingExpensesTotalApr.ToString("N0");
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = commutingExpensesTotalMay.ToString("N0");
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = commutingExpensesTotalJun.ToString("N0");
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = commutingExpensesTotalJul.ToString("N0");
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = commutingExpensesTotalAug.ToString("N0");
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = commutingExpensesTotalSep.ToString("N0");
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // salary statutory welfare expenses
                    if (salaryTypeSummeryCount == 10)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = welfareExpensesTotalOct.ToString("N0");
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = welfareExpensesTotalNov.ToString("N0");
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = welfareExpensesTotalDec.ToString("N0");
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = welfareExpensesTotalJan.ToString("N0");
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = welfareExpensesTotalFeb.ToString("N0");
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = welfareExpensesTotalMar.ToString("N0");
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = welfareExpensesTotalApr.ToString("N0");
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = welfareExpensesTotalMay.ToString("N0");
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = welfareExpensesTotalJun.ToString("N0");
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = welfareExpensesTotalJul.ToString("N0");
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = welfareExpensesTotalAug.ToString("N0");
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = welfareExpensesTotalSep.ToString("N0");
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // provision for statotory welfare expenses for bonus
                    if (salaryTypeSummeryCount == 11)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = welfareExpensesBonusTotalOct.ToString("N0");
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = welfareExpensesBonusTotalNov.ToString("N0");
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = welfareExpensesBonusTotalDec.ToString("N0");
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = welfareExpensesBonusTotalJan.ToString("N0");
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = welfareExpensesBonusTotalFeb.ToString("N0");
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = welfareExpensesBonusTotalMar.ToString("N0");
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = welfareExpensesBonusTotalApr.ToString("N0");
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = welfareExpensesBonusTotalMay.ToString("N0");
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = welfareExpensesBonusTotalJun.ToString("N0");
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = welfareExpensesBonusTotalJul.ToString("N0");
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = welfareExpensesBonusTotalAug.ToString("N0");
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = welfareExpensesBonusTotalSep.ToString("N0");
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // total statutory benefites
                    if (salaryTypeSummeryCount == 12)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = statutoryTotalOct.ToString("N0");
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = statutoryTotalNov.ToString("N0");
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = statutoryTotalDec.ToString("N0");
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = statutoryTotalJan.ToString("N0");
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = statutoryTotalFeb.ToString("N0");
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = statutoryTotalMar.ToString("N0");
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = statutoryTotalApr.ToString("N0");
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = statutoryTotalMay.ToString("N0");
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = statutoryTotalJun.ToString("N0");
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = statutoryTotalJul.ToString("N0");
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = statutoryTotalAug.ToString("N0");
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = statutoryTotalSep.ToString("N0");
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // total expenses
                    if (salaryTypeSummeryCount == 13)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = expensesTotalOct.ToString("N0");
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = expensesTotalNov.ToString("N0");
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = expensesTotalDec.ToString("N0");
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = expensesTotalJan.ToString("N0");
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = expensesTotalFeb.ToString("N0");
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = expensesTotalMar.ToString("N0");
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = expensesTotalApr.ToString("N0");
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = expensesTotalMay.ToString("N0");
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = expensesTotalJun.ToString("N0");
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = expensesTotalJul.ToString("N0");
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = expensesTotalAug.ToString("N0");
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = expensesTotalSep.ToString("N0");
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }




                    //rowCount++;
                    rowCountForFuture++;
                    salaryTypeSummeryCount++;
                }
                #endregion



                rowCount = rowCount + 2;

                sheet.Cells[rowCount, 1, rowCount, 14].Style.Font.Bold = true;
                sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                sheet.Cells[rowCount, 1].Value = "Costing";
                rowCount++;

                octTotal = 0; novTotal = 0; decTotal = 0; janTotal = 0; febTotal = 0; marTotal = 0; aprTotal = 0; mayTotal = 0; junTotal = 0; julTotal = 0; augTotal = 0; sepTotal = 0;
                // costing
                foreach (var item in salaryAssignmentDtos)
                {
                    sheet.Cells[rowCount, 2].Value = item.Grade.GradeName;


                    foreach (var singleAssignment in item.ForecastAssignmentViewModels)
                    {
                        oct += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 10).Total);
                        nov += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 11).Total);
                        dec += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 12).Total);
                        jan += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 1).Total);
                        feb += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 2).Total);
                        mar += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 3).Total);
                        apr += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 4).Total);
                        may += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 5).Total);
                        jun += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 6).Total);
                        jul += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 7).Total);
                        aug += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 8).Total);
                        sep += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 9).Total);
                    }
                    sheet.Cells[rowCount, 3].Value = oct.ToString("N0");
                    octTotal += oct;
                    sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 4].Value = nov.ToString("N0");
                    novTotal += nov;
                    sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 5].Value = dec.ToString("N0");
                    decTotal += dec;
                    sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 6].Value = jan.ToString("N0");
                    janTotal += jan;
                    sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 7].Value = feb.ToString("N0");
                    febTotal += feb;
                    sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 8].Value = mar.ToString("N0");
                    marTotal += mar;
                    sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 9].Value = apr.ToString("N0");
                    aprTotal += apr;
                    sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 10].Value = may.ToString("N0");
                    mayTotal += may;
                    sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 11].Value = jun.ToString("N0");
                    junTotal += jun;
                    sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 12].Value = jul.ToString("N0");
                    julTotal += jul;
                    sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 13].Value = aug.ToString("N0");
                    augTotal += aug;
                    sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 14].Value = sep.ToString("N0");
                    sepTotal += sep;
                    sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                    rowCount++;
                }


                sheet.Cells[rowCount, 2].Value = "合計";
                sheet.Cells[rowCount, 3].Value = octTotal.ToString("N0");
                sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 4].Value = novTotal.ToString("N0");
                sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 5].Value = decTotal.ToString("N0");
                sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 6].Value = janTotal.ToString("N0");
                sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 7].Value = febTotal.ToString("N0");
                sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 8].Value = marTotal.ToString("N0");
                sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 9].Value = aprTotal.ToString("N0");
                sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 10].Value = mayTotal.ToString("N0");
                sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 11].Value = junTotal.ToString("N0");
                sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 12].Value = julTotal.ToString("N0");
                sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 13].Value = augTotal.ToString("N0");
                sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[rowCount, 14].Value = sepTotal.ToString("N0");
                sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                rowCount++;

                rowCount = rowCount + 2;

                sheet.Cells[rowCount, 1, rowCount, 14].Style.Font.Bold = true;
                sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                sheet.Cells[rowCount, 2].Value = "工数";
                rowCount++;

                if (assignmentsWithSectionAndCompany.Count > 0)
                {
                    List<ForecastAssignmentViewModel> customizedList = assignmentsWithSectionAndCompany.ToList();
                    List<ForecastAssignmentViewModel> tempList = null;
                    octTotal = 0; novTotal = 0; decTotal = 0; janTotal = 0; febTotal = 0; marTotal = 0; aprTotal = 0; mayTotal = 0; junTotal = 0; julTotal = 0; augTotal = 0; sepTotal = 0;

                    foreach (var item in assignmentsWithSectionAndCompany)
                    {

                        tempList = customizedList.Where(l => l.SectionId == item.SectionId && l.CompanyId == item.CompanyId).ToList();
                        if (tempList.Count > 0)
                        {
                            foreach (var removableItem in tempList)
                            {
                                customizedList.Remove(removableItem);
                            }



                            sheet.Cells[rowCount, 1].Value = tempList[0].SectionName;
                            sheet.Cells[rowCount, 2].Value = tempList[0].CompanyName;

                            foreach (var tempItem in tempList)
                            {

                                oct += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 10).Points);
                                nov += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 11).Points);
                                dec += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 12).Points);
                                jan += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 1).Points);
                                feb += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 2).Points);
                                mar += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 3).Points);
                                apr += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 4).Points);
                                may += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 5).Points);
                                jun += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 6).Points);
                                jul += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 7).Points);
                                aug += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 8).Points);
                                sep += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 9).Points);
                            }

                            sheet.Cells[rowCount, 3].Value = oct;
                            octTotal += oct;
                            sheet.Cells[rowCount, 4].Value = nov;
                            novTotal += nov;
                            sheet.Cells[rowCount, 5].Value = dec;
                            decTotal += dec;
                            sheet.Cells[rowCount, 6].Value = jan;
                            janTotal += jan;
                            sheet.Cells[rowCount, 7].Value = feb;
                            febTotal += feb;
                            sheet.Cells[rowCount, 8].Value = mar;
                            marTotal += mar;
                            sheet.Cells[rowCount, 9].Value = apr;
                            aprTotal += apr;
                            sheet.Cells[rowCount, 10].Value = may;
                            mayTotal += may;
                            sheet.Cells[rowCount, 11].Value = jun;
                            junTotal += jun;
                            sheet.Cells[rowCount, 12].Value = jul;
                            julTotal += jul;
                            sheet.Cells[rowCount, 13].Value = aug;
                            augTotal += aug;
                            sheet.Cells[rowCount, 14].Value = sep;
                            sepTotal += sep;

                            oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                            rowCount++;
                        }
                    }

                    sheet.Cells[rowCount, 2].Value = "合計";
                    sheet.Cells[rowCount, 3].Value = octTotal;
                    sheet.Cells[rowCount, 4].Value = novTotal;
                    sheet.Cells[rowCount, 5].Value = decTotal;
                    sheet.Cells[rowCount, 6].Value = janTotal;
                    sheet.Cells[rowCount, 7].Value = febTotal;
                    sheet.Cells[rowCount, 8].Value = marTotal;
                    sheet.Cells[rowCount, 9].Value = aprTotal;
                    sheet.Cells[rowCount, 10].Value = mayTotal;
                    sheet.Cells[rowCount, 11].Value = junTotal;
                    sheet.Cells[rowCount, 12].Value = julTotal;
                    sheet.Cells[rowCount, 13].Value = augTotal;
                    sheet.Cells[rowCount, 14].Value = sepTotal;
                    rowCount++;

                    rowCount = rowCount + 2;
                    sheet.Cells[rowCount, 1, rowCount, 14].Style.Font.Bold = true;
                    sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.BackgroundColor.SetColor(Color.LightGray);


                    sheet.Cells[rowCount, 2].Value = "金額";
                    rowCount++;

                    customizedList = assignmentsWithSectionAndCompany.ToList();
                    octTotal = 0; novTotal = 0; decTotal = 0; janTotal = 0; febTotal = 0; marTotal = 0; aprTotal = 0; mayTotal = 0; junTotal = 0; julTotal = 0; augTotal = 0; sepTotal = 0;
                    foreach (var item in assignmentsWithSectionAndCompany)
                    {

                        tempList = customizedList.Where(l => l.SectionId == item.SectionId && l.CompanyId == item.CompanyId).ToList();
                        if (tempList.Count > 0)
                        {
                            foreach (var removableItem in tempList)
                            {
                                customizedList.Remove(removableItem);
                            }



                            sheet.Cells[rowCount, 1].Value = tempList[0].SectionName;
                            sheet.Cells[rowCount, 2].Value = tempList[0].CompanyName;

                            foreach (var tempItem in tempList)
                            {

                                oct += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 10).Total);
                                nov += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 11).Total);
                                dec += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 12).Total);
                                jan += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 1).Total);
                                feb += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 2).Total);
                                mar += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 3).Total);
                                apr += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 4).Total);
                                may += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 5).Total);
                                jun += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 6).Total);
                                jul += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 7).Total);
                                aug += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 8).Total);
                                sep += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 9).Total);
                            }

                            sheet.Cells[rowCount, 3].Value = oct.ToString("N0");
                            octTotal += oct;
                            sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells[rowCount, 4].Value = nov.ToString("N0");
                            novTotal += nov;
                            sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells[rowCount, 5].Value = dec.ToString("N0");
                            decTotal += dec;
                            sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells[rowCount, 6].Value = jan.ToString("N0");
                            janTotal += jan;
                            sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells[rowCount, 7].Value = feb.ToString("N0");
                            febTotal += feb;
                            sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells[rowCount, 8].Value = mar.ToString("N0");
                            marTotal += mar;
                            sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells[rowCount, 9].Value = apr.ToString("N0");
                            aprTotal += apr;
                            sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells[rowCount, 10].Value = may.ToString("N0");
                            mayTotal += may;
                            sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells[rowCount, 11].Value = jun.ToString("N0");
                            junTotal += jun;
                            sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells[rowCount, 12].Value = jul.ToString("N0");
                            julTotal += jul;
                            sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells[rowCount, 13].Value = aug.ToString("N0");
                            augTotal += aug;
                            sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            sheet.Cells[rowCount, 14].Value = sep.ToString("N0");
                            sepTotal += sep;
                            sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                            rowCount++;
                        }
                    }
                    sheet.Cells[rowCount, 2].Value = "合計";
                    sheet.Cells[rowCount, 3].Value = octTotal.ToString("N0");
                    sheet.Cells[rowCount, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 4].Value = novTotal.ToString("N0");
                    sheet.Cells[rowCount, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 5].Value = decTotal.ToString("N0");
                    sheet.Cells[rowCount, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 6].Value = janTotal.ToString("N0");
                    sheet.Cells[rowCount, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 7].Value = febTotal.ToString("N0");
                    sheet.Cells[rowCount, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 8].Value = marTotal.ToString("N0");
                    sheet.Cells[rowCount, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 9].Value = aprTotal.ToString("N0");
                    sheet.Cells[rowCount, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 10].Value = mayTotal.ToString("N0");
                    sheet.Cells[rowCount, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 11].Value = junTotal.ToString("N0");
                    sheet.Cells[rowCount, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 12].Value = julTotal.ToString("N0");
                    sheet.Cells[rowCount, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 13].Value = augTotal.ToString("N0");
                    sheet.Cells[rowCount, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    sheet.Cells[rowCount, 14].Value = sepTotal.ToString("N0");
                    sheet.Cells[rowCount, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    rowCount++;


                }

                #region salary master

                var sheetSalaryMaster = package.Workbook.Worksheets.Add("SalaryMaster");
                List<Department> departments = _departmentBLL.GetAllDepartments().OrderBy(dep => dep.Id).ToList();
                List<SalaryMasterExportDto> salaryMasterExportDtos = new List<SalaryMasterExportDto>();
                var unitPriceTypeIds = _salaryBLL.GetUnitPriceTypeIdByYear(2022);
                foreach (var unitPriceTypeId in unitPriceTypeIds)
                {
                    foreach (var grade in grades)
                    {
                        salaryMasterExportDtos.Add(_salaryBLL.GetUnitPriceTypeWithGradeSalaryByYear(2022, grade.Id, unitPriceTypeId));
                    }
                }

                int rowCountSalaryMaster = 6;
                // year
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Value = "FY2022";

                rowCountSalaryMaster++;
                // heading 1
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Value = "期初目標";
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Style.Font.Color.SetColor(Color.White);
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);


                sheetSalaryMaster.Cells[rowCountSalaryMaster, 7].Value = "下修目標";
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 7].Style.Font.Color.SetColor(Color.White);
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 7].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheetSalaryMaster.Cells[rowCountSalaryMaster, 8].Value = "期初目標";
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 8].Style.Font.Color.SetColor(Color.White);
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 8].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);


                sheetSalaryMaster.Cells[rowCountSalaryMaster, 9].Value = "下修目標";
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 9].Style.Font.Color.SetColor(Color.White);
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 9].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);
                rowCountSalaryMaster++;

                // heading 2
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 4].Value = "";
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 4].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheetSalaryMaster.Cells[rowCountSalaryMaster, 5].Value = "";
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetSalaryMaster.Cells[rowCountSalaryMaster, 5].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                int extendedRow = 5;
                foreach (var item in departments)
                {
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Value = item.DepartmentName;
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Merge = true;
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.Font.Bold = true;
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.Font.Color.SetColor(Color.White);
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);
                    extendedRow += 2;
                }

                rowCountSalaryMaster++;
                int prevSalaryTypeId = 0;

                foreach (var item in salaryMasterExportDtos)
                {
                    if (item.GradeSalaryTypes.Count > 1)
                    {
                        if (prevSalaryTypeId != item.UnitPriceType.Id)
                        {
                            sheetSalaryMaster.Cells[rowCountSalaryMaster, 4].Value = item.UnitPriceType.TypeName;
                            prevSalaryTypeId = item.UnitPriceType.Id;
                        }

                        // for grade name
                        sheetSalaryMaster.Cells[rowCountSalaryMaster, 5].Value = item.GradeSalaryTypes[0].GradeName;
                        // for beginning target
                        sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Value = item.GradeSalaryTypes[0].GradeLowPoints.ToString("N0");
                        sheetSalaryMaster.Cells[rowCountSalaryMaster, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        // for beginning target
                        sheetSalaryMaster.Cells[rowCountSalaryMaster, 7].Value = item.GradeSalaryTypes[0].GradeHighPoints.ToString("N0");
                        sheetSalaryMaster.Cells[rowCountSalaryMaster, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        // for downward revision target
                        sheetSalaryMaster.Cells[rowCountSalaryMaster, 8].Value = item.GradeSalaryTypes[1].GradeLowPoints.ToString("N0");
                        sheetSalaryMaster.Cells[rowCountSalaryMaster, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        // for downward revision target
                        sheetSalaryMaster.Cells[rowCountSalaryMaster, 9].Value = item.GradeSalaryTypes[1].GradeHighPoints.ToString("N0");
                        sheetSalaryMaster.Cells[rowCountSalaryMaster, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        rowCountSalaryMaster++;
                    }



                }






                var sheetCommonMaster = package.Workbook.Worksheets.Add("CommonMaster");


                List<CommonMaster> commonMasters = _commonMasterBLL.GetCommonMasters();

                sheetCommonMaster.Cells[6, 4].Value = "";
                sheetCommonMaster.Cells[6, 4].Style.Font.Color.SetColor(Color.White);
                sheetCommonMaster.Cells[6, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetCommonMaster.Cells[6, 4].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheetCommonMaster.Cells[6, 5].Value = "昇給率";
                sheetCommonMaster.Cells[6, 5].Style.Font.Color.SetColor(Color.White);
                sheetCommonMaster.Cells[6, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetCommonMaster.Cells[6, 5].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheetCommonMaster.Cells[6, 6].Value = "残業固定時間";
                sheetCommonMaster.Cells[6, 6].Style.Font.Color.SetColor(Color.White);
                sheetCommonMaster.Cells[6, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetCommonMaster.Cells[6, 6].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);


                sheetCommonMaster.Cells[6, 7].Value = "賞与引当金比率";
                sheetCommonMaster.Cells[6, 7].Style.Font.Color.SetColor(Color.White);
                sheetCommonMaster.Cells[6, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetCommonMaster.Cells[6, 7].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheetCommonMaster.Cells[6, 8].Value = "賞与引当金定数";
                sheetCommonMaster.Cells[6, 8].Style.Font.Color.SetColor(Color.White);
                sheetCommonMaster.Cells[6, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetCommonMaster.Cells[6, 8].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                sheetCommonMaster.Cells[6, 9].Value = "給与法定福利費比率";
                sheetCommonMaster.Cells[6, 9].Style.Font.Color.SetColor(Color.White);
                sheetCommonMaster.Cells[6, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetCommonMaster.Cells[6, 9].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);

                int rowCountCommonMaster = 7;


                if (commonMasters.Count > 0)
                {
                    foreach (var commonMaster in commonMasters)
                    {
                        sheetCommonMaster.Cells[rowCountCommonMaster, 4].Value = commonMaster.GradeName;
                        sheetCommonMaster.Cells[rowCountCommonMaster, 5].Value = commonMaster.SalaryIncreaseRate;
                        sheetCommonMaster.Cells[rowCountCommonMaster, 6].Value = commonMaster.OverWorkFixedTime;
                        sheetCommonMaster.Cells[rowCountCommonMaster, 7].Value = commonMaster.BonusReserveRatio;
                        sheetCommonMaster.Cells[rowCountCommonMaster, 8].Value = commonMaster.BonusReserveConstant;
                        sheetCommonMaster.Cells[rowCountCommonMaster, 9].Value = commonMaster.WelfareCostRatio;

                        rowCountCommonMaster++;
                    }
                }


                #endregion

                var excelData = package.GetAsByteArray();
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                var fileName = explanation.ExplanationName + ".xlsx";
                return File(excelData, contentType, fileName);

            }
        }

        public double GetTotalManPoints(List<SalaryAssignmentDto> salaryAssignmentDtos, int monthId, int gradeId)
        {
            double points = 0;

            foreach (var item in salaryAssignmentDtos)
            {
                if (item.Grade.Id==gradeId)
                {
                    foreach (var singleAssignment in item.ForecastAssignmentViewModels)
                    {
                        points += Convert.ToDouble(singleAssignment.forecasts.SingleOrDefault(f => f.Month == monthId).Points);
                    }
                }
            }

            return points;
        }
        public double GetTotalOverTime(List<SalaryAssignmentDto> salaryAssignmentDtos, int monthId, int gradeId)
        {
            double points = 0;

            foreach (var item in salaryAssignmentDtos)
            {
                if (item.Grade.Id == gradeId)
                {
                    foreach (var singleAssignment in item.ForecastAssignmentViewModels)
                    {
                        points += Convert.ToDouble(singleAssignment.forecasts.SingleOrDefault(f => f.Month == monthId).OverTime);
                    }
                }
            }

            return points;
        }
        public ActionResult DataExports()
        {
            return View(new ExportViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }

        public JsonResult ViewDataByDepartment(int departmentId)
        {
            List<ForecastAssignmentViewModel> forecastAssignmentViewModels = new List<ForecastAssignmentViewModel>();
            var _department = _departmentBLL.GetDepartmentByDepartemntId(departmentId);

            using (var client = new HttpClient())
            {
                string uri = "" + _utility.Address + "/api/utilities/SearchForecastEmployee?employeeName=&sectionId=&departmentId=" + departmentId + "&inchargeId=&roleId=&explanationId=&companyId=&status=&year=";
                client.BaseAddress = new Uri(uri);
                //HTTP GET
                var responseTask = client.GetAsync("");
                responseTask.Wait();
                 
                var result = responseTask.Result;
                if (result.IsSuccessStatusCode)
                {
                    var readTask = result.Content.ReadAsAsync<List<ForecastAssignmentViewModel>>();
                    readTask.Wait();

                    forecastAssignmentViewModels = readTask.Result;
                }
                else //web api sent error response 
                {
                    //log response status here..

                    forecastAssignmentViewModels = new List<ForecastAssignmentViewModel>();

                    // ModelState.AddModelError(string.Empty, "Server error. Please contact administrator.");
                }
            }

            foreach (var forecastItem in forecastAssignmentViewModels)
            {
                if (forecastItem.CompanyName.ToLower() != "mw" || forecastItem.SectionName.ToLower() != "mw")
                {
                    forecastItem.GradePoint = "";
                }
            }


            return Json(forecastAssignmentViewModels, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetAllExplanationsByDepartmentId(int departmentId)
        {
            return Json(_explanationsBLL.GetAllExplanationsByDepartmentId(departmentId), JsonRequestBehavior.AllowGet);
        }

        public JsonResult ViewDataByAllocation(int departmentId = 0, int explanationId = 0)
        {
            List<Grade> grades = _gradeBll.GetAllGrade();

            List<SalaryAssignmentDto> salaryAssignmentDtos = new List<SalaryAssignmentDto>();
            var department = _departmentBLL.GetDepartmentByDepartemntId(departmentId);
            var explanation = _explanationsBLL.GetExplanationByExplanationId(explanationId);
            var salaries = _salaryBLL.GetAllSalaryPoints();


            // get all data with department and allocation
            List<ForecastAssignmentViewModel> assignmentsWithForecast = _exportBLL.AssignmentsByAllocation(departmentId, explanationId);
            //filtered by grade id
            List<ForecastAssignmentViewModel> assignmentsWithGrade = assignmentsWithForecast.Where(a => a.GradeId != null).ToList();
            //filtered by section and company
            List<ForecastAssignmentViewModel> assignmentsWithSectionAndCompany = assignmentsWithForecast.Where(a => a.SectionId != null && a.CompanyId != null).ToList();

            foreach (var item in grades)
            {
                List<ForecastAssignmentViewModel> forecastAssignmentViews = new List<ForecastAssignmentViewModel>();

                SalaryAssignmentDto salaryAssignmentDto = new SalaryAssignmentDto();
                salaryAssignmentDto.Grade = item;

                var gradeSalaryTypes = _exportBLL.GetGradeUnitPriceTypes(item.Id, departmentId, 2022, 1);

                foreach (var gradeSalary in gradeSalaryTypes)
                {
                    List<ForecastAssignmentViewModel> filteredAssignmentsByGradeSalaryTypeId = assignmentsWithGrade.Where(a => a.GradeId == gradeSalary.GradeId.ToString()).ToList();
                    forecastAssignmentViews.AddRange(filteredAssignmentsByGradeSalaryTypeId);
                }

                salaryAssignmentDto.ForecastAssignmentViewModels = forecastAssignmentViews;
                salaryAssignmentDtos.Add(salaryAssignmentDto);

            }
            int rowCount = 5;
            string allocaionTableBodyTd = "", allocaionTableBodyTrStart = "", allocaionTableBodyTrEnd = "";
            allocaionTableBodyTrStart = "<tr>";
            allocaionTableBodyTrEnd = "</tr>";
            // allocaionTableBodyTd = allocaionTableBodyTd +"td>人数</td>";

            decimal oct = 0;
            decimal nov = 0;
            decimal dec = 0;
            decimal jan = 0;
            decimal feb = 0;
            decimal mar = 0;
            decimal apr = 0;
            decimal may = 0;
            decimal jun = 0;
            decimal jul = 0;
            decimal aug = 0;
            decimal sep = 0;

            decimal octTotal = 0;
            decimal novTotal = 0;
            decimal decTotal = 0;
            decimal janTotal = 0;
            decimal febTotal = 0;
            decimal marTotal = 0;
            decimal aprTotal = 0;
            decimal mayTotal = 0;
            decimal junTotal = 0;
            decimal julTotal = 0;
            decimal augTotal = 0;
            decimal sepTotal = 0;

            //grade wise count
            allocaionTableBodyTd = allocaionTableBodyTd + "<tr>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td rowspan='17'>人数</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "</tr>";
            foreach (var item in salaryAssignmentDtos)
            {
                allocaionTableBodyTd = allocaionTableBodyTd + "<tr>";
                //allocaionTableBodyTd = allocaionTableBodyTd +"<td></td>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + item.Grade.GradeName + "</td>";
                foreach (var singleAssignment in item.ForecastAssignmentViewModels)
                {
                    oct += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 10).Points);
                    nov += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 11).Points);
                    dec += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 12).Points);
                    jan += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 1).Points);
                    feb += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 2).Points);
                    mar += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 3).Points);
                    apr += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 4).Points);
                    may += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 5).Points);
                    jun += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 6).Points);
                    jul += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 7).Points);
                    aug += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 8).Points);
                    sep += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 9).Points);
                }
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + oct + "</td>";
                octTotal += oct;
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + nov + "</td>";
                novTotal += nov;
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + dec + "</td>";
                decTotal += dec;
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + jan + "</td>";
                janTotal += jan;
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + feb + "</td>";
                febTotal += feb;
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + mar + "</td>";
                marTotal += mar;
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + apr + "</td>";
                aprTotal += apr;
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + may + "</td>";
                mayTotal += may;
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + jun + "</td>";
                junTotal += jun;
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + jul + "</td>";
                julTotal += jul;
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + aug + "</td>";
                augTotal += aug;
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + sep + "</td>";
                sepTotal += sep;
                oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                rowCount++;
                allocaionTableBodyTd = allocaionTableBodyTd + "</tr>";
            }
            allocaionTableBodyTd = allocaionTableBodyTd + "<tr>";
            //allocaionTableBodyTd = allocaionTableBodyTd +"<td></td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td>合計</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + octTotal + "</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + novTotal + "</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + decTotal + "</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + janTotal + "</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + febTotal + "</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + marTotal + "</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + aprTotal + "</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + mayTotal + "</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + junTotal + "</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + julTotal + "</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + augTotal + "</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + sepTotal + "</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "</tr>";

            // common master wise count
            allocaionTableBodyTd = allocaionTableBodyTd + "<tr>";
            allocaionTableBodyTd = allocaionTableBodyTd + "<td rowspan='16'>1人あたりの時間外勤務見込<br>(みなし時間（固定時間）<br>を含む残業時間<br>を入力してください</td>";
            allocaionTableBodyTd = allocaionTableBodyTd + "</tr>";

            decimal octOverTime = 0;
            decimal novOverTime = 0;
            decimal decOverTime = 0;
            decimal janOverTime = 0;
            decimal febOverTime = 0;
            decimal marOverTime = 0;
            decimal aprOverTime = 0;
            decimal mayOverTime = 0;
            decimal junOverTime = 0;
            decimal julOverTime = 0;
            decimal augOverTime = 0;
            decimal sepOverTime = 0;
            //over time
            foreach (var item in salaryAssignmentDtos)
            {
               

                foreach (var singleAssignment in item.ForecastAssignmentViewModels)
                {
                    octOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 10).OverTime);
                    novOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 11).OverTime);
                    decOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 12).OverTime);
                    janOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 1).OverTime);
                    febOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 2).OverTime);
                    marOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 3).OverTime);
                    aprOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 4).OverTime);
                    mayOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 5).OverTime);
                    junOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 6).OverTime);
                    julOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 7).OverTime);
                    augOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 8).OverTime);
                    sepOverTime += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 9).OverTime);
                }
                allocaionTableBodyTd = allocaionTableBodyTd + "<tr>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + item.Grade.GradeName +"</td>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + octOverTime + "</td>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + novOverTime + "</td>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + decOverTime + "</td>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + janOverTime + "</td>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + febOverTime + "</td>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + marOverTime + "</td>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + aprOverTime + "</td>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + mayOverTime + "</td>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + junOverTime + "</td>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + julOverTime + "</td>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + augOverTime + "</td>";
                allocaionTableBodyTd = allocaionTableBodyTd + "<td>" + sepOverTime + "</td>";

                rowCount++;
                allocaionTableBodyTd = allocaionTableBodyTd + "</tr>";



                //valuesWithGrades.Add(new ValuesWithGrade { GradeId = item.Grade.Id,GradeName=item.Grade.GradeName,Point= octTotal,MonthId=10 });
                octOverTime = 0; novOverTime = 0; decOverTime = 0; janOverTime = 0; febOverTime = 0; marOverTime = 0; aprOverTime = 0; mayOverTime = 0; junOverTime = 0; julOverTime = 0; augOverTime = 0; sepOverTime = 0;
                rowCount++;
            }

            rowCount++;
            int rowCountForFuture = rowCount;
            rowCount += 12;
            // 2. regular total
            double regularSalaryTotalOct = 0, regularSalaryTotalNov = 0, regularSalaryTotalDec = 0, regularSalaryTotalJan = 0, regularSalaryTotalFeb = 0, regularSalaryTotalMar = 0, regularSalaryTotalApr = 0, regularSalaryTotalMay = 0, regularSalaryTotalJun = 0, regularSalaryTotalJul = 0, regularSalaryTotalAug = 0, regularSalaryTotalSep = 0;
            // 3. fixed total
            double fixedSalaryTotalOct = 0, fixedSalaryTotalNov = 0, fixedSalaryTotalDec = 0, fixedSalaryTotalJan = 0, fixedSalaryTotalFeb = 0, fixedSalaryTotalMar = 0, fixedSalaryTotalApr = 0, fixedSalaryTotalMay = 0, fixedSalaryTotalJun = 0, fixedSalaryTotalJul = 0, fixedSalaryTotalAug = 0, fixedSalaryTotalSep = 0;
            // 4. over time
            double overTimeSalaryTotalOct = 0, overTimeSalaryTotalNov = 0, overTimeSalaryTotalDec = 0, overTimeSalaryTotalJan = 0, overTimeSalaryTotalFeb = 0, overTimeSalaryTotalMar = 0, overTimeSalaryTotalApr = 0, overTimeSalaryTotalMay = 0, overTimeSalaryTotalJun = 0, overTimeSalaryTotalJul = 0, overTimeSalaryTotalAug = 0, overTimeSalaryTotalSep = 0;
            // 5. total Salary
            double salaryTotalOct = 0, salaryTotalNov = 0, salaryTotalDec = 0, salaryTotalJan = 0, salaryTotalFeb = 0, salaryTotalMar = 0, salaryTotalApr = 0, salaryTotalMay = 0, salaryTotalJun = 0, salaryTotalJul = 0, salaryTotalAug = 0, salaryTotalSep = 0;
            // 6. Miscellaneous Wages
            double mwTotalOct = 0, mwTotalNov = 0, mwTotalDec = 0, mwTotalJan = 0, mwTotalFeb = 0, mwTotalMar = 0, mwTotalApr = 0, mwTotalMay = 0, mwTotalJun = 0, mwTotalJul = 0, mwTotalAug = 0, mwTotalSep = 0;
            // 7. Dispatch Fee
            double dispatchFeeTotalOct = 0, dispatchFeeTotalNov = 0, dispatchFeeTotalDec = 0, dispatchFeeTotalJan = 0, dispatchFeeTotalFeb = 0, dispatchFeeTotalMar = 0, dispatchFeeTotalApr = 0, dispatchFeeTotalMay = 0, dispatchFeeTotalJun = 0, dispatchFeeTotalJul = 0, dispatchFeeTotalAug = 0, dispatchFeeTotalSep = 0;
            // 8. Provision for Employee Bonus
            double employeeBonusTotalOct = 0, employeeBonusTotalNov = 0, employeeBonusTotalDec = 0, employeeBonusTotalJan = 0, employeeBonusTotalFeb = 0, employeeBonusTotalMar = 0, employeeBonusTotalApr = 0, employeeBonusTotalMay = 0, employeeBonusTotalJun = 0, employeeBonusTotalJul = 0, employeeBonusTotalAug = 0, employeeBonusTotalSep = 0;
            // 9. Commuting Expenses
            double commutingExpensesTotalOct = 0, commutingExpensesTotalNov = 0, commutingExpensesTotalDec = 0, commutingExpensesTotalJan = 0, commutingExpensesTotalFeb = 0, commutingExpensesTotalMar = 0, commutingExpensesTotalApr = 0, commutingExpensesTotalMay = 0, commutingExpensesTotalJun = 0, commutingExpensesTotalJul = 0, commutingExpensesTotalAug = 0, commutingExpensesTotalSep = 0;
            // 10. Salary Statutory Welfare Expenses
            double welfareExpensesTotalOct = 0, welfareExpensesTotalNov = 0, welfareExpensesTotalDec = 0, welfareExpensesTotalJan = 0, welfareExpensesTotalFeb = 0, welfareExpensesTotalMar = 0, welfareExpensesTotalApr = 0, welfareExpensesTotalMay = 0, welfareExpensesTotalJun = 0, welfareExpensesTotalJul = 0, welfareExpensesTotalAug = 0, welfareExpensesTotalSep = 0;
            // 11. Provision for Statotory Welfare Expenses for Bonus
            double welfareExpensesBonusTotalOct = 0, welfareExpensesBonusTotalNov = 0, welfareExpensesBonusTotalDec = 0, welfareExpensesBonusTotalJan = 0, welfareExpensesBonusTotalFeb = 0, welfareExpensesBonusTotalMar = 0, welfareExpensesBonusTotalApr = 0, welfareExpensesBonusTotalMay = 0, welfareExpensesBonusTotalJun = 0, welfareExpensesBonusTotalJul = 0, welfareExpensesBonusTotalAug = 0, welfareExpensesBonusTotalSep = 0;
            // 12. Total Statutory Benefites
            double statutoryTotalOct = 0, statutoryTotalNov = 0, statutoryTotalDec = 0, statutoryTotalJan = 0, statutoryTotalFeb = 0, statutoryTotalMar = 0, statutoryTotalApr = 0, statutoryTotalMay = 0, statutoryTotalJun = 0, statutoryTotalJul = 0, statutoryTotalAug = 0, statutoryTotalSep = 0;
            // 13. Total Expenses
            double expensesTotalOct = 0, expensesTotalNov = 0, expensesTotalDec = 0, expensesTotalJan = 0, expensesTotalFeb = 0, expensesTotalMar = 0, expensesTotalApr = 0, expensesTotalMay = 0, expensesTotalJun = 0, expensesTotalJul = 0, expensesTotalAug = 0, expensesTotalSep = 0;

            
            //other grade with entities
            rowCount++;
            int count = 1;
            string strOtherGradeTd = "";
            foreach (var item in salaryAssignmentDtos)
            {
                strOtherGradeTd = strOtherGradeTd + "<tr>";
                strOtherGradeTd = strOtherGradeTd + "<td rowspan='14'>" + item.Grade.GradeName + "</td>";
                strOtherGradeTd = strOtherGradeTd + "</tr>";

                //salary allow regular
                double totalRegularOct = 0, totalRegularNov = 0, totalRegularDec = 0, totalRegularJan = 0, totalRegularFeb = 0, totalRegularMar = 0, totalRegularApr = 0, totalRegularMay = 0, totalRegularJun = 0, totalRegularJul = 0, totalRegularAug = 0, totalRegularSep = 0;
                // salary allow fixed
                double totalFixedOct = 0, totalFixedNov = 0, totalFixedDec = 0, totalFixedJan = 0, totalFixedFeb = 0, totalFixedMar = 0, totalFixedApr = 0, totalFixedMay = 0, totalFixedJun = 0, totalFixedJul = 0, totalFixedAug = 0, totalFixedSep = 0;
                // salary allow oertime
                double totalOverOct = 0, totalOverNov = 0, totalOverDec = 0, totalOverJan = 0, totalOverFeb = 0, totalOverMar = 0, totalOverApr = 0, totalOverMay = 0, totalOverJun = 0, totalOverJul = 0, totalOverAug = 0, totalOverSep = 0;
                // total salary
                double totalSalaryOct = 0, totalSalaryNov = 0, totalSalaryDec = 0, totalSalaryJan = 0, totalSalaryFeb = 0, totalSalaryMar = 0, totalSalaryApr = 0, totalSalaryMay = 0, totalSalaryJun = 0, totalSalaryJul = 0, totalSalaryAug = 0, totalSalarySep = 0;
                // Miscellaneous Wages
                double mWagesOct = 0, mWagesNov = 0, mWagesDec = 0, mWagesJan = 0, mWagesFeb = 0, mWagesMar = 0, mWagesApr = 0, mWagesMay = 0, mWagesJun = 0, mWagesJul = 0, mWagesAug = 0, mWagesSep = 0;
                //dispatch fee
                double dispatchFeeOct = 0, dispatchFeeNov = 0, dispatchFeeDec = 0, dispatchFeeJan = 0, dispatchFeeFeb = 0, dispatchFeeMar = 0, dispatchFeeApr = 0, dispatchFeeMay = 0, dispatchFeeJun = 0, dispatchFeeJul = 0, dispatchFeeAug = 0, dispatchFeeSep = 0;
                // employee bonus
                double employeeBonusOct = 0, employeeBonusNov = 0, employeeBonusDec = 0, employeeBonusJan = 0, employeeBonusFeb = 0, employeeBonusMar = 0, employeeBonusApr = 0, employeeBonusMay = 0, employeeBonusJun = 0, employeeBonusJul = 0, employeeBonusAug = 0, employeeBonusSep = 0;
                //commuting expenses
                double commutingExpensesOct = 0, commutingExpensesNov = 0, commutingExpensesDec = 0, commutingExpensesJan = 0, commutingExpensesFeb = 0, commutingExpensesMar = 0, commutingExpensesApr = 0, commutingExpensesMay = 0, commutingExpensesJun = 0, commutingExpensesJul = 0, commutingExpensesAug = 0, commutingExpensesSep = 0;
                // welfare expenses
                double wExpensesOct = 0, wExpensesNov = 0, wExpensesDec = 0, wExpensesJan = 0, wExpensesFeb = 0, wExpensesMar = 0, wExpensesApr = 0, wExpensesMay = 0, wExpensesJun = 0, wExpensesJul = 0, wExpensesAug = 0, wExpensesSep = 0;
                // welfare expenses bonuses
                double wExpBonusOct = 0, wExpBonusNov = 0, wExpBonusDec = 0, wExpBonusJan = 0, wExpBonusFeb = 0, wExpBonusMar = 0, wExpBonusApr = 0, wExpBonusMay = 0, wExpBonusJun = 0, wExpBonusJul = 0, wExpBonusAug = 0, wExpBonusSep = 0;
                // total statutory
                double totalStatutoryOct = 0, totalStatutoryNov = 0, totalStatutoryDec = 0, totalStatutoryJan = 0, totalStatutoryFeb = 0, totalStatutoryMar = 0, totalStatutoryApr = 0, totalStatutoryMay = 0, totalStatutoryJun = 0, totalStatutoryJul = 0, totalStatutoryAug = 0, totalStatutorySep = 0;

                int innerCount = 1;
                int salaryTypeCount = 1;
                foreach (var salaryType in _unitPriceTypeBLL.GetAllSalaryTypes())
                {
                    strOtherGradeTd = strOtherGradeTd + "<tr>";
                    strOtherGradeTd = strOtherGradeTd + "<td>" + salaryType.SalaryTypeName + "</td>";
                    //strOtherGradeTd = strOtherGradeTd +"</tr>";                           

                    // alligned with the serial of salary type
                    // executives
                    if (salaryTypeCount == 1)
                    {
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                    }

                    // salary allowance (regular)
                    if (salaryTypeCount == 2)
                    {
                        double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                        double beginningValue = _salaryBLL.GetGradeUnitPriceType(departmentId, 1, 2022, item.Grade.Id).GradeLowPoints;
                        double commonValue = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).SalaryIncreaseRate);

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointOct * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalRegularOct += manpointOct * beginningValue * commonValue;
                        regularSalaryTotalOct += totalRegularOct;
                        double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointOct * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalRegularNov += manpointNov * beginningValue * commonValue;
                        regularSalaryTotalNov += totalRegularNov;
                        double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointDec * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalRegularDec += manpointDec * beginningValue * commonValue;
                        regularSalaryTotalDec += totalRegularDec;
                        double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJan * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalRegularJan += manpointJan * beginningValue * commonValue;
                        regularSalaryTotalJan += totalRegularJan;
                        double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointFeb * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalRegularFeb += manpointFeb * beginningValue * commonValue;
                        regularSalaryTotalFeb += totalRegularFeb;

                        double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointMar * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalRegularMar += manpointMar * beginningValue * commonValue;
                        regularSalaryTotalMar += totalRegularMar;

                        double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointApr * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalRegularApr += manpointApr * beginningValue * commonValue;
                        regularSalaryTotalApr += totalRegularApr;

                        double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointMay * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalRegularMay += manpointMay * beginningValue * commonValue;
                        regularSalaryTotalMay += totalRegularMay;

                        double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJun * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalRegularJun += manpointJun * beginningValue * commonValue;
                        regularSalaryTotalJun += totalRegularJun;

                        double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJul * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalRegularJul += manpointJul * beginningValue * commonValue;
                        regularSalaryTotalJul += totalRegularJul;

                        double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointAug * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalRegularAug += manpointAug * beginningValue * commonValue;
                        regularSalaryTotalAug += totalRegularAug;

                        double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointSep * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalRegularSep += manpointSep * beginningValue * commonValue;
                        regularSalaryTotalSep += totalRegularSep;
                    }

                    // salary allowance (fixed)
                    if (salaryTypeCount == 3)
                    {
                        double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                        double beginningValue = _salaryBLL.GetGradeUnitPriceType(departmentId, 2, 2022, item.Grade.Id).GradeLowPoints;
                        double commonValue = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).SalaryIncreaseRate);

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointOct * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalFixedOct += manpointOct * beginningValue * commonValue;
                        fixedSalaryTotalOct += totalFixedOct;

                        double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointNov * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalFixedNov += manpointNov * beginningValue * commonValue;
                        fixedSalaryTotalNov += totalFixedNov;

                        double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointDec * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalFixedDec += manpointDec * beginningValue * commonValue;
                        fixedSalaryTotalDec += totalFixedDec;

                        double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJan * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalFixedJan += manpointJan * beginningValue * commonValue;
                        fixedSalaryTotalJan += totalFixedJan;


                        double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointFeb * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalFixedFeb += manpointFeb * beginningValue * commonValue;
                        fixedSalaryTotalFeb += totalFixedFeb;

                        double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointMar * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalFixedMar += manpointMar * beginningValue * commonValue;
                        fixedSalaryTotalMar += totalFixedMar;


                        double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointApr * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalFixedApr += manpointApr * beginningValue * commonValue;
                        fixedSalaryTotalApr += totalFixedApr;

                        double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointMay * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalFixedMay += manpointMay * beginningValue * commonValue;
                        fixedSalaryTotalMay += totalFixedMay;

                        double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJun * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalFixedJun += manpointJun * beginningValue * commonValue;
                        fixedSalaryTotalJun += totalFixedJun;

                        double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJul * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalFixedJul += manpointJul * beginningValue * commonValue;
                        fixedSalaryTotalJul += totalFixedJul;

                        double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointAug * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalFixedAug += manpointAug * beginningValue * commonValue;
                        fixedSalaryTotalAug += totalFixedAug;

                        double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointSep * beginningValue * commonValue).ToString("N0") + "</td>";

                        totalFixedSep += manpointSep * beginningValue * commonValue;
                        fixedSalaryTotalSep += totalFixedSep;

                    }

                    // salary allowance (overtime)
                    if (salaryTypeCount == 4)
                    {
                        if (item.Grade.Id==12)
                        {

                        }
                        CommonMaster commonMaster = _commonMasterBLL.GetCommonMasters().Where(cm => cm.GradeId == item.Grade.Id).SingleOrDefault();
                        double salaryMasterOverTimeLowPoints = _exportBLL.GetGradeUnitPriceTypes(item.Grade.Id, department.Id, 2022, 3).SingleOrDefault().GradeLowPoints;

                        double calculatedOverTimeOct = GetTotalOverTime(salaryAssignmentDtos, 10, item.Grade.Id);
                        double manPointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                        
                        if (calculatedOverTimeOct > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverOct = manPointOct * (calculatedOverTimeOct - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalOct += totalOverOct;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + totalOverOct.ToString("N0") + "</td>";


                        double calculatedOverTimeNov = GetTotalOverTime(salaryAssignmentDtos, 11, item.Grade.Id);
                        double manPointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);

                        if (calculatedOverTimeNov > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverNov = manPointNov * (calculatedOverTimeNov - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalNov += totalOverNov;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + totalOverNov.ToString("N0") + "</td>";

                        double calculatedOverTimeDec = GetTotalOverTime(salaryAssignmentDtos, 12, item.Grade.Id);
                        double manPointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);

                        if (calculatedOverTimeDec > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverDec = manPointDec * (calculatedOverTimeDec - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalDec += totalOverDec;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + totalOverDec.ToString("N0") + "</td>";

                        double calculatedOverTimeJan = GetTotalOverTime(salaryAssignmentDtos, 1, item.Grade.Id);
                        double manPointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                        if (calculatedOverTimeJan > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverJan = manPointJan * (calculatedOverTimeJan - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalJan += totalOverJan;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + totalOverJan.ToString("N0") + "</td>";

                        double calculatedOverTimeFeb = GetTotalOverTime(salaryAssignmentDtos, 2, item.Grade.Id);
                        double manPointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);

                        if (calculatedOverTimeFeb > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverFeb = manPointFeb * (calculatedOverTimeFeb - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalFeb += totalOverFeb;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + totalOverFeb.ToString("N0") + "</td>";

                        double calculatedOverTimeMar = GetTotalOverTime(salaryAssignmentDtos, 3, item.Grade.Id);
                        double manPointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);

                        if (calculatedOverTimeMar > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverMar = manPointMar * (calculatedOverTimeMar - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalMar += totalOverMar;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + totalOverMar.ToString("N0") + "</td>";

                        double calculatedOverTimeApr = GetTotalOverTime(salaryAssignmentDtos, 4, item.Grade.Id);
                        double manPointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);

                        if (calculatedOverTimeApr > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverApr = manPointApr * (calculatedOverTimeApr - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalApr += totalOverApr;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + totalOverApr.ToString("N0") + "</td>";

                        double calculatedOverTimeMay = GetTotalOverTime(salaryAssignmentDtos, 5, item.Grade.Id);
                        double manPointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);

                        if (calculatedOverTimeMay > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverMay = manPointMay * (calculatedOverTimeMay - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalMay += totalOverMay;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + totalOverMay.ToString("N0") + "</td>";

                        double calculatedOverTimeJun = GetTotalOverTime(salaryAssignmentDtos, 6, item.Grade.Id);
                        double manPointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);

                        if (calculatedOverTimeJun > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverJun = manPointJun * (calculatedOverTimeJun - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalJun += totalOverJun;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + totalOverJun.ToString("N0") + "</td>";

                        double calculatedOverTimeJul = GetTotalOverTime(salaryAssignmentDtos, 7, item.Grade.Id);
                        double manPointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                        if (calculatedOverTimeJul > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverJul = manPointJul * (calculatedOverTimeJul - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalJul += totalOverJul;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + totalOverJul.ToString("N0") + "</td>";

                        double calculatedOverTimeAug = GetTotalOverTime(salaryAssignmentDtos, 8, item.Grade.Id);
                        double manPointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);

                        if (calculatedOverTimeAug > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverAug = manPointAug * (calculatedOverTimeAug - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalAug += totalOverAug;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + totalOverAug.ToString("N0") + "</td>";

                        double calculatedOverTimeSept = GetTotalOverTime(salaryAssignmentDtos, 9, item.Grade.Id);
                        double manPointSept = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);

                        if (calculatedOverTimeSept > Convert.ToDouble(commonMaster.OverWorkFixedTime))
                        {
                            totalOverSep = manPointSept * (calculatedOverTimeSept - Convert.ToDouble(commonMaster.OverWorkFixedTime)) * salaryMasterOverTimeLowPoints * Convert.ToDouble(commonMaster.SalaryIncreaseRate);
                        }
                        overTimeSalaryTotalSep += totalOverSep;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + totalOverSep.ToString("N0") + "</td>";

                    }                    

                    // salary allowance (total)
                    if (salaryTypeCount == 5)
                    {
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalRegularOct + totalFixedOct + totalOverOct).ToString("N0") + "</td>";

                        totalSalaryOct += totalRegularOct + totalFixedOct + totalOverOct;
                        salaryTotalOct += totalSalaryOct;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalRegularNov + totalFixedNov + totalOverNov).ToString("N0") + "</td>";

                        totalSalaryNov += totalRegularNov + totalFixedNov + totalOverNov;
                        salaryTotalNov += totalSalaryNov;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalRegularDec + totalFixedDec + totalOverDec).ToString("N0") + "</td>";

                        totalSalaryDec += totalRegularDec + totalFixedDec + totalOverDec;
                        salaryTotalDec += totalSalaryDec;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalRegularJan + totalFixedJan + totalOverJan).ToString("N0") + "</td>";

                        totalSalaryJan += totalRegularJan + totalFixedJan + totalOverJan;
                        salaryTotalJan += totalSalaryJan;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalRegularFeb + totalFixedFeb + totalOverFeb).ToString("N0") + "</td>";

                        totalSalaryFeb += totalRegularFeb + totalFixedFeb + totalOverFeb;
                        salaryTotalFeb += totalSalaryFeb;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalRegularMar + totalFixedMar + totalOverMar).ToString("N0") + "</td>";

                        totalSalaryMar += totalRegularMar + totalFixedMar + totalOverMar;
                        salaryTotalMar += totalSalaryMar;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalRegularApr + totalFixedApr + totalOverApr).ToString("N0") + "</td>";

                        totalSalaryApr += totalRegularApr + totalFixedApr + totalOverApr;
                        salaryTotalApr += totalSalaryApr;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalRegularMay + totalFixedMay + totalOverMay).ToString("N0") + "</td>";

                        totalSalaryMay += totalRegularMay + totalFixedMay + totalOverMay;
                        salaryTotalMay += totalSalaryMay;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalRegularJun + totalFixedJun + totalOverJun).ToString("N0") + "</td>";

                        totalSalaryJun += totalRegularJun + totalFixedJun + totalOverJun;
                        salaryTotalJun += totalSalaryJun;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalRegularJul + totalFixedJul + totalOverJul).ToString("N0") + "</td>";

                        totalSalaryJul += totalRegularJul + totalFixedJul + totalOverJul;
                        salaryTotalJul += totalSalaryJul;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalRegularAug + totalFixedAug + totalOverAug).ToString("N0") + "</td>";

                        totalSalaryAug += totalRegularAug + totalFixedAug + totalOverAug;
                        salaryTotalAug += totalSalaryAug;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalRegularSep + totalFixedSep + totalOverSep).ToString("N0") + "</td>";

                        totalSalarySep += totalRegularSep + totalFixedSep + totalOverSep;
                        salaryTotalSep += totalSalarySep;
                    }

                    // Miscellneous Wages
                    if (salaryTypeCount == 6)
                    {
                        double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                        double beginningValue = _salaryBLL.GetGradeUnitPriceType(departmentId, salaryType.Id, 2022, item.Grade.Id).GradeLowPoints;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointOct * beginningValue).ToString("N0") + "</td>";

                        mWagesOct += manpointOct * beginningValue;
                        mwTotalOct += mWagesOct;

                        double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointNov * beginningValue).ToString("N0") + "</td>";

                        mWagesNov += manpointNov * beginningValue;
                        mwTotalNov += mWagesNov;

                        double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointDec * beginningValue).ToString("N0") + "</td>";

                        mWagesDec += manpointDec * beginningValue;
                        mwTotalDec += mWagesDec;

                        double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJan * beginningValue).ToString("N0") + "</td>";

                        mWagesJan += manpointJan * beginningValue;
                        mwTotalJan += mWagesJan;

                        double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointFeb * beginningValue).ToString("N0") + "</td>";

                        mWagesFeb += manpointFeb * beginningValue;
                        mwTotalFeb += mWagesFeb;

                        double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointMar * beginningValue).ToString("N0") + "</td>";

                        mWagesMar += manpointMar * beginningValue;
                        mwTotalMar += mWagesMar;

                        double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointApr * beginningValue).ToString("N0") + "</td>";

                        mWagesApr += manpointApr * beginningValue;
                        mwTotalApr += mWagesApr;

                        double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointMay * beginningValue).ToString("N0") + "</td>";

                        mWagesMay += manpointMay * beginningValue;
                        mwTotalMay += mWagesMay;

                        double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJun * beginningValue).ToString("N0") + "</td>";

                        mWagesJun += manpointJun * beginningValue;
                        mwTotalJun += mWagesJun;

                        double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJul * beginningValue).ToString("N0") + "</td>";

                        mWagesJul += manpointJul * beginningValue;
                        mwTotalJul += mWagesJul;

                        double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointAug * beginningValue).ToString("N0") + "</td>";

                        mWagesAug += manpointAug * beginningValue;
                        mwTotalAug += mWagesAug;

                        double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointSep * beginningValue).ToString("N0") + "</td>";

                        mWagesSep += manpointSep * beginningValue;
                        mwTotalSep += mWagesSep;
                    }

                    // Dispatch Fee
                    if (salaryTypeCount == 7)
                    {
                        double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                        double beginningValue = _salaryBLL.GetGradeUnitPriceType(departmentId, salaryType.Id, 2022, item.Grade.Id).GradeLowPoints;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointOct * beginningValue).ToString("N0") + "</td>";

                        dispatchFeeOct += manpointOct * beginningValue;
                        dispatchFeeTotalOct += dispatchFeeOct;

                        double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointNov * beginningValue).ToString("N0") + "</td>";

                        dispatchFeeNov += manpointNov * beginningValue;
                        dispatchFeeTotalNov += dispatchFeeNov;

                        double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointDec * beginningValue).ToString("N0") + "</td>";

                        dispatchFeeDec += manpointDec * beginningValue;
                        dispatchFeeTotalDec += dispatchFeeDec;

                        double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJan * beginningValue).ToString("N0") + "</td>";

                        dispatchFeeJan += manpointJan * beginningValue;
                        dispatchFeeTotalJan += dispatchFeeJan;

                        double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointFeb * beginningValue).ToString("N0") + "</td>";

                        dispatchFeeFeb += manpointFeb * beginningValue;
                        dispatchFeeTotalFeb += dispatchFeeFeb;

                        double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointMar * beginningValue).ToString("N0") + "</td>";

                        dispatchFeeMar += manpointMar * beginningValue;
                        dispatchFeeTotalMar += dispatchFeeMar;

                        double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointApr * beginningValue).ToString("N0") + "</td>";

                        dispatchFeeApr += manpointApr * beginningValue;
                        dispatchFeeTotalApr += dispatchFeeApr;

                        double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointMay * beginningValue).ToString("N0") + "</td>";

                        dispatchFeeMay += manpointMay * beginningValue;
                        dispatchFeeTotalMay += dispatchFeeMay;

                        double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJun * beginningValue).ToString("N0") + "</td>";

                        dispatchFeeJun += manpointJun * beginningValue;
                        dispatchFeeTotalJun += dispatchFeeJun;

                        double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJul * beginningValue).ToString("N0") + "</td>";

                        dispatchFeeJul += manpointJul * beginningValue;
                        dispatchFeeTotalJul += dispatchFeeJul;

                        double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointAug * beginningValue).ToString("N0") + "</td>";

                        dispatchFeeAug += manpointAug * beginningValue;
                        dispatchFeeTotalAug += dispatchFeeAug;

                        double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointSep * beginningValue).ToString("N0") + "</td>";

                        dispatchFeeSep += manpointSep * beginningValue;
                        dispatchFeeTotalSep += dispatchFeeSep;
                    }

                    // Provision for Employee Bonus
                    if (salaryTypeCount == 8)
                    {
                        double overWorkFixedTime = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).OverWorkFixedTime);
                        double bonusReserveConstant = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).BonusReserveConstant);
                        double valueOct = (totalRegularOct * overWorkFixedTime) / bonusReserveConstant;
                        valueOct = Double.IsNaN(valueOct) ? 0 : valueOct;
                        valueOct = Double.IsInfinity(valueOct) ? 0 : valueOct;
                        employeeBonusOct += valueOct;
                        employeeBonusTotalOct += employeeBonusOct;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + valueOct.ToString("N0") + "</td>";

                        double valueNov = (totalRegularNov * overWorkFixedTime) / bonusReserveConstant;
                        valueNov = Double.IsNaN(valueNov) ? 0 : valueNov;
                        valueNov = Double.IsInfinity(valueNov) ? 0 : valueNov;
                        employeeBonusNov += valueNov;
                        employeeBonusTotalNov += employeeBonusNov;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + valueNov.ToString("N0") + "</td>";

                        double valueDec = (totalRegularDec * overWorkFixedTime) / bonusReserveConstant;
                        valueDec = Double.IsNaN(valueDec) ? 0 : valueDec;
                        valueDec = Double.IsInfinity(valueDec) ? 0 : valueDec;
                        employeeBonusDec += valueDec;
                        employeeBonusTotalDec += employeeBonusDec;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + valueDec.ToString("N0") + "</td>";

                        double valueJan = (totalRegularJan * overWorkFixedTime) / bonusReserveConstant;
                        valueJan = Double.IsNaN(valueJan) ? 0 : valueJan;
                        valueJan = Double.IsInfinity(valueJan) ? 0 : valueJan;
                        employeeBonusJan += valueJan;
                        employeeBonusTotalJan += employeeBonusJan;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + valueJan.ToString("N0") + "</td>";

                        double valueFeb = (totalRegularFeb * overWorkFixedTime) / bonusReserveConstant;
                        valueFeb = Double.IsNaN(valueFeb) ? 0 : valueFeb;
                        valueFeb = Double.IsInfinity(valueFeb) ? 0 : valueFeb;
                        employeeBonusFeb += valueFeb;
                        employeeBonusTotalFeb += employeeBonusFeb;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + valueFeb.ToString("N0") + "</td>";

                        double valueMar = (totalRegularMar * overWorkFixedTime) / bonusReserveConstant;
                        valueMar = Double.IsNaN(valueMar) ? 0 : valueMar;
                        valueMar = Double.IsInfinity(valueMar) ? 0 : valueMar;
                        employeeBonusMar += valueMar;
                        employeeBonusTotalMar += employeeBonusMar;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + valueMar.ToString("N0") + "</td>";

                        double valueApr = (totalRegularApr * overWorkFixedTime) / bonusReserveConstant;
                        valueApr = Double.IsNaN(valueApr) ? 0 : valueApr;
                        valueApr = Double.IsInfinity(valueApr) ? 0 : valueApr;
                        employeeBonusApr += valueApr;
                        employeeBonusTotalApr += employeeBonusApr;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + valueApr.ToString("N0") + "</td>";

                        double valueMay = (totalRegularMay * overWorkFixedTime) / bonusReserveConstant;
                        valueMay = Double.IsNaN(valueMay) ? 0 : valueMay;
                        valueMay = Double.IsInfinity(valueMay) ? 0 : valueMay;
                        employeeBonusMay += valueMay;
                        employeeBonusTotalMay += employeeBonusMay;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + valueMay.ToString("N0") + "</td>";

                        double valueJun = (totalRegularJun * overWorkFixedTime) / bonusReserveConstant;
                        valueJun = Double.IsNaN(valueJun) ? 0 : valueJun;
                        valueJun = Double.IsInfinity(valueJun) ? 0 : valueJun;
                        employeeBonusJun += valueJun;
                        employeeBonusTotalJun += employeeBonusJun;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + valueJun.ToString("N0") + "</td>";

                        double valueJul = (totalRegularJul * overWorkFixedTime) / bonusReserveConstant;
                        valueJul = Double.IsNaN(valueJul) ? 0 : valueJul;
                        valueJul = Double.IsInfinity(valueJul) ? 0 : valueJul;
                        employeeBonusJul += valueJul;
                        employeeBonusTotalJul += employeeBonusJul;

                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + valueJul.ToString("N0") + "</td>";

                        double valueAug = (totalRegularAug * overWorkFixedTime) / bonusReserveConstant;
                        valueAug = Double.IsNaN(valueAug) ? 0 : valueAug;
                        valueAug = Double.IsInfinity(valueAug) ? 0 : valueAug;
                        employeeBonusAug += valueAug;
                        employeeBonusTotalAug += employeeBonusAug;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + valueAug.ToString("N0") + "</td>";

                        double valueSep = (totalRegularSep * overWorkFixedTime) / bonusReserveConstant;
                        valueSep = Double.IsNaN(valueSep) ? 0 : valueSep;
                        valueSep = Double.IsInfinity(valueSep) ? 0 : valueSep;
                        employeeBonusSep += valueSep;
                        employeeBonusTotalSep += employeeBonusSep;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + valueSep.ToString("N0") + "</td>";
                    }


                    //commuting expenses
                    if (salaryTypeCount == 9)
                    {
                        double manpointOct = GetTotalManPoints(salaryAssignmentDtos, 10, item.Grade.Id);
                        double beginningValue = _salaryBLL.GetGradeUnitPriceType(departmentId, 4, 2022, item.Grade.Id).GradeLowPoints;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointOct * beginningValue).ToString("N0") + "</td>";

                        commutingExpensesOct += manpointOct * beginningValue;
                        commutingExpensesTotalOct += commutingExpensesOct;
                        double manpointNov = GetTotalManPoints(salaryAssignmentDtos, 11, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointNov * beginningValue).ToString("N0") + "</td>";

                        commutingExpensesNov += manpointNov * beginningValue;
                        commutingExpensesTotalNov += commutingExpensesNov;

                        double manpointDec = GetTotalManPoints(salaryAssignmentDtos, 12, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointDec * beginningValue).ToString("N0") + "</td>";

                        commutingExpensesDec += manpointDec * beginningValue;
                        commutingExpensesTotalDec += commutingExpensesDec;

                        double manpointJan = GetTotalManPoints(salaryAssignmentDtos, 1, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJan * beginningValue).ToString("N0") + "</td>";

                        commutingExpensesJan += manpointJan * beginningValue;
                        commutingExpensesTotalJan += commutingExpensesJan;

                        double manpointFeb = GetTotalManPoints(salaryAssignmentDtos, 2, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointFeb * beginningValue).ToString("N0") + "</td>";

                        commutingExpensesFeb += manpointFeb * beginningValue;
                        commutingExpensesTotalFeb += commutingExpensesFeb;

                        double manpointMar = GetTotalManPoints(salaryAssignmentDtos, 3, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointMar * beginningValue).ToString("N0") + "</td>";

                        commutingExpensesMar += manpointMar * beginningValue;
                        commutingExpensesTotalMar += commutingExpensesMar;

                        double manpointApr = GetTotalManPoints(salaryAssignmentDtos, 4, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointApr * beginningValue).ToString("N0") + "</td>";

                        commutingExpensesApr += manpointApr * beginningValue;
                        commutingExpensesTotalApr += commutingExpensesApr;

                        double manpointMay = GetTotalManPoints(salaryAssignmentDtos, 5, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointMay * beginningValue).ToString("N0") + "</td>";

                        double manpointJun = GetTotalManPoints(salaryAssignmentDtos, 6, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJun * beginningValue).ToString("N0") + "</td>";

                        commutingExpensesJun += manpointJun * beginningValue;
                        commutingExpensesTotalJun += commutingExpensesJun;

                        double manpointJul = GetTotalManPoints(salaryAssignmentDtos, 7, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointJul * beginningValue).ToString("N0") + "</td>";

                        commutingExpensesJul += manpointJul * beginningValue;
                        commutingExpensesTotalJul += commutingExpensesJul;

                        double manpointAug = GetTotalManPoints(salaryAssignmentDtos, 8, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointAug * beginningValue).ToString("N0") + "</td>";

                        commutingExpensesAug += manpointAug * beginningValue;
                        commutingExpensesTotalAug += commutingExpensesAug;


                        double manpointSep = GetTotalManPoints(salaryAssignmentDtos, 9, item.Grade.Id);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (manpointSep * beginningValue).ToString("N0") + "</td>";

                        commutingExpensesSep += manpointSep * beginningValue;
                        commutingExpensesTotalSep += commutingExpensesSep;
                    }

                    // welfare expenses
                    if (salaryTypeCount == 10)
                    {
                        double commonValue = Convert.ToDouble(_commonMasterBLL.GetCommonMasters().SingleOrDefault(cm => cm.GradeId == item.Grade.Id).WelfareCostRatio);
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + ((totalSalaryOct + mWagesOct + commutingExpensesOct) * commonValue).ToString("N0") + "</td>";

                        wExpensesOct += (totalSalaryOct + mWagesOct + commutingExpensesOct) * commonValue;
                        welfareExpensesTotalOct += wExpensesOct;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + ((totalSalaryNov + mWagesNov + commutingExpensesNov) * commonValue).ToString("N0") + "</td>";

                        wExpensesNov += (totalSalaryNov + mWagesNov + commutingExpensesNov) * commonValue;
                        welfareExpensesTotalNov += wExpensesNov;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + ((totalSalaryDec + mWagesDec + commutingExpensesDec) * commonValue).ToString("N0") + "</td>";

                        wExpensesDec += (totalSalaryDec + mWagesDec + commutingExpensesDec) * commonValue;
                        welfareExpensesTotalDec += wExpensesDec;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + ((totalSalaryJan + mWagesJan + commutingExpensesJan) * commonValue).ToString("N0") + "</td>";

                        wExpensesJan += (totalSalaryJan + mWagesJan + commutingExpensesJan) * commonValue;
                        welfareExpensesTotalJan += wExpensesJan;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + ((totalSalaryFeb + mWagesFeb + commutingExpensesFeb) * commonValue).ToString("N0") + "</td>";

                        wExpensesFeb += (totalSalaryFeb + mWagesFeb + commutingExpensesFeb) * commonValue;
                        welfareExpensesTotalFeb += wExpensesFeb;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + ((totalSalaryMar + mWagesMar + commutingExpensesMar) * commonValue).ToString("N0") + "</td>";

                        wExpensesMar += (totalSalaryMar + mWagesMar + commutingExpensesMar) * commonValue;
                        welfareExpensesTotalMar += wExpensesMar;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + ((totalSalaryApr + mWagesApr + commutingExpensesApr) * commonValue).ToString("N0") + "</td>";

                        wExpensesApr += (totalSalaryApr + mWagesApr + commutingExpensesApr) * commonValue;
                        welfareExpensesTotalApr += wExpensesApr;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + ((totalSalaryMay + mWagesMay + commutingExpensesMay) * commonValue).ToString("N0") + "</td>";

                        wExpensesMay += (totalSalaryMay + mWagesMay + commutingExpensesMay) * commonValue;
                        welfareExpensesTotalMay += wExpensesMay;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + ((totalSalaryJun + mWagesJun + commutingExpensesJun) * commonValue).ToString("N0") + "</td>";

                        wExpensesJun += (totalSalaryJun + mWagesJun + commutingExpensesJun) * commonValue;
                        welfareExpensesTotalJun += wExpensesJun;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + ((totalSalaryJul + mWagesJul + commutingExpensesJul) * commonValue).ToString("N0") + "</td>";

                        wExpensesJul += (totalSalaryJul + mWagesJul + commutingExpensesJul) * commonValue;
                        welfareExpensesTotalJul += wExpensesJul;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + ((totalSalaryAug + mWagesAug + commutingExpensesAug) * commonValue).ToString("N0") + "</td>";

                        wExpensesAug += (totalSalaryAug + mWagesAug + commutingExpensesAug) * commonValue;
                        welfareExpensesTotalAug += wExpensesAug;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + ((totalSalarySep + mWagesSep + commutingExpensesSep) * commonValue).ToString("N0") + "</td>";

                        wExpensesSep += (totalSalarySep + mWagesSep + commutingExpensesSep) * commonValue;
                        welfareExpensesTotalSep += wExpensesSep;
                    }

                    // welfare expenses bonus
                    if (salaryTypeCount == 11)
                    {
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + 0.ToString("N0") + "</td>";

                    }


                    // total statutory
                    if (salaryTypeCount == 12)
                    {
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (wExpensesOct + wExpBonusOct).ToString("N0") + "</td>";

                        totalStatutoryOct += wExpensesOct + wExpBonusOct;
                        statutoryTotalOct += totalStatutoryOct;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (wExpensesNov + wExpBonusNov).ToString("N0") + "</td>";

                        totalStatutoryNov += wExpensesNov + wExpBonusNov;
                        statutoryTotalNov += totalStatutoryNov;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (wExpensesDec + wExpBonusDec).ToString("N0") + "</td>";

                        totalStatutoryDec += wExpensesDec + wExpBonusDec;
                        statutoryTotalDec += totalStatutoryDec;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (wExpensesJan + wExpBonusJan).ToString("N0") + "</td>";

                        totalStatutoryJan += wExpensesJan + wExpBonusJan;
                        statutoryTotalJan += totalStatutoryJan;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (wExpensesFeb + wExpBonusFeb).ToString("N0") + "</td>";

                        totalStatutoryFeb += wExpensesFeb + wExpBonusFeb;
                        statutoryTotalFeb += totalStatutoryFeb;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (wExpensesMar + wExpBonusMar).ToString("N0") + "</td>";

                        totalStatutoryMar += wExpensesMar + wExpBonusMar;
                        statutoryTotalMar += totalStatutoryMar;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (wExpensesApr + wExpBonusApr).ToString("N0") + "</td>";

                        totalStatutoryApr += wExpensesApr + wExpBonusApr;
                        statutoryTotalApr += totalStatutoryApr;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (wExpensesMay + wExpBonusMay).ToString("N0") + "</td>";

                        totalStatutoryMay += wExpensesMay + wExpBonusMay;
                        statutoryTotalMay += totalStatutoryMay;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (wExpensesJun + wExpBonusJun).ToString("N0") + "</td>";

                        totalStatutoryJun += wExpensesJun + wExpBonusJun;
                        statutoryTotalJun += totalStatutoryJun;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (wExpensesJul + wExpBonusJul).ToString("N0") + "</td>";

                        totalStatutoryJul += wExpensesJul + wExpBonusJul;
                        statutoryTotalJul += totalStatutoryJul;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (wExpensesAug + wExpBonusAug).ToString("N0") + "</td>";

                        totalStatutoryAug += wExpensesAug + wExpBonusAug;
                        statutoryTotalAug += totalStatutoryAug;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (wExpensesSep + wExpBonusSep).ToString("N0") + "</td>";

                        totalStatutorySep += wExpensesSep + wExpBonusSep;
                        statutoryTotalSep += totalStatutorySep;

                    }

                    // total expenses
                    if (salaryTypeCount == 13)
                    {
                        expensesTotalOct += totalSalaryOct + mWagesOct + dispatchFeeOct + employeeBonusOct + commutingExpensesOct + totalStatutoryOct;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalSalaryOct + mWagesOct + dispatchFeeOct + employeeBonusOct + commutingExpensesOct + totalStatutoryOct).ToString("N0") + "</td>";

                        expensesTotalNov += totalSalaryNov + mWagesNov + dispatchFeeNov + employeeBonusNov + commutingExpensesNov + totalStatutoryNov;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalSalaryNov + mWagesNov + dispatchFeeNov + employeeBonusNov + commutingExpensesNov + totalStatutoryNov).ToString("N0") + "</td>";

                        expensesTotalDec += totalSalaryDec + mWagesDec + dispatchFeeDec + employeeBonusDec + commutingExpensesDec + totalStatutoryDec;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalSalaryDec + mWagesDec + dispatchFeeDec + employeeBonusDec + commutingExpensesDec + totalStatutoryDec).ToString("N0") + "</td>";

                        expensesTotalJan += totalSalaryJan + mWagesJan + dispatchFeeJan + employeeBonusJan + commutingExpensesJan + totalStatutoryJan;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalSalaryJan + mWagesJan + dispatchFeeJan + employeeBonusJan + commutingExpensesJan + totalStatutoryJan).ToString("N0") + "</td>";

                        expensesTotalFeb += totalSalaryFeb + mWagesFeb + dispatchFeeFeb + employeeBonusFeb + commutingExpensesFeb + totalStatutoryFeb;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalSalaryFeb + mWagesFeb + dispatchFeeFeb + employeeBonusFeb + commutingExpensesFeb + totalStatutoryFeb).ToString("N0") + "</td>";

                        expensesTotalMar += totalSalaryMar + mWagesMar + dispatchFeeMar + employeeBonusMar + commutingExpensesMar + totalStatutoryMar;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalSalaryMar + mWagesMar + dispatchFeeMar + employeeBonusMar + commutingExpensesMar + totalStatutoryMar).ToString("N0") + "</td>";

                        expensesTotalApr += totalSalaryApr + mWagesApr + dispatchFeeApr + employeeBonusApr + commutingExpensesApr + totalStatutoryApr;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalSalaryApr + mWagesApr + dispatchFeeApr + employeeBonusApr + commutingExpensesApr + totalStatutoryApr).ToString("N0") + "</td>";

                        expensesTotalMay += totalSalaryMay + mWagesMay + dispatchFeeMay + employeeBonusMay + commutingExpensesMay + totalStatutoryMay;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalSalaryMay + mWagesMay + dispatchFeeMay + employeeBonusMay + commutingExpensesMay + totalStatutoryMay).ToString("N0") + "</td>";

                        expensesTotalJun += totalSalaryJun + mWagesJun + dispatchFeeJun + employeeBonusJun + commutingExpensesJun + totalStatutoryJun;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalSalaryJun + mWagesJun + dispatchFeeJun + employeeBonusJun + commutingExpensesJun + totalStatutoryJun).ToString("N0") + "</td>";

                        expensesTotalJul += totalSalaryJul + mWagesJul + dispatchFeeJul + employeeBonusJul + commutingExpensesJul + totalStatutoryJul;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalSalaryJul + mWagesJul + dispatchFeeJul + employeeBonusJul + commutingExpensesJul + totalStatutoryJul).ToString("N0") + "</td>";

                        expensesTotalAug += totalSalaryAug + mWagesAug + dispatchFeeAug + employeeBonusAug + commutingExpensesAug + totalStatutoryAug;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalSalaryAug + mWagesAug + dispatchFeeAug + employeeBonusAug + commutingExpensesAug + totalStatutoryAug).ToString("N0") + "</td>";

                        expensesTotalSep += totalSalarySep + mWagesSep + dispatchFeeSep + employeeBonusSep + commutingExpensesSep + totalStatutorySep;
                        strOtherGradeTd = strOtherGradeTd + "<td style='text-align:right;'>" + (totalSalarySep + mWagesSep + dispatchFeeSep + employeeBonusSep + commutingExpensesSep + totalStatutorySep).ToString("N0") + "</td>";
                    }
                    rowCount++;
                    innerCount++;
                    salaryTypeCount++;
                }

                strOtherGradeTd = strOtherGradeTd + "</tr>";
                count++;

            }

            string salaryTypeTd = "";
            //region salary types total
            int salaryTypeSummeryCount = 1;
            foreach (var item in _unitPriceTypeBLL.GetAllSalaryTypes())
            {
                salaryTypeTd = salaryTypeTd + "<tr>";
                salaryTypeTd = salaryTypeTd + "<td>" + item.SalaryTypeName + "</td>";
                //salaryTypeTd = salaryTypeTd + "</tr>";
                // executive
                if (salaryTypeSummeryCount == 1)
                {
                    salaryTypeTd = salaryTypeTd + "<td></td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + 0 + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + 0 + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + 0 + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + 0 + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + 0 + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + 0 + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + 0 + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + 0 + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + 0 + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + 0 + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + 0 + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + 0 + "</td>";
                }

                // regular
                if (salaryTypeSummeryCount == 2)
                {
                    salaryTypeTd = salaryTypeTd + "<td></td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + regularSalaryTotalOct.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + regularSalaryTotalNov.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + regularSalaryTotalDec.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + regularSalaryTotalJan.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + regularSalaryTotalFeb.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + regularSalaryTotalMar.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + regularSalaryTotalApr.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + regularSalaryTotalMay.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + regularSalaryTotalJun.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + regularSalaryTotalJul.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + regularSalaryTotalAug.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + regularSalaryTotalSep.ToString("N0") + "</td>";
                }
                // fixed
                if (salaryTypeSummeryCount == 3)
                {
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'></td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + fixedSalaryTotalOct.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + fixedSalaryTotalNov.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + fixedSalaryTotalDec.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + fixedSalaryTotalJan.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + fixedSalaryTotalFeb.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + fixedSalaryTotalMar.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + fixedSalaryTotalApr.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + fixedSalaryTotalMay.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + fixedSalaryTotalJun.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + fixedSalaryTotalJul.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + fixedSalaryTotalAug.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + fixedSalaryTotalSep.ToString("N0") + "</td>";
                }
                // overtime
                if (salaryTypeSummeryCount == 4)
                {
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'></td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + overTimeSalaryTotalOct.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + overTimeSalaryTotalNov.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + overTimeSalaryTotalDec.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + overTimeSalaryTotalJan.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + overTimeSalaryTotalFeb.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + overTimeSalaryTotalMar.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + overTimeSalaryTotalApr.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + overTimeSalaryTotalMay.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + overTimeSalaryTotalJun.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + overTimeSalaryTotalJul.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + overTimeSalaryTotalAug.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + overTimeSalaryTotalSep.ToString("N0") + "</td>";

                }
                // total salary
                if (salaryTypeSummeryCount == 5)
                {
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'></td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + salaryTotalOct.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + salaryTotalNov.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + salaryTotalDec.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + salaryTotalJan.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + salaryTotalFeb.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + salaryTotalMar.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + salaryTotalApr.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + salaryTotalMay.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + salaryTotalJun.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + salaryTotalJul.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + salaryTotalAug.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + salaryTotalSep.ToString("N0") + "</td>";
                }
                // miscellaneous wages
                if (salaryTypeSummeryCount == 6)
                {
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'></td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + mwTotalOct.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + mwTotalNov.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + mwTotalDec.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + mwTotalJan.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + mwTotalFeb.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + mwTotalMar.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + mwTotalApr.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + mwTotalMay.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + mwTotalJun.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + mwTotalJul.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + mwTotalAug.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + mwTotalSep.ToString("N0") + "</td>";
                }
                // dispatch fee
                if (salaryTypeSummeryCount == 7)
                {
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'></td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + dispatchFeeTotalOct.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + dispatchFeeTotalNov.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + dispatchFeeTotalDec.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + dispatchFeeTotalJan.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + dispatchFeeTotalFeb.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + dispatchFeeTotalMar.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + dispatchFeeTotalApr.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + dispatchFeeTotalMay.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + dispatchFeeTotalJul.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + dispatchFeeTotalAug.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + dispatchFeeTotalSep.ToString("N0") + "</td>";
                }
                // provision for employee bonus
                if (salaryTypeSummeryCount == 8)
                {
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'></td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + employeeBonusTotalOct.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + employeeBonusTotalNov.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + employeeBonusTotalDec.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + employeeBonusTotalJan.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + employeeBonusTotalFeb.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + employeeBonusTotalMar.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + employeeBonusTotalApr.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + employeeBonusTotalMay.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + employeeBonusTotalJun.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + employeeBonusTotalJul.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + employeeBonusTotalAug.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + employeeBonusTotalSep.ToString("N0") + "</td>";
                }
                // commuting expenses
                if (salaryTypeSummeryCount == 9)
                {
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'></td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + commutingExpensesTotalOct.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + commutingExpensesTotalNov.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + commutingExpensesTotalDec.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + commutingExpensesTotalJan.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + commutingExpensesTotalFeb.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + commutingExpensesTotalMar.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + commutingExpensesTotalApr.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + commutingExpensesTotalMay.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + commutingExpensesTotalJun.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + commutingExpensesTotalJul.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + commutingExpensesTotalAug.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + commutingExpensesTotalSep.ToString("N0") + "</td>";
                }
                // salary statutory welfare expenses
                if (salaryTypeSummeryCount == 10)
                {
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'></td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesTotalOct.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesTotalNov.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesTotalDec.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesTotalJan.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesTotalFeb.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesTotalMar.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesTotalApr.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesTotalMay.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesTotalJun.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesTotalJul.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesTotalAug.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesTotalSep.ToString("N0") + "</td>";
                }
                // provision for statotory welfare expenses for bonus
                if (salaryTypeSummeryCount == 11)
                {
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'></td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesBonusTotalOct.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesBonusTotalNov.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesBonusTotalDec.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesBonusTotalJan.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesBonusTotalFeb.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesBonusTotalMar.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesBonusTotalApr.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesBonusTotalMay.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesBonusTotalJun.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesBonusTotalJul.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesBonusTotalAug.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + welfareExpensesBonusTotalSep.ToString("N0") + "</td>";
                }
                // total statutory benefites
                if (salaryTypeSummeryCount == 12)
                {
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'></td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + statutoryTotalOct.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + statutoryTotalNov.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + statutoryTotalDec.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + statutoryTotalJan.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + statutoryTotalFeb.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + statutoryTotalMar.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + statutoryTotalApr.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + statutoryTotalMay.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + statutoryTotalJun.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + statutoryTotalJul.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + statutoryTotalAug.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + statutoryTotalSep.ToString("N0") + "</td>";
                }
                // total expenses
                if (salaryTypeSummeryCount == 13)
                {
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'></td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + expensesTotalOct.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + expensesTotalNov.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + expensesTotalDec.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + expensesTotalJan.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + expensesTotalFeb.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + expensesTotalMar.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + expensesTotalApr.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + expensesTotalMay.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + expensesTotalJun.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + expensesTotalJul.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + expensesTotalAug.ToString("N0") + "</td>";
                    salaryTypeTd = salaryTypeTd + "<td style='text-align:right;'>" + expensesTotalSep.ToString("N0") + "</td>";
                }
                salaryTypeTd = salaryTypeTd + "</tr>";
                rowCountForFuture++;
                salaryTypeSummeryCount++;
            }
            //endregion

            rowCount = rowCount + 2;
            string strAllocationCostingBody = "";
            strAllocationCostingBody = strAllocationCostingBody + "<tr>";
            strAllocationCostingBody = strAllocationCostingBody + "<td rowspan='16'>Costing</td>";
            strAllocationCostingBody = strAllocationCostingBody + "</tr>";
            rowCount++;

            octTotal = 0; novTotal = 0; decTotal = 0; janTotal = 0; febTotal = 0; marTotal = 0; aprTotal = 0; mayTotal = 0; junTotal = 0; julTotal = 0; augTotal = 0; sepTotal = 0;
            int indexCount = 1;
            // costing
            foreach (var item in salaryAssignmentDtos)
            {
                strAllocationCostingBody = strAllocationCostingBody + "<tr>";
                if(indexCount == 15)
                {
                    strAllocationCostingBody = strAllocationCostingBody + "<td>合計</td>";
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>"+ octTotal.ToString("N0") + "</td>";
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + novTotal.ToString("N0") + "</td>";
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + decTotal.ToString("N0") + "</td>";
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + janTotal.ToString("N0") + "</td>";
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + febTotal.ToString("N0") + "</td>";
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + marTotal.ToString("N0") + "</td>";
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + aprTotal.ToString("N0") + "</td>";
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + mayTotal.ToString("N0") + "</td>";
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + junTotal.ToString("N0") + "</td>";
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + julTotal.ToString("N0") + "</td>";
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + augTotal.ToString("N0") + "</td>";
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + sepTotal.ToString("N0") + "</td>";
                    rowCount++;
                }
                else
                {
                    strAllocationCostingBody = strAllocationCostingBody + "<td>" + item.Grade.GradeName + "</td>";
                    foreach (var singleAssignment in item.ForecastAssignmentViewModels)
                    {
                        oct += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 10).Total);
                        nov += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 11).Total);
                        dec += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 12).Total);
                        jan += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 1).Total);
                        feb += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 2).Total);
                        mar += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 3).Total);
                        apr += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 4).Total);
                        may += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 5).Total);
                        jun += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 6).Total);
                        jul += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 7).Total);
                        aug += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 8).Total);
                        sep += Convert.ToDecimal(singleAssignment.forecasts.SingleOrDefault(f => f.Month == 9).Total);
                    }
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + oct.ToString("N0") + "</td>";
                    octTotal += oct;
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + nov.ToString("N0") + "</td>";
                    novTotal += nov;
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + dec.ToString("N0") + "</td>";
                    decTotal += dec;
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + jan.ToString("N0") + "</td>";
                    janTotal += jan;
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + feb.ToString("N0") + "</td>";
                    febTotal += feb;
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + mar.ToString("N0") + "</td>";
                    marTotal += mar;
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + apr.ToString("N0") + "</td>";
                    aprTotal += apr;
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + may.ToString("N0") + "</td>";
                    mayTotal += may;
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + jun.ToString("N0") + "</td>";
                    junTotal += jun;
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + jul.ToString("N0") + "</td>";
                    julTotal += jul;
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + aug.ToString("N0") + "</td>";
                    augTotal += aug;
                    strAllocationCostingBody = strAllocationCostingBody + "<td style='text-align:right;'>" + sep.ToString("N0") + "</td>";
                    sepTotal += sep;
                }                


                

                indexCount +=1;                
                strAllocationCostingBody = strAllocationCostingBody + "</tr>";
                oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                rowCount++;
            }
            rowCount = rowCount + 2;
            string strCompanyWiseAllocationBody = "";
            strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<tr>";
            strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='font-weight: bold;'>工数</td>";
            strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "</tr>";

            if (assignmentsWithSectionAndCompany.Count > 0)
            {
                List<ForecastAssignmentViewModel> customizedList = assignmentsWithSectionAndCompany.ToList();
                List<ForecastAssignmentViewModel> tempList = null;
                octTotal = 0; novTotal = 0; decTotal = 0; janTotal = 0; febTotal = 0; marTotal = 0; aprTotal = 0; mayTotal = 0; junTotal = 0; julTotal = 0; augTotal = 0; sepTotal = 0;

                foreach (var item in assignmentsWithSectionAndCompany)
                {
                    strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<tr>";                    

                    tempList = customizedList.Where(l => l.SectionId == item.SectionId && l.CompanyId == item.CompanyId).ToList();
                    if (tempList.Count > 0)
                    {
                        foreach (var removableItem in tempList)
                        {
                            customizedList.Remove(removableItem);
                        }
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:left;'>" + tempList[0].SectionName + "</td>";
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:left;'>" + tempList[0].CompanyName + "</td>";
                        foreach (var tempItem in tempList)
                        {
                            oct += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 10).Points);
                            nov += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 11).Points);
                            dec += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 12).Points);
                            jan += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 1).Points);
                            feb += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 2).Points);
                            mar += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 3).Points);
                            apr += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 4).Points);
                            may += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 5).Points);
                            jun += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 6).Points);
                            jul += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 7).Points);
                            aug += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 8).Points);
                            sep += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 9).Points);
                        }
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + oct + "</td>";                        
                        octTotal += oct;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + nov + "</td>";                        
                        novTotal += nov;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + dec + "</td>";                        
                        decTotal += dec;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + jan + "</td>";                        
                        janTotal += jan;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + feb + "</td>";                        
                        febTotal += feb;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + mar + "</td>";                        
                        marTotal += mar;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + apr + "</td>";                        
                        aprTotal += apr;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + may + "</td>";                       
                        mayTotal += may;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + jun + "</td>";                        
                        junTotal += jun;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + jul + "</td>";                        
                        julTotal += jul;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + aug + "</td>";                        
                        augTotal += aug;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + sep + "</td>";                        
                        sepTotal += sep;

                        oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                        rowCount++;
                    }
                    strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "</tr>";
                }
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<tr>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td></td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:left;'>合計</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + octTotal + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + novTotal + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + decTotal + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + janTotal + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + febTotal + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + marTotal + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + aprTotal + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + mayTotal + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + junTotal + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + julTotal + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + augTotal + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td>" + sepTotal + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "</tr>";
                rowCount++;

                rowCount = rowCount + 2;

                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<tr>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='font-weight: bold;'>金額</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "</tr>";
                rowCount++;

                customizedList = assignmentsWithSectionAndCompany.ToList();
                octTotal = 0; novTotal = 0; decTotal = 0; janTotal = 0; febTotal = 0; marTotal = 0; aprTotal = 0; mayTotal = 0; junTotal = 0; julTotal = 0; augTotal = 0; sepTotal = 0;
                foreach (var item in assignmentsWithSectionAndCompany)
                {
                    strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<tr>";                    
                    tempList = customizedList.Where(l => l.SectionId == item.SectionId && l.CompanyId == item.CompanyId).ToList();
                    if (tempList.Count > 0)
                    {
                        foreach (var removableItem in tempList)
                        {
                            customizedList.Remove(removableItem);
                        }
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:left;'>" + tempList[0].SectionName + "</td>";
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:left;'>" + tempList[0].CompanyName + "</td>";

                        foreach (var tempItem in tempList)
                        {

                            oct += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 10).Total);
                            nov += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 11).Total);
                            dec += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 12).Total);
                            jan += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 1).Total);
                            feb += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 2).Total);
                            mar += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 3).Total);
                            apr += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 4).Total);
                            may += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 5).Total);
                            jun += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 6).Total);
                            jul += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 7).Total);
                            aug += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 8).Total);
                            sep += Convert.ToDecimal(tempItem.forecasts.SingleOrDefault(f => f.Month == 9).Total);
                        }
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + oct.ToString("N0") + "</td>";                        
                        octTotal += oct;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + nov.ToString("N0") + "</td>";                        
                        novTotal += nov;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + dec.ToString("N0") + "</td>";                        
                        decTotal += dec;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + jan.ToString("N0") + "</td>";                        
                        janTotal += jan;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + feb.ToString("N0") + "</td>";                        
                        febTotal += feb;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + mar.ToString("N0") + "</td>";                        
                        marTotal += mar;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + apr.ToString("N0") + "</td>";                        
                        aprTotal += apr;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + may.ToString("N0") + "</td>";                        
                        mayTotal += may;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + jun.ToString("N0") + "</td>";                        
                        junTotal += jun;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + jul.ToString("N0") + "</td>";                        
                        julTotal += jul;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + aug.ToString("N0") + "</td>";                        
                        augTotal += aug;
                        strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + sep.ToString("N0") + "</td>";                        
                        sepTotal += sep;                        

                        oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                        rowCount++;
                    }
                    strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "</tr>";
                }
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<tr>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td></td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:left;'>合計</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + octTotal.ToString("N0") + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + novTotal.ToString("N0") + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + decTotal.ToString("N0") + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + janTotal.ToString("N0") + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + febTotal.ToString("N0") + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + marTotal.ToString("N0") + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + aprTotal.ToString("N0") + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + mayTotal.ToString("N0") + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + junTotal.ToString("N0") + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + julTotal.ToString("N0") + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + augTotal.ToString("N0") + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "<td style='text-align:right;'>" + sepTotal.ToString("N0") + "</td>";
                strCompanyWiseAllocationBody = strCompanyWiseAllocationBody + "</tr>";              
                rowCount++;
            }


            //get the allocation list table
            string allocationList = "";
            allocationList = allocationList + "<table class='' id='allocation_wise_view_list'> ";
            allocationList = allocationList + "    <thead id='allocation_wise_tbl_header'>";
            allocationList = allocationList + "        <tr>";
            allocationList = allocationList + "            <th></th>";
            allocationList = allocationList + "            <th></th>";
            allocationList = allocationList + "            <th class='man_month_cost_allocation'>10</th>";
            allocationList = allocationList + "            <th class='man_month_cost_allocation'>11</th>";
            allocationList = allocationList + "            <th class='man_month_cost_allocation'>12</th>";
            allocationList = allocationList + "            <th class='man_month_cost_allocation'>1</th>";
            allocationList = allocationList + "            <th class='man_month_cost_allocation'>2</th>";
            allocationList = allocationList + "            <th class='man_month_cost_allocation'>3</th>";
            allocationList = allocationList + "            <th class='man_month_cost_allocation'>4</th>";
            allocationList = allocationList + "            <th class='man_month_cost_allocation'>5</th>";
            allocationList = allocationList + "            <th class='man_month_cost_allocation'>6</th>";
            allocationList = allocationList + "            <th class='man_month_cost_allocation'>7</th>";
            allocationList = allocationList + "            <th class='man_month_cost_allocation'>8</th>";
            allocationList = allocationList + "            <th class='man_month_cost_allocation'>9</th>";
            allocationList = allocationList + "        </tr>";
            allocationList = allocationList + "    </thead>";
            allocationList = allocationList + "    <tbody id='view_data_allocation_wise'>";

            allocationList = allocationList + allocaionTableBodyTd + salaryTypeTd+ strOtherGradeTd+ strAllocationCostingBody+ strCompanyWiseAllocationBody;

            allocationList = allocationList + "    </tbody>";
            allocationList = allocationList + "</table> ";

            _exportDepartmentAllocationViewModel.AllocationWiseExportList = allocationList;
            _exportDepartmentAllocationViewModel.SalaryMaster = "<div><table><th>sl</th><tr><td>salary master</td></tr></table></div>";
            _exportDepartmentAllocationViewModel.CommonMaster = "<div><table><th>sl</th><tr><td>common master</td></tr></table></div>";
            return Json(_exportDepartmentAllocationViewModel, JsonRequestBehavior.AllowGet);
        }
    }
}