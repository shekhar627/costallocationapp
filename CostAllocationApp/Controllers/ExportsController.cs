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

        public ExportsController()
        {
            _departmentBLL = new DepartmentBLL();
            _exportBLL = new ExportBLL();
            _salaryBLL = new SalaryBLL();
            _explanationsBLL = new ExplanationsBLL();
            _gradeBll = new GradeBLL();
            _commonMasterBLL = new CommonMasterBLL();
            _unitPriceTypeBLL = new UnitPriceTypeBLL();
        }
        // GET: Exports
        public FileResult ExportBySection(int sectionId=0,string sectionName = "")
        {
            List<ForecastAssignmentViewModel> forecastAssignmentViewModels = new List<ForecastAssignmentViewModel>();
            using (var client = new HttpClient())
            {
                string uri = "http://198.38.92.119:8081/api/utilities/SearchForecastEmployee?employeeName=&sectionId=" + sectionId + "&departmentId=&inchargeId=&roleId=&explanationId=&companyId=&status=&year=";
                //string uri = "http://localhost:59198/api/utilities/SearchForecastEmployee?employeeName=&sectionId=" + sectionId + "&departmentId=&inchargeId=&roleId=&explanationId=&companyId=&status=&year=";
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
                            if (forecastItem.Points==0)
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
                htmlTable+=@"
                 </tbody>
               </table>
                </body>
                </html> ";
            }
            UTF8Encoding uTF8Encoding = new UTF8Encoding();
            return File(uTF8Encoding.GetBytes(htmlTable), "application/vnd.ms-excel", sectionName+".xls");
        }

        //public FileResult PlanningLayout()
        //{
        //    List<ForecastAssignmentViewModel> forecastAssignmentViewModels = new List<ForecastAssignmentViewModel>();
        //    using (var client = new HttpClient())
        //    {
        //        client.BaseAddress = new Uri("http://198.38.92.119:8081/api/utilities/SearchForecastEmployee?employeeName=&sectionId=&departmentId=&inchargeId=&roleId=&explanationId=&companyId=&status=&year=");
        //        //HTTP GET
        //        var responseTask = client.GetAsync("");
        //        responseTask.Wait();

        //        var result = responseTask.Result;
        //        if (result.IsSuccessStatusCode)
        //        {
        //            var readTask = result.Content.ReadAsAsync<List<ForecastAssignmentViewModel>>();
        //            readTask.Wait();

        //            forecastAssignmentViewModels = readTask.Result;
        //        }
        //        else //web api sent error response 
        //        {
        //            //log response status here..

        //            forecastAssignmentViewModels = new List<ForecastAssignmentViewModel>();

        //            // ModelState.AddModelError(string.Empty, "Server error. Please contact administrator.");
        //        }
        //    }
        //    string htmlTable = "";
        //    if (forecastAssignmentViewModels.Count > 0)
        //    {
        //        int i = 1;
        //        htmlTable = @"<!DOCTYPE html>
        //        <html>
        //            <head>
        //                <style>
        //                #dev_placement {
        //                  border-collapse: collapse;
        //                  width: 50%;
        //                }

        //                #dev_placement td, #dev_placement th {
        //                  border: 1px solid #ddd;
        //                }
        //                </style>
        //            </head>
        //        <body>

        //        <table id='dev_placement'>
        //        <thead>
        //          <tr>
        //            <th>No</th>
        //            <th>FullName</th>
        //            <th>Section</th>
        //            <th>Department</th>
        //            <th>Grade</th>
        //            <th>UnitPrice</th>
        //            <th>10</th>
        //            <th>11</th>
        //            <th>12</th>
        //            <th>1</th>
        //            <th>2</th>
        //            <th>3</th>
        //            <th>4</th>
        //            <th>5</th>
        //            <th>6</th>
        //            <th>7</th>
        //            <th>8</th>
        //            <th>9</th>
        //          </tr>
        //        </thead>
        //        <tbody>";
        //        foreach (var item in forecastAssignmentViewModels)
        //        {
        //            string td = "";
        //            if (item.forecasts.Count > 0)
        //            {
        //                foreach (var forecastItem in item.forecasts)
        //                {
        //                    if (forecastItem.Points == 0)
        //                    {
        //                        td += $@"<td style='background-color:#808080;padding-right:20px;'> </td>";
        //                    }
        //                    else
        //                    {
        //                        td += $@"<td style='padding-right:20px;'>{forecastItem.Points}</td>";
        //                    }

        //                }
        //            }
        //            htmlTable += $@"
        //                  <tr>
        //                    <td>{i}</td>
        //                    <td>{item.EmployeeName}</td>
        //                    <td>{item.SectionName}</td>
        //                    <td>{item.DepartmentName}</td>
        //                    <td>{item.GradePoint}</td>
        //                    <td>{item.UnitPrice}</td>
        //                    {td}
        //                  </tr>
        //            ";
        //            i++;
        //        }
        //        htmlTable += @"
        //         </tbody>
        //       </table>
        //        </body>
        //        </html> ";
        //    }
        //    UTF8Encoding uTF8Encoding = new UTF8Encoding();
        //    return File(uTF8Encoding.GetBytes(htmlTable), "application/vnd.ms-excel", "PlanningLayout.xls");
        //}


        public ActionResult DataExportByDepartment()
        {
            return View(new ExportViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }

        [HttpPost]
        public ActionResult DataExportByDepartment(int departmentId=0)
        {
            List<ForecastAssignmentViewModel> forecastAssignmentViewModels = new List<ForecastAssignmentViewModel>();
            var _department = _departmentBLL.GetDepartmentByDepartemntId(departmentId);

            using (var client = new HttpClient())
            {
                string uri = "http://198.38.92.119:8081/api/utilities/SearchForecastEmployee?employeeName=&sectionId=&departmentId=" + departmentId + "&inchargeId=&roleId=&explanationId=&companyId=&status=&year=";
                //string uri = "http://localhost:59198/api/utilities/SearchForecastEmployee?employeeName=&sectionId=&departmentId="+departmentId+"&inchargeId=&roleId=&explanationId=&companyId=&status=&year=";
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
                if (forecastItem.CompanyName.ToLower()!="mw" || forecastItem.SectionName.ToLower()!="mw")
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
                sheet.Cells["B1"].Style.Font.Bold = true;;
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

                if (forecastAssignmentViewModels.Count>0)
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
                var fileName = _department.DepartmentName+ ".xlsx";
                return File(excelData, contentType, fileName);
            }
        }

        public ActionResult DataExportByAllocation()
        {
            return View(new ExportViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }

        class ValuesWithGrade
        {
            public string GradeName { get; set; }
            public int GradeId { get; set; }
            public float Point { get; set; }
            public int MonthId { get; set; }
        }

        [HttpPost]
        public ActionResult DataExportByAllocation(int departmentId = 0, int explanationId = 0)
        {

            List<Grade> grades = _gradeBll.GetAllGrade();



            List<SalaryAssignmentDto> salaryAssignmentDtos = new List<SalaryAssignmentDto>();
            var department = _departmentBLL.GetDepartmentByDepartemntId(departmentId);
            var explanation = _explanationsBLL.GetExplanationByExplanationId(explanationId);
            var salaries = _salaryBLL.GetAllSalaryPoints();
            List<ValuesWithGrade> valuesWithGrades = new List<ValuesWithGrade>();
           

            // get all data with department and allocation
            List<ForecastAssignmentViewModel> assignmentsWithForecast = _exportBLL.AssignmentsByAllocation(departmentId, explanationId);
            //filtered by grade id
            List<ForecastAssignmentViewModel> assignmentsWithGrade = assignmentsWithForecast.Where(a=>a.GradeId!=null).ToList();
            //filtered by section and company
            List<ForecastAssignmentViewModel> assignmentsWithSectionAndCompany = assignmentsWithForecast.Where(a=>a.SectionId!=null && a.CompanyId!=null).ToList();

            foreach (var item in grades)
            {
                List<ForecastAssignmentViewModel> forecastAssignmentViews = new List<ForecastAssignmentViewModel>();

                SalaryAssignmentDto salaryAssignmentDto = new SalaryAssignmentDto();
                salaryAssignmentDto.Grade = item;

                var gradeSalaryTypes = _exportBLL.GetGradeSalaryTypes(item.Id,departmentId,2022,2);

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
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

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
                sheet.Cells[3, 1].Style.Fill.BackgroundColor.SetColor(1,36,64,98);

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

                #region common master
                sheet.Cells[rowCount, 1].Value = "1人あたりの時間外勤務見込 \n (みなし時間（固定時間）\n を含む残業時間 \n を入力してください";
                foreach (var item in _commonMasterBLL.GetCommonMasters())
                {
                    sheet.Cells[rowCount, 2].Value = item.GradeName;
                    sheet.Cells[rowCount, 3].Value = item.OverWorkFixedTime;
                 
                    sheet.Cells[rowCount, 4].Value = item.OverWorkFixedTime;
                    
                    sheet.Cells[rowCount, 5].Value = item.OverWorkFixedTime;
                    
                    sheet.Cells[rowCount, 6].Value = item.OverWorkFixedTime;
                    
                    sheet.Cells[rowCount, 7].Value = item.OverWorkFixedTime;
                    
                    sheet.Cells[rowCount, 8].Value = item.OverWorkFixedTime;
                    
                    sheet.Cells[rowCount, 9].Value = item.OverWorkFixedTime;
                    
                    sheet.Cells[rowCount, 10].Value = item.OverWorkFixedTime;
                    
                    sheet.Cells[rowCount, 11].Value = item.OverWorkFixedTime;
                    
                    sheet.Cells[rowCount, 12].Value = item.OverWorkFixedTime;

                    sheet.Cells[rowCount, 13].Value = item.OverWorkFixedTime;

                    sheet.Cells[rowCount, 14].Value = item.OverWorkFixedTime;

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
                    
                    if (count%2==0)
                    {
                        // even
                        sheet.Cells[rowCount, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCount, 1].Style.Fill.BackgroundColor.SetColor(1, 218, 238, 243);
                    }
                    else
                    {
                        // odd
                        sheet.Cells[rowCount, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCount, 1].Style.Fill.BackgroundColor.SetColor(1, 252,213,180);
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
                    foreach (var salaryType in _unitPriceTypeBLL.GetAllUnitPriceTypes())
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

                            sheet.Cells[rowCount, 6].Value =0.ToString("N0");
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


                            sheet.Cells[rowCount, 12].Value =0.ToString("N0");
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


                            sheet.Cells[rowCount, 14].Value =0.ToString("N0");
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
                            double beginningValue = _salaryBLL.GetGradeSalaryType(departmentId, salaryType.Id, 2022, item.Grade.Id).GradeLowPoints;
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
                            sheet.Cells[rowCount, 5].Value = (manpointDec* beginningValue* commonValue).ToString("N0");
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
                            sheet.Cells[rowCount, 6].Value = (manpointJan* beginningValue* commonValue).ToString("N0");
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
                            sheet.Cells[rowCount, 7].Value = (manpointFeb* beginningValue* commonValue).ToString("N0");
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
                            sheet.Cells[rowCount, 8].Value = (manpointMar* beginningValue* commonValue).ToString("N0");
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
                            sheet.Cells[rowCount, 9].Value = (manpointApr* beginningValue* commonValue).ToString("N0");
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
                            sheet.Cells[rowCount, 10].Value = (manpointMay*beginningValue*commonValue).ToString("N0");
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
                            sheet.Cells[rowCount, 11].Value = (manpointJun*beginningValue*commonValue).ToString("N0");
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
                            sheet.Cells[rowCount, 12].Value = (manpointJul* beginningValue* commonValue).ToString("N0");
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
                            sheet.Cells[rowCount, 13].Value = (manpointAug* beginningValue*commonValue).ToString("N0");
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
                            sheet.Cells[rowCount, 14].Value = (manpointSep*beginningValue*commonValue).ToString("N0");
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
                            double beginningValue = _salaryBLL.GetGradeSalaryType(departmentId, salaryType.Id, 2022, item.Grade.Id).GradeLowPoints;
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
                        // salary allowance (total)
                        if (salaryTypeCount == 5)
                        {
                            sheet.Cells[rowCount, 3].Value = (totalRegularOct+totalFixedOct+totalOverOct).ToString("N0");
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
                            double beginningValue = _salaryBLL.GetGradeSalaryType(departmentId, salaryType.Id, 2022, item.Grade.Id).GradeLowPoints;
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
                            double beginningValue = _salaryBLL.GetGradeSalaryType(departmentId, salaryType.Id, 2022, item.Grade.Id).GradeLowPoints;
                            

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
                            double beginningValue = _salaryBLL.GetGradeSalaryType(departmentId, salaryType.Id, 2022, item.Grade.Id).GradeLowPoints;
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
                            sheet.Cells[rowCount, 3].Value = (wExpensesOct+wExpBonusOct).ToString("N0");
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
                            expensesTotalOct += totalSalaryOct + mWagesOct + dispatchFeeOct + employeeBonusOct + commutingExpensesOct + totalSalaryOct;
                            sheet.Cells[rowCount, 3].Value = (totalSalaryOct + mWagesOct + dispatchFeeOct + employeeBonusOct + commutingExpensesOct + totalSalaryOct).ToString("N0");
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

                            expensesTotalNov += totalSalaryNov + mWagesNov + dispatchFeeNov + employeeBonusNov + commutingExpensesNov + totalSalaryNov;
                            sheet.Cells[rowCount, 4].Value = (totalSalaryNov + mWagesNov + dispatchFeeNov + employeeBonusNov + commutingExpensesNov + totalSalaryNov).ToString("N0");
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
                            expensesTotalDec += totalSalaryDec + mWagesDec + dispatchFeeDec + employeeBonusDec + commutingExpensesDec + totalSalaryDec;
                            sheet.Cells[rowCount, 5].Value = (totalSalaryDec + mWagesDec + dispatchFeeDec + employeeBonusDec + commutingExpensesDec + totalSalaryDec).ToString("N0");
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
                            expensesTotalJan += totalSalaryJan + mWagesJan + dispatchFeeJan + employeeBonusJan + commutingExpensesJan + totalSalaryJan;
                            sheet.Cells[rowCount, 6].Value = (totalSalaryJan + mWagesJan + dispatchFeeJan + employeeBonusJan + commutingExpensesJan + totalSalaryJan).ToString("N0");
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

                            expensesTotalFeb += totalSalaryFeb + mWagesFeb + dispatchFeeFeb + employeeBonusFeb + commutingExpensesFeb + totalSalaryFeb;
                            sheet.Cells[rowCount, 7].Value = (totalSalaryFeb + mWagesFeb + dispatchFeeFeb + employeeBonusFeb + commutingExpensesFeb + totalSalaryFeb).ToString("N0");
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
                            expensesTotalMar += totalSalaryMar + mWagesMar + dispatchFeeMar + employeeBonusMar + commutingExpensesMar + totalSalaryMar;
                            sheet.Cells[rowCount, 8].Value = (totalSalaryMar + mWagesMar + dispatchFeeMar + employeeBonusMar + commutingExpensesMar + totalSalaryMar).ToString("N0");
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
                            expensesTotalApr += totalSalaryApr + mWagesApr + dispatchFeeApr + employeeBonusApr + commutingExpensesApr + totalSalaryApr;
                            sheet.Cells[rowCount, 9].Value = (totalSalaryApr + mWagesApr + dispatchFeeApr + employeeBonusApr + commutingExpensesApr + totalSalaryApr).ToString("N0");
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
                            expensesTotalMay += totalSalaryMay + mWagesMay + dispatchFeeMay + employeeBonusMay + commutingExpensesMay + totalSalaryMay;
                            sheet.Cells[rowCount, 10].Value = (totalSalaryMay + mWagesMay + dispatchFeeMay + employeeBonusMay + commutingExpensesMay + totalSalaryMay).ToString("N0");
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
                            expensesTotalJun += totalSalaryJun + mWagesJun + dispatchFeeJun + employeeBonusJun + commutingExpensesJun + totalSalaryJun;
                            sheet.Cells[rowCount, 11].Value = (totalSalaryJun + mWagesJun + dispatchFeeJun + employeeBonusJun + commutingExpensesJun + totalSalaryJun).ToString("N0");
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

                            expensesTotalJul += totalSalaryJul + mWagesJul + dispatchFeeJul + employeeBonusJul + commutingExpensesJul + totalSalaryJul;
                            sheet.Cells[rowCount, 12].Value = (totalSalaryJul + mWagesJul + dispatchFeeJul + employeeBonusJul + commutingExpensesJul + totalSalaryJul).ToString("N0");
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
                            expensesTotalAug += totalSalaryAug + mWagesAug + dispatchFeeAug + employeeBonusAug + commutingExpensesAug + totalSalaryAug;
                            sheet.Cells[rowCount, 13].Value = (totalSalaryAug + mWagesAug + dispatchFeeAug + employeeBonusAug + commutingExpensesAug + totalSalaryAug).ToString("N0");
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

                            expensesTotalSep += totalSalarySep + mWagesSep + dispatchFeeSep + employeeBonusSep + commutingExpensesSep + totalSalarySep;
                            sheet.Cells[rowCount, 14].Value = (totalSalarySep + mWagesSep + dispatchFeeSep + employeeBonusSep + commutingExpensesSep + totalSalarySep).ToString("N0");
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
                foreach (var item in _unitPriceTypeBLL.GetAllUnitPriceTypes())
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
                        sheet.Cells[rowCountForFuture, 3].Value = regularSalaryTotalOct;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = regularSalaryTotalNov;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = regularSalaryTotalDec;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = regularSalaryTotalJan;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = regularSalaryTotalFeb;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = regularSalaryTotalMar;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = regularSalaryTotalApr;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = regularSalaryTotalMay;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = regularSalaryTotalJun;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = regularSalaryTotalJul;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = regularSalaryTotalAug;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = regularSalaryTotalSep;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // fixed
                    if (salaryTypeSummeryCount == 3)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = fixedSalaryTotalOct;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = fixedSalaryTotalNov;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = fixedSalaryTotalDec;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = fixedSalaryTotalJan;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = fixedSalaryTotalFeb;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = fixedSalaryTotalMar;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = fixedSalaryTotalApr;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = fixedSalaryTotalMay;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = fixedSalaryTotalJun;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = fixedSalaryTotalJul;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = fixedSalaryTotalAug;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = fixedSalaryTotalSep;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // overtime
                    if (salaryTypeSummeryCount == 4)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = overTimeSalaryTotalOct;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = overTimeSalaryTotalNov;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = overTimeSalaryTotalDec;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = overTimeSalaryTotalJan;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = overTimeSalaryTotalFeb;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = overTimeSalaryTotalMar;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = overTimeSalaryTotalApr;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = overTimeSalaryTotalMay;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = overTimeSalaryTotalJun;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = overTimeSalaryTotalJul;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = overTimeSalaryTotalAug;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = overTimeSalaryTotalSep;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // total salary
                    if (salaryTypeSummeryCount == 5)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = salaryTotalOct;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = salaryTotalNov;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = salaryTotalDec;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = salaryTotalJan;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = salaryTotalFeb;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = salaryTotalMar;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = salaryTotalApr;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = salaryTotalMay;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = salaryTotalJun;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = salaryTotalJul;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = salaryTotalAug;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = salaryTotalSep;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // miscellaneous wages
                    if (salaryTypeSummeryCount == 6)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = mwTotalOct;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = mwTotalNov;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = mwTotalDec;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = mwTotalJan;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = mwTotalFeb;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = mwTotalMar;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = mwTotalApr;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = mwTotalMay;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = mwTotalJun;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = mwTotalJul;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = mwTotalAug;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = mwTotalSep;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // dispatch fee
                    if (salaryTypeSummeryCount == 7)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = dispatchFeeTotalOct;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = dispatchFeeTotalNov;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = dispatchFeeTotalDec;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = dispatchFeeTotalJan;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = dispatchFeeTotalFeb;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = dispatchFeeTotalMar;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = dispatchFeeTotalApr;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = dispatchFeeTotalMay;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = dispatchFeeTotalJun;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = dispatchFeeTotalJul;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = dispatchFeeTotalAug;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = dispatchFeeTotalSep;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // provision for employee bonus
                    if (salaryTypeSummeryCount == 8)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = employeeBonusTotalOct;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = employeeBonusTotalNov;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = employeeBonusTotalDec;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = employeeBonusTotalJan;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = employeeBonusTotalFeb;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = employeeBonusTotalMar;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = employeeBonusTotalApr;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = employeeBonusTotalMay;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = employeeBonusTotalJun;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = employeeBonusTotalJul;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = employeeBonusTotalAug;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = employeeBonusTotalSep;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // commuting expenses
                    if (salaryTypeSummeryCount == 9)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = commutingExpensesTotalOct;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = commutingExpensesTotalNov;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = commutingExpensesTotalDec;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = commutingExpensesTotalJan;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = commutingExpensesTotalFeb;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = commutingExpensesTotalMar;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = commutingExpensesTotalApr;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = commutingExpensesTotalMay;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = commutingExpensesTotalJun;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = commutingExpensesTotalJul;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = commutingExpensesTotalAug;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = commutingExpensesTotalSep;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // salary statutory welfare expenses
                    if (salaryTypeSummeryCount == 10)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = welfareExpensesTotalOct;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = welfareExpensesTotalNov;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = welfareExpensesTotalDec;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = welfareExpensesTotalJan;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = welfareExpensesTotalFeb;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = welfareExpensesTotalMar;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = welfareExpensesTotalApr;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = welfareExpensesTotalMay;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = welfareExpensesTotalJun;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = welfareExpensesTotalJul;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = welfareExpensesTotalAug;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = welfareExpensesTotalSep;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // provision for statotory welfare expenses for bonus
                    if (salaryTypeSummeryCount == 11)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = welfareExpensesBonusTotalOct;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = welfareExpensesBonusTotalNov;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = welfareExpensesBonusTotalDec;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = welfareExpensesBonusTotalJan;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = welfareExpensesBonusTotalFeb;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = welfareExpensesBonusTotalMar;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = welfareExpensesBonusTotalApr;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = welfareExpensesBonusTotalMay;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = welfareExpensesBonusTotalJun;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = welfareExpensesBonusTotalJul;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = welfareExpensesBonusTotalAug;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = welfareExpensesBonusTotalSep;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // total statutory benefites
                    if (salaryTypeSummeryCount == 12)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = statutoryTotalOct;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = statutoryTotalNov;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = statutoryTotalDec;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = statutoryTotalJan;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = statutoryTotalFeb;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = statutoryTotalMar;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = statutoryTotalApr;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = statutoryTotalMay;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = statutoryTotalJun;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = statutoryTotalJul;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = statutoryTotalAug;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = statutoryTotalSep;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    }
                    // total expenses
                    if (salaryTypeSummeryCount == 13)
                    {
                        sheet.Cells[rowCountForFuture, 3].Value = expensesTotalOct;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 4].Value = expensesTotalNov;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 5].Value = expensesTotalDec;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 6].Value = expensesTotalJan;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 7].Value = expensesTotalFeb;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 8].Value = expensesTotalMar;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 9].Value = expensesTotalApr;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 10].Value = expensesTotalMay;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 11].Value = expensesTotalJun;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 12].Value = expensesTotalJul;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 13].Value = expensesTotalAug;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowCountForFuture, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                        sheet.Cells[rowCountForFuture, 14].Value = expensesTotalSep;
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
                        
                        tempList = customizedList.Where(l=>l.SectionId==item.SectionId && l.CompanyId==item.CompanyId).ToList();
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
                List<Department> departments = _departmentBLL.GetAllDepartments().OrderBy(dep=>dep.Id).ToList();
                List<SalaryMasterExportDto> salaryMasterExportDtos = new List<SalaryMasterExportDto>();
                var salaryTypeIds = _salaryBLL.GetSalaryTypeIdByYear(2022);
                foreach (var salaryTypeId in salaryTypeIds)
                {
                    foreach (var grade in grades)
                    {
                         salaryMasterExportDtos.Add(_salaryBLL.GetSalaryTypeWithGradeSalaryByYear(2022, grade.Id, salaryTypeId));
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
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow+1, rowCountSalaryMaster, extendedRow+2].Value = item.DepartmentName;
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Merge = true;
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.Font.Bold = true;
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.Font.Color.SetColor(Color.White);
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow + 1, rowCountSalaryMaster, extendedRow + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheetSalaryMaster.Cells[rowCountSalaryMaster, extendedRow+1, rowCountSalaryMaster, extendedRow+2].Style.Fill.BackgroundColor.SetColor(1, 36, 64, 98);
                    extendedRow += 2;
                }

                rowCountSalaryMaster++;
                int prevSalaryTypeId = 0;

                foreach (var item in salaryMasterExportDtos)
                {
                    if (item.GradeSalaryTypes.Count > 1)
                    {
                        if (prevSalaryTypeId != item.SalaryType.Id)
                        {
                            sheetSalaryMaster.Cells[rowCountSalaryMaster, 4].Value = item.SalaryType.SalaryTypeName;
                            prevSalaryTypeId = item.SalaryType.Id;
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

                #endregion

                #region common master

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


                if (commonMasters.Count>0)
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

        public double GetTotalManPoints(List<SalaryAssignmentDto> salaryAssignmentDtos,int monthId,int gradeId)
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
        IDictionary<string,string> GetCommonMaster()
        {
            Dictionary<string, string> values = new Dictionary<string, string>();
            values.Add("取締役/監査役","");
            values.Add("G10", "");
            values.Add("G9", "");
            values.Add("G8", "");
            values.Add("G7", "");
            values.Add("G6", "");
            values.Add("G5", "45");
            values.Add("G4", "45");
            values.Add("G3", "45");
            values.Add("G2", "45");
            values.Add("G1", "45");
            values.Add("ハンディキャップ", "");

            return values;

        }

        List<string> GetAllGradeEntities()
        {
            return new List<string>() {
                "役員報酬",
                "給与手当（定例給与分",
                "給与手当（固定時間分",
                "給与手当(残業手当分",
                "給与手当合計",
                "雑給",
                "派遣費",
                "従業員賞与引当金",
                "給与法定福利費",
                "賞与法定福利費引当",
                "法定福利費合計",
                "通勤費",
                "経費合計"
            };
        }

        Dictionary<string,List<string>> GetOtherGrades()
        {
            Dictionary<string, List<string>> getOtherGradeWithEntities = new Dictionary<string, List<string>>();
            getOtherGradeWithEntities.Add("G10（役員)", GetAllGradeEntities());
            getOtherGradeWithEntities.Add("G9", GetAllGradeEntities());
            getOtherGradeWithEntities.Add("G8", GetAllGradeEntities());
            getOtherGradeWithEntities.Add("G7", GetAllGradeEntities());
            getOtherGradeWithEntities.Add("G6", GetAllGradeEntities());
            getOtherGradeWithEntities.Add("G5",new List<string>() { "役員報酬" });
            getOtherGradeWithEntities.Add("G4", GetAllGradeEntities());
            getOtherGradeWithEntities.Add("G3", GetAllGradeEntities());
            getOtherGradeWithEntities.Add("G2", GetAllGradeEntities());
            getOtherGradeWithEntities.Add("G1", GetAllGradeEntities());
            getOtherGradeWithEntities.Add("ハンディキャップ", GetAllGradeEntities());
            getOtherGradeWithEntities.Add("クルー", GetAllGradeEntities());
            getOtherGradeWithEntities.Add("派遣社員（管理)", GetAllGradeEntities());

            return getOtherGradeWithEntities;
        }
    }
}