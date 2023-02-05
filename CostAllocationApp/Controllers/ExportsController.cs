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

namespace CostAllocationApp.Controllers
{
    public class ExportsController : Controller
    {
        private DepartmentBLL _departmentBLL = null;
        private ExportBLL _exportBLL = null;
        private SalaryBLL _salaryBLL = null;
        private ExplanationsBLL _explanationsBLL = null;
        public ExportsController()
        {
            _departmentBLL = new DepartmentBLL();
            _exportBLL = new ExportBLL();
            _salaryBLL = new SalaryBLL();
            _explanationsBLL = new ExplanationsBLL();
        }
        // GET: Exports
        public FileResult ExportBySection(int sectionId=0,string sectionName = "")
        {
            List<ForecastAssignmentViewModel> forecastAssignmentViewModels = new List<ForecastAssignmentViewModel>();
            using (var client = new HttpClient())
            {
                //string uri = "http://198.38.92.119:8081/api/utilities/SearchForecastEmployee?employeeName=&sectionId=" + sectionId + "&departmentId=&inchargeId=&roleId=&explanationId=&companyId=&status=&year=";
                string uri = "http://localhost:59198/api/utilities/SearchForecastEmployee?employeeName=&sectionId=" + sectionId + "&departmentId=&inchargeId=&roleId=&explanationId=&companyId=&status=&year=";
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
            using (var client = new HttpClient())
            {
                //string uri = "http://198.38.92.119:8081/api/utilities/SearchForecastEmployee?employeeName=&sectionId=" + sectionId + "&departmentId=&inchargeId=&roleId=&explanationId=&companyId=&status=&year=";
                string uri = "http://localhost:59198/api/utilities/SearchForecastEmployee?employeeName=&sectionId=&departmentId="+departmentId+"&inchargeId=&roleId=&explanationId=&companyId=&status=&year=";
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
                        sheet.Cells["F" + count].Value = Convert.ToDecimal(item.UnitPrice);

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
                                sheet.Cells["S" + count].Value = Convert.ToDecimal(item.forecasts[i].Total);
                                sheet.Cells["T" + count].Value = Convert.ToDecimal(item.forecasts[i].Total);
                                sheet.Cells["U" + count].Value = Convert.ToDecimal(item.forecasts[i].Total);
                                sheet.Cells["V" + count].Value = Convert.ToDecimal(item.forecasts[i].Total);
                                sheet.Cells["W" + count].Value = Convert.ToDecimal(item.forecasts[i].Total);
                                sheet.Cells["X" + count].Value = Convert.ToDecimal(item.forecasts[i].Total);
                                sheet.Cells["Y" + count].Value = Convert.ToDecimal(item.forecasts[i].Total);
                                sheet.Cells["Z" + count].Value = Convert.ToDecimal(item.forecasts[i].Total);
                                sheet.Cells["AA" + count].Value = Convert.ToDecimal(item.forecasts[i].Total);
                                sheet.Cells["AB" + count].Value = Convert.ToDecimal(item.forecasts[i].Total);
                                sheet.Cells["AC" + count].Value = Convert.ToDecimal(item.forecasts[i].Total);
                                sheet.Cells["AD" + count].Value = Convert.ToDecimal(item.forecasts[i].Total);
                            }
                        }
                        count++;
                    } 
                }


                var excelData = package.GetAsByteArray();
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                var fileName = "MyWorkbook.xlsx";
                return File(excelData, contentType, fileName);
            }
        }

        public ActionResult DataExportByAllocation()
        {
            return View(new ExportViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }

        [HttpPost]
        public ActionResult DataExportByAllocation(int departmentId = 0, int explanationId = 0)
        {
            List<SalaryAssignmentDto> salaryAssignmentDtos = new List<SalaryAssignmentDto>();
            var explanation = _explanationsBLL.GetExplanationByExplanationId(explanationId);
            var salaries = _salaryBLL.GetAllSalaryPoints();

            // get all data with department and allocation
            List<ForecastAssignmentViewModel> assignmentsWithForecast = _exportBLL.AssignmentsByAllocation(departmentId, explanationId);
            //filtered by grade id
            List<ForecastAssignmentViewModel> assignmentsWithGrade = assignmentsWithForecast.Where(a=>a.GradeId!=null).ToList();
            //filtered by section and company
            List<ForecastAssignmentViewModel> assignmentsWithSectionAndCompany = assignmentsWithForecast.Where(a=>a.SectionId!=null && a.CompanyId!=null).ToList();
            //var aa = (from a in assignmentsWithSectionAndCompany
            //          group a by a.SectionId).ToList();  

            foreach (var item in salaries)
            {
                List<ForecastAssignmentViewModel> filteredAssignmentsBySalaryId = assignmentsWithGrade.Where(a=>a.GradeId==item.Id.ToString()).ToList();
                salaryAssignmentDtos.Add(new SalaryAssignmentDto { Salary = item,ForecastAssignmentViewModels = filteredAssignmentsBySalaryId });
            }

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells[1, 1].Value = "Number of People";
                sheet.Cells[1, 1].Style.Font.Bold = true;
                sheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells[1, 2].Value = "Grade";
                sheet.Cells[1, 2].Style.Font.Bold = true;
                sheet.Cells[1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells[1, 3].Value = "10";
                sheet.Cells[1, 3].Style.Font.Bold = true;
                sheet.Cells[1, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells[1, 4].Value = "11";
                sheet.Cells[1, 4].Style.Font.Bold = true;
                sheet.Cells[1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells[1, 5].Value = "12";
                sheet.Cells[1, 5].Style.Font.Bold = true;
                sheet.Cells[1, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 5].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells[1, 6].Value = "1";
                sheet.Cells[1, 6].Style.Font.Bold = true;
                sheet.Cells[1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells[1, 7].Value = "2";
                sheet.Cells[1, 7].Style.Font.Bold = true;
                sheet.Cells[1, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 7].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells[1, 8].Value = "3";
                sheet.Cells[1, 8].Style.Font.Bold = true;
                sheet.Cells[1, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 8].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells[1, 9].Value = "4";
                sheet.Cells[1, 9].Style.Font.Bold = true;
                sheet.Cells[1, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 9].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


                sheet.Cells[1, 10].Value = "5";
                sheet.Cells[1, 10].Style.Font.Bold = true;
                sheet.Cells[1, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 10].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                sheet.Cells[1, 11].Value = "6";
                sheet.Cells[1, 11].Style.Font.Bold = true;
                sheet.Cells[1, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 11].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


                sheet.Cells[1, 12].Value = "7";
                sheet.Cells[1, 12].Style.Font.Bold = true;
                sheet.Cells[1, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 12].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


                sheet.Cells[1, 13].Value = "8";
                sheet.Cells[1, 13].Style.Font.Bold = true;
                sheet.Cells[1, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 13].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);


                sheet.Cells[1, 14].Value = "9";
                sheet.Cells[1, 14].Style.Font.Bold = true;
                sheet.Cells[1, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, 14].Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);

                int rowCount = 2;
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

                //man month point
                foreach (var item in salaryAssignmentDtos)
                {
                    sheet.Cells[rowCount, 2].Value = item.Salary.SalaryGrade;

                    

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
                    sheet.Cells[rowCount, 4].Value = nov;
                    sheet.Cells[rowCount, 5].Value = dec;
                    sheet.Cells[rowCount, 6].Value = jan;
                    sheet.Cells[rowCount, 7].Value = feb;
                    sheet.Cells[rowCount, 8].Value = mar;
                    sheet.Cells[rowCount, 9].Value = apr;
                    sheet.Cells[rowCount, 10].Value = may;
                    sheet.Cells[rowCount, 11].Value = jun;
                    sheet.Cells[rowCount, 12].Value = jul;
                    sheet.Cells[rowCount, 13].Value = aug;
                    sheet.Cells[rowCount, 14].Value = sep;

                    oct = 0; nov = 0; dec = 0; jan = 0;feb = 0;mar = 0;apr = 0;may = 0;jun = 0;jul = 0;aug = 0;sep = 0;
                    rowCount++;
                }

                rowCount = rowCount + 2;

                sheet.Cells[rowCount, 1, rowCount, 14].Style.Font.Bold = true;
                sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                sheet.Cells[rowCount, 1].Value = "Costing";
                rowCount = rowCount + 3;

                // costing
                foreach (var item in salaryAssignmentDtos)
                {
                    sheet.Cells[rowCount, 2].Value = item.Salary.SalaryGrade;



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
                    sheet.Cells[rowCount, 3].Value = oct;
                    sheet.Cells[rowCount, 4].Value = nov;
                    sheet.Cells[rowCount, 5].Value = dec;
                    sheet.Cells[rowCount, 6].Value = jan;
                    sheet.Cells[rowCount, 7].Value = feb;
                    sheet.Cells[rowCount, 8].Value = mar;
                    sheet.Cells[rowCount, 9].Value = apr;
                    sheet.Cells[rowCount, 10].Value = may;
                    sheet.Cells[rowCount, 11].Value = jun;
                    sheet.Cells[rowCount, 12].Value = jul;
                    sheet.Cells[rowCount, 13].Value = aug;
                    sheet.Cells[rowCount, 14].Value = sep;

                    oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                    rowCount++;
                }


                rowCount = rowCount + 2;




                if (assignmentsWithSectionAndCompany.Count > 0)
                {

                    sheet.Cells[rowCount, 1, rowCount, 14].Style.Font.Bold = true;
                    sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                    sheet.Cells[rowCount, 2].Value = "工数";
                    rowCount++;
                    foreach (var item in assignmentsWithSectionAndCompany)
                    {
                        sheet.Cells[rowCount, 1].Value = item.SectionName;
                        sheet.Cells[rowCount, 2].Value = item.CompanyName;


                        //foreach (var singleAssignment in item.ForecastAssignmentViewModels)
                        //{
                            oct += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 10).Points);
                            nov += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 11).Points);
                            dec += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 12).Points);
                            jan += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 1).Points);
                            feb += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 2).Points);
                            mar += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 3).Points);
                            apr += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 4).Points);
                            may += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 5).Points);
                            jun += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 6).Points);
                            jul += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 7).Points);
                            aug += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 8).Points);
                            sep += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 9).Points);
                        //}
                        sheet.Cells[rowCount, 3].Value = oct;
                        sheet.Cells[rowCount, 4].Value = nov;
                        sheet.Cells[rowCount, 5].Value = dec;
                        sheet.Cells[rowCount, 6].Value = jan;
                        sheet.Cells[rowCount, 7].Value = feb;
                        sheet.Cells[rowCount, 8].Value = mar;
                        sheet.Cells[rowCount, 9].Value = apr;
                        sheet.Cells[rowCount, 10].Value = may;
                        sheet.Cells[rowCount, 11].Value = jun;
                        sheet.Cells[rowCount, 12].Value = jul;
                        sheet.Cells[rowCount, 13].Value = aug;
                        sheet.Cells[rowCount, 14].Value = sep;

                        oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                        rowCount++;
                    }

                    rowCount = rowCount + 2;

                    sheet.Cells[rowCount, 1, rowCount, 14].Style.Font.Bold = true;
                    sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 1, rowCount, 14].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    sheet.Cells[rowCount, 2].Value = "金額";
                    rowCount++;

                    foreach (var item in assignmentsWithSectionAndCompany)
                    {
                        sheet.Cells[rowCount, 1].Value = item.SectionName;
                        sheet.Cells[rowCount, 2].Value = item.CompanyName;


                        //foreach (var singleAssignment in item.ForecastAssignmentViewModels)
                        //{
                        oct += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 10).Total);
                        nov += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 11).Total);
                        dec += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 12).Total);
                        jan += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 1).Total);
                        feb += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 2).Total);
                        mar += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 3).Total);
                        apr += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 4).Total);
                        may += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 5).Total);
                        jun += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 6).Total);
                        jul += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 7).Total);
                        aug += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 8).Total);
                        sep += Convert.ToDecimal(item.forecasts.SingleOrDefault(f => f.Month == 9).Total);
                        //}
                        sheet.Cells[rowCount, 3].Value = oct;
                        sheet.Cells[rowCount, 4].Value = nov;
                        sheet.Cells[rowCount, 5].Value = dec;
                        sheet.Cells[rowCount, 6].Value = jan;
                        sheet.Cells[rowCount, 7].Value = feb;
                        sheet.Cells[rowCount, 8].Value = mar;
                        sheet.Cells[rowCount, 9].Value = apr;
                        sheet.Cells[rowCount, 10].Value = may;
                        sheet.Cells[rowCount, 11].Value = jun;
                        sheet.Cells[rowCount, 12].Value = jul;
                        sheet.Cells[rowCount, 13].Value = aug;
                        sheet.Cells[rowCount, 14].Value = sep;

                        oct = 0; nov = 0; dec = 0; jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0;
                        rowCount++;
                    }
                }





                var excelData = package.GetAsByteArray();
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                var fileName = explanation.ExplanationName+".xlsx";
                return File(excelData, contentType, fileName);

            }
        }
    }
}