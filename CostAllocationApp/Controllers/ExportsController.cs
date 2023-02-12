﻿using System;
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
            var _department = _departmentBLL.GetDepartmentByDepartemntId(departmentId);

            using (var client = new HttpClient())
            {
                //string uri = "http://198.38.92.119:8081/api/utilities/SearchForecastEmployee?employeeName=&sectionId=&departmentId=" + departmentId + "&inchargeId=&roleId=&explanationId=&companyId=&status=&year=";
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

        [HttpPost]
        public ActionResult DataExportByAllocation(int departmentId = 0, int explanationId = 0)
        {
            List<SalaryAssignmentDto> salaryAssignmentDtos = new List<SalaryAssignmentDto>();
            var department = _departmentBLL.GetDepartmentByDepartemntId(departmentId);
            var explanation = _explanationsBLL.GetExplanationByExplanationId(explanationId);
            var salaries = _salaryBLL.GetAllSalaryPoints();

            // get all data with department and allocation
            List<ForecastAssignmentViewModel> assignmentsWithForecast = _exportBLL.AssignmentsByAllocation(departmentId, explanationId);
            //filtered by grade id
            List<ForecastAssignmentViewModel> assignmentsWithGrade = assignmentsWithForecast.Where(a=>a.GradeId!=null).ToList();
            //filtered by section and company
            List<ForecastAssignmentViewModel> assignmentsWithSectionAndCompany = assignmentsWithForecast.Where(a=>a.SectionId!=null && a.CompanyId!=null).ToList();

            foreach (var item in salaries)
            {
                List<ForecastAssignmentViewModel> filteredAssignmentsBySalaryId = assignmentsWithGrade.Where(a=>a.GradeId==item.Id.ToString()).ToList();
                salaryAssignmentDtos.Add(new SalaryAssignmentDto { Salary = item,ForecastAssignmentViewModels = filteredAssignmentsBySalaryId });
            }

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

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
                foreach (var item in GetCommonMaster())
                {
                    sheet.Cells[rowCount, 2].Value = item.Key;
                    sheet.Cells[rowCount, 3].Value = item.Value;
                 
                    sheet.Cells[rowCount, 4].Value = item.Value;
                    
                    sheet.Cells[rowCount, 5].Value = item.Value;
                    
                    sheet.Cells[rowCount, 6].Value = item.Value;
                    
                    sheet.Cells[rowCount, 7].Value = item.Value;
                    
                    sheet.Cells[rowCount, 8].Value = item.Value;
                    
                    sheet.Cells[rowCount, 9].Value = item.Value;
                    
                    sheet.Cells[rowCount, 10].Value = item.Value;
                    
                    sheet.Cells[rowCount, 11].Value = item.Value;
                    
                    sheet.Cells[rowCount, 12].Value = item.Value;

                    sheet.Cells[rowCount, 13].Value = item.Value;

                    sheet.Cells[rowCount, 14].Value = item.Value;

                    rowCount++;
                }
                #endregion
                rowCount++;
                #region grade entities
                foreach (var item in GetAllGradeEntities())
                {
                    sheet.Cells[rowCount, 1].Value = item;
                    sheet.Cells[rowCount, 3].Value = 0;
                    sheet.Cells[rowCount, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 3].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCount, 4].Value = 0;
                    sheet.Cells[rowCount, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 4].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCount, 5].Value = 0;
                    sheet.Cells[rowCount, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 5].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCount, 6].Value = 0;
                    sheet.Cells[rowCount, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 6].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCount, 7].Value = 0;
                    sheet.Cells[rowCount, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 7].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCount, 8].Value = 0;
                    sheet.Cells[rowCount, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 8].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCount, 9].Value = 0;
                    sheet.Cells[rowCount, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 9].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCount, 10].Value = 0;
                    sheet.Cells[rowCount, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 10].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCount, 11].Value = 0;
                    sheet.Cells[rowCount, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 11].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCount, 12].Value = 0;
                    sheet.Cells[rowCount, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 12].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCount, 13].Value = 0;
                    sheet.Cells[rowCount, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 13].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    sheet.Cells[rowCount, 14].Value = 0;
                    sheet.Cells[rowCount, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[rowCount, 14].Style.Fill.BackgroundColor.SetColor(1, 216, 228, 188);
                    rowCount++;
                }
                #endregion

                #region other grade with entities
                rowCount++;
                int count = 1;
                foreach (var item in GetOtherGrades())
                {
                    sheet.Cells[rowCount, 1].Value = item.Key;
                    
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

                    int innerCount = 1;
                    foreach (var item1 in item.Value)
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

                        sheet.Cells[rowCount, 2].Value = item1;
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

                        sheet.Cells[rowCount, 3].Value = 0;

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

                        sheet.Cells[rowCount, 4].Value = 0;

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

                        sheet.Cells[rowCount, 5].Value = 0;

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

                        sheet.Cells[rowCount, 6].Value = 0;

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

                        sheet.Cells[rowCount, 7].Value = 0;

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

                        sheet.Cells[rowCount, 8].Value = 0;

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

                        sheet.Cells[rowCount, 9].Value = 0;

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

                        sheet.Cells[rowCount, 10].Value = 0;

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

                        sheet.Cells[rowCount, 11].Value = 0;

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

                        sheet.Cells[rowCount, 12].Value = 0;

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

                        sheet.Cells[rowCount, 13].Value = 0;

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

                        sheet.Cells[rowCount, 14].Value = 0;

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
                        rowCount++;
                        innerCount++;
                    }

                    count++;

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


                var excelData = package.GetAsByteArray();
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                var fileName = explanation.ExplanationName + ".xlsx";
                return File(excelData, contentType, fileName);

            }
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