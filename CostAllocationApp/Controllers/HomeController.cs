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

                            tempAssignmentId = EmployeeAssignment();
                            EmployeeAssignmentViewModel employeeAssignment = employeeAssignmentBLL.GetAssignmentById(tempAssignmentId);
                            
                            tempRow += $@"10_{dt_.Rows[i][1].ToString()}_{dt_.Rows[i][13].ToString()},11_{dt_.Rows[i][2].ToString()}_{dt_.Rows[i][14].ToString()},12_{dt_.Rows[i][3].ToString()}_{dt_.Rows[i][15].ToString()},1_{dt_.Rows[i][4].ToString()}_{dt_.Rows[i][16].ToString()},2_{dt_.Rows[i][5].ToString()}_{dt_.Rows[i][17].ToString()},3_{dt_.Rows[i][6].ToString()}_{dt_.Rows[i][18].ToString()},4_{dt_.Rows[i][7].ToString()}_{dt_.Rows[i][19].ToString()},5_{dt_.Rows[i][8].ToString()}_{dt_.Rows[i][20].ToString()},6_{dt_.Rows[i][9].ToString()}_{dt_.Rows[i][21].ToString()},7_{dt_.Rows[i][10].ToString()}_{dt_.Rows[i][22].ToString()},8_{dt_.Rows[i][11].ToString()}_{dt_.Rows[i][23].ToString()},9_{dt_.Rows[i][12].ToString()}_{dt_.Rows[i][24].ToString()}";

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
    }
}