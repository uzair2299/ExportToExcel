using ClosedXML.Excel;
using FastMember;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebApplication1
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            List<Student> students = new List<Student>
            {
                new Student(){Id=1,Name="Uzair",RollNumber=235},
                new Student(){Id=1,Name="Uzair 11",RollNumber=236},
                new Student(){Id=1,Name="Uzair 22",RollNumber=237},
                new Student(){Id=1,Name="Uzair 33",RollNumber=238},
                new Student(){Id=1,Name="Uzair 55",RollNumber=239},
                new Student(){Id=1,Name="Uzair 66",RollNumber=240},
                new Student(){Id=1,Name="Uzair 77",RollNumber=241},
                new Student(){Id=1,Name="Uzair 88",RollNumber=242},
                new Student(){Id=1,Name="Uzair 99",RollNumber=235},
            };

            DataTable dataTable = new DataTable();
            using (var reader = ObjectReader.Create(students))
            {
                dataTable.Load(reader);
            }
            using(XLWorkbook wb = new XLWorkbook())
            {
               var ws = wb.Worksheets.Add(dataTable, "RecordsClosedXMl");
                ws.Tables.FirstOrDefault().ShowAutoFilter = false;
                ws.Tables.FirstOrDefault().Theme = XLTableTheme.None;
                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                string excelName = "Record";
                Response.ContentType= "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment; filename=" + excelName + ".xlsx");
                using (var memoryStream = new MemoryStream())
                {
                    wb.SaveAs(memoryStream);
                    memoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }
            //ExcelPackage excel = new ExcelPackage();
            //var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
            //workSheet.TabColor = System.Drawing.Color.Black;
            //workSheet.DefaultRowHeight = 12;

            //Header of table  
            //  
            //workSheet.Row(1).Height = 20;
            //workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //workSheet.Row(1).Style.Font.Bold = true;
            //workSheet.Cells[1, 1].Value = "ID";
            //workSheet.Cells[1, 2].Value = "Name";
            //workSheet.Cells[1, 3].Value = "Roll Number";
            //body of table
            //int recordIndex = 2;
            //foreach (var student in students)
            //{
            //    workSheet.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
            //    workSheet.Cells[recordIndex, 2].Value = student.Id;
            //    workSheet.Cells[recordIndex, 3].Value = student.Name;
            //    recordIndex++;
            //}


            //workSheet.Column(1).AutoFit();
            //workSheet.Column(2).AutoFit();
            //workSheet.Column(3).AutoFit();
            //string excelName = "Record";
            //using (var memoryStream = new MemoryStream())
            //{
            //    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //    Response.AddHeader("content-disposition", "attachment; filename=" + excelName + ".xlsx");
            //    excel.SaveAs(memoryStream);
            //    memoryStream.WriteTo(Response.OutputStream);
            //    Response.Flush();
            //    Response.End();
            //   }
        }

        public class Student
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public int RollNumber { get; set; }
        }
    }
}