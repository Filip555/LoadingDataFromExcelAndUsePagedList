using OfficeOpenXml;
using PagedList;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LoadingExcelAndPagedList.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index(int pageNumber = 1, int pageSize = 5)
        {
            var model = new Example();
            var path = new FileInfo(@"C:\Users\Filip555\Desktop\Zeszyt1.xlsx");
            List<Example> list = new List<Example>();
            using (var package = new ExcelPackage(path))
            {
                var workbook = package.Workbook;
                var current = workbook.Worksheets[1];
                var rowCount = current.Dimension.End.Row;
                var columnCount = current.Dimension.End.Column;


                for (int i = 1; i < rowCount + 1; i++)
                {
                    list.Add(new Example()
                    {
                        ID = int.Parse(current.Cells[i, 1].Value.ToString()),
                        Question = current.Cells[i, 2].Value.ToString(),
                        Answer = current.Cells[i, 3].Value.ToString()
                    });
                }
            }
            PagedList<Example> pagedList = new PagedList<Example>(list, pageNumber, pageSize);
            return View(pagedList);
        }
    }
    public class Example
    {
        public int ID { get; set; }
        public string Question { get; set; }
        public string Answer { get; set; }
    }
}