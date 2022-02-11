using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using ExportDatabaseToExcel.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportDatabaseToExcel.Controllers
{
    public class ChartController : Controller
    {
        private ExportDatabaseToExcelDBContext db = new ExportDatabaseToExcelDBContext();

        // GET: Chart
        public ActionResult Index()
        {
            return View(db.Charts.ToList());
        }
        private static ExcelPackage GenerateExcelFile(IEnumerable<ExportDatabaseToExcel.Models.Chart> datasource)
        {

            ExcelPackage pck = new ExcelPackage();

            //Create the worksheet 
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Sheet 1");

            // Sets Headers
            ws.Cells[1, 1].Value = "Id";
            ws.Cells[1, 2].Value = "Description";
            ws.Cells[1, 3].Value = "CodeUd";
            ws.Cells[1, 4].Value = "Type";

            // Inserts Data
            for (int i = 0; i < datasource.Count(); i++)
            {
                ws.Cells[i + 2, 1].Value = datasource.ElementAt(i).Id;
                ws.Cells[i + 2, 2].Value = datasource.ElementAt(i).Description;
                ws.Cells[i + 2, 3].Value = datasource.ElementAt(i).CodeUd;
                ws.Cells[i + 2, 4].Value = datasource.ElementAt(i).Type;
            }

            // Format Header of Table
            using (ExcelRange rng = ws.Cells["A1:C1"])
            {

                rng.Style.Font.Bold = true;
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid; //Set Pattern for the background to Solid 
                rng.Style.Fill.BackgroundColor.SetColor(Color.Gold); //Set color to DarkGray 
                rng.Style.Font.Color.SetColor(Color.Black);
            }
            return pck;
        }

        public FileContentResult Download()
        {

            var fileDownloadName = String.Format("FileName.xlsx");
            const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";


            // Pass your ef data to method
            ExcelPackage package = GenerateExcelFile(db.Charts.ToList());

            var fsr = new FileContentResult(package.GetAsByteArray(), contentType);
            fsr.FileDownloadName = fileDownloadName;

            return fsr;
        }


        // GET: Chart/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Chart chart = db.Charts.Find(id);
            if (chart == null)
            {
                return HttpNotFound();
            }
            return View(chart);
        }

        // GET: Chart/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Chart/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,Description,CodeUd,Type")] Chart chart)
        {
            if (ModelState.IsValid)
            {
                db.Charts.Add(chart);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(chart);
        }

        // GET: Chart/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Chart chart = db.Charts.Find(id);
            if (chart == null)
            {
                return HttpNotFound();
            }
            return View(chart);
        }

        // POST: Chart/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,Description,CodeUd,Type")] Chart chart)
        {
            if (ModelState.IsValid)
            {
                db.Entry(chart).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(chart);
        }

        // GET: Chart/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Chart chart = db.Charts.Find(id);
            if (chart == null)
            {
                return HttpNotFound();
            }
            return View(chart);
        }

        // POST: Chart/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Chart chart = db.Charts.Find(id);
            db.Charts.Remove(chart);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
