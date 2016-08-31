﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CaptchaMvc.HtmlHelpers;
using MelavaWebsite.Common;
using MelavaWebsite.Models;
using NPOI.HSSF.UserModel;

namespace MelavaWebsite.Controllers
{
    public class RegistrationController : Controller
    {
        private DbPersonDetails db = new DbPersonDetails();
        private Helper helper = new Helper();


        public ActionResult Home()
        {
            return View();

        }

        public ActionResult ProgramDetails()
        {
            return View();

        }

        public ActionResult Venue()
        {
            return View();

        }

        public ActionResult AboutUs()
        {
            return View();

        }

        public ActionResult Login()
        {
            return View();

        }

        // GET: /Registration/

        public ActionResult Index()
        {
            return View(db.Persons.ToList());
        }

        public ActionResult ContactUs()
        {
            return View();
        }

        //
        // GET: /Registration/Details/5

        //public ActionResult Details(int id = 0)
        //{
        //    PersonDetails persondetails = db.Persons.Find(id);
        //    if (persondetails == null)
        //    {
        //        return HttpNotFound();
        //    }
        //    return View(persondetails);
        //}

        //
        // GET: /Registration/Create

        //public ActionResult Create()
        //{
        //    return View();
        //}

        //
        // POST: /Registration/Create

        //[HttpPost]
        //public ActionResult Create(PersonDetails persondetails)
        //{
        //    if (ModelState.IsValid)
        //    {

        //        db.Persons.Add(persondetails);
        //        db.SaveChanges();
        //        return RedirectToAction("Index");
        //    }

        //    return View(persondetails);
        //}

        //
        // GET: /Registration/Edit/5

        //public ActionResult Edit(int id = 0)
        //{
        //    PersonDetails persondetails = db.Persons.Find(id);
        //    if (persondetails == null)
        //    {
        //        return HttpNotFound();
        //    }
        //    return View(persondetails);
        //}

        //
        // POST: /Registration/Edit/5

        //[HttpPost]
        //public ActionResult Edit(PersonDetails persondetails)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        db.Entry(persondetails).State = EntityState.Modified;
        //        db.SaveChanges();
        //        return RedirectToAction("Index");
        //    }
        //    return View(persondetails);
        //}

        //
        // GET: /Registration/Delete/5

        //public ActionResult Delete(int id = 0)
        //{
        //    PersonDetails persondetails = db.Persons.Find(id);
        //    if (persondetails == null)
        //    {
        //        return HttpNotFound();
        //    }
        //    return View(persondetails);
        //}

        //
        // POST: /Registration/Delete/5

        //[HttpPost, ActionName("Delete")]
        //public ActionResult DeleteConfirmed(int id)
        //{
        //    PersonDetails persondetails = db.Persons.Find(id);
        //    db.Persons.Remove(persondetails);
        //    db.SaveChanges();
        //    return RedirectToAction("Index");
        //}

        public ActionResult Register()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Register(PersonDetails persondetails)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    //if (this.IsCaptchaValid("Captcha is not valid"))
                    //{

                    //    return RedirectToAction("ThankYouPage");
                    //}  

                    if (Request.Files.Count > 0)
                    {
                        var file = Request.Files[0];
                        if (file != null && (!string.IsNullOrEmpty(file.FileName)))
                        {
                            file.SaveAs(HttpContext.Server.MapPath("~/CandidatePhotos/") + file.FileName);
                            persondetails.ImagePath = "~/CandidatePhotos/" + file.FileName;
                        }
                    }
                    persondetails.CreatedDate = DateTime.Now;
                    db.Persons.Add(persondetails);
                    db.SaveChanges();
                    ViewData["Success"] = "Success";
                    if (db.Persons != null && db.Persons.Local.Count > 0)
                    {
                        PersonDetails updatedPersondetails = db.Persons.Local[0];
                        helper.SendEmail(updatedPersondetails);
                    }
                }
                catch (Exception ex)
                {
                    return RedirectToAction("Index", "Error");
                }

            }
            return View(persondetails);
        }

        protected override void Dispose(bool disposing)
        {
            db.Dispose();
            base.Dispose(disposing);
        }


        public FileResult ExportList()
        {
            List<ExcelColumConfiguration> lstexcelColumConfiguration = CofigureColumns();
            IEnumerable<PersonDetails> lstEntities = db.Persons.ToList();
            return GenerateExcel<PersonDetails>(lstEntities.OrderBy(m => m.Id), "CandidateList", lstexcelColumConfiguration);
        }

        public FileResult GenerateExcel<T>(IEnumerable<T> lstT, string FileName, List<ExcelColumConfiguration> lstexcelColumConfiguration)
        {
            MemoryStream output = new MemoryStream();
            ExcelManager objExcelManager = new ExcelManager();
            var workbook = new HSSFWorkbook();
            try
            {
                //Write the workbook to a memory stream
                workbook = objExcelManager.GetWorkBookFor<T>(lstT, lstexcelColumConfiguration);
                workbook.Write(output);
            }
            catch (Exception ex)
            {
                //Logger.Error(string.Format("Error occured in Controller:Base Action:GenerateExcel ExceptionMessage:{0}", ex.Message), ex);
                //DisplayMessages(MessageType.EROR, BankBrandingResource.ErrorMessage_Generic);
            }

            return File(output.ToArray(),   //The binary data of the XLS file
                "application/vnd.ms-excel", //MIME type of Excel files
                FileName + ".xls");     //Suggested file name in the "Save as" dialog which will be displayed to the end user
        }

        private List<ExcelColumConfiguration> CofigureColumns()
        {
            List<ExcelColumConfiguration> columnList = new List<ExcelColumConfiguration>();
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 0, DBColumnName = "Id", ColumnHeaderName = "Candidate Id", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 1, DBColumnName = "AnubandhId", ColumnHeaderName = "Anubandh Id", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 2, DBColumnName = "Name", ColumnHeaderName = "Name", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 3, DBColumnName = "BirthName", ColumnHeaderName = "Birth Name", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 4, DBColumnName = "FirstGotra", ColumnHeaderName = "First Gotra", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 5, DBColumnName = "SecondGotra", ColumnHeaderName = "Second Gotra", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 6, DBColumnName = "Gender", ColumnHeaderName = "Gender", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 7, DBColumnName = "Age", ColumnHeaderName = "Age", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 8, DBColumnName = "Height", ColumnHeaderName = "Height", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 9, DBColumnName = "Education", ColumnHeaderName = "Education", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 10, DBColumnName = "BirthPlace", ColumnHeaderName = "Birth Place", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 11, DBColumnName = "Occupation", ColumnHeaderName = "Occupation", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 12, DBColumnName = "ContactNumber", ColumnHeaderName = "Contact Number", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 13, DBColumnName = "Email", ColumnHeaderName = "Email", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 14, DBColumnName = "Address", ColumnHeaderName = "Address", ColumnWidth = 15 * 256 });
            columnList.Add(new ExcelColumConfiguration { ColumnIndex = 15, DBColumnName = "CreatedDate", ColumnHeaderName = "Created Date", ColumnWidth = 15 * 256 });
            return columnList;
        }
    }
}