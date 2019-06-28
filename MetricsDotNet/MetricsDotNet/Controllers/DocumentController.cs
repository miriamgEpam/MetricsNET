using System.Collections.Generic;
using System.Threading.Tasks;
using MetricsDotNet.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.FileProviders;
using MetricsDotNet.ExtensionMethods;
using System.IO;
using System;
using OfficeOpenXml;
using System.Text;
using MetricsDotNet.ViewModels;

namespace MetricsDotNet.Controllers
{
    public class DocumentController : Controller
    {
        private readonly IDocumentService _documentService;
        private readonly IFileProvider _fileProvider;
        public DocumentController(IDocumentService documentService, IFileProvider fileProvider)
        {
            _documentService = documentService;
            this._fileProvider = fileProvider;
        }
        public IActionResult Index()
        {
            return View();
        }
        [HttpPost("Document")]
        public async Task<IActionResult> Index(IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return Content("file not selected");
                }
                var path = Path.Combine(
                            Directory.GetCurrentDirectory(), "wwwroot",
                            file.GetFilename());
                using (var stream = new FileStream(path, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                TempData["excelPath"] = path;
                return RedirectToAction("ProcessDocument");
            }
            catch (Exception e)
            {

                throw new Exception(e.Message);
            }
        }

        public IActionResult ProcessDocument()
        {
            try
            {

                string excelPath = null;
                if (TempData.ContainsKey("excelPath"))
                {
                    excelPath = TempData["excelPath"].ToString();
                }
                var processedFile = _documentService.UploadExcelDocument(excelPath);

                return View(processedFile);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        [HttpPost("SaveInfo")]
        public IActionResult SaveInfo(FieldsCapturedViewModel fieldsCapturedVM)
        {
            try
            {
                _documentService.SendInfoToService(fieldsCapturedVM);
                return View();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }


        }

    }


}