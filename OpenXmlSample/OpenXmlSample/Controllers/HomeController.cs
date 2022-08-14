using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OpenXmlSample.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlSample.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }
 
         
        public async Task<IActionResult> ExportTemplates()
        {  
            var path = Path.Combine(Directory.GetCurrentDirectory(), "templates", "templates.docx");

            byte[] byteArray = System.IO.File.ReadAllBytes(path);

            string exportName = $"Export_{Guid.NewGuid()}.docx";
            string newPath = Path.Combine(Directory.GetCurrentDirectory(), "Exports", exportName);

            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length); 
                using (WordprocessingDocument doc = WordprocessingDocument.Open(stream, true))
                {
                    doc.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
                    doc.InsertText("Test", "MahmoudAhmadAbujahel");
                }
                using (FileStream fs = new FileStream(newPath, FileMode.Create))
                { 
                    stream.WriteTo(fs); 
                }
            }
              
            return File(System.IO.File.ReadAllBytes(newPath), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"); 

        }


    }
}
