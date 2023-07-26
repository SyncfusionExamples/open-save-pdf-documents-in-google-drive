using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
using Newtonsoft.Json;
using Syncfusion.EJ2.PdfViewer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace PdfViewerWebService
{
    [Route("[controller]")]
    [ApiController]
    public class PdfViewerController : ControllerBase
    {

        private IWebHostEnvironment _hostingEnvironment;
        //Initialize the memory cache object   
        public IMemoryCache _cache;
        string[] Scopes = { DriveService.Scope.Drive };
        private IConfiguration _configuration;
        public readonly string folderId;
        private readonly string applicationName;
        public readonly string credentialPath;

        public PdfViewerController(IWebHostEnvironment hostingEnvironment, IMemoryCache cache, IConfiguration configuration)
        {
            _hostingEnvironment = hostingEnvironment;
            _cache = cache;
            _configuration = configuration;

            folderId = _configuration.GetValue<string>("FolderId");
            credentialPath = _configuration.GetValue<string>("CredentialPath");
            applicationName = _configuration.GetValue<string>("ApplicationName");
        }

        [HttpPost("Load")]
        [Microsoft.AspNetCore.Cors.EnableCors("MyPolicy")]
        [Route("[controller]/Load")]
        //Post action for Loading the PDF documents   

        public async Task<IActionResult> Load([FromBody] Dictionary<string, string> jsonObject)
        {
            Console.WriteLine("Load called");
            //Initialize the PDF viewer object with memory cache object
            PdfRenderer pdfviewer = new PdfRenderer(_cache);
            MemoryStream stream = new MemoryStream();
            object jsonResult = new object();

            if (jsonObject != null && jsonObject.ContainsKey("document"))
            {
                if (bool.Parse(jsonObject["isFileName"]))
                {
                    string objectName = jsonObject["document"];

                    UserCredential credential;

                    using (var memStream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read))
                    {
                        string credPath = "token.json";
                        credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load(memStream).Secrets,
                            Scopes,
                            "user",
                            CancellationToken.None,
                            new FileDataStore(credPath, true));
                    }

                    // Create Google Drive API service.
                    var service = new DriveService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = applicationName,
                    });

                    //This object is used to construct a request to list files from Google Drive.
                    FilesResource.ListRequest listRequest = service.Files.List();
                    listRequest.Q = "mimeType='application/pdf' and '" + folderId + "' in parents and trashed=false";
                    listRequest.Fields = "files(id, name)";
                    var files = await listRequest.ExecuteAsync();

                    // Process the list of files (you can use 'files' to retrieve the list of files)
                    string fileIdToDownload = null;
                    foreach (var file in files.Files)
                    {
                        string fileId = file.Id;
                        string fileName = file.Name;
                        if (fileName == objectName)
                        {
                            // Save the matching fileId
                            fileIdToDownload = fileId;
                            break;
                        }
                    }
                    string fileIds = fileIdToDownload;
                    var request = service.Files.Get(fileIds);
                    await request.DownloadAsync(stream);
                    stream.Position = 0;
                }
                else
                {
                    byte[] bytes = Convert.FromBase64String(jsonObject["document"]);
                    stream = new MemoryStream(bytes);
                }
            }
            jsonResult = pdfviewer.Load(stream, jsonObject);
            return Content(JsonConvert.SerializeObject(jsonResult));
        }

        [AcceptVerbs("Post")]
        [HttpPost("Bookmarks")]
        [Microsoft.AspNetCore.Cors.EnableCors("MyPolicy")]
        [Route("[controller]/Bookmarks")]
        //Post action for processing the bookmarks from the PDF documents
        public IActionResult Bookmarks([FromBody] Dictionary<string, string> jsonObject)
        {
            //Initialize the PDF Viewer object with memory cache object
            PdfRenderer pdfviewer = new PdfRenderer(_cache);
            var jsonResult = pdfviewer.GetBookmarks(jsonObject);
            return Content(JsonConvert.SerializeObject(jsonResult));
        }

        [AcceptVerbs("Post")]
        [HttpPost("RenderPdfPages")]
        [Microsoft.AspNetCore.Cors.EnableCors("MyPolicy")]
        [Route("[controller]/RenderPdfPages")]
        //Post action for processing the PDF documents  
        public IActionResult RenderPdfPages([FromBody] Dictionary<string, string> jsonObject)
        {
            //Initialize the PDF Viewer object with memory cache object
            PdfRenderer pdfviewer = new PdfRenderer(_cache);
            object jsonResult = pdfviewer.GetPage(jsonObject);
            return Content(JsonConvert.SerializeObject(jsonResult));
        }

        [AcceptVerbs("Post")]
        [HttpPost("RenderPdfTexts")]
        [Microsoft.AspNetCore.Cors.EnableCors("MyPolicy")]
        [Route("[controller]/RenderPdfTexts")]
        //Post action for processing the PDF texts  
        public IActionResult RenderPdfTexts([FromBody] Dictionary<string, string> jsonObject)
        {
            //Initialize the PDF Viewer object with memory cache object
            PdfRenderer pdfviewer = new PdfRenderer(_cache);
            object jsonResult = pdfviewer.GetDocumentText(jsonObject);
            return Content(JsonConvert.SerializeObject(jsonResult));
        }

        [AcceptVerbs("Post")]
        [HttpPost("RenderThumbnailImages")]
        [Microsoft.AspNetCore.Cors.EnableCors("MyPolicy")]
        [Route("[controller]/RenderThumbnailImages")]
        //Post action for rendering the ThumbnailImages
        public IActionResult RenderThumbnailImages([FromBody] Dictionary<string, string> jsonObject)
        {
            //Initialize the PDF Viewer object with memory cache object
            PdfRenderer pdfviewer = new PdfRenderer(_cache);
            object result = pdfviewer.GetThumbnailImages(jsonObject);
            return Content(JsonConvert.SerializeObject(result));
        }
        [AcceptVerbs("Post")]
        [HttpPost("RenderAnnotationComments")]
        [Microsoft.AspNetCore.Cors.EnableCors("MyPolicy")]
        [Route("[controller]/RenderAnnotationComments")]
        //Post action for rendering the annotations
        public IActionResult RenderAnnotationComments([FromBody] Dictionary<string, string> jsonObject)
        {
            //Initialize the PDF Viewer object with memory cache object
            PdfRenderer pdfviewer = new PdfRenderer(_cache);
            object jsonResult = pdfviewer.GetAnnotationComments(jsonObject);
            return Content(JsonConvert.SerializeObject(jsonResult));
        }
        [AcceptVerbs("Post")]
        [HttpPost("ExportAnnotations")]
        [Microsoft.AspNetCore.Cors.EnableCors("MyPolicy")]
        [Route("[controller]/ExportAnnotations")]
        //Post action to export annotations
        public IActionResult ExportAnnotations([FromBody] Dictionary<string, string> jsonObject)
        {
            PdfRenderer pdfviewer = new PdfRenderer(_cache);
            string jsonResult = pdfviewer.ExportAnnotation(jsonObject);
            return Content(jsonResult);
        }
        [AcceptVerbs("Post")]
        [HttpPost("ImportAnnotations")]
        [Microsoft.AspNetCore.Cors.EnableCors("MyPolicy")]
        [Route("[controller]/ImportAnnotations")]
        //Post action to import annotations
        public IActionResult ImportAnnotations([FromBody] Dictionary<string, string> jsonObject)
        {
            PdfRenderer pdfviewer = new PdfRenderer(_cache);
            string jsonResult = string.Empty;
            object JsonResult;
            if (jsonObject != null && jsonObject.ContainsKey("fileName"))
            {
                string documentPath = GetDocumentPath(jsonObject["fileName"]);
                if (!string.IsNullOrEmpty(documentPath))
                {
                    jsonResult = System.IO.File.ReadAllText(documentPath);
                }
                else
                {
                    return this.Content(jsonObject["document"] + " is not found");
                }
            }
            else
            {
                string extension = Path.GetExtension(jsonObject["importedData"]);
                if (extension != ".xfdf")
                {
                    JsonResult = pdfviewer.ImportAnnotation(jsonObject);
                    return Content(JsonConvert.SerializeObject(JsonResult));
                }
                else
                {
                    string documentPath = GetDocumentPath(jsonObject["importedData"]);
                    if (!string.IsNullOrEmpty(documentPath))
                    {
                        byte[] bytes = System.IO.File.ReadAllBytes(documentPath);
                        jsonObject["importedData"] = Convert.ToBase64String(bytes);
                        JsonResult = pdfviewer.ImportAnnotation(jsonObject);
                        return Content(JsonConvert.SerializeObject(JsonResult));
                    }
                    else
                    {
                        return this.Content(jsonObject["document"] + " is not found");
                    }
                }
            }
            return Content(jsonResult);
        }

        [AcceptVerbs("Post")]
        [HttpPost("ExportFormFields")]
        [Microsoft.AspNetCore.Cors.EnableCors("MyPolicy")]
        [Route("[controller]/ExportFormFields")]
        public IActionResult ExportFormFields([FromBody] Dictionary<string, string> jsonObject)

        {
            PdfRenderer pdfviewer = new PdfRenderer(_cache);
            string jsonResult = pdfviewer.ExportFormFields(jsonObject);
            return Content(jsonResult);
        }

        [AcceptVerbs("Post")]
        [HttpPost("ImportFormFields")]
        [Microsoft.AspNetCore.Cors.EnableCors("MyPolicy")]
        [Route("[controller]/ImportFormFields")]
        public IActionResult ImportFormFields([FromBody] Dictionary<string, string> jsonObject)
        {
            PdfRenderer pdfviewer = new PdfRenderer(_cache);
            jsonObject["data"] = GetDocumentPath(jsonObject["data"]);
            object jsonResult = pdfviewer.ImportFormFields(jsonObject);
            return Content(JsonConvert.SerializeObject(jsonResult));
        }

        [AcceptVerbs("Post")]
        [HttpPost("Unload")]
        [Microsoft.AspNetCore.Cors.EnableCors("MyPolicy")]
        [Route("[controller]/Unload")]
        //Post action for unloading and disposing the PDF document resources  
        public IActionResult Unload([FromBody] Dictionary<string, string> jsonObject)
        {
            //Initialize the PDF Viewer object with memory cache object
            PdfRenderer pdfviewer = new PdfRenderer(_cache);
            pdfviewer.ClearCache(jsonObject);
            return this.Content("Document cache is cleared");
        }


        [HttpPost("Download")]
        [Microsoft.AspNetCore.Cors.EnableCors("MyPolicy")]
        [Route("[controller]/Download")]
        //Post action for downloading the PDF documents

        public async Task<IActionResult> Download([FromBody] Dictionary<string, string> jsonObject)
        {
            //Initialize the PDF Viewer object with memory cache object
            PdfRenderer pdfviewer = new PdfRenderer(_cache);
            string documentBase = pdfviewer.GetDocumentAsBase64(jsonObject);
            byte[] documentBytes = Convert.FromBase64String(documentBase.Split(",")[1]);
            string documentId = jsonObject["documentId"];
            string result = Path.GetFileNameWithoutExtension(documentId);
            string fileName = result + "_downloaded.pdf";
          
            UserCredential credential;

            using (var memStream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(memStream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true));
            }
          
            // Create the Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });
            var fileMetadata = new Google.Apis.Drive.v3.Data.File()
            {
                Name = fileName,
                Parents = new List<string> { folderId }
            };
            FilesResource.CreateMediaUpload request;
            using (var stream = new MemoryStream(documentBytes))
            {
                request = service.Files.Create(fileMetadata, stream, "application/pdf");
                request.Fields = "id";
                object value = await request.UploadAsync();
            }
            return Content(documentBase);
        }

        [HttpPost("PrintImages")]
        [Microsoft.AspNetCore.Cors.EnableCors("MyPolicy")]
        [Route("[controller]/PrintImages")]
        //Post action for printing the PDF documents
        public IActionResult PrintImages([FromBody] Dictionary<string, string> jsonObject)
        {
            //Initialize the PDF Viewer object with memory cache object
            PdfRenderer pdfviewer = new PdfRenderer(_cache);
            object pageImage = pdfviewer.GetPrintImage(jsonObject);
            return Content(JsonConvert.SerializeObject(pageImage));
        }

        //Gets the path of the PDF document
        private string GetDocumentPath(string document)
        {
            string documentPath = string.Empty;
            if (!System.IO.File.Exists(document))
            {
                var path = _hostingEnvironment.ContentRootPath;
                if (System.IO.File.Exists(path + "/Data/" + document))
                    documentPath = path + "/Data/" + document;
            }
            else
            {
                documentPath = document;
            }
            Console.WriteLine(documentPath);
            return documentPath;
        }
        // GET api/values
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/values/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        }
    }
}
