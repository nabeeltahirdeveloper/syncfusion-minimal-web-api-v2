
using Microsoft.AspNetCore.Mvc;
using Syncfusion.Pdf;
using Syncfusion.DocIORenderer;
using Syncfusion.EJ2.DocumentEditor;
using WDocument = Syncfusion.DocIO.DLS.WordDocument;
using WFormatType = Syncfusion.DocIO.FormatType;
using Syncfusion.EJ2.SpellChecker;
// using Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense;
using Newtonsoft.Json;

namespace DocumentEditorCore
{
    /// <summary>
    /// Contains the server-side helper functions for JavaScript Document editor control.
    /// </summary>
    public class DocumentEditorHelper
    {
        internal List<Syncfusion.EJ2.SpellChecker.DictionaryData>? spellDictCollection;
        internal string? path;
        internal string? personalDictPath;
        public DocumentEditorHelper()
        {
            //check the spell check dictionary path environment variable value and assign default data folder
            //if it is null.
            path = Path.Combine("\\App_Data\\");
            //Set the default spellcheck.json file if the json filename is empty.
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("ORg4AjUWIQA/Gnt2VVhhQlFaclhJXGFWfVJpTGpQdk5xdV9DaVZUTWY/P1ZhSXxRdkFhXH5ccXVVQGZeVEUsdfs=");
            string jsonFileName = Path.Combine("\\App_Data\\spellcheck.json");
            if (System.IO.File.Exists(jsonFileName))
            {
                string jsonImport = System.IO.File.ReadAllText(jsonFileName);
                List<DictionaryData> spellChecks = JsonConvert.DeserializeObject<List<DictionaryData>>(jsonImport);
                spellDictCollection = new List<DictionaryData>();
                //construct the dictionary file path using customer provided path and dictionary name
                foreach (var spellCheck in spellChecks)
                {
                    spellDictCollection.Add(new DictionaryData(spellCheck.LanguadeID, Path.Combine(path, spellCheck.DictionaryPath), Path.Combine(path, spellCheck.AffixPath)));
                    personalDictPath = Path.Combine(path, spellCheck.PersonalDictPath);
                }
            }
        }

        public string? Import(IFormCollection data)
        {
            if (data.Files.Count == 0)
                return null;
            Stream stream = new MemoryStream();
            IFormFile file = data.Files[0];
            int index = file.FileName.LastIndexOf('.');
            string type = index > -1 && index < file.FileName.Length - 1 ?
                file.FileName.Substring(index) : ".docx";
            file.CopyTo(stream);
            stream.Position = 0;

            Syncfusion.EJ2.DocumentEditor.WordDocument document = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(stream, GetFormatType(type.ToLower()));
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
            document.Dispose();
            return json;
        }

        internal static FormatType GetFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            switch (format.ToLower())
            {
                case ".dotx":
                case ".docx":
                case ".docm":
                case ".dotm":
                    return FormatType.Docx;
                case ".dot":
                case ".doc":
                    return FormatType.Doc;
                case ".rtf":
                    return FormatType.Rtf;
                case ".txt":
                    return FormatType.Txt;
                case ".xml":
                    return FormatType.WordML;
                case ".html":
                    return FormatType.Html;
                default:
                    throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }
        }
        internal static WFormatType GetWFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            switch (format.ToLower())
            {
                case ".dotx":
                    return WFormatType.Dotx;
                case ".docx":
                    return WFormatType.Docx;
                case ".docm":
                    return WFormatType.Docm;
                case ".dotm":
                    return WFormatType.Dotm;
                case ".dot":
                    return WFormatType.Dot;
                case ".doc":
                    return WFormatType.Doc;
                case ".rtf":
                    return WFormatType.Rtf;
                case ".html":
                    return WFormatType.Html;
                case ".txt":
                    return WFormatType.Txt;
                case ".xml":
                    return WFormatType.WordML;
                case ".odt":
                    return WFormatType.Odt;
                default:
                    throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }
        }
        public void Save(SaveParameter data)
        {
            string name = data.FileName != null ? data.FileName : "Saveddoc.docx";
            Console.WriteLine("data.FileName: " + data.FileName);

            string format = RetrieveFileType(name);

            if (string.IsNullOrEmpty(name))
            {
                name = "Document1.doc";
            }
            WDocument document = WordDocument.Save(data.Content);
            FileStream fileStream = new FileStream(name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            document.Save(fileStream, GetWFormatType(format));
            document.Close();
            fileStream.Close();
        }


        public FileStreamResult ExportSFDT(SaveParameter data)
        {

            string name = data.FileName != null ? data.FileName : "Saveddoc.docx";
            string apiKey = data.ApiKey != null ? data.ApiKey: "";
            string ourApiKey = "dGhpcyBpcyBzeW5jZnVzaW9uIGFwaSBzZXJ2ZXI=";
            
            Console.WriteLine("name>>" + name);
            // string format = RetrieveFileType(name);
            int index = name.LastIndexOf('.');
            string format = index > -1 && index < name.Length - 1 ?
                name.Substring(index) : ".doc";
            if (string.IsNullOrEmpty(name))
            {
                name = "Document1.doc";
            }
            WDocument document = WordDocument.Save(data.Content);
            // return SaveDocument(document, format, name);
            Stream stream = new MemoryStream();
            string contentType = "";
            if (format == ".pdf")
            {
                contentType = "application/pdf";
            }
            else
            {
                WFormatType type = GetWFormatType(format);
                switch (type)
                {
                    case WFormatType.Rtf:
                        contentType = "application/rtf";
                        break;
                    case WFormatType.WordML:
                        contentType = "application/xml";
                        break;
                    case WFormatType.Html:
                        contentType = "application/html";
                        break;
                    case WFormatType.Dotx:
                        contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.template";
                        break;
                    case WFormatType.Docx:
                        contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                        break;
                    case WFormatType.Doc:
                        contentType = "application/msword";
                        break;
                    case WFormatType.Dot:
                        contentType = "application/msword";
                        break;
                }
                if (apiKey == ourApiKey){

                document.Save(stream, type);
                }
            }
            document.Close();
            stream.Position = 0;
            return new FileStreamResult(stream, contentType)
            {
                FileDownloadName = name
            };
        }
public FileStreamResult ExportSFDTtoPDF(SaveParameter data)
        {
            string name = data.FileName != null ? data.FileName : "Saveddoc.docx";
            string apiKey = data.ApiKey != null ? data.ApiKey: "";
            string ourApiKey = "dGhpcyBpcyBzeW5jZnVzaW9uIGFwaSBzZXJ2ZXI=";
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("ORg4AjUWIQA/Gnt2VVhhQlFaclhJXGFWfVJpTGpQdk5xdV9DaVZUTWY/P1ZhSXxRdkFhXH5ccsdfsXVVQGZeVEU=");

            
// string format = RetrieveFileType(name);
            int index = name.LastIndexOf('.');
            string format = index > -1 && index < name.Length - 1 ?
                name.Substring(index) : ".doc";
            if (string.IsNullOrEmpty(name))
            {
                name = "Document1.doc";
            }
            WDocument document = WordDocument.Save(data.Content);
            // return SaveDocument(document, format, name);
            Stream stream = new MemoryStream();
            string contentType = "";
            if (format == ".pdf")
            {
                contentType = "application/pdf";
            }
            else
            {
                WFormatType type = GetWFormatType(format);
                switch (type)
                {
                    case WFormatType.Rtf:
                        contentType = "application/rtf";
                        break;
                    case WFormatType.WordML:
                        contentType = "application/xml";
                        break;
                    case WFormatType.Html:
                        contentType = "application/html";
                        break;
                    case WFormatType.Dotx:
                        contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.template";
                        break;
                    case WFormatType.Docx:
                        contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                        break;
                    case WFormatType.Doc:
                        contentType = "application/msword";
                        break;
                    case WFormatType.Dot:
                        contentType = "application/msword";
                        break;
                }
                document.Save(stream, type);
            }
            document.Close();
            stream.Position = 0;
            // return new FileStreamResult(stream, contentType)
            // {
            //     FileDownloadName = name
            // };
            // if (data.Files.Count == 0)
            //     return null;
            Stream streamSFDT = new MemoryStream();
            string typeSFDT = ".docx";

            Syncfusion.DocIO.DLS.WordDocument doc = new Syncfusion.DocIO.DLS.WordDocument(stream, Syncfusion.DocIO.FormatType.Docx);
            Console.WriteLine("This is doc");
            //Instantiation of DocIORenderer for Word to PDF conversion 
            DocIORenderer render = new DocIORenderer();
            Console.WriteLine("This is render");
            //Converts Word document into PDF document 
            PdfDocument pdfDocument = render.ConvertToPDF(doc);
            Console.WriteLine("This is conversion");
            Stream streamData = new MemoryStream();
            Console.WriteLine("This is memory");
            
            //Saves the PDF file
            if (apiKey == ourApiKey)
            {
            pdfDocument.Save(streamData);
            }

            Console.WriteLine("This is save");
            streamData.Position = 0;
            Console.WriteLine("This is postion");
            pdfDocument.Close();         
            Console.WriteLine("This is close");
            // documentData.Close();
            Console.WriteLine("This is done", streamData, "stream end");
            
            return new FileStreamResult(streamData, "application/pdf")
            {
                FileDownloadName = "temp.pdf"
            };

            
        }

        private string RetrieveFileType(string name)
        {
            int index = name.LastIndexOf('.');
            string format = index > -1 && index < name.Length - 1 ?
                name.Substring(index) : ".doc";
            return format;
        }


        

    }


}
