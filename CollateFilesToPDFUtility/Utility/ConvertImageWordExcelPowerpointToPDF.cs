using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Collections;
using msExcel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Web;
using Microsoft.Office.Core;
//using System.Web.Hosting;

namespace IPBSSPWebServices.Utilities
{
    public class ConvertWordExcelPowerpointToPDF
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool SetVolumeLabel(string lpRootPathName, string lpVolumeName);
        public static object missing = System.Reflection.Missing.Value;
        private static string sourcefolder;
        private static string destinationfile;
        private static IList fileList = new ArrayList();
        public string SourceFolder
        {
            get { return sourcefolder; }
            set { sourcefolder = value; }
        }
        public string DestinationFile
        {
            get { return destinationfile; }
            set { destinationfile = value; }
        }       
        public static void AddFile(string pathnname)
        {
            fileList.Add(pathnname);
        }       
        public static string ConvertExcelToPdf(string excelFileIn, string pdffilefolder)
        {
            string fileURL = string.Empty;
                 msExcel.Application excel = new msExcel.Application();
                try
                {
                    excel.Visible = false;
                    excel.ScreenUpdating = false;
                    excel.DisplayAlerts = false;

                    msExcel.Workbook wbk = excel.Workbooks.Open(excelFileIn, missing,
                    missing, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing,
                    missing, missing, missing);
                    wbk.Activate();
              
                    // A4 papersize
              
                    Worksheet ws = wbk.Application.ActiveSheet as Worksheet;

                    var _with1 = ws.PageSetup;
                    _with1.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4;
                    _with1.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                        System.IO.File.Copy(excelFileIn, excelFileIn.Replace(".xlsx", ".pdf"), true);
                        fileURL = excelFileIn.Replace(".xlsx", ".pdf");
                   
                    msExcel.XlFixedFormatType fileFormat = msExcel.XlFixedFormatType.xlTypePDF;
                   
                    // Save document into PDF Format
                    wbk.ExportAsFixedFormat(fileFormat, excelFileIn.Replace(".xlsx", ".pdf"),
                    missing, missing, missing,
                    missing, missing, missing,
                    missing);

                    object saveChanges = msExcel.XlSaveAction.xlDoNotSaveChanges;
                    ((msExcel._Workbook)wbk).Close(saveChanges, missing, missing);
                    wbk = null;
                }
                catch (Exception e)
                {
                  return e.Message.ToString();
                }
                finally
                {
                ((msExcel._Application)excel).Quit();
                excel = null;
                }
                return fileURL;
        }
        #region Converting word to PDF
        // C# doesn’t have optional arguments so we’ll need a dummy value
        // object oMissing = System.Reflection.Missing.Value;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="FilePath"></param>
        /// <param name="FileName"></param>
        /// <param name="pdffilefolder"></param>
        /// <returns></returns>
        public static string  ConvertWordToPdf(String FilePath, String FileName,string pdffilefolder)
        {
            string fileURL = string.Empty;
              
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            try
            {
                // Get a Word file
              
                word.Visible = false;
                word.ScreenUpdating = false;
                // Cast as Object for word Open method
                Object filename = (Object)FilePath;              
                // Use the dummy value as a placeholder for optional arguments
                Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref filename, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing);
                doc.Activate();
                //object outputFileName = System.Web.HttpContext.Current.Server.MapPath("~/" + pdffilefolder + "/" + Path.GetFileName(FileName).Replace(".docx", ".pdf"));//Path.Combine(HttpRuntime.AppDomainAppPath, "Functional doc.docx").Replace(".docx", ".pdf"); 
                object outputFileName = FilePath.Replace(".docx", ".pdf");
                System.IO.File.Copy(FilePath, FilePath.Replace(".docx", ".pdf"), true);
                object fileFormat = WdSaveFormat.wdFormatPDF;
                // Save document into PDF Format
                doc.SaveAs(ref outputFileName,
                ref fileFormat, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing);
                // Close the Word document, but leave the Word application open.
                // doc has to be cast to type _Document so that it will find the
                // correct Close method.
                fileURL = FilePath.Replace(".docx", ".pdf");
                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)doc).Close(ref saveChanges, ref missing, ref missing);
                doc = null;
                // word has to be cast to type _Application so that it will find
                // the correct Quit method.
                ((Microsoft.Office.Interop.Word._Application)word).Quit(ref missing, ref missing, ref missing);
                word = null;
            }
            catch (Exception e)
            {
                return e.Message.ToString();
            }
            return fileURL;
        }
        #endregion
        /// <summary>
        /// 
        /// </summary>
        /// <param name="FileName"></param>
        /// <param name="pdffilefolder"></param>
        /// <returns></returns>
        public static string ConvertPPTXToPDF(string FileName,string pdffilefolder)
        {
            Microsoft.Office.Interop.PowerPoint.Application app = new Microsoft.Office.Interop.PowerPoint.Application();
            string sourcePptx = FileName;
            string targetPpt = string.Empty;
                System.IO.File.Copy(FileName, FileName.Replace(".ppt", ".pdf"), true);
                targetPpt = FileName.Replace(".ppt", ".pdf");                     
            object missing = Type.Missing;
            Microsoft.Office.Interop.PowerPoint.Presentation pptx = app.Presentations.Open(sourcePptx, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
            pptx.SaveAs(targetPpt, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPDF, MsoTriState.msoTrue);
            app.Quit();
            return targetPpt;
        }

        /// <summary>
        /// Convert image to PDF
        /// </summary>
        /// <param name="path"></param>
        /// <param name="siteURL"></param>
        /// <returns></returns>
        public static byte[] ConvertImageIntoSinglePDF(string path, string siteURL,byte[] fileBytes)
        {
            try
            {
                iTextSharp.text.Document doc = new iTextSharp.text.Document();
                doc.SetPageSize(PageSize.A4);

                var ms = new MemoryStream();
                {
                    PdfCopy pdf = new PdfCopy(doc, ms);
                    doc.Open();

                    byte[] data = fileBytes;
                    doc.NewPage();
                    iTextSharp.text.Document imageDocument = null;
                    PdfWriter imageDocumentWriter = null;
                    switch (Path.GetExtension(path).ToLower().Trim('.'))
                    {
                        case "bmp":
                        case "gif":
                        case "jpg":
                        case "png":
                            imageDocument = new iTextSharp.text.Document();
                            using (var imageMS = new MemoryStream())
                            {
                                imageDocumentWriter = PdfWriter.GetInstance(imageDocument, imageMS);
                                imageDocument.Open();
                                if (imageDocument.NewPage())
                                {
                                    var image = iTextSharp.text.Image.GetInstance(data);
                                    image.Alignment = Element.ALIGN_CENTER;
                                    if (image.Width > doc.PageSize.Width)
                                    {
                                        image.ScaleToFit(doc.PageSize.Width - 10, doc.PageSize.Height - 10);
                                    }
                                    if (!imageDocument.Add(image))
                                    {
                                        throw new Exception("Unable to add image to page!");
                                    }
                                    imageDocument.Close();
                                    imageDocumentWriter.Close();
                                    PdfReader imageDocumentReader = new PdfReader(imageMS.ToArray());
                                    var page = pdf.GetImportedPage(imageDocumentReader, 1);
                                    pdf.AddPage(page);
                                    imageDocumentReader.Close();
                                }
                            }
                            break;

                        case "pdf":
                            var rdr = new PdfReader(data);
                            for (int i = 0; i < rdr.NumberOfPages; i++)
                            {
                                pdf.AddPage(pdf.GetImportedPage(rdr, i + 1));
                            }
                            pdf.FreeReader(rdr);
                            rdr.Close();
                            break;
                        default:
                            // not supported image format:
                            // skip it (or throw an exception if you prefer)
                            break;
                    }

                }
                if (doc.IsOpen()) doc.Close();
                return ms.ToArray();
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public static byte[] ConvertTextFileToPDF(string path, string siteUrl, byte[] fileBytes)
        {
            iTextSharp.text.Document doc = new iTextSharp.text.Document();
            doc.SetPageSize(PageSize.A4);
            StreamReader rdr = new StreamReader(path);
            var ms = new MemoryStream();
            {
                PdfCopy pdf = new PdfCopy(doc, ms);
                doc.Open();
                //

                //using (SPSite spCurrentSite = new SPSite(siteUrl))
                //{
                //    using (SPWeb objSpWeb = spCurrentSite.OpenWeb())
                //    {

                byte[] data = fileBytes;
                doc.NewPage();
                iTextSharp.text.Document imageDocument = null;
                PdfWriter imageDocumentWriter = null;

                imageDocument = new iTextSharp.text.Document();
                using (var imageMS = new MemoryStream())
                {
                    imageDocumentWriter = PdfWriter.GetInstance(doc, new FileStream(path, FileMode.Create));
                    imageDocument.Open();
                    if (imageDocument.NewPage())
                    {
                        var image = new iTextSharp.text.Paragraph(rdr.ReadToEnd()); //iTextSharp.text.Image.GetInstance(data);

                        imageDocument.Close();
                        imageDocumentWriter.Close();
                        PdfReader imageDocumentReader = new PdfReader(imageMS.ToArray());
                        var page = pdf.GetImportedPage(imageDocumentReader, 1);
                        pdf.AddPage(page);
                        imageDocumentReader.Close();
                    }
                }




                //            }
                //        }
                //    //});
                //}


                if (doc.IsOpen()) doc.Close();
                return ms.ToArray();








                //StreamReader rdr = new StreamReader(path);

                ////Create a New instance on Document Class

                //iTextSharp.text.Document doc = new iTextSharp.text.Document();

                ////Create a New instance of PDFWriter Class for Output File

                //PdfWriter.GetInstance(doc, new FileStream(path, FileMode.Create));

                ////Open the Document

                //doc.Open();

                ////Add the content of Text File to PDF File

                //doc.Add(new iTextSharp.text.Paragraph(rdr.ReadToEnd()));

                ////Close the Document

                //doc.Close();
                //byte[] bytes = System.IO.File.ReadAllBytes();
                //return bytes;
            }
        }
      
    }

}
