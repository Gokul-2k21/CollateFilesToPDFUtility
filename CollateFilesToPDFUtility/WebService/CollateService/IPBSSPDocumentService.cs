using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Services;
using System.Drawing;
using System.Data.SqlClient;
using System.Threading;
using System.Data;
using IPBSSPWebServices.Utilities;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace IPBSSPWebServices.Layouts.WebServices
{
    public class IPBSSPDocumentService : WebService
    {
        #region Declaration of Variables
        string message = string.Empty;
        public static object missing = System.Reflection.Missing.Value;
        private static string sourcefolder;
        private static string destinationfile;
        private static IList fileList = new ArrayList();
        string storedConvertedPDFFiles = "ConvertedPDFFiles";
        string fileURL = string.Empty;
        string errormessage= string.Empty;
        string  currentdocumentname =string.Empty;
        List<DocumentDetails> _documents = new List<DocumentDetails>();

        #endregion


        /// <summary>
        /// Collate all the files based on sequence
        /// </summary>
        /// <param name="siteUrl"></param>send the site url 
        /// <param name="FileURL"></param>send the file url of which the 
        /// <param name="libraryName"></param>send the library name where the documents are stored
        /// <param name="mergedFolderName"></param>send the folder name where the merge/collated files copy to be stored
        /// <returns></returns>                      
        [WebMethod]
        public string CollateFilesInPdf(List<DocumentEntity> filesList, String DestinationFolder, String FileName)
        {
            byte[] mergeBytes = null;
            byte[] fileBytes = null;
            string MergeResult = "";
            string Error = "";
            try
            {
                DestinationFolder = DestinationFolder + "\\" + FileName + ".pdf";


                //-----SORTING BASED ON SEQUENCE
                filesList.Sort((x, y) => {
                    int ret = String.Compare(x.sequence.ToString(), y.sequence.ToString());
                    return ret;
                });

                //Perform the conversion in memory first
                using (MemoryStream ms = new MemoryStream())
                {

                    //Using the ItextSharp Document
                    using (Document document = new Document())
                    {
                        //Using the ITextSharp PdfCopy to create a PDF document in the memory stream
                        using (PdfCopy copy = new PdfCopy(document, ms))
                        {

                            //Open the document before any changes can be made.
                            document.Open();
                            PdfReader reader = null;
                            //Loop through each file
                            for (int i = 0; i <= filesList.Count - 1; i++)
                            {
                                if (filesList[i] != null)
                                {
                                    if (!string.IsNullOrEmpty(filesList[i].docUrl))
                                    {
                                        var filepath = filesList[i].docUrl;
                                        string filename = Path.GetFileName(filepath);
                                        currentdocumentname = string.Empty;
                                        currentdocumentname = filename;
                                        if (Path.GetExtension(filename).ToLower().Contains("pdf"))
                                        {
                                            bool result=IsValidPdf(filepath,null);
                                            if (result == false)
                                            {
                                                errormessage = "FileName : "+ filename +" is Corrupted !.\n"+ errormessage;
                                                return errormessage;
                                            }
                                            else
                                            {
                                                reader = new PdfReader(filepath);
                                            }
                                        }
                                        else if (Path.GetExtension(filename).ToLower().Contains("bmp") ||
                                                               Path.GetExtension(filename).ToLower().Contains("jpg") ||
                                                               Path.GetExtension(filename).ToLower().Contains("png") ||
                                                               Path.GetExtension(filename).ToLower().Contains("gif"))
                                        {
                                            fileBytes = System.IO.File.ReadAllBytes(filepath);
                                            byte[] resultarray = ConvertWordExcelPowerpointToPDF.ConvertImageIntoSinglePDF(filepath, filename, fileBytes);
                                            bool result = IsValidPdf("", resultarray);
                                            if (result == false)
                                            {
                                                errormessage = "FileName : " + filename + " is Corrupted !.\n" + errormessage;
                                                return errormessage;
                                            }
                                            else
                                            {
                                                reader = new PdfReader(resultarray);
                                            }
                                        }
                                        else if (Path.GetExtension(filename).ToLower().Contains("ppt"))
                                        {
                                            string pptURL = ConvertWordExcelPowerpointToPDF.ConvertPPTXToPDF(filepath, storedConvertedPDFFiles);
                                            bool result = IsValidPdf(pptURL, null);
                                            if (result == false)
                                            {
                                                errormessage = "FileName : " + filename + " is Corrupted !.\n" + errormessage;
                                                return errormessage;
                                            }
                                            else
                                            {
                                                reader = new PdfReader(File.ReadAllBytes(pptURL));
                                            }
                                        }
                                        else if (Path.GetExtension(filename).ToLower().Contains("doc") || Path.GetExtension(filename).ToLower().Contains("docx"))
                                        {
                                            string docURL = ConvertWordExcelPowerpointToPDF.ConvertWordToPdf(filepath, filename, storedConvertedPDFFiles);
                                            bool result = IsValidPdf(docURL, null);
                                            if (result == false)
                                            {
                                                errormessage = "FileName : " + filename + " is Corrupted !.\n" + errormessage;
                                                return errormessage;
                                            }
                                            else
                                            {
                                                reader = new PdfReader(File.ReadAllBytes(docURL));
                                            }
                                        }
                                        else if (Path.GetExtension(filename).ToLower().Contains("xls") || Path.GetExtension(filename).ToLower().Contains("xlsx"))
                                        {
                                            string fileExcelURL = ConvertWordExcelPowerpointToPDF.ConvertExcelToPdf(filepath, storedConvertedPDFFiles);
                                            bool result = IsValidPdf(fileExcelURL, null);
                                            if (result == false)
                                            {
                                                errormessage = "FileName : " + filename + " is Corrupted !.\n" + errormessage;
                                                return errormessage;
                                            }
                                            else
                                            {
                                                reader = new PdfReader(File.ReadAllBytes(fileExcelURL));
                                            }
                                        }

                                        int n = reader.NumberOfPages;
                                        //Loop through each page in the current PDF file
                                        for (int page = 0; page < n-1;)
                                        {
                                            //Import the page to the PDF document in the memory stream.
                                            copy.AddPage(copy.GetImportedPage(reader, ++page));
                                        }
                                    }
                                }
                            }
                        }

                        mergeBytes = ms.ToArray();
                        File.WriteAllBytes(DestinationFolder, mergeBytes);
                        MergeResult = "Files Merged and Saved To Destination Folder!.";
                        return MergeResult;
                    }
                }

            }
            catch (Exception ex)
            {
                if (errormessage != null)
                {
                    return errormessage;
                }
                else
                {
                    Error = "Error Occured in File Name : " + currentdocumentname + " \n  Error : " + ex.Message.ToString();
                    return Error;
                }
            }

        }
        private bool IsValidPdf(string filepath,byte[] imagefile)
        {
            bool Ret = true;

            PdfReader reader = null;

            try
            {
                if (filepath != "")
                {
                    reader = new PdfReader(filepath);
                }
                else
                {
                    reader = new PdfReader(imagefile);
                }
            }
            catch (Exception ex)
            {
                errormessage= ex.Message.ToString();
                Ret = false;
            }

            return Ret;
        }

    }
}
