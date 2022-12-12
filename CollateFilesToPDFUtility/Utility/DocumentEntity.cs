using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IPBSSPWebServices.Utilities
{
    public class DocumentEntity
    {
        public string docUrl { get; set; }

        public int sequence { get; set; }
    }
    public class BulkFileBytes
    {
        public byte[] fileByte { get; set; }
        public string libraryName { get; set; }
        public string siteURL { get; set; }
        public string folderURL { get; set; }
        public string folderName { get; set; }
        public string childFolderName { get; set; }
        public string fileName { get; set; }
        public int sequence { get; set; }
    }
    public class DocumentDetails
    {
       public string fileName {get; set;}
       public string fileURL { get; set; }
       public string folderName { get; set; }
       public string childFolderName { get; set; }
       public string invoiceID { get; set; }
       public string certiificateID { get; set; }
       public string activityID { get; set; }
       public string milestoneID { get; set; }
       public string libraryName { get; set; }
       public string siteURL { get; set; }
       public string folderURL { get; set; }
       public string mergedFilesFolderName { get; set; }
       public string CollatedFilesFolderName { get; set; }
       public string logFilePath { get; set; }
       public string itemID { get; set; }
       public string attachmentType { get; set; }
       public string recordId {get;set;}
       public string levelId { get; set; }
       public string collateSrNo { get; set; }
       public string CollateFileName { get; set; }
       public string DestinationFolder { get; set; }
       public string DestinationChildFolder { get; set; }
       public string SourceFolder { get; set; }
       public string SourceChildFolder { get; set; }
    }
}
