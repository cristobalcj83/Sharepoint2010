namespace SharePoint2010Migration.Web.Models
{
    public class MigrationRunResult
    {
        public int ProcessedCount { get; set; }
        public int InsertedOrUpdatedCount { get; set; }
        public int SkippedNonPdfCount { get; set; }
        public int ErrorCount { get; set; }
        public string LastError { get; set; }
    }
}
