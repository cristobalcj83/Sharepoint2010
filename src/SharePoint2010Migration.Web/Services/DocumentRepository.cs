using System;
using System.Data;
using System.Data.SqlClient;
using Microsoft.SharePoint.Client;

namespace SharePoint2010Migration.Web.Services
{
    public class DocumentRepository
    {
        private readonly string _connectionString;

        public DocumentRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        public void Upsert(ListItem item, File file, byte[] contentBytes)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
MERGE dbo.DocumentArchive AS target
USING (SELECT @SharePointItemId AS SharePointItemId) AS source
ON target.SharePointItemId = source.SharePointItemId
WHEN MATCHED THEN UPDATE SET
    FileName = @FileName,
    ServerRelativeUrl = @ServerRelativeUrl,
    ContentType = @ContentType,
    CreatedUtc = @CreatedUtc,
    ModifiedUtc = @ModifiedUtc,
    Author = @Author,
    Editor = @Editor,
    CustomMetadataJson = @CustomMetadataJson,
    PdfContent = @PdfContent,
    MigratedAtUtc = GETUTCDATE()
WHEN NOT MATCHED THEN
    INSERT (SharePointItemId, FileName, ServerRelativeUrl, ContentType, CreatedUtc, ModifiedUtc, Author, Editor, CustomMetadataJson, PdfContent)
    VALUES (@SharePointItemId, @FileName, @ServerRelativeUrl, @ContentType, @CreatedUtc, @ModifiedUtc, @Author, @Editor, @CustomMetadataJson, @PdfContent);";

                cmd.Parameters.AddWithValue("@SharePointItemId", item.Id);
                cmd.Parameters.AddWithValue("@FileName", (object)file.Name ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ServerRelativeUrl", (object)file.ServerRelativeUrl ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ContentType", item["ContentType"] != null ? item["ContentType"].ToString() : (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@CreatedUtc", item["Created"] ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@ModifiedUtc", item["Modified"] ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Author", item["Author"] != null ? item["Author"].ToString() : (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Editor", item["Editor"] != null ? item["Editor"].ToString() : (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@CustomMetadataJson", "{}");

                var contentParam = cmd.Parameters.Add("@PdfContent", SqlDbType.VarBinary, -1);
                contentParam.Value = contentBytes;

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }
    }
}
