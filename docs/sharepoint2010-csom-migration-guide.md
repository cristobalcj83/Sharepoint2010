# SharePoint 2010 CSOM Prototype: Extract PDFs + Metadata to SQL Server

This guide shows how to build a **local prototype web app** that reads documents and metadata from a SharePoint 2010 document library and stores them in a non-SharePoint SQL Server database.

## 1) Recommended architecture

1. Build an ASP.NET web app (or worker hosted by web app) on your local machine.
2. Use SharePoint 2010 CSOM to query list items and download file bytes.
3. Map SharePoint fields to your SQL schema.
4. Store:
   - Metadata in SQL columns.
   - PDF content in a `VARBINARY(MAX)` column (your “blob format”).
5. Keep a migration log table to support retries and incremental runs.

## 2) What to install (local machine)

For SharePoint 2010 CSOM you typically need:

- **Visual Studio** (2019/2022 is fine for editing old .NET Framework projects).
- **.NET Framework 4.0 or 4.5** target pack (4.0 is closest to SharePoint 2010 era; 4.5 also often works in migration tools).
- SharePoint 2010 CSOM assemblies (copy from a SharePoint server, commonly under `C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\14\ISAPI`):
  - `Microsoft.SharePoint.Client.dll`
  - `Microsoft.SharePoint.Client.Runtime.dll`
- **SQL Server** (local SQL Express/Developer edition is fine).
- **IIS Express** or local IIS if hosting a classic ASP.NET web app.

> Note: SharePoint 2010 does not support modern OAuth flows like SharePoint Online. For on-premises 2010 you usually use Windows/domain credentials.

## 3) Suggested SQL schema

```sql
CREATE TABLE dbo.DocumentArchive
(
    Id INT IDENTITY(1,1) PRIMARY KEY,
    SharePointItemId INT NOT NULL,
    FileName NVARCHAR(260) NOT NULL,
    ServerRelativeUrl NVARCHAR(1024) NOT NULL,
    ContentType NVARCHAR(256) NULL,
    CreatedUtc DATETIME NULL,
    ModifiedUtc DATETIME NULL,
    Author NVARCHAR(256) NULL,
    Editor NVARCHAR(256) NULL,
    CustomMetadataJson NVARCHAR(MAX) NULL,
    PdfContent VARBINARY(MAX) NOT NULL,
    MigratedAtUtc DATETIME NOT NULL DEFAULT(GETUTCDATE())
);

CREATE UNIQUE INDEX UX_DocumentArchive_SharePointItemId
    ON dbo.DocumentArchive(SharePointItemId);
```

This stores metadata in normal columns and keeps PDF bytes in `VARBINARY(MAX)`.

## 4) CSOM extraction flow (C# example)

```csharp
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using Microsoft.SharePoint.Client;

public class SharePointDocumentMigrator
{
    private readonly string _siteUrl;
    private readonly string _libraryTitle;
    private readonly string _sqlConnectionString;

    public SharePointDocumentMigrator(string siteUrl, string libraryTitle, string sqlConnectionString)
    {
        _siteUrl = siteUrl;
        _libraryTitle = libraryTitle;
        _sqlConnectionString = sqlConnectionString;
    }

    public void Run(string userName, string password, string domain)
    {
        using (var context = new ClientContext(_siteUrl))
        {
            context.Credentials = new System.Net.NetworkCredential(userName, password, domain);

            var list = context.Web.Lists.GetByTitle(_libraryTitle);
            var caml = new CamlQuery
            {
                ViewXml = @"<View Scope='RecursiveAll'>
                                <Query>
                                    <Where>
                                        <Eq>
                                            <FieldRef Name='FSObjType' />
                                            <Value Type='Integer'>0</Value>
                                        </Eq>
                                    </Where>
                                </Query>
                             </View>"
            };

            ListItemCollection items = list.GetItems(caml);
            context.Load(items,
                c => c.Include(
                    i => i.Id,
                    i => i["FileLeafRef"],
                    i => i["FileRef"],
                    i => i["ContentType"],
                    i => i["Created"],
                    i => i["Modified"],
                    i => i["Author"],
                    i => i["Editor"],
                    i => i.File));

            context.ExecuteQuery();

            foreach (var item in items)
            {
                var file = item.File;
                context.Load(file);
                context.ExecuteQuery();

                if (!file.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    continue;

                FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, file.ServerRelativeUrl);
                byte[] bytes;
                using (var ms = new System.IO.MemoryStream())
                {
                    fileInfo.Stream.CopyTo(ms);
                    bytes = ms.ToArray();
                }

                SaveToSql(item, file, bytes);
            }
        }
    }

    private void SaveToSql(ListItem item, Microsoft.SharePoint.Client.File file, byte[] pdfBytes)
    {
        using (var conn = new SqlConnection(_sqlConnectionString))
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

            // Serialize your extra fields to JSON here if needed.
            cmd.Parameters.AddWithValue("@CustomMetadataJson", "{}");

            var p = cmd.Parameters.Add("@PdfContent", SqlDbType.VarBinary, -1);
            p.Value = pdfBytes;

            conn.Open();
            cmd.ExecuteNonQuery();
        }
    }
}
```

## 5) Important implementation notes

- Use **paging** for large libraries (CAML RowLimit + ListItemCollectionPosition).
- Keep an **incremental strategy** (e.g., filter by `Modified >= lastRunUtc`).
- Add retry + logging for transient connectivity failures.
- If metadata contains user fields or taxonomy-like fields, parse CSOM objects carefully (often lookup/value pairs).
- Store credentials securely (Windows Credential Manager/DPAPI/app config encryption), never hard-code passwords.

## 6) Minimal web app shape

For your prototype, a simple structure is enough:

- `/Migrate` page with form inputs:
  - Site URL
  - Library name
  - Domain/user/password
  - SQL connection string
  - Optional “Modified since” date
- A “Run migration” button that invokes the migrator service.
- A status panel showing count of processed documents, inserted/updated rows, and errors.

## 7) Production-readiness checklist

- Run under a service account with least privileges on SharePoint + SQL.
- Add duplicate detection based on `SharePointItemId` and/or URL + hash.
- Optionally compute SHA-256 of PDF bytes for integrity checks.
- Add audit table for each migration batch.
- Validate SQL backup/restore and rollback plan before cutover.

## 8) Step-by-step local installation and first run

Follow this sequence exactly on your laptop + VM setup.

### Step 1: Prepare your SharePoint 2010 VM

1. Ensure your document library exists and contains a few PDF test files.
2. Confirm the site URL is reachable from your laptop browser (for example `http://sp2010vm/sites/migration`).
3. Confirm the user account you will use has at least **Read** on the document library.
4. On the VM, note the location of CSOM DLLs (usually):
   - `C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\14\ISAPI\Microsoft.SharePoint.Client.dll`
   - `C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\14\ISAPI\Microsoft.SharePoint.Client.Runtime.dll`

### Step 2: Install local prerequisites on your laptop

1. Install **Visual Studio 2019 or 2022** with .NET desktop/web tooling.
2. Install **.NET Framework developer pack** (4.0 or 4.5).
3. Install **SQL Server Express/Developer** and **SSMS**.
4. Verify IIS Express is available (installed with Visual Studio).

### Step 3: Create the prototype web app project

1. Open Visual Studio.
2. Create a new **ASP.NET Web Application (.NET Framework)** project.
3. Choose **MVC** template (or Empty + MVC).
4. Set target framework to **.NET Framework 4.5** (or 4.0 if required by your environment).

### Step 4: Add SharePoint 2010 CSOM references

1. Copy the two DLLs from the VM to a local folder in your project, for example `lib\sp2010\`.
2. In Visual Studio: **Project > Add Reference > Browse**.
3. Select:
   - `Microsoft.SharePoint.Client.dll`
   - `Microsoft.SharePoint.Client.Runtime.dll`
4. Set both references to:
   - `Copy Local = true`
   - `Specific Version = false` (if needed to avoid version-binding issues).

### Step 5: Create SQL database and migration table

1. In SSMS, create a database, for example `SharePointMigrationLab`.
2. Run the SQL script from section **3) Suggested SQL schema** in this document.
3. Confirm table exists:

```sql
SELECT TOP 1 * FROM dbo.DocumentArchive;
```

### Step 6: Add configuration values

In `Web.config` add app settings / connection string values:

```xml
<appSettings>
  <add key="SharePointSiteUrl" value="http://sp2010vm/sites/migration" />
  <add key="SharePointLibraryTitle" value="Shared Documents" />
  <add key="SharePointDomain" value="CONTOSO" />
  <add key="SharePointUser" value="spreader" />
</appSettings>

<connectionStrings>
  <add name="MigrationDb"
       connectionString="Server=localhost;Database=SharePointMigrationLab;Trusted_Connection=True;"
       providerName="System.Data.SqlClient" />
</connectionStrings>
```

> For prototype only, you may prompt password on the page at runtime. For production, use a secure secret storage approach.

### Step 7: Add migration service code

1. Create a service class (for example `Services/SharePointDocumentMigrator.cs`).
2. Paste and adapt the sample from section **4) CSOM extraction flow**.
3. Add a simple MVC controller endpoint, e.g. `POST /Migrate/Run`, that calls `Run(user, password, domain)`.

### Step 8: Run the web app and execute first migration

1. Press **F5** to run locally.
2. Open your migration page and provide inputs:
   - Site URL
   - Library name
   - Domain/user/password
3. Start migration.
4. Check SQL result:

```sql
SELECT COUNT(*) AS MigratedPdfCount FROM dbo.DocumentArchive;
SELECT TOP 10 SharePointItemId, FileName, DATALENGTH(PdfContent) AS PdfBytes
FROM dbo.DocumentArchive
ORDER BY Id DESC;
```

### Step 9: Validate metadata mapping

1. Pick one PDF item in SharePoint.
2. Compare SharePoint fields (`Created`, `Modified`, `Editor`, custom columns) with row data in SQL.
3. If custom fields are missing, extend your mapping and store extras in `CustomMetadataJson`.

### Step 10: Troubleshooting checklist

- **401/403 errors:** confirm credentials and library permissions.
- **Could not load file or assembly (SharePoint.Client):** check reference versions and `Copy Local = true`.
- **SQL login failure:** verify connection string and SQL authentication mode.
- **Timeout on large libraries:** implement CAML paging and incremental load by `Modified` date.
- **Non-PDF files are skipped:** expected with current sample; remove extension filter if you want all document types.

---

If you want, next step can be a full **starter project structure** (ASP.NET MVC + service classes + SQL repository + config examples) that you can open directly in Visual Studio.
