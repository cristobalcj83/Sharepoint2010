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
GO

CREATE UNIQUE INDEX UX_DocumentArchive_SharePointItemId
ON dbo.DocumentArchive(SharePointItemId);
GO
