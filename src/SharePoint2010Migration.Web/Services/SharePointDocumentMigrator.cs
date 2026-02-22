using System;
using System.IO;
using System.Net;
using Microsoft.SharePoint.Client;
using SharePoint2010Migration.Web.Models;

namespace SharePoint2010Migration.Web.Services
{
    public class SharePointDocumentMigrator
    {
        private readonly DocumentRepository _repository;

        public SharePointDocumentMigrator(DocumentRepository repository)
        {
            _repository = repository;
        }

        public MigrationRunResult Run(MigrationRunRequest request)
        {
            var result = new MigrationRunResult();

            using (var context = new ClientContext(request.SiteUrl))
            {
                context.Credentials = new NetworkCredential(request.UserName, request.Password, request.Domain);

                var list = context.Web.Lists.GetByTitle(request.LibraryTitle);
                ListItemCollectionPosition position = null;

                do
                {
                    var query = BuildQuery(position, request.ModifiedSinceUtc);
                    var items = list.GetItems(query);

                    context.Load(items,
                        c => c.ListItemCollectionPosition,
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
                        result.ProcessedCount++;

                        try
                        {
                            var file = item.File;
                            context.Load(file);
                            context.ExecuteQuery();

                            if (!file.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                            {
                                result.SkippedNonPdfCount++;
                                continue;
                            }

                            var bytes = DownloadFileBytes(context, file.ServerRelativeUrl);
                            _repository.Upsert(item, file, bytes);
                            result.InsertedOrUpdatedCount++;
                        }
                        catch (Exception ex)
                        {
                            result.ErrorCount++;
                            result.LastError = ex.Message;
                        }
                    }

                    position = items.ListItemCollectionPosition;
                }
                while (position != null);
            }

            return result;
        }

        private static CamlQuery BuildQuery(ListItemCollectionPosition position, DateTime? modifiedSinceUtc)
        {
            var modifiedFilter = string.Empty;
            if (modifiedSinceUtc.HasValue)
            {
                modifiedFilter = string.Format(@"
                    <And>
                        <Eq>
                            <FieldRef Name='FSObjType' />
                            <Value Type='Integer'>0</Value>
                        </Eq>
                        <Geq>
                            <FieldRef Name='Modified' />
                            <Value IncludeTimeValue='TRUE' Type='DateTime'>{0:yyyy-MM-ddTHH:mm:ssZ}</Value>
                        </Geq>
                    </And>", modifiedSinceUtc.Value);
            }
            else
            {
                modifiedFilter = @"
                    <Eq>
                        <FieldRef Name='FSObjType' />
                        <Value Type='Integer'>0</Value>
                    </Eq>";
            }

            return new CamlQuery
            {
                ListItemCollectionPosition = position,
                ViewXml = string.Format(@"<View Scope='RecursiveAll'>
                    <Query>
                        <Where>
                            {0}
                        </Where>
                    </Query>
                    <RowLimit Paged='TRUE'>200</RowLimit>
                </View>", modifiedFilter)
            };
        }

        private static byte[] DownloadFileBytes(ClientContext context, string serverRelativeUrl)
        {
            var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, serverRelativeUrl);

            using (var ms = new MemoryStream())
            {
                fileInfo.Stream.CopyTo(ms);
                return ms.ToArray();
            }
        }
    }
}
