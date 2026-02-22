# SharePoint 2010 Migration Prototype (Code)

## Structure

- `src/SharePoint2010Migration.Web` - ASP.NET MVC (.NET Framework) web app.
- `database/create-document-archive.sql` - SQL table setup.
- `lib/sp2010` - place SharePoint 2010 CSOM DLLs here.

## Required DLLs

Copy these into `lib/sp2010`:

- `Microsoft.SharePoint.Client.dll`
- `Microsoft.SharePoint.Client.Runtime.dll`

## Run

1. Open `src/SharePoint2010Migration.Web/SharePoint2010Migration.Web.csproj` in Visual Studio.
2. Restore NuGet packages.
3. Run SQL script in your target database.
4. Start app and open `/Admin/Index`.
5. In Admin Settings, configure:
   - SharePoint CSOM DLL absolute paths
   - Site URL + Library title
   - Domain, user, and password
6. Save settings, then open `/Migrate/Index` and run migration.
