# SharePoint 2010 Migration Prototype (Code)

## Structure

- `src/SharePoint2010Migration.Web` - ASP.NET MVC (.NET Framework) web app.
- `src/SharePoint2010Migration.Runner` - console executable project.
- `database/create-document-archive.sql` - SQL table setup.
- `lib/sp2010` - place SharePoint 2010 CSOM DLLs here.

## Required DLLs

Copy these into `lib/sp2010`:

- `Microsoft.SharePoint.Client.dll`
- `Microsoft.SharePoint.Client.Runtime.dll`

## Web app flow

1. Open `Sharepoint2010.sln` in Visual Studio.
2. Restore NuGet packages.
3. Run SQL script in your target database.
4. Start app and open `/Admin/Index`.
5. In Admin Settings, configure:
   - SharePoint CSOM DLL absolute paths
   - Site URL + Library title
   - Domain, user, and password
6. Save settings, then open `/Migrate/Index` and run migration.

## Executable (.exe) flow

1. Build from Visual Studio (Release) **or** run `build-runner.bat` on Windows.
2. The executable is generated at:
   - `src\SharePoint2010Migration.Runner\bin\Release\SharePoint2010Migration.Runner.exe`
3. Ensure admin settings were already saved by the web app (`App_Data/admin-settings.json`), or pass path as argument:
   - `SharePoint2010Migration.Runner.exe "..\SharePoint2010Migration.Web\App_Data\admin-settings.json"`
