
<img align="left" src="Images/ReadMe/Portfolio.gif" width="250px">

<a href="https://social.msdn.microsoft.com/profile/aduguid/"><img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/msdn_profile2.png#3" width="150" valign="top" align="right" alt="MSDN Profile">
</a><a href="https://stackexchange.com/users/12441351/aduguid?tab=accounts"><img src="https://stackexchange.com/users/flair/12441351.png#2" width="175" height="50" alt="profile for aduguid on Stack Exchange, a network of free, community-driven Q&amp;A sites" title="profile for aduguid on Stack Exchange, a network of free, community-driven Q&amp;A sites" valign="top" align="right"/></a><a href="https://social.msdn.microsoft.com/profile/aduguid/"></a>


<br>
<br>
<br>
<br>
<br>
<br>

<table style="width:100%">
    <caption>
        <h2>
            Table of Contents
        </h2>
    </caption>

<tr valign="top">
<td align="left" valign="middle">
<img align="center" src="Images/ReadMe/autocad.png" width="32px">
<a href="#autocad">AutoCAD</a>
</td>
</tr>

<tr valign="top">
<td align="left" valign="middle">
<img align="center" src="Images/ReadMe/office.png" width="32px">
<a href="#office-addins">Office</a>
</td>
</tr>

<tr valign="top">
<td align="left" valign="middle">
<img align="center" src="Images/ReadMe/ssrs.png" width="32px">
<a href="#ssrs-reports">Reporting Services</a>
</td>
</tr>

<tr valign="top">
<td align="left" valign="middle">
<img align="center" src="Images/ReadMe/visualstudio.png" width="32px">
<a href="#visual-studio">Visual Studio</a>
</td>
</tr>

</table>

<br>
<br>
<br>
<br>

<html lang="en" xmlns="http://www.w3.org/1999/xhtml">

<a id="user-content-autocad" class="anchor" href="#autocad" aria-hidden="true"> </a>
<table style="width:100%">
    <caption>
        <h2>
            <img align="left" src="Images/ReadMe/autocad.png" width="32px">
            AutoCAD Projects
        </h2>
    </caption>
    <tr valign="top">
        <td width="33%">
            <kbd>
                AutoCAD Automation Plotting Tool
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/csharpwpfautocadbatchutilityprocess.png" align="top" width="256px" title="AutoCAD Automation Plotting Tool" />
                <br>
                <br>
                <span style="max-width:256px;">Written in VB.NET, Windows Presentation Foundation (WPF), it allows the user to batch plot AutoCAD files with options to convert drawings to pdf files. Plot styles are applied to the drawings during the conversion. Logging is implemented with Log4Net to track errors during the conversion of 200,000 files. There was less than a 0.5% failure in the process. The processing time took 3 virtual machines 5 days to complete. The median processing time was 10 seconds per file. It can also send the files to a plotter with a selected page size in the active space either model or paper. There is also an option to add or remove a layered watermark using a .dwt file. </span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                AutoCAD Automation Extracting Tool
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/csharpwpfautocadbatchutilityextract.png" align="top" width="256px" title="AutoCAD Automation Extracting Tool" />
                <br>
                <br>
                <span style="max-width:256px;">This applicaiton allows the user to loop through each entity in a piping and instrumentation drawing to extract entities and attribute values. Depending on the parameter options entities are extracted within clouding for new project work or no clouding as with as building at project completion. The application uses stored layer names to determine the cloud types. A standard list of entities and attributes are used to query the extracted data for reporting in SSRS. There is an option in the settings to use either a SQL Server localDB or a server database. SQL Server localDB is used for the portable version. The original project was migrated from VBA/dvb project file. This application is in Visual Studio 2015 in VB .NET/WPF using the AutoCAD 2016 object model. The procedures are multi-threaded using background workers. The application uses late binding for a current session of AutoCAD and early binding for a new session. ClickOnce deployment is used for the install. The API documentation is done with Microsoft Sandcastle. The application diagrams are done with Microsoft Visio. The “As Is” documentation is done in a markdown file..</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                Title Block Tool
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/dvb.titleblock.tool2.png" align="top" width="256px" title="Title Block Tool" />
                <br>
                <br>
                <span style="max-width:256px;">Title Block Tool was developed in a compiled VBA/dvb file with the database in Microsoft Access & Oracle. Microsoft Access is used for the portable version. The custom AutoCAD menu is in DIESEL. It extracts title blocks from multiple files by using standardized naming conventions for the title block names. The title block and attribute names are stored in a template table and referenced during the extraction process. A custom edit attributes form is used to select values from dropdown lists to improve data quality. The application also allows the users to bulk update the title blocks from the database back to the original files. The modules and classes are exported with a procedure for source control check-in. The install was done in Wise Installer.</span>
                <br>
            </kbd>
        </td>
    </tr>
    
</table>
<br>
<br>
<br>

<a id="user-content-office-addins" class="anchor" href="#office-addins" aria-hidden="true"> </a>
<table style="width:100%">
    <caption>
        <h2>
            <img align="left" src="Images/ReadMe/office.png" width="32px">
            <a href="https://github.com/Office-projects">Office Projects</a>
        </h2>
    </caption>
    <tr valign="top">
        <td width="33%">
            <kbd>
                <a href="https://github.com/Office-projects/Script-Help/blob/master/README.md">Microsoft Excel Script Help (VSTO)</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/vsto.excel.ribbon.scripthelp.gif" align="top" width="256px" title="Microsoft Excel Script Help" />
                <br>
                <br>
                <span style="max-width:256px;">Written in C#, VB.NET and VBA, it allows the user to use a table to create different SQL scripts.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
              <kbd>
                <a href="https://github.com/Office-projects/Server-Actions/blob/master/README.md">Microsoft Excel Server Help (VSTO)</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/vsto.excel.ribbon.serverhelp.gif" align="top" width="256px" title="Microsoft Excel Server Help" />
                <br>
                <br>
                <span style="max-width:256px;">Written in C#, VB.NET and VBA to ping servers as well as create a Microsoft Remote Desktop Manager file.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
              <kbd>
                Microsoft Excel Engineering Markup (VSTO)
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/vsto.excel.ribbon.markup.gif" align="top" width="256px" title="Microsoft Excel Engineering Markup" />
                <br>
                <br>
                <span style="max-width:256px;">Written in C#, VB.NET and VBA, it allows the user to mark up files with revision clouding around the selected cells in different colors.</span>
                <br>
            </kbd>
        </td>
    </tr>
    <tr valign="top">
        <td width="33%">
            <kbd>
                <a>Microsoft Excel Cell Extract Addin (VBA)</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/vba.excel.ribbon.cellextract.gif" align="top" width="256px" title="Microsoft Excel Cell Extract Addin" />
                <br>
                <br>
                <span style="max-width:256px;">Written in VBA, this project was done as an Excel Addin with a ribbon to import enrollment data from 500 files. The median import time is 2 seconds per file. I used a column mapping to the cell references in multiple sheets to generate the import specification.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                <a href="https://github.com/Office-projects/Excel-Timesheet/blob/master/README.md">Microsoft Excel Timesheet Ribbon (VBA)</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/Office-projects/Excel-Timesheet/master/Images/ReadMe/screenshot.png" align="top" width="256px" title="Microsoft Excel Timesheet Ribbon" />
                <br>
                <br>
                <span style="max-width:256px;">This Add-In is used to produce a timesheet file with functionality to import your Google Timeline.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                <a href="https://github.com/Office-projects/Excel-Favorites/blob/master/README.md">Microsoft Excel Favorites Ribbon (VSTO)</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/vsto.excel.ribbon.favorites.gif" align="top" width="256px" title="Microsoft Excel Favorites Ribbon" />
                <br>
                <br>
                <span style="max-width:256px;">Written in C#, VB.NET and VBA, it allows the user to create a custom Excel favorites ribbon.</span>
                <br>
            </kbd>
        </td>
    </tr>
    <tr valign="top">
        <td width="33%">
            <kbd>
                <a href="https://github.com/Office-projects/Access-examples/blob/master/README.md">Microsoft Access VBA Examples</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/vba.access.examples.png" align="top" width="256px" title="Microsoft Access VBA Examples" />
                <br>
                <br>
                <span style="max-width:256px;">Written in VBA, here are examples of reports and code for MS Access.</span>
                <br>
            </kbd>
        </td>
                </td>
        <td width="33%">
            <kbd>
                <a href="https://github.com/Office-projects/Visio-Shape-Extract/blob/master/README.md">Microsoft Visio Shape Export (VSTO)</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/vsto.visio.shape.extract2.png" align="top" width="256px" title="Microsoft Visio Shape Export" />
                <br>
                <br>
                <span style="max-width:256px;">Written in C#, VB.NET and VBA, it allows the user to export the shape attributes to a .csv file.</span>
                <br>
            </kbd>
        </td>
     </tr>
</table>

<br>
<br>
<br>

<a id="user-content-ssrs-reports" class="anchor" href="#ssrs-reports" aria-hidden="true"> </a>
<table style="width:100%">
    <caption>
        <h2>
            <img align="left" src="Images/ReadMe/ssrs.png" width="32px">
            <a href="https://github.com/SQL-Server-projects">Reporting Services Projects</a>
        </h2>
    </caption>
    <tr valign="top">
            <td width="33%">
            <kbd>
                <a href="https://gist.githubusercontent.com/aduguid/4905cd812ef2c86ad3d026be852c62ad/raw/deff6b7cd79f2b729aa31e62795cfa32956cf4f3/SSRS.heatmap_example">Heat Map Calendar</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrsheatmap_calendar.png" align="top" width="256px" title="Heat Map Calendar" />
                <br>
                <br>
                <span style="max-width:256px;">This report is a heat map calendar of events. It uses a variable for the base color.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                Geospatial Map example
                <br>
                <br>
                <img align="center" src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrs.geospatial.map.png" width="256px" title="Geospatial Map" />
                <br>
                <br>
                <span style="max-width:256px;">This is an example of a geospatial map.</span>
                <br>
            </kbd>
        </td>
        <td width="33%" align="center" >
             <kbd>
                Mobile Report Publisher example
                <br>
                <br>
                <img align="center" src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrs.mobile.report.png" width="64px" title="Mobile Report Publisher" />
                <br>
                <br>
                <span style="max-width:256px;">This is an example from SQL Server Mobile Report Publisher.</span>
                <br>
            </kbd>
        </td>
            </tr>
    <tr valign="top">
        <td width="33%">
            <kbd>
                <a href="https://raw.githubusercontent.com/aduguid/SqlServerReportingServices/master/ServerDatabase/Database%20Dictionary.rdl">Data Dictionary</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrsdatadictionary.png" align="top" width="256px" title="Data Dictionary" />
                <br>
                <br>
                <span style="max-width:256px;">This report is used for querying the data dictionary of a SQL Server database.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                <a href="https://raw.githubusercontent.com/aduguid/SqlServerReportingServices/master/ServerDatabase/Scheduled%20Jobs.rdl">Scheduled Jobs Gantt Chart</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrsscheduledjobs.png" align="top" width="256px" title="Scheduled Jobs Gantt Chart" />
                <br>
                <br>
                <span style="max-width:256px;">This report is used for querying the scheduled jobs for a SQL Server database.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                <a href="https://github.com/aduguid/SqlServerReportingServices/blob/master/ServerDatabase/Activity%20Moniter.rdl">Activity Monitor</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrsactivitymonitor.png" align="top" width="256px" title="Activity Monitor" />
                <br>
                <br>
                <span style="max-width:256px;">This report queries the activity monitor from SQL Server.</span>
                <br>
            </kbd>
        </td>
    </tr>
    <tr valign="top">
        <td width="33%">
            <kbd>
                <a href="https://raw.githubusercontent.com/aduguid/SqlServerReportingServices/master/ServerReports/Report%20List.rdl">Report Listing</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrsreportlisting.png" align="top" width="256px" title="Report Listing" />
                <br>
                <br>
                <span style="max-width:256px;">This report is used for querying the deployed SSRS reports, their subscriptions and their execution logs.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                <a href="https://raw.githubusercontent.com/aduguid/SqlServerReportingServices/master/ServerReports/Subscriptions.rdl">Report Subscriptions</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrsreportsubscriptions.png" align="top" width="256px" title="Report Subscriptions" />
                <br>
                <br>
                <span style="max-width:256px;">This report is used for querying the deployed SSRS subscriptions.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                <a href="https://github.com/aduguid/SqlServerReportingServices/blob/master/ServerReports/Execution%20Log.rdl">Report Execution Log</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrsreportexecutionlog.png" align="top" width="256px" title="Report Execution Log" />
                <br>
                <br>
                <span style="max-width:256px;">This report is used for querying the report server exection log table.</span>
                <br>
            </kbd>
        </td>
    </tr>
    <tr valign="top">
        <td width="33%">
            <kbd>
                <a href="https://github.com/aduguid/SqlServerReportingServices/blob/master/ServerReports/Report%20Documentation.rdl">Report Documentation</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrsreportdocumentation.png" align="top" width="256px" title="Report Documentation" />
                <br>
                <br>
                <span style="max-width:256px;">This report is used for querying the deployed report XML.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
               <a href="https://github.com/aduguid/SqlServerReportingServices/blob/master/ServerReports/Data%20Sources.rdl">Report Data Sources</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrsreportdatasources.png" align="top" width="256px" title="Report Data Sources" />
                <br>
                <br>
                <span style="max-width:256px;">This report is used for querying the deployed data sources.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                <a href="https://github.com/aduguid/SqlServerReportingServices/blob/master/ServerReports/Permissions.rdl">Report Server Folder Permissions</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrsreportfolderpermissions.png?" align="top" width="256px" title="Report Server Folder Permissions" />
                <br>
                <br>
                <span style="max-width:256px;">This report is used for querying the report server folder permissions.</span>
                <br>
            </kbd>
        </td>
    </tr>
    <tr valign="top">
        <td width="33%">
            <kbd>
                S Curve Cumulative Progress
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrsscurve.png" align="top" width="256px" title="S Curve Cumulative Progress" />
                <br>
                <br>
                <span style="max-width:256px;">This report is used for querying the cumulative hours or sales quantities plotted against time. It is dynamicly grouped by either quarter, month or weekending.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                Weekly Calendar
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/ssrsnestedtablixmatrixcalendar.png" align="top" width="256px" title="Weekly Calendar" />
                <br>
                <br>
                <span style="max-width:256px;">This report uses a nested tablix inside of a matrix to create a calendar view.</span>
                <br>
            </kbd>
        </td>
    </tr>
    </tr>
</table>
<br>
<br>
<br>

<a id="user-content-visual-studio" class="anchor" href="#visual-studio" aria-hidden="true"> </a>
<table style="width:100%">
    <caption>
        <h2>
            <img align="left" src="Images/ReadMe/visualstudio.png" width="32px">
            <a href="https://github.com/Visual-Studio-projects">Visual Studio Projects</a>
        </h2>
    </caption>
    <tr valign="top">
        <td width="33%">
            <kbd>
                <a href="https://github.com/aduguid/ActiveDirectorySearch/blob/master/README.md">Active Directory Search</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/csharpwinformactivediretorysearch2.png" align="top" width="256px" title="Active Directory Search" />
                <br>
                <br>
                <span style="max-width:256px;">Written in C#/VB.NET, Windows Forms, it allows the user to search Active Directory and save the results to a .csv file</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                <a href="https://github.com/aduguid/DocumentumScriptAdministrator/blob/master/README.md">Documentum Script Administrator</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/csharpwpfdocumentumscriptadministrator.png" align="top" width="256px" title="Documentum Script Administrator" />
                <br>
                <br>
                <span style="max-width:256px;">Written in C#/VB.NET, Windows Presentation Foundation (WPF), it allows the user to run a file as one command against the idql32/64.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                <a href="https://raw.githubusercontent.com/aduguid/SharePointWebparts/master/Timezones_with_Temp.dwp">SharePoint Webpart Timezones & Temperature</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/sharepointwebparttimezoneweatherhyperlink.gif" align="top" width="256px" title="SharePoint Webpart Timezones with Temperature" />
                <br>
                <br>
                <span style="max-width:256px;">Written in HTML/Javascript, it displays timezones with hyperlinks to temperature. The time updates every second.</span>
                <br>
            </kbd>
        </td>
    </tr>
    <tr valign="top">
        <td width="33%">
            <kbd>
                <a>Visual Studio Team Services Dashboard (Report Project)</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/vsts.dashboard.reportsproject.png" align="top" width="256px" title="Visual Studio Team Services Dashboard (Report Project)" />
                <br>
                <br>
                <span style="max-width:256px;">Dashboard for SSRS report solution project in Visual Studio Team Services.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                <a>Visual Studio Team Services Dashboard (All Projects)</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/vsts.dashboard.allprojects.png" align="top" width="256px" title="Visual Studio Team Services Dashboard (All Projects)" />
                <br>
                <br>
                <span style="max-width:256px;">Dashboard for all projects in Visual Studio Team Services.</span>
                <br>
            </kbd>
        </td>
        <td width="33%">
            <kbd>
                <a href="https://github.com/aduguid/SoftwarePortfolio/edit/master/README.md">Mark Up/Mark Down Example</a>
                <br>
                <br>
                <img src="https://raw.githubusercontent.com/aduguid/SoftwarePortfolio/master/Images/ReadMe/markupmarkdownexample.png" align="top" width="256px" title="Mark Up/Mark Down Example" />
                <br>
                <br>
                <span style="max-width:256px;">Written in HTML, it displays this page.</span>
                <br>
            </kbd>
        </td>
    </tr>

</table>
<br>
<br>
<br>

</html>
