# PBI Automated Checks

This Power BI external tool runs automated checks on a report and generates a list of exceptions to warn you of potential problems.   
![PBIAutomatedChecks](https://github.com/flavio-meneses/PBIAutomatedChecks/assets/59139011/82a2a772-be75-4d48-9376-13e4de4c36fe)

The checks currently supported are:

1. Page Structure
   1. Have filters been applied and forgotten in the report, page or visuals?
   2. Are the page sizes default, or have they been changed?
   3. Is the report landing page the first page? (first page user sees)  
   
2. Data - these are 100% flexible and can be built with DAX Studio. Examples:
   1. Check the row count of tables you know shouldn’t change (e.g. Calendar table);
   2. Look for specific figures you’re expecting (e.g. max date, sum of sales for FY19/20, etc);
   3. Define % variance criteria and flag outliers (e.g. forecast is up 10% from last month, needs investigation);
   4. Compare data source figures vs Power BI figures to check if any discrepancies were introduced during data processing;
   5. Any other applicable to your specific report!


## Requirements

To run the "Page Structure" checks you'll need Admin rights to the machine to install the tool.  

To run the "Data" checks you'll need Global Administrator/Application Administrator role in Azure AD and Power BI Admin role to setup a [Service Principal](https://learn.microsoft.com/en-us/power-bi/developer/embedded/embed-service-principal) (detailed instructions coming soon). If you don't want, or are unable to set this up, you can disable these checks. You will also need a Power BI Pro Account to run these "Data" checks. 


## Install
1. Launch PowerShell as an Administrator

<img src="https://user-images.githubusercontent.com/126802604/228036208-c79594dd-db5d-4768-8b4c-c3e59adad06c.png" width="50%">

2. Copy/paste the following command and press enter:

```
Invoke-Expression -Command (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/flavio-meneses/PBIAutomatedChecks/main/Install.ps1" | Select-Object -ExpandProperty Content)
```

This will download the required files to your "Downloads" folder and install them on your machine.  

After installation, the tool will be visible when you open the "External Tools" ribbon in Power BI Desktop. If you had Power BI open during installation please restart it.  
<img src="https://user-images.githubusercontent.com/59139011/239099639-6883d4ea-f13a-4c5b-b4d4-1d4e2b19dc7a.png" width="50%">

## Configure checks

1. Navigate to "C:\PowerBI_AutomatedChecks"  
2. Open the "Settings.json" file and replace the placeholders with the variables required. Please make sure to include double backslashes \\\ for any file or folder paths;
3. If you're not running the "Data" checks you can ignore the "systemSettings" and should make the propery "runDAXtest" 0 for each of the reports; 
4. If you are running the "Data" checks, go to each of the "Tests" folder and update the "DAXTests.dax" file with the Data checks you want for that report.  

## Use tool

1. In Power BI, go to the "External tools" menu and click the "Automated Checks" button. This will run the checks and generate an automated report that will be saved to your desktop.  
2. If there's a check you don't want to run, you can go to the "Tests" folder for that report and add it to the "IgnoreTests.json" file. For example:
```
[
  {
    "Check Type": "File Structure",
    "Check Name": "Page dimensions for page 'Page 1'"
  },
  {
    "Check Type": "File Structure",
    "Check Name": "Report landing page is first page"
  }
]
```
