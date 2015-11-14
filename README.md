# Microsoft.PowerBI.PowerShell
**this module requires Microsoft.ADAL.PowerShell module, which you can find below.<br/>
https://github.com/kenakamu/Microsoft.ADAL.PowerShell/

### Overview 
With Microsoft.PowerBI.PowerShell module, it is so easy to consume PowerBI now!<br/>
Refer to following link for PowerBI API details.<br/>
https://msdn.microsoft.com/en-us/library/mt147898.aspx

**Where do I find the latest relase?**
Releases are found on the [Release Page](https://github.com/kenakamu/Microsoft.PowerBI.PowerShell/releases)

###How to setup modules
<p>1. Download Microsoft.PowerBI.Powershell.zip.</p> 
<p>2. Right click the downloaded zip file and click "Properties". </p> 
<p>3. Check "Unblock" checkbox and click "OK", or simply click "Unblock" button depending on OS versions. </p> 
![Image of Unblock](https://i1.gallery.technet.s-msft.com/powershell-functions-for-16c5be31/image/file/142582/1/unblock.png)
<p>4. Extract the zip file and copy "Microsoft.ADAL.PowerShell" folder to one of the following folders:<br/>
  * %USERPROFILE%\Documents\WindowsPowerShell\Modules<br/>
  * %WINDIR%\System32\WindowsPowerShell\v1.0\Modules<br/>
<p>5. You may need to change Execution Policy to load the module. You can do so by executing following command. </p> 
```PowerShell
 Set-ExecutionPolicy –ExecutionPolicy RemoteSigned –Scope CurrentUser
```
Please refer to 
[Set-ExecutionPolicy](https://technet.microsoft.com/en-us/library/ee176961.aspx) 
for more information.
<p>6. Open PowerShell and run following command to load the module. </p> 
```PowerShell
# Import Micrsoft.PowerBI.Powershell module 
Import-Module Microsoft.PowerBI.Powershell
```

####Example:Connect to PowerBI
First of all, you connect to PowerBI by using Connect-PowerBI function. <br/>
You need to register an application to Azure Active Directory before to obtain ClientId.
```PowerShell
Connect-PowerBI -AuthorityName contoso.onmicrosoft.com -ClientId bf922382-cdc4-43d4-995c-0f90ecdeda21 `
-UserName user@contoso.onmicrosoft.com -Password password
```

If you want to connect to Group rather than your own, use Switch-PowerBIContext function to swtich it.
```PowerShell
Switch-PowerBIContext -GroupId ce88923a-b885-4d11-997a-a240e73fb6b5
```
To back to your own context, use -Me switch
```PowerShell
Switch-PowerBIContext -Me
```
####Example: Create DataSet
When you want to create your own PowerBI dataset, you need to define tables and its schema at the same time.
```PowerShell
# Define columns
$col1 = New-PowerBIColumn -ColumnName ID -ColumnType Int64
$col2 = New-PowerBIColumn -ColumnName Data -ColumnType String
# Define table with two columns
$table1 = New-PowerBITable -TableName SampleTable1 -Columns $col1,$col2

# Define more columns
$col3 = New-PowerBIColumn -ColumnName ID -ColumnType Int64
$col4 = New-PowerBIColumn -ColumnName Date -ColumnType DateTime
$col5 = New-PowerBIColumn -ColumnName Detail -ColumnType String
$col6 = New-PowerBIColumn -ColumnName Result -ColumnType Double
# Define another table with four columns
$table2 = New-PowerBITable -TableName SampleTable2 -Columns $col3,$col4,$col5,$col6

# Define dataset with two tables
$dataset = New-PowerBIDataSet -DataSetName SampleDataSet -Tables $table1,$table2

# Create the dataset
Add-PowerBIDataSet -DataSet $dataset
```
####Example: Update Schema
In case you need to update existing table's schema, use Update-PowerBITableSchema.
```PowerShell
# Define columns
$col1 = New-PowerBIColumn -ColumnName ID -ColumnType Int64
$col2 = New-PowerBIColumn -ColumnName Data -ColumnType String
$col3 = New-PowerBIColumn -ColumnName Date -ColumnType DateTime
# Define table with three columns, You need to use same table name
$table1 = New-PowerBITable -TableName SampleTable1 -Columns $col1,$col2,$col3
# Update the table schema
Update-PowerBITableSchema -DataSetId 4b644350-f745-48dd-821c-f008350199a8 -TableName SampleTable1
```
####Example: Insert rows
Once you define table, it's time to insert rows! You can insert rows from script or from csv.
```PowerShell
# Insert rows inline.
Add-PowerBIRows -DataSetId 4b644350-f745-48dd-821c-f008350199a8 -TableName SampleTable1 `
-Rows @{"ID"=1;"Data"="1"},@{"ID"=2;"Data"="2"}

# Insert rows from CSV which has same schema as the table
Add-PowerBIRows -DataSetId 4b644350-f745-48dd-821c-f008350199a8 -TableName SampleTable1 `
-Rows (Import-Csv -Path ".\data.csv")
```
###How to get command details
Each command has detail explanation.
<p>Run following command to get all commands.</p>
```PowerShell
Get-Command -Module Microsoft.PowerBI.PowerShell
```
<p>Run following command to get help.</p>
```PowerShell
Get-Help Add-PowerBIRows -Detailed
```
