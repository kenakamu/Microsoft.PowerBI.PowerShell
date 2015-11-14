# Copyright © Microsoft Corporation.  All Rights Reserved.
# This code released under the terms of the 
# Microsoft Public License (MS-PL, http://opensource.org/licenses/ms-pl.html.)
# Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
# THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
# INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE. 
# We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that. 
# You agree: 
# (i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; 
# (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; 
# and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code 

# Connection Functions
function Connect-PowerBI{

<#
 .SYNOPSIS
 Stores parameters to scrip variable to be used in other functions.

 .DESCRIPTION
 The Connect-PowerBI cmdlet lets you store parameters to scrip variable to be used in other functions.

 .parameter AuthorityName
 Azure Active Directory Name or Guid. i.e.)contoso.onmicrosoft.com

 .parameter ClientId
 A registerered ClientId as application to the Azure Active Directory.

 .parameter UserName
 A username to login to PowerBI.com.

 .parameter Password
 A password for UserName. 

 .parameter GroupId
 A guid of Group. 

 .EXAMPLE
 Connect-PowerBI -AuthorityName contoso.onmicrosoft.com -ClientId bf922382-cdc4-43d4-995c-0f90ecdeda21 -UserName user@contoso.onmicrosoft.com -Password password 

 This examples connects to PowerBI instance of contoso.onmicrosoft.com by using specified UserName/Password.

 .EXAMPLE
 Connect-PowerBI -AuthorityName contoso.onmicrosoft.com -ClientId bf922382-cdc4-43d4-995c-0f90ecdeda21 -UserName user@contoso.onmicrosoft.com -Password password -GroupdId ce88923a-b885-4d11-997a-a240e73fb6b5

 This examples connects to a group of PowerBI instance of contoso.onmicrosoft.com by using specified UserName/Password.
#>

    PARAM(
        [parameter(Mandatory=$true)]
        [string] $AuthorityName,
        [parameter(Mandatory=$true)]
        [string] $ClientId,
        [parameter(Mandatory=$true)]
        [string] $UserName,
        [parameter(Mandatory=$true)]
        [string] $Password,
        [parameter(Mandatory=$false)]
        [string] $GroupId
    )

    # Storing variable to script level
    $script:authorityName = $AuthorityName
    $script:PowerBIClientId = $ClientId
    $script:PowerBIUserName = $UserName
    $script:PowerBIPassword = $Password
    $script:PowerBIResourceId = "https://analysis.windows.net/powerbi/api"
    $script:PowerBIBaseAddress = "https://api.powerbi.com/v1.0/myorg/"
    if($GroupId -ne '')
    {
        $script:PowerBIBaseAddress += "groups/$GroupId/"
    }

    $script:PowerBIheader = @{"Content-Type"="application/json";"Authorization"="Bearer " + (Get_PowerBIAccessToken)} 
}

function Switch-PowerBIContext{

<#
 .SYNOPSIS
 Switches PowerBI context to a group or me.

 .DESCRIPTION
 The Switch-PowerBIContext cmdlet lets you switches PowerBI context to a group or me.

 .parameter GroupId
 A guid of Group. You can get them by running Get-PowerBIGroups once you run Connect-PowerBIApi without GroupId.

 .EXAMPLE
 Switch-PowerBIContext -GroupId ce88923a-b885-4d11-997a-a240e73fb6b5

 This example switches PowerBI address to https://api.powerbi.com/v1.0/myorg/groups/ce88923a-b885-4d11-997a-a240e73fb6b5/

 .EXAMPLE
 Switch-PowerBIContext -Me

 This example switches PowerBI address to https://api.powerbi.com/v1.0/myorg/

#>

    PARAM(
        [parameter(Mandatory=$true, parameterSetName="GroupId")]
        [string] $GroupId,
        [parameter(Mandatory=$true, parameterSetName="Me")]
        [switch] $Me
    )
      
    $script:PowerBIBaseAddress = "https://api.powerbi.com/v1.0/myorg/"
    if($GroupId -ne '')
    {
        $script:PowerBIBaseAddress += "groups/$GroupId/"
    }
    if($Me)
    {
        $script:PowerBIBaseAddress = "https://api.powerbi.com/v1.0/myorg/"
    }
    
    Write-Verbose "Current PowerBI address is $PowerBIBaseAddress."
}

# Private AccessToken obtain function
function Get_PowerBIAccessToken{
    return Get-ADALAccessToken -AuthorityName $authorityName -ClientId $PowerBIClientId `
    -ResourceId $PowerBIResourceId `
    -UserName $PowerBIUserName -Password $PowerBIPassword
}

# DataSet Operations
# https://msdn.microsoft.com/en-us/library/mt203562.aspx
function Add-PowerBIDataSet{
<#
 .SYNOPSIS
 Adds PowerBI dataset.

 .DESCRIPTION
 The Add-PowerBIDataSet cmdlet lets you add PowerBI dataset. You need to include Table definition as well.

 .parameter DataSet
 DataSet object to be created. You need to include Table definition as well.

 .EXAMPLE
 PS C:\>$col1 = New-PowerBIColumn -ColumnName ID -ColumnType Int64
 PS C:\>$col2 = New-PowerBIColumn -ColumnName Data -ColumnType String
 PS C:\>$table1 = New-PowerBITable -TableName SampleTable1 -Columns $col1,$col2
 PS C:\>
 PS C:\>$col3 = New-PowerBIColumn -ColumnName ID -ColumnType Int64
 PS C:\>$col4 = New-PowerBIColumn -ColumnName Date -ColumnType DateTime
 PS C:\>$col5 = New-PowerBIColumn -ColumnName Detail -ColumnType String
 PS C:\>$col6 = New-PowerBIColumn -ColumnName Result -ColumnType Double
 PS C:\>$table2 = New-PowerBITable -TableName SampleTable2 -Columns $col3,$col4,$col5,$col6
 PS C:\>
 PS C:\>$dataset = New-PowerBIDataSet -DataSetName SampleDataSet -Tables $table1,$table2
 PS C:\>
 PS C:\>Add-PowerBIDataSet -DataSet $dataset

 This example instantiate a table with two columns and another table with four columns, and instantiate a dataset.
 Then, it creates the dataset in PowerBI.
#>
    PARAM(
        [parameter(Mandatory=$true)]
        [string]$DataSet
    )
        
    # Send request and return result only.
    $result = Invoke-RestMethod -Method Post -Uri ($PowerBIBaseAddress + "datasets") -Headers $PowerBIheader -Body $DataSet
    return $result.id
}

# https://msdn.microsoft.com/en-us/library/mt203567.aspx
function Get-PowerBIDataSets{
<#
 .SYNOPSIS
 Gets all PowerBI datasets.

 .DESCRIPTION
 The Get-PowerBIDataSets cmdlet lets you retrieve PowerBI datasets for your organization.

 .EXAMPLE
 Get-PowerBIDataSets

 id                                   name               
 --                                   ----               
 4b644350-f745-48dd-821c-f008350199a8 DataSet1
 d77cd0fc-f310-4547-97fa-47c5ccf7f9e1 DataSet2     
 3f08bb1b-4f9e-4be7-939f-750ddbb629de DataSet3
 ...  
#>

    $result = Invoke-RestMethod -Method Get -Uri ($PowerBIBaseAddress + "datasets") -Headers $PowerBIheader
    return $result.value
}

# Table Operations
# https://msdn.microsoft.com/en-us/library/mt203556.aspx
function Get-PowerBITables{
<#
 .SYNOPSIS
 Gets all PowerBI Tables for specified DataSet.

 .DESCRIPTION
 The Get-PowerBIDataSets cmdlet lets you retrieve PowerBI datasets for your organization.

 .parameter DataSetId
 The Id of dataset.

 .EXAMPLE
 Get-PowerBITables -DataSetId 4b644350-f745-48dd-821c-f008350199a8

 name      
 ----      
 PowerBISampleTable1
 PowerBISampleTable2
 ...   
#>
    PARAM(
        [parameter(Mandatory=$true)]
        [string]$DataSetId
    )

    $result = Invoke-RestMethod -Method Get -Uri ($PowerBIBaseAddress + "datasets/$DataSetID/tables") -Headers $PowerBIheader
    return $result.value
}

# https://msdn.microsoft.com/en-us/library/mt203560.aspx
function Update-PowerBITableSchema{
<#
 .SYNOPSIS
 Updates PowerBI Table Schema.

 .DESCRIPTION
 The Update-PowerBITableSchema cmdlet lets you update PowerBI Table Schema.

 .parameter DataSetId
 A DataSetId of the table.

 .parameter TableName
 Updating Table name

 .parameter TableSchema
 A Table object

 .EXAMPLE
 PS C:\>$col1 = New-PowerBIColumn -ColumnName ID -ColumnType Int64
 PS C:\>$col2 = New-PowerBIColumn -ColumnName Data -ColumnType String
 PS C:\>$col3 = New-PowerBIColumn -ColumnName Date -ColumnType DateTime
 PS C:\>$table1 = New-PowerBITable -TableName SampleTable1 -Columns $col1,$col2,$col3
 PS C:\>
 PS C:\>Update-PowerBITableSchema -DataSetId 4b644350-f745-48dd-821c-f008350199a8 -TableName SampleTable1

 This example update SampleTable1 Table Schema with three columns.
#>
    PARAM(
        [parameter(Mandatory=$true)]
        [string]$DataSetId,
        [parameter(Mandatory=$true)]
        [string]$TableName,
        [parameter(Mandatory=$true)]
        [string]$TableSchema

    )
        
    # Send request and return result only.
    $result = Invoke-RestMethod -Method Put -Uri ($PowerBIBaseAddress + "datasets/$DataSetId/tables/$TableName") -Headers $PowerBIheader -Body $TableSchema
}

# Row Operations
# https://msdn.microsoft.com/en-us/library/mt203561.aspx
function Add-PowerBIRows{
<#
 .SYNOPSIS
 Adds Rows to PowerBI table.

 .DESCRIPTION
 The Add-PowerBIRows cmdlet lets you add rows to PowerBI table. 

 .parameter DataSetId
 A ataSet Id which the table resides.

 .parameter TableName
 A Table Name to insert data.

 .parameter Rows
 Actual Data to be inserted. Rows are array of hashtable. i.e.) @{"Column1"="Value1;"Column2"="Value2"}

 .EXAMPLE
 Add-PowerBIRows -DataSetId 4b644350-f745-48dd-821c-f008350199a8 -TableName Table1 -Rows @{"Column1"="Value1;"Column2"="Value2"},@{"Column1"="Value1;"Column2"="Value2"}

 This example inserts two rows to Table1.

 .EXAMPLE
 Add-PowerBIRows -DataSetId 4b644350-f745-48dd-821c-f008350199a8 -TableName Table1 -Rows (Import-Csv -Path ".\data.csv")

 This example inserts rows from CSV to Table1.

#>
    PARAM(
        [parameter(Mandatory=$true)]
        [string]$DataSetId,
        [parameter(Mandatory=$true)]
        [string]$TableName,
        [parameter(Mandatory=$true)]
        [array]$Rows
    )

    $rows = "{'rows': " + (ConvertTo-Json $Rows) + "}"
    $result = Invoke-RestMethod -Method Post -Uri ($PowerBIBaseAddress + "datasets/$DataSetId/tables/$TableName/rows") -Headers $PowerBIheader -Body $rows
    return $result.id
}

# https://msdn.microsoft.com/en-us/library/mt238041.aspx
function Remove-PowerBIRows{
<#
 .SYNOPSIS
 Removes all Rows from PowerBI table.

 .DESCRIPTION
 The Remove-PowerBIRows cmdlet lets you remove rows from PowerBI table. 

 .parameter DataSetId
 DataSet Id which the table resides.

 .parameter TableName
 Table Name to delete rows.

 .EXAMPLE
 Remove-PowerBIRows -DataSetId 4b644350-f745-48dd-821c-f008350199a8 -TableName Table1

 OK.  
#>
    PARAM(
        [parameter(Mandatory=$true)]
        [string]$DataSetId,
        [parameter(Mandatory=$true)]
        [string]$TableName
    )
    
    try
    {
        Invoke-RestMethod -Method Delete -Uri ($PowerBIBaseAddress + "datasets/$DataSetId/tables/$TableName/rows") -Headers $PowerBIheader
    }
    catch
    {
        Write-Warning "Falied to delete rows."
    }
}

# Group Operations
# https://msdn.microsoft.com/en-us/library/mt243842.aspx
function Get-PowerBIGroups{
<#
 .SYNOPSIS
 Gets all PowerBI groups.

 .DESCRIPTION
 The Get-PowerBIGroups cmdlet lets you retrieve PowerBI groups.

 .EXAMPLE
 Get-PowerBIGroups
 id                                   name        
 --                                   ----        
 ce88923a-b885-4d11-997a-a240e73fb6b5 PowerBIGroup

 This example gets groups which current user belongs to.
#>

    $result = Invoke-RestMethod -Method Get -Uri ($PowerBIBaseAddress + "groups") -Headers $PowerBIheader
    return $result.value
}

# Other Util functions
function New-PowerBIDataSet{
<#
 .SYNOPSIS
 Creates New PowerBI dataset object.

 .DESCRIPTION
 The New-PowerBIDataSet cmdlet lets you create New PowerBI dataset object.

 .parameter DataSetName
 DataSet Name.

 .parameter Tables
 An array of Table objects

 .EXAMPLE
 PS C:\>$col1 = New-PowerBIColumn -ColumnName ID -ColumnType Int64
 PS C:\>$col2 = New-PowerBIColumn -ColumnName Data -ColumnType String
 PS C:\>$table1 = New-PowerBITable -TableName SampleTable1 -Columns $col1,$col2
 PS C:\>
 PS C:\>$col3 = New-PowerBIColumn -ColumnName ID -ColumnType Int64
 PS C:\>$col4 = New-PowerBIColumn -ColumnName Date -ColumnType DateTime
 PS C:\>$col5 = New-PowerBIColumn -ColumnName Detail -ColumnType String
 PS C:\>$col6 = New-PowerBIColumn -ColumnName Result -ColumnType Double
 PS C:\>$table2 = New-PowerBITable -TableName SampleTable2 -Columns $col3,$col4,$col5,$col6
 PS C:\>
 PS C:\>$dataset = New-PowerBIDataSet -DataSetName SampleDataSet -Tables $table1,$table2
 PS C:\>
 PS C:\>Add-PowerBIDataSet -DataSet $dataset

 This example instantiate a table with two columns and another table with four columns, and instantiate a dataset.
 Then, it creates the dataset in PowerBI.
#>
    PARAM(
        [parameter(Mandatory=$true)]
        [string]$DataSetName,
        [parameter(Mandatory=$true)]
        [array]$Tables
    )

    return "{'name': '$DataSetName', 'tables': [" + ($Tables -join ",") + "]}"
}

function New-PowerBITable{
<#
 .SYNOPSIS
 Creates New PowerBI table object.

 .DESCRIPTION
 The New-PowerBITable cmdlet lets you create New PowerBI table object.

 .parameter TableName
 Table Name.

 .parameter Columns
 An array of Column objects.

 .EXAMPLE
 PS C:\>$col1 = New-PowerBIColumn -ColumnName ID -ColumnType Int64
 PS C:\>$col2 = New-PowerBIColumn -ColumnName Data -ColumnType String
 PS C:\>$table1 = New-PowerBITable -TableName SampleTable1 -Columns $col1,$col2
 PS C:\>
 PS C:\>$col3 = New-PowerBIColumn -ColumnName ID -ColumnType Int64
 PS C:\>$col4 = New-PowerBIColumn -ColumnName Date -ColumnType DateTime
 PS C:\>$col5 = New-PowerBIColumn -ColumnName Detail -ColumnType String
 PS C:\>$col6 = New-PowerBIColumn -ColumnName Result -ColumnType Double
 PS C:\>$table2 = New-PowerBITable -TableName SampleTable2 -Columns $col3,$col4,$col5,$col6
 PS C:\>
 PS C:\>$dataset = New-PowerBIDataSet -DataSetName SampleDataSet -Tables $table1,$table2
 PS C:\>
 PS C:\>Add-PowerBIDataSet -DataSet $dataset

 This example instantiate a table with two columns and another table with four columns, and instantiate a dataset.
 Then, it creates the dataset in PowerBI.
#>
    PARAM(
        [parameter(Mandatory=$true)]
        [string]$TableName,
        [parameter(Mandatory=$true)]
        [array]$Columns
    )

    return "{'name': '$TableName', 'columns': [" + ($Columns -join ",") + "]}"
}

function New-PowerBIColumn{
<#
 .SYNOPSIS
 Creates new PowerBI column object.

 .DESCRIPTION
 The New-PowerBIColumn cmdlet lets you create new PowerBI column object.

 .parameter ColumnName
 A column name.

 .parameter ColumnType
 A type of the column. Type can be String, Int64, DateTime, Boolean, or Double

 .EXAMPLE
 PS C:\>$col1 = New-PowerBIColumn -ColumnName ID -ColumnType Int64
 PS C:\>$col2 = New-PowerBIColumn -ColumnName Data -ColumnType String
 PS C:\>$table1 = New-PowerBITable -TableName SampleTable1 -Columns $col1,$col2
 PS C:\>
 PS C:\>$col3 = New-PowerBIColumn -ColumnName ID -ColumnType Int64
 PS C:\>$col4 = New-PowerBIColumn -ColumnName Date -ColumnType DateTime
 PS C:\>$col5 = New-PowerBIColumn -ColumnName Detail -ColumnType String
 PS C:\>$col6 = New-PowerBIColumn -ColumnName Result -ColumnType Double
 PS C:\>$table2 = New-PowerBITable -TableName SampleTable2 -Columns $col3,$col4,$col5,$col6
 PS C:\>
 PS C:\>$dataset = New-PowerBIDataSet -DataSetName SampleDataSet -Tables $table1,$table2
 PS C:\>
 PS C:\>Add-PowerBIDataSet -DataSet $dataset

 This example instantiate a table with two columns and another table with four columns, and instantiate a dataset.
 Then, it creates the dataset in PowerBI.
#>
    PARAM(
        [parameter(Mandatory=$true)]
        [string]$ColumnName,
        [parameter(Mandatory=$true)]
        [ValidateSet("String","Int64","DateTime","Boolean","Double")]
        [string]$ColumnType
    )
    
    return "{'name':'$ColumnName','dataType':'$ColumnType'}"
}
