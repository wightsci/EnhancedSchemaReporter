<#
.SYNOPSIS

Reports on the Active Directory Domain Services (ADDS) schema for an ADDS forest.

.DESCRIPTION

The Simple Schema Reporter is designed to connect to the schema of the current ADDS forest
and report on the attributes of the class(es) specified by the ClassName parameter, or provide a
list of classes using the ListClasses parameter.

There is no requirement for the Active Directory cmdlets to be present on the machine running this
script as .Net is used directly.

Reports are available as HTML, XML or CSV files or can be output directly onto the Clipboard as an HTML table.

An option is available to immediately view the report(s) generated using the ViewOutput parameter.

The HTMLFile output contains a JavaScript function to sort the results of the table by any column heading.
Please be aware the the performance of the sorting script is poor, so please bear with it whilst it sorts.
Many of the schema entries have ~400 properties!

Find the current version on GitHub at: https://github.com/wightsci/EnhanceSchemaReporter

.PARAMETER ListClasses
This switch parameter specifies that the report will be a list of all classes available,
rather than details about one class.

.PARAMETER ClassName
Specifies the Schema Class(es) to report on. The default class name is User. This parameter 
acccepts a comma-separated list. The list of class names is not validated before an attempt
is made to obtain schema information.

.PARAMETER ReportType
Specifies the type of report. This can be one or more of:
    HTMLFile
    HTMLClipboard
    CSVFile
    XMLFile

The default report type is HTMLFile. Multiple report formats should be comma-separated.

.PARAMETER ReportName
Specifies the file name to be used for the report. If no ReportName parameter is
provided, a system generated file name is used, based on the date, time and schema class.
If a ReportName parameter is provided then the schema class name is appended.

.PARAMETER ViewOutput
Specifies whether the report is displayed using the default application
for its file type. The default value is False.

.EXAMPLE
    SimpleSchemaReporter.ps1

Runs SimpleSchemaReporter with all defaults. An HTML report is generated for the User class
with a system generated name.

.EXAMPLE
    SimpleSchemaReporter.ps1 -ListClasses

Runs SimpleSchemaReporter to list classes available in ADDS. An HTML report is generated.

.EXAMPLE
    SimpleSchemaReporter.ps1 -ClassName Computer

An HTML report is generated for the Computer class with a system generated name.

.EXAMPLE
    SimpleSchemaReporter.ps1 -ClassName Computer -ViewOutput

An HTML report is generated for the Computer class with a system generated name, the report
is displayed in the user's default HTML viewer.

.EXAMPLE
    SimpleSchemaReporter.ps1 -ClassName Computer,User,Contact -ReportType HTML,CSV

An HTML report and a CSV report is generated for the Computer, User and Contact classes with system generated names.



.INPUTS

None. You cannot pipe objects to SimpleSchemaReporter.ps1

#>
Param (
    [Parameter(Mandatory=$False,ParameterSetName='ForClass')]
    [String[]]
    $ClassName = 'User',
    [Parameter(Mandatory=$True,ParameterSetName='ListClass')]
    [Switch]
    $ListClasses,
    [Parameter(Mandatory=$False,ParameterSetName='ForClass')]
    [Parameter(Mandatory=$False,ParameterSetName='ListClass')]
    [ValidateSet('HTMLFile','HTMLClipboard','XMLFile','CSVFile')]
    [String[]]
    $ReportType = 'HTMLFile',
    [Parameter(Mandatory=$False,ParameterSetName='ForClass')]
    [Parameter(Mandatory=$False,ParameterSetName='ListClass')]
    [String]
    $ReportName,
    [Parameter(Mandatory=$False,ParameterSetName='ForClass')]
    [Parameter(Mandatory=$False,ParameterSetName='ListClass')]
    [Switch]
    $ViewOutput = $False 
)
<#
Tested on Windows 10/Server 2019
#>
$oldstylesheet = @"
<style>
body {
    font-family: Calibri, Segoe, Arial, Sans-Serif;
    font-size: 9pt;
}
table {
    border-collapse: collapse;
    border-color: RoyalBlue;
    border-style: solid;
    width: 100%;
}
th, td {
    text-align: left;
    padding: 0.5em;
    border-collapse: collapse;
    border-color: RoyalBlue;
    border-style: solid;
    border-width: 1pt;
    width: 100%;
}
th {
    background-color: SteelBlue;
    color: White;
    cursor: pointer;
}
tr:nth-child(odd) {
    background-color: CornflowerBlue;
}
h1 {
    color: SteelBlue;
}
.reportfooter {
    font-size: 8pt;
    color: SteelBlue;
}
</style>
"@

$stylesheet = @"
<style>
h1 {
color: skyblue;
}
body {
font-family: Calibri, Segoe, Arial, Sans-Serif;
font-size: 9pt;
}
table#myTable tr:nth-child(even),  table#myTable th  {
background-color: lightskyblue;
}
table#myTable th {
color: white
}
</style>
"@

#JQuery and DataTables
$htmlScript = @'
<link rel='Stylesheet' type='text/css' href="https://cdn.datatables.net/1.10.20/css/jquery.dataTables.min.css"></link>
<script src="http://code.jquery.com/jquery-3.4.1.min.js" integrity="sha256-CSXorXvZcTkaix6Yvo6HppcZGetbYMGWSFlBw8HfCJo=" crossorigin="anonymous">
</script>
<script src='https://cdn.datatables.net/1.10.20/js/jquery.dataTables.min.js'></script>
<script>
$(document).ready( function () {
$('#myTable').DataTable();
} );
</script> 
'@ 

function classprops ($classname) {
#Getting the properties of the passed-in class
    #Using a Generic.List to avoid += overheads
    $classproperties = New-Object System.Collections.Generic.List[System.Object]
    $class = $schema.FindClass($classname)
    foreach ($property in $class.MandatoryProperties) {
        $property | Add-Member -MemberType NoteProperty -Name "Mandatory" -Value $True
        if  (@($constrattrnames ) -contains $property.Name) {
            $property | Add-Member -MemberType NoteProperty -Name "Constructed" -Value $True
        }
        else {
            $property | Add-Member -MemberType NoteProperty -Name "Constructed" -Value $False
        }
        $classproperties.Add($property)   
    }
    foreach ($property in $class.OptionalProperties) {
        $property | Add-Member -MemberType NoteProperty -Name "Mandatory" -Value $False
        if  (@($constrattrnames ) -contains $property.Name) {
            $property | Add-Member -MemberType NoteProperty -Name "Constructed" -Value $True
        }
        else {
            $property | Add-Member -MemberType NoteProperty -Name "Constructed" -Value $False
        }
        $classproperties.Add($property)   
    }
    Return $classproperties
}

function findcontsrattr {
   #Finding constructed Attributes
   $searcher = New-Object System.DirectoryServices.DirectorySearcher
   $searcher.SearchRoot = $schema.GetDirectoryEntry()
   $searcher.Filter = '(&(systemFlags:1.2.840.113556.1.4.803:=4)(objectClass=attributeSchema))'
   $constrattrs = $searcher.FindAll().GetDirectoryEntry()
   Return ($constrattrs | Select-Object -Expand ldapDisplayName)
}

filter selectattributes {
    $_ | Select-Object Name,CommonName,OID,Syntax,Mandatory,Constructed,IsSingleValued,IsInAnr,RangeLower,RangeUpper,Link,LinkId
}

filter selectclasses {
    $_.FindAllClasses() | Select-Object Name,CommonName,subClassOf
}

function addjscript($html) {
    #Modify HTML to add JQuery and format table (remove colgroup, add thead and tbody)
    $htmldata = [xml]($($html))
    $htmlHeadTag = $htmldata.html.head ; 
    $htmlTableTag = $htmldata.html.body.table
    #Remove COLGROUP tag
    $htmlTableTag.RemoveChild($htmlTableTag.colgroup)
    #Create THEAD tag
    $theadTag = $htmldata.CreateElement("thead",$htmldata.html.NamespaceURI)
    $htmlTableTag.InsertBefore($theadTag,$htmlTableTag.tr[0])
    #Move firt row of TABLE into THEAD
    $theadTag.PrependChild($htmlTableTag.tr[0])
    #Create TBODY tag
    $tbodyTag = $htmldata.CreateElement("tbody",$htmldata.html.NamespaceURI)
    $htmlTableTag.InsertBefore($tbodyTag,$htmlTableTag.tr[0])
    #Move all remaining rows into TBODY
    $htmlTableTag.tr | ForEach-Object { $tbodyTag.AppendChild($_) }
    #Set TABLE ID
    $htmlTableTag.SetAttribute('id','myTable')
    #Set TABLE CLASS
    $htmlTableTag.SetAttribute('class','nowrap')
    return $htmldata
}

function generateclasslist($Type, $BaseName, $View, $schemaobject) {
    $htmlPre = @"
    <h1>Class Report</h1>
"@
    $htmlPost = @"
    <div class="reportfooter">Generated by Simple Schema Reporter.</div>
"@

$htmlHead = @"
<title>Class Report</title>
$htmlscript
$stylesheet
"@

switch ($type) {
    'HTMLFile' {
        $outputfilename = "$($basename).html"
        $outputdata = $schemaobject | selectclasses | Sort-Object -Property Name | ConvertTo-Html -Head $htmlHead -PreContent $htmlPre -PostContent $htmlPost
        $htmlpage = addjscript $outputdata
        Out-File -FilePath $outputfilename -InputObject $htmlpage.html.OuterXml
    }
    'HTMLClipboard' {
        $outputdata = $schemaobject | selectclasses | Sort-Object -Property Name | ConvertTo-Html -Fragment -PreContent $htmlPre -PostContent $htmlPost
        Set-Clipboard -Value $outputdata
    }
    'XMLFile' {
        $outputfilename = "$($basename).xml"
        $outputdata = $schemaobject | selectclasses | Sort-Object -Property Name | ConvertTo-XML -As String -NoTypeInformation
        Out-File -FilePath $outputfilename -InputObject $outputdata
    }
    'CSVFile' {
        $outputfilename = "$($basename).csv"
        $outputdata = $schemaobject | selectclasses | Sort-Object -Property Name
        $outputdata | Export-CSV -NoTypeInformation -Path $outputfilename
    }
}
Write-Verbose "$type Class Report created"
if ($View -and ($type -ne 'HTMLClipboard')) { Start-Process $outputfilename }   
}

function generatereport ($Schema, $Type, $BaseName, $View) {
$htmlPre = @"
        <h1>Schema Report for $schemaToReport Class</h1>
"@
        $htmlPost = @"
        <div class="reportfooter">Generated by Simple Schema Reporter.</div>
"@

$htmlHead = @"
<title>Schema Report for $schemaToReport Class</title>
$htmlscript
$stylesheet
"@
switch ($type) {
    'HTMLFile' {
        $outputfilename = "$($basename).html"
        $outputdata = $Schema | selectattributes | Sort-Object -Property Name | ConvertTo-Html -Head $htmlHead -PreContent $htmlPre -PostContent $htmlPost
        #Add a Javascript to allow table sorting
        $htmldata = addjscript $outputdata
        Out-File -FilePath $outputfilename -InputObject $htmldata.html.OuterXml
    }
    'HTMLClipboard' {
        $outputdata = $Schema | selectattributes | Sort-Object -Property Name | ConvertTo-Html -Fragment -PreContent $htmlPre -PostContent $htmlPost
        Set-Clipboard -Value $outputdata
    }
    'XMLFile' {
        $outputfilename = "$($basename).xml"
        $outputdata = $Schema | selectattributes | Sort-Object -Property Name | ConvertTo-XML -As String -NoTypeInformation
        Out-File -FilePath $outputfilename -InputObject $outputdata
    }
    'CSVFile' {
        $outputfilename = "$($basename).csv"
        $outputdata = $Schema | selectattributes | Sort-Object -Property Name
        $outputdata | Export-CSV -NoTypeInformation -Path $outputfilename
    }
}
Write-Verbose "$type Report created for $schemaToReport Class"
if ($View -and ($type -ne 'HTMLClipboard')) { Start-Process $outputfilename }
}

#region main code
#Getting the schema
$schema = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySchema]::GetCurrentSchema()

if ($ListClasses.IsPresent) {
    ForEach ($reportTypetoRun in $ReportType) {
        if (($Null -eq $ReportName) -or ('' -eq $ReportName)) {
            $ReportDisplayName = "SimpleSchema-Class-List-$((Get-Date).ToString('yyyy-MM-dd-HH-mm-ss'))"
            }
        else {
        Write-Verbose "***$ReportName***"
            $ReportDisplayName = "$ReportName-Class-List"
        }
        generateclasslist -Type $reportTypetoRun -View $ViewOutput -BaseName $ReportDisplayName -SchemaObject $schema
    }
}
else {
    #Getting Constructed Attributes
    $constrattrnames = findcontsrattr
    #Creating the reports
    ForEach ($schemaToReport in $ClassName) {
        $classSchema = classprops $schemaToReport
        ForEach ($reportTypetoRun in $ReportType) {
            if (($Null -eq $ReportName) -or ('' -eq $ReportName)) {
                $ReportDisplayName = "SimpleSchema-$($schemaToReport)-$((Get-Date).ToString('yyyy-MM-dd-HH-mm-ss'))"
                }
            else {
            Write-Verbose "***$ReportName***"
                $ReportDisplayName = "$ReportName-$($schemaToReport)"
            }
            generatereport -Schema $classSchema -Type $reportTypeToRun -BaseName $ReportDisplayName -View $ViewOutput
        }
    }
}
#endregion

#SimpleSchemaReporter.ps1