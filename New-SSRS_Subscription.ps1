<#
.SYNOPSIS
    Creates SSRS Report Subscription using the SSRS Web Service
	
.DESCRIPTION
    Creates SSRS Report Subscription using the SSRS Web Service
   
.EXAMPLE
    
	
.Inputs
    

.Outputs
	

.NOTES
    https://msdn.microsoft.com/en-us/library/ms154020(v=sql.130).aspx
    https://msdn.microsoft.com/en-us/library/reportservice2010.parametervalue.aspx


    1) The ExtensionSettings holds the Subscription Type settings(case sensitive):

    a) "Report Server Email"
        Parameters:
        TO
        CC
        BCC
        ReplyTo
        IncludeReport = True/False
        RenderFormat = EXCELOPENXML, IMAGE, XML, PPTX, CSV, PDF (Landscape), PDF, MHTML, WORDOPENXML, PDF (Portrait)
        Priority = NORMAL, HIGH
        Subject = @ReportName
        Comment
        IncludeLink = True/False

            
    b) "Report Server Fileshare"
        Parameters:
        FILENAME = "report.pdf"
        PATH = "\\fileserver\mnt\subfolder"
        RENDER_FORMAT = PDF, MHTML, IMAGE, CSV, XML, EXCELOPENXML, PDF (Landscape), PPTX, WORDOPENXML, PDF (Portrait)
        WRITEMODE = None, OverWrite, AutoIncrement
        FILEEXTN = True/False - Add an Extension based on Type (.PDF)
        USERNAME = Share Creds
        PASSWORD = Share Creds
        DEFAULTCREDENTIALS
    
    
    2) The MatchData or Schedule XML Type:
    https://msdn.microsoft.com/en-us/library/reportservice2005.recurrencepattern(v=sql.130).aspx

    Monthly (Calendar Days of Selected Months):
    <ScheduleDefinition>
		<StartDateTime>2017-01-01T07:00:00.000-05:00</StartDateTime>
            <MonthlyRecurrence>
                <Days>1</Days>
                <MonthsOfYear>
                    <January>true</January>
                    <February>true</February>
                    <March>true</March>
                    <April>true</April>
                    <May>true</May>
                    <June>true</June>
                    <July>true</July>
                    <August>true</August>
                    <September>true</September>
                    <October>true</October>
                    <November>true</November>
                    <December>true</December>
                </MonthsOfYear>
            </MonthlyRecurrence>
    <ScheduleDefinition>

    Weekly (M-F at 0700):
    <ScheduleDefinition>
		<StartDateTime>2017-01-01T07:00:00.000-05:00</StartDateTime>
		<WeeklyRecurrence>
			<WeeksInterval>1</WeeksInterval>
				<DaysOfWeek>
					<Monday>true</Monday>
					<Tuesday>true</Tuesday>
					<Wednesday>true</Wednesday>
					<Thursday>true</Thursday>
					<Friday>true</Friday>
				</DaysOfWeek>
		</WeeklyRecurrence>
	</ScheduleDefinition>

    Daily (at 0600):
    <ScheduleDefinition>
	    <StartDateTime>2017-01-01T06:00:00.000-05:00</StartDateTime>
        <DailyRecurrence>
            <DaysInterval>1</DaysInterval>
        </DailyRecurrence>
    </ScheduleDefinition>

    Once (at 1025):
    <ScheduleDefinition>
        <StartDateTime>02/20/2017 10:25:00</StartDateTime>
    </ScheduleDefinition>


    3) Report Parameters
    For populating report drop-downs with multiple selections, repeat the ParameterValue block as needed

    <ReportParameter>
	  <ParameterValue>
		  <Name>Center</Name>
		  <Value>1</Value>
	  </ParameterValue>
	  <ParameterValue>
		  <Name>Center</Name>
		  <Value>2</Value>
	  </ParameterValue>
	  <ParameterValue>
		  <Name>Center</Name>
		  <Value>3</Value>
	  </ParameterValue>
	  <ParameterValue>
		  <Name>Center</Name>
		  <Value>4</Value>
	  </ParameterValue>

	  <ParameterValue>
		  <Name>Month</Name>
		  <Value>January</Value>
	  </ParameterValue>
      <ParameterValue>
		  <Name>Month</Name>
		  <Value>February</Value>
	  </ParameterValue>	

	  <ParameterValue>
		  <Name>Requirement</Name>
		  <Value>0</Value>
	  </ParameterValue>
	  <ParameterValue>
		  <Name>Requirement</Name>
		  <Value>1</Value>
	  </ParameterValue>
	</ReportParameter>

#>

Param(
    [parameter(Position=0,mandatory=$true,ValueFromPipeline)]
    [ValidateNotNullOrEmpty()]
    [string]$SSRSServer
)


# ----------------------------
# -- Functions rule the earth
# ----------------------------
function CreateSSRSSubscription
{

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][string]$prmMatchData,
        [Parameter(Mandatory=$true)][string]$prmSchedule
        )

# SSRS Server URI
$ReportServerUri  = "http://$SSRSServer/ReportServer/ReportService2010.asmx"

# Open Web Service Connection
$rs2010 += New-WebServiceProxy -Uri $ReportServerUri -UseDefaultCredential;  

# Get Types from SSRS Webservice Namespace
$type = $rs2010.GetType().Namespace

# Define Object Types for Subscription property call
# http://stackoverflow.com/questions/25984874/not-able-to-create-objects-in-powershell-for-invoking-a-web-service
# http://stackoverflow.com/questions/32611187/using-complex-objects-via-a-web-service-from-powershell
# This XML Fragment holds Three sections
# 1) Extension Settings (Email or Fileshare, where, who etc)
# 2) Schedule
# 3) Report Parameters

$ExtensionSettingsDataType = ($type + '.ExtensionSettings')
$ActiveStateDataType = ($type + '.ActiveState')
$ParmValueDataType = ($type + '.ParameterValue')

# Create New ExtensionSettings Object based on Type 
$extSettings = New-Object ($ExtensionSettingsDataType)
$AllReportParameters = New-Object ($ParmValueDataType)

# Function Call parameters setup
$extensionParams = @()
$rptParamArray = @()


# Load Subscription build parameters from an XML File with includes a section for the schedule
[xml]$xml = $prmMatchData
$xSubscription = $xml.Subscription
$xExtensionSettings = $xSubscription.ExtensionSettings

# Get more Report parameters
#$report = [string]::Join("", $xSubscription.ReportPath, $xSubscription.ReportName) 
$report = $xSubscription.ReportPath+$xSubscription.ReportName
$desc = $xSubscription.Description
$event = $xSubscription.EventType
$extSettings.Extension = $xExtensionSettings.DeliveryExtension

# Get Schedule from a direct XML Definition
$scheduleXml = $prmSchedule

# Get the extension settings parameter values from the XML Fragment
$xExtParams = $xExtensionSettings.ParameterValues.ParameterValue

foreach ($p in $xExtParams) {
	$param = New-Object ($ParmValueDataType)
	$param.Name = $p.Name
	$param.Value = $p.Value
	$extensionParams += $param
}
# Build up object
$extSettings.ParameterValues = $extensionParams


# Get Actual Report Parameters from XML Fragment
$ReportParameters= $xml.Subscription.ReportParameter.ParameterValue
foreach ($rp in $ReportParameters) {
    $rparam = New-Object ($ParmValueDataType)
	$rparam.Name = $rp.Name
	$rparam.Value = $rp.Value
	$rptParamArray += $rparam
}
# BuildUpObject from individual elements
$AllReportParameters = $rptParamArray


# Call the WebService
try
{
    $subscriptionID = $rs2010.CreateSubscription($report, $extSettings, $desc, $event, $scheduleXml, $AllReportParameters)
    Write-Output("Created Subscription ID: {0}" -f $subscriptionID)
}
catch
{
    Write-Output ("Exception: {0} Inner: {1}" -f $_.Exception.Message, $_.Exception.Message.InnerException)
    $error[0] | fl -force

    $rs2010 = $null
}



$rs2010 = $null

}



# ---------------------------------
# -- Generate Multiple Subs Example
# ---------------------------------

[string]$myReportName="/SomeCoolReport_Where_Parameters_Change_Frequently"
[string]$myReportPath="/Sales Reports Folder"
[string]$myLocation = "1"
[string]$myMonth="April"
[string]$myYear="2017"


# Create Sub for Each Location
foreach ($Location in (1..10))
{
    $myLocation = $Location
    $myMonth="April"
    $myYear="2017"
    $myEmail = "ReportConsumer@Domain.com"
    $mySubDescription = "Special Report for Location "+$myLocation + " - "+$myMonth + " " +$myYear

	$myMatchData =
	"<Subscription>
		<ReportName>$myReportName</ReportName>
		<ReportPath>$myReportPath</ReportPath>
		<Description>$mySubDescription</Description>
		<EventType>TimedSubscription</EventType>
		<ExtensionSettings>
			<DeliveryExtension>Report Server Email</DeliveryExtension>
			<ParameterValues>
				<ParameterValue>
					<Name>TO</Name>
					<Value>$myEmail</Value>
				</ParameterValue>
				<ParameterValue>
					<Name>IncludeReport</Name>
					<Value>True</Value>
				</ParameterValue>
				<ParameterValue>
					<Name>RenderFormat</Name>
					<Value>MHTML</Value>
				</ParameterValue>
				<ParameterValue>
					<Name>Subject</Name>
					<Value>$mySubDescription</Value>
				</ParameterValue>
				<ParameterValue>
					<Name>Priority</Name>
					<Value>NORMAL</Value>
				</ParameterValue>
			</ParameterValues>
		</ExtensionSettings>

		<ReportParameter>
		  <ParameterValue>
			  <Name>Center</Name>
			  <Value>$myLocation</Value>
		  </ParameterValue>
		  <ParameterValue>
			  <Name>Month</Name>
			  <Value>$MyMonth</Value>
		  </ParameterValue>
		  <ParameterValue>
			  <Name>Year</Name>
			  <Value>$myYear</Value>
		  </ParameterValue>
		</ReportParameter>

	</Subscription>"

    [string]$MySkedMonth = (get-date).Month
    [string]$myscheduleXml = 
    "
    <ScheduleDefinition>
        <StartDateTime>$myYear-$MySkedMonth-15T06:00:00.000-04:00</StartDateTime>
    </ScheduleDefinition>
    "

    CreateSSRSSubscription $myMatchData $myscheduleXml
}



