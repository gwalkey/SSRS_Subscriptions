<#
.SYNOPSIS
    Create SSRS Report Subscription via SSRS WebService using Powershell
	
.DESCRIPTION
    Create SSRS Report Subscription via SSRS WebService using Powershell
	
.EXAMPLE
  
.Inputs
  $ReportServerUri, $MatchData, ReportName, ReportPath

.Outputs

	
.NOTES
	
.LINK
	
	
#>
Set-StrictMode -Version latest;

$ReportServerUri  = "http://SSRS2016/ReportServer/ReportService2010.asmx"
$sitePath = "/" 
 
$rs2010 += New-WebServiceProxy -Uri $ReportServerUri -UseDefaultCredential;  


# Get Types from Namespace
$type = $rs2010.GetType().Namespace

# Define Object Types for Subscription property call
$ExtensionSettingsDataType = ($type + '.ExtensionSettings')
$ActiveStateDataType = ($type + '.ActiveState')
$ParmValueDataType = ($type + '.ParameterValue')

# Create New ExtensionSettings Object based on Type 
$extSettings = New-Object ($ExtensionSettingsDataType)

# Function Call parameters setup
$extensionParams = @()
$rptParams = @()

# XML Fragment holds all the Subcription Parameters
[string]$MatchData =
"<Subscription>
	<ReportName>TestReport1</ReportName>
	<ReportPath>/</ReportPath>
	<Description>Created By Powershell</Description>
	<EventType>TimedSubscription</EventType>
	<ExtensionSettings>
		<DeliveryExtension>Report Server Email</DeliveryExtension>
        <ParameterValues>
	        <ParameterValue>
		        <Name>TO</Name>
		        <Value>user@domain.com</Value>
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
		        <Value>Testing Powershell Subscription Creation</Value>
	        </ParameterValue>
        </ParameterValues>
	</ExtensionSettings>
	<ScheduleXML>
		<![CDATA[
			<ScheduleDefinition>
				<StartDateTime>2017-01-01T08:00:00.000-05:00</StartDateTime>
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
		]]>
	</ScheduleXML>
	<ReportParameter>
	  <ParameterValue>
		  <Name></Name>
		  <Value></Value>
	  </ParameterValue>
	</ReportParameter>
</Subscription>"

# Load Subscription build parameters from an XML File with includes a section for the schedule
[xml]$xml = $MatchData
$xSubscription = $xml.Subscription
$xExtensionSettings = $xSubscription.ExtensionSettings

# Get more Function paremeters
$report = [string]::Join("", $xSubscription.ReportPath, $xSubscription.ReportName) 
$desc = $xSubscription.Description
$event = $xSubscription.EventType
$scheduleXml = $xSubscription.ScheduleXML.'#cdata-section'
$extSettings.Extension = $xExtensionSettings.DeliveryExtension

# read the extension setting parameter values
$xExtParams = $xExtensionSettings.ParameterValues.ParameterValue

foreach ($p in $xExtParams) {
	$param = New-Object ($ParmValueDataType)
	$param.Name = $p.Name
	$param.Value = $p.Value
	$extensionParams += $param
}
# Build up object
$extSettings.ParameterValues = $extensionParams

# Prep Report Param Values (null this time)
$paramValues = @()

# Do It
try
{
    $subscriptionID = $rs2010.CreateSubscription($report,$extSettings,$desc, $event, $scheduleXml, $paramValues)
    Write-Output("SubscriptionID: {0}" -f $subscriptionID)
}
catch
{
    Write-Output ("Exception: {0} Inner: {1}" -f $_.Exception.Message, $_.Exception.Message.InnerException)
    $error[0] | fl -force
}



$rs2010 = $null
