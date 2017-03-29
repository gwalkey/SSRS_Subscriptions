<#
.SYNOPSIS
    Deletes SSRS Report Subscriptions for One report
	
.DESCRIPTION
    Deletes SSRS Report Subscriptions for One report
   
.EXAMPLE
    
	
.Inputs
    

.Outputs
	

.NOTES
    https://msdn.microsoft.com/en-us/library/ms154020(v=sql.130).aspx

    1) The ExtensionSettings holds the Subscription Type settings(case sensitive):

    a) Report Server Email
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

            
    b) Report Server Fileshare
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

#>


Param(
    [parameter(Position=0,mandatory=$true,ValueFromPipeline)]
    [ValidateNotNullOrEmpty()]
    [string]$SSRSServer,
    [parameter(Position=1,mandatory=$true,ValueFromPipeline)]
    [ValidateNotNullOrEmpty()]
    [string]$Report
)

$ReportServerUri  = "http://$SSRSServer/ReportServer/ReportService2010.asmx"

# Optional -class parameter? 
$rs2010 += New-WebServiceProxy -Uri $ReportServerUri -UseDefaultCredential;  

# Get Types from Namespace
$type = $rs2010.GetType().Namespace

# Define Object Types for Subscription property call
# http://stackoverflow.com/questions/25984874/not-able-to-create-objects-in-powershell-for-invoking-a-web-service
# http://stackoverflow.com/questions/32611187/using-complex-objects-via-a-web-service-from-powershell

$ExtensionSettingsDataType = ($type + '.ExtensionSettings')
$ActiveStateDataType = ($type + '.ActiveState')
$ParmValueDataType = ($type + '.ParameterValue')

# Create typed parameters the method needs
$extSettings = New-Object ($ExtensionSettingsDataType)
$paramSettings = New-Object ($ParmValueDataType)
$activeSettings = New-Object ($ActiveStateDataType)
$desc = ""
$status = ""
$eventType = ""
$matchdata = ""

# Call the WebService
try
{
    $subscriptions = $rs2010.ListSubscriptions($report)
    if ($subscriptions -ne $null)
    {
        Write-Output("Subscriptions for Report {0} `r`n" -f $report)

        # Show Subs
        foreach($sub in $subscriptions)
        {
            # Sub            
            Write-Output("== Deleting Subscription {0} ==" -f $sub.subscriptionID)

            # Delete
            $status = $rs2010.DeleteSubscription($sub.SubscriptionID)

        }
    }
}
catch
{
    Write-Output ("Exception: {0} Inner: {1}" -f $_.Exception.Message, $_.Exception.Message.InnerException)
    $error[0] | fl -force
}



$rs2010 = $null
