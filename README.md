# Create, Delete, List SQL Server Reporting Services Subscriptions using Powershell and SOAP

Based on Code from Jasmin Mistry

https://gist.github.com/jasmin-mistry/9de9daf74cbc1d56a33984f2787ad2ea


This Code will be updated shortly to include the option to persist subscriptions to external JSON Files

This code assumes you have permissions to the SSRS SOAP Webservice

*<b>Get-SSRS_Subscriptions</b> -  lists all the subscriptions for a given Report*

Parameters: If your SSRS server is called 'SSRS2016' and your top-level folder structure has a folder named 'Sales' and a report in that folder called 'Sales Projections',  you need to set the script parameters as 'SSRS2016' and '/Sales/Sales Projections' accordingly

*<b>New-SSRS_Subscription</b> -  Creates a subscription for a given Report using a variety of SSRS options*

Parameters: Script Parameter $SSRSServer should be set to your SSRS instance

Internal Parameters: The powershell Comment block at the top details all the possible Subscription Options including

Email or Fileshare

File share path

Email Recipients

Schedule Time and Repeating frequency

Report Parameter options include a sample for a multi-select combo-type parameter

The Scheduling Options are maintained in the XML Fragment


*<b>Remove-SSRS_Subscriptions</b> -  Deletes all subscriptions for a given Report*

Parameters: SSRS Instance and full Report Path (eg '/Sales/Western Division/Sales by Manager')

*In SSRS 2017 and the Power BI Reporting Server, a REST API exists and is much easier to use*
