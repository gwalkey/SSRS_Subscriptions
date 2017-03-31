# Create, Delete, List SQL Server Reporting Services Subscriptions from Powershell

Based on Code from Jasmin Mistry

https://gist.github.com/jasmin-mistry/9de9daf74cbc1d56a33984f2787ad2ea

This code assumes you have permissions to the SSRS Webservice

<b>Get-SSRS_Subscriptions</b> -  lists all the subscriptions for a given Report

Parameters: If your SSRS server is called 'SSRS2016' and your top-level folder structure has a folder named 'Sales' and a report in that folder called 'Sales Projections',  you need to set the script parameters as 'SSRS2016' and '/Sales/Sales Projections' accordingly

<b>New-SSRS_Subscription</b> -  Creates a subscription for a given Report using a variety of SSRS options

Parameters: Script Parameter $SSRSServer should be set to your SSRS instance

Internal Parameters: The powershell Comment block at the stop details all the possible Subscription Options including
Email or Fileshare
File share path
Email Recipients
