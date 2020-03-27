#Testing of COVID-19 up to date data import into excel.

Excel document sends query to https://www.worldometers.info/coronavirus/ getting data from table 0.
Then under query properties, refresh has been set to every 60 seconds.


On a fourum on how to update queries using VBA https://social.technet.microsoft.com/Forums/en-US/5f20a3c0-937e-47ab-91ad-806353720510/refreshing-queries-w-vba?forum=powerquery

ThisWorkbook.Connections("Query - Data").Refresh

and so on..., but as the referenced thread indicates, 
you need to turn off background refresh for the connections prior to refreshing.  