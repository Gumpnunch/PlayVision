<% 
'declare the variable that will hold new connection object
Dim Connection   
'create an ADO connection object
Set Connection=Server.CreateObject("ADODB.Connection")

'declare the variable that will hold the connection string
Dim ConnectionString 
'define connection string, specify database driver and location of the database
ConnectionString ="DRIVER={SQL Server};SERVER=162.13.114.34;UID=RMS_SA;" & _
"PWD=mY%!IyP9n/y_6W.;DATABASE=RMS_DB_C1279"

'open the connection to the database
Connection.Open ConnectionString

'declare the variable that will hold our new object
Dim Recordset   
'create an ADO recordset object
Set Recordset=Server.CreateObject("ADODB.Recordset")

'declare the variable that will hold the SQL statement
Dim SQL   
SQL=""
SQL=SQL & "SELECT  count(distinct Site.SiteCode)	  "
SQL=SQL & "FROM "
SQL=SQL & "[dbo].[ARA_Location] as loc "
SQL=SQL & "INNER JOIN [dbo].[ARA_AsbestosStatus] as stat on stat.ARA_AsbestosStatusId = loc.ARA_AsbestosStatusId "
SQL=SQL & "INNER JOIN [dbo].[Room] as room on room.IdRoom = loc.roomid "
SQL=SQL & "INNER JOIN [dbo].Floor as floor on floor.IdFloor = room.IdFloor "
SQL=SQL & "INNER JOIN [dbo].Building as building on building.IdBuilding = floor.IdBuilding "
SQL=SQL & "INNER JOIN [dbo].Site as Site on Site.IdSite = building.IdSite "
SQL=SQL & "WHERE  "
SQL=SQL & "site.IsActive = 1 "
SQL=SQL & "and isnull(stat.acute,0) =0  "
SQL=SQL & "and ISNULL(building.IsHidden, 0) <> 1 and ISNULL(floor.IsHidden, 0) <> 1 and ISNULL(room.IsHidden, 0) <> 1 and ISNULL(loc.Obsolete, 0) <> 1 "
SQL=SQL & "and site.sitecode not in ( "
SQL=SQL & "SELECT  distinct Site.SiteCode	  "
SQL=SQL & "FROM "
SQL=SQL & "[dbo].[ARA_Location] as loc "
SQL=SQL & "INNER JOIN [dbo].[ARA_AsbestosStatus] as stat on stat.ARA_AsbestosStatusId = loc.ARA_AsbestosStatusId "
SQL=SQL & "INNER JOIN [dbo].[Room] as room on room.IdRoom = loc.roomid "
SQL=SQL & "INNER JOIN [dbo].Floor as floor on floor.IdFloor = room.IdFloor "
SQL=SQL & "INNER JOIN [dbo].Building as building on building.IdBuilding = floor.IdBuilding "
SQL=SQL & "INNER JOIN [dbo].Site as Site on Site.IdSite = building.IdSite "
SQL=SQL & "WHERE  "
SQL=SQL & "site.IsActive = 1 "
SQL=SQL & "and isnull(stat.acute,0) >0  "
SQL=SQL & "and ISNULL(building.IsHidden, 0) <> 1 and ISNULL(floor.IsHidden, 0) <> 1 and ISNULL(room.IsHidden, 0) <> 1 and ISNULL(loc.Obsolete, 0) <> 1) "

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL, Connection

If Recordset.EOF Then
Response.Write("000")
Else
'if there are records then loop through the fields
Do While NOT Recordset.Eof   
Response.write Recordset(0)
Response.write "<br>"   
Recordset.MoveNext    
Loop
End If

Recordset.Close
Set Recordset=Nothing
Connection.Close
Set Connection=Nothing
%> 