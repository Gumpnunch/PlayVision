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
SQL=SQL & " "
SQL=SQL & "	SELECT	  tt.IdSite "
SQL=SQL & "		, tt.SiteCode "
SQL=SQL & "		, max(tt.P1) as P1 "
SQL=SQL & "		, max(tt.P2) as P2 "
SQL=SQL & "		, max(tt.P3) as P3 "
SQL=SQL & "		, max(tt.P4) as P4 "
SQL=SQL & "		, max(tt.P5) as P5 "
SQL=SQL & "from "
SQL=SQL & "( "
SQL=SQL & "SELECT  "
SQL=SQL & "	site.IdSite, site.SiteCode, count(stat.name) as P1, 0 as P2, 0 as P3, 0 as P4, 0 as P5 "
SQL=SQL & "FROM  "
SQL=SQL & "	[dbo].[ARA_Location] loc "
SQL=SQL & "INNER JOIN [dbo].[ARA_AsbestosStatus] as stat on stat.ARA_AsbestosStatusId = loc.ARA_AsbestosStatusId "
SQL=SQL & "INNER JOIN [dbo].[Room] as room on room.IdRoom = loc.roomid "
SQL=SQL & "INNER JOIN [dbo].Floor as floor on floor.IdFloor = room.IdFloor "
SQL=SQL & "INNER JOIN [dbo].Building as building on building.IdBuilding = floor.IdBuilding "
SQL=SQL & "INNER JOIN [dbo].Site as Site on Site.IdSite = building.IdSite "
SQL=SQL & "where  "
SQL=SQL & "	stat.name = 'Present'  and site.IsActive = 1 and ISNULL(building.IsHidden, 0) <> 1 and ISNULL(floor.IsHidden, 0) <> 1 and ISNULL(room.IsHidden, 0) <> 1 and ISNULL(loc.Obsolete, 0) <> 1      "
SQL=SQL & "group by  "
SQL=SQL & "	site.idsite, site.sitecode, stat.name "
SQL=SQL & " "
SQL=SQL & "Union All		 "
SQL=SQL & "SELECT  "
SQL=SQL & "	site.IdSite, site.SiteCode, 0 as P1, count(stat.name) as P2, 0 as P3, 0 as P4, 0 as P5 "
SQL=SQL & "FROM  "
SQL=SQL & "	[dbo].[ARA_Location] loc "
SQL=SQL & "INNER JOIN [dbo].[ARA_AsbestosStatus] as stat on stat.ARA_AsbestosStatusId = loc.ARA_AsbestosStatusId "
SQL=SQL & "INNER JOIN [dbo].[Room] as room on room.IdRoom = loc.roomid "
SQL=SQL & "INNER JOIN [dbo].Floor as floor on floor.IdFloor = room.IdFloor "
SQL=SQL & "INNER JOIN [dbo].Building as building on building.IdBuilding = floor.IdBuilding "
SQL=SQL & "INNER JOIN [dbo].Site as Site on Site.IdSite = building.IdSite "
SQL=SQL & "where  "
SQL=SQL & "	stat.name = 'Strongly Presumed'  and site.IsActive = 1  and ISNULL(building.IsHidden, 0) <> 1 and ISNULL(floor.IsHidden, 0) <> 1 and ISNULL(room.IsHidden, 0) <> 1 and ISNULL(loc.Obsolete, 0) <> 1    "
SQL=SQL & "group by  "
SQL=SQL & "	site.idsite, site.sitecode, stat.name "
SQL=SQL & " "
SQL=SQL & "Union All		 "
SQL=SQL & "SELECT  "
SQL=SQL & "	site.IdSite, site.SiteCode, 0 as P1, 0 as P2, count(stat.name) as P3, 0 as P4, 0 as P5 "
SQL=SQL & "FROM  "
SQL=SQL & "	[dbo].[ARA_Location] loc "
SQL=SQL & "INNER JOIN [dbo].[ARA_AsbestosStatus] as stat on stat.ARA_AsbestosStatusId = loc.ARA_AsbestosStatusId "
SQL=SQL & "INNER JOIN [dbo].[Room] as room on room.IdRoom = loc.roomid "
SQL=SQL & "INNER JOIN [dbo].Floor as floor on floor.IdFloor = room.IdFloor "
SQL=SQL & "INNER JOIN [dbo].Building as building on building.IdBuilding = floor.IdBuilding "
SQL=SQL & "INNER JOIN [dbo].Site as Site on Site.IdSite = building.IdSite "
SQL=SQL & "where  "
SQL=SQL & "	stat.name = 'Presumed'  and site.IsActive = 1  and ISNULL(building.IsHidden, 0) <> 1 and ISNULL(floor.IsHidden, 0) <> 1 and ISNULL(room.IsHidden, 0) <> 1 and ISNULL(loc.Obsolete, 0) <> 1    "
SQL=SQL & "group by  "
SQL=SQL & "	site.idsite, site.sitecode, stat.name "
SQL=SQL & " "
SQL=SQL & "Union All		 "
SQL=SQL & "SELECT  "
SQL=SQL & "	site.IdSite, site.SiteCode, 0 as P1, 0 as P2, 0 as P3, count(stat.name) as P4, 0 as P5 "
SQL=SQL & "FROM  "
SQL=SQL & "	[dbo].[ARA_Location] loc "
SQL=SQL & "INNER JOIN [dbo].[ARA_AsbestosStatus] as stat on stat.ARA_AsbestosStatusId = loc.ARA_AsbestosStatusId "
SQL=SQL & "INNER JOIN [dbo].[Room] as room on room.IdRoom = loc.roomid "
SQL=SQL & "INNER JOIN [dbo].Floor as floor on floor.IdFloor = room.IdFloor "
SQL=SQL & "INNER JOIN [dbo].Building as building on building.IdBuilding = floor.IdBuilding "
SQL=SQL & "INNER JOIN [dbo].Site as Site on Site.IdSite = building.IdSite "
SQL=SQL & "where  "
SQL=SQL & "	stat.name = 'Removed'  and site.IsActive = 1  and ISNULL(building.IsHidden, 0) <> 1 and ISNULL(floor.IsHidden, 0) <> 1 and ISNULL(room.IsHidden, 0) <> 1 and ISNULL(loc.Obsolete, 0) <> 1    "
SQL=SQL & "group by  "
SQL=SQL & "	site.idsite, site.sitecode, stat.name "
SQL=SQL & " "
SQL=SQL & " "
SQL=SQL & "Union All		 "
SQL=SQL & "SELECT  "
SQL=SQL & "	site.IdSite, site.SiteCode, 0 as P1, 0 as P2, 0 as P3, 0 as P4, count(stat.name) as P5 "
SQL=SQL & "FROM  "
SQL=SQL & "	[dbo].[ARA_Location] loc "
SQL=SQL & "INNER JOIN [dbo].[ARA_AsbestosStatus] as stat on stat.ARA_AsbestosStatusId = loc.ARA_AsbestosStatusId "
SQL=SQL & "INNER JOIN [dbo].[Room] as room on room.IdRoom = loc.roomid "
SQL=SQL & "INNER JOIN [dbo].Floor as floor on floor.IdFloor = room.IdFloor "
SQL=SQL & "INNER JOIN [dbo].Building as building on building.IdBuilding = floor.IdBuilding "
SQL=SQL & "INNER JOIN [dbo].Site as Site on Site.IdSite = building.IdSite "
SQL=SQL & "where  "
SQL=SQL & "	stat.name = 'Not Suspected'  and site.IsActive = 1 and ISNULL(building.IsHidden, 0) <> 1 and ISNULL(floor.IsHidden, 0) <> 1 and ISNULL(room.IsHidden, 0) <> 1 and ISNULL(loc.Obsolete, 0) <> 1     "
SQL=SQL & "group by  "
SQL=SQL & "	site.idsite, site.sitecode, stat.name "
SQL=SQL & "				 "
SQL=SQL & ") tt "
SQL=SQL & "group by tt.IdSite, tt.SiteCode "
SQL=SQL & "order by tt.SiteCode "

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL, Connection

dim oPut
dim tmp
dim intCount

intcount = 0

If Recordset.EOF Then
	oPut="000"
Else
	oPut = "["
'if there are records then loop through the fields
	Do While NOT Recordset.Eof   
		intcount = intcount + 1

		if oPut <> "[" then oPut = oPut & ","
		oPut = oPut & "{" & chr(34) & "Site_ID" & chr(34) & ":" & chr(34) & Recordset("idSite") & chr(34) & ","
		oPut = oPut & chr(34) & "Site_Code" & chr(34) & ":" & chr(34) & Recordset("SiteCode") & chr(34) & ","
		oPut = oPut & chr(34) & "Positive" & chr(34) & ":" & chr(34) & Recordset("P1") & chr(34) & ","
		oPut = oPut & chr(34) & "Strongly_Presumed" & chr(34) & ":" & chr(34) & Recordset("P2") & chr(34) & ","
		oPut = oPut & chr(34) & "Presumed" & chr(34) & ":" & chr(34) & Recordset("P3") & chr(34) & "}"

		Recordset.MoveNext    
	Loop
	oPut=oPut & "]"
End If

response.write(oPut)
Recordset.Close
Set Recordset=Nothing
Connection.Close
Set Connection=Nothing
%> 