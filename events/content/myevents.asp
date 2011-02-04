<!-- #include file="../section.asp" -->


<%
'---------------------------------------------------------------------------------
' Description:	View all of the events that I'm hosting, or attending.
' History:		10/26/1999 - KDILL - Created
'---------------------------------------------------------------------------------
%>


<font size=1>Upcoming events for <b><% =Session("LongName") %></b>:</font><p>



<table border=0 width="100%" cellspacing=0 cellpadding=3>
	<tr><td width=10></td><td width="10%"></td><td width="90%"></td></tr>
<%
' SELECT DISTINCT dbo.Event.Event_ID, dbo.Event.Title,
'                    MIN(dbo.Event_Date.EventDate) AS StartDate, dbo.Event_Type.ImageLoc,
'                    dbo.Event_Type.Type, dbo.Event_Sub_Type.Sub_Type
' FROM         dbo.Event INNER JOIN dbo.Event_Date ON dbo.Event.Event_ID 
'                        = dbo.Event_Date.Event_ID INNER JOIN dbo.Event_Sub_Type ON 
'                        dbo.Event.Event_Sub_Type_ID = dbo.Event_Sub_Type.Event_Sub_Type_ID 
'                        INNER JOIN dbo.Event_Type ON dbo.Event_Sub_Type.Event_Type_ID 
'                        = dbo.Event_Type.Event_Type_ID LEFT OUTER JOIN Event_Attendees ON 
'                        dbo.Event.Event_ID = Event_Attendees.Event_ID
' WHERE      (dbo.Event.Contact_ID = 43) AND (dbo.Event_Date.EventDate > GETDATE ()) 
'                        OR (dbo.Event_Date.EventDate > GETDATE ()) AND (dbo.Event.Creator_ID = 43) 
'                        OR (dbo.Event_Date.EventDate > GETDATE ()) AND (Event_Attendees.Employee_ID = 43)
' GROUP BY dbo.Event.Event_ID, dbo.Event.Title, dbo.Event_Sub_Type.Sub_Type,
'                    dbo.Event_Type.ImageLoc, dbo.Event_Type.Type
' ORDER BY MIN(dbo.Event_Date.EventDate)

id = Session("ID")
sql =	"SELECT DISTINCT Event.Event_ID, Event.Title," &_
		"                   MIN(Event_Date.EventDate) AS StartDate, Event_Type.ImageLoc," &_
		"                   Event_Type.Type, Event_Sub_Type.Sub_Type " &_
		"FROM         Event INNER JOIN Event_Date ON Event.Event_ID " &_
		"                       = Event_Date.Event_ID INNER JOIN Event_Sub_Type ON " &_
		"                       Event.Event_Sub_Type_ID = Event_Sub_Type.Event_Sub_Type_ID " &_
		"                       INNER JOIN Event_Type ON Event_Sub_Type.Event_Type_ID " &_
		"                       = Event_Type.Event_Type_ID LEFT OUTER JOIN Event_Attendees ON " &_
		"                       Event.Event_ID = Event_Attendees.Event_ID " &_
		"WHERE      ((Event.Contact_ID = " & id & ") OR (Event.Creator_ID = " & id & ") OR (Event_Attendees.Employee_ID = " & id & ") OR (Event_Type.Event_Type_ID = 'E')) " &_
		"					AND (Event_Date.EventDate > GETDATE ()) " &_
		"GROUP BY Event.Event_ID, dbo.Event.Title, Event_Sub_Type.Sub_Type," &_
		"                   Event_Type.ImageLoc, Event_Type.Type " &_
		"ORDER BY MIN(Event_Date.EventDate)"

set rs = DBQuery(sql)
while not rs.eof
	event_id = rs("Event_ID")
	event_title = rs("Sub_Type") & ", " & rs("Title")
	event_startdate = rs("StartDate")
	event_image = rs("ImageLoc")
	event_alt = rs("Type") & ": " & rs("Sub_Type")
	set rs1 = DBQuery("select LongDesc from Event where Event_ID = " & event_id)
	event_desc = trim(rs1(0))
	set rs2 = DBQuery("select * from Event_Date where Event_ID = " & event_id & " order by EventDate")
%>
	<tr>
		<td valign=top><img src="<% =event_image %>" alt="<% =event_alt %>" height=16 width=16 border=0></td>
		<td colspan=2><font size=1><a href="details.asp?id=<% =event_id %>"><b><% =event_title %></b></a><br>
<% if event_desc <> "" then response.write event_desc %>
		</tr>
	<tr>
		<td>&nbsp;</td>
		<td valign=top>&nbsp;</td>
		<td><font size=1>
<%
while not rs2.eof
	event_date = rs2("EventDate")
	if DateDiff("d", Date, event_date) < 7 then
		response.write "<b>" & event_date & "</b>"
	else
		response.write event_date
	end if
	response.write "&nbsp;&nbsp;(" & WeekDayName(WeekDay(rs2("EventDate"))) & ")"
	if (trim(rs2("StartTime")) <> "") or (trim(rs2("EndTime")) <> "") then
		response.write ":&nbsp;&nbsp;&nbsp;" & rs2("StartTime") & " - " & rs2("EndTime")
		end if
	response.write "<br>"
	rs2.movenext
	wend
rs2.close
%>
			</font></td>
		</tr>
<%
	rs1.close
	rs.movenext
	wend
rs.close
%>
</table>

<!-- #include file="../../footer.asp" -->
