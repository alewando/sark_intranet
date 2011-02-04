<!--
Developer:    Kevin Dill
Date:         
Description:  Listing of events by event type.
-->


<!-- #include file="../section.asp" -->


<%
'-----------------
'  Dim variables  
'-----------------
Dim Sql, CalendarDate(31)
Dim imgBirthday, imgTimesheets, imgTemp
Dim tsMidMonth, tsEndMonth, tsMidMonthDay, tsEndMonthDay
Dim FirstDayOfMonth, FirstDayOfNextMonth, DayOfMonth, EventImage
Dim htmlBirthdays, htmlEvents, LastEventType


'-----------------------
'  Set constant values  
'-----------------------
imgTemp = "images/calendar_temp.gif"
imgBirthday = "images/calendar_birthday.gif"
imgTimesheets = "images/calendar_timesheets.gif"


'---------------------
'  Get calendar date  
'---------------------
rdate = request.querystring("date")
if rdate = "" then rdate = Date
rmonth = DatePart("m", rdate)
ryear = DatePart("yyyy",rdate)
FirstDayOfMonth = DateAdd("d", -(DatePart("d", rdate)-1), rdate)
FirstDayOfNextMonth = DateAdd("m", 1, FirstDayOfMonth)
LastDayOfMonth = DateAdd("d", -1, FirstDayOfNextMonth)


'----------------------------------------
'  Function to build an individual item  
'----------------------------------------
function buildItem(link, textval, imageloc, imagealt)
	dim item
	if not isnull(imageloc) then item = "<img src=" & chr(34) & imageloc & chr(34) & " alt=" & chr(34) & imagealt & chr(34) & " height=16 width=16 border=0>"
	if not isnull(textval) then item = item & textval
	if not isnull(link) then item = "<a style='text-decoration:none' href=" & chr(34) & link & chr(34) & " onMouseOver=" & chr(34) & "window.status='View detail.'" & chr(34) & " onMouseOut=" & chr(34) & "window.status=''" & chr(34) & ">" & item & "</a>"
	buildItem = item
end function


'----------------------
'  Build event header  
'----------------------
rmonthname = MonthName(rmonth)
%>
<center>
<table border=0><tr><td>
<a href="bydate.asp?date=<%=DateAdd("yyyy",-1,rdate)%>"><img src="images/calendar_prev_year.gif" alt="Previous Year" height=16 width=16 border=0></a>
<a href="bydate.asp?date=<%=DateAdd("m",-1,rdate)%>"><img src="images/calendar_prev_month.gif" alt="Previous Month" height=16 width=16 border=0></a>
</td><td valign=center>&nbsp;&nbsp;<font size=2 face='ms sans serif, arial, geneva'><b><%=rmonthname & ", " & ryear%></b></font>&nbsp;&nbsp;</td><td>
<a href="bydate.asp?date=<%=DateAdd("m",1,rdate)%>"><img src="images/calendar_next_month.gif" alt="Next Month" height=16 width=16 border=0></a>
<a href="bydate.asp?date=<%=DateAdd("yyyy",1,rdate)%>"><img src="images/calendar_next_year.gif" alt="Next Year" height=16 width=16 border=0></a>
</td></tr></table>
</center>
<hr size=1>



<%
'--------------------------------
'  Build list of events by date  
'--------------------------------
sql =	"SELECT DISTINCT Type, Sub_Type, ImageLoc, Event_ID, Title " &_
		"FROM v_events_by_date " &_
		"WHERE (EventDate >= '" & FirstDayOfMonth & "' AND EventDate < '" & FirstDayOfNextMonth & "') " &_
		"ORDER BY Type, Sub_Type"
set rs = DBQuery(sql)
LastEventType = ""
while not rs.eof
	if LastEventType <> rs("Type") then response.write("<br><b>" & rs("Type") & ":</b><br>")
	EventImage = rs("ImageLoc")
	if isnull(rs("ImageLoc")) then EventImage = imgTemp
	response.write("&nbsp;&nbsp;&nbsp;&nbsp;" & buildItem("details.asp?id=" & rs("Event_ID") & "&lastpage=bytype&date=" & server.URLEncode(rdate), " " & rs("Sub_Type") & ":&nbsp;&nbsp;" & rs("Title"), EventImage, "") & "<br>" & NL)
	LastEventType = rs("Type")
	rs.movenext
	wend
rs.close
set rs = nothing


'-----------------------
'  Build birthday list  
'-----------------------
sql =	"SELECT EmployeeID, FirstName, LastName, Username, Birthday " & _
		"FROM Employee " & _
		"WHERE (Birthday IS NOT NULL) and (DATEPART(month, Birthday)=" & rmonth & ") " & _
		"ORDER BY LastName, FirstName"
set rs = DBQuery(sql)
response.write("<br><b>Birthdays:</b><br>")
if rs.eof then
	response.write("&nbsp;&nbsp;&nbsp;&nbsp;<font color=gray>[None]</font>")
else
	while not rs.eof
		response.write("&nbsp;&nbsp;&nbsp;&nbsp;" & buildItem("../../directory/content/details.asp?EmpID=" & rs("EmployeeID"),  " " & rs("FirstName") & " " & rs("LastName") & " (" & DatePart("m", rs("Birthday")) & "/" & DatePart("d", rs("Birthday")) & ")", imgBirthday, "") & "<br>" & NL)
		rs.movenext
		wend
	end if
rs.close
set rs = nothing
%>


<!-- #include file="../../footer.asp" -->
