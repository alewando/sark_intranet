<!--
Developer:    Kevin Dill
Date:         10/20/1998
Description:  Listing of events by date.
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


'-----------------------
'  Build birthday list  
'-----------------------
sql =	"SELECT EmployeeID, FirstName, LastName, Username, DATEPART(day, Birthday) as DayOfBirth, Birthday " & _
		"FROM Employee " & _
		"WHERE (Birthday IS NOT NULL) and (DATEPART(month, Birthday)=" & rmonth & ") " & _
		"ORDER BY LastName"
set rs = DBQuery(sql)
while not rs.eof
	DayOfMonth = rs("DayOfBirth")
	CalendarDate(DayOfMonth) = CalendarDate(DayOfMonth) & "&nbsp;&nbsp;&nbsp;&nbsp;" & buildItem("/" & Application("Web") & "/directory/content/details.asp?EmpID=" & rs("EmployeeID"),  " Birthday:&nbsp;&nbsp;" & rs("FirstName") & " " & rs("LastName"), imgBirthday, "") & "<br>" & NL
	rs.movenext
	wend
rs.close
set rs = nothing


'--------------------------------
'  Build list of events by date  
'--------------------------------
sql =	"SELECT Sub_Type, ImageLoc, Event_ID, Title, StartTime, EndTime, DATEPART(day, EventDate) as DayOfMonth " &_
		"FROM v_events_by_date " &_
		"WHERE (EventDate >= '" & FirstDayOfMonth & "' AND EventDate < '" & FirstDayOfNextMonth & "') " &_
		"ORDER BY Type, Sub_Type"
set rs = DBQuery(sql)
while not rs.eof
	DayOfMonth = rs("DayOfMonth")
	EventImage = rs("ImageLoc")
	if isnull(rs("ImageLoc")) then EventImage = imgTemp
	CalendarDate(DayOfMonth) = CalendarDate(DayOfMonth) & "&nbsp;&nbsp;&nbsp;&nbsp;" & buildItem("details.asp?id=" & rs("Event_ID") & "&lastpage=bydate&date=" & server.URLEncode(rdate), " " & rs("Sub_Type") & ":&nbsp;&nbsp;" & rs("Title") & " (" & rs("StartTime") & " - " & rs("EndTime") & ")", EventImage, "") & "<br>" & NL
	rs.movenext
	wend
rs.close
set rs = nothing


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
'-------------------------------
'  Build event display by date  
'-------------------------------
for CurDay = 1 to DatePart("d", LastDayOfMonth)
	if not isempty(CalendarDate(CurDay)) then
		response.write("<br>" & rmonthname & " " & CurDay & ", " & ryear & "&nbsp;&nbsp;(" & WeekDayName(DatePart("w", DateSerial(ryear, rmonth, CurDay))) & ")<br>" & CalendarDate(CurDay) & NL)
		end if
	next
%>


<!-- #include file="../../footer.asp" -->
