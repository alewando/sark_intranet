<!--
Developer:    Kevin Dill
Date:         08/27/1998
Description:  Displays a calendar with event icons for events during
              the indicated month.
-->


<!-- #include file="../section.asp" -->

<%
'-----------------
'  Dim variables  
'-----------------
Dim Sql, CalendarDate(31), calCellWidth, calCellHeight, html
Dim imgBirthday, imgTimesheets, imgTemp
Dim tsMidMonth, tsEndMonth, tsMidMonthDay, tsEndMonthDay
Dim FirstDayOfMonth, FirstDayOfNextMonth, DateEnd, DayOfMonth, EventImage


'-----------------------
'  Set constant values  
'-----------------------
calCellWidth = 54
calCellHeight = 50
imgTemp = "images/calendar_temp.gif"
imgBirthday = "images/calendar_birthday.gif"
imgTimesheets = "images/calendar_timesheets.gif"


'---------------------
'  Get calendar date  
'---------------------
rdate = request.querystring("date")
if rdate = "" then rdate = Date
rmonth = DatePart("m", rdate)
if rmonth < 10 then rmonth = "0" & rmonth
ryear = DatePart("yyyy",rdate)
FirstDayOfMonth = DateAdd("d", -(DatePart("d", rdate)-1), rdate)
FirstDayOfNextMonth = DateAdd("m", 1, FirstDayOfMonth)
DateEnd = DateAdd("d", -1, FirstDayOfNextMonth)


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
	CalendarDate(DayOfMonth) = CalendarDate(DayOfMonth) & buildItem("../../directory/content/details.asp?EmpID=" & rs("EmployeeID"),  NULL, imgBirthday, rs("FirstName") & " " & rs("LastName") & "'s birthday")
	rs.movenext
	wend
rs.close
set rs = nothing


'--------------------
'  Build event list  
'--------------------
sql =	"SELECT e.Event_ID, e.Title, et.ImageLoc, est.Sub_Type, ed.StartTime, ed.EndTime, DATEPART(day, ed.EventDate) as DayOfMonth " & _
		"FROM Event e, Event_Date ed, Event_Type et, Event_Sub_Type est " & _
		"WHERE (e.Event_ID = ed.Event_ID) " & _
			"AND (est.Event_Sub_Type_ID = e.Event_Sub_Type_ID) " & _
			"AND (est.Event_Type_ID = et.Event_Type_ID) " & _
			"AND ((ed.EventDate >= '" & FirstDayOfMonth & "') AND (ed.EventDate < '" & FirstDayOfNextMonth & "')) "
set rs = DBQuery(sql)
while not rs.eof
	DayOfMonth = rs("DayOfMonth")
	EventImage = imgTemp
	html = ""
	if not isnull(rs("ImageLoc")) then EventImage = rs("ImageLoc")
	if (trim(rs("StartTime")) <> "") or (trim(rs("EndTime")) <> "") then
		html = NL & " (" & rs("StartTime") & " - " & rs("EndTime") & ")"
		end if
	CalendarDate(DayOfMonth) = CalendarDate(DayOfMonth) & buildItem("details.asp?id=" & rs("Event_ID") & "&lastpage=calendar&date=" & server.URLEncode(rdate), NULL, EventImage, rs("Sub_Type") & ": " & rs("Title") & html)
	rs.movenext
	wend
rs.close
set rs = nothing


'-------------------------
'  Build timesheet dates  
'-------------------------
midDay = 15
while instr(CalendarDate(midDay), "Holiday:")
	midDay = midDay - 1
	wend
endDay = DatePart("d", DateEnd)
while instr(CalendarDate(endDay), "Holiday:")
	endDay = endDay - 1
	wend
tsMidMonth = DateSerial(ryear, rmonth, midDay)
tsEndMonth = DateSerial(ryear, rmonth, endDay)
if DatePart("w", tsMidMonth) = 1 then tsMidMonth = DateAdd("d", -2, tsMidMonth)
if DatePart("w", tsMidMonth) = 7 then tsMidMonth = DateAdd("d", -1, tsMidMonth)
if DatePart("w", tsEndMonth) = 1 then tsEndMonth = DateAdd("d", -2, tsEndMonth)
if DatePart("w", tsEndMonth) = 7 then tsEndMonth = DateAdd("d", -1, tsEndMonth)
tsMidMonthDay = DatePart("d", tsMidMonth)
tsEndMonthDay = DatePart("d", tsEndMonth)
CalendarDate(tsMidMonthDay) = buildItem(NULL, NULL, imgTimesheets, "Timesheets due by noon!") & CalendarDate(tsMidMonthDay)
CalendarDate(tsEndMonthDay) = buildItem(NULL, NULL, imgTimesheets, "Timesheets due by noon!") & CalendarDate(tsEndMonthDay)
%>


<center>
<font color=maroon>Click on a date to add a new event...</font>
<table border=1 cellspacing=0 cellpadding=0><tr><td bgcolor="silver">
	
	<table border=0 width="100%">
	<tr>
		<td align=left width=48 nowrap>
			<a href="Calendar.asp?date=<%=DateAdd("yyyy",-1,rdate)%>"><img src="images/calendar_prev_year.gif" alt="Previous Year" height=16 width=16 border=0></a>&nbsp;&nbsp;
			<a href="Calendar.asp?date=<%=DateAdd("m",-1,rdate)%>"><img src="images/calendar_prev_month.gif" alt="Previous Month" height=16 width=16 border=0></a>
		</td>
		<td align=center width=48><font size=1 face='ms sans serif, arial, geneva'>
		<% if not Session("isGuest") then %>
			<a href="addnew.asp?date=<%=rdate%>">Add</a>
		<% end if %>
		</font></td>
		<td align=center nowrap><font size=2 color="#000000" face="ms sans serif, arial, geneva">
			<b><%=MonthName(rmonth) & " " & DatePart("yyyy",rdate)%></b></font>
		</td>
		<td align=center width=48><font size=1 face='ms sans serif, arial, geneva'>
			<a href="calendar.asp?date=">
			<!--<img src="images/calendar_today.gif" height=16 width=16 alt="Today" border=0>-->
			Today</a>
		</font></td>
		<td align=right width=48 nowrap>
			<a href="Calendar.asp?date=<%=DateAdd("m",1,rdate)%>"><img src="images/calendar_next_month.gif" alt="Next Month" height=16 width=16 border=0></a>&nbsp;&nbsp;
			<a href="Calendar.asp?date=<%=DateAdd("yyyy",1,rdate)%>"><img src="images/calendar_next_year.gif" alt="Next Year" height=16 width=16 border=0></a>
		</td>
	</tr>
	</table>

	<table border=1 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center><font size=1 face='ms sans serif, arial, geneva'>Sun</font></td>
		<td align=center><font size=1 face='ms sans serif, arial, geneva'>Mon</font></td>
		<td align=center><font size=1 face='ms sans serif, arial, geneva'>Tue</font></td>
		<td align=center><font size=1 face='ms sans serif, arial, geneva'>Wed</font></td>
		<td align=center><font size=1 face='ms sans serif, arial, geneva'>Thu</font></td>
		<td align=center><font size=1 face='ms sans serif, arial, geneva'>Fri</font></td>
		<td align=center><font size=1 face='ms sans serif, arial, geneva'>Sat</font></td>
	</tr>


<%
'---------------------------------------
'	Build the Calendar dynamically		
'	taking into account the events		
'	going on in that month (and adding	
'	links for those days.				
'---------------------------------------
CurDate = DateAdd("d",-(DatePart("w", FirstDayOfMonth)-1), FirstDayOfMonth)
DateLinkEnabled = false
'----------------------------------------------------
'	Loop through days to display, starting with the	 
'	first day of the first week of which to display	 
'----------------------------------------------------
do
	'-----------------------
	' Create calendar cell  
	'-----------------------
	calMonth = DatePart("m", CurDate)
	calDay = DatePart("d", CurDate)
	holiday = false
	if calDay = 1 then DateLinkEnabled = not DateLinkEnabled
	if DatePart("w",CurDate) = 1 then response.write("	<tr>" & NL)
	response.write("		<td align=left valign=top height=" & calCellHeight & " width=" & calCellWidth)
	if CurDate = Date and DateLinkEnabled then response.write(" bgcolor=yellow")
	if DatePart("m",CurDate) <> DatePart("m", rdate) then response.write(" bgcolor=gray")
	response.write("><a style='color:black; text-decoration: none' href='addnew.asp?show=yes&date=" & CurDate & "' onMouseOver=" & chr(34) & "top.status='Add new event'" & chr(34) & " onMouseOut=" & chr(34) & "top.status=''" & chr(34) & ">")
	response.write("<font size=1 face=arial>" & DatePart("d",CurDate) & "</font></a><br>")
	if DateLinkEnabled then response.write(CalendarDate(calDay))
	response.write("</td>" & NL)
	if DatePart("w",CurDate) = 7 then response.write("	</tr>" & NL)
	PrevDate = CurDate
	CurDate = DateAdd("d",1,CurDate)
loop until ((CurDate >= FirstDayOfNextMonth) and (DatePart("w",PrevDate) = 7))


'-----------------------------
'  Build list of event types  
'-----------------------------
Dim Legend(2), toggle
toggle = 0
set rs = DBQuery("select Type, ImageLoc from Event_Type order by Type")
while not rs.eof
	EventImage = imgTemp
	if not isnull(rs("ImageLoc")) then EventImage = rs("ImageLoc")
	Legend(toggle) = Legend(toggle) & "<img src='" & EventImage & "' height=16 width=16 border=0>&nbsp;&nbsp;" & rs("Type") & "<br>"
	toggle = toggle + 1
	if toggle > 2 then toggle = 0
	rs.movenext
	wend
rs.close
set rs = nothing
%>


</table>
</td></tr></table>

<table border=0><tr><td valign=top><font size=1>
<%=Legend(0)%>
</font></td><td width=30>&nbsp;</td><td valign=top><font size=1>
<%=Legend(1)%>
</font></td><td width=30>&nbsp;</td><td valign=top><font size=1>
<%=Legend(2)%>
</font></td></tr></table>

</center>


<!-- #include file="../../footer.asp" -->
