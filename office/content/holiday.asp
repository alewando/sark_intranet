<!--
Developer:    Kevin Dill
Date:         11/09/98
Description:  Holiday Schedule
History:      01/19/1999 - Changed so that it is dynamic
-->


<!-- #include file="../section.asp" -->

<%
holidayID = 10
ryear = request("year")
rmonth = request("month")
if ryear = "" then ryear = Year(Date)
%>

<table align=center cellpadding=5>
	<tr><td colspan=2 align=center><FONT SIZE=1 face="ms sans serif, arial, geneva"><B>
		<% = ryear %> Holiday Schedule</B><br>&nbsp;
		</td></tr>
<%
sql = "select e.event_id, e.title, ed.eventdate from event e, event_date ed " & _
		"where e.event_sub_type_id = " & holidayID & _
		" and e.event_id = ed.event_id" & _
		" and DATEPART(year, ed.eventdate) = " & ryear & _
            " and ed.eventdate >= '" & date() & "'" &  _
		" order by ed.eventdate"
set rs = DBQuery(sql)
eventID = 0
if rs.eof then response.write("<tr><td colspan=""2""><font size=""1"" face=""ms sans serif, arial, geneva"">No remaining holidays in " & ryear & ".</font></td></tr>")
while not rs.eof
	curEventID = rs("event_id")
	curDate = rs("eventdate")
%>

<%	if eventID <> curEventID then %>
	<tr><td valign=top><FONT SIZE=1 face="ms sans serif, arial, geneva">
		<% = rs("title") %>&nbsp;&nbsp;</FONT></td>
		<td align=right><FONT SIZE=1 face="ms sans serif, arial, geneva"><%
	else
		response.write("<br>")
		end if
	response.write(WeekDayName(WeekDay(curDate)) & ", " & MonthName(Month(curDate)) & " " & Day(curDate))
	eventID = rs("event_id")
	rs.movenext
	if rs.eof then
		response.write("</FONT></td></tr>")
	else if eventID <> rs("event_id") then 
		response.write("</FONT></td></tr>")
		end if
		end if
	wend
rs.close
%>
</table>

<br>

<% ryear = ryear + 1 %>

<table align=center cellpadding=5>
	<tr><td colspan=2 align=center><FONT SIZE=1 face="ms sans serif, arial, geneva"><B>
		<% = ryear %> Holiday Schedule</B><br>&nbsp;
		</td></tr>
<%
sql = "select e.event_id, e.title, ed.eventdate from event e, event_date ed " & _
		"where e.event_sub_type_id = " & holidayID & _
		" and e.event_id = ed.event_id" & _
		" and DATEPART(year, ed.eventdate) = " & ryear & _
		"order by ed.eventdate"
set rs = DBQuery(sql)
eventID = 0
while not rs.eof
	curEventID = rs("event_id")
	curDate = rs("eventdate")
%>

<%	if eventID <> curEventID then %>
	<tr><td valign=top><FONT SIZE=1 face="ms sans serif, arial, geneva">
		<% = rs("title") %>&nbsp;&nbsp;</FONT></td>
		<td align=right><FONT SIZE=1 face="ms sans serif, arial, geneva"><%
	else
		response.write("<br>")
		end if
	response.write(WeekDayName(WeekDay(curDate)) & ", " & MonthName(Month(curDate)) & " " & Day(curDate))
	eventID = rs("event_id")
	rs.movenext
	if rs.eof then
		response.write("</FONT></td></tr>")
	else if eventID <> rs("event_id") then 
		response.write("</FONT></td></tr>")
		end if
		end if
	wend
rs.close
%>
</table>

<!-- #include file="../../footer.asp" -->
