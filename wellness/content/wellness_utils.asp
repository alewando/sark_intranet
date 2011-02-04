<!--
Developer:    Adam Lewandowski
Date:         10/9/2000
Description:  Wellness program utility functions
-->

<%

Function getActivityDays(empID)
 sql = "SELECT COUNT(Activity_Date) activityCount FROM wellness_activity WHERE employee_id="& empID & _
       " AND activity_date>='" & wellnessStartDate & "' AND activity_date <= '" & _
       wellnessEndDate & "' AND sark_sponsored=0"
 set rsCount = DBQuery(sql)
 activityCount = rsCount("activityCount")
 sql = "SELECT COUNT(Activity_Date) activityCount FROM wellness_activity WHERE employee_id="& empID & _
       " AND activity_date>='" & wellnessStartDate & "' AND activity_date <= '" & _
       wellnessEndDate & "' AND sark_sponsored=1"
 set rsCount = DBQuery(sql)
 activityCount = activityCount + (2 * rsCount("activityCount"))
 getActivityDays = activityCount
End Function

' Adds the activity records to the database
Sub addRecords()
 addDate = Request("dateSelected")
 
 ' Is this date a sark sponsored event?
 sql = "SELECT count(*) c FROM Wellness_Events WHERE event_date='" & addDate & "'"
 set rsCount = DBQuery(sql)
 sarkSponsored=false
 if not rsCount.eof then
  if rsCount("c") > 0 then
   sarkSponsored=true
  end if
 end if

 ' Update / Delete the requested date
 sql = "SELECT sark_sponsored FROM Wellness_Activity WHERE employee_id=" & empID & " AND activity_date='" & addDate & "'"
 'Response.Write("sql=" & sql & "<BR>")
 set rsAdd = DBQuery(sql)
 if rsAdd.eof then
  sql = "INSERT INTO Wellness_Activity (employee_id, activity_date, sark_sponsored) VALUES (" & empID & ",'" & addDate & "', 0)"
  'Response.Write("sql=" & sql & "<BR>")
  DBQuery(sql)
 elseif rsAdd("sark_sponsored")=0 then
  if sarkSponsored then
	sql = "UPDATE Wellness_Activity SET sark_sponsored=1 WHERE employee_id=" & empID & " AND activity_date ='" & addDate & "'"   
	'Response.Write("sql=" & sql & "<BR>")
	DBQuery(sql)    
  else
   sql = "DELETE FROM Wellness_Activity WHERE employee_id=" & empID & " AND activity_date ='" & addDate & "'"   
   'Response.Write("sql=" & sql & "<BR>")
   DBQuery(sql)
  end if
 elseif rsAdd("sark_sponsored")=True then
  sql = "DELETE FROM Wellness_Activity WHERE employee_id=" & empID & " AND activity_date ='" & addDate & "'"   
  'Response.Write("sql=" & sql & "<BR>")
  DBQuery(sql)
 end if
End Sub

Sub showCalendar(calMonth, calYear, empID)

 monthNames = array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")

 Response.Write("<!-- Calendar: month=" & calMonth & ", year=" & calYear & " -->") 
 Dim dates
 set dates = CreateObject("Scripting.Dictionary")
 Dim sarkSponsoredDates
 set sarkSponsoredDates = CreateObject("Scripting.Dictionary")
 Dim eventDescriptions
 set eventDescriptions = CreateObject("Scripting.Dictionary") 
 Dim otherDates
 set otherDates = CreateObject("Scripting.Dictionary")
 
 ' Retrieve prior activity dates
 sql = "SELECT rtrim(CONVERT(char, activity_date, 101)) as date, sark_sponsored FROM wellness_activity " & _
       "WHERE employee_id = " & empID
 set rsDates = DBQuery(sql) 
 While not rsDates.eof
  'Response.Write("Adding activity date &quot;" & rsDates("date") & "&quot;<BR>")
  dates.Add Cstr(rsDates("date")),"1" 
  if rsDates("sark_sponsored") then
   sarkSponsoredDates.Add Cstr(rsDates("date")), "1"
   'Response.Write("Add sark-sponsored date: " & Cstr(rsDates("date")) & "<BR>")
  end if 
  rsDates.moveNext
 Wend
 
 ' Retrieve sark-sponsored activity dates
 sql = "SELECT rtrim(CONVERT(char, event_date, 101)) as date, Event_Description FROM Wellness_Events WHERE sark_sponsored=1" 
 set rsEvents = DBQuery(sql)
 Dim sarkEvents
 set sarkEvents = CreateObject("Scripting.Dictionary")
 While not rsEvents.eof
  sarkEvents.Add Cstr(rsEvents("date")),"1" 
  eventDescriptions.Add Cstr(rsEvents("date")), Cstr(rsEvents("Event_Description"))
  'Response.Write("Sark Event on:"&Cstr(rsEvents("date"))&"<BR>")
  rsEvents.moveNext
 Wend
 
 ' Retrieve other activity dates
 if Session("ID") = 24 then
  otherDates.Add "10/21/2000","1"
 end if

%>

<FORM ACTION="activity.asp" METHOD=POST NAME=cal>
<INPUT TYPE=HIDDEN NAME=action VALUE=addRecords>
<INPUT TYPE=HIDDEN NAME=empID VALUE=<%=empID%>>
<INPUT TYPE=HIDDEN NAME=month VALUE=<%=calMonth%>>
<INPUT TYPE=HIDDEN NAME=year VALUE=<%=calYear%>>
<CENTER>
<TABLE BORDER=0>
<%
 startDate = CDate(wellnessStartDate)
 endDate = CDate(wellnessEndDate)
 dispDate = CDate(calMonth & "/01/" & calYear)
 nowDOM=Cstr(Day(Now))
 if len(nowDOM)=1 then
  nowDOM = "0"+nowDOM
 end if
 today = Month(Now) & "/" & nowDOM & "/" & Year(Now)
 If dispDate < startDate or dispDate > endDate Then
  %>
  <TR><TD ALIGN=CENTER COLSPAN=3 BGCOLOR=#FFFFFF>
   <FONT COLOR=RED><B>Note:</B> The Wellness Program runs from <%=startDate%> to <%=endDate%><BR>
   Dates marked on this calendar will not count towards your program goal</FONT>
  </TD></TR>
  <%
 End If
 %>
 <TR><TD ALIGN=CENTER COLSPAN=3 BGCOLOR=#FFFFFF>
 <B><%=monthNames(DatePart("m", dispDate)-1)%>&nbsp;<%=DatePart("yyyy", dispDate)%></B>
 </TD></TR>
<%
  ' Links to previous/next month's calendar
  Response.Write("<TR>")
  prevMonth = calMonth-1
  if prevMonth=0 then
   prevMonth=12
  end if
  if prevMonth=12 then
   prevYear = calYear-1
  else 
   prevYear = calYear
  end if 
  
  Response.Write("<TD align=left><A HREF='activity.asp?month=" & prevMonth & "&year=" & prevYear & "'> &lt;&lt; Prev</A></TD>")
  Response.Write("<TD></TD>")
  
  nextMonth=calMonth+1
  if nextMonth=13 then
   nextMonth=1
  end if
  if nextMonth=1 then 
   nextYear=calYear+1
  else
   nextYear=calYear
  end if
  Response.Write("<TD align=right><A HREF='activity.asp?month=" & nextMonth & "&year=" & nextYear & "'>Next &gt;&gt;</A></TD>")
  
  Response.Write("</TR>") 
  %>
 <TR><TD ALIGN=CENTER COLSPAN=3 BGCOLOR=#FFFFFF>
  <TABLE BGCOLOR="#000000" CELLPADDING=1><TR><TD>
   <TABLE BORDER=0 CELLPADDING=3>
   <%
    ' Print days of week
    Response.Write("<TR>")
    dow = array("Sun&nbsp;", "Mon&nbsp;", "Tue&nbsp;", "Wed&nbsp;", "Thu&nbsp;", "Fri&nbsp;", "Sat&nbsp;")
    For i = 1 to 7
     Response.Write("<TH BGCOLOR=#FFFFFF>" & dow(i-1) & "</TH>")
    Next
    Response.Write("</TR>")
    
    ' Print out empty cells until we get to the first of the month
    Response.Write("<TR>")
    startDOW = DatePart("w", dispDate)-1
    dayCount=0
    While dayCount<startDOW 
     Response.Write("<TD BGCOLOR=#FFFFFF>&nbsp;</TD>")
     dayCount = dayCount+1
    Wend
    
    ' Print out cells of the days of the month, breaking for new weeks
    daysInMonths = array(31,29,31,30,31,30,31,31,30,31,30,31)
    daysInMonth = daysInMonths(DatePart("m",dispDate)-1)
    dom = 1
    While dom <= daysInMonth
     if len(dom)=1 then
      dom = "0" & dom
     end if
     dateVal = calMonth & "/" & dom & "/" & calYear
     'Background color of cells depends on previous activity
     if Dates.Exists(dateVal) then
      if otherDates.Exists(dateVal) then 
       Response.Write("<TD class=otherActivityDate>")
      elseif sarkSponsoredDates.Exists(dateVal) then
	   Response.Write("<TD class=sarkActivityDate>")
      else     		
       Response.Write("<TD class=activityDate>")
      end if
     else
      Response.Write("<TD class=date>")
     end if
     'Response.Write("<span>")
     if dateVal = today then
      Response.Write("[")
     end if
     Response.write("<A ")
     if sarkEvents.Exists(dateVal) then
      Response.Write("STYLE='color:red' ")
     end if
     if eventDescriptions.Exists(dateVal) then
      Response.Write("TITLE='" & CStr(eventDescriptions.Item(dateVal)) & "' ")
     end if
     Response.Write("HREF='activity.asp?action=addRecords&month=" & calMonth & "&year=" & calYear & "&dateSelected="&dateVal&"'>" & dom & "</A>")
     if sarkEvents.Exists(dateVal) then
      Response.Write("</FONT>")
     end if
     if dateVal = today then
      Response.Write("]")
     else
      'Response.Write("dateVal=" & dateVal & ", today=" & today)
     end if
     Response.Write("</TD>" & vbCRLF)
     dom = dom + 1
     dayCount = (dayCount+1) mod 7
     if dayCount = 0 then
      Response.Write("</TR><TR>")
     end if
    Wend
    
    'Fill out the rest of the table with empty cells
    while dayCount <> 0 and dayCount < 7
     Response.Write("<TD BGCOLOR=#FFFFFF>&nbsp</TD>")
     dayCount = dayCount+1
    wend
    Response.Write("</TR>")
   %>
   </TABLE>
  </TD></TR>
  </TABLE>
 </TD></TR>
 <%

   
%>
</TABLE>
 <P>
 <B>Directions:</B> To log activity for a date, click on it. Click again for a Sark-sponsored 
 activity. Click a third time to remove the date.<BR>
 <TABLE BORDER=0>
  <TR><TD WIDTH=20 CLASS=activityDate>&nbsp;</TD><TD>= Completed Day</TD></TR>
  <TR><TD WIDTH=20 CLASS=sarkActivityDate>&nbsp;</TD><TD>= Sark sponsored Activity</TD></TR>
 </TABLE>
</CENTER>
</FORM>

<%
End Sub
%>