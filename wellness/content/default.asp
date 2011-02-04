
<%
Response.Expires=0                                ' - works with IE 4.0 browsers. 
Response.AddHeader "Pragma","no-cache"            ' - works with proxy servers. 
Response.AddHeader "cache-control", "no-store"    ' - works with IE 5.0 browsers. 
%>
<!--
Developer:    Adam Lewandowski
Date:         10/9/2000
Description:  Wellness program main page
-->
<!-- #include file="../section.asp" -->
<!-- #include file="wellness_utils.asp" -->

<STYLE>
BODY {
 font-family: ms sans serif, arial, geneva;
 font-size: 30pt;
}
.EmployeeHeader {
 font-family: verdana, arial, geneva;
 font-weight: bold;
 color: blue;
}
</STYLE>

<BODY BGOLOR="#FFFFFF">

<CENTER>
<H4>Wellness Program</H4>
</CENTER>
<HR>
<%
 empID = Session("ID")
 ' Get employee name
 sql = "SELECT FirstName, LastName FROM Employee WHERE employeeID = " & empID
 set rsEmp = DBQuery(sql)
 ls_name = rsEmp("FirstName") & "&nbsp;" & rsEmp("LastName")
%>
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=100%>
<TR><TD COLSPAN=2><CENTER><span class="EmployeeHeader"><%=ls_name%></span></CENTER><BR></TD></TR>
<TR>
 <TD WIDTH=50% VALIGIN=TOP>
 <!-- Shirt Size selector -->
 <%
  ' Update employee record with shirt size
  If Request.Form("Action") = "addShirtSize" Then
   sql = "UPDATE Employee SET ShirtSize='" & Request.Form("shirtSize") & "' WHERE employeeid =" & empID
   Response.Write("<!-- SQL: " & sql & " -->")
   set rs = DBQuery(sql)
  End If
  ' Display shirt size choices if necessary
  sql = "SELECT ShirtSize FROM Employee WHERE employeeid=" & empID
  set rs = DBQuery(sql)
  If rs.eof or IsNull(rs("ShirtSize")) or rs("ShirtSize")="" Then
   %>
   <H3>Shirt Size</H3>
   Select your shirt size:
   <FORM ACTION="<%=Request.ServerVariables("SCRIPT_NAME")%>" METHOD=POST>
   <INPUT TYPE=HIDDEN NAME="Action" VALUE="addShirtSize">
   <SELECT NAME="shirtSize">
   <OPTION VALUE="S">S</OPTION>
   <OPTION VALUE="M">M</OPTION>
   <OPTION VALUE="L">L</OPTION>
   <OPTION VALUE="XL">XL</OPTION>
   <OPTION VALUE="XXL">XXL</OPTION>
   </SELECT>
   <INPUT TYPE=SUBMIT VALUE="OK">
   </FORM>
   <%
  End If
 %> 
  <!-- Progress Bar -->
  <%
   '-------------------------------
   ' Calculate progress bar values
   '-------------------------------
   activity_count = getActivityDays(empID)
   total_days = round(DateDiff("d", wellnessStartDate, wellnessEndDate) / 2)
   days_left =  total_days - activity_count
   total_bar_width = 150
   completed_bar_percent = activity_count / total_days
   completed_bar_width = total_bar_width * completed_bar_percent
   empty_bar_width = total_bar_width - completed_bar_width   
   percent = round(completed_bar_percent * 100)
  %>
  <H4> Progress: </H4>
  <center>
  <IMG SRC="images/bar_red.jpg" WIDTH=<%=completed_bar_width%> HEIGHT=20
   ALT="<%=activity_count%> days complete" BORDER=0 HSPACE=0><IMG
   SRC="images/bar_gray.jpg" WIDTH=<%=empty_bar_width%> HEIGHT=20
   ALT="<%=days_left%> days to go" BORDER=0 HSPACE=0>
  <P>
  <font size="-1">
  <%
   If percent >= 100 Then
    Response.Write("<B>Congratulations!<B> You have completed the Wellness Program!<BR>")
   else
    Response.Write(activity_count & " / " & total_days & " (" & percent & "%)<BR>")
    Response.Write("You have " & days_left & " days of excercise left. <BR>")
   End If
  %>
  </center>
  </font>
  
  <P>
  <CENTER><A HREF="activity.asp">Activity Tracking</A><BR></CENTER>
 </TD>
 <TD WIDTH=50%>
  <!-- Upcoming events -->
  <H4>Upcoming events</H4>
  <UL>
  <%
   sql = "SELECT event_date, event_description, event_link, sark_sponsored FROM Wellness_Events " & _
         "WHERE event_date >= getDate() ORDER BY event_date"
   set rsEvents = DBQuery(sql)
   While not rsEvents.eof
    if rsEvents("event_link") <> "" then
	%>
	<LI><A HREF="<%=rsEvents("event_link")%>" TARGET="_blank"><%=rsEvents("event_description")%></A> -
	 <%=rsEvents("event_date")%>
    <% 
    else 
    %>
	<LI><%=rsEvents("event_description")%></A> - <%=rsEvents("event_date")%>
	<%
	end if 
   	rsEvents.moveNext
   Wend
  %>
  </UL>
  
  <!-- Current Leaders -->
  <HR>
  <H4>Current Leaders</H4>
  <UL>
  <%
   sql = "SELECT firstName + ' ' + lastName as name, employee_id, COUNT(activity_date) + " &_
         "(SELECT COUNT(*) " & _
         " FROM wellness_activity " & _
         " WHERE employee_id = a.employee_id AND " & _
         "  sark_sponsored = 1 " & _
         ") as activityCount " & _
         "FROM Wellness_activity a, employee e " & _
         "WHERE a.employee_id = e.employeeid AND " & _
         "activity_date >= '10/1/2000' AND " & _
         "activity_date <= '12/31/2000' " & _
         "GROUP BY firstName + ' ' + lastName, employee_id " & _
         "ORDER BY activityCount desc"
   set rsCount = DBQuery(sql)
   count = 0
   While not rsCount.eof and count<7
    Response.Write("<LI>")
    Response.Write("<A HREF='../../directory/content/details.asp?EmpID=" & rsCount("employee_id") & "'>" & rsCount("name") & "</A>")
    Response.Write(" - " & rsCount("activityCount") & " days. </LI>")
    count = count+1
    rsCount.moveNext
   Wend
  %>
  </UL>
 </TD>
</TR>
</TABLE>

<!-- #include file="../../footer.asp" -->
