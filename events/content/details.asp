<%
response.buffer = true
action = request.form("action")
if action = "delete" then
    set rst = DBQuery("select * from Event_Contacts where Event_ID= " & request.form("id"))
    'If this event is in Event_Contacts (people to contact), delete it from there as well
    set rs = DBQuery("delete from Event_Contacts where Event_ID= " & request.form("id"))
	set rs = DBQuery("delete from Event_Date where Event_ID = " & request.form("id"))
	set rs = DBQuery("delete from Event where Event_ID = " & request.form("id"))
	Response.Redirect(Request("lastpage") & ".asp?date=" & Request("date"))
	end if
%>




<!--
Developer:    Kevin Dill
Date:         12/30/1998
Description:  Details about a specific event.
-->

<% pageTitle = "Details" %>

<!-- #include file="../section.asp" -->

<%
rdate = Request.QueryString("date")
if request("id") = "" or (not isnumeric(request("id"))) then
	response.write(ErrorMsg)
else

	dim EventDates, CreationDate, Creator_ID, Contact_ID, contact, creator, Description, Evnt_Dte
	dim Title, SubType, ImageLoc, Category, ID

	set rs = DBQuery("select e.CreationDate, e.Creator_ID, e.Contact_ID, e.Description, e.Title, contact.firstname as contactf, contact.lastname as contactl, creator.firstname as creatorf, creator.lastname as creatorl, est.Sub_Type, et.ImageLoc, et.Type from Event e, Employee contact, Employee creator, Event_Sub_Type est, Event_Type et where e.Event_ID = " & request("id") & " and e.Contact_ID = contact.EmployeeID and e.Creator_ID = creator.EmployeeID and e.Event_Sub_Type_ID = est.Event_Sub_Type_ID and est.Event_Type_ID = et.Event_Type_ID")
	
	if rs.eof then
		
		response.write(ErrorMsg)
		
	else
	
		CreationDate = rs("CreationDate")
		Creator_ID = rs("Creator_ID")
		Contact_ID = rs("Contact_ID")
		contact = rs("contactf") & " " & rs("contactl")
		creator = rs("creatorf") & " " & rs("creatorl")
		Description = rs("Description")
		Title = rs("Title")
		SubType = rs("Sub_Type")
		ImageLoc = rs("ImageLoc")
		Category = rs("Type")
		rs.close
	
		set rs = DBQuery("select EventDate, StartTime, EndTime from Event_Date where Event_ID = " & request("id") & " order by EventDate")
		while not rs.eof
			EventDates = EventDates & "<br>" & rs("EventDate") & " (" & WeekDayName(WeekDay(rs("EventDate"))) & ")"
			if (trim(rs("StartTime")) <> "") or (trim(rs("EndTime")) <> "") then
				EventDates = EventDates & ":&nbsp;&nbsp;&nbsp;" & rs("StartTime") & " - " & rs("EndTime")
				end if
			rs.movenext
			wend
		rs.close

		ID = Session("ID")
%>

<center>
<img src="<%=ImageLoc%>" height=16 width=16 alt="<%=Category%>" border=0>&nbsp;&nbsp;
<font color=navy><b><%=Title%></b></font>&nbsp;&nbsp;(<%=SubType%>)
<table border=0 cellpadding=0 cellspacing=8>
	<tr>
		<td colspan=3 valign=top><font size=1 face="ms sans serif, arial, helvetica">
		<b>Description:</b><br>
		<%=Description%>
		</font></td>
	</tr>
	<tr>
		<td valign=top nowrap><font size=1 face="ms sans serif, arial, helvetica">
			<b>Dates:</b><%=EventDates%></font></td>

		<td width=10>&nbsp;</td>
		
		<td valign=top nowrap><font size=1 face="ms sans serif, arial, helvetica">
			<b>Contact:</b>&nbsp;&nbsp;<a href="../../directory/content/details.asp?EmpId=<%=Contact_ID%>"><%=contact%></a><br>
			Created <%=FormatDateTime(CreationDate,2)%> by:&nbsp;&nbsp;<a href="../../directory/content/details.asp?EmpId=<%=Creator_ID%>"><%=creator%></a>
			</font></td>
	</tr>
</table>


<script language=javascript>
<!--
function editEvent(){
	if (confirm("You have selected to EDIT this event...")){
		window.location.href = "edit.asp?id=<%=request("id")%>&date=<%=Server.URLEncode(rdate)%>&lastpage=<%=request("lastpage")%>"
		}
	}
function deleteEvent(){
	if (confirm("You have selected to DELETE this event...")){
		document.frmInfo.action.value = "delete"
		document.frmInfo.submit()
		}
	}
function goBack(){
	window.location.href='<% =request("lastpage") & ".asp?date=" & rdate %>'
	}
// -->
</script>

<form name="frmInfo" method=post action="details.asp">
<%if (Creator_ID = ID) or (Contact_ID = ID)  or hasRole("WebMaster") then %>
<input type=button class=button value="  Edit  " onClick="editEvent();">
<input type=button class=button value="Delete" onClick="deleteEvent();">
<% end if %>
<input type=button class=button value=" Back " onClick="window.location.href='calendar.asp?date=<%date%>'">
	<input type=hidden name="action" value="">
	<input type=hidden name="id" value="<% =request("id") %>">
	<input type=hidden name="date" value="<% =rdate %>">
	<input type=hidden name="lastpage" value="<% =request("lastpage") %>">
</form>
</center>

<%
		end if
	end if
%>

<!-- #include file="../../footer.asp" -->
