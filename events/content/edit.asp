<%
response.buffer = true
Dim sessItems  'Keep track of items throughout session
'-------------------------------------------------------
'  Check to see if we should save info to the database  
'  and redirect to the calendar page.                   
'-------------------------------------------------------
if request.form("final") <> "" then
	'INSERT CODE to extract out the values and save in database
	sql = "select employeeID from employee where username ='" & Session("userName") & "'"
	set rs = DBQuery(sql)
	sql = "Update Event set Title = '" & replace(request.form("ShortDesc"),"'","''") & "', Description = '" & replace(request.form("LongDesc"),"'","''") & "', Contact_ID =" & Request("admin") & ", Event_Sub_Type_ID =" & request.form("type") & ", SignUp = 0, Creator_ID = " & rs(0) & ", CreationDate = getdate() "
	sql = sql & "where event_id = " & Request.form("eventid")
	set rs = DBQuery(sql)
	'We have new information to insert into the database, so we must delete the old stuff. However,
	'to delete the old stuff we must delete the contacts as well, even if we don't want to. So here we do
	'our deletes, and then we'll add them in the for loop.
	DBQuery("Delete from Event_Contacts where Event_ID=" &Request.Form("eventid"))
	DBQuery("Delete from Event_Date where Event_ID = " & Request.form("eventid"))

	for each x in request.form("setdate")
		sql =	"INSERT INTO Event_Date (Event_ID, EventDate, StartTime, EndTime) " & _
				"VALUES (" & request.form("eventid") & ",'" & x & "','" & request.form(x & "start") & "','" & request.form(x & "end") & "')"
		set rs = DBQuery(sql)
		set rst = DBQuery("select max(Event_Date_ID) from Event_Date")
		edate = rst(0)
		'Add the users to be notified if reminders are checked
		if Request.Form("Remndr") = "on" then
		   set reqFrmSelected = Request.Form("Selected")
		   for each myName in reqFrmSelected
		      'namehere = Request.Form("TempArrayUName(5)")
		      newsql = "INSERT INTO Event_Contacts (Event_ID, Event_Date_ID, Employee_User_Name, EventDate, Email_Days_Before) " &_
		               "VALUES (" & Request.Form("eventid") & "," & edate & ",'" & myName & "','" & x & "'," & Cint(Request.Form("DaysInAdv")) & ")"
		      set rst = DBQuery(newsql)
		      'We want to fire off an e-mail now notifying everyone we selected of the event.
		      txtBody = session("Longname") & " has modified an event scheduled for " &x& ", and has requested that you be notified. You will receive a reminder e-mail " &Request.Form("DaysInAdv")& " days before the event."
		      txtBody = txtBody & chr(10) & chr(10) & "--" & Request.Form("ShortDesc")
		      txtBody = txtBody & chr(10) & "--" & Request.Form("LongDesc")
		      txtSubject = "Event Notification"
		      txtRecipients = "" &myName& "@sark.com"
		      txtCC = ""
		      SendMail txtSubject, txtRecipients, txtCC, txtBody
		   next
		end if
	next
	
	response.redirect("calendar.asp?date=" & server.URLEncode(request.form("rdate")))
	
	end if
pageTitle = "Edit"
%>



<!-- #include file="../section.asp" -->


<%
if request.form("setdate") = "" then
	'---------------------------------------------
	'  Display the first page to allow a user to  
	'  enter a new event and select the dates.    
	'---------------------------------------------
%>

<script language=javascript>
<!--
var items = 0; 
function validate(aNumber){
   	var result = true
	if (result && (document.frmInfo.ShortDesc.value == "")){
		alert("Please enter a short title for this event.")
		document.frmInfo.ShortDesc.focus()
		result = false
		}
	if (result && (document.frmInfo.ShortDesc.value == "")){
		alert("Please enter a more detailed description for this event.")
		document.frmInfo.ShortDesc.focus()
		result = false
		}
    if (result && (aNumber != 0)) {
       items = aNumber
    }
	if (result && (items < 1)){
		alert("Please select one or more days.");
		result = false
		}
	return result
	}
function toggle(calDay){
	if (calDay.checked){
		items++
		}
	else{
		items--
		}
	}
// -->
</script>

<form name=frmInfo method=post action="edit.asp" onSubmit="return (validate(<%=sessItems%>));">
<input type=hidden name=eventid value="<%=Request("id")%>"></input>
<table border=0><tr><td>
<%
'---------------------------------------
'	Build the Calendar dynamically		
'	taking into account the events		
'	going on in that month (and adding	
'	links for those days.				
'---------------------------------------
rdate = request.querystring("date")
if rdate = "" then rdate = Date
showDate = (request.querystring("show") <> "")
FirstDayOfMonth = DateAdd("d",-(DatePart("d", rdate)-1), rdate)
CurDate = DateAdd("d",-(DatePart("w", FirstDayOfMonth)-1), FirstDayOfMonth)
response.write("<table border=1 cellspacing=0>")
response.write("<tr bgcolor=silver><td colspan=7 align=center><font size=1><b>" & MonthName(DatePart("m", rdate)) & " " & DatePart("yyyy", rdate) & "</b></font></td></tr>")
response.write("<tr bgcolor=gray><td align=center><b>s</b></td><td align=center><b>m</b></td><td align=center><b>t</b></td><td align=center><b>w</b></td><td align=center><b>t</b></td><td align=center><b>f</b></td><td align=center><b>s</b></td></tr>")
DateLinkEnabled = false

	
do
	'-----------------------
	' Create calendar cell  
	'-----------------------
	calDay = DatePart("d", CurDate)
	if calDay = 1 then DateLinkEnabled = not DateLinkEnabled
	if DatePart("w",CurDate) = 1 then response.write("	<tr>" & NL)
	response.write("<td")
	if DatePart("m",CurDate) <> DatePart("m", rdate) then
		response.write(" bgcolor=silver")
	else
		response.write(" bgcolor=#dddddd")
	end if
	response.write("><font size=1 color=navy face='arial'>")
	response.write(DatePart("d",CurDate) & "<br>")
	sql = "select EventDate from Event_Date where Event_Id = " & Request("id") 
	set checkedrs = DBQuery(sql)
	response.write("<input type=checkbox name=setdate value=" & CurDate)
	while not checkedrs.eof 
		if checkedrs(0) = CurDate then 
			response.write(" checked " & NL)
			end if
		checkedrs.movenext()
		wend
	Response.Write(" onClick='toggle(this);'></font></td>")
	checkedrs.close()
	checkedrs = null
	
	if DatePart("w",CurDate) = 7 then response.write("</tr>" & NL)
	PrevDate = CurDate
	CurDate = DateAdd("d",1,CurDate)
	DisplayLink = false
loop until ((DatePart("m",PrevDate) > DatePart("m",rdate) or DatePart("yyyy",PrevDate) > DatePart("yyyy",rdate)) and DatePart("w",PrevDate) = 7)
response.write("</table>")
%>

<script language=javascript>
<!--
for (i=0; i < document.frmInfo.elements.length; i++){
	element = document.frmInfo.elements[i]
	if ((element.type == "checkbox") && (element.checked)) items++
	}
// -->
</script>

<input type=hidden name=rdate value="<%=rdate%>">

</td><td width=20>&nbsp;</td><td align=left valign=top>

<font size=1 color=navy>
Please fill in the following information and check the day(s) on which
the event will occur.  Both the creator and the designated contact will
be able to edit the event later.
</font><hr>

<table border=0>
	<tr><td><font size=1 color=maroon>Contact:</font></td><td><select name=admin>
<%
sessionID = 0
sql = "select Contact_ID from event where event.event_ID = " & Request("id")
Response.Write(sql)
set rsexists = DBQuery(sql)
set rs = DBQuery("select employeeid, firstname, lastname, username from employee order by lastname, firstname")
while not rs.eof

Response.Write("<option value=" & rs("employeeid"))
	if rsexists(0) = rs(0) then
		sessionID = rs("employeeid")

		Response.write(" SELECTED>")
	else
		Response.Write("> ")
	end if
	response.write(rs("lastname") & ", " & rs("firstname") & "</option>" & NL)
	
	rs.movenext
	wend
rs.close()
rsexists.close()
%>
</select></td></tr>

	<tr><td><font size=1 color=maroon>Type:</font></td><td><select name=type>
<%

set rsexists = DBQuery("select et.event_sub_type_ID from Event et, Event_Sub_Type est where et.Event_Sub_Type_ID = est.Event_Sub_Type_ID and et.Event_ID = " & Request("id"))
set rs = DBQuery("select et.event_type_id, est.event_sub_type_id, et.type, est.sub_type from event_type et, event_sub_type est where est.event_type_id = et.event_type_id order by type, sub_type")

while not rs.eof
	response.write("<option value='" & rs("event_sub_type_id") & "'")
	if rs(1) = rsexists(0) then
	Response.Write(" SELECTED ")
	end if
	Response.Write(" >" & rs("type") & ", " & rs("sub_type") & "</option>" & NL)
	rs.movenext
wend
	
rs.close()
rsexists.close()

sql = "select title, description, employee.username from event, employee where event_ID = " & Request("id") & " and employee.employeeID = event.creator_ID"
set rs = DBQuery(sql)
%>
</select></td></tr>
	<tr><td><font size=1 color=maroon>Title:</font></td><td><input type=text name=ShortDesc maxlength=50 value="<%=trim(rs(0))%>"></td></tr>
	<tr><td><font size=1 color=maroon>Desc:</font></td><td><input type=text name=LongDesc maxlength=255 value="<%=trim(rs(1))%>"></td></tr>
	<tr><td><font size=1 color=maroon>Creator:</font></td><td><b><%=lcase(rs(2))%></b></td></tr>
</table>
<%
rs.close()

set rs = DBQuery("select * from event_date where event_id = " & Request("id"))
while not rs.eof
	response.write("<input type=hidden name=" & FormatDateTime(rs("eventdate"),2) & "start value='" & rs("StartTime") & "'>" & NL)
	response.write("<input type=hidden name=" & FormatDateTime(rs("eventdate"),2) & "end value='" & rs("EndTime") & "'>" & NL)
	rs.movenext
	wend
%>
<input type=hidden name="LastPage" value="<%=request("lastpage")%>">

<br><input type=button class=button value="Cancel" onClick="window.location.href='calendar.asp?date=<%=rdate%>'" id=button1 name=button1><%if sessionID > 0 then%>
<input type=submit class=button value="Next >>" id=submit1 name=submit1><input type=hidden name=creator value="<%=sessionID%>"><%end if%>

</td></tr></table>
</form>



<%
	else
	'-------------------------------------------------
	'  Display event confirmation and allow the user  
	'  to enter start and end times for each date.    
	'-------------------------------------------------
   
   'We want to set the number of items for the session in case we need to go back
   for each d in Request.Form("setdate")
      'Response.Write("Test")
      sessItems=sessItems+1
   next
%>


<center>
<form name="frmInfo" method=post action="edit.asp">
<input type=hidden name=final value=yes>
<input type=hidden name=eventid value="<%=Request.form("eventid")%>"></input>
<input type=hidden name=rdate value="<%=request.form("rdate")%>"></input>


<font color=navy>Please fill in the start/end times for the selected dates:</font><p>

<table width="90%" border=0><tr><td width="50%" valign=top><font size=1>
	<b>Title:</b> <%=request.form("ShortDesc")%><input type=hidden name=ShortDesc value="<%=request.form("ShortDesc")%>"><br>
	<b>Description:</b> <%=request.form("LongDesc")%><input type=hidden name=LongDesc value="<%=request.form("LongDesc")%>">

</font></td><td width=10>&nbsp;</td><td width="50%" valign=top nowrap><font size=1>

<table border=0>
<%
for each x in request.form("setdate")
	response.write("<tr><td align=right nowrap><font size=1>" & x & "</font></td><td nowrap><font size=1><input type=hidden name=setdate value='" & x & "'>")
	response.write("<input type=text name='" & x & "start' size=7 maxlength=10 value='" & request.form(x & "start") & "'>")
	response.write(" to <input type=text name='" & x & "end' size=7 maxlength=10 value='" & request.form(x & "end") & "'>")
	response.write("</font></td></tr>" & NL)
	next%>
</table>

</font></td></tr>

<%'Checkbox for selecting that the user wants to send an e-mail reminder to people%>

<tr><td colspan=3><font size=1>
	<br>
	<input type=checkbox name="Remndr" CHECKED><b>Check here</b> to send an automatic e-mail reminder to the people you select below. (Optional).
	<br>
	</font></td>
</tr><%
'This is the number of days in advance of the event that the user wants to send the e-mail
'set nrq=DBQuery("select Email_Days_Before from EventContacts where Event_ID="&Request.Form("eventid")&"")
%>
<tr><td colspan=3><font size=1>
	<br>
	<input type=text name="DaysInAdv" size=2 maxlength=3 value=3><b>Enter</b> the number of days in advance of the event that this e-mail should be sent.
	<br>
	</font></td>
</tr>

</table>
<%
' This section sets up the table from which we can choose the names of people we wish to receive
' e-mail reminders. See the Skills_Search.asp document in the documents folder to view the basis of
' this section.
%>
<table class=tableShadow width=100% cellspacing=0 cellpadding=2 border=0 bgcolor=#ffffcc>
                   <th width='45%'></th>
                   <th width='10%'></th>
                   <th width='45%'></th>

   <tr> 
      <td valign=bottom><font size=2 face= "ms sans serif, arial, geneva"  color = black><b>Not Selected</b></td>
      <td valign=bottom><font size=2 face= "ms sans serif, arial, geneva"  color = black>&nbsp;</td>
      <td valign=bottom><font size=2 face= "ms sans serif, arial, geneva"  color = black><b>Selected<br></b></td>
   </tr>

   <TR>
      <TD colspan=3><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD>
   </TR>  

   <tr>
      <td><select MULTIPLE size=10 name="NotSelected" id=NotSelected style="{ width: 100%; }" onfocus="setSelectList(document.frmInfo.Selected,false)";> 
 
      <%
      'Here we are just querying the database for employee names and user names
      'These will be displayed in the table under "Not Selected" so they can be selected
              
      mysql = "select a.lastname, a.firstname, a.username from employee a where a.username not in (select Employee_User_Name from Event_Contacts where Event_ID='"&Request.form("eventid")&"' ) order by a.lastname"
    
      'Notice all of the sizes here are hard coded. Since we don't have 300 people in the branch,
      'this should not cause a problem. However, if you need more than 300, you'll have to change
      'the sizes here.
      Dim TempArrayFName(300)  'store the person's first name
      Dim TempArrayLName(300)  'store the person's last name
      Dim TempArrayUName(300)  'store the person's user name
      set rsName = DBQuery(mysql) 'result set
      Dim TempSelected(300)    'Selection properties
    
      'Manipulate the result set
      i=0
      While Not rsName.eof
        TempArrayFName(i) = rsName("FirstName")
        TempArrayLName(i) = rsName("LastName")
        TempArrayUName(i) = rsName("Username")
        TempSelected(i) = false
       
        'Now print the names out
        ' Response.Write "'" & TempArrayLName(i) & ", " & TempArrayFName(i) &"'<br>"
      
      %>
      
      <option value=<%=TempArrayUName(i)%>><%=(TempArrayLName(i) & ", " & TempArrayFName(i))%></option>
      
      <%
        rsName.movenext
        i=i+1
      Wend
  
      rsName.close
      set rsName = Nothing    'CLEAR!
      %>
   
      </select></td>
	  	  <td align=center>
	     <table border=0>
	     <tr>
	        <td><input type=button value=" --> " id=button2 name=button2 onClick="moveOption(document.frmInfo.NotSelected,document.frmInfo.Selected,false,false);"></td>
	     </tr>
     	     <tr>	        <td><input type=button value=" <-- " id=button3 name=button3 onClick="moveOption(document.frmInfo.Selected,document.frmInfo.NotSelected,true,false);"></td>	     </tr>     
         </table>	  </td>

	  <td ><select MULTIPLE size=10 name="Selected" id=Selected style="{ width: 100%; }" onfocus="setSelectList(document.frmInfo.NotSelected,false)";>
 	     <%
      
      'Here we are just querying the database for employee names and user names
      'That are already selected for this event
      'These will be displayed in the table under "Selected"
              
      newsql = "select a.lastname, a.firstname, a.username from employee a INNER JOIN Event_Contacts ec ON a.username=ec.Employee_User_Name WHERE ec.Event_ID="&Request.form("eventid")&" ORDER BY a.lastname"
    
      'Notice all of the sizes here are hard coded. Since we don't have 300 people in the branch,
      'this should not cause a problem. However, if you need more than 300, you'll have to change
      'the sizes here.
      Dim SelArrayFName(300)  'store the person's first name
      Dim SelArrayLName(300)  'store the person's last name
      Dim SelArrayUName(300)  'store the person's user name
      set rName = DBQuery(newsql) 'result set
      Dim SelSelected(300)    'Selection properties
    
      'Manipulate the result set
      i=0
      While Not rName.eof
        SelArrayFName(i) = rName("FirstName")
        SelArrayLName(i) = rName("LastName")
        SelArrayUName(i) = rName("Username")
        SelSelected(i) = false
       
        'Now print the names out
        ' Response.Write "'" & TempArrayLName(i) & ", " & TempArrayFName(i) &"'<br>"
      
      %>
      
      <option value=<%=SelArrayUName(i)%>><%=(SelArrayLName(i) & ", " & SelArrayFName(i))%></option>
      
      <%
        rName.movenext
        i=i+1
      Wend
  
      rName.close
      set rName = Nothing    'CLEAR!
      %>
 	     
	  </select></td>
   </tr>

   <TR>
      <TD><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD>
   </TR>  
 
</table>

<br><br>
<input type=button class=button name=backbutt value=" Back " onClick="window.location='javascript:history.back(1);'">
<input type=button class=button value="Cancel" onClick="window.location.href='calendar.asp?date=<%=request.form("rdate")%>'" id=button4 name=button4>
<input type=submit class=button value="Finish" onClick='javasctipt:doSubmit()'; id=submit2 name=submit2>
<input type=hidden name="type" value="<%=request.form("type")%>">
<input type=hidden name="Admin" value="<%=request.form("Admin")%>">
<input type=hidden name="Creator" value="<%=request.form("Creator")%>">

</form>
</center>



<%end if%>

<SCRIPT LANGUAGE="javascript">

/* function that makes sure the information submitted is selected */
function doSubmit()
{
	frm=document.frmInfo;
	setSelectList(frm.Selected,true);	//Turn on all items for server side to process
	//frm.submit();	
}

 
/*
 * Sets the "selected" property for all options of a select list.
 */
function setSelectList(oSelect,value)
{
	var i;
	
	for (i = 0; i < oSelect.length; i++) 
	{       
		oSelect[i].selected=value;
		//alert(oSelect[i].value);
	}
}  
   
/*
 * Function is used to move selected options between the "NotSelected" and "Selected" lists.
 *
 * The boolean argument "deSelectFlag" indicates whether the movement is from "Selected" to
 * "NotSelected". This is required since the option text and values in the "Selected" list are 
 * slightly different than the "NotSelected" list option values.
 * 
 * The boolean argument "resetFlag" indicates when a form clear is occuring
 * 
 */
function moveOption(from,to,deSelectFlag,clearFlag)
{
	var flag=false;
	var i,j,l;
	textList = new Array();
	valueList = new Array();
	
	l=0;
	for (i = 0; i < from.length; i++) 
	{       
		j = to.length;   
		if ( from[i].selected )
		{
    		oText =  from[i].text; 
			oValue = from[i].value;

			// Creates a new option in the destination list
			to[j] = new Option (oText, oValue, false, false);
			//to[j].selected = true;
			flag=true;					//Flag that a movement has occurred
		}
	} 

	if (flag)
	{
		//---------------------------------------------------------------------
		// At least one option was moved and the from list needs these
		// removed.
		//---------------------------------------------------------------------
		for (i = from.length-1; i >= 0; i--) 
		{   
			if (from[i].selected)
			{
				from[i].selected=false;
				from[i] = null;
				l=i;	//first option in list 
			}
		}
		if ( from.length > 0 )
		{
			if (l >= from.length)
				l=from.length-1;
				
			from[0].selected=true;
		}
	}
	else if (clearFlag)
	{
		// No message in this situation
	}
	else if (deSelectFlag)
	{
		alert("Please high-light a name first from the Selected list.");
	}
	else
	{
		alert("Please high-light a name first from the Not Selected list.");
	}
}

</SCRIPT>
<!-- #include file="../../footer.asp" -->
