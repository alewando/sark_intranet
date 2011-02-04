<!--
Developer:    mapalmer & nyates
Date:         12/08/1999
Description:  Allows editing of employee data
-->


<!-- #include file="../../script.asp" -->

<html>

<head>
<title>Update Employee</title>
<!-- #include file="../../style.htm" -->
</head>


<%
'-----------------------------------------------
'	Execute Database Query for Profile Info		
'-----------------------------------------------
empID = Request.QueryString("EmpID") 
set rs = DBQuery("select * from employee where employeeid = " & empID)
if not rs.eof then
	ls_insertupdate = "E"
	'-----------------------------------------------
	'	Get Profile Info							
	'-----------------------------------------------
	ls_fname = trim(rs("FirstName"))
	ls_lname = trim(rs("LastName"))
	ls_address = trim(rs("Address"))
	ls_city = trim(rs("City"))
	ls_state = trim(rs("State"))
	ls_zip = trim(rs("Zip"))
	ls_hphone = trim(rs("HomePhone"))
	ls_spouse = trim(rs("SpouseName"))
	ls_pager = trim(rs("PagerNumber"))
	ls_pager_email = trim(rs("PagerEmail"))
	ls_birthday = trim(rs("Birthday"))
	ls_personal_website = trim(rs("PersonalWebsite"))
	ls_clientPhone = trim(rs("Client_Phone"))
	ls_cell = trim(rs("CellPhone"))
	ls_vmail = trim(rs("VoiceMail"))
	ls_ClientEmail = trim(rs("ClientEmail"))
	ls_client_ID = trim(rs("Client_ID"))
end if
rs.close
%>

<body bgcolor=silver><center>

<table border=0 width='90%'><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
Please enter your profile information.
This is used to help introduce you to the rest of our company.
Please avoid the use of single or double quotes.
Thank you. <%=lsclient_id%>
</b></font></td></tr></table>


<form NAME="frmInfo" ACTION="updateemp.asp">
<INPUT TYPE="Hidden" NAME="EmployeeID" VALUE="<%=empID%>"> 

<table border=0>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">First Name:</font></td>
	<td><INPUT TYPE="Text" NAME="FirstName" VALUE="<%=ls_fname%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Last Name:</font></td>
	<td><INPUT TYPE="Text" NAME="LastName" VALUE="<%=ls_lname%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Address:</font></td>
	<td><INPUT TYPE="Text" NAME="Address" VALUE="<%=ls_address%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">City:</font></td>
	<td><INPUT TYPE="Text" NAME="City" VALUE="<%=ls_city%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">State:</font></td>
	<td><INPUT TYPE="Text" NAME="State" VALUE="<%=ls_state%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Zip:</font></td>
	<td><INPUT TYPE="Text" NAME="Zip" VALUE="<%=ls_zip%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Home Phone:</font></td>
	<td><INPUT TYPE="Text" NAME="HomePhone" VALUE="<%=ls_hphone%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Pager Number:</font></td>
	<td><INPUT TYPE="Text" NAME="PagerNumber" VALUE="<%=ls_pager%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Pager Email:</font></td>
	<td><INPUT TYPE="Text" NAME="PagerEmail" VALUE="<%=ls_pager_email%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Cell Phone:</font></td>
	<td><INPUT TYPE="Text" NAME="CellPhone" VALUE="<%=ls_Cell%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Spouse Name:</font></td>
	<td><INPUT TYPE="Text" NAME="SpouseName" VALUE="<%=ls_spouse%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Birthday:</font></td>
	<td><INPUT TYPE="Text" NAME="Birthday" VALUE="<%=ls_birthday%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Personal Website: http://</font></td>
	<td><INPUT TYPE="Text" NAME="PersonalWebsite" VALUE="<%=ls_personal_website%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Voice Mail:</font></td>
	<td><INPUT TYPE="Text" NAME="VoiceMail" VALUE="<%=ls_vmail%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Client Name:</font></td>
	<td>
		<Select NAME="Client_ID">
			<%set rs = DBQuery("select Client_ID, ClientName from client order by ClientName")
				if not rs.eof and not rs.bof then
					do while not rs.eof%>
						<Option Value=<%=rs("Client_ID")%> <%if trim(rs("Client_ID")) = ls_client_id then%> Selected<%end if%>><%=trim(rs("ClientName"))%>
						<% rs.movenext
					loop
					rs.close
				end if%>			
		</Select>
	</td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Client Phone:</font></td>
	<td><INPUT TYPE="Text" NAME="Client_Phone" VALUE="<%=ls_clientphone%>" size=40 maxlength=255></td>
</tr>

<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Client Email:</font></td>
	<td><INPUT TYPE="Text" NAME="ClientEmail" VALUE="<%=ls_clientemail%>" size=40 maxlength=255></td>
</tr>

<%
'-------------------------------
'get client information
'-------------------------------
'set rs = DBQuery("select Client_ID, ClientName from client order by ClientName)

'rs.close 
%>


</table>

<INPUT TYPE="Hidden" NAME="InsertUpdate" VALUE="<%=ls_insertupdate%>"><br>
<table align=center width=75% border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Update" OnClick='document.frmInfo.submit();' id=button1 name=button1>
			<input type=button class=button value="Cancel" OnClick='window.close();' id=button2 name=button2>
		</td>
	</tr>
</table>

</form>
</center>

</body>
</html>
