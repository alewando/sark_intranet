<% response.expires = 0 %>
<!-- #include file="../../script.asp" -->


<%
'---------------------------------------------------------------------------------
' Description:	Add a new project.
' History:		06/10/1999 - KDILL - Created
'---------------------------------------------------------------------------------
%>


<html>
<head>
<% if request("projectid") = "" then%>
<title>Add new project</title>
<% else %>
<title>Edit project</title>
<% end if %>
<!-- #include file="../../style.htm" -->
</head>


<body bgcolor=silver><font face="ms sans serif, arial, geneva">

<center><font size=3>

<%
if request("action") <> "" then
	if request("projectid") <> "" then
		
		if request("action") = "add" then
			
			' UPDATE EXISTING PROJECT...
			sql = "UPDATE Project SET "
			sql = sql & "Client_ID = " & request("ClientID")
			sql = sql & ", Project_Name = '" & Clean(request("Project_Name")) & "'"
			sql = sql & ", Project_Desc = '" & Clean(request("Project_Desc")) & "'"
			sql = sql & ", Client_Contacts = '" & Clean(request("Client_Contacts")) & "'"
			if request("StartDate") = "" then
				sql = sql & ", StartDate = NULL"
			else
				sql = sql & ", StartDate = '" & request("StartDate") & "'"
			end if
			if request("EndDate") = "" then
				sql = sql & ", EndDate = NULL"
			else
				sql = sql & ", EndDate = '" & request("EndDate") & "'"
			end if
			sql = sql & " WHERE Project_ID = " & request("ProjectID")
			DBQuery(sql)
			response.write("Project updated...")
			
		else if request("action") = "delete" then
			
			sql = "DELETE FROM Project where Project_ID = " & request("projectid")
			response.write("<!-- sql = '" & sql & "' -->")
			DBQuery(sql)
			response.write("Project deleted...")
			end if
			
		end if
		
	else
		
		' INSERT NEW PROJECT...
		sqlfields = "Client_ID, Project_Name, Project_Desc, Client_Contacts, OtherTech"
		sqlvalues = request("ClientID") & ", '" & Clean(request("Project_Name")) & "','" & Clean(request("Project_Desc")) & "','" & Clean(request("Client_Contacts")) & "',''"
		if request("StartDate") <> "" then
			sqlfields = sqlfields & ", StartDate"
			sqlvalues = sqlvalues & ", '" & request("StartDate") & "'"
			end if
		if request("EndDate") <> "" then
			sqlfields = sqlfields & ", EndDate"
			sqlvalues = sqlvalues & ", '" & request("EndDate") & "'"
			end if
		sql = "INSERT INTO Project (" & sqlfields & ") VALUES (" & sqlvalues & ")"
		DBQuery(sql)
		set rs = DBQuery("select @@IDENTITY from Project")
		ProjectID = rs(0)
		rs.close
		response.write("Project added...")
		
		end if
%>


</font>

<script language=javascript>
<!--
var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}
function closeWin(){
<% if request("projectid") = "" then %>
	if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.href="details.asp?clientid=<% =request("ClientID") %>&projectid=<% =ProjectID %>"
<% else %>
	if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.reload()
<% end if %>
	window.close()
	}
i = setTimeout("closeWin()", 700)
// -->
</script>

<%
else
%>

<script language=javascript>
<!--
function valid(fld, title){
	var result = true
	if (fld.value == ""){
		fld.select()
		alert("Please fill in the " + title + ".");
		fld.focus()
		result = false
		}
	return result
	}

function validDate(datefld, fldname){
	var result = true
	var testval = datefld.value
	var finalval = ""
	if (testval != ""){
		for (i=0; i < testval.length; i++){
			ch = testval.substring(i,i+1)
			if ("0123456789/".indexOf(ch) > -1) finalval += ch
			}
		if (Date.parse(finalval) == "NaN"){
			alert("Bad value")
			result = false
			datefld.select()
			alert("Please enter an appropriate " + fldname + " (mm/dd/yyyy)")
			datefld.focus()
			}
		else {datefld.value = finalval}
		}
	return result
	}

function validateFields(){
	var result = false
	if (document.frmInfo.action.value == "delete") {return true}
	else{
		if (valid(document.frmInfo.Project_Name, "project name")){
		if (valid(document.frmInfo.Project_Desc, "description")){
			result = true
			if (result) result = validDate(document.frmInfo.StartDate, "start date")
			if (result) result = validDate(document.frmInfo.EndDate, "end date")
			}}
		}
	return result
	}
// -->
</script>


<form name="frmInfo" method="post" action="details_add.asp" onSubmit="return (validateFields())">
<input type=hidden name=projectid value="<%=request("ProjectID")%>">

<table border=0>

	<tr>
		<td valign=middle><font size=1 face="ms sans serif, arial, geneva">
			Client:
			</font></td>
		<td width=10>&nbsp;</td>
		<td nowrap><font size=1 face="ms sans serif, arial, geneva">
			<select name="ClientID">
<%
ClientID = CInt(request("ClientID"))

' Get list of clients...
sql = "select Client_ID, ClientName from Client order by ClientName"
set rs = DBQUery(sql)
while not rs.eof
	response.write("<option value='" & rs("Client_ID") & "'")
	if ClientID = rs("Client_ID") then response.write(" selected")
	response.write(">" & rs("ClientName") & "</option>" & NL)
	rs.movenext
	wend
rs.close

' If existing project...
if request("ProjectID") <> "" then
	sql = "select * from Project where Project_ID = " & request("ProjectID")
	set rs = DBQuery(sql)
	if not rs.eof then
		Project_Name = rs("Project_Name")
		Client_Contacts = rs("Client_Contacts")
		Project_Desc = rs("Project_Desc")
		if not isnull(rs("StartDate")) then StartDate = rs("StartDate")
		if not isnull(rs("EndDate")) then EndDate = rs("EndDate")
		end if
	end if
%>
			</select>&nbsp;
			</td>
	</tr>
	
	<tr>
		<td valign=top><font size=1 face="ms sans serif, arial, geneva">
			Project / Contacts:
			</font></td>
		<td>&nbsp;</td>
		<td>
			<input type=text name="Project_Name" value="<% =Project_Name %>" size=25 maxlength=100>
			<input type=text name="Client_Contacts" value="<% =Client_Contacts %>" size=25 maxlength=100>
			</td>
	</tr>
	
	<tr>
		<td valign=top><font size=1 face="ms sans serif, arial, geneva">
			Description:
			</font></td>
		<td>&nbsp;</td>
		<td><textarea name="Project_Desc" cols=55 rows=11><% =Project_Desc %></textarea></td>
	</tr>
	
	<tr>
		<td valign=middle><font size=1 face="ms sans serif, arial, geneva">
			Start / End Dates:
			</font></td>
		<td>&nbsp;</td>
		<td>
			<input type=text name=StartDate value="<% =StartDate %>" size=10 maxlength=10>&nbsp;
			<input type=text name=EndDate value="<% =EndDate %>" size=10 maxlength=10>
			</td>
	</tr>
	
</table>

<br>
<input type=hidden name=action value="add">
<% if request("ProjectID") <> "" then %>
<input type=button class=button value="Delete Project" onClick="if (confirm('Are you sure you wish to delete this project?')){document.frmInfo.action.value='delete'; document.frmInfo.submit();}">&nbsp;&nbsp;&nbsp;&nbsp;
<% end if %>
<input type=submit class=button value=" Save Project "><input type=button class=button value=" Cancel " onClick="window.close()">

</form>

<% end if %>

</font></center>


</body>
</html>


<%
on error resume next
DataConn.close
set DataConn = nothing
%>