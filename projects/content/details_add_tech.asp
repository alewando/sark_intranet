<% response.expires = 0 %>
<!-- #include file="../../script.asp" -->


<%
'---------------------------------------------------------------------------------
' Description:	Adds technologies to a project.
' History:		06/10/1999 - KDILL - Created
'---------------------------------------------------------------------------------
%>


<html>

<head>
<title>Add/Remove "<% =request("ProjectName") %>" technologies</title>
<!-- #include file="../../style.htm" -->
</head>

<body bgcolor=silver><center>

<%
ProjectID = CInt(request("ProjectID"))
if request("action") = "update" then
%>

<font size=3 face="ms sans serif, arial, geneva">
Saving...
</font>

<%
	'-----------------------------------------------
	'	Update project technology listing			
	'-----------------------------------------------
	DBQuery("DELETE FROM Project_Tech_Xref WHERE Project_ID = " & ProjectID)
	for each x in request("Tech")
		DBQuery("INSERT INTO Project_Tech_Xref (Project_ID, Tech_ID) values (" & ProjectID & ", " & x & ")")
		next
	DBQuery("UPDATE Project SET OtherTech = '" & request("OtherTech") & "' WHERE Project_ID = " & ProjectID)
%>

<script language=javascript>
<!--
var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}
function closeWin(){
	if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.reload();
	window.close()
	}
i = setTimeout("closeWin()", 700)
// -->
</script>

<% else %>


<table border=0 width='90%'><tr><td><font size=1 face="ms sans serif, arial, geneva">
<b>Please select the technologies which were employed on the project.</b>&nbsp;&nbsp;
If the technology is not listed, please enter it in the "Other technologies" area.
</font></td></tr></table>

<form NAME="frmInfo" ACTION="details_add_tech.asp">
<input type=hidden name="ProjectID" value="<%=request("ProjectID") %>">
<input type=hidden name="ProjectName" value="<% =request("ProjectName") %>">
<input type=hidden name="ClientID" value="<% =request("ClientID") %>">

<TABLE border=0><tr><td><font size=1 face='ms sans serif, arial, geneva'>
<%
'-------------------------------------------------
'	Execute Database Query for Tech Listing Info  
'-------------------------------------------------
sql = "SELECT Tech.Tech_ID, Tech.Tech_Name, Tech_Area.Tech_Area, Project_Tech_Xref.Project_ID FROM Tech INNER JOIN Tech_Area ON Tech.Tech_Area_ID = Tech_Area.Tech_Area_ID LEFT OUTER JOIN Project_Tech_Xref ON Tech.Tech_ID = Project_Tech_Xref.Tech_ID Where Tech.Approved = 1 ORDER BY Tech_Area.Tech_Area, Tech.Tech_Name"
set rs = DBQuery(sql)
i = 0
while not rs.eof
	if (i <> 0) and (i <> rs("Tech_ID")) then
		Response.Write("<INPUT TYPE=checkbox NAME=tech VALUE=" & tech_id & " " & checked & ">" & tech_name & "<BR>" & NL)
		checked = ""
		end if
	if tech_area <> rs("Tech_Area") then response.write("<p><b>" & rs("Tech_Area") & "</b><br>" & NL)
	tech_id = rs("Tech_ID")
	tech_name = rs("Tech_Name")
	tech_area = rs("Tech_Area")
	i = tech_id
	if not isnull(rs("Project_ID")) then
		if CInt(rs("Project_ID")) = ProjectID then
			checked = "checked"
			end if
		end if
	rs.movenext
	wend
rs.close
Response.Write("<INPUT TYPE=checkbox NAME=tech VALUE=" & tech_id & " " & checked & ">" & tech_name & "<BR>" & NL)
%>
<p>
<b>Other technologies</b>:<br>
<%
OtherTech = ""
set rs = DBQuery("select OtherTech from Project where Project_ID = " & ProjectID)
if not isnull(rs("OtherTech")) then OtherTech = rs("OtherTech")
rs.close
%>
<input type=text name="OtherTech" value="<% =OtherTech %>" size=30 maxlength=255>

</font></td></tr></TABLE>

<input type=hidden name="action" value="update"><br>
<table align=center width=75% border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Save Technologies" OnClick='document.frmInfo.submit();'>
			<input type=button class=button value="Cancel" OnClick='window.close();'>
		</td>
	</tr>
</table>
</form>


<% end if %>

</center>

</body>
</html>

<%
on error resume next
DataConn.close
set DataConn = nothing
%>
