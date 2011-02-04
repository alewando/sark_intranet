<%@ Language=VBScript %>
<!--
Developer:	  MAPGAR
Date:         03/27/2000
Description:  Allows a SARK client to be added.
-->

<!-- #include file="../section.asp" -->

<%
if Request.Form("Submitted") = "True" then
	clientName=Request.Form("txtClientName")
	Web=Request.Form("txtWebsite")
	
	if mid(web,1,4)=lcase("http") then
		if mid(web,5,7)="://" then
			Web=mid(Web,8,len(Web))
		end if
	end if
	
	if clientName<> "" then
		sql="Select * from Client where ClientName='" & clientName & "'"
		set rs = dbQuery(sql)
		if rs.bof then
			sSQL="Insert into Client (ClientName, Website) values ('" & clientName & "', '" & Web & "')"
			dbquery(sSQL)
		end if
	end if
end if

%>

<SCRIPT LANGUAGE=javascript>
<!--
function CheckInfo()
	{		
		if (document.frmAddClient.txtClientName.value=="") {
			alert("Client Name cannot be blank!")
			}
		else {
			if (document.frmAddClient.txtClientName.value!=""){
				document.frmAddClient.Submitted.value = "True"
				document.frmAddClient.submit()
			}
		}
	}
	
function GoHome()
	{
		window.navigate("default.asp")
	}
//-->
</SCRIPT>


<HTML>
<HEAD>
<title>Add New Client</title>

</HEAD>
<BODY>

<table border=0>
<tr>
	<td colspan=2><font size=2 color=red>
		<b>ADD CLIENT</b></font>
	</td>
</tr>
<tr>
    <td width=100%><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>
To add a new client to the database, enter the clients name and 
Web Site and then click 'Submit' when finished:</STRONG></font></td></tr></table>

<form NAME="frmAddClient" Action="add_client.asp" method=post>
<input type=hidden name="Submitted" value="False">
<table border=0 width=100%>
<tr>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Client Name:</STRONG> </FONT></P></TD>
      <td></td>
      <td><INPUT size="40" name="txtClientName" ></td></tr>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Web Site:</STRONG></FONT></P></TD>
	<TD width=7><FONT face="ms sans serif, arial, geneva" size=1 color=black><STRONG>http://</STRONG></FONT></TD>
    <TD><INPUT name="txtWebsite" size="40"></TD></TR>
<tr></tr>
<tr></tr>
<tr>
    <td></td><td></td><td><INPUT type=button class=button onclick='CheckInfo();' value=' Submit '>&nbsp
    <input type=button class=button value=' Cancel ' OnClick='GoHome();'></td></tr>
</form></TR></table>

</BODY>

<!-- #include file="../../footer.asp" -->

</HTML>
