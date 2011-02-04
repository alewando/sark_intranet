<%@ Language=VBScript %>
<!--
Developer:	  MAPGAR
Date:         10/11/2000
Description:  Allows a SARK branch to be edited.
-->

<!-- #include file="../section.asp" -->

<%
sSQL = "Select * from Branch ORDER BY Branch_Name"
set Rs = dbQuery(sSQL)
if Request.Form("Submitted") = "True" and Request.form("branch") <> "" then
	BranchName=Request.Form("txtBranchName")
	Web=Request.Form("txtWebsite")
	if mid(web,1,4)=lcase("http") then
		if mid(web,5,7)="://" then
			Web=mid(Web,8,len(Web))
		end if
	end if
	IAddress=Request.Form("txtAddress")
	ICity=Request.Form("txtCity")
	IState=Request.Form("txtState")
	IZip=Request.Form("txtZip")
	IPhone=Request.Form("txtPhone")
	if IPhone = "(xxx) xxx-xxxx" then
		IPhone = ""
	end if
	IFax=Request.Form("txtFax")
	if IFax = "(xxx) xxx-xxxx" then
		IFax = ""
	end if
	ITollFree=Request.Form("txtTollFree")
	if ITollFree = "(xxx) xxx-xxxx" then
		ITollFree = ""
	end if

	sSQL="Update Branch set Branch_Name='" & BranchName & "', BranchURL='" & Web & "', Address='" & IAddress & "', City='" & ICity & "', State='" & IState & "', Zip='" & IZip & "', Phone='" & IPhone & "', Fax='" & IFax & "', TollFree='" & ITollFree & "' where Branch_ID='" & Request.Form("BID") & "'" 
	dbquery(sSQL)
elseif Request.Form("Submitted") = "False" and Request.Form("branch") <> "" then
	sSQL = "Select * from Branch where Branch_ID = '" & Request.Form("branch") & "'"
	set Rs=dbquery(sSQL)
end if
%>

<SCRIPT LANGUAGE=javascript>
<!--
function CheckInfo()
	{		
		if (document.frmEditBranch.txtBranchName.value=="") {
			alert("Branch Name cannot be blank!")
			}
		else {
			if (document.frmEditBranch.txtBranchName.value!=""){
				document.frmEditBranch.Submitted.value = "True"
				document.frmEditBranch.branch.value=document.frmEditBranch.txtBranchName.value

				document.frmEditBranch.submit()
			}
		}
	}
	
function GoHome()
	{
		window.navigate("edit_branch.asp")
	}

function DBUpdated()
	{
		document.frmEditBranch.branch.value=""
		document.frmEditBranch.Submitted.value="False"
		document.frmEditBranch.submit()
	}
	
function FormatNumber(val, field)
	{
		document.frmEditBranch.SubmitButton.disabled=false
		if (val.length != 14) {
			alert("Number not formatted Properly.  Use ''(xxx) xxx-xxxx''")
			document.frmEditBranch.SubmitButton.disabled=true
		}
	}

function Reload()
	{
		document.frmEditBranch.branch.value=document.frmEditBranch.txtBranchName.value
		document.frmEditBranch.submit()
	}
//-->
</SCRIPT>
<HTML>
<HEAD>
<title>Edit Branch</title>
</HEAD>
<BODY>

<%if Request.Form("branch") = "" then%>
<table border=0>
<tr>
	<td colspan=2><font size=2 color=red>
		<b>EDIT BRANCH</b></font>
	</td>
</tr>

<form NAME="frmEditBranch" Action="edit_branch.asp" method=post>
<input type=hidden name="Submitted" value="False">
<input type=hidden name="branch" value="">
<table border=0 width=100%>

<tr>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Branch Name:</STRONG> </FONT></P></TD>
      <td></td>
      <td><select name="txtBranchName" onChange="JavaScript:Reload();">
		  <%Response.Write("<option value=" & chr(34) & chr(34) & chr(34) & chr(34) & chr(34) & chr(34) & "></option>" & NL)
		while not Rs.eof
			response.write("<option value=" & chr(34) & Rs("Branch_ID") & chr(34))
			response.write(">" & Rs("Branch_Name") & "</option>" & NL)
			Rs.movenext
		wend
		Rs.movefirst%>
		</select></td></tr>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Web Site:</STRONG></FONT></P></TD>
	<TD width=7><FONT face="ms sans serif, arial, geneva" size=1 color=black><STRONG>http://</STRONG></FONT></TD>
    <TD><INPUT name="txtWebsite" size="40" value=""></TD></TR>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Address:</STRONG></FONT></P></TD>
	  <td></td>	
    <TD><INPUT name="txtAddress" size="40" value=""></TD></TR>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>City:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtCity" size="40" value=""></TD></TR>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>State:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtState" size="40" value=""></TD></TR>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Zip:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtZip" size="40" value=""></TD></TR>
<TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Phone:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtPhone" size="40" onChange="JavaScript:FormatNumber(document.frmEditBranch.txtPhone.value, document.frmEditBranch.txtPhone.name);" maxlength=14 value=""></TD></TR>
<TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Fax:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtFax" size="40" onChange="JavaScript:FormatNumber(document.frmEditBranch.txtFax.value, document.frmEditBranch.txtFax.name);" maxlength=14 value=""></TD></TR>
<TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Toll Free:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtTollFree" size="40" onChange="JavaScript:FormatNumber(document.frmEditBranch.txtTollFree.value, document.frmEditBranch.txtTollFree.name);" maxlength=14 value=""></TD></TR>
<tr></tr>
<tr></tr>
<tr>
    <td></td><td></td><td><INPUT type=button class=button onclick='CheckInfo();' value='  Edit  ' name="SubmitButton">&nbsp
    <input type=button class=button value=' Cancel ' OnClick='GoHome();' name="CancelButton"></td></tr>
<%elseif Request.Form("Submitted") = "False" and Request.Form("branch") <> "" then%>
<table border=0>
<tr>
	<td colspan=2><font size=2 color=red>
		<b>EDIT BRANCH</b></font>
	</td>
</tr>

<form NAME="frmEditBranch" Action="edit_branch.asp" method=post>
<input type=hidden name="Submitted" value="False">
<input type=hidden name="branch" value="">
<table border=0 width=100%>

<tr>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Branch Name:</STRONG> </FONT></P></TD>
      <td></td>
      	<td><input name="txtBranchName" size="40" value="<%=Rs("Branch_Name")%>"></td></tr>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Web Site:</STRONG></FONT></P></TD>
	<TD width=7><FONT face="ms sans serif, arial, geneva" size=1 color=black><STRONG>http://</STRONG></FONT></TD>
    <TD><INPUT name="txtWebsite" size="40" value="<%=Rs("BranchURL")%>"></TD></TR>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Address:</STRONG></FONT></P></TD>
	  <td></td>	
    <TD><INPUT name="txtAddress" size="40" value="<%=Rs("Address")%>"></TD></TR>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>City:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtCity" size="40" value="<%=Rs("City")%>"></TD></TR>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>State:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtState" size="40" value="<%=Rs("State")%>"></TD></TR>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Zip:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtZip" size="40" value="<%=Rs("Zip")%>"></TD></TR>
<TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Phone:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtPhone" size="40" onChange="JavaScript:FormatNumber(document.frmEditBranch.txtPhone.value, document.frmEditBranch.txtPhone.name);" maxlength=14 value="<%=Rs("Phone")%>"></TD></TR>
<TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Fax:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtFax" size="40" onChange="JavaScript:FormatNumber(document.frmEditBranch.txtFax.value, document.frmEditBranch.txtFax.name);" maxlength=14 value="<%=Rs("Fax")%>"></TD></TR>
<TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Toll Free:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtTollFree" size="40" onChange="JavaScript:FormatNumber(document.frmEditBranch.txtTollFree.value, document.frmEditBranch.txtTollFree.name);" maxlength=14 value="<%=Rs("TollFree")%>"></TD></TR>
<tr><td><input type=hidden name="BID" size="40" value="<%=Rs("Branch_ID")%>"></td>
</tr>
<tr></tr>
<tr>
    <td></td><td></td><td><INPUT type=button class=button onclick='CheckInfo();' value='  Edit  ' name="SubmitButton">&nbsp
    <input type=button class=button value=' Cancel ' OnClick='GoHome();' name="CancelButton"></td></tr>
<%elseif Request.Form("Submitted") = "True" then%>
<table border=0>
<tr></tr>
<form NAME="frmEditBranch" Action="edit_branch.asp" method=post>
<input type=hidden name="Submitted" value="False">
<input type=hidden name="branch" value="">
<table border=0 width=100%>

	<tr></tr>
	<tr>
	<td></td><td></td><td><font color="blue"><strong>Database Updated!</strong></font></td></tr>
	<tr>
    <td></td><td></td><td><input type=button class=button value='  OK  ' OnClick='DBUpdated();'></td></tr>
<%end if%>
</form></TR></table>

</BODY>

<!-- #include file="../../footer.asp" -->

</HTML>
