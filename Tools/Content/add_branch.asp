<%@ Language=VBScript %>
<!--
Developer:	  MAPGAR
Date:         10/09/2000
Description:  Allows a SARK branch to be added.
-->

<!-- #include file="../section.asp" -->

<%
if Request.Form("Submitted") = "True" then
	BranchName=Request.Form("txtBranchName")
	Web=Request.Form("txtWebsite")
	if mid(web,1,4)=lcase("http") then
		if mid(web,5,7)="://" then
			Web=mid(Web,8,len(Web))
		end if
	end if
	Address=Request.Form("txtAddress")
	City=Request.Form("txtCity")
	State=Request.Form("txtState")
	Zip=Request.Form("txtZip")
	Phone=Request.Form("txtPhone")
	if Phone = "(xxx) xxx-xxxx" then
		Phone = ""
	end if
	Fax=Request.Form("txtFax")
	if Fax = "(xxx) xxx-xxxx" then
		Fax = ""
	end if
	TollFree=Request.Form("txtTollFree")
	if TollFree = "(xxx) xxx-xxxx" then
		TollFree = ""
	end if
	
	sSQL="Insert into Branch (Branch_Name, BranchURL, Address, City, State, Zip, Phone, Fax, TollFree) " & _
	     "values ('" & BranchName & "', '" & Web & "', '" & Address & "', '" & City & "', '" & State & "', '" & Zip & "', '" & Phone & "', '" & Fax & "', '" & TollFree & "')"
	dbquery(sSQL)
end if

%>

<SCRIPT LANGUAGE=javascript>
<!--
function CheckInfo()
	{		
		if (document.frmAddBranch.txtBranchName.value=="") {
			alert("Branch Name cannot be blank!")
			}
		else {
			if (document.frmAddBranch.txtBranchName.value!=""){
				document.frmAddBranch.Submitted.value = "True"
				document.frmAddBranch.submit()
			}
		}
	}
	
function GoHome()
	{
		window.navigate("default.asp")
	}

function FormatNumber(val, field)
	{
		document.frmAddBranch.SubmitButton.disabled=false
		if (val.length != 14) {
			alert("Number not formatted Properly.  Use ''(xxx) xxx-xxxx''")
			document.frmAddBranch.SubmitButton.disabled=true
		}
	}
//-->
</SCRIPT>
<HTML>
<HEAD>
<title>Add New Branch</title>
</HEAD>
<BODY>

<table border=0>
<tr>
	<td colspan=2><font size=2 color=red>
		<b>ADD BRANCH</b></font>
	</td>
</tr>

<form NAME="frmAddBranch" Action="add_branch.asp" method=post>
<input type=hidden name="Submitted" value="False">
<table border=0 width=100%>
<tr>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Branch Name:</STRONG> </FONT></P></TD>
      <td></td>
      <td><INPUT size="40" name="txtBranchName" ></td></tr>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Web Site:</STRONG></FONT></P></TD>
	<TD width=7><FONT face="ms sans serif, arial, geneva" size=1 color=black><STRONG>http://</STRONG></FONT></TD>
    <TD><INPUT name="txtWebsite" size="40"></TD></TR>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Address:</STRONG></FONT></P></TD>
	  <td></td>	
    <TD><INPUT name="txtAddress" size="40"></TD></TR>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>City:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtCity" size="40"></TD></TR>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>State:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtState" size="40"></TD></TR>
  <TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Zip:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtZip" size="40"></TD></TR>
<TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Phone:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtPhone" size="40" onChange="JavaScript:FormatNumber(document.frmAddBranch.txtPhone.value, document.frmAddBranch.txtPhone.name);" value="(xxx) xxx-xxxx" maxlength=14></TD></TR>
<TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Fax:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtFax" size="40" onChange="JavaScript:FormatNumber(document.frmAddBranch.txtFax.value, document.frmAddBranch.txtFax.name);" value="(xxx) xxx-xxxx" maxlength=14></TD></TR>
<TR>
    <TD>
      <P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Toll Free:</STRONG></FONT></P></TD>
	<td></td>	
	<TD><INPUT name="txtTollFree" size="40" onChange="JavaScript:FormatNumber(document.frmAddBranch.txtTollFree.value, document.frmAddBranch.txtTollFree.name);" value="(xxx) xxx-xxxx" maxlength=14></TD></TR>
<tr></tr>
<tr></tr>
<tr>
    <td></td><td></td><td><INPUT type=button class=button onclick='CheckInfo();' value=' Submit ' name="SubmitButton">&nbsp
    <input type=button class=button value=' Cancel ' OnClick='GoHome();' name="CancelButton"></td></tr>
</form></TR></table>

</BODY>

<!-- #include file="../../footer.asp" -->

</HTML>
