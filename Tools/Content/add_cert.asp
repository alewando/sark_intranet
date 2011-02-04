<% @language=VBscript%>

<!--
Developer:	  SSeissiger
Date:         08/10/2000
Description:  Allows a certificaiton to be added.
-->

<!-- #include file="../section.asp" -->

<%
if Request.Form("Submitted") = "True" then
	cert_entered = Request.form("txtCert_name")
end if
	
if cert_entered <> "" then
	sql="Select * from Certifications where cert_name ='" & cert_entered & "'"
	set rs = DBQuery(sql)
		if rs.bof then
			sSQL="Insert into Certifications (cert_name) " & _
			     "values ('" & cert_entered & "')"
			set rs = DBQuery(sSQL)
		end if
end if
	
%>

<HTML>
<HEAD>
<title>Add New Certification</title>

</HEAD>
<BODY>

<table border=0>
<tr>
	<td colspan=2><font size=2 color=red>
		<b>ADD CERTIFICATION</b></font>
	</td>
</tr>
<tr>
    <td width=100%><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>
<!--To add a new certification to the database, enter the certification name and --> 
<!--then click 'Submit' when finished:</STRONG></font>-->
	</td>
</tr></table>

<form NAME="frmAddCert" Action="add_cert.asp" method=post>
	<input type=hidden name="Submitted" value="False">
	<table border=0 width=100%>
<tr>
    <TD>
		<P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Certification Name:</STRONG> </FONT></P>
    </TD>
    <TD></TD>
    <TD><INPUT size="50" name="txtCert_name" >
    </TD>
</tr>
<tr></tr>
<tr>
   <TD></TD>
   <TD></TD>
   <TD>	<INPUT type=button class=button onclick='CheckInfo();' value=' Submit ' id=button1 name=button1>&nbsp
		<INPUT type=button class=button value=' Cancel ' OnClick='GoHome();' id=button2 name=button2>
   </td>
</tr>
</form>
</table>
</BODY>

<!-- #include file="../../footer.asp" -->

</HTML>

<SCRIPT LANGUAGE=javascript>
<!--
function CheckInfo()
	{		
		if (document.frmAddCert.Submitted.value = "True")
		{
			cert_entered = document.frmAddCert.textCert_name
		}
		if (document.frmAddCert.txtCert_name.value=="")
		{
			alert("Certification cannot be blank!")
		}
		else
		{
			if (document.frmAddCert.txtCert_name.value!="")
			{
				document.frmAddCert.Submitted.value = "True"
				document.frmAddCert.submit()
			}
		}
	}
	
function GoHome()
	{
		window.navigate("default.asp")
	}
//-->
</SCRIPT>