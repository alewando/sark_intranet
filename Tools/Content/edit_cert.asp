<%@ Language = Vbscript%>
<!--
Developer:	  Sseissiger
Date:         03/27/2000
Description:  Allows a SARK Cert to be added.
-->

<!-- #include file="../section.asp" -->

<%
		if Request.Form("Submitted")="True" then
		ctName=Request.Form("Cert_Name")
		ls_CertID=Request.Form("Cert_ID")
		
		if ctName<> "" then
			queryString = "UPDATE Certifications SET Cert_Name = '" & ctName & "' WHERE Cert_ID = " & ls_CertID
			DBQuery(queryString)
		end if

	elseif Request.form("Selected")="True" then
		set rs=dbquery("Select Cert_ID, Cert_Name from Certifications " & _
		               "where Cert_ID='" & Request.Form("Cert_ID") & "'")
		ls_CName=trim(rs("Cert_Name"))
		ls_CertID=rs("Cert_ID")
		rs.close
	end if
%>
<SCRIPT LANGUAGE=javascript>
<!--
	function Changes()
	{
				document.frmInfo.Submitted.value="True"
				document.frmInfo.submit()
	}
	
	function GoHome()
	{
		window.navigate("default.asp")
	}
	function LoadForm()
	{
		document.frmInfo.Selected.value="True"
		document.frmInfo.submit()
	}
//-->
</SCRIPT>

<HTML>
<HEAD>
</HEAD>
<body>

<table border=0><tr><td><font size=1 face="ms sans serif, arial, geneva" color=blue><b><STRONG>
  <font size=2 color=red><b>EDIT Certification</b></font><p>
  <!-- This form is used to change a Cert's name or website URL. -->
  </font></td></tr></table>

<form NAME="frmInfo" ACTION="edit_Cert.asp" method=post>
  <INPUT TYPE=Hidden NAME="InsertUpdate" VALUE="<%=ls_CertID%>"><br>
  <INPUT TYPE=Hidden NAME="Submitted" VALUE="False">
  <INPUT TYPE=Hidden NAME="Selected" VALUE="False">

  <table border=0>
  <tr>
	<td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Certification Name:</STRONG></font></td>
	<td></td>
	<td>
		<Select NAME="Cert_ID" onChange='javascript:LoadForm();'>
			<%	set rs = DBQuery("select * from Certifications order by Cert_Name")
				if not rs.eof and not rs.bof then
					Response.Write("<option value=" & chr(34) & chr(34) & "></option>")
					do while not rs.eof%>
							<Option Value=<%=trim(rs("Cert_ID"))%>><%=trim(rs("Cert_Name"))%>
						<% rs.movenext
					loop
					rs.close
					if ls_CName<>"" and Request.form("Submitted")="False" then%>
						<Option selected value=<%=ls_CertID%>><%=ls_CName%>
					<%end if
				
					end if%>
		</Select>
	</td>
  </tr>

  <tr><td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Certification Name:</STRONG></font></td>
	<td></td>
	<%if ls_CName="" and Request.form("Submitted")="False" then
		Response.Write "<td><INPUT TYPE='Text' NAME='Cert_Name' VALUE='' size=40 maxlength=255></td></tr>"
	else
		Response.Write ("<td><INPUT TYPE='Text' NAME='Cert_Name' VALUE=" & chr(34) & ls_CName & chr(34) & " size=40 maxlength=255></td></tr>")
	end if%>
	<!--<td><INPUT TYPE="Text" NAME="Cert_Name" VALUE="<%=ls_CName%>" size=40 maxlength=255></td></tr>-->
	
</table>

<table align=center width=75% border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Update" OnClick='Changes();' id=button1 name=button1>
			<input type=button class=button value="Cancel" OnClick='GoHome();' id=button2 name=button2>
		</td>
	</tr>
	<tr>
		<td align=center>
			<%	if mid(chg,1,6)="Update" then
					Response.Write ("<b><Font face='ms sans serif, arial, geneva' size=1 color=red><Strong>Database Updated!</Strong></font></b>")
				end if%>
		</td>
	</tr>
</table>

</form>
</body>
</HTML>
<!-- #include file="../../footer.asp" -->