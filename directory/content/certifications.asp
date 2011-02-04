<!--
Developer:    GINGRICE
Date:         09/10/1998
Description:  Allows specifying Technical Subject Experts
-->


<!-- #include file="../../script.asp" -->

<html>

<head>
<title>Certifications</title>
<!-- #include file="../../style.htm" -->
</head>


<%
'-------------------------------------------------
'	Execute Database Query for Tech Experts Info  
'-------------------------------------------------
NL = chr(13) & chr(10)
empID = Request.QueryString("EmpID")
' Remember to only select records where tech.Approved = 1!
ls_sql = "SELECT Certifications.Cert_ID, Certifications.Cert_Name," _
    & " Employee_Cert_Xref.Employee_ID," _
    & " Employee_Cert_Xref.Employee_Cert_Xref" _
	& " FROM Employee_Cert_Xref RIGHT OUTER JOIN" _
    & " Certifications ON" _
	& " Employee_Cert_Xref.Cert_ID = Certifications.Cert_ID" _
	& " WHERE (Employee_Cert_Xref.Employee_ID = " & empID & ")" _
	& " UNION" _
	& " SELECT Certifications.Cert_ID, Certifications.Cert_Name, null, null" _
	& " FROM Certifications" _
	& " WHERE Certifications.Cert_ID NOT IN" _
	& "		(SELECT Certifications.Cert_ID" _
	& "		FROM Employee_Cert_Xref RIGHT OUTER JOIN" _
    & "		Certifications ON" _
    & "		Employee_Cert_Xref.Cert_ID = Certifications.Cert_ID" _
	& "		WHERE (Employee_Cert_Xref.Employee_ID = " & empID & "))" _
	& " ORDER BY Certifications.Cert_Name"

Response.Write("<!--" & vbNewLine)
Response.Write("EmpID = " & empID & "<BR>" & vbNewLine)
Response.write("ls_sql = '" & ls_sql & "'" & vbNewLine)
Response.Write("-->")

set rs = DBQuery(ls_sql)
%>

<body bgcolor=silver link=blue alink=purple vlink=blue onLoad="window.resizeBy(0, -190)">
<center>

<table border=0 width='90%'><tr><td align=center><font size=1 face="ms sans serif, arial, geneva">
<b>Please specify the appropriate certifications.</b>
</FONT></td></tr></table>

<form NAME="frmInfo" ACTION="updateemp.asp">
<INPUT TYPE="Hidden" NAME="EmployeeID" VALUE="<%=empID%>">

<TABLE border=0 width='40%'><tr><td><font size=1 face='ms sans serif, arial, geneva'>
<%
while not rs.eof
	ls_empid = rs("Employee_ID")
	ls_CertID = rs("Cert_ID")
	ls_CertName = rs("Cert_Name")
	'If ls_empid is null, then the user has not specified themselves
	'as an expert
	Response.Write("<INPUT TYPE=checkbox NAME=certs VALUE=" & rs("Cert_ID"))
	if not isnull(rs("Employee_ID")) then 
	    Response.Write(" CHECKED")
	end if
	Response.Write(">" & ls_CertName & "<BR>" & NL)
	rs.movenext
wend

rs.close
set rs = nothing

'------------------------------------------------
' Fields added for specifying MCP # 
' (supporting code allows for other cert #'s, but only MCP is used 
' at this time
'------------------------------------------------
 ls_sql = "SELECT category_id, cert_number FROM employee_cert_numbers WHERE employee_id=" & empID
 set rs = DBQuery(ls_sql)
%> 
<HR>
<INPUT TYPE=checkbox NAME=CertNum VALUE=1
<% 
 If not rs.eof and Not IsNull(rs("cert_number")) Then
  Response.Write(" CHECKED ")
  ls_certNum = rs("cert_number")
 else
  ls_certNum = ""
 End If
%>
>
MCP # <INPUT TYPE=text NAME=CertNum1Value VALUE="<%=ls_certNum%>" MAXLENGTH=15 SIZE=10 >
<BR>

</font></td></tr></TABLE>

<INPUT TYPE="Hidden" NAME="InsertUpdate" VALUE="CERT"><br>
<table align=center width=75% border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Update Certifications" OnClick='document.frmInfo.submit();'>
			<input type=button class=button value="Cancel" OnClick='window.close();'>
		</td>
	</tr>
</table>
</form>
</center>

</body>
</html>

