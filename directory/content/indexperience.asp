<!--
Developer:    GINGRICE
Date:         09/10/1998
Description:  Allows specifying Technical Subject Experts
-->


<!-- #include file="../../script.asp" -->

<html>

<head>
<title>Industry Experience</title>
<!-- #include file="../../style.htm" -->
</head>


<%
'-------------------------------------------------
'	Execute Database Query for Tech Experts Info  
'-------------------------------------------------
NL = chr(13) & chr(10)
empID = Request.QueryString("EmpID")
' Remember to only select records where tech.Approved = 1!
ls_sql = "SELECT Industry.Industry_ID, Industry.Industry_Name," _
    & " Industry.Approved, Employee_Industry_Xref.Employee_ID," _
    & " Employee_Industry_Xref.Employee_Ind_Xref" _
	& " FROM Employee_Industry_Xref RIGHT OUTER JOIN" _
    & " Industry ON" _
	& " Employee_Industry_Xref.Industry_ID = Industry.Industry_ID" _
	& " WHERE (Industry.Approved = 1) AND" _
    & " (Employee_Industry_Xref.Employee_ID = " & empID & ")" _
	& " UNION" _
	& " SELECT Industry.Industry_ID, Industry.Industry_Name, Industry.Approved, null, null" _
	& " FROM Industry" _
	& " WHERE (Industry.Approved = 1) AND Industry.Industry_ID NOT IN" _
	& "		(SELECT Industry.Industry_ID" _
	& "		FROM Employee_Industry_Xref RIGHT OUTER JOIN" _
    & "		Industry ON" _
    & "		Employee_Industry_Xref.Industry_ID = Industry.Industry_ID" _
	& "		WHERE (Industry.Approved = 1) AND" _
    & "		(Employee_Industry_Xref.Employee_ID = " & empID & "))" _
	& " ORDER BY Industry.Industry_Name"

Response.Write("<!--" & vbNewLine)
Response.Write("EmpID = " & empID & "<BR>" & vbNewLine)
Response.write("ls_sql = '" & ls_sql & "'" & vbNewLine)
Response.Write("-->")

set rs = DBQuery(ls_sql)
%>
<body bgcolor=silver link=blue alink=purple vlink=blue onLoad="window.resizeBy(0, -150)"><center>

<table border=0 width='90%'><tr><td><font size=1 face="ms sans serif, arial, geneva">
<b>Please specify the industries in which you consider yourself to have significant experience in.</b><p>
<font color=maroon><b>Note:</b> By selecting an industry, you are designating yourself as a reference to other SARKs, and may receive phone calls or emails relating to the specified industries.</font>
</FONT></td></tr></table>

<form NAME="frmInfo" ACTION="updateemp.asp">
<INPUT TYPE="Hidden" NAME="EmployeeID" VALUE="<%=empID%>">

<TABLE border=0 width='30%'><tr><td><font size=1 face='ms sans serif, arial, geneva'>
<%
while not rs.eof
	ls_empid = rs("Employee_ID")
	ls_industryid = rs("Industry_ID")
	ls_industryname = rs("Industry_Name")
	'If ls_empid is null, then the user has not specified themselves
	'as an expert
	Response.Write("<INPUT TYPE=checkbox NAME=ind_expert VALUE=" & rs("Industry_ID"))
	if not isnull(rs("Employee_ID")) then 
	    Response.Write(" CHECKED")
	end if
	Response.Write(">" & ls_IndustryName & "<BR>" & NL)
	rs.movenext
wend

rs.close
set rs = nothing

%>
</font></td></tr></TABLE>

<INPUT TYPE="Hidden" NAME="InsertUpdate" VALUE="IND"><br>
<table align=center width=75% border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Update Experience" OnClick='document.frmInfo.submit();'>
			<input type=button class=button value="Cancel" OnClick='window.close();'>
		</td>
	</tr>
</table>
</form>
</center>

</body>
</html>

