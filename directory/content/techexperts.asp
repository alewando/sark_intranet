<!--
Developer:    GINGRICE
Date:         09/10/1998
Description:  Allows specifying Technical Subject Experts
-->

<!-- #include file="../../script.asp" -->

<html>

<head>
<title>Technology expertise</title>
<!-- #include file="../../style.htm" -->

<SCRIPT language = javascript>
<!--function OpenWin(page, EmpID, WinTitle){
	winTech= window.open(page + "?empID=" + EmpID, WinTitle,"height=500,width=500,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbars=0")
	}

// --></SCRIPT>

<%
'-------------------------------------------------
'	Execute Database Query for Tech Experts Info  
'-------------------------------------------------
NL = chr(13) & chr(10)
empID = Request.QueryString("EmpID")
' Remember to only select records where tech.Approved = 1!

ls_sql = "SELECT employee_tech_xref.employee_id, " & _
         "       tech.tech_id, " & _
         "       tech_area.tech_area, " & _
         "       tech.tech_name " & _
         "FROM tech, tech_area, employee_tech_xref " & _
         "WHERE tech.tech_id = employee_tech_xref.tech_id " & _
         "  AND tech.tech_area_id = tech_area.tech_area_id " & _
         "  AND employee_tech_xref.employee_id = " & empid & _
         "  AND Tech.Approved = 1 " & _
         "UNION " & _
         "SELECT null, " & _
         "       tech.tech_id, " & _
         "       tech_area.tech_area, " & _
         "       tech.tech_name " & _
         "FROM tech, tech_area " & _
         "WHERE tech.tech_area_id = tech_area.tech_area_id " & _
         "  AND Tech.Approved = 1 " & _
         "  AND tech.tech_id NOT IN " & _
         "(SELECT tech.tech_id " & _
         " FROM employee_tech_xref, Tech " & _
         " WHERE employee_tech_xref.tech_id = tech.tech_id " & _
         "   AND employee_tech_xref.employee_id = " & empID & _
         "   AND Tech.Approved = 1 " & _
         ")" & _
         "ORDER BY tech_area, tech_name"


Response.Write("<!--" & vbNewLine)
Response.Write("EmpID = " & empID & "<BR>" & vbNewLine)
Response.write("ls_sql = '" & ls_sql & "'" & vbNewLine)
Response.Write("-->")

set rs = DBQuery(ls_sql)
%>

<body bgcolor=silver link=blue alink=purple vlink=blue><center>

<table border=0 width='90%'><tr><td><font size=1 face="ms sans serif, arial, geneva">
<b>Please specify the technologies in which you consider yourself a subject matter expert.</b><p>
<font color=maroon><b>Note:</b> As an expert, you are designating yourself as a reference to other consultants, and may receive phone calls or emails relating to the specified technology.</font>
</FONT></td></tr></table>

<form NAME="frmInfo" ACTION="updateemp.asp">
<INPUT TYPE="Hidden" NAME="EmployeeID" VALUE="<%=empID%>">

<TABLE border=0><tr><td><font size=1 face='ms sans serif, arial, geneva'>
<%
while not rs.eof
	ls_empid = rs("employee_id")
	ls_techid = rs("tech_id")
	ls_techname = rs("tech_area") & ", " & rs("tech_name")
	
	'If ls_empid is null, then the user has not specified themselves
	'as an expert
	Response.Write("<INPUT TYPE=checkbox NAME=tech_expert VALUE=" & rs("tech_id"))
	if not isnull(rs("employee_id")) then Response.Write(" CHECKED")
	Response.Write(">" & ls_techname & "<BR>" & NL)
	rs.movenext
	wend
rs.close
set rs = nothing

%>
</font></td></tr></TABLE>

<INPUT TYPE="Hidden" NAME="InsertUpdate" VALUE="T"><br>
<table align=center width=75% border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Add Technology" OnClick='window.close();OpenWin("addnewtech.asp", "<%=empID%>", "AddNewTech");'>
			<input type=button class=button value="Update Expertise" OnClick='document.frmInfo.submit();'>
			<input type=button class=button value="Cancel" OnClick='window.close();'>
		</td>
	</tr>
</table>
</form>
</center>

</body>
</html>

