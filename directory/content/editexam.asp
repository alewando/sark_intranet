<!--
Developer:    GINRICE code modified by DMARTIN
Date:         08/28/2000
Description:  Allows editing of exam info
-->


<!-- #include file="../../script.asp" -->

<html>

<head>
<title>Update Exam Info</title>
<!-- #include file="../../style.htm" -->
</head>


<%
'-----------------------------------------------
'	Execute Database Query for Exam Info		
'-----------------------------------------------
empID = Request.QueryString("EmpID") 
set rs = DBQuery("SELECT Exam_Name" _
	& " FROM Exam_list INNER JOIN" _
	& " Exams_passed ON" _
    & " Exams_passed.Exams_passed_ID = Exam_list.Exam_ID" _
	& " WHERE Exams_passed.EmployeeID = " & empid & "")


if not rs.eof then
	ls_insertupdate = "U"
	'-----------------------------------------------
	'	Get Exam Info							
	'-----------------------------------------------
	ls_examid =	  trim(rs("exam_id"))
	ls_examname = 	trim(rs("exam_name"))
	ls_examdate = 	trim(rs("exam_date"))
	ls_examcomments =	trim(rs("comments"))

else
	ls_insertupdate = "I"
end if
%>

<body bgcolor=silver><center>

<table border=0 width="90%"><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
This form allows for updating exam information.
</b></font></td></tr></table>

Response.Write("<center><font color=navy><b>" & rs("examname") & " " & rs("empid") & ",&nbsp;&nbsp;" & ls_title & "</b></font></center><br>") 






<FORM action=updateexam2.asp name=frmInfo>
<INPUT TYPE="hidden" NAME="EmployeeID" VALUE="<%=empID%>">&nbsp; 

<table border=0 style="HEIGHT: 193px; WIDTH: 452px">
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Exam ID:</font></td>
	<td><INPUT NAME="examid" VALUE=" <%=ls_examid%>" size=10 maxlength=25></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Exam Name:</font></td>
	<td><INPUT NAME="examname" VALUE=" <%=ls_examname%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Date Passed:</font></td>
	<td><INPUT NAME="datepassed" VALUE=" <%=ls_examdate%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Comments:</font></td>
	<td><INPUT NAME="comments" VALUE=" <%=ls_examcomments%>" size=40 maxlength=255></td>
</tr>
  <TR>

		<td align=middle>
			<input type=button class=button value="Update Exam" OnClick='document.frmInfo.submit();'>
			<input type=button class=button value="Cancel" OnClick='window.close();' style="LEFT: 124px; TOP: 4px">
		</td></TR>
</table></FORM>
</center>

</body>
</html>
