<!--
Developer:    GINRICE code modified by DMARTIN
Date:         08/28/2000
Description:  Allows editing of exam info
-->


<!-- #include file="../../script.asp" -->

<html>

<head>
<title>Add Exam Info "addexam.asp"</title>
<!-- #include file="../../style.htm" -->
</head>
<body bgcolor=silver onLoad="window.resizeBy(0, -200)"> 

<%
'-----------------------------------------------
'	Execute Database Query for Exam Info		
'-----------------------------------------------
empID = Request.QueryString("EmpID") 
set rs = DBQuery("SELECT Exam_List.Exam_Name, Exam_List.Exam_ID," _
    & " Exams_passed.Exam_Date, Exams_passed.Comments" _
	& " FROM Exam_list INNER JOIN" _
	& " Exams_passed ON" _
    & " Exams_passed.Exam_ID = Exam_list.Exam_ID" _
	& " WHERE Exams_passed.EmployeeID = " & empid & "")


if not rs.eof then
	
	'-----------------------------------------------
	'	Get Exam Info							
	'-----------------------------------------------
	ls_examid =	  trim(rs("Exam_ID"))
	ls_examname = 	trim(rs("Exam_Name"))
	ls_examdate = 	trim(rs("Exam_Date"))
	ls_examcomments =	trim(rs("Comments"))

	ls_insertupdate = "U"
else
	ls_insertupdate = "I"
end if
%>

<body bgcolor=silver><center>

<table border=0 width="90%" Height="" ><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
This form allows for the adding of exam information. 
</b></font></td></tr></table>

</script>


<FORM action=updateexam2.asp name=frmInfo>
<INPUT TYPE="hidden" NAME="EmployeeID" VALUE="<%=empID%>">&nbsp; 

<table border=0 style="HEIGHT: 193px; WIDTH: 452px">
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Exam Name:</font></td>
	<td>
		<Select NAME="Exam_ID">
			<%set rs = DBQuery("select Exam_ID, Exam_Name from Exam_List order by Exam_Name")
				if not rs.eof and not rs.bof then
					do while not rs.eof%>
						<Option Value=<%=rs("Exam_ID")%> <%if trim(rs("Exam_ID")) = ls_examid then%> Selected<%end if%>><%=trim(rs("Exam_Name"))%>
						<% rs.movenext
					loop
					rs.close
				end if%>			
		</Select>
	</td>
</tr>

<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Date Passed:</font></td>
	<td><INPUT NAME="exam_date" VALUE="<%=date()%>" size=11 maxlength=11></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Comments:</font></td>
	<td><INPUT NAME="comments" VALUE="" size=40 maxlength=255></td>
</tr>
  <TR>

		<td align=middle colspan=2>
	<INPUT TYPE="Hidden" NAME="InsertUpdate" VALUE="<%=ls_insertupdate%>"><br>
            <input type=button class=button value="Submit" OnClick='document.frmInfo.submit()' id=button1 name=button1>
			<input type=button class=button value="Cancel" OnClick='window.close();' id=button2 name=button2>
		</td></TR>
</table></FORM>
</center>

</body>
</html>
