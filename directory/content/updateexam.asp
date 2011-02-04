<!--
Developer:    David Martin
Date:         08/31/2000
Description:  Allows admin to update an Exam
-->

<!-- #include file="../../script.asp" -->

<HTML>
<HEAD>
<TITLE>Edit Exam "updateexam.asp"</TITLE>
<!-- #include file="../../style.htm" -->

<%

epid = Request.QueryString("epid")
'Response.Write("select Exam_id, Employeeid, Exam_date, Exams_passed_id, comments from exams_passed where exams_passed_id = "& epid)
Response.Write("<!-- epid: " & epid & "-->") 
%>

	<form name=frmInfo action='deleteexam.asp'>
	<INPUT Type=Hidden Name='DelFlag' Value=''>	<INPUT Type=Hidden Name='epid' Value=<% =epid %>>		
	</form>
</HEAD>

<%
	'-------------------------------------------
	'	Execute Database Query		
	'-------------------------------------------

empID = Request.QueryString("employeeID")

qry="select Exams_passed.Exam_id, Exam_list.Exam_Name, Exams_passed.Employeeid, Exams_passed.Exam_date, Exams_passed.Exams_passed_id, Exams_passed.comments from exams_passed, Exam_list where exams_passed_id = "& epid
'Response.Write "qry='" & qry & "'<br>"
set rs = DBQuery(qry)


	'-----------------------------------------------
	'	Get exam info							
	'-----------------------------------------------
	ls_examid = rs("exam_id")
	ls_examname = rs("exam_name")
	ls_examdate = rs("exam_date")
	ls_comments = rs("comments")	
	
rs.close

%>

<body bgcolor=silver onLoad="window.resizeBy(0, -200)"> 
<!--<center> -->

<table border=0><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
Please enter employee exam information.
</b></font></td></tr></table>


<form NAME="frmUpdate"  ACTION="editPreviousexam.asp">
<INPUT TYPE="Hidden" NAME="EmployeeID" VALUE="<%=empID%>"> 

<table border=0 height=200>

<tr>
	<td align=left><font size=1 face="ms sans serif, arial, geneva">Exam Passed Id:</font></td>
	<td><INPUT TYPE="Text" NAME="txtep_id" VALUE="<%=epid%>" size=4 maxlength=50 readonly></td>
</tr>



<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Exam Name:</font></td>
	<td>
		<Select NAME="txtexam_id">
			<%set rs = DBQuery("select Exam_ID, Exam_Name from Exam_List order by Exam_name")
			
							if not rs.eof and not rs.bof then
					do while not rs.eof %>
						<Option Value=<%=rs("Exam_ID")%> <%if trim(rs("Exam_ID")) = trim(ls_examid) then%> &nbsp; Selected
						<%end if%>><%=trim(rs("Exam_Name"))%>
						<% rs.movenext
					loop
					
					rs.close
				end if%>			
		</Select>
	</td>
</tr>

<tr>
	<td align=left><font size=1 face="ms sans serif, arial, geneva">Date Passed:</font></td>
	<td><INPUT TYPE="Text" NAME="txtexam_date" VALUE="<%=ls_examdate%>" size=11 maxlength=11></td>
</tr>

<tr>
	<td align=left><font size=1 face="ms sans serif, arial, geneva">Comments:</font></td>
	<td><INPUT TYPE=text NAME="txtComments" VALUE="<%=ls_Comments%>" SIZE=50 MAXLENGTH="250"></td>
</tr>

</table>

	<p>
	<INPUT TYPE=button CLASS=button VALUE="Update" OnClick='submit()' id=button1 name=button1>
	<INPUT TYPE=button CLASS=button VALUE="Cancel" OnClick='window.close();' id=button2 name=button2>
	<INPUT type=button class=button value='Delete ' onclick='deleteexam();' id=button3 name=button3>

</form>

</BODY>

<script language="JavaScript" type="text/javascript">

		
<!--

function deleteexam()
{
	if (confirm("Are you sure you want to delete this exam?"))
	{
		document.frmInfo.DelFlag.value = "True"
		document.frmInfo.submit()
	}
}

// -->
</script>
</HTML>

