<!--
Developer:    David Martin
Date:         08/31/2000
Description:  Allows admin to update an Class
-->

<!-- #include file="../../script.asp" -->

<HTML>
<HEAD>
<TITLE>Edit Class "updateclass.asp"</TITLE>
<!-- #include file="../../style.htm" -->

<%

ctid = Request.QueryString("ctid")
'Response.Write("select Exam_id, Employeeid, Exam_date, Exams_passed_id, comments from exams_passed where exams_passed_id = "& epid)
Response.Write("<!-- ctid: " & ctid & "-->") 
%>

	<form name=frmInfo action='deleteclass.asp'>
	<INPUT Type=Hidden Name='DelFlag' Value=''>	<INPUT Type=Hidden Name='ctid' Value=<% =ctid %>>		
	</form>
</HEAD>

<%
	'-------------------------------------------
	'	Execute Database Query		
	'-------------------------------------------

empID = Request.QueryString("employeeID")

qry="SELECT Classes.Class_ID, Class_list.Class_Name, Classes.EmployeeID, Classes.Class_Start_Date, Classes.Class_taken_ID, Classes.comments, Classes.Class_End_Date FROM Classes, Class_list WHERE Class_taken_ID = "& ctid 
'Response.Write "qry='" & qry & "'<br>"
set rs = DBQuery(qry)


	'-----------------------------------------------
	'	Get exam info							
	'-----------------------------------------------
	ls_classid = trim(rs("class_id"))
	ls_classname = trim(rs("class_name"))
	ls_classStartdate = trim(rs("class_Start_Date"))
	ls_classEnddate = trim(rs("Class_End_Date"))
	ls_comments = trim(rs("comments"))
	
	
rs.close

%>

<body bgcolor=silver onLoad="window.resizeBy(0, -200)"> 
<!--<center> -->

<table border=0><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
Please enter employee class information.
</b></font></td></tr></table>


<form NAME="frmUpdate"  ACTION="editPreviousclass.asp">
<INPUT TYPE="Hidden" NAME="EmployeeID" VALUE="<%=empID%>"> 

<table border=0 height=200>

<tr>
	<td align=left><font size=1 face="ms sans serif, arial, geneva">Class Id:</font></td>
	<td><INPUT TYPE="Text" NAME="txtct_id" VALUE="<%=ctid%>" size=4 maxlength=50 readonly></td>
</tr>



<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Class Name:</font></td>
	<td>
		<Select NAME="txtclass_id">
			<%set rs = DBQuery("select Class_ID, Class_Name from Class_List order by Class_ID")
				if not rs.eof and not rs.bof then
					do while not rs.eof%>
						<Option Value=<%=trim(rs("Class_ID"))%> <%if trim(rs("Class_ID")) = ls_Classid then%> Selected<%end if%>><%=trim(rs("Class_Name"))%>
						<% rs.movenext
					loop
					rs.close
				end if%>			
		</Select>
	</td>
</tr>

<tr>
	<td align=left><font size=1 face="ms sans serif, arial, geneva">Start Date:</font></td>
	<td><INPUT TYPE="Text" NAME="txtclass_start_date" VALUE="<%=ls_classStartdate%>" size=11 maxlength=11></td>
</tr>

<tr>
	<td align=left><font size=1 face="ms sans serif, arial, geneva">End Date:</font></td>
	<td><INPUT TYPE="Text" NAME="txtclass_end_date" VALUE="<%=ls_classEnddate%>" size=11 maxlength=11></td>
</tr>


<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Location:</font></td>
	<td>
		<Select NAME="txtlocation_ID">
			<%set rs = DBQuery("select Location_ID, Location_Name from Location order by Location_ID")
				if not rs.eof and not rs.bof then
					do while not rs.eof%>
						<Option Value=<%=rs("Location_ID")%> <%if trim(rs("Location_ID")) = ls_locationid then%> Selected<%end if%>><%=trim(rs("Location_Name"))%>
						<% rs.movenext
					loop
					rs.close
				end if%>			
		</Select>
	</td>
</tr>

<tr>
	<td align=left><font size=1 face="ms sans serif, arial, geneva">Comments:</font></td>
	<td><INPUT TYPE=text NAME="txtComments" VALUE="<%=ls_Comments%>" SIZE=50 MAXLENGTH="250"></td>
</tr>

</table>

	<p>
	<INPUT TYPE=button CLASS=button VALUE="Update" OnClick='submit();' id=button1 name=button1>
	<INPUT TYPE=button CLASS=button VALUE="Cancel" OnClick='window.close();' id=button2 name=button2>
	<INPUT type=button class=button value='Delete ' onclick='deleteclass();' id=button3 name=button3>

</form>

</BODY>

<script language="JavaScript" type="text/javascript">
<!--

function deleteclass()
{
	if (confirm("Are you sure you want to delete this class?"))
	{
		document.frmInfo.DelFlag.value = "True"
		document.frmInfo.submit()
	}
}

// -->
</script>
</HTML>

