<!--
Developer:	  David Martin
Date:         08/31/2000
Description:  Allows an exam to be added.
Modifications: Adam Lewandowski - 10/02/2000
               Added Exam Category field
-->
<% @language=VBscript%>

<!-- #include file="../section.asp" -->

<%
if Request.Form("Submitted") = "True" then
	exam_entered = Request.form("txtexam_name")
	examn_entered = Request.form("txtexam_number")
	exam_category = Request.Form("txtexam_category")
	fcert = Request.Form("fcert")
	
end if
	
if exam_entered <> "" then
	sql="Select * from Exam_list where exam_name ='" & exam_entered & "'"
	set rs = DBQuery(sql)
		if rs.bof then
			sSQL="Insert into Exam_list (exam_name, exam_number, cert, Category_ID) " & _
			     "values ('" & exam_entered & "', '" & examn_entered & "', '" & _
			     fcert & "', " & exam_category & ")"
			set rs = DBQuery(sSQL)
		end if
end if
	
%>

<HTML>
<HEAD>
<title>Add New Exam</title>

</HEAD>
<BODY>

<table border=0>
<tr>
	<td colspan=2><font size=2 color=red>
		<b>ADD EXAM</b></font>
	</td>
</tr>
<tr>
    <td width=100%><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>
<!--To add a new certification to the database, enter the certification name and --> 
<!--then click 'Submit' when finished:</STRONG></font>-->
	</td>
</tr></table><br>

<form NAME="frmAddexam" Action="add_exam.asp" method=post>
	<input type=hidden name="Submitted" value="False"><br><br>
	<table border=0 width=100%>
<tr>
    <TD>
		<P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Exam Name:</STRONG> </FONT></P>
    </TD>
    <TD></TD>
    <TD><INPUT size="50" name="txtexam_name" >
    </TD>
</tr>
<tr>
    <TD>
		<P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Exam Number:</STRONG> </FONT></P>
    </TD>
    <TD></TD>
    <TD><INPUT size="10" maxlength="10" name="txtexam_number" >
    </TD>
</tr>
<tr>
    <TD>
		<P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Exam Category:</STRONG> </FONT></P>
    </TD>
    <TD></TD>
    <TD><SELECT name="txtexam_category">
    <%
     set rs = DBQuery("SELECT Category_ID, Category_Description FROM Exam_Categories")
     do while not rs.eof
      Response.Write("<OPTION VALUE='" & rs("Category_ID") & "'>" & rs("Category_Description"))
      rs.movenext
     loop
     rs.close
    %>
    </SELECT>
    </TD>
</tr>
<tr><td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Certification:</STRONG></font></td>
	<td></td>
	
	<td><INPUT TYPE=checkbox NAME='fcert' VALUE='1'></td></tr> 
 
</tr>
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
		if (document.frmAddexam.Submitted.value = "True")
		{
			exam_entered = document.frmAddexam.textexam_name
			examn_entered = document.frmAddexam.textexam_number
			fcert = document.frmAddexam.fcert
		}
		if (document.frmAddexam.txtexam_name.value=="")
		{
			alert("Exam cannot be blank!")
		}
		if (document.frmAddexam.txtexam_number.value=="")
		{
			alert("Exam number cannot be blank!")
		}
		else
		{
			
			
				document.frmAddexam.Submitted.value = "True"
				document.frmAddexam.submit()
			
		}
	}
	
function GoHome()
	{
		window.navigate("default.asp")
	}
//-->
</SCRIPT>