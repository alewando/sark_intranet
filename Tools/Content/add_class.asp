<!--
Developer:	  David Martin
Date:         08/31/2000
Description:  Allows an exam to be added.
-->
<% @language=VBscript%>

<!-- #include file="../section.asp" -->

<%
if Request.Form("Submitted") = "True" then
	class_entered = Request.form("txtclass_name")
	
end if
	
if class_entered <> "" then
	sql="Select * from class_list where class_name ='" & class_entered & "'"
	set rs = DBQuery(sql)
		if rs.bof then
			sSQL="Insert into Class_list (class_name) " & _
			     "values ('" & class_entered & "')"
			set rs = DBQuery(sSQL)
		end if
end if
	
%>

<HTML>
<HEAD>
<title>Add New Class</title>

</HEAD>
<BODY>

<table border=0>
<tr>
	<td colspan=2><font size=2 color=red>
		<b>ADD CLASS</b></font>
	</td>
</tr>
<tr>
    <td width=100%><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>
<!--To add a new class to the database, enter the class name and --> 
<!--then click 'Submit' when finished:</STRONG></font>-->
	</td>
</tr></table><br>

<form NAME="frmAddclass" Action="add_class.asp" method=post>
	<input type=hidden name="Submitted" value="False"><br><br>
	<table border=0 width=100%>
<tr>
    <TD>
		<P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Class Name:</STRONG> </FONT></P>
    </TD>
    <TD></TD>
    <TD><INPUT size="50" name="txtclass_name" >
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
		if (document.frmAddclass.Submitted.value = "True")
		{
			class_entered = document.frmAddclass.textclass_name
			
		}
		if (document.frmAddclass.txtclass_name.value=="")
		{
			alert("Class cannot be blank!")
		}
		
		else
		{
			
			
				document.frmAddclass.Submitted.value = "True"
				document.frmAddclass.submit()
			
		}
	}
	
function GoHome()
	{
		window.navigate("default.asp")
	}
//-->
</SCRIPT>