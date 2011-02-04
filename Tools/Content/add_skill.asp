<!--
Developer:     Todd Breuer
Date:          10/07/2000
Description:   Allows a new skill to be added.  It will also add the new skill to the 
               Skills_Inventory table
Modifications: 
-->
<% @language=VBscript%>

<!-- #include file="../section.asp" -->

<%
if Request.Form("Submitted") = "True" then
	skill_name = Request.form("txtskill_name")
	group_category = Request.Form("txtgroup_category")
	other_skill = Request.Form("otherSkill")

	if other_skill = "" then
		other_skill = "0"
	end if
end if
	
if skill_name <> "" then
	sql="Select * from Skills where SkillName ='" & skill_name & "'"
	set rs = DBQuery(sql)
		if rs.bof then
			sSQL="Insert into Skills (SkillName, GroupID, isOtherSkill) " & _
			     "values ('" & skill_name & "', " & group_category & ", " & _
			     other_skill & ")"
			set rs = DBQuery(sSQL)

			set rs=DBQuery("Select SkillID from Skills Where SkillName = '" & _
			    skill_name & "'")
			skill_ID=rs("SkillID")

			sSQL="Insert into Skills_Inventory (EmployeeID, SkillRanking, " & _
			     "SkillID, SkillValue) (Select Distinct EmployeeID, 1, " & _
			     skill_ID & ", '' from Skills_Inventory)"
			set rs=DBQuery(sSQL)
		end if
end if
	
%>

<HTML>
<HEAD>
<title>Add New Skill</title>

</HEAD>
<BODY>

<table border=0>
<tr>
	<td colspan=2><font size=2 color=red>
		<b>ADD SKILL</b></font>
	</td>
</tr>
<tr>
    <td width=100%><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>
<!--To add a new certification to the database, enter the certification name and --> 
<!--then click 'Submit' when finished:</STRONG></font>-->
	</td>
</tr></table><br>

<form NAME="frmAddSkill" Action="add_skill.asp" method=post>
	<input type=hidden name="Submitted" value="False"><br><br>
	<table border=0 width=100%>
<tr>
    <TD>
		<P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Skill Name:</STRONG> </FONT></P>
    </TD>
    <TD></TD>
    <TD><INPUT size="50" name="txtskill_name" >
    </TD>
</tr>
<tr>
    <TD>
		<P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Group Category:</STRONG> </FONT></P>
    </TD>
    <TD></TD>
    <TD><SELECT name="txtgroup_category">
    <%
     set rs = DBQuery("SELECT GroupID, GroupName FROM Skills_Groups")
     do while not rs.eof
      Response.Write("<OPTION VALUE='" & rs("GroupID") & "'>" & rs("GroupName"))
      rs.movenext
     loop
     rs.close
    %>
    </SELECT>
    </TD>
</tr>
<tr><td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Is Other Skill Type:</STRONG></font></td>
	<td></td>
	
	<td><INPUT TYPE=checkbox NAME='otherSkill' VALUE='1'></td></tr> 
 
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
		if (document.frmAddSkill.Submitted.value = "True")
		{
			skill_name = document.frmAddSkill.textskill_name
			group_category = document.frmAddSkill.textgroup_category
			other_skill = document.frmAddSkill.otherSkill
		}
		if (document.frmAddSkill.txtskill_name.value=="")
		{
			alert("Skill Name cannot be blank!")
		}
		else
		{
				document.frmAddSkill.Submitted.value = "True"
				document.frmAddSkill.submit()
		}
	}
	
function GoHome()
	{
		window.navigate("default.asp")
	}
//-->
</SCRIPT>