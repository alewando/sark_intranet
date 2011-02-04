<XML ID="skills_inv" src="skills_inv.asp?EmployeeID=<% =request("EmpID") %>"></XML>
<!-- #include file="../../script.asp" -->
<html>
<form id="skills" name="skills" action="SkillsSubmit.asp" method="POST">
<head>
<title>Skills Inventory</title>
<!-- #include file="../../style.htm" -->
</head>

<body bgcolor=silver>
 
<center><p><b>Skills Inventory</b></p>
  <table border=0 cellspacing=0 cellpadding=0 datasrc="#skills_inv" width="460">
	<tr>
		<td width="160" valign=top>
			<b>&nbsp Employee Name &nbsp</b> 
		</td>
		<td width="300" valign=top>
			<span datafld="FirstName"></span>&nbsp<span datafld="LastName"></span>
		</td>
	</tr>
	<tr>
		<td width="160" valign=top>
			<b>&nbsp Date Modified &nbsp</b> 
		</td>
		<td width="300" valign=top>
			<span datafld="DateSkillsModified"></span>
		</td>
	</tr>
  </table>
</center>

<p>This document provides an assessment of consultant technical skills and should be
used to facilitate discussion between the reviewer and consultant concerning
the progress of technical education and in setting consultant goals.</p>

<p>A preliminary assessment should be filled out by the consultant prior to the
review. The reviewer should note any areas of change from previous reviews,
discern areas for future development and make any rating adjustments. The
reviewer should be prepared to discuss these items with the consultant at the
time of the review.</p>

<b>Explanation of Ratings</b>

<p>
<table border=1 cellspacing=1 cellpadding=1 width="90%">
 <tr>
  <td width="30%" valign=top>
	<p>N/A</p>
  </td>
  <td width="60%" valign=top >
	<p>Consultant has had no exposure to the technology.</p>
  </td>
 </tr>
 <tr>
  <td width="30%" valign=top>
	<p>Basic Education</p>
  </td>
  <td width="60%" valign=top>
	<p>Consultant has attended a class in which the technology was taught or is self-taught,
		but has no actual work experience.<o:p></o:p></span></p>
  </td>
 </tr>
 <tr>
  <td width="30%" valign=top>
	<p>Some Experience</p>
  </td>
  <td width="60%" valign=top>
	<p>Consultant has had some exposure to the technology as it is applied in a working environment.</p>
  </td>
 </tr>
 <tr>
  <td width="30%" valign=top>
	<p>Productive w/minimal assistance</p>
  </td>
  <td width="60%" valign=top>
	<p>Consultant can be immediately productive in a working environment, with infrequent assistance.</p>
  </td>
 </tr>
 <tr>
  <td width="30%" valign=top>
	<p>Productive w/o assistance</p>
  </td>
  <td width="60%" valign=top>
	<p>Consultant can be immediately productive in a working environment, with no assistance.</p>
  </td>
 </tr>
 <tr>
  <td width="30%" valign=top>
	<p>Expert</p>
  </td>
  <td width="60%" valign=top>
	<p>Consultant has extensive, detailed working knowledge of the technology and can assist
		others with complex technical issues.</p>
  </td>
 </tr>
</table>
</p>
<p>

<table id=table datasrc=#skills_inv datafld="SkillGroup" border=0 width="460" cellpadding=1>
	<tr><tr></tr>
		<td valign=top width="460">
			<b><div datafld="GroupName"></div></b>
		</td>
		<tr>
			<td valign=top width="460">
				<table  border=1 width="460" cellpadding=0 datasrc=#skills_inv datafld=SkillGroup.Skill>
					<tr>
						<td valign=top width="380">
							<div datafld="SkillGroup.Skill.SkillName" dataformatas=HTML></div>
						</td>
						<td valign=top width="80">
							<input TYPE="hidden" NAME="SkillID" DATASRC="#skills_inv" span DATAFLD="SkillGroup.Skill.SkillID">
							<select name="SkillRanking" datasrc="#skills_inv" span datafld="SkillGroup.Skill.SkillRanking">
							<option value=1>N/A
							<option value=2>Basic Education
							<option value=3>Some Experience
							<option value=4>Productive w/minimal assistance
							<option value=5>Productive w/o assistance
							<option value=6>Expert
							</select>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</tr>
</table>

<center>
<br>
<input name="Submit" type="SUBMIT" class=button value="Update Skills">
<input name="Close" type=button class=button value="Cancel" OnClick='window.close();'>
</center>
<%
sql = "select e.employeeid from employee e where e.username = " & "'" & session("Username") & "'"set rs = DBQuery(sql)empID = rs("employeeid")
if (trim(empID) = trim(request("EmpID"))) Then %>
<script>
	document.skills.SkillID.disabled = false
	top.skills.SkillRanking.disabled = false
	top.skills.Submit.disabled = false
</script>
<% Else %>
<script>
	document.skills.SkillID.disabled = true
	document.skills.SkillRanking.disabled = true
	document.skills.Submit.disabled = true
</script>
<% End If %>
</form>
</body>
</html>
