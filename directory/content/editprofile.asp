<!--
Developer:    GINGRICE
Date:         09/10/1998
Description:  Allows editing of profile
-->


<!-- #include file="../../script.asp" -->

<html>

<head>
<title>Update Profile</title>
<!-- #include file="../../style.htm" -->
</head>


<%
'-----------------------------------------------
'	Execute Database Query for Profile Info		
'-----------------------------------------------
empID = Request.QueryString("EmpID") 
set rs = DBQuery("select * from profile where employee_id = " & empID)
if not rs.eof then
	ls_insertupdate = "U"
	'-----------------------------------------------
	'	Get Profile Info							
	'-----------------------------------------------
	ls_college =	trim(rs("college"))
	ls_hobbies = 	trim(rs("hobbies_sports_interests"))
	ls_vacation = 	trim(rs("dream_vacation"))
	ls_show =		trim(rs("favorite_show_ad"))
	ls_inspired = 	trim(rs("person_inspired"))
	ls_motto = 		trim(rs("favorite_motto"))
	ls_funniest = 	trim(rs("funniest_embarrassing"))
	ls_lotto = 		trim(rs("lotto"))
	ls_pet_peeve = 	trim(rs("pet_peeve"))
	ls_possession =	trim(rs("prized_possession"))
	ls_challenge = 	trim(rs("biggest_challenge"))
else
	ls_insertupdate = "I"
end if
%>
<body bgcolor=silver onLoad="window.resizeBy(0, -150)"><center>

<table border=0 width='90%'><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
Please enter your profile information.
This is used to help introduce you to the rest of our company.
Please avoid the use of single or double quotes.
Thank you.
</b></font></td></tr></table>


<form NAME="frmInfo" ACTION="updateemp.asp">
<INPUT TYPE="Hidden" NAME="EmployeeID" VALUE="<%=empID%>"> 

<table border=0>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">College:</font></td>
	<td><INPUT TYPE="Text" NAME="college" VALUE="<%=ls_college%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Hobbies:</font></td>
	<td><INPUT TYPE="Text" NAME="hobbies" VALUE="<%=ls_hobbies%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Dream Vacation:</font></td>
	<td><INPUT TYPE="Text" NAME="vacation" VALUE="<%=ls_vacation%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Favorite TV Show/Ad:</font></td>
	<td><INPUT TYPE="Text" NAME="favoriteshow" VALUE="<%=ls_show%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Person Who Inspired Me Most:</font></td>
	<td><INPUT TYPE="Text" NAME="inspired" VALUE="<%=ls_inspired%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Favorite Motto:</font></td>
	<td><INPUT TYPE="Text" NAME="motto" VALUE="<%=ls_motto%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Funniest / Most Embarrassing:</font></td>
	<td><INPUT TYPE="Text" NAME="moment" VALUE="<%=ls_funniest%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">If I won the lotto...:</font></td>
	<td><INPUT TYPE="Text" NAME="lotto" VALUE="<%=ls_lotto%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Pet Peeve:</font></td>
	<td><INPUT TYPE="Text" NAME="petpeeve" VALUE="<%=ls_pet_peeve%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Prized Possession:</font></td>
	<td><INPUT TYPE="Text" NAME="possession" VALUE="<%=ls_possession%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Biggest Challenge:</font></td>
	<td><INPUT TYPE="Text" NAME="challenge" VALUE="<%=ls_challenge%>" size=40 maxlength=255></td>
</tr>
</table>

<INPUT TYPE="Hidden" NAME="InsertUpdate" VALUE="<%=ls_insertupdate%>"><br>
<table align=center width=75% border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Update Profile" OnClick='document.frmInfo.submit();'>
			<input type=button class=button value="Cancel" OnClick='window.close();'>
		</td>
	</tr>
</table>

</form>
</center>

</body>
</html>
