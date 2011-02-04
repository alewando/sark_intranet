<!--
Developer:	  SSeissiger
Date:         08/21/2000
Description:  
			  Part of the AM report set
-->
<% @language=VBscript%>
<% if (hasRole("AccountManager")) then %>
<!-- #include file="../section.asp" -->

<html>
<head><title>Intranet Survey Results</title></head>

<body>
<%
sql1_1 = "select count(q1_1) as onceDay FROM online_survey where q1_1 = '1'"
set rs1= DBQuery(sql1_1)
onceDay = rs1("onceDay")

sql1_2 = "select count(q1_2) as twiceDay FROM online_survey where q1_2 = '1'"
set rs2= DBQuery(sql1_2)
twiceDay = rs2("twiceday")

sql1_3 = "select count(q1_3) as onceWeek FROM online_survey where q1_3 = '1'"
set rs3= DBQuery(sql1_3)
onceWeek = rs3("onceWeek")

sql1_4 = "SELECT count(q1_4) as twiceWeek FROM online_survey where q1_4 = '1'"
set rs4= DBQuery(sql1_4)
twiceWeek = rs4("twiceWeek")

sql1_5 = "SELECT count(q1_5) as weekend FROM online_survey where q1_5 = '1'"
set rs5= DBQuery(sql1_5)
weekend = rs5("weekend")


sql6= "SELECT ROUND(avg(CAST(q2 AS DECIMAL(38,3))),2) as usefulness FROM online_survey"
set rs6= DBQuery(sql6)
usefulness = rs6("usefulness")


sql7 = "SELECT q3 FROM online_survey WHERE q3 != 'null'"
set rs7 = DBQuery(sql7)
notUsed = rs7("q3")


sql8 = "SELECT q4 FROM online_survey WHERE q4 != 'null'"
set rs8 = DBQuery(sql8)
likeSee = rs8("q4")

sql9 = "select count('q_2') as entries from online_survey"
set rs9 = DBQuery(sql9)
Entries = rs9("entries")

sql10 = "select count('employeeID') as emps from Employee"
set rs10 = DBQuery(sql10)
Emps = rs10("emps")

%>
<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Intranet Survey Results</b>
<br></br>
	<table>
	<tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=215 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Completed Surveys</b></td>
		<td colspan=1 width=90 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b><%=entries%> of <%=emps%></b></td>
	</tr>
	<tr></tr>
	<tr></tr>
	<tr></tr>
	<tr></tr>
	<tr>
		<td colspan=1 width=215 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Once a Day</b></td>
		<td colspan=1 width=250 align=left><font size=2 face="ms sans serif, arial, geneva" color=black><b><%=onceDay%></b></td>
	</tr>
	<tr></tr>
	<tr></tr>
	<tr>
		<td colspan=1 width=215 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>More Than Once a Day</b></td>
		<td colspan=1 width=250 align=left><font size=2 face="ms sans serif, arial, geneva" color=black><b><%=twiceDay%></b></td>
	</tr>
	<tr></tr>
	<tr></tr>
	<tr>
		<td colspan=1 width=215 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Once a Week</b></td>
		<td colspan=1 width=250 align=left><font size=2 face="ms sans serif, arial, geneva" color=black><b><%=onceWeek%></b></td>
	</tr>
	<tr></tr>
	<tr></tr>
	<tr>
		<td colspan=1 width=215 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>More Than Once a Week</b></td>
		<td colspan=1 width=250 align=left><font size=2 face="ms sans serif, arial, geneva" color=black><b><%=twiceWeek%></b></td>
	</tr>
	<tr></tr>
	<tr></tr>
	<tr>
		<td colspan=1 width=215 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>On the Weekend</b></td>
		<td colspan=1 width=250 align=left><font size=2 face="ms sans serif, arial, geneva" color=black><b><%=weekend%></b></td>
	</tr>
	<tr></tr>
	<tr></tr>
	<tr>
		<td colspan=1 width=215 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Usefulness (1 = most usefull)</b></td>
		<td colspan=1 width=250 align=left><font size=2 face="ms sans serif, arial, geneva" color=black><b><%=usefulness%></b></td>
	</tr>
	<tr></tr>
	<tr></tr>
	<!--<td colspan=1 width=215 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Items Not Used</b></td>
	-->
	<%'while not rs7.eof %>
	<!--	<tr>
		<td colspan=1 width=215 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon></td>
		<td colspan=1 width=250 align=left><font size=2 face="ms sans serif, arial, geneva" color=black><b><%=rs7("q3")%></b></td>
	</tr>
	<tr></tr>
	<tr></tr>
	-->
		<%'rs7.movenext
	'wend
	'rs7.close
	%>
	
	<td colspan=1 width=215 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>New Items Suggested</b></td>
	<%while not rs8.eof %>
	<tr>
		<td colspan=1 width=215 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon></td>
		<td colspan=1 width=250 align=left><font size=2 face="ms sans serif, arial, geneva" color=black><b><%=rs8("q4")%></b></td>
	</tr>
	<tr></tr>
	<tr></tr>
	
		<%rs8.movenext
	wend
	rs8.close
	%>
	
</table>
</body>
</html>

<!-- #include file="../../footer.asp" -->
<% else %>
	<b><h2><font color=red>You are not currently an account manager,<br> therefore can not view the Account Manager Reports.
	</b></h2></font>
<%end if%>

