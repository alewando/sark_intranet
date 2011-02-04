<!--
Developer:	  Julie Walters
Date:         01/02/2001
Description:  Report - SARK reviews past due
			  Part of the AM report set
-->
<% @language=VBscript%>
<%
'--------------------------------------------------
'Added By Julie Walters - Allows CDolan to see who's
'1 year reviews have not been held yet but are past due 
'on their review date.
'--------------------------------------------------

if (hasRole("AccountManager")) or hasRole("SponsorAdmin") then %>
<!-- #include file="../section.asp" -->

<html>
<head><title>Reviews Past Due</title></head>

<body>
<%
sql = "SELECT r.next_review, DATEDIFF(dd, r.next_review, GETDATE()) AS review_days, " &_
	"e.lastname AS CLastName, e.FirstName AS cFirstName, e.employeeID AS cEmpID " &_
	"FROM review r LEFT JOIN employee e ON r.employeeid = e.employeeid " &_
	"WHERE DATEDIFF(dd, r.next_review, GETDATE()) > 0 " &_
    "ORDER BY e.ClastName, r.next_review"

set rsReview = DBQuery(sql)
%>
<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Reviews Past Due</b>
<p>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>
	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
		<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Consultant</b></td>
		<td colspan=1 width=140 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Review Date</b></td>
		<td colspan=1 width=140 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Days Past Due</b></td>
	</tr>

	<%while not rsReview.eof
		cLastName = rsReview("cLastName")
		cFirstName = rsReview("cFirstName")
		cEmpID = rsReview("cEmpID")
		nextReview = rsReview("next_review")
		rvwDays = rsReview("review_days")
	%>
		<tr>
		<td colspan=1 width=140 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<a href='../../directory/content/details.asp?EmpID=<%=cEmpID%>'><%=cLastName%>,&nbsp;<%=cFirstName%></a>
		</td>
		<td colspan=1 width=100 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%=nextReview	%>
		</td>
		<td colspan=1 width=100 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%=rvwDays %>
		</td>
		</tr>
		<%rsReview.movenext
	wend
	rsReview.close
	%>
</table>
</body>
</html>

<!-- #include file="../../footer.asp" -->
<% else %>
	<b><h2><font color=red>You are not currently an account manager, therefore can not view the Account Manager Reports.
	</b></h2></font>

<%end if%>