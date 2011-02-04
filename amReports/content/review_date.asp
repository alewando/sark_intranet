<!--
Developer:	  SSeissiger
Date:         08/16/2000
Description:  Report - sark reviews by date
			  Part of the AM report set
-->
<% @language=VBscript%>
<%
'--------------------------------------------------
'Added By SSeissiger - Allows TBreuer to see the
'Review and sponsor reports for the 
'Sponsorship Program
'--------------------------------------------------

if (hasRole("AccountManager")) or hasRole("SponsorAdmin") then %>
<!-- #include file="../section.asp" -->

<html>
<head><title>Reviews by Date</title></head>

<body>
<%
sql = "SELECT r.next_review, e.lastname as CLastName,e.FirstName as cFirstName, e.employeeID as cEmpID, " &_
		"s.lastname as sLastName, s.firstname as sFirstName, s.EmployeeID as sEmpID, " &_
		"am.lastname as amLastName, am.firstname as amFirstName, am.EmployeeID as amEmpID " &_
		"FROM review r left join employee e on r.employeeid = e.employeeid " &_
		"left join employee s on r.sponsor_id = s.employeeid " &_
		"left join employee am on r.account_mgr_id = am.employeeid " &_
		"ORDER BY r.next_review" 

set rsReview = DBQuery(sql)
%>
<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Reviews by Date</b>
<p>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>
	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
		<td colspan=1 width=140 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Date</b></td>
		<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Consultant</b></td>
		<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Sponsor</b></td>
		<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Account Mgr</b></td>
	</tr>

	<%while not rsReview.eof
		cLastName = rsReview("cLastName")
		cFirstName = rsReview("cFirstName")
		cEmpID = rsReview("cEmpID")
		sLastName = rsReview("sLastName")
		sFirstName = rsReview("sFirstName")
		sEmpID = rsReview("sEmpID")
		amLastName = rsReview("amLastName")
		amFirstName = rsReview("amFirstName")
		amEmpID = rsReview("amEmpID")
	%>
		<tr>
		<td colspan=1 width=100 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%=rsReview("next_review")	%>
		</td>
		<td colspan=1 width=140 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<a href='../../directory/content/details.asp?EmpID=<%=cEmpID%>'><%=cLastName%>,&nbsp;<%=cFirstName%></a>
		</td>
		<td colspan=1 width=140 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%if not isnull(sLastName) then %>
			<a href='../../directory/content/details.asp?EmpID=<%=sEmpID%>'><%=sLastName%>,&nbsp;<%=sFirstName%></a>
				<%else%>&nbsp;
				<%end if%>
		</td>
		<td colspan=1 width=140 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%if not isnull(amLastName) then %>
			<a href='../../directory/content/details.asp?EmpID=<%=amEmpID%>'><%=amLastName%>,&nbsp;<%=amFirstName%></a>
				<%else%>&nbsp;
				<%end if%>
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