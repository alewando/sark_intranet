<!--
Developer:	  SSeissiger
Date:         08/16/2000
Description:  Report - sark reviews by date
			  Part of the AM report set
-->
<%
'--------------------------------------------------
'Added By SSeissiger - Allows TBreuer to see the
'Review and sponsor reports for the 
'Sponsorship Program
'--------------------------------------------------

if (hasRole("AccountManager")) or hasRole("SponsorAdmin") then %>
<!-- #include file="../section.asp" -->

<script language=javascript>
function toggle(item){
	var docItem = document.all(item)
	if (docItem.style.display == ""){
		docItem.style.display = "none"
		document.all(item + "Header").innerHTML = "(collapsed)"
		}
	else {
		docItem.style.display = ""
		document.all(item + "Header").innerHTML = ""
		}
	}
</script>
<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Next Review Date Listed by Account Manager</b>
<br></br>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<br></br>
<TABLE BORDER=0 WIDTH=450>
<tr><td ALIGN=LEFT VALIGN=top>
<!--''''''''''''''''''''''''''''''''''''''''''''''
'Code to display a list of Account Managers 
'List Sarks under them - sorted by review date
'And include sponsor
''''''''''''''''''''''''''''''''''''''''''''''-->
<%

sql = "SELECT r.next_review, s.lastname as sLastName, s.firstname as sFirstName, s.EmployeeID as sEmpID, " &_
		"am.lastname as amLastName, am.firstname as amFirstName, am.EmployeeID as amEmpID, " &_
		"c.lastname as cLastName, c.FirstName as cFirstName, c.EmployeeID as cEmpID " &_
		"FROM review r left join employee c on r.employeeid = c.employeeid " &_
		"left join employee s on r.sponsor_id = s.employeeid " &_
		"join employee am on r.account_mgr_id = am.employeeid " &_
		"ORDER BY am.LastName,r.next_review" 

set rsReview = DBQuery(sql)
prev_amLastName = ""
%>
<table width=400 cellspacing=0 cellpadding=4 border=0 >
<tr><td colspan=2 height=1>
</td></tr>


<%while not rsReview.eof
	cLastName = rsReview("cLastName")
	cFirstName = rsReview("cFirstName")
	amLastName = trim(rsReview("amLastName"))
	amFirstName = rsReview("amFirstName")
	cEmpID = rsReview("cEmpID")
	amEmpID = rsReview("amEmpID")
	
	sEmpID = rsReview("sEmpID")
	sLastName = rsReview("sLastName")
	sFirstName = rsReview("sFirstName")
	review = rsReview("next_review")
	
	if prev_amLastName = "" OR prev_amLastName <> amLastName then%>
	<tr>
		<td colspan=1 width=150 align=left><font size=1 face= "ms sans serif, arial, geneva"  color = maroon>
		<b><span onClick="toggle('<%=amLastName%>')" alt="Expand or collapse this section." style="cursor:hand"><%=amLastName%>,&nbsp;<%=amFirstName%></b>
		<span id='<%=amLastName &  "Header"%>' style="">&nbsp;&nbsp;(collapsed)</span>
		</td>
	<%end if%>
		<span id='<%=amLastName%>' style="display:none">
			<% if prev_amLastName = amLastName then %>
					<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=Maroon>
					<td valign=top width=90 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>
					<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>
				</tr>
			<%end if%>
				<tr>
				<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>
				<td width = 150 valign=top colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>
				<%if not isnull(sLastName) then %>
				<a href='../../directory/content/details.asp?EmpID=<%=sEmpID%>'><%=sLastName%>,&nbsp;<%=sFirstName%></a><br>
				<%end if%>
				</td>
			 
					
						cFirstName = rsReview("cFirstName")
						amLastName = trim(rsReview("amLastName"))
						amFirstName = rsReview("amFirstName")
						cEmpID = rsReview("cEmpID")
						sEmpID = rsReview("sEmpID")
						sLastName = rsReview("sLastName")
						sFirstName = rsReview("sFirstName")
					else
					end if
						
			</table>
		</span>
	</td></tr>
	

<%wend
rsReview.close

</tr>

</table>
<!-- #include file="../../footer.asp" -->
<% else %>
<b><h2><font color=red>You are not currently an account manager, therefore can not view the Account Manager Reports.
</b></h2></font>

<%end if%>