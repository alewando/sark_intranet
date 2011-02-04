<!--
Developer:	  SSeissiger
Date:         08/16/2000
Description:  Report - Sponsor/Sponsoree with review date
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
<b>Sponsors and Sponsorees With Next Review Date</b>
<br></br>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<br></br>
<TABLE BORDER=0 WIDTH=450>
<tr><td ALIGN=LEFT VALIGN=top>
<%
sql = "SELECT r.next_review, c.lastname as cLastName,c.FirstName as cFirstName, c.EmployeeID as cEmpID, " &_
		"s.lastname as sLastName, s.firstname as sFirstName " &_
		"FROM review r, employee c, employee s " &_
		"WHERE r.employeeid = c.employeeid AND r.sponsor_id = s.employeeid " &_
		"ORDER BY s.LastName, s.FirstName, r.next_review"

set rsReview = DBQuery(sql)
prev_sLastName = ""
%>
<table width=450 cellspacing=0 cellpadding=4 border=0 >
<tr><td colspan=2 height=1>
</td></tr>

<%while not rsReview.eof
	cLastName = rsReview("cLastName")
	cFirstName = rsReview("cFirstName")
	sLastName = rsReview("sLastName")
	sFirstName = rsReview("sFirstName")
	cEmpID = rsReview("cEmpID")
	review = rsReview("next_review")
	
	
	if prev_sLastName = "" OR prev_sLastName <> sLastName then%>
	<tr>
		<td width=200><font size=1 face= "ms sans serif, arial, geneva"  color = blue>
		<b><span onClick="toggle('<%=sLastName%>')" alt="Expand or collapse this section." style="cursor:hand"><%=sLastName%>,&nbsp;<%=sFirstName%></b>
		<span id='<%=sLastName &  "Header"%>' style="">(collapsed)</span>
		</td>	</tr>
	<%end if%>	<%prev_sLastName = sLastName%>		<tr><td colspan=2>
		<span id='<%=sLastName%>' style="display:none">			<table>									<% while prev_sLastName = sLastName%>
				<tr>
				<td width=150 valign=top><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;<a href='../../directory/content/details.asp?EmpID=<%=cEmpID%>'><%=cLastName%>,&nbsp;<%=cFirstName%></a><br>				</td>				<td width=100  valign=top><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;<%=review%></a><br>				</td>				</tr>
			 				<%	rsReview.movenext										if not rsReview.eof then						cLastName = rsReview("cLastName")
						cFirstName = rsReview("cFirstName")
						sLastName = trim(rsReview("sLastName"))
						sFirstName = rsReview("sFirstName")
						cEmpID = rsReview("cEmpID")						review = rsReview("next_review")
					else						prev_sLastName=""
					end if			   wend%>	
						
			</table>
		</span>		
	</td></tr>
	

<%wend
rsReview.close%>
</td>
</tr></table>
</tr>
</table>
<!-- #include file="../../footer.asp" -->
<% else %>
	<b><h2><font color=red>You are not currently an account manager,<br> therefore can not view the Account Manager Reports.
	</b></h2></font>

<%end if%>

