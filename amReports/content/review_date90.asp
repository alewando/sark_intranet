<!--
Developer:	  SSeissiger
Date:         08/16/2000
Description:  Report - Sark reviews in the next (30/45/60/90) days out by A/M
			  Part of the AM report set
Modified:	  12/26/2000
Modified By:  Julie Walters (added breakouts by date ranges instead of just 1-45 days.
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
<b>Reviews in the Next 90 Days - Listed by Account Manager</b>
<br></br>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<br></br>
<TABLE BORDER=0 WIDTH=450>
<tr><td ALIGN=LEFT VALIGN=top>
<!--''''''''''''''''''''''''''''''''''''''''''''''
'Code to display a list of Account Managers 
'List Sarks under them - sorted by review date - in the next 90 days
'And include sponsor
''''''''''''''''''''''''''''''''''''''''''''''-->

<!--SQL to select employee's with 0-30 days out to review date-->
<%
sql = "SELECT r.next_review, DATEDIFF(dd, GETDATE(), next_review) AS reviewdays, " &_
	  "s.lastname AS sLastName, s.firstname AS sFirstName, s.EmployeeID AS sEmpID, " &_
      "am.lastname AS amLastName, am.firstname AS amFirstName, am.EmployeeID AS amEmpID, " &_
      "c.lastname AS cLastName, c.FirstName AS cFirstName, c.EmployeeID AS cEmpID " &_
      "FROM review r LEFT JOIN employee c ON r.employeeid = c.employeeid LEFT JOIN " &_
	  "employee s ON r.sponsor_id = s.employeeid JOIN employee am ON r.account_mgr_id " &_
	  " = am.employeeid WHERE DATEDIFF(dd, GETDATE(), next_review) > 0 AND " &_
	  "DATEDIFF(dd, GETDATE(), next_review) < 91 " &_
	  "ORDER BY am.LastName, am.FirstName, r.next_review"
	  
set rsReview = DBQuery(sql)
prev_amLastName = ""
%>

<%if rsReview.eof then %>
	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=400 align=left><font size=2 face='ms sans serif, arial, geneva' color=red>
	<b>Query Returned No Records</b>
	</td>
	</tr>
<% else %>

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
	rvwdays = rsReview("reviewdays")
	
	if prev_amLastName = "" OR prev_amLastName <> amLastName then%>
	<tr>
		<td colspan=1 width=150 align=left><font size=1 face= "ms sans serif, arial, geneva"  color = maroon>
		<b><span onClick="toggle('<%=amLastName%>')" alt="Expand or collapse this section." style="cursor:hand"><%=amLastName%>,&nbsp;<%=amFirstName%></b>
		<span id='<%=amLastName &  "Header"%>' style="">&nbsp;&nbsp;(collapsed)</span>
		</td>
		<!--<td colspan=1 width=175 width=70 align=center><font size=1 face= "ms sans serif, arial, geneva"  color = maroon><b>Sponsor</b>
		</td>
		-->	</tr>
	<%end if%>	<%prev_amLastName = amLastName%>		<tr><td colspan=2>
		<span id='<%=amLastName%>' style="display:none">			<table>			<!--Insert header and employees for reviews in next 1-30 days-->
			<% if prev_amLastName = amLastName then %>				<tr>
					<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>					&nbsp;&nbsp;<b>1-30 Days</b>					</td>
					<td valign=top width=90 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>					&nbsp;&nbsp;<b>Date</b>					</td>
					<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>					&nbsp;&nbsp;<b>Sponsor</b>					</td>
				</tr>				<% while prev_amLastName = amLastName and rvwdays > 0 and rvwdays < 31%>
				<tr>
				<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;<a href='../../directory/content/details.asp?EmpID=<%=cEmpID%>'><%=cLastName%>,&nbsp;<%=cFirstName%></a><br>				</td>				<td valign=top width = 90 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;<%=review%></a><br>
				<td width = 150 valign=top colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;
				<%if not isnull(sLastName) then %>
				<a href='../../directory/content/details.asp?EmpID=<%=sEmpID%>'><%=sLastName%>,&nbsp;<%=sFirstName%></a><br>				</td>				<%else%>&nbsp;
				<%end if%>
				</td>				</tr>
			 				<%	rsReview.movenext
										if not rsReview.eof then						cLastName = rsReview("cLastName")
						cFirstName = rsReview("cFirstName")
						amLastName = trim(rsReview("amLastName"))
						amFirstName = rsReview("amFirstName")
						cEmpID = rsReview("cEmpID")						amEmpID = rsReview("amEmpID")						
						sEmpID = rsReview("sEmpID")
						sLastName = rsReview("sLastName")
						sFirstName = rsReview("sFirstName")						review = rsReview("next_review")						rvwdays = rsReview("reviewdays")
					else						prev_amLastName=""
					end if			   wend%>				   
			   <!--Insert header and employees for reviews in next 31-45 days-->				<tr>
					<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>					&nbsp;&nbsp;<b>31-45 Days</td>
					<td valign=top width=90 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>					&nbsp;&nbsp;<b>Date</b>					</td>
					<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>					&nbsp;&nbsp;<b>Sponsor</b>					</td>
				</tr>
				<% while prev_amLastName = amLastName and rvwdays > 30 and rvwdays < 46%>
				<tr>
				<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;<a href='../../directory/content/details.asp?EmpID=<%=cEmpID%>'><%=cLastName%>,&nbsp;<%=cFirstName%></a><br>				</td>				<td valign=top width = 90 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;<%=review%></a><br>
				<td width = 150 valign=top colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;
				<%if not isnull(sLastName) then %>
				<a href='../../directory/content/details.asp?EmpID=<%=sEmpID%>'><%=sLastName%>,&nbsp;<%=sFirstName%></a><br>				</td>				<%else%>&nbsp;
				<%end if%>
				</td>				</tr>
			 				<%	rsReview.movenext
										if not rsReview.eof then						cLastName = rsReview("cLastName")
						cFirstName = rsReview("cFirstName")
						amLastName = trim(rsReview("amLastName"))
						amFirstName = rsReview("amFirstName")
						cEmpID = rsReview("cEmpID")						amEmpID = rsReview("amEmpID")						
						sEmpID = rsReview("sEmpID")
						sLastName = rsReview("sLastName")
						sFirstName = rsReview("sFirstName")						review = rsReview("next_review")						rvwdays = rsReview("reviewdays")
					else						prev_amLastName=""
					end if			   wend%>				   				<!--Insert header and employees for reviews in next 46-60 days-->				<tr>
					<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>					&nbsp;&nbsp;<b>46-60 Days</td>
					<td valign=top width=90 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>					&nbsp;&nbsp;<b>Date</b>					</td>
					<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>					&nbsp;&nbsp;<b>Sponsor</b>					</td>
				</tr>				<% while prev_amLastName = amLastName and rvwdays > 45 and rvwdays < 61%>
				<tr>
				<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;<a href='../../directory/content/details.asp?EmpID=<%=cEmpID%>'><%=cLastName%>,&nbsp;<%=cFirstName%></a><br>				</td>				<td valign=top width = 90 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;<%=review%></a><br>
				<td width = 150 valign=top colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;
				<%if not isnull(sLastName) then %>
				<a href='../../directory/content/details.asp?EmpID=<%=sEmpID%>'><%=sLastName%>,&nbsp;<%=sFirstName%></a><br>				</td>				<%else%>&nbsp;
				<%end if%>
				</td>				</tr>
			 				<%	rsReview.movenext
										if not rsReview.eof then						cLastName = rsReview("cLastName")
						cFirstName = rsReview("cFirstName")
						amLastName = trim(rsReview("amLastName"))
						amFirstName = rsReview("amFirstName")
						cEmpID = rsReview("cEmpID")						amEmpID = rsReview("amEmpID")						
						sEmpID = rsReview("sEmpID")
						sLastName = rsReview("sLastName")
						sFirstName = rsReview("sFirstName")						review = rsReview("next_review")						rvwdays = rsReview("reviewdays")
					else						prev_amLastName=""
					end if			   wend%>				   				<!--Insert header and employees for reviews in next 61-90 days-->
				<tr>
					<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>					&nbsp;&nbsp;<b>61-90 Days</td>
					<td valign=top width=90 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>					&nbsp;&nbsp;<b>Date</b>					</td>
					<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>					&nbsp;&nbsp;<b>Sponsor</b>					</td>
				</tr>				<% while prev_amLastName = amLastName and rvwdays > 60 and rvwdays < 91%>
				<tr>
				<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;<a href='../../directory/content/details.asp?EmpID=<%=cEmpID%>'><%=cLastName%>,&nbsp;<%=cFirstName%></a><br>				</td>				<td valign=top width = 90 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;<%=review%></a><br>
				<td width = 150 valign=top colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;
				<%if not isnull(sLastName) then %>
				<a href='../../directory/content/details.asp?EmpID=<%=sEmpID%>'><%=sLastName%>,&nbsp;<%=sFirstName%></a><br>				</td>				<%else%>&nbsp;
				<%end if%>
				</td>				</tr>
			 				<%	rsReview.movenext
										if not rsReview.eof then						cLastName = rsReview("cLastName")
						cFirstName = rsReview("cFirstName")
						amLastName = trim(rsReview("amLastName"))
						amFirstName = rsReview("amFirstName")
						cEmpID = rsReview("cEmpID")						amEmpID = rsReview("amEmpID")						
						sEmpID = rsReview("sEmpID")
						sLastName = rsReview("sLastName")
						sFirstName = rsReview("sFirstName")						review = rsReview("next_review")						rvwdays = rsReview("reviewdays")
					else						prev_amLastName=""
					end if			   wend%>				<%end if%>	
				
			</table>
		</span>		
	</td></tr>
	

<%wend
rsReview.close%>
</td>
</tr></table>
<% end if%>
</tr>
</table>
<!-- #include file="../../footer.asp" -->


<% else %>
<b><h2><font color=red>You are not currently an account manager,<br> therefore can not view the Account Manager Reports.
</b></h2></font>

<%end if%>
