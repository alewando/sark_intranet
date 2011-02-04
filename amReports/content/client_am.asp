<!--
Developer:	  SSeissiger
Date:         08/17/2000
Description:  Report - client list by account manager with sales rep also listed
			  Part of the AM report set
-->
<% @language=VBscript%>

<%if (hasRole("AccountManager")) then %>
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
<b>Client List with Sales Rep -- By Account Manager</b>
<p>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>

<TABLE BORDER=0 WIDTH=450>
<tr><td ALIGN=LEFT VALIGN=top>
<!--''''''''''''''''''''''''''''''''''''''''''''''
'Code to display a list of Account Managers 
'List clients and sales rep under each A/M
''''''''''''''''''''''''''''''''''''''''''''''-->

<%
sql = "select c.clientname,sr.firstname as srFirstName, sr.lastname as srLastName, sr.employeeID as srEmpID, " &_
"am.firstname as amFirstName, am.lastname as amLastName " &_
"from client c left join employee am on " &_
"c.Employee_AcctMgr_ID = am.employeeID " &_
"join employee sr on " &_
"c.Employee_SalesRep_ID = sr.employeeID " &_
"ORDER BY am.lastname, c.clientname"

set rs = DBQuery(sql)

prev_amLastName = ""
%>
<table width=400 cellspacing=0 cellpadding=4 border=0 >
<tr><td colspan=2 height=1>
</td></tr>

	<%while not rs.eof
		srLastName = rs("srLastName")
		srFirstName = rs("srFirstName")
		srEmpID = rs("srEmpID")
		amLastName = rs("amLastName")
		amFirstName = rs("amFirstName")
		client = rs("clientname")
		
		if prev_amLastName = "" OR prev_amLastName <> amLastName then%>
		<tr>
			<td colspan=1 width=200 align=left><font size=1 face= "ms sans serif, arial, geneva"  color = maroon>
			<b><span onClick="toggle('<%=amLastName%>')" alt="Expand or collapse this section." style="cursor:hand"><%=amLastName%>,&nbsp;<%=amFirstName%></b>
			<span id='<%=amLastName &  "Header"%>' style="">&nbsp;&nbsp;(collapsed)</span>
			</td>
		</tr>
		<%end if%>
				<%prev_amLastName = amLastName%>			<tr><td colspan=2>
			<span id='<%=amLastName%>' style="display:none">		<table>			<% if prev_amLastName = amLastName then %>				<tr>
					<td valign=top width=200 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=Maroon>					&nbsp;&nbsp;</td>
					<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=maroon>					&nbsp;&nbsp;<b>Sales Rep</b>					</td>
				</tr>			<%end if%>					<% while prev_amLastName = amLastName%>
				<tr>
					<td valign=top width = 200 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>					&nbsp;&nbsp;<%=client%></a><br>					</td>
					<td valign=top width=150 colspan=1 align=left><font size=1 face= "ms sans serif, arial, geneva" color=black>					&nbsp;&nbsp;<a href='../../directory/content/details.asp?EmpID=<%=srEmpID%>'><%=srLastName%>,&nbsp;<%=srFirstName%></a><br>					</td>				</tr>
			 				<%	rs.movenext
														if not rs.eof then					srLastName = rs("srLastName")
					srFirstName = rs("srFirstName")
					srEmpID = rs("srEmpID")
					amLastName = rs("amLastName")
					amFirstName = rs("amFirstName")
					client = rs("clientname")
				else					prev_amLastName=""
				end if				wend%>	
						
			</table>
		</span>	</td></tr>
	

<%wend
rs.close%>
</td>
</tr></table>
</tr>
</table>
		
<!-- #include file="../../footer.asp" -->

<% else %>
	<b><h2><font color=red>You are not currently an account manager,<br> therefore can not view the Account Manager Reports.
	</b></h2></font>

<%end if%>
