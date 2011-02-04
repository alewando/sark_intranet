<%@ Language = Vbscript%>
<!--
Developer:	  MAPGAR
Date:         03/27/2000
Description:  Allows a SARK client to be added.
-->

<!-- #include file="../section.asp" -->

<%
	'dim SESS, Edits
	if Request.Form("Submitted")="True" then
		clName=Request.Form("Client_Name")
		clWeb=Request.Form("Client_Website")
		ls_ClientID=Request.Form("Client_ID")
		ls_acct_mgr_id=trim(Request.Form("txtAccountManagerID"))
		ls_sales_rep_id=trim(Request.Form("txtSalesRepID"))
		
		'set SESS = server.CreateObject("Client.Session")
		'sess.dbConnect Application("DataConn_ConnectionString")
		'set edits=sess.EDIT
		
		if clname<>"" then
			'CHG=ls_ClientID & clName & clWeb
			'clName = Replace(clName, "'", "''")
			'CHG=edits.ChgClient(ls_ClientID, clName, clweb)
			queryString = "UPDATE Client SET ClientName = '" & _
			              clName & "', WebSite = '" & clWeb & "', Employee_AcctMgr_ID = " & _
			              ls_acct_mgr_id & ", Employee_SalesRep_ID = " & ls_sales_rep_id & _
			              " WHERE Client_ID = " & ls_ClientID
			DBQuery(queryString)
		end if

	elseif Request.form("Selected")="True" then
		set rs=dbquery("Select Client_ID, ClientName, Website, Employee_AcctMgr_ID, " & _
		               "Employee_SalesRep_ID from client where Client_ID='" & Request.Form("Client_ID") & "'")
		ls_CName=rs("ClientName")
		ls_CWeb=rs("Website")
		ls_ClientID=rs("Client_ID")
		ls_acct_mgr_id=trim(rs("Employee_AcctMgr_ID"))
		ls_sales_rep_id=trim(rs("Employee_SalesRep_ID"))
		rs.close
	end if
%>
<SCRIPT LANGUAGE=javascript>
<!--
	function Changes()
		{
			if (document.frmInfo.Client_Website.value!="") {
				document.frmInfo.Submitted.value="True"
				document.frmInfo.submit()
				}
			else {
				alert("Enter a value for the Client's WebSite")
			}
		}
	function GoHome()
	{
		window.navigate("default.asp")
	}
	function LoadForm()
	{
		document.frmInfo.Selected.value="True"
		document.frmInfo.submit()
	}
//-->
</SCRIPT>

<HTML>
<HEAD>
</HEAD>
<body>

<table border=0><tr><td><font size=1 face="ms sans serif, arial, geneva" color=blue><b><STRONG>
  <font size=2 color=red><b>EDIT CLIENT</b></font><p>
  <!-- This form is used to change a client's name or website URL. -->
  </font></td></tr></table>

<form NAME="frmInfo" ACTION="edit_client.asp" method=post>
  <INPUT TYPE=Hidden NAME="InsertUpdate" VALUE="<%=ls_ClientID%>"><br>
  <INPUT TYPE=Hidden NAME="Submitted" VALUE="False">
  <INPUT TYPE=Hidden NAME="Selected" VALUE="False">

  <table border=0>
  <tr>
	<td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Client Name:</STRONG></font></td>
	<td></td>
	<td>
		<Select NAME="Client_ID" onChange='javascript:LoadForm();'>
			<%	set rs = DBQuery("select Client_ID, ClientName from client order by ClientName")
				if not rs.eof and not rs.bof then
					Response.Write("<option value=" & chr(34) & chr(34) & "></option>")
					do while not rs.eof%>
							<Option Value=<%=trim(rs("Client_ID"))%>><%=trim(rs("ClientName"))%>
						<% rs.movenext
					loop
					rs.close
					if ls_CName<>"" and Request.form("Submitted")="False" then%>
						<Option selected value=<%=ls_ClientID%>><%=ls_CName%>
					<%end if
				end if%>
		</Select>
	</td>
  </tr>

  <tr><td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Client Name:</STRONG></font></td>
	<td></td>
	<%if ls_CName="" and Request.form("Submitted")="False" then
		Response.Write "<td><INPUT TYPE='Text' NAME='Client_Name' VALUE='' size=40 maxlength=255></td></tr>"
	else
		Response.Write ("<td><INPUT TYPE='Text' NAME='Client_Name' VALUE=" & chr(34) & ls_CName & chr(34) & " size=40 maxlength=255></td></tr>")
	end if%>
	<!--<td><INPUT TYPE="Text" NAME="Client_Name" VALUE="<%=ls_CName%>" size=40 maxlength=255></td></tr>-->
	
  <tr><td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Client Website:</STRONG></font></td>
	<td align=right><font face="ms sans serif, arial, geneva" size=1 color=black><STRONG>http://</STRONG></font></td>
	<%if ls_CWeb="" and Request.form("Submitted")="False" then
		Response.Write "<td><INPUT TYPE='Text' NAME='Client_Website' VALUE='' size=40 maxlength=255></td></tr>"
	else
		Response.Write ("<td><INPUT TYPE='Text' NAME='Client_Website' VALUE='" & ls_CWeb & "' size=40 maxlength=255></td></tr>")
	end if%>
	<!--<td><INPUT TYPE="Text" NAME="Client_Website" VALUE="<%=ls_CWeb%>" size=40 maxlength=255></td></tr>-->
	
<tr>
	<td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Account Manager:</STRONG></font></td>
	<td></td>
	<td>
		<select name="txtAccountManagerID">
			<%
			  'Note: EmployeeTitle_ID of 5, 6, or 16 corresponds to an account manager position
			  set rs = DBQuery("select employeeID, firstname, lastname from employee " & _
			                   "where EmployeeTitle_ID in (5, 6, 16) order by lastname")
			  
			  if not rs.eof and not rs.bof then
					Response.Write("<option value=" & chr(34) & chr(34) & "></option>")
					do while not rs.eof
					
					  s_empID = trim(rs("employeeID"))
					  s_fn = trim(rs("firstname"))
					  s_ln = trim(rs("lastname"))
					%>
					  <Option Value=<%=s_empID%>><%=s_fn & " " & s_ln%></option>
					<% rs.movenext
					loop
					rs.close
					
					if ls_CName<>"" and Request.form("Submitted")="False" then%>
					  <Option selected value=
					    <%if not isNull(ls_acct_mgr_id) then
					        Response.Write(ls_acct_mgr_id)
					      else
					        Response.Write(" ")
					      end if
					    %>>
						<%
						  if not isNull(ls_acct_mgr_id) then
						    set rs = DBQuery("select firstname, lastname from employee " & _
						                   "where employeeID = " & ls_acct_mgr_id)
						    Response.Write(trim(rs("firstname")) & " " & trim(rs("lastname")))
						    rs.close
						  end if
						%>
					  </option>		
			<%
				    end if
			   end if
			%>
		</select>
	</td>
	
</tr>
<tr>
	<td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Sales Representative:</STRONG></font></td>
	<td></td>
	<td>
		<select name="txtSalesRepID">
			<%
			  'Note: EmployeeTitle_ID of 9, 10, 15, or 17 corresponds to a sales rep position
			  set rs = DBQuery("select employeeID, firstname, lastname from employee " & _
			                   "where EmployeeTitle_ID in (9, 10, 15, 17) order by lastname")
			  
			  if not rs.eof and not rs.bof then
					Response.Write("<option value=" & chr(34) & chr(34) & "></option>")
					do while not rs.eof%>
							<Option Value=<%=trim(rs("employeeID"))%>><%=trim(rs("firstname")) & " " & trim(rs("lastname"))%></option>
						<% rs.movenext
					loop
					rs.close
					
					if ls_CName<>"" and Request.form("Submitted")="False" then%>
					  <Option selected value=
					    <%if not isNull(ls_sales_rep_id) then
					        Response.Write(ls_sales_rep_id)
					      else
					        Response.Write(" ")
					      end if
					    %>>
						<%
						  if not isNull(ls_sales_rep_id) then
						    set rs = DBQuery("select firstname, lastname from employee " & _
						                   "where employeeID = " & ls_sales_rep_id)
						    Response.Write(trim(rs("firstname")) & " " & trim(rs("lastname")))
						    rs.close
						  end if
						%>
					  </option>		
			<%
				    end if
			   end if
			%>
		</select>
	</td>
</tr>
</table>

<table align=center width=75% border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Update" OnClick='Changes();'>
			<input type=button class=button value="Cancel" OnClick='GoHome();'>
		</td>
	</tr>
	<tr>
		<td align=center>
			<%	if mid(chg,1,6)="Update" then
					Response.Write ("<b><Font face='ms sans serif, arial, geneva' size=1 color=red><Strong>Database Updated!</Strong></font></b>")
				end if%>
		</td>
	</tr>
</table>

</form>
</body>
</HTML>
<!-- #include file="../../footer.asp" -->