<%@ Language = Vbscript%>
<!--
Developer:	  David Martin
Date:         09/20/2000
Description:  Allows a SARK Class to be added.
-->

<!-- #include file="../section.asp" -->

<%

classid = Request.QueryString("ls_classid")
Response.Write("<!-- classid: " & classid & "-->") 
%>


<HTML>

<HEAD>

</HEAD>



<%
		if Request.Form("Submitted")="True" then
		className=Request.Form("class_Name")
		ls_classID=Request.Form("class_ID")
		
		
		if className<> "" then
			
			queryString = "UPDATE class_list SET class_Name = '" & className & "' WHERE class_ID = " & ls_classID
		
			DBQuery(queryString)
		end if
		
        
		elseif Request.Form("deleted")="True" then
		className=Request.Form("class_Name")
		ls_classID=Request.Form("class_ID")
		
		
		if className<> "" then
			
			queryString = "DELETE from class_list WHERE class_id = " & ls_classid  
		
			DBQuery(queryString)
		end if





	elseif Request.form("Selected")="True" then
		set rs=dbquery("Select class_ID, class_Name from class_list " & _
		               "where class_ID='" & Request.Form("class_ID") & "'")
		ls_className=trim(rs("class_Name"))
		ls_classID=trim(rs("class_ID"))
		
		rs.close
	end if
	
%>
<body>

<table border=0><tr><td><font size=1 face="ms sans serif, arial, geneva" color=blue><b><STRONG>
  <font size=2 color=red><b>EDIT CLASS</b></font><p>
  <!-- This form is used to change a Cert's name or website URL. -->
  </font></td></tr></table>

<form NAME="frmInfo" ACTION="edit_class.asp" method=post>
  <INPUT TYPE=Hidden NAME="InsertUpdate" VALUE="<%=ls_classID%>"><br>
  <INPUT TYPE=Hidden NAME="Submitted" VALUE="False">
  <INPUT TYPE=Hidden NAME="Selected" VALUE="False">
  <INPUT TYPE=Hidden NAME="deleted" VALUE="False">

  <table border=0>
  <tr>
	<td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Class Name:</STRONG></font></td>
	<td></td>
	<td>
		<Select NAME="class_ID" onChange='javascript:LoadForm();'>
			<%	set rs = DBQuery("select * from class_list order by class_name")
				if not rs.eof and not rs.bof then
					Response.Write("<option value=" & chr(34) & chr(34) & "></option>")
					do while not rs.eof%>
							<Option Value=<%=trim(rs("class_ID"))%>><%=left(trim(rs("class_Name")),55)%> 
						<% rs.movenext
					loop
					rs.close
					if ls_className<>"" and Request.form("Submitted")="False" then%>
						<Option selected value=<%=ls_classID%>><%=left(ls_classname, 55)%>
						
					
						
					<%end if
					
				
					end if%>
		</Select>
	</td>
  </tr>

  <tr><td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Class Name:</STRONG></font></td>
	<td></td>
	<%if ls_className="" and Request.form("Submitted")="False" then
		Response.Write "<td><INPUT TYPE='Text' NAME='class_Name' VALUE='' size=55 maxlength=255></td></tr>"
	else
		Response.Write ("<td><INPUT TYPE='Text' NAME='class_Name' VALUE=" & chr(34) & left(ls_classname, 55) & chr(34) & " size=55 maxlength=255></td></tr><br>")
	end if%>
		
		 
	
	
</table>

<table align=center width=75% border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Update" OnClick='Changes();' id=button1 name=button1>
			<input type=button class=button value='Delete' OnClick='deleteexam();' id=button2 name=button2>
			<input type=button class=button value="Cancel" OnClick='GoHome();' id=button3 name=button3>
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
<SCRIPT LANGUAGE=javascript>
<!--
	function Changes()
	{
				document.frmInfo.Submitted.value="True"
				document.frmInfo.submit()
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
	
	function deleteexam()
{
	if (confirm("Are you sure you want to delete this class?"))
	{
		document.frmInfo.deleted.value="True"
				document.frmInfo.submit()
	}
}
	
//-->
</SCRIPT>
</HTML>







<!-- #include file="../../footer.asp" -->