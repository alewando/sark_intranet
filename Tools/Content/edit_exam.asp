<%@ Language = Vbscript%>
<!--
Developer:	  David Martin
Date:         08/31/2000
Description:  Allows a SARK Exam to be added.
Modifications: Adam Lewandowski - 10/02/2000
               Added Exam Category field

-->

<!-- #include file="../section.asp" -->

<%

examid = Request.QueryString("ls_examid")
Response.Write("<!-- examid: " & examid & "-->") 
%>


<HTML>

<HEAD>

</HEAD>



<%
		if Request.Form("Submitted")="True" then
		   ctName=Request.Form("exam_Name")
		   ls_examID=Request.Form("exam_ID")
		   ls_examnumber=Request.Form("exam_number")
		   ls_cert=Request.Form("cert")
		   ls_examcategory=Request.Form("exam_category")
		
		
		if ctName<> "" then
			
			'queryString = "UPDATE exam_list SET exam_Name = '" & ctName & "', exam_number = '" & ls_examnumber & "', cert = '" & ls_cert &"' WHERE exam_ID = " & ls_examID
		
		     queryString = "UPDATE exam_list SET exam_Name = '" & ctName & "', exam_number = '" & ls_examnumber & "', Category_ID='" & ls_examcategory & "', cert = '" & replace(ls_cert,"True","1") &"' WHERE exam_ID = " & ls_examID
			'Response.Write("SQL=" & queryString & "<BR>")
			DBQuery(queryString)
		end if
		
        
		elseif Request.Form("deleted")="True" then
		       ctName=Request.Form("exam_Name")
		       ls_examID=Request.Form("exam_ID")
		       ls_examnumber=Request.Form("exam_number")
		       ls_cert=Request.Form("cert")
		
		if ctName<> "" then
			
			queryString = "DELETE from Exam_list WHERE exam_id = " & ls_examid  
		
			DBQuery(queryString)
		end if





	elseif Request.form("Selected")="True" then
		set rs=dbquery("Select exam_ID, exam_Name, exam_number, Category_ID, cert from exam_list " & _
		               "where exam_ID='" & Request.Form("exam_ID") & "'")
		ls_CName=trim(rs("exam_Name"))
		ls_examID=trim(rs("exam_ID"))
		ls_examnumber=trim(rs("exam_number"))
		ls_examcategory=trim(rs("Category_ID"))
		ls_cert=trim(rs("cert"))
		rs.close
	end if
	
%>
<body>

<table border=0><tr><td><font size=1 face="ms sans serif, arial, geneva" color=blue><b><STRONG>
  <font size=2 color=red><b>EDIT Exams</b></font><p>
  <!-- This form is used to change a Cert's name or website URL. -->
  </font></td></tr></table>

<form NAME="frmInfo" ACTION="edit_exam.asp" method=post>
  <INPUT TYPE=Hidden NAME="InsertUpdate" VALUE="<%=ls_examID%>"><br>
  <INPUT TYPE=Hidden NAME="Submitted" VALUE="False">
  <INPUT TYPE=Hidden NAME="Selected" VALUE="False">
  <INPUT TYPE=Hidden NAME="deleted" VALUE="False">

  <table border=0>
  <tr>
	<td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Exam Name:</STRONG></font></td>
	<td></td>
	<td>
		<Select NAME="exam_ID" onChange='javascript:LoadForm();'>
			<%	set rs = DBQuery("select * from exam_list order by exam_name")
				if not rs.eof and not rs.bof then
					Response.Write("<option value=" & chr(34) & chr(34) & "></option>")
					do while not rs.eof%>
							<Option Value=<%=trim(rs("exam_ID"))%>><%=trim(left(rs("exam_Name"),45))%> 
						<% rs.movenext
					loop
					rs.close
					if ls_CName<>"" and Request.form("Submitted")="False" then%>
						<Option selected value=<%=ls_examID%>><%=trim(left(ls_CName,45))%>
						
					
						
					<%end if
					
				
					end if%>
		</Select>
	</td>
  </tr>

  <tr><td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Exam Name:</STRONG></font></td>
	<td></td>
	<%if ls_CName="" and Request.form("Submitted")="False" then
		Response.Write "<td><INPUT TYPE='Text' NAME='exam_Name' VALUE=" & chr(34) & (ls_CName) & chr(34)  & " size=48 maxlength=255></td></tr>"
	else
		Response.Write ("<td><INPUT TYPE='Text' NAME='exam_Name' VALUE=" & chr(34) & (ls_CName) & chr(34) & " size=48 maxlength=255></td></tr><br>")
	end if%>
		<tr><td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Exam Number:</STRONG></font></td>
	<td></td>
	<%if Request.form("Submitted")="True" or Request.form("deleted")="True" then
		Response.Write "<td><INPUT TYPE='Text' NAME='exam_number' VALUE='' size=10 maxlength=10></td></tr>"
	else
		Response.Write ("<td><INPUT TYPE='Text' NAME='exam_number' VALUE=" & chr(34) & ls_examnumber & chr(34) & " size=10 maxlength=10></td></tr>")
	end if%>
  <tr><td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Exam Category:</STRONG></font></td>    
   <td></td>
   <td>
   <SELECT NAME="exam_category">
    <OPTION VALUE="">
    <%
     set rs = DBQuery("SELECT Category_ID, Category_Description FROM Exam_Categories")
     do while not rs.eof
      'Response.Write("<!-- rs(Category_ID)='" & rs("Category_ID") & "', ls_examcategory='" & ls_examcategory & "' -->")
      Response.Write("<OPTION VALUE='" & rs("Category_ID") & "' ")
      if Request.Form("Submitted")<>"True" and Request.Form("deleted")<>"True" and trim(rs("Category_ID")) = trim(ls_examcategory) then
       Response.Write("SELECTED ")
      end if
      Response.Write(">" & rs("Category_Description"))
      rs.movenext
     loop
     rs.close
    %>    
    </SELECT>
   </td>
  </tr>	
		 <tr><td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Certification?:</STRONG></font></td>
	<td></td>
	
		<%if Request.form("Submitted")="True" or Request.form("deleted")="True" then
		Response.Write "<td><INPUT TYPE=checkbox NAME='cert' VALUE='True' ></td></tr>"
	else
			Response.Write "<td><INPUT TYPE=checkbox NAME='cert' VALUE='True'"

	if ls_cert="True" then 
		Response.Write( " CHECKED ></td></tr>") 
	else 
		Response.Write( " ></td></tr>") 

	end if
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
	if (confirm("Are you sure you want to delete this exam?"))
	{
		document.frmInfo.deleted.value="True"
				document.frmInfo.submit()
	}
}
	
//-->
</SCRIPT>
</HTML>







<!-- #include file="../../footer.asp" -->