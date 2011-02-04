<%@ Language = Vbscript%>
<!--
Developer:	  Adam Lewandowski
Date:         10/11/2000
Description:  Allows a SARK Exam to be added.
-->
<!-- #include file="../section.asp" -->

<HTML>

<%
techID = Request("techID")
Response.Write("<!-- techid: " & techID & "-->") 
ls_action=Request("action")
ls_techID=Request("techID")
ls_description=Request("description")
ls_areaID=Request("areaID")
		
if Request.Form("Submitted")="True" then
 ' Save stuff to the database
 sql = "UPDATE Tech SET Tech_Desc='" & ls_description & "', Tech_Area_ID=" & ls_areaID & " WHERE Tech_ID="&techID
 'Response.Write("sql="&sql&"<BR>")
 DBQuery(sql)
elseif Request.Form("deleted")= "True" then
 ' Remove stuff from the database
 ctName=Request.Form("exam_Name")
 ls_examID=Request.Form("exam_ID")
 ls_examnumber=Request.Form("exam_number")
 ls_cert=Request.Form("cert")
elseif Request.form("Selected")="True" then
 ' Load stuff from the database
 set rs=dbquery("Select * from Tech where tech_ID=" & techID)
 ls_techID=trim(rs("tech_id"))
 ls_TechName=trim(rs("tech_Name"))
 ls_description=trim(rs("tech_desc"))
 ls_areaID=trim(rs("tech_area_ID"))
 rs.close
end if
	
%>
<body>

<table border=0><tr><td><font size=1 face="ms sans serif, arial, geneva" color=blue><b><STRONG>
  <font size=2 color=red><b>Edit Technology</b></font><p>
  </font></td></tr></table>
<form NAME="frmInfo" ACTION="<%Request.ServerVariables("SCRIPT_NAME")%>" method=post>
  <!-- <INPUT TYPE=Hidden NAME="techID" VALUE="<%=techID%>"> -->
  <INPUT TYPE=Hidden NAME="Submitted" VALUE="False">
  <INPUT TYPE=Hidden NAME="Selected" VALUE="False">
  <INPUT TYPE=Hidden NAME="deleted" VALUE="False">

  <table border=0>
  <tr>
	<td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue>
	<STRONG>Technology:</STRONG></font></td>
	<td></td>
	<td>
		<Select NAME="techID" onChange='javascript:LoadForm();'>
			<%	set rs = DBQuery("select * from Tech order by Tech_name")
				if not rs.eof and not rs.bof then
					Response.Write("<option value=" & chr(34) & chr(34) & "></option>")
					do while not rs.eof%>
							<Option Value=<%=trim(rs("tech_ID"))%>
							<% 
							if trim(rs("tech_ID")) = techID then
							  Response.Write(" SELECTED ")
							end if
							%>
							>
							<%=trim(left(rs("tech_Name"),45))%> 
						<% rs.movenext
					loop
					rs.close
				end if%>
		</Select>
	</td>
  </tr>
  <tr><td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue>
   <STRONG>Description:</STRONG></font></td>
	<td></td>
	<td><textarea NAME='description' ROWS=5 COLS=45><%=ls_description%></textarea>
	</td></tr>
  <tr><td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue>
   <STRONG>Area:</STRONG></font></td>    
   <td></td>
   <td>
   <SELECT NAME="areaID">
    <OPTION VALUE="">
    <%
     set rs = DBQuery("SELECT Tech_Area_ID, Tech_Area FROM Tech_Area")
     do while not rs.eof
      Response.Write("<OPTION VALUE='" & rs("Tech_Area_ID") & "' ")
      if trim(rs("Tech_Area_ID")) = ls_areaID then
       Response.Write("SELECTED ")
      end if
      Response.Write(">" & rs("Tech_Area"))
      rs.movenext
     loop
     rs.close
    %>    
    </SELECT>
   </td>
  </tr>	
	
</table>

<table align=center width=75% border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Update" OnClick='Changes();' id=button1 name=button1>
			<input type=button class=button value='Delete' OnClick='deleteexam();' id=button2 name=button2>
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
	if (confirm("Are you sure you want to delete this technology?"))
	{
		document.frmInfo.deleted.value="True"
				document.frmInfo.submit()
	}
}
	
//-->
</SCRIPT>
</HTML>







<!-- #include file="../../footer.asp" -->