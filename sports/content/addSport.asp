<%@ Language = Vbscript%>
<!--
Developer:	  Adam Lewandowski
Date:         10/24/2000
Description:  Tool for webmaster to add/remove a sport.
Modifications: 

-->

<!-- #include file="../section.asp" -->

<%
on error resume next
action = Request("Action")
removeSportID = Request("sport_ID")
newSportName = Request("newSportName")

if action="addSport" then
 sql="INSERT INTO Sports (Sport_Name) VALUES ('" & newSportName & "')"
 DBQuery(sql)
end if

if action="removeSport" then
 sql="DELETE FROM Sports WHERE Sport_ID=" & removeSportID
 'Response.Write("sql=" & sql & "<BR>")
 DBQuery(sql)
end if
%>


<HTML>

<HEAD>

</HEAD>



<%

	
%>
<body>

<table border=0><tr><td><font size=1 face="ms sans serif, arial, geneva" color=blue><b><STRONG>
  <font size=2 color=red><b>Add/Remove a Sport</b></font><p>
  </font></td></tr>
</table>

<form NAME="frmSport" ACTION="<%=Request.ServerVariables("SCRIPT_NAME")%>" method=post>
  <INPUT TYPE=Hidden NAME="Action" VALUE=""><br>
  <table border=0>
  <tr>
   <td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue>
    <STRONG>New Sport:</STRONG></font>
   </td>
   <td><INPUT TYPE=TEXT NAME=newSportName SIZE=30></td>
   <TD><INPUT TYPE=BUTTON VALUE="Add Sport" CLASS=button onClick="javascript:addSport()"></TD>
  </TR>
  <TR>
   <TD align=left><font face="ms sans serif, arial, geneva" size=1 color=blue>
    <STRONG>Sport to Remove:</STRONG></font>
   </td>
   <td>
    <Select NAME="sport_ID">
    <% 
     set rs = DBQuery("select * from Sports order by Sport_Name")
      do while not rs.eof %>
       <Option Value=<%=trim(rs("sport_ID"))%>><%=trim(left(rs("sport_Name"),45))%> 
       <% rs.movenext
      loop
      rs.close %>
	</Select>
	</td>
	<td><INPUT TYPE=BUTTON VALUE="Remove Sport" CLASS=button onClick="javascript:removeSport()"></TD>
  </tr>
  </table>
</form>
</body>

<SCRIPT LANGUAGE=javascript>
<!--
	function addSport() {
		document.frmSport.Action.value="addSport";
		document.frmSport.submit();
	}
	
	function removeSport() {
		document.frmSport.Action.value="removeSport";
		document.frmSport.submit();
	}
	
//-->
</SCRIPT>
</HTML>







<!-- #include file="../../footer.asp" -->