<%@ Language = Vbscript%>
<!--
Developer:	  Adam Lewandowski
Date:         10/24/2000
Description:  Tool for webmaster to assign/remove managers from a sport.
Modifications: 

-->

<!-- #include file="../section.asp" -->

<%
action = Request("Action")
sportID = Request("sport_ID")
newManagerID = Request("newManagerID")
removeManagerID = Request("removeManagerID")

if action = "addManager" then
 sql = "INSERT INTO Sports_Managers (sport_id, employee_id) VALUES (" & _
       sportID & "," & newManagerID & ")"
 DBQuery(sql)
end if

if action = "removeManager" then 
 sql = "DELETE FROM Sports_Managers WHERE sport_id = " & sportID & _
       " AND employee_id = " & removeManagerID
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
  <font size=2 color=red><b>Add/Remove Managers</b></font><p>
  </font></td></tr>
</table>

<form NAME="frmSport" ACTION="<%=Request.ServerVariables("SCRIPT_NAME")%>" method=post>
  <INPUT TYPE=Hidden NAME="Action" VALUE="<%=action%>">
  <INPUT TYPE=Hidden NAME="removeManagerID" VALUE="">
  <table border=0>
  <TR>
   <TD align=left><font face="ms sans serif, arial, geneva" size=1 color=blue>
    <STRONG>Sport:</STRONG></font>
   </td>
   <td>
    <Select NAME="sport_ID" onChange="loadForm()">
    <% 
     set rs = DBQuery("select * from Sports order by Sport_Name")
      do while not rs.eof %>
       <Option Value=<%=trim(rs("sport_ID"))%>
       <% if CInt(trim(rs("sport_ID"))) = CInt(sportID) then 
        Response.Write(" SELECTED ")
       end if %>>
       <%=trim(left(rs("sport_Name"),45))%> 
       <% rs.movenext
      loop
      rs.close %>
	</Select>
	</td>
  </tr>
  <% 
  if action <> "" then 
   %><TR><TD><B>Managers</B></TD></TR><%
   sql="SELECT e.employeeID, e.FirstName, e.LastName, s.sport_ID " & _ 
       "FROM Sports_Managers sm, Sports s, Employee e " & _
       "WHERE s.sport_id	= sm.sport_id AND sm.employee_id = e.employeeID " & _
       "AND sm.sport_id = " & sportID
   set rsManagers = DBQuery(sql)
   do while not rsManagers.eof
    %>
    <TR><TD><%=rsManagers("FirstName") & " " & rsManagers("LastName")%></TD>
     <TD><INPUT TYPE=BUTTON VALUE="Remove Manager" CLASS=button onClick="removeManager(<%=rsManagers("employeeID")%>)"></TD>
    </TR>
    <%
    rsManagers.movenext
   loop
   rsManagers.close
   %><TR><TD><SELECT NAME=newManagerID><%
   sql="SELECT employeeID, FirstName, LastName FROM Employee ORDER BY LastName, FirstName"
   set rsEmps = DBQuery(sql)
   do while not rsEmps.eof
    Response.Write("<OPTION VALUE='" & rsEmps("employeeID") & "'>" & rsEmps("LastName") & ", " & rsEmps("FirstName") & vbCRLF)
    rsEmps.movenext
   loop
   rsEmps.close
   %>
   </SELECT></TD><TD>
   <INPUT TYPE=BUTTON VALUE="Add Manager" CLASS=button onClick="addManager()">
   </TD></TR>
   <%
  end if 
  %>
  </table>
</form>
</body>

<SCRIPT LANGUAGE=javascript>
<!--
    function loadForm() {
		document.frmSport.Action.value="loadSport";
		document.frmSport.submit();
    }
	function addManager() {
		document.frmSport.Action.value="addManager";
		document.frmSport.submit();
	}
	
	function removeManager(ID) {
		document.frmSport.Action.value="removeManager";
		document.frmSport.removeManagerID.value=ID;
		document.frmSport.submit();
	}
	
//-->
</SCRIPT>
</HTML>







<!-- #include file="../../footer.asp" -->