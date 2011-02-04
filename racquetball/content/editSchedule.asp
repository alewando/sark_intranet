<%@ Language = Vbscript%>
<!--
Developer:    Adam Lewandowski
Date:         11/21/2000
Description:  Edit racquetball leagues & schedules
-->
<!-- #include file="../section.asp" -->

<%
 action      = Request("action")
 actionValue = Request("actionValue")
 leagueID    = Request("leagueID")
 leagueName  = Request("leagueName")
 newGameWeek = Request("newGameWeek")
 newGamePlayerA = Request("newGamePlayerA")
 newGamePlayerB = Request("newGamePlayerB")
 
 Response.Write("<!-- Form Variables: " & vbCRLF)
 Response.Write("action        = " & action & vbCRLF)
 Response.Write("actionValue   = " & actionValue & vbCRLF)
 Response.Write("leagueID      = " & leagueID & vbCRLF)
 Response.Write("leagueName    = " & leagueName & vbCRLF)
 Response.Write("newGameWeek   = " & newGameWeek & vbCRLF)
 Response.Write("newGamePlayerA= " & newGamePlayerA & vbCRLF)
 Response.Write("newGamePlayerB= " & newGamePlayerB & vbCRLF)
 Response.Write("-->" & vbCRLF)
 
 ' Create a new league
 if action="createLeague" then
  sql = "INSERT INTO Racquetball_League (Description) VALUES (" & _
        "'" & clean(leagueName) & "')"
  DBQuery(sql)
 end if
 
 'Remove a league
 if action="deleteLeague" then
  sql = "DELETE FROM Racquetball_Schedule WHERE league_ID=" & leagueID
  DBQuery(sql)
  sql = "DELETE FROM Racquetball_Players WHERE league_ID=" & leagueID
  DBQuery(sql)
  sql = "DELETE FROM Racquetball_League WHERE league_ID=" & leagueID
  DBQuery(sql)
  leagueID="-1"
 end if
 
 'Add a player to the league
 if action="addPlayer" then
  sql = "INSERT INTO Racquetball_Players (league_ID, employee_ID) VALUES (" & _
        leagueID & "," & actionValue & ")"
  DBQuery(sql)
 end if
 
 'Remove a player from the league
 if action="removePlayer" then
  sql = "DELETE FROM Racquetball_Players WHERE League_ID=" & leagueID & _
        " AND employee_ID=" & actionValue
  DBQuery(sql)
 end if
 
 'Add game
 if action="addGame" then
  sql = "INSERT INTO Racquetball_Schedule (league_id, week, playerA, playerB) VALUES (" & _
        leagueID & ",'" & newGameWeek & "'," & newGamePlayerA & "," & newGamePlayerB & ")"
  DBQuery(sql)
 end if
 
 'Remove Game
 if action="removeGame" then
  sql = "DELETE FROM Racquetball_Schedule WHERE league_ID=" & leagueID & _
        " AND match_ID = " & actionValue
  DBQuery(sql)
 end if

 'Load league information
 if leagueID<>"" and leagueID<>"-1" then
  sql = "SELECT * from Racquetball_League WHERE league_ID=" & leagueID
  set rs = DBQuery(sql)
  leagueName = rs("description")
 end if
 
%>

<HTML><HEAD>
<TITLE>Sark Racquetball - Edit Leagues & Schedules</TITLE>
<LINK rel="stylesheet" type="text/css" href="..\style.css"></HEAD>
<BODY>

<FORM NAME="frmSched" ACTION="<%=Request.ServerVariables("SCRIPT_NAME")%>" METHOD=POST>
<INPUT TYPE=HIDDEN NAME="action" VALUE="">
<INPUT TYPE=HIDDEN NAME="actionValue" VALUE="">
<!-- 
 <INPUT TYPE=HIDDEN NAME="leagueID" VALUE="<%=leagueID%>">
-->

<DIV ALIGN=CENTER CLASS=leagueHeader>
Racquetball Leagues
</DIV>
<HR>
<TABLE BORDER=0>
 <TR>
  <TD>Select league</TD>
  <TD>
   <SELECT NAME="leagueID" onChange="loadLeague()">
   <%
    Response.Write("<OPTION VALUE='-1'")
    if leagueID="" then
     Response.Write(" SELECTED ")
    end if
    Response.Write("> </OPTION>")
    sql = "SELECT * FROM Racquetball_League ORDER BY league_ID"
    set rs = DBQuery(sql)
    do while not rs.eof 
     Response.Write("<OPTION VALUE='" & rs("league_ID") & "' ")
     if CInt(rs("league_ID"))=CInt(leagueID) then
      Response.Write(" SELECTED ")
     end if 
     Response.Write(">" & rs("description") & "</OPTION>" & vbCRLF)
     rs.moveNext
    loop
   %>
   </SELECT>
  </TD>
 <% if leagueID<>"" then %>
  <TD>
   <INPUT TYPE=BUTTON CLASS=button VALUE="Delete League" onClick="deleteLeague()">
  </TD>
 <% end if %>
 </TR>
 <TR>
  <TD>League Name:</TD>
  <TD><INPUT TYPE=TEXT NAME="leagueName"></TD>
  <TD><INPUT TYPE=BUTTON CLASS=button VALUE="Create League" onClick="createLeague()"></TD> 
 </TR>
</TABLE>
<HR>
<% if leagueID<>"" and leagueID<>"-1" then %>
<DIV CLASS=standingsHeader>
Players
</DIV>
<TABLE BORDER=0>
<%
 sql = "SELECT employee_ID, lastName, firstName FROM Racquetball_Players rp, Employee e " & _
       "WHERE rp.employee_ID = e.employeeID " & _
       "AND league_ID=" & leagueID & " " & _
       "ORDER BY lastName, firstName"
 set rs = DBQuery(sql)
 do while not rs.eof
  %>
   <TR>
    <TD><%=rs("lastName")%>, <%=rs("firstName")%></TD>
    <TD><INPUT TYPE=Button CLASS=button VALUE="Remove Player" onClick="removePlayer(<%=rs("employee_id")%>)"></TD>
   </TR>
  <%
  rs.moveNext
 loop
%>
<TR>
 <TD>
 <SELECT NAME="newPlayer">
  <%
  sql = "SELECT firstName, lastName, employeeID " & _
        "FROM Employee " & _
        "WHERE EmployeeID NOT IN " & _
        "(SELECT employee_ID FROM Racquetball_Players WHERE League_ID=" & leagueID & ") " & _
        "ORDER BY lastName, firstName"
  set rs=DBQuery(sql)
  do while not rs.eof
   Response.Write("<OPTION VALUE='" & rs("employeeID") & "'>")
   Response.Write(rs("lastName") & ", " & rs("firstName") & "</OPTION>")
   rs.moveNext
  loop
  %>
 </SELECT>
 </TD>
 <TD>
  <INPUT TYPE=Button CLASS=button VALUE="Add Player" onClick="addPlayer()">
 </TD>
</TR>
</TABLE>
<HR>
<DIV CLASS=standingsHeader>
Schedule
</DIV>
<TABLE BORDER=0>
 <TR>
  <TH>Week</TH>
  <TH>Player A</TH>
  <TH>Player B</TH>
 </TR>
 <%
  sql = "SELECT match_ID, week, e1.firstName+' '+e1.lastName as nameA, e2.firstName+' '+e2.lastName as nameB " & _
        "FROM Racquetball_Schedule rs, Employee e1, Employee e2 " & _
        "WHERE rs.playerA = e1.employeeID AND rs.playerB = e2.employeeID " & _
        " AND League_ID=" & leagueID & _
        " ORDER BY Week, playerA"
  set rs = DBQuery(sql)
  do while not rs.eof
   %>
   <TR>
    <TD CLASS=weekHeader><%=rs("week")%></TD>
    <TD CLASS=playerLoser><%=rs("nameA")%></TD>
    <TD CLASS=playerLoser><%=rs("nameB")%></TD>
    <TD><INPUT TYPE=BUTTON CLASS=button VALUE="Remove Game" onClick="removeGame(<%=rs("match_ID")%>)"></TD>
   </TR>
   <%
   rs.moveNext
  loop
 %>
 <TR>
  <TD>
   <INPUT TYPE=TEXT NAME="newGameWeek" VALUE="<%=newGameWeek%>" SIZE=10><BR>
   <FONT SIZE="-4">Ex: 12/31/2000</FONT>
  </TD>
  <TD>
   <SELECT NAME=newGamePlayerA>
   <%
    sql = "SELECT firstName, lastName, employee_ID FROM Racquetball_Players rp, Employee e " & _
          "WHERE rp.employee_ID = e.employeeID AND league_ID=" & leagueID & _
          " ORDER BY lastName, firstName"
    set rs=DBQuery(sql)
    do while not rs.eof
     %>
     <OPTION VALUE="<%=rs("employee_ID")%>"><%=rs("LastName")%>, <%=rs("FirstName")%></OPTION>
     <%
     rs.moveNext
    loop
   %>
   </SELECT>
  </TD>

  <TD>
   <SELECT NAME=newGamePlayerB>
   <%
    sql = "SELECT firstName, lastName, employee_ID FROM Racquetball_Players rp, Employee e " & _
          "WHERE rp.employee_ID = e.employeeID AND league_ID=" & leagueID & _
          " ORDER BY lastName, firstName"
    set rs=DBQuery(sql)
    do while not rs.eof
     %>
     <OPTION VALUE="<%=rs("employee_ID")%>"><%=rs("LastName")%>, <%=rs("FirstName")%></OPTION>
     <%
     rs.moveNext
    loop
   %>
   </SELECT>
  </TD>
  <TD>
   <INPUT TYPE=BUTTON CLASS=button VALUE="Add Game" onClick="addGame()">
  </TD>
 </TR>
</TABLE>
<% end if 'end if leagueID<>"" %>
</FORM>

<script language="JavaScript">
function loadLeague() {
 if(document.frmSched.leagueID.value == -1) {
  return;
 }
 document.frmSched.action.value="loadLeague";
 document.frmSched.actionValue.value=document.frmSched.leagueID.value;
 document.frmSched.submit();
}

function deleteLeague() {
 if(document.frmSched.leagueID.value == -1) {
  return;
 }
 if(confirm("Are you sure you want to delete this league?")) {
  if(confirm("Are you *really* sure?")) {
   document.frmSched.action.value="deleteLeague";
   document.frmSched.actionValue.value=document.frmSched.leagueID.value;
   document.frmSched.submit();   
  }
 }
}

function createLeague() {
 if(document.frmSched.leagueName.value=="") {
  alert("Please enter a league name");
  document.frmSched.leagueName.select();
  return;
 }
 document.frmSched.action.value="createLeague";
 document.frmSched.actionValue.value = document.frmSched.leagueName.value;
 document.frmSched.submit();
}

function addPlayer() {
 document.frmSched.action.value="addPlayer";
 document.frmSched.actionValue.value = document.frmSched.newPlayer.value;
 document.frmSched.submit();
}

function removePlayer(empID) {
 document.frmSched.action.value="removePlayer";
 document.frmSched.actionValue.value = empID;
 document.frmSched.submit();
}

function addGame() {
 if(document.frmSched.newGameWeek.value=="") {
  alert("Enter a date");
  document.frmSched.newGameWeek.select();
  return;
 }
 if(document.frmSched.newGamePlayerA.value==document.frmSched.newGamePlayerB.value) {
  alert("Please select two different players.");
  return;
 }
 document.frmSched.action.value="addGame";
 document.frmSched.submit();
}

function removeGame(matchID) {
 document.frmSched.action.value="removeGame";
 document.frmSched.actionValue.value=matchID;
 document.frmSched.submit();
}

</script>
<!-- #include file="../../footer.asp" -->
</BODY></HTML>
