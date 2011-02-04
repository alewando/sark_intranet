<%@ Language = Vbscript%>
<!--
Developer:    Adam Lewandowski
Date:         11/21/2000
Description:  League Schedules
-->
<!-- #include file="../section.asp" -->

<%
 leagueID = Request("leagueID")
 action = Request("action")
 
 Response.Write("<!-- leagueID=" & leagueID & " -->" & vbCRLF)
 Response.Write("<!-- action=" & action & " -->" & vbCRLF)
 
 if leagueID <> "" then 
  ' Load league description
  sql = "SELECT * FROM Racquetball_League WHERE League_ID=" & leagueID
  set rs = DBQuery(sql)
  leagueName = rs("Description")
 end if
 
 'Update scores
 if action="updateScores" then
  For Each strItem in Request.Form
   if left(strItem,len("MatchWinner")) = "MatchWinner" then
    matchNum = right(strItem, len(strItem)-len("MatchWinner"))
    winner = Request.Form(strItem)
    'Response.Write("Updating match " & matchNum & " winner to " & winner & "<BR>")
    sql = "UPDATE Racquetball_Schedule SET Winner=" & winner & " WHERE Match_ID=" & matchNum
    DBQuery(sql)
   end if
  Next
 end if
%>

<HTML><HEAD>
<TITLE>Sark Racquetball - <%=leagueName%></TITLE>
<LINK rel="stylesheet" type="text/css" href="..\style.css"></HEAD>
<BODY>

<P><SPAN CLASS=leagueHeader><%=leagueName%></SPAN></P>
<SPAN CLASS=standingsHeader>Standings</SPAN>
<DIV>
<TABLE width="100%" border=0 class=StandingsTable>
<%
 sql = "SELECT Employee_ID, LastName, FirstName FROM Racquetball_Players rp, Employee e " & _
       "WHERE rp.Employee_ID = e.employeeid " &_
       "AND League_ID=" & leagueID & _
       "ORDER BY LastName, FirstName" 
 set rsPlayers=DBQuery(sql)
 do while not rsPlayers.eof
  %><TR><TD><A HREF="../../directory/content/details.asp?EmpID=<%=rsPlayers("Employee_ID")%>">
  <%=rsPlayers("FirstName") & " " & rsPlayers("LastName")%></A></TD>
  <TD>
  <%
   ' Get win/loss record
   set rsRecord=DBQuery("SELECT COUNT(*) as wins FROM Racquetball_Schedule WHERE " & _
                 "winner IS NOT NULL " & _
                 "AND winner = " & rsPlayers("Employee_ID") & _
                 " AND league_ID=" & leagueID)
   if not isNull(rsRecord) and not isNull(rsRecord("wins")) then 
    wins=rsRecord("wins")
   else
    wins=0
   end if
   set rsRecord=DBQuery("SELECT COUNT(*) as losses FROM Racquetball_Schedule WHERE (playerA = " & rsPlayers("Employee_ID") & _
                 " OR playerB = " & rsPlayers("Employee_ID") & ") AND winner IS NOT NULL " & _
                 "AND winner <> " & rsPlayers("Employee_ID") & _
                 " AND league_ID=" & leagueID)
   if not isNull(rsRecord) and not isNull(rsRecord("losses")) then 
    losses=rsRecord("losses")
   else
    losses=0
   end if
   Response.Write(wins & " - " & losses)
  %>
  </TD>
  </TR>
 <%
  rsPlayers.moveNext
 loop
%>
</TABLE>
</DIV>

<P>
<SPAN CLASS=scheduleHeader>Schedule</SPAN>
<TABLE BORDER=0 BGCOLOR="#f0f0f0">
<TR COLSPAN=2><TD>Key</TD>
<TR CLASS=lateGame><TD WIDTH=15>&nbsp;</TD><TD STYLE="background-color: transparent"><font size="-2">No Score Reported</font></TD></TR>
<TR><TD WIDTH=15 CLASS=playerWinner>&nbsp;</TD><TD STYLE="background-color: transparent"><font size="-2">Winner</font></TD></TR>
</TABLE>
<BR>
<%
 if (hasRole("RacquetballAdmin")) then 
  %>
  <FORM NAME=frmRB ACTION="<%=Request.ServerVariables("SCRIPT_NAME")%>" METHOD=POST>
  <INPUT NAME="leagueID" TYPE=HIDDEN VALUE="<%=leagueID%>">
  <%
 end if
 sql = "SELECT distinct(week) as week FROM Racquetball_Schedule WHERE League_ID=" & leagueID
 set rsWeeks=DBQuery(sql)
 do while not rsWeeks.eof
  Response.Write("<SPAN CLASS=weekHeader>" & rsWeeks("week") & "</SPAN><BR>")
  %><TABLE BORDER=0 WIDTH=100% CELLSPACING=0 CELLPADDING=0><%
  sql = "SELECT match_ID,playerA, playerB, e1.firstName+' '+e1.lastName as NameA, " & _
        "e2.firstName+' '+e2.lastName as NameB, winner, " & _
        "CASE WHEN (week < (GETDATE() - 7) AND winner is null) THEN 1 ELSE 0 END AS late " & _
        "FROM Racquetball_Schedule rs, Employee e1, Employee e2 " & _
        "WHERE rs.playerA=e1.employeeid AND rs.playerB=e2.employeeid " & _
        "AND League_ID=" & leagueID & " AND Week='" & rsWeeks("week") & "'"
  'Response.Write("sql = " & sql & "<BR>")
  set rsSched = DBQuery(sql)
  do while not rsSched.eof
   if rsSched("late")=0 then
    Response.Write("<TR>")
   else
    Response.Write("<TR CLASS=lateGame>")
   end if  
   %>
    <TD width=45%
    <%if rsSched("winner")=rsSched("playerA") then
      Response.Write(" class=playerWinner ")
     else
      Response.Write(" class=playerLoser ")
     end if
    %>
    >
     <% if hasRole("RacquetballAdmin") then %>
      <INPUT TYPE=RADIO NAME="MatchWinner<%=rsSched("Match_ID")%>" VALUE="<%=rsSched("playerA")%>"
      <%if rsSched("winner")=rsSched("playerA") then
        Response.Write(" CHECKED ")
       end if
      %>>
     <% end if %>
     <% if rsSched("playerA")=Session("ID") then
      Response.Write("<B>" & rsSched("NameA") & "</B>")
     else
      Response.Write(rsSched("NameA"))
     end if 
     %>
    </TD>
    <TD width=10%>vs</TD>
    <TD width=45% 
    <%if rsSched("winner")=rsSched("playerB") then
      Response.Write(" class=playerWinner ")
     else
      Response.Write(" class=playerLoser ")
     end if
    %>
    >
     <% if hasRole("RacquetballAdmin") then %>
      <INPUT TYPE=RADIO NAME="MatchWinner<%=rsSched("Match_ID")%>" VALUE="<%=rsSched("playerB")%>"
      <%if rsSched("winner")=rsSched("playerB") then
        Response.Write(" CHECKED ")
       end if
      %>>
     <% end if %>
     <% if rsSched("playerB")=Session("ID") then
      Response.Write("<B>" & rsSched("NameB") & "</B>")
     else
      Response.Write(rsSched("NameB"))
     end if 
     %>
    </TD>
   </TR>
   <%
   rsSched.moveNext
  loop
  Response.Write("</TABLE>")
  rsWeeks.moveNext
  Response.Write("<BR>")
 loop
 
 if (hasRole("RacquetballAdmin")) then 
  %>
  <CENTER><INPUT CLASS=button TYPE=BUTTON onClick="updateScores()" VALUE="Update Scores"></CENTER>
  <INPUT TYPE=HIDDEN NAME="action" VALUE="">
  </FORM>
  <%
 end if
%>
<script language="JavaScript">
function updateScores() {
 document.frmRB.action.value="updateScores";
 document.frmRB.submit();
}
</script>
<!-- #include file="../../footer.asp" -->
</BODY></HTML>
