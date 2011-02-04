 <%@ Language = Vbscript%>
<!--
Developer:	  Adam Lewandowski
Date:         11/8/2000
Description:  Display sports information
Modifications: 

-->

<!-- #include file="../section.asp" -->

<%
'Get form variables
action = Request("action")
sportID = Request("sportID")
seasonNum = Request("seasonNum")
Response.Write("<!-- seasonNum=" & seasonNum & " -->" & vbCRLF)
Response.Write("<!-- sportID=" & sportID & " -->" & vbCRLF)
Response.Write("<!-- action=" & action & " -->" & vbCRLF)
Response.Write("<!-- PATH_INFO = " & Request.ServerVariables("PATH_INFO") & " -->" & vbCRLF)
Response.Write("<!-- QUERY_STRING = " & Request.ServerVariables("QUERY_STRING") & " -->" & vbCRLF)
Response.Write("<!-- Application(Web) = " & Application("Web") & " -->" & vbCRLF)

'Check for sport info
sql = "SELECT count(*) as c FROM Sports_Season WHERE Sport_ID=" & sportID
set rsCheck = DBQuery(sql)
if rsCheck("c") > 0 then

 'Find the largest used season number for this sport
 if sportID <> "" and seasonNum="" then
  sql = "SELECT Season_Number num FROM Sports_Season WHERE Season_Start =" & _
        "(SELECT MAX(Season_Start) FROM Sports_Season  WHERE Sport_ID = " & sportID & ") " & _
        " AND Sport_ID = " & sportID
  set rs = DBQuery(sql)
  if isNull(rs) or isNull(rs("num")) then
   seasonNum = ""
  else
   seasonNum = CInt(rs("num"))
  end if 
 end if 


'Update Game Result
if action="updateGameResult" then 
 updateGameNum = Request("updateGameNum")
 updateGameScore = Request("updateGameScore")
 updateGameResult = Request("updateGameResult")
 sql = "UPDATE Sports_Schedule SET Win='" & Clean(updateGameResult) & "' " & _
       ", Score='" & Clean(updateGameScore) & "' " & _
       " WHERE " & _
       "Sport_ID = " & sportID & " AND Season_Number=" & seasonNum & _
       " AND Game_Number=" & updateGameNum
 'Response.Write("Sql="&sql&"<BR>")
 DBQuery(sql)
 if updateGameResult <> "0" then
 sql = "UPDATE Sports_Schedule SET Played=1 " & _
       "WHERE " & _
       "Sport_ID = " & sportID & " AND Season_Number=" & seasonNum & _
       " AND Game_Number=" & updateGameNum
 DBQuery(sql)
 end if
end if

 ' Load sport information if a new sport was selected 
 sportName = "None"
 sportDescription = "N/A"
 sportNotes = "None"
 if sportID <> "" then
  sql = "SELECT * FROM Sports WHERE Sport_ID = " & sportID
  set rsSport = DBQuery(sql)
  sportName = rsSport("Sport_Name")
  if sportName="" then
   sportName="None"
  end if
  sportDescription = rsSport("Sport_Description")
  if sportDescription="" then
   sportDescription = "N/A"
  end if
  sportNotes = rsSport("Notes")
  if sportNotes="" then
   sportNotes="None"
  end if
 end if

 ' Read season information for selected season
 if sportID <> "" and seasonNum <> "" then
  ' Get Season Info
  sql = "SELECT * from Sports_Season WHERE sport_id=" & sportID & " AND season_number=" & seasonNum 
  set rsSeason = DBQuery(sql)
  seasonDescription = rsSeason("Season_Description")
  seasonStart = Cstr(rsSeason("Season_Start"))
  seasonEnd = Cstr(rsSeason("Season_End"))
  rsSeason.close
  ' Get Win/Loss/Tie counts
  winCount=0
  lossCount=0
  tieCount=0
  sql = "SELECT COUNT(*) as ct FROM Sports_Schedule WHERE sport_id=" & sportID & _
        " AND season_number=" & seasonNum & " AND played=1 AND Win='W'"
  set rsCount = DBQuery(sql)
  if not isNull(rsCount) then 
 	winCount = CInt(rsCount("ct"))
  end if
  sql = "SELECT COUNT(*) as ct FROM Sports_Schedule WHERE sport_id=" & sportID & _
        " AND season_number=" & seasonNum & " AND played=1 AND Win='L'"
  set rsCount = DBQuery(sql)
  if not isNull(rsCount) then 
 	lossCount = CInt(rsCount("ct"))
  end if
  sql = "SELECT COUNT(*) as ct FROM Sports_Schedule WHERE sport_id=" & sportID & _
        " AND season_number=" & seasonNum & " AND played=1 AND Win='T'"
  set rsCount = DBQuery(sql)
  if not isNull(rsCount) then
 	tieCount = CInt(rsCount("ct"))
  end if
 else
  seasonDescription = ""
  seasonStart = ""
  seasonEnd = ""
 end if

else ' if rsCheck("c")
 %>
 No information has been entered for this sport<BR>
 <%
end if
%>

<HTML>

<HEAD>
<TITLE>Sark <%=sportName%></TITLE>
<LINK REL="stylesheet" TYPE="text/css" HREF="sports.css">
</HEAD>

<%

	
%>
<body>
<!-- 
 PATH_INF0=<%=Request.ServerVariables("PATH_INFO")%><BR> 
-->
<CENTER>
<SPAN CLASS=sportHeader><%=sportName%></SPAN>
</CENTER><P>
<TABLE CLASS=tableShadowSilver>
<TR><TD>
<B>Description:</B><BR>
<DIV CLASS=sportDescription>
<%=sportDescription%>
</DIV>
</TD></TR>
</TABLE>
<P>
<TABLE CLASS=tableShadowSilver>
<TR><TD>
<B>Notes</B> (Directions, etc):<BR>
<DIV CLASS=sportDescription>
<%=sportNotes%>
</DIV>
</TD></TR>
</TABLE>
<P>
<TABLE CLASS=tableShadowSilver>
<TR><TD>
<B>Manager</B><BR>
<DIV CLASS=sportDescription>
<%
 sql = "SELECT Employee_ID, FirstName, LastName FROM Sports_Managers, Employee" & _
       " WHERE Employee_ID=employeeid AND Sport_ID=" & sportID
 set rsManagers = DBQuery(sql)
 do while not rsManagers.eof
  Response.Write("<A HREF='../../directory/content/details.asp?EmpID=" & rsManagers("Employee_ID") & "'>")
  Response.Write(rsManagers("FirstName") & " " & rsManagers("LastName") & "</A><BR>" & vbCRLF)
  rsManagers.moveNext
 loop
%>
</DIV>
</TD></TR>
</TABLE>

<form NAME="frmSport" ACTION="<%=Request.ServerVariables("SCRIPT_NAME")%>" method=post>
  <INPUT TYPE=Hidden NAME="sportID" VALUE="<%=sportID%>">
  <INPUT TYPE=Hidden NAME="action" VALUE="">
  <INPUT TYPE=Hidden NAME="updateGameNum" VALUE="">
  <INPUT TYPE=Hidden NAME="updateGameResult" VALUE="">
  <INPUT TYPE=Hidden NAME="updateGameScore" VALUE="">
  <BR>
  <!-- Season selection drop-down -->
  <B>Season</B>
  <SELECT NAME="seasonNum" onChange="seasonSelect()">
  <%
   sql = "SELECT * FROM Sports_Season WHERE Sport_ID=" & sportID & " ORDER BY Season_Start"
   set rsSeason = DBQuery(sql)
   do while not rsSeason.eof
    Response.Write("<OPTION VALUE='" & rsSeason("Season_Number") & "'")
    if CInt(rsSeason("Season_Number"))=CInt(seasonNum) then
     Response.Write(" SELECTED")
    end if
    Response.Write(">" & rsSeason("Season_Description") & "</OPTION>")
    rsSeason.moveNext
   loop
  %>
  </SELECT>
<P>
<SPAN CLASS=sportRecord>
<B>Record:</B><%=winCount%>-<%=lossCount%>-<%=tieCount%><BR>
</SPAN>
<P>
<DIV CLASS=tableShadow>
<B>Schedule:</B>
<TABLE border=0 WIDTH=100%>
<TR>
 <TH><font color=navy>Date</font></TH>
 <TH><font color=navy>Time</font></TH>
 <TH><font color=navy>Opp</font></TH>
 <TH><font color=navy>W/L/T</font></TH>
 <TH><font color=navy>Score</font></TH>
</TR>
<%
 sql = "SELECT * FROM Sports_Schedule WHERE Sport_ID=" & sportID & " AND Season_Number=" & seasonNum & " ORDER BY Game_Date"
 set rsSched = DBQuery(sql)
 do while not rsSched.eof
  %>
  <TR>
   <TD align=center><%=rsSched("Game_Date")%></TD>
   <TD align=center><%=rsSched("Game_Time")%></TD>
   <TD align=center><%=rsSched("Opponent")%></TD>
   <TD align=center>
    <% if isManagerOf(sportID) then
     Response.Write("<SELECT NAME=WinLossTie" & rsSched("Game_Number") & ">")
     Response.Write("<OPTION VALUE='W'")
     if ucase(rsSched("Win"))="W" then 
      Response.Write(" SELECTED ")
     end if 
     Response.Write("> Win" & vbCRLF)

     Response.Write("<OPTION VALUE='L'")
     if ucase(rsSched("Win"))="L" then 
      Response.Write(" SELECTED ")
     end if 
     Response.Write("> Loss" & vbCRLF)

     Response.Write("<OPTION VALUE='T'")
     if ucase(rsSched("Win"))="T" then 
      Response.Write(" SELECTED ")
     end if 
     Response.Write("> Tie" & vbCRLF)

     Response.Write("<OPTION VALUE='0'")
     if isNull(rsSched("Win")) or rsSched("Win")="0" or rsSched("Win")="" then 
      Response.Write(" SELECTED ")
     end if 
     Response.Write("> N/A" & vbCRLF)

     Response.Write("</SELECT>")
     Response.Write("</TD><TD>")
     Response.Write("<INPUT TYPE=TEXT SIZE=10 NAME='game" & rsSched("Game_Number") & "Score' VALUE='" & rsSched("Score") & "'>")
     Response.Write("</TD><TD>")
     Response.Write("<INPUT TYPE=BUTTON CLASS=button VALUE=""Update"" " & _
                    "onClick='updateResult(" & rsSched("Game_Number") & _
                    ", document.frmSport.WinLossTie" & rsSched("Game_Number") & ".value" & _
                    ", game" & rsSched("Game_Number") & "Score.value)'>")
    else 
     if isNull(rsSched("Win")) or rsSched("Win")="" or rsSched("Win")="0" then
      Response.Write("&nbsp</TD>")
      if (not isNull(rsSched("Score"))) or rsSched("Score")<>"" then
       Response.Write("<TD>" & rsSched("Score") & "</TD>")
      else
       Response.Write("<TD>&nbsp;</TD>")
      end if
     else
      Response.Write(rsSched("Win") & "</TD>")
      if (not isNull(rsSched("Score"))) or rsSched("Score")<>"" then
       Response.Write("</TD><TD align=center>" & rsSched("Score") & "</TD>")
      else
       Response.Write("</TD><TD>&nbsp;</TD>")
      end if
     end if
    end if%>
   </TD>
  </TR>
  <%
  rsSched.moveNext
 loop
%>
</TABLE>
</DIV>
</form>
</body>

<SCRIPT LANGUAGE=javascript>
<!--
function seasonSelect() {
 document.frmSport.action.value="selectSeason";
 document.frmSport.submit();
}

function foo(arg1, arg2) {
 alert("hi " + arg1 + " " + arg2);
}

function updateResult(gameNum, gameResult, gameScore) {
 document.frmSport.action.value="updateGameResult";
 document.frmSport.updateGameNum.value = gameNum;
 document.frmSport.updateGameResult.value = gameResult;
 document.frmSport.updateGameScore.value = gameScore;
 document.frmSport.submit();
}
//-->
</SCRIPT>
</HTML>







<!-- #include file="../../footer.asp" -->