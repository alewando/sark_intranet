<%@ Language = Vbscript%>
<!--
Developer:	  Adam Lewandowski
Date:         10/24/2000
Description:  Tool for sport Mangers to edit sport information.
Modifications: 

-->

<!-- #include file="../section.asp" -->

<%
'Get form variables
action = Request("action")
sportID = Request("sportID")
notes = Request("notes")
description = Request("description")
seasonNum = Request("seasonNum")
seasonDescription = Request("seasonDescription")
seasonStart = Request("seasonStartDate")
seasonEnd = Request("seasonEndDate")

'Find the largest used season number for this sport
if sportID <> "" then
 sql = "SELECT MAX(season_number) as num FROM Sports_Season WHERE sport_ID = " & sportID
 set rs = DBQuery(sql)
 if isNull(rs("num")) then
  maxSeasonNum = 0
 else
  maxSeasonNum = CInt(rs("num"))
 end if 
 'Response.Write("Max Season = " & maxSeasonNum & "<BR>")
end if

'Create a new season if requested
if action = "updateSeason" then
 if seasonNum = "-1" then
  sql = "INSERT INTO Sports_Season " & _
        "(sport_id, season_number, season_description, season_start, season_end) " & _
        "VALUES (" & sportID & "," & (maxSeasonNum+1) & ",'" & replace(seasonDescription, "'", "''") & "'" & _
        ",'" & seasonStart & "','" & seasonEnd & "')"
  'Response.Write("sql = " & sql & "<BR>")
  DBQuery(sql)
  seasonNum = maxSeasonNum+1
 else
  sql = "UPDATE Sports_Season SET " & _
        "season_description = '" & replace(seasonDescription, "'", "''") & "', " & _
        "season_start = '" & seasonStart & "', " & _
        "season_end = '" & seasonEnd & "' " & _
        "WHERE sport_id = " & sportID & " AND season_number = " & seasonNum
  'Response.Write("sql = " & sql & "<BR>")
  DBQuery(sql)   
  seasonNum = maxSeasonNum+1
 end if
end if

' Load sport information if a new sport was selected
if action = "loadSport" then
 sql = "SELECT * FROM Sports WHERE Sport_ID = " & sportID
 set rsSport = DBQuery(sql)
 description = rsSport("Sport_Description")
 notes = rsSport("Notes")
 if(Request("sportChanged")="true") then
	seasonNum = -1
 end if
end if

' Read sport name
if sportID<>"" and sportID<>"-1" then
 sql = "SELECT * FROM Sports WHERE Sport_ID = " & sportID
 set rsSport = DBQuery(sql)
 description = rsSport("Sport_Description")
 notes = rsSport("Notes")
 sportName = rsSport("Sport_Name") 
end if

' Update sport information
if action = "update" then
 sql = "UPDATE Sports SET Sport_Description='" & replace(Request("description"),"'","''") & "' " & _
       ", Notes='" & Request("notes") & "' " & _
       "WHERE Sport_ID = " & sportID
 'Response.Write("sql = " & sql & "<BR>")
 DBQuery(sql)
end if

' Remove game from schedule
if action = "removeGame" then
 removeGameID = Request("removeGameID")
 'Get Game Date (for use in deletes below)
 sql = "SELECT Game_Date FROM Sports_Schedule WHERE Sport_ID=" & sportID & _
       " AND Season_Number=" & seasonNum & " AND Game_Number=" & removeGameID
 set rsDate = DBQuery(sql)
 delGameDate = rsDate("Game_Date")
 
 ' Delete from sports schedule
 sql = "DELETE FROM Sports_Schedule WHERE Sport_ID=" & sportID & _
       " AND Season_Number=" & seasonNum & " AND Game_Number=" & removeGameID
 DBQuery(sql)
 
 'Remove from Wellness calendar
 sql = "DELETE FROM Wellness_Events WHERE event_date='" & delGameDate & "' AND " & _
       "event_description='" & sportName & "'"
 DBQuery(sql)
 
 'Remove from Sark Events calendar
 sql = "SELECT Event.event_ID from Event, Event_Date WHERE " & _
       "Event.event_ID = Event_Date.event_ID " & _
       " AND eventDate='" & delGameDate & "' AND " & _
       "Title='" & sportName & "'"
 set rsID = DBQuery(sql)
 eventID = rsID("event_ID")
 if not rsID.eof then
  sql = "DELETE FROM Event_Date WHERE Event_ID=" & eventID
  DBQuery(sql)
  sql = "DELETE FROM Event WHERE Event_ID=" & eventID
  DBQuery(sql)
 end if
end if

' Add game to schedule
if action = "addGame" then
 newGameDate = Request("newGameDate")
 newGameTime = Request("newGameTime")
 newGameOpponent = Request("newGameOpponent")
 ' Get next game number
 sql = "SELECT MAX(Game_Number) as Game_Number FROM Sports_Schedule WHERE Sport_ID=" & sportID & _
       " AND Season_Number=" & seasonNum
 set rsGameNum = DBQuery(sql)
 newGameNum = rsGameNum("Game_Number") + 1
 if isNull(newGameNum) or newGameNum="" then
  newGameNum=1
 end if
 sql = "INSERT INTO Sports_Schedule (Sport_ID, Season_Number, Game_Number, Game_Date, Game_Time, Played, Win, Opponent) VALUES " & _
       "(" & sportID & "," & seasonNum & "," & newGameNum & ",'" & newGameDate & "', " & _
       "'" & newGameTime & "'" & _
       ",0,0,'" & clean(newGameOpponent) & "')"
 'Response.Write("sql=" & sql & "<BR>")
 DBQuery(sql)
 
 'Add to wellness calendar
 eventDesc = sportName
 'if trim(newGameOpponent) <> "" then
 ' eventDesc = eventDesc & " vs " & newGameOpponent
 'end if
 sql = "INSERT INTO Wellness_Events (event_date, event_description, sark_sponsored) " & _
       "VALUES ('" & newGameDate & "', '" & eventDesc & "', 1)"
 DBQuery(sql)
 
 'Add to Sark Event Calendar
 sql = "INSERT INTO Event (Title, Description, Contact_ID, Event_Sub_Type_ID, Creator_ID, SignUp, CreationDate) " & _
       "VALUES ('" & clean(eventDesc) & "', '" & clean(newGameOpponent) & "', " & _ 
       Session("ID") & ", 24" & _
       ", " & session("ID") & ",0, GETDATE())"
 DBQuery(sql)
 sql = "SELECT MAX(Event_ID) as ID FROM Event WHERE Title='" & clean(eventDesc) & "'"
 set rsID = DBQuery(sql)
 eventID = rsID("ID")
 sql = "INSERT INTO Event_Date (event_ID, eventDate, StartTime, EndTime) VALUES (" & _
       eventID & ", '" & newGameDate & "', '" & newGameTime & "','')"
 DBQuery(sql)
end if

' Read season information for selected season
if sportID <> "" and seasonNum <> "" and seasonNum <> "-1" then
 sql = "SELECT * from Sports_Season WHERE sport_id=" & sportID & " AND season_number=" & seasonNum 
 'Response.Write("sql=" & sql & "<BR>")
 set rsSeason = DBQuery(sql)
 if not rsSeason.eof then
  seasonDescription = rsSeason("Season_Description")
  seasonStart = Cstr(rsSeason("Season_Start"))
  seasonEnd = Cstr(rsSeason("Season_End"))
  rsSeason.close
 end if
else
 seasonDescription = ""
 seasonStart = ""
 seasonEnd = ""
end if


%>


<HTML>

<HEAD>

</HEAD>



<%

	
%>
<body>

<table border=0><tr><td><font size=1 face="ms sans serif, arial, geneva" color=blue><b><STRONG>
  <font size=2 color=red><b>Edit a Sport</b></font></font></td></tr>
</table>

<form NAME="frmSport" ACTION="<%=Request.ServerVariables("SCRIPT_NAME")%>" method=post>
  <INPUT TYPE=Hidden NAME="action" VALUE=""><br>
  <table border=0 class=tableShadow>
  <!-- Sport Drop-down -->
  <TR>
   <TD align=left><font face="ms sans serif, arial, geneva" size=1 color=blue>
    <STRONG>Sport:</STRONG></font>
   </td>
   <td>
    <INPUT TYPE=HIDDEN NAME="sportChanged" VALUE="false">
    <Select NAME="sportID" onChange="document.frmSport.sportChanged.value=true; loadSport()">
    <%
     if sportID="" or isNull(sportID) then
      %><OPTION VALUE=""></OPTION><%
     end if
     sql = "SELECT s.sport_ID, s.sport_Name FROM Sports s " & _
           "WHERE s.Sport_ID in " & getManagedSports() & " " & _
           "ORDER BY Sport_Name"
     'Response.Write("sql="&sql&"<BR>")
     set rs = DBQuery(sql)
     do while not rs.eof 
      Response.Write("<Option Value=" & trim(rs("sport_ID")) )
      if Cint(rs("sport_ID")) = CInt(sportID) then
       Response.Write(" SELECTED ")
      end if
      %>
      ><%=trim(left(rs("sport_Name"),45))%> 
      <% rs.movenext
     loop
     rs.close %>
	</Select>
	</td>
  </tr>
 <%
 if sportID<>"" then  
 %>
  <!-- Description box -->
  <TR>
   <TD align=left><font face="ms sans serif, arial, geneva" size=1 color=blue>
    <STRONG>Description:</STRONG></font><BR>
    <FONT face="ms sans serif, arial, geneva" SIZE="-2">(HTML allowed)</FONT><BR>
    <INPUT TYPE=BUTTON CLASS=button VALUE="Preview" onClick="preview('Description', document.frmSport.description.value)">
   </TD>
   <TD>
    <TEXTAREA NAME=description ROWS=5 COLS=45 WRAP=Virtual><%=description%></TEXTAREA>
   </TD>   
  </TR>
  
  <!-- Notes -->
  <TR>
   <TD align=left><font face="ms sans serif, arial, geneva" size=1 color=blue>
    <STRONG>Notes:</STRONG></font><BR><FONT face="arial" size=-2 color=black>
    Directions, Links, etc.</FONT><BR>
    <FONT face="ms sans serif, arial, geneva" SIZE="-2">(HTML allowed)</FONT><BR>
    <INPUT TYPE=BUTTON CLASS=button VALUE="Preview" onClick="preview('Notes', document.frmSport.notes.value)" id=BUTTON1 name=BUTTON1>
   </TD>
   <TD>
    <TEXTAREA NAME=notes ROWS=5 COLS=45 WRAP=Virtual><%=notes%></TEXTAREA>
   </TD>   
  </TR>
  
  
  <!-- Submit Button -->
  <TR>
   <TD COLSPAN=2 ALIGN=CENTER>
    <INPUT TYPE=BUTTON VALUE="Submit" CLASS=button onClick="submitForm()">
   </TD>
  </TR>  
  
  <% 
  end if 'End if sportID<>""
  %>
  </table>

 <%if sportID <> "" and description<>"" then %>
  <table border=0> <!-- this table has 1 row, 2 columns. A column for season start/end, a column for schedule -->
  <TR>
  <TD VALIGN=TOP>
  <table border=0 class=tableShadow>
  <TR><TD COLSPAN=2 ALIGN=LEFT><H5>Season Info</H5></TD></TR>
  <!-- Season Drop-down -->
  <TR>
   <TD align=left><font face="ms sans serif, arial, geneva" size=1 color=blue>
    <STRONG>Season:</STRONG></font>
   </td>
   <td>
    <Select NAME="seasonNum" onChange="seasonChange()">
    <OPTION VALUE="-1"
    <% if seasonNum="" then
     Response.Write(" SELECTED ")
    end if
    %>>Add New Season</OPTION>
    <%
     sql = "SELECT season_number, season_description FROM Sports_Season " & _
           "WHERE Sport_ID = " & sportID & " " & _
           "ORDER BY Season_Start"
     set rs = DBQuery(sql)
     do while not rs.eof 
      Response.Write("<Option Value=" & trim(rs("season_number")) )
      if seasonNum <> -1 and rs("season_number") = CInt(seasonNum) then
       Response.Write(" SELECTED seasonNum=" & Cint(seasonNum) & " rs(season_number)=" & Cint(rs("season_number")) & "")
      end if
      %>
      ><%=trim(left(rs("season_description"),45))%> 
      <% rs.movenext
     loop
     rs.close %>
	</Select>
	</td>
  </tr>

  <!-- Season Description -->
  <TR>
   <TD align=left><font face="ms sans serif, arial, geneva" size=1 color=blue>
    <STRONG>Description:</STRONG></font>
   </td>
   <td>
    <INPUT TYPE=TEXT NAME=seasonDescription VALUE="<%=seasonDescription%>">
   </td>
  </tr>
  
  <!-- Season Start Date -->
  <TR>
   <TD align=left><font face="ms sans serif, arial, geneva" size=1 color=blue>
    <STRONG>Start Date:</STRONG></font><BR>
    <FONT face="arial" SIZE=-2>mm/dd/yyyy</FONT>
   </td>
   <td>
    <INPUT TYPE=TEXT NAME=seasonStartDate VALUE="<%=seasonStart%>">
   </td>
  </tr>
  
  <!-- Season End Date -->
  <TR>
   <TD align=left><font face="ms sans serif, arial, geneva" size=1 color=blue>
    <STRONG>End Date:</STRONG></font><BR>
    <FONT face="arial" SIZE=-2>mm/dd/yyyy</FONT>
   </td>
   <td>
    <INPUT TYPE=TEXT NAME=seasonEndDate VALUE="<%=seasonEnd%>">
   </td>
  </tr>

  <!-- Season Submit Button -->
  <TR>
   <TD COLSPAN=2 ALIGN=CENTER>
    <INPUT TYPE=BUTTON VALUE="Update Season Info" CLASS=button onClick="updateSeason()">
   </TD>
  </TR>  
    
  </table>
  </TD>
  <% if seasonNum <> "-1" and seasonNum <> "" then %>
  <TD VALIGN=TOP>
  <!-- Season Schedule-->
  <INPUT TYPE=HIDDEN NAME="removeGameID" VALUE="">
  <TABLE BORDER=0 class=tableShadow>
  <TR><TD COLSPAN=2 ALIGN=LEFT><H5>Schedule</H5></TD></TR>
  <%
   ' Retrieve existing schedule
   sql = "SELECT * FROM Sports_Schedule WHERE sport_id = " & sportID & _
         " AND season_number = " & seasonNum & " ORDER BY game_date"
   'Response.Write("sql=" & sql & "<BR>")
   set rsSched = DBQuery(sql)
   do while not rsSched.eof
    Response.Write("<TR><TD>" & rsSched("Game_Date") & " " & rsSched("Opponent") & " - " & rsSched("Game_Time") & "</TD>")
    Response.Write("<TD><INPUT TYPE=BUTTON VALUE='Remove Game' " & _
                   "CLASS=button onClick='removeGame(" & rsSched("Game_Number") & ")'></TD></TR>" & vbCRLF)
    rsSched.movenext
   loop
  %>
   <TR>
    <TD colspan=2>
     <TABLE BORDER=0 BGCOLOR="#bbbbaa" width=100%>
     <TR><TD>
      <INPUT TYPE=TEXT NAME=newGameDate SIZE=10 MAXLENGTH=10>
      <FONT face="arial" SIZE=-2>mm/dd/yyyy</FONT><BR>
      <INPUT TYPE=TEXT NAME=newGameTime SIZE=15>
      <FONT face="arial" SIZE=-2>Game Time</FONT><BR>
      <INPUT TYPE=TEXT NAME=newGameOpponent SIZE=15>
      <FONT face="arial" SIZE=-2>Oponent</FONT>
     </TD></TR>
     <TR>   
      <TD ALIGN=CENTER><INPUT NAME="btnAddGame" TYPE=BUTTON VALUE="Add Game" CLASS=button onClick='addGame()'></TD>
     </TR>     
     </TABLE>
   </TD>
   </TR>
  </TABLE>
  </TD>
  <% end if %>
  </TR>
  </table>
 <% end if 'End if sportID<>"" %>
</form>
<form name=frmPreview ACTION="preview.asp" METHOD=POST TARGET="_blank">
 <INPUT TYPE=HIDDEN NAME=title>
 <INPUT TYPE=HIDDEN NAME=text>
</FORM>
</body>

<SCRIPT LANGUAGE=javascript>
<!--
function submitForm() {
 document.frmSport.action.value="update";
 document.frmSport.submit();
}

function loadSport() {
 document.frmSport.action.value="loadSport";
 document.frmSport.submit();
}

function updateSeason() {
 if(document.frmSport.seasonDescription.value=="" ) {
    alert("Please enter a description for the season");
    return;
 } 
 if(    document.frmSport.seasonStartDate.value=="" ) {
    alert("Please enter a season starting date");
    return;
 } 
 if(document.frmSport.seasonEndDate.value =="" ) {
    alert("Please enter a season ending date ");
    return;
 }
 document.frmSport.action.value="updateSeason";
 document.frmSport.submit();
}

function seasonChange() {
 //alert("seasonChange(), seasonNum="+document.frmSport.seasonNum.value);
 if (document.frmSport.seasonNum.value != -1) {
  loadSport()
 } else {
  document.frmSport.seasonDescription.value="";
  document.frmSport.seasonStartDate.value="";
  document.frmSport.seasonEndDate.value="";
  if (document.frmSport.btnAddGame) {
   document.frmSport.btnAddGame.disabled=true;
  }
 }
}

function removeGame(gameID) {
 document.frmSport.action.value="removeGame";
 document.frmSport.removeGameID.value=gameID;
 document.frmSport.submit();
}

function addGame() {
 if(document.frmSport.newGameDate.value=="" ) {
    alert("Please enter a game date");
    return;
 } 
 if(document.frmSport.newGameTime.value=="" ) {
    alert("Please enter a game time");
    return;
 } 
 if(document.frmSport.newGameOpponent.value=="" ) {
    alert("Please enter a game opponent");
    return;
 } 
 document.frmSport.action.value="addGame";
 document.frmSport.submit();
}

function preview(title, text) {
 document.frmPreview.title.value=title;
 document.frmPreview.text.value=text;
 document.frmPreview.submit();
}

//-->
</SCRIPT>
</HTML>







<!-- #include file="../../footer.asp" -->