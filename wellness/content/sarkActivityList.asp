<!--
Developer:    Adam Lewandowski
Date:         10/11/2000
Description:  Wellness program - Sark-sponsored activities
-->
<!-- #include file="../section.asp" -->
<!-- #include file="wellness_utils.asp" -->
<%
 action           = Request("action")
 eventID          = Request("eventID")
 eventDate        = Request("eventDate")
 eventDescription = Request("eventDescription")
 eventLink        = Request("eventLink")
 sarkSponsored    = Request("sarkSponsored")
 if sarkSponsored <> "1" then 
  sarkSponsored="0"
 end if
 
 if action="addEvent" then
  sql = "INSERT INTO Wellness_Events(Event_Date, Event_Description, Event_Link, Sark_Sponsored) "  & _
        "VALUES ('" & eventDate & "', '" & clean(eventDescription) & "', " 
  if eventLink="" then
   sql = sql & "NULL"
  else
   sql = sql & "'" & eventLink & "'"
  end if
  sql = sql & ", " & sarkSponsored & ")"
  'Response.Write("sql=" & sql & "<br>")
  DBQuery(sql)
 end if
 
 if action="deleteEvent" and not isNull(eventID) and eventID<>"" then
  sql = "DELETE FROM Wellness_Events WHERE event_ID=" & eventID
  DBQuery(sql)
 end if
 
%>
<STYLE>
BODY {
 font-family: ms sans serif, arial, geneva;
 font-size: 30pt;
}
.EmployeeHeader {
 font-family: ms sans serif, arial, geneva;
}
</STYLE>

<BODY BGOLOR="#FFFFFF">

<CENTER><H4>Sark-Sponsored Events</H4></CENTER>
<P>
The following events count for 2 days in the Wellness Program. When you go to log
 your participation in one of the events listed below, click twice on the date. 
 Sark-sponsored events will show up as red on the calendar.

<%
 sql="SELECT * FROM Wellness_Events ORDER BY event_date"
 set rsEvents = DBQuery(sql)
 Response.Write("<UL>")
 while not rsEvents.eof
  Response.Write("<LI>")
  if hasRole("WellnessAdmin") then
   Response.Write("<INPUT TYPE=BUTTON CLASS=button VALUE='Delete' onClick='deleteEvent(" & _
    rsEvents("Event_ID") & ")'>" & vbCRLF)
  end if
  if not IsNull(rsEvents("event_link")) then
	Response.Write("<A HREF='" & rsEvents("event_link") & "' TARGET='_blank'>" & _
	 rsEvents("event_description") & "</A> - " & rsEvents("event_date"))
  else
	Response.Write(rsEvents("event_description") & " - " & rsEvents("event_date"))
  end if
  rsEvents.moveNext
 wend
 Response.Write("</UL>")
 if hasRole("WellnessAdmin") then
 %>
  <FORM NAME="frmEvent" ACTION="<%=Request.ServerVariables("SCRIPT_NAME")%>" METHOD=POST>
  <INPUT TYPE=HIDDEN NAME="action" VALUE="addEvent">
  <INPUT TYPE=HIDDEN NAME="eventID" VALUE="">
  <TABLE BORDER=0 BGCOLOR="#EEEEEE">
   <TR>
    <TD><font color=red>*</font>Event Date</TD>
    <TD>
     <INPUT TYPE=TEXT SIZE=15 NAME=eventDate >
    </TD>
   </TR>
   <TR>
    <TD><font color=red>*</font>Description</TD>
    <TD>
     <INPUT TYPE=TEXT SIZE=40 NAME=eventDescription >
    </TD>
   </TR>
   <TR>
    <TD>Link<BR><FONT SIZE="-2">Ex: http://www.getwell.com/</FONT></TD>
    <TD>
     <INPUT TYPE=TEXT SIZE=40 NAME=eventLink >
    </TD>
   </TR>
   <TR>
    <TD ALIGN=RIGHT>
     <INPUT TYPE=CHECKBOX VALUE=1 NAME=sarkSponsored >
    </TD>
    <TD>SARK-sponsored</TD>
   </TR>
   <TR>
    <TD COLSPAN=2 ALIGN=CENTER>
     <INPUT TYPE=BUTTON VALUE="Add New Event" CLASS=button onClick='addEvent()'>
    </TD>
   </TR>
  </TABLE>
  </FORM>
  <SCRIPT LANGUAGE="JavaScript">
   function addEvent() {
    if (document.frmEvent.eventDate.value=="") {
     alert("Please enter a date");
     document.frmEvent.eventDate.select();
     return;
    }
    if (document.frmEvent.eventDescription.value=="") {
     alert("Please enter a description");
     document.frmEvent.eventDescription.select();
     return;
    }
    document.frmEvent.action.value="addEvent";
    document.frmEvent.submit();
   }
   
   function deleteEvent(eventID) {
    document.frmEvent.eventID.value=eventID;
    document.frmEvent.action.value="deleteEvent";
    document.frmEvent.submit();
   }
  </SCRIPT>
 <%
 end if
%>

<!-- #include file="../../footer.asp" -->
