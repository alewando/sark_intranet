<!--
Developer:    Sark Beach
Date:         12/1999
Description:  User preference page.
-->


<!-- #include file="../section.asp" -->

<%
 if hasRole("WebMaster") then
  %>
  <B>Webmaster's Instructions</B><BR>
  As webmaster, you can <A HREF="addSport.asp">add/remove sports</A> and 
  <A HREF="assignManager.asp">assign managers.</A>
  <P>
  <%
 end if
 if isManager() then
  %>
  <B>Manager's Instructions</B><BR>
  You are the manager of a sports team. Use the <A HREF="editSport.asp">sports maintenance</A>
  page to enter schedules, directions, etc. for your sport. You can use the normal sport 
  display pages (listed below and in the drop-down box at the top of the page) to record
  wins (no losses, please) and scores.<BR>
  <%
 end if
%>
<P>
<H3>Sark Sports</H3>
If you would like to have a sport added, contact <A HREF="javascript:SendEmail('cdolan')">Chris Dolan</A>.
<UL>
<%
 sql = "SELECT Sport_ID, Sport_Name FROM Sports Order By Sport_Name"
 set rsSports = DBQuery(sql)
  
 do while not rsSports.eof 
  Response.Write("<LI><A HREF='showSport.asp?sportID=" & rsSports("Sport_ID") & "'>")
  Response.Write(rsSports("Sport_Name") & "</A>")
  rsSports.moveNext
 loop
%>
<!-- Put other sport links here -->
<LI><A HREF='../../racquetball/content/about.asp'>Racquetball</A>
</UL>

<SCRIPT Language=JavaScript>
function SendEmail(recipient) {
 alert("Sending to " + recipient); winEmail= window.open("../../email.asp?editto=yes&recipient=" + recipient + "&footer=Sark Sports Page", "SendEmail","height=320,width=520,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1");
}
</SCRIPT>
<!-- #include file="../../footer.asp" -->
