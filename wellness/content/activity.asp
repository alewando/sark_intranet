<%
Response.Expires=0                                ' - works with IE 4.0 browsers. 
Response.AddHeader "Pragma","no-cache"            ' - works with proxy servers. 
Response.AddHeader "cache-control", "no-store"    ' - works with IE 5.0 browsers. 
%>
<!--
Developer:    Adam Lewandowski
Date:         10/9/2000
Description:  Wellness program activty tracking page
-->
<!-- #include file="../section.asp" -->
<!-- #include file="wellness_utils.asp" -->

<STYLE>
BODY {
 font-family: ms sans serif, arial, geneva;
 font-size: 30pt;
}

A {
 color: #0000CC;
 text-decoration: none;
 cursor: hand;
}
A:hover {
 color: #FF0000;
 text-decoration: none;
 cursor: hand;
}
A:visited {
 color: #0000CC;
 text-decoration: none;
 cursor: hand;
}

.date {
 background-color: #FFFFFF;
 font-family: sans-serif, arial, geneva;
 font-size: 8pt;
 text-align: center;
}
.activityDate {
 background-color: #CCCCCC;
 font-family: sans-serif, arial, geneva;
 font-size: 8pt;
 text-align: center;
}
.sarkActivityDate {
 background-color: #FFCCCC;
 font-family: sans-serif, arial, geneva;
 font-size: 8pt;
 text-align: center;
}
.otherActivityDate {
 background-color: #A860FA;
 font-family: sans-serif, arial, geneva;
 font-size: 8pt;
 text-align: center;
}

</STYLE>

<BODY BGOLOR="#FFFFFF">

<%
 empID = Session("ID")
 if IsEmpty(Request("month")) then
  cMonth = Month(Now)
 else
  cMonth = Request("month")
 end if
 
 if IsEmpty(Request("year")) then
  cYear = Year(Now)
 else
  cYear = Request("year")
 end if 
 
 ' Get employee name
 sql = "SELECT FirstName, LastName FROM Employee WHERE employeeID = " & empID
 set rsEmp = DBQuery(sql)
 ls_name = rsEmp("FirstName") & "&nbsp;" & rsEmp("LastName")
%>

<%
 action = Request("action")
 if action="" then
  action="showCalendar"
 end if
 
 if action="showCalendar" then
	showCalendar cMonth, cYear, empID
 elseif action="addRecords" then
    addRecords
    showCalendar cMonth, cYear, empID
 elseif  action="deleteRecords" then
    showCalendar cMonth, cYear, empID
 end if
%>

<!-- #include file="../../footer.asp" -->
