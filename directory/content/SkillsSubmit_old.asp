<!-- #include file="../../script.asp" -->
<script language=javascript>
<!--
var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}

function closeWindow(){
	//if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.reload();
	window.close();
	}
// -->
function OpenWin(page, EmpID, WinTitle){
	winTech= window.open(page + "?empID=" + EmpID, WinTitle,"height=615,width=500,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbars=1")
	}

// -->
</script>


<html>
<head>
<title>Update Skills Inventory</title>
</head>
<body bgcolor=silver onLoad="closeWindow()">

<%
sql = "select e.employeeid from employee e where e.username = " & "'" & session("Username") & "'"
set rs = DBQuery(sql)
empID = rs("employeeid")
sql = "SELECT s.EmployeeID " &_
	  "FROM Skills_Inventory s " &_ 
	  "WHERE s.EmployeeID = " & "'" & empID & "'"
	  set rs = DBQuery(sql)
	  
Dim intSkillID
Dim intSkillRanking
Dim arySkillIDs()
Dim arySkillRankings()
Dim j
intSkillID = 0
intSkillRanking = 0
	 
'Populate an array with the skills.
For Each item in Request.Form("SkillID")
  ReDim Preserve arySkillIDs(intSkillID + 1)
  intSkillID = intSkillID + 1
  arySkillIDs(intSkillID) = Request.Form("SkillID")(intSkillID)
Next
	
For Each item in Request.Form("SkillRanking")
  ReDim Preserve arySkillRankings(intSkillRanking + 1)
  intSkillRanking = intSkillRanking + 1
  arySkillRankings(intSkillRanking) = Request.Form("SkillRanking")(intSkillRanking)
Next

IF rs.eof THEN
	'-------------------------------------------
	'	Insert Skills Info				
	'-------------------------------------------
	For j = 1 to UBound(arySkillIDs)
		SkillID = arySkillIDs(j)
		SkillRanking = arySkillRankings(j)
		sql = "Insert into Skills_Inventory (EmployeeID, SkillRanking, SkillID) " &_
		  "Values ('" & empID & "','" & SkillRanking & "' , '" & SkillID & "')"
		DBQuery(sql)
	Next
Else 
	'-------------------------------------------
	'	Update Skills Info						
	'-------------------------------------------
	For j = 1 to UBound(arySkillIDs)
		SkillID = arySkillIDs(j)
		SkillRanking = arySkillRankings(j)
		sql = "UPDATE Skills_Inventory " &_
			  "SET SkillRanking = " & SkillRanking &_ 
			  "WHERE SkillID = " & SkillID &_
			  "AND EmployeeID = " & empID
		DBQuery(sql)
	Next
End If
rs.close
currDate = FormatDateTime(Date)
sql = "UPDATE Employee " &_
	  "SET DateSkillsModified = convert(datetime,'" & currDate & "',1) " &_ 
	  "WHERE EmployeeID = " & empID
DBQuery(sql)
%>
</body>
</html>
