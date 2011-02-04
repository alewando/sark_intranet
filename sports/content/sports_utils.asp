<!--
Developer:    Adam Lewandowski
Date:         10/9/2000
Description:  Sports utility functions
-->

<%
' isManager returns true if the currently logged in user is a manager
'  for at least one sport
Function isManager
 sql = "SELECT COUNT(*) as c FROM Sports_Managers WHERE Employee_ID = " & Session("ID")
 set rsCount = DBQuery(sql)
 if rsCount("c") > 0 then
  isManager = true
 else
  isManager = false
 end if
 if hasRole("WebMaster") then
  isManager = true
 end if
End Function

Function isManagerOf(sportID)
 sql = "SELECT COUNT(*) as c FROM Sports_Managers WHERE Employee_ID = " & Session("ID") & _
       " AND Sport_ID=" & sportID
 set rsCount = DBQuery(sql)
 if rsCount("c") > 0 then
  isManagerOf = true
 else
  isManagerOF = false
 end if
 if hasRole("WebMaster") then
  isManagerOf = true
 end if       
End Function

' Returns a string of sports managed by the current user 
' usable in an SQL 'IN' clause 
Function getManagedSports
 if hasRole("WebMaster") then
  sql = "SELECT sport_id FROM Sports ORDER BY sport_id"
 else
  sql = "SELECT sport_id FROM Sports_Managers WHERE employee_id = " & Session("ID")
 end if
 set rsSports = DBQuery(sql)
 sSports = "("
 do while not rsSports.eof
  sSports = sSports & rsSports("sport_id") & ","
  rsSports.movenext
 loop
 rsSports.close
 sSports = Left(sSports, Len(sSports)-1) & ")"
 getManagedSports = sSports
End Function
%>