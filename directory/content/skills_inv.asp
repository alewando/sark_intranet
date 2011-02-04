<!-- #include file="../../script.asp" -->
<%
Response.ContentType = "text/xml"
empid = Request.QueryString("EmployeeID")Response.Write "<?xml version='1.0' ?>"

sql = "SELECT s.EmployeeID " &_
	  "FROM Skills_Inventory s " &_ 
	  "WHERE s.EmployeeID = " & empid
set rs = DBQuery(sql)

If rs.eof Then
	sql = "SELECT e.EmployeeID, e.FirstName, e.LastName " &_
			  "FROM Employee e " &_ 
			  "WHERE e.EmployeeID = " & empid
	set rsEmployee = DBQuery(sql)

	Response.Write "<Employee>"

	while not rsEmployee.eof
		for i=0 to rsEmployee.fields.count-1
			Response.Write "<" & rsEmployee.fields(i).name & ">"
			Response.Write rsEmployee.fields(i)
			Response.Write "</" & rsEmployee.fields(i).name & ">"
		next
	rsEmployee.movenext
	wend
	rsEmployee.close
	
	Response.Write "<DateSkillsModified>"
	Response.Write "Not submitted yet..."
	Response.Write "</DateSkillsModified>"	
	sql = "SELECT sg.GroupName, s.SkillName, s.SkillID, sr.SkillRanking " &_
		  "FROM Skills s, Skills_Groups sg, Skills_Rankings sr " &_
		  "WHERE s.GroupID = sg.GroupID " &_
		  "AND sr.SkillRanking = 1 " &_
		  "ORDER BY s.GroupID"
	set rsEmployee = DBQuery(sql)
Else
	sql = "SELECT e.EmployeeID, e.FirstName, e.LastName, e.DateSkillsModified " &_
			  "FROM Employee e " &_ 
			  "WHERE e.EmployeeID = " & empid
	set rsEmployee = DBQuery(sql)

	Response.Write "<Employee>"

	while not rsEmployee.eof
		for i=0 to rsEmployee.fields.count-1
			Response.Write "<" & rsEmployee.fields(i).name & ">"
			Response.Write rsEmployee.fields(i)
			Response.Write "</" & rsEmployee.fields(i).name & ">"
		next
	rsEmployee.movenext
	wend
	rsEmployee.close

	sql = "SELECT sg.GroupName, s.SkillName, s.SkillID, sr.SkillRanking " &_
		  "FROM Skills_Inventory si, Skills_Groups sg, Skills s, Skills_Rankings sr " &_ 
		  "WHERE si.EmployeeID = " & empid &_
		  "AND si.SkillRanking = sr.SkillRanking " &_
		  "AND s.SkillID = si.SkillID " &_
		  "AND s.GroupID = sg.GroupID " &_
		  "ORDER BY s.GroupID"
	set rsEmployee = DBQuery(sql)
End If
rs.close

field = ""
j = 0
while not rsEmployee.eof
	for i=0 to rsEmployee.fields.count-1
		If rsEmployee.fields(i).name = "GroupName" THEN
			If j=0 THEN			
				Response.Write "<SkillGroup>"
				Response.Write "<GroupName>"
				Response.Write rsEmployee.fields(i)
				Response.Write "</GroupName>"
				j = 1
				field = rsEmployee.fields(i)
			Else
				If field <> rsEmployee.fields(i) THEN
					Response.Write "</SkillGroup>" 				
					Response.Write "<SkillGroup>"
					Response.Write "<GroupName>"
					Response.Write rsEmployee.fields(i)
					Response.Write "</GroupName>"
					field = rsEmployee.fields(i)
				End If
			End If
		ElseIf rsEmployee.fields(i).name = "SkillName" THEN
			Response.Write "<Skill>"
			Response.Write "<" & rsEmployee.fields(i).name & ">"
			Response.Write rsEmployee.fields(i)
			Response.Write "</" & rsEmployee.fields(i).name & ">"
		Else 
			Response.Write "<" & rsEmployee.fields(i).name & ">"
			Response.Write rsEmployee.fields(i)
			Response.Write "</" & rsEmployee.fields(i).name & ">"
			If rsEmployee.fields(i).name = "SkillRanking" THEN
				Response.Write "</Skill>"
			End If
		End If
	next
rsEmployee.movenext
wend
rsEmployee.close

Response.Write "</SkillGroup>"
Response.Write "</Employee>"
%>








