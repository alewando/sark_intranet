<!-- #include file="../../script.asp" -->
<%
Response.ContentType = "text/xml"
empid = Request.QueryString("EmpID")
Response.Write "<?xml version=""1.0""?>"
Response.Write "<?xml-stylesheet type=""text/xsl"" href=""SkillsInv.xsl"" ?>"

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

	empID = rsEmployee.fields(0)
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
	Response.Write "Initial Entry"
	Response.Write "</DateSkillsModified>"
	
	sql = "SELECT e.EmployeeID FROM Employee e WHERE e.username = '" & session("Username") & "'"
	set rsUser = DBQuery(sql)
	userID = rsUser.fields(0)
	If userID = empID Then
		Response.Write "<CanUpdate>"
		Response.Write "True"
		Response.Write "</CanUpdate>"
	Else
		Response.Write "<CanUpdate>"
		Response.Write "False"
		Response.Write "</CanUpdate>"
	End If 
	rsUser.close
	
	sql = "SELECT sg.PageBreak, sg.GroupName, s.isOtherSkill, SkillValue = s.SkillName, s.SkillName, s.SkillID, sr.SkillRanking " &_
		  "FROM Skills s, Skills_Groups sg, Skills_Rankings sr " &_
		  "WHERE s.GroupID = sg.GroupID " &_
		  "AND sr.SkillRanking = 1 " &_
		  "ORDER BY s.GroupID, s.isOtherSkill, s.SkillName "
	set rsEmployee = DBQuery(sql)
Else
	sql = "SELECT e.EmployeeID, e.FirstName, e.LastName, e.DateSkillsModified " &_
			  "FROM Employee e " &_ 
			  "WHERE e.EmployeeID = " & empid
	set rsEmployee = DBQuery(sql)

	Response.Write "<Employee>"

	empID = rsEmployee.fields(0)
	for i=0 to rsEmployee.fields.count-1
		Response.Write "<" & rsEmployee.fields(i).name & ">"
		Response.Write rsEmployee.fields(i)
		Response.Write "</" & rsEmployee.fields(i).name & ">"
	next
	rsEmployee.close

	sql = "SELECT e.EmployeeID FROM Employee e WHERE e.username = '" & session("Username") & "'"
	set rsUser = DBQuery(sql)
	userID = rsUser.fields(0)
	If userID = empID Then
		Response.Write "<CanUpdate>"
		Response.Write "True"
		Response.Write "</CanUpdate>"
	Else
		Response.Write "<CanUpdate>"
		Response.Write "False"
		Response.Write "</CanUpdate>"
	End If 
	rsUser.close
	
	sql = "SELECT sg.PageBreak, sg.GroupName, s.isOtherSkill, si.SkillValue, s.SkillName, s.SkillID, sr.SkillRanking " &_
		  "FROM Skills_Inventory si, Skills_Groups sg, Skills s, Skills_Rankings sr " &_ 
		  "WHERE si.EmployeeID = " & empid &_
		  "AND si.SkillRanking = sr.SkillRanking " &_
		  "AND s.SkillID = si.SkillID " &_
		  "AND s.GroupID = sg.GroupID " &_
		  "ORDER BY s.GroupID, s.isOtherSkill, s.SkillName "
	set rsEmployee = DBQuery(sql)
End If
rs.close

field = ""
j = 0
pb=0

while not rsEmployee.eof
	for i=0 to rsEmployee.fields.count-1
		If rsEmployee.fields(i).name = "PageBreak" THEN
			pageBreakValue = rsEmployee.fields(i)
			If pb=0 THEN		
				Response.Write "<SkillGroup>"
				Response.Write "<PageBreak>"
				Response.Write pageBreakValue
				Response.Write "</PageBreak>"
				pb=1
			End If
		ElseIf rsEmployee.fields(i).name = "GroupName" THEN
			If j=0 THEN			
				Response.Write "<GroupName>"
				Response.Write rsEmployee.fields(i)
				Response.Write "</GroupName>"
				j = 1
				field = rsEmployee.fields(i)
			Else
				If field <> rsEmployee.fields(i) THEN
					Response.Write "</SkillGroup>" 				
					Response.Write "<SkillGroup>"
					Response.Write "<PageBreak>"
					Response.Write pageBreakValue
					Response.Write "</PageBreak>"
					Response.Write "<GroupName>"
					Response.Write rsEmployee.fields(i)
					Response.Write "</GroupName>"
					field = rsEmployee.fields(i)
				End If
			End If
		ElseIf rsEmployee.fields(i).name = "isOtherSkill" THEN
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








