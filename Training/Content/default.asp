<%
'-----------------------------------------------------------
'  This page redirects to either the user's preferred       
'  section page, or to the about section page, depending    
'  on if there is a "DefaultPage" cookie.                   
'-----------------------------------------------------------
response.buffer = true
sectionPage = request.cookies("trainingPage")
if sectionPage = "" then
	sectionPage = "about_training.asp"
else
	i = instr(sectionPage, "~")
	while i > 0
		sectionPage = left(sectionPage, i-1) & "=" & mid(sectionPage, i+1)
		i = instr(sectionPage, "~")
		wend
	end if
response.redirect sectionPage
%>
