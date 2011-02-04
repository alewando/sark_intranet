<!--
Developer:    Erica Gingrich
Date:         09/11/98
Description:  About Directory Section
-->
<!-- #include file="../section.asp" -->

<!--<BGSOUND SRC="sounds/ringing.wav" LOOP=2> -->
<img src="images/phone.gif" height=48 width=48 vspace=1 hspace=10 border=0 align=left>
Looking for a SARK phone number?  Then you've come to the right place.  Here you will
find <a href="clients.asp">client</a> phone numbers,
<a href="employee.asp">alphabetical</a> listings,
and <a href="emp_pics.asp">pictures</a> of all your co-workers.  
<p>
Want to learn more about a fellow SARK?  Then check out their details page by
clicking on their name.  And don't forget to keep your own SARK profile
and subject matter expertise up to date.
<p>
Also included in this section are the locations and phone numbers of the other
SARK <a href="branches.asp">branches</a>.
<p>


<table border=0 width='100%' cellpadding=10 cellspacing=0><tr><td bgcolor=#ffffcc>
<font size=1 face="ms sans serif, arial, geneva">

<font color=navy><B>Can you name this Mystery SARK?</B></font><BR><BR>

<% 
'----------------------------
'	Execute Database Query   
'	returns random employee  
'   record id                
'----------------------------
set rs = DBQuery("select round(rand() * count(*), 0)'round' from profile")  

ll_randomsark = rs("round")

'---------------------------------
'	Loop through all rows         
'   in profile table until        
'   we reach the random employee  
'---------------------------------

set rs = DBQuery("select * from profile")

for ll_index = 1 to ll_randomsark - 1
	rs.movenext
next

if not (isnull(rs("college")) or trim(rs("college")) = "") then
	Response.Write("<b>" & "College:" & "<br></b>" & rs("college") & "<br><br>")
end if
'		if not (isnull(rs("hobbies_sports_interests")) or trim(rs("hobbies_sports_interests")) = "" ) then
'			Response.Write("<b>" & "Hobbies, Sports & Interests:" & "<br></b>" & rs("hobbies_sports_interests") & "<br><br>")
'		end if
if not (isnull(rs("dream_vacation")) or trim(rs("dream_vacation")) = "" ) then
	Response.Write("<b>" & "Dream Vacation:" & "<br></b>" & rs("dream_vacation") & "<br><br>")
end if
'		if not (isnull(rs("favorite_show_ad")) or trim(rs("favorite_show_ad")) = "" ) then
'			Response.Write("<b>" & "Favorite Show/Ad:" & "<br></b>" & rs("favorite_show_ad") & "<br><br>")
'		end if
'		if not (isnull(rs("person_inspired")) or trim(rs("person_inspired")) = "" ) then
'			Response.Write("<b>" & "Person Who Inspired me the Most:" & "<br></b>" & rs("person_inspired") & "<br><br>")
'		end if
if not (isnull(rs("favorite_motto")) or trim(rs("favorite_motto")) = "" ) then
	Response.Write("<b>" & "Favorite Motto:" & "<br></b>" & rs("favorite_motto") & "<br><br>")
end if
'		if not (isnull(rs("funniest_embarrassing")) or trim(rs("funniest_embarrassing")) = "" ) then
'			Response.Write("<b>" & "Funniest or Most Embarrassing Moment:" & "<br></b>" & rs("funniest_embarrassing") & "<br><br>")
'		end if 
if not (isnull(rs("lotto")) or trim(rs("lotto")) = "" ) then
	Response.Write("<b>" & "If I Won the Lotto...:" & "<br></b>" & rs("lotto") & "<br><br>")
end if
if not (isnull(rs("pet_peeve")) or trim(rs("pet_peeve")) = "" ) then
	Response.Write("<b>" & "Pet Peeve:" & "<br></b>" &  rs("pet_peeve") & "<br><br>")
end if
'		if not (isnull(rs("prized_possession")) or trim(rs("prized_possession")) = "" ) then
'			Response.Write("<b>" & "Prized Possession:" & "<br></b>" & rs("prized_possession") & "<br><br>")
'		end if
if not (isnull(rs("biggest_challenge")) or trim(rs("biggest_challenge")) = "" ) then
	Response.Write("<b>" & "Biggest Challenge:" & "<br></b>" & rs("biggest_challenge") & "<br><br>")
end if
Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a style='text-decoration:none' href='details.asp?EmpId=" & rs("employee_id") & "'><B>Who am I?</B>")
%>

</font></td></tr></table>


<!-- #include file="../../footer.asp" -->
