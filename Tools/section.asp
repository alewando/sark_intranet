<!-- #include file="../script.asp" -->
<%
'----------------------
'  SECTION PROPERTIES  
'----------------------
sectionTitle = "WebMaster Tools"
sectionDir = "tools"

buildNavItem "About", "default.asp"
if hasRole("WebMaster") then
 buildNavItem "Add Employee", "add_emp.asp"
 buildNavItem "Edit Employee", "chg_emp.asp"
 buildNavItem "Delete Employee", "del_emp.asp"
 buildNavItem "Edit Emp Security", "edit_user_role.asp"
 buildNavItem "Add Skill to S/I", "add_skill.asp"
 buildNavItem "Add Client", "add_client.asp"
 buildNavItem "Edit Client", "edit_client.asp"
 buildNavItem "Approve Postings", "approve_postings.asp"
 buildNavItem "Approve New Tech", "techapproval.asp"
 buildNavItem "Edit Tech", "edit_technology.asp"
 buildNavItem "Skills Inventory Report", "list_skillsInv.asp"
 buildNavItem "Edit Announcements", "updateannouncement.asp?Announcement_ID=1"
 buildNavItem "Add Exam", "add_exam.asp"
 buildNavItem "Edit Exam", "edit_exam.asp"
end if
if hasRole("WebMaster") or hasRole("TrainingCoordinator") then
 buildNavItem "Add Class", "add_class.asp"
 buildNavItem "Edit Class", "edit_class.asp"
end if
if hasRole("WebMaster") then
 buildNavItem "Add Certification", "add_cert.asp"
 buildNavItem "Edit Certification", "edit_cert.asp"
 buildNavItem "Add Solution Services Member", "add_ss.asp"
 buildNavItem "Edit Solution Services Member", "edit_ss.asp"
 buildNavItem "Add Branch", "add_branch.asp"
 buildNavItem "Edit Branch", "edit_branch.asp"
 buildNavItem "Edit Announcements", "updateannouncement.asp?Announcement_ID=1"
end if
%>
<!-- #include file="../header.asp" -->
