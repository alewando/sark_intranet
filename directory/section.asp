<!-- #include file="../script.asp" -->
	// -------------------------------------------------------------
	// SERVER-SIDE SNIFF
	// -------------------------------------------------------------

	function BrowserData(sUA) {
		var iMSIE = sUA.indexOf("MSIE");
		this.userAgent = sUA;
		this.browser = (iMSIE > -1) ? "MSIE" : "Other";
		this.majorVer = parseInt(sUA.substring(iMSIE + 5, iMSIE + 6));
		this.getsSkills = ("MSIE" == this.browser && 5 <= this.majorVer);
	}
	var oBD = new BrowserData(new String(Request.ServerVariables("HTTP_USER_AGENT")));
</SCRIPT>
<%
'----------------------
'  SECTION PROPERTIES  
'----------------------
sectionTitle = "Directory"
sectionDir = "directory"

'SetWebMaster "cdolan"

buildNavItem "About this section", "about.asp"
buildNavItem "Alphabetical", "employee.asp"
buildNavItem "Alpha. with Pictures", "emp_pics.asp"
buildNavItem "Branch Info", "branches.asp"
buildNavItem "Client Lists", "clients.asp"
If hasRole("Sales") Then
%>

	// --------------------------
	// CLIENT-SIDE SNIFF
	// --------------------------

	function BrowserData(sUA) {
		var iMSIE = sUA.indexOf("MSIE");
		this.userAgent = sUA;
		this.browser = (-1 != iMSIE) ? "MSIE" : "Other";
		this.majorVer = parseInt(sUA.substring(iMSIE + 5, iMSIE + 6));
		this.getsSkills = ("MSIE" == this.browser && 5 <= this.majorVer);
	
	}
	var oBD = new BrowserData(navigator.userAgent);
//--></SCRIPT>