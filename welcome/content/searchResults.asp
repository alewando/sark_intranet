<!--
Developer:    Kevin Dill
Date:         08/17/1998
Description:  Search results page designed to work with several
              different search engines.
-->


<html>

<head>
<title>Search Results</title>

<%
'------------------------------
'  Get search text and engine  
'------------------------------
searchRoot = ""
search = trim(request.querystring("search"))
searchEngine = UCase(trim(request.querystring("searchEngine")))
if searchEngine = "" then searchEngine = request.cookies("SearchEngine")
%>

<script language="javascript">
<!--
function setCookie(key, val){
	var today = new Date()
	var expires = new Date()
	expires.setTime(today.getTime() + 1000*60*60*24*365)
	document.cookie = key + "=" + val + ";expires=" + expires.toGMTString()
	}

if ("<%=searchEngine%>" != "") setCookie("SearchEngine", "<%=searchEngine%>");
// -->
</script>

</head>

<body bgcolor=white>

<%
'------------------------------------------------
'  Build submit page depending on search engine  
'------------------------------------------------
if (searchEngine <> "") then
	select case searchEngine
%>

<%		case "ALTAVISTA" searchRoot="http://www.altavista.com"%>

<FORM name=frmSearch method=GET action="http://www.altavista.com/cgi-bin/query">
<INPUT type=hidden name=pg value=q>
<input type=hidden name=what value="web">
<input type=hidden name=k1 value="en">
<INPUT type=hidden name=q value="<%=search%>">
</form>

<%		case "EXCITE" searchRoot="http://www.excite.com"%>

<form name=frmSearch action="http://search.excite.com/search.gw" method=get>
<input type=hidden name=search value="<%=search%>"></form>

<%		case "HOTBOT" searchRoot="http://www.hotbot.com"%>

<form name=frmSearch action="http://www.hotbot.com/default.asp">
<INPUT type=hidden name="MT" value="<%=search%>">
<input type=hidden name="SM" value="MC">
</form>

<%		case "INFOSEEK" searchRoot="http://www.infoseek.com"%>

<form name=frmSearch action="http://www.infoseek.com/Titles" method=get>
<input type="hidden" name="qt" value="<%=search%>">
<input type="hidden" name="col" value="WW">
<input type=hidden name="sv" value="IS">
<input type=hidden name="lk" value="noframes">
<input type=hidden name="nh" value="10">
</form>

<%		case "LOOKSMART" searchRoot="http://www.looksmart.com"%>

<form name=frmSearch action="http://www.looksmart.com/r_search" method=GET>
<input type=hidden name=look value=>
<input name=key type=hidden value="<%=search%>">
<input type=hidden name=search value=0>
</form>

<%		case "LYCOS" searchRoot="http://www.lycos.com"%>

<form name=frmSearch action="http://www.lycos.com/cgi-bin/pursuit" method=get>
<input type=hidden name=matchmode value=and>
<input type=hidden name=cat value=lycos>
<input type=hidden name=query value="<%=search%>">
</form>

<%		case "NORTHERNLIGHT" searchRoot="http://www.northernlight.com"%>

<form name=frmSearch action="http://www.northernlight.com/nlquery.fcg" method=get>
<input type=hidden name=qr value="<%=search%>">
<input type=hidden name=si value="">
<input type=hidden name=cb value=0>
<input type=hidden name=cc value="">
<input type=hidden name=us value=025>
</form>

<%		case "WEBCRAWLER" searchRoot="http://www.webcrawler.com"%>

<form name=frmSearch action="http://www.webcrawler.com/cgi-bin/WebQuery" method=GET>
<input type=hidden name=searchText value="<%=search%>">
</form>

<%		case "YAHOO!" searchRoot="http://www.yahoo.com"%>

<form name=frmSearch action="http://search.yahoo.com/bin/search"><input type=hidden name=p value="<%=search%>"></form>

<%		end select%>

<script language="javascript">
<%if search="" then%>
window.location.href = "<%=searchRoot%>"
<%else%>
document.frmSearch.submit()
<%end if%>
</script>
<%	end if%>

</body>
</html>
