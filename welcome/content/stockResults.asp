<%response.buffer=true%>
<!--
Developer:    Kevin Dill
Date:         09/01/1998
Description:  Stock results.
-->


<html>

<head>
<title>Search Results</title>

<script language=javascript>
<!--
function submitInfo(){
	document.frmInfo.submit()
	}
// -->
</script>

</head>


<%
stock = trim(request("stock"))
response.write("<body bgcolor=white")
if stock <> "" then
	response.write(" onLoad='submitInfo()'>")
	response.write("<form name=frmInfo method=post action='http://www.news.com/Investor/Lookup/Processing/1,239,0~1~0~~~~~~~~,00.html'>")
	response.write("<input type=hidden name=ticker value='" & stock & "'>")
	if len(stock) < 6 then
		response.write("<input type=hidden name=type value='symbol'>")
	else
		response.write("<input type=hidden name=type value='name'>")
		end if
	response.write("</form>")
else
	response.write(">Please enter a stock name or symbol, and press Enter.")
	end if
%>

</body>
</html>
