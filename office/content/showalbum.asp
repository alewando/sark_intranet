<!-- #include file="../section.asp" -->
<HTML>
<HEAD>
  <TITLE>Software Architects Photo Album</TITLE>
  <STYLE>  
  
  .primaryColor
    {
    font-family: Arial;
    font-size: 12pt;
    background-color: #ffffff;
    color: #000000;
    }
  .secondaryColor
    {
    font-family: Arial;
    font-size: 12pt;
    background-color: #990000;
    color: white;
    text-align: center;
    }
  .highlightColor
    {
    font-family: Arial;
    font-size: 12pt;
    background-color: #cccccc;
    color: #ffffff;
    }
  </STYLE>
</HEAD>

<BODY bgcolor="#ffffff">
<CENTER>
  <TABLE class="secondaryColor" width="450" border="0" cellspacing="4">
  <TR>
    <TD>
      <B>
      <%= mid(request.querystring("id"),3) %>
      </B>
    </TD>
  </TR>
  <TR>
    <TD>
      <TABLE class="primaryColor" width="100%">
      <TR>
        <TD class="highlightColor" height="4" colspan="2"></TD>
      </TR>
<%
    dim fso, oFolder, oFiles, iCounter
    dim iStartCount, iTotalCount
    iStartCount = request.querystring("start")
      if iStartCount = "" then iStartCount = 1
    set fso = CreateObject("Scripting.FileSystemObject")
    set oFolder = fso.GetFolder(server.mappath(".") & "/../album/" & request.querystring("id"))
    set oFiles = oFolder.files
    iCounter = 0
    iTotalCount = oFiles.count
    for each oFile in oFiles
      iCounter = iCounter + 1
      if (iCounter > (iStartCount - 1)) and (iCounter < (iStartCount + 10)) then
        if iCounter mod 2 = 1 then
          response.write("      <TR>" & vbCRLF)
          response.write("        <TD align=""center"">" & vbCRLF)
          response.write("          <A href=""showpicture.asp?id=" & server.urlencode(request.querystring("id")) & "/" & oFile.name & """ target=""new"">" & vbCRLF)
          response.write("          <IMG border=""0"" src=""../album/" & request.querystring("id") & "/" & oFile.name & """ height=""100"">" & vbCRLF)
          response.write("          </A>" & vbCRLF)
          response.write("          <BR>" & vbCRLF)
          response.write("          <FONT style=""font-size:10""><B>" & left(oFile.name, len(oFile.name) - 4) & "</B></FONT>" & vbCRLF)
          response.write("        </TD>" & vbCRLF)
        else
          response.write("        <TD align=""center"">" & vbCRLF)
          response.write("          <A href=""showpicture.asp?id=" & request.querystring("id") & "/" & oFile.name & """ target=""new"">" & vbCRLF)
          response.write("          <IMG border=""0"" src=""../album/" & request.querystring("id") & "/" & oFile.name & """ height=""100"">" & vbCRLF)
          response.write("          </A>" & vbCRLF)
          response.write("          <BR>" & vbCRLF)
          response.write("          <FONT style=""font-size:10""><B>" & left(oFile.name, len(oFile.name) - 4) & "</B></FONT>" & vbCRLF)
          response.write("        </TD>" & vbCRLF)
          response.write("      </TR>" & vbCRLF)
          response.write("      <TR>" & vbCRLF)
          response.write("        <TD class=""highlightColor"" height=""4"" colspan=""4""></TD>" & vbCRLF)
          response.write("      </TR>" & vbCRLF)
        end if
      end if
    next
    if iCounter mod 1 = 0 then
      response.write("        <TD align=""center"">" & vbCRLF)
      response.write("        </TD>" & vbCRLF)
      response.write("      </TR>" & vbCRLF)
    end if
    set oFiles = nothing
    set oFolder = nothing
    set fso = nothing

    response.write("      <TR>" & vbCRLF)
    response.write("        <TD align=""center"">" & vbCRLF)
    if iStartCount > 10 then 
      response.write("          <A href=""showAlbum.asp?id=" & server.urlencode(request.querystring("id")) & "&start=" & iStartCount - 10 & """>&lt;&lt;BACK&lt;&lt;</A>" & vbCRLF)
    end if
    response.write("        </TD>" & vbCRLF)
    response.write("        <TD align=""center"">" & vbCRLF)
    if (iStartCount + 10) < iTotalCount then 
      response.write("          <A href=""showAlbum.asp?id=" & server.urlencode(request.querystring("id")) & "&start=" & iStartCount + 10 & """>&gt;&gt;NEXT&gt;&gt;</A>" & vbCRLF)
    end if
    response.write("        </TD>" & vbCRLF)
    response.write("      </TR>" & vbCRLF)
%>
      </TABLE>
    </TD>
  </TR>
  </TABLE>
  <BR>
  <A href="album.asp">Return to Main Page</A>
</CENTER>
<!-- #include file="../../footer.asp" -->
</BODY>
</HTML>
