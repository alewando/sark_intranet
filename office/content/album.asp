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
<BODY bgcolor="#ffffff" background="/common/images/logo_bg_white.jpg">
<CENTER>
  <TABLE class="secondaryColor" width="100%" border="0" cellspacing="4">
  <TR>
    <TD>
      <B>
      The Software Architects Photo Album
      </B>
    </TD>
  </TR>
  <TR>
    <TD>
      <TABLE class="primaryColor" width="100%">
      <TR>
        <TD class="highlightColor" height="4"></TD>
      </TR>
      <TR>
        <TD align="center">
          <!-- PHOTO ALBUM LIST -->
<%

    dim fso, oFolder, oSubFolders, oSubFolder, sFolderName
    set fso = CreateObject("Scripting.FileSystemObject")
    set oFolder = fso.GetFolder(server.mappath(".") & "\..\album")
    set oSubFolders = oFolder.SubFolders
    for each oSubFolder in oSubFolders
      for iCounter = len(oSubFolder) to 1 step -1
        if mid(oSubFolder, iCounter, 1) = "\" then
          sFolderName = mid(oSubFolder, iCounter + 1)
          exit for
        end if
      next
      response.write("<A href=""showAlbum.asp?id=" & server.urlencode(sFolderName) & """>" & mid(sFolderName, 3) & "</A>" & vbCRLF)
      response.write("<BR>" & vbCRLF)
    next
    set oSubFolders = nothing
    set oFolder = nothing
    set fso = nothing

%>
          <!-- PHOTO ALBUM LIST END -->
        </TD>
      </TR>
      </TABLE>
    </TD>
  </TR>
  </TABLE>
</CENTER>

</BODY>
</HTML>


<!-- #include file="../../footer.asp" -->
