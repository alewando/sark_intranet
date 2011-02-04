Dave Podnar created this file on 8/3/00

This file contains relevant information for the Solution Services repository 
and its use of the Posting Acceptor, which is a part of Site Server Express.

The best example information regarding the Posting Acceptor can be found in
the MSDN workshop at these URL's:
http://msdn.microsoft.com/workshop/server/asp/server052499.asp
http://msdn.microsoft.com/workshop/server/asp/server062899.asp

Additional information can be found by searching the MS Knowledge base using
keywords "Posting Acceptor".

Here are the high level steps which were performed:
- Install the posting acceptor if necessary.
- If necessary, create an IIS virtual directory for the scripts directory
- Create an explorer folder for the uploaded files
- Create an IIS virtual directory for the upload directory

The web server must have at least Basic Authentication set as an Authentication
Method for the Posting Acceptor to work. There is an MS KB entry about anonymous
posting, but I didn't try it.

To Install the posting acceptor do this:
1) Run the NT 4 Option Pack setup
2) Select Microsoft Site Server Express
3) Click Show Subcomponents.
4) Select Publishing - Posting Acceptor.
5) Click OK to continue the wizard.

Verify the cpshost.dll file is in the IIS scripts directory. If so, the 
Posting Acceptor is installed.

The production intranet site (www.sarkcincinnati.com/intranet) did not have
the IIS scripts directory as a virtual directory. If this is the case, make sure 
you add the scripts directory as a virtual directory. The only required IIS 
permission is execute. IIS Read and Write permission are not needed and probably 
should not be set, for security purposes. 

Create a folder named "Uploads" using file explorer. The "Uploads" name is not 
required by the Posting Acceptor but is coded into the scripts which were created
for the intranet repository. Admin, System and Everyone should have full control
directory permissions. On the production server these permissions were set:
Domain Admins - Full Control
Domain Users - Read
Everyone - Full
System - Full
Trickle\Intranet Full

Now, create a virtual directory named "Uploads" using IIS which points to where ever
the physical "Uploads" directory is. The IIS permissions for this directory should
only be "Read" and "Write". Script or Execute should NOT be set! This would allow
someone to upload a script and then execute it.

After this step you should try a test posting to see if everything is working. There
is a TestUpload.asp script in solutionservices/content which can be used for
initial testing. In the event, you chose a different folder name than "Uploads",
make sure to revise the TestUpload.asp script.

Due to different versions of the cpshost.dll, different errors may be encountered.
Good Luck working through these if they occur. I did not find any information
which gave instructions on what to do if errors are encountered. 

While on the subject of different versions of cpshost.dll. You'll notice that 
version 6.1.27.0 is used on both test and production. However, there is also 
a newer version of the file (don't know the specific version) which was renamed 
to cpshost.dll.new since it kept writing the text "Repost OK" after the
Add_Document.asp script wrote "Successfully uploaded file xyz". I was unable
to turn this off, so we're using the older version.

As far as the repository application specifics, it is very simple. A file system
based repository was chosen instead of database tables for these reasons:
- File system was straight forward and simple
- Database tables have BLOB limitations
- Future file searching functionality should be easy

The ASP FileSystemObject is used to navigate through the directories and perform
physical folder create, delete and document move and delete. I encountered problems
with folder names using special characters, so we restrict to only Alpha, Numeric
and spaces.

The top level directories of the repository must match the Solution Services
practice areas in the Tech_Area.Tech_Area columns. In the event a top level 
directory is added which isn't in Tech_Area.Tech_Area, only the WebMaster will
be able to create sub folders and delete documents.
