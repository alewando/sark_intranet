<%
'----------------------------------------------------------
' Common Constants and Variables for the Repository section
'----------------------------------------------------------

'Directory names, don't terminate with "/" or "\"
const UPLOAD_DIR = "Uploads" 'where documents are uploaded to
const REPOSITORY_ROOT_DIR = "repository" 'Root (top level) of the repository

'Script names
const VIEW_REPOSITORY_SCRIPT = "View_Repository.asp"
const ADD_DOCUMENT_SCRIPT = "Add_Document.asp"
const NOT_PRIVILEGED_SCRIPT = "Not_Privileged.asp"

'Values for Hidden ACTION field of Forms
const ACTION_DELETE = "DELETE"
const ACTION_UPLOAD = "UPLOAD"

'Values for Permission error messages
const NO_PERMISSION_ERROR = 0
const NO_DELETE_DOCUMENT_PERMISSION = 1
const NO_DELETE_FOLDER_PERMISSION = 2

%>