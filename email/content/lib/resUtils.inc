<%


Sub addReservation()
	Set objNewMsg = objCalendar.Messages.Add
	BindData(objNewMsg)
	objNewMsg.Update

	Set objCalendar = objOMSession.GetFolder(bstrCalendarID, NULL)
	Set objNewMsg = objCalendar.Messages.GetLast()

	Set objNewMsg2 = objNewMsg.CopyTo(objRoom.ID, objRoom.StoreID)
	objNewMsg2.Update

	rem Add custom field to use for tracking updates to Public Folder Reservation
	objNewMsg.Fields.Add "AppointmentLink", 8, objNewMsg2.ID
	objNewMsg2.Fields.Add "AppointmentLink", 8, objNewMsg.ID

	rem For Debugging
	'getMessages(objFolderRooms)
	'Response.write objNewMsg.Fields.Item("AppointmentLink") & "<br>"
	'Response.write objNewMsg2.Fields.Item("AppointmentLink") & "<br>"	
	'Response.write "DONE"
End Sub

Function getEndTime(bstrDuration,bstrStartDate,bstrStartTime)
	dDate = DateValue(bstrStartDate)
	dTime = TimeValue(bstrStartTime)
	
	iDuration = CSng(bstrDuration)
	
	iDay = Day(dDate)
	iMonth = Month(dDate)
	iYear = Year(dDate)
	iHour = Hour(dTime)
	iMinute = Minute(dTime)	
	if iDuration < 1 then
		iDuration = Cint(iDuration * 60)
		iMinute = iDuration + iMinute
		if iMinute > 60 then
			iMinute = iMinute - 60
			iHour = iHour + 1
			if iHour > 23 then
				iHour = iHour - 23
				iDay = iDay + 1
				' Need to continue this but you get the point
			end if
		end if
	else
		iHour = iHour + iDuration
			if iHour > 23 then
				iHour = iHour - 23
				iDay = iDay + 1
				' Need to continue this but you get the point
			end if		
	end if
	getEndTime = DateSerial(iYear,iMonth,iDay) + TimeSerial(iHour, iMinute, 0)
End Function



Function findSubFolders(objSubFolder,strFind)
	Set findSubFolders = Nothing
	if NOT objSubFolder.Folders.Count = 0 then	
		For Each Folder in objSubFolder.Folders
			if strFind = Folder.Name then
'				Response.write "<i>" & Folder.name & "</i><br>"						
				Set findSubFolders = Folder
				Exit Function
			end if
			findSubFolders Folder, strFind
		Next
	end if
End Function

Sub BindData(objNewMsg)

		objNewMsg.Importance=1 'Normal Importance
		objNewMsg.Type = "IPM.Appointment.Reservation"
		objNewMsg.TimeReceived = Now()
		objNewMsg.Subject = bstrPurpose 		'Request.Form("subject")
		objNewMsg.Location = objRoom.Name 	'Request.Form("location")
		objNewMsg.Text = ""		'Request.Form("message")

'		Sets Reminder values, default is True 	
'		    if Request.Form("Reminder") <> "" then
'			    objNewMsg.ReminderSet = True
'		    	objNewMsg.ReminderMinutesBeforeStart = CInt(Request.Form("reminderTime")) 
'	        else
'		objNewMsg.ReminderSet = False       
'	    	end if
        
		objNewMsg.BusyStatus = ActMsgBusy 	'Cint(Request.Form("showtimeas"))

		' START/END TIMES leave last in binding
' 	       fIsAllDay = BindStartEndTimes(1)

'		Used for Recurring Appointments			
'	  	if objNewMsg.IsRecurring = true Then
'   		Set objRecur = objNewMsg.GetRecurrencePattern
'	    	objRecur.PatternStartDate    = dtStart
'   	 	objRecur.StartTime    = dtStart
'    		objRecur.EndTime      = dtEnd
'  		Else   
			objNewMsg.StartTime	= dStartTime		'dtStart
			objNewMsg.EndTime	= dEndTime		'Date() + TimeSerial(23, 0, 0) 'dtEnd
' 			Must set StartTime and EndTime first
			objNewMsg.AllDayEvent = False 
'		end if

End Sub ' bindData

Sub getSubFoldersValues(objSubFolder)
	if NOT objSubFolder.Folders.Count = 0 then
		For Each Folder in objSubFolder.Folders
			bstrSubFoldersName = bstrSubFolderName + comma + Folder.Name 
			bstrSubFoldersID = bstrSubFolderID + comma + Folder.ID
			comma = ","
		Next
	end if
End Sub



Sub getMessages(objSubFolder)
	if NOT objSubFolder.Messages.Count = 0 then
		Response.write objSubFolder.Name
'		Response.write getTableHeader(arHeadtxt,arHeadsz)	
		For Each objMessage in objSubFolder.Messages
			getFields(objMessage)
'			Response.write objMessage.Type	
'			Response.write objMessage.StartTime
		Next		
	end if
End Sub

Sub getSubFolders(objSubFolder,boMessages)
	if NOT objSubFolder.Folders.Count = 0 then
		For Each Folder in objSubFolder.Folders
			Response.write "<i>" & Folder.name & "</i><br>"	
		Next
	end if
End Sub

Function getPublicFolder() 
	Set getPublicFoler = Nothing
	For Each InfoStore IN objOMSession.InfoStores
		If InfoStore.Name = "Public Folders" then
			tagPFRootID = &H66310102 ' PR_IPM_PUBLIC_FOLDERS_ENTRYID
			strPFRootID = InfoStore.Fields(tagPFRootID) ' entry ID
			Set objPFRoot = objOMSession.GetFolder(strPFRootID, InfoStore.ID)
'			For Each objPublicFolder In objPFRoot.Folders
'				Response.write objPublicFolder.Name & "<br>"
'			Next			
		end if
	Next
	Set getPublicFolder = objPFRoot
End Function

Sub getFields(objNewMsg)

			Response.write "CDP Message Class is " & objNewMsg.Class & "<br>"
			For each Field in objNewMsg.Fields
				if NOT IsArray(Field.Value) then
					Response.write "<i>&H" & Hex(Field.ID) & "</i>  "
					Response.write CStr(Field.Value) & "<br>"
								if NOT Err = 0 then
									Response.write "Error on FieldID: &H" & Hex(Field.ID) + "<br>"
									Err.Clear
								end if
				else
					For each arField in Field.Value
						Response.write "IS ARRAY<br>"
						Response.write "<i>&H" & Hex(Field.ID) & "</i>"						
						Response.write " " & CStr(arField) & "<br>"
						Response.write "ARRAY END"			
					Next
					Response.write "<br>"
				end if
			Next
			if objNewMsg.Class = 3 then
				For each Att in objNewMsg.Attachments
					Response.write "<br>Attachment<br>"
					getFields(Att)
				Next
			end if

End Sub

%>