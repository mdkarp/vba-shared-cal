VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub run()
   Dim NS As Outlook.NameSpace
   Dim newCalFolder As Outlook.Folder
   
   Dim curCal As Items
   
   Set NS = Application.GetNamespace("MAPI")
   ' default calendar
   'Cornell Tech - Academic & Campus Events <cutech-academic@cornell.edu>
   Set curCalFolder = GetFolderPath("\Calendar")
   
   
        Dim CalendarFolder As Outlook.Folder
        Dim objOwner As Outlook.Recipient
           
           'calendar you are copying to
        Set objOwner = NS.CreateRecipient("cutech-academic@cornell.edu")
        objOwner.Resolve
              
         
         If objOwner.Resolved Then
           'MsgBox objOwner.Name
         
          Set newCalFolder = NS.GetSharedDefaultFolder(objOwner, olFolderCalendar)
          Set Items = newCalFolder.Items
         End If
   
   Set curCal = Items
   'calendar you are copying to
   Set newCalFolder = GetFolderPath("mk2298@cornell.edu\Calendar\TestCal")
   Set NS = Nothing
   
   
   ' copy
   
   curCal.IncludeRecurrences = True
   
    Dim cAppt As AppointmentItem
    Dim moveCal As AppointmentItem
    Dim objAppointment As AppointmentItem
     
     On Error Resume Next
     
    For Each objAppointment In curCal
     
       'If Item.BusyStatus = olBusy Then
        
       Set cAppt = newCalFolder.Items.Add(olAppointmentItem)
        
       With cAppt
           .Subject = "Copied: " & objAppointment.Subject
           .Start = objAppointment.Start
           .Duration = objAppointment.Duration
           .Location = objAppointment.Location
           .Body = objAppointment.Body
       End With
        
       ' set the category after it's moved to force EAS to sync changes
        'Set moveCal = cAppt.Move(newCalFolder)
       ' moveCal.Categories = "moved"
        cAppt.Save
    
       'End If
    Next

   
End Sub

Function GetFolderPath(ByVal FolderPath As String) As Outlook.Folder
    Dim oFolder As Outlook.Folder
    Dim FoldersArray As Variant
    Dim i As Integer
         
    On Error GoTo GetFolderPath_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set oFolder = Application.Session.Folders.Item(FoldersArray(0))
    If Not oFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = oFolder.Folders
            Set oFolder = SubFolders.Item(FoldersArray(i))
            If oFolder Is Nothing Then
                Set GetFolderPath = Nothing
            End If
        Next
    End If
    'Return the oFolder
    Set GetFolderPath = oFolder
    Exit Function
         
GetFolderPath_Error:
    Set GetFolderPath = Nothing
    Exit Function
End Function


 
Sub CopyAttachments(objSourceItem, objTargetItem)
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set fldTemp = fso.GetSpecialFolder(2) ' TemporaryFolder
   strPath = fldTemp.Path & "\"
   For Each objAtt In objSourceItem.Attachments
      strFile = strPath & objAtt.FileName
      objAtt.SaveAsFile strFile
      objTargetItem.Attachments.Add strFile, , , objAtt.DisplayName
      fso.DeleteFile strFile
   Next
 
   Set fldTemp = Nothing
   Set fso = Nothing
End Sub

