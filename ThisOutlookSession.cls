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
   Dim readCalFolder As Outlook.Folder
   
   Dim newCal As Items
   Dim curCal As Items
   
   Set NS = Application.GetNamespace("MAPI")
   
   
   'calendar you are copying to
   Set newCalFolder = GetFolderPath("mk2298@cornell.edu\Calendar\TestCal")
   
   ' dump calendar to replace everything
   Dim delAppt As AppointmentItem
   
   Set newCal = newCalFolder.Items
   
   newCal.Sort "[Start]"
   newCal.IncludeRecurrences = True
   
   'On Error Resume Next
     
    For Each delAppt In newCal
        delAppt.Delete
    Next
   
   
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
         
          Set readCalFolder = NS.GetSharedDefaultFolder(objOwner, olFolderCalendar)
          Set Items = readCalFolder.Items
         End If
   
   Set curCal = Items

   Set NS = Nothing
   
   
   ' copy
   curCal.Sort "[Start]"
   curCal.IncludeRecurrences = True
   
    Dim cAppt As AppointmentItem
    Dim moveCal As AppointmentItem
    Dim objAppointment As AppointmentItem
     
     On Error Resume Next
     
    For Each objAppointment In curCal
     
       'If Item.BusyStatus = olBusy Then
        
       Set cAppt = newCalFolder.Items.Add(olAppointmentItem)
        
       With cAppt
           .Subject = objAppointment.Subject
           .Start = objAppointment.Start
           .Duration = objAppointment.Duration
           .Location = objAppointment.Location
           .Body = objAppointment.Body
           .Categories = objAppointment.Categories
       End With
        
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

