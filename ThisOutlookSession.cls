VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim WithEvents curCal As Items
Attribute curCal.VB_VarHelpID = -1
Dim newCalFolder As Outlook.Folder
  
Private Sub Application_Startup()
   Dim NS As Outlook.NameSpace
   Set NS = Application.GetNamespace("MAPI")
   ' default calendar
   Set curCal = NS.GetDefaultFolder(olFolderCalendar).Items
   'calendar you are copying to
   Set newCalFolder = GetFolderPath("me@outlook.com\Calendar")
   Set NS = Nothing
End Sub
  
Private Sub curCal_ItemAdd(ByVal Item As Object)
Dim cAppt As AppointmentItem
Dim moveCal As AppointmentItem
 
If Item.BusyStatus = olBusy Then
 
Set cAppt = Application.CreateItem(olAppointmentItem)
 
With cAppt
    .Subject = "Copied: " & Item.Subject
    .Start = Item.Start
    .Duration = Item.Duration
    .Location = Item.Location
    .Body = Item.Body
End With
 
' set the category after it's moved to force EAS to sync changes
 Set moveCal = cAppt.Move(newCalFolder)
 moveCal.Categories = "moved"
 moveCal.Save
 
    End If
 End Sub
 
 
Private Sub curCal_ItemChange(ByVal Item As Object)
Dim cAppt As AppointmentItem
Dim objAppointment As AppointmentItem
Dim strStart, strSubject As String
     
 On Error Resume Next
     
strSubject = "Copied: " & Item.Subject
strStart = Item.Start
 
For Each objAppointment In newCalFolder.Items
 If objAppointment.Subject = strSubject And objAppointment.Start = strStart Then
         Set cAppt = objAppointment
   End If
 Next
  
With cAppt
    .Subject = "Copied: " & Item.Subject
    .Start = Item.Start
    .Duration = Item.Duration
    .Location = Item.Location
    .Body = Item.Body
    .Save
End With
  
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


