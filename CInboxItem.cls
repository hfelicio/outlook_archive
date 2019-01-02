VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInboxItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' private variables  & const
Private mlArchiveYear       As Long
Private mlDaysToKeep        As Long
Private mlItemsMoved        As Long
Private mlTotalItems        As Long
Private mbIncludeUnread     As Boolean
Private mbIncludeSubFolders As Boolean
Private mbSkipEncrypted     As Boolean
Private mobjArchiveFolder   As Outlook.folder

Private Const mc_ModuleName             As String = "CInboxItem"
Private Const mc_SkipString             As String = "NOT_FOR_ARCHIVE"
Private Const mc_MessageClassEncrypted  As String = "IPM.Note.SMIME"

Public Event OnInfo(ByVal Message As String)
Public Event OnError(ByVal Message As String)
Public Event ItemMoved(ByVal MovedItems As Long, ByVal TotalItems As Long)
Public Event ItemMovedText(ByVal ItemText As String)
Public Event OnPathChanged(ByVal FolderPath As String)
Public Event ArchiveFolderChange(ByVal FolderPath As String)

' public read/write properties

Public Property Get ArchiveYear() As Long
    ArchiveYear = mlArchiveYear
End Property

Public Property Let ArchiveYear(ByVal value As Long)
    mlArchiveYear = value
End Property

Public Property Get DaysToKeep() As Long
    DaysToKeep = mlDaysToKeep
End Property

Public Property Let DaysToKeep(ByVal value As Long)
    mlDaysToKeep = value
End Property

Public Property Get ItemsMoved() As Long
    ItemsMoved = mlItemsMoved
End Property

Public Property Get TotalItems() As Long
    TotalItems = mlTotalItems
End Property

Public Property Get IncludeUnread() As Boolean
    IncludeUnread = mbIncludeUnread
End Property

Public Property Let IncludeUnread(ByVal value As Boolean)
    mbIncludeUnread = value
End Property

Public Property Get IncludeSubFolders() As Boolean
    IncludeSubFolders = mbIncludeSubFolders
End Property

Public Property Let IncludeSubFolders(ByVal value As Boolean)
    mbIncludeSubFolders = value
End Property

Public Property Get SkipEncrypted() As Boolean
    SkipEncrypted = mbSkipEncrypted
End Property

Public Property Let SkipEncrypted(ByVal value As Boolean)
    mbSkipEncrypted = value
End Property

' public read-only properties




' private procedures

Private Sub Class_Initialize()
    PDebug mc_ModuleName & "->Class_Initialize"
    mlArchiveYear = 0
    mlDaysToKeep = 10000
End Sub

Private Sub Class_Terminate()
    PDebug mc_ModuleName & "->Class_Terminate"
End Sub

Public Sub StartArchiving()
    PDebug mc_ModuleName & "->StartArchiving->Start "
    On Error GoTo errHandler
    mlTotalItems = 0
    mlItemsMoved = 0
    ProcessFolderRec GetSourceFolder
cleanup:
    PDebug mc_ModuleName & "->StartArchiving->Exit "
    Exit Sub
errHandler:
    RaiseEvent OnError(PError(mc_ModuleName, "StartArchiving", Err.Number, Err.Description))
    Resume Next
    GoTo cleanup
End Sub

Private Sub ProcessFolderRec(ByVal oFolder As Outlook.MAPIFolder)
    PDebug mc_ModuleName & "->ProcessFolderRec->Start " & oFolder.Name
    Dim lngCount           As Long
    Dim lTotalMoved As Long
    On Error GoTo errHandler
    
    If (Not IsObject(oFolder) Or oFolder Is Nothing) Then
        RaiseError mc_ModuleName, "ProcessFolderRec", "oFolder variable is nothing."
    End If
    
    If oFolder.Description = mc_SkipString Then
        If Me.DaysToKeep > 0 Then Exit Sub
    End If
    
    If IsOperationAborted Then Exit Sub
    
    RaiseEvent OnInfo("Folder " & oFolder.Name & " contains " & oFolder.items.Count & " item(s)")
    
    'main process call
    If oFolder.items.Count > 0 Then
        lTotalMoved = MoveFolderItems(oFolder, Me.SkipEncrypted)
        RaiseEvent OnPathChanged(oFolder.Name & " (" & lTotalMoved & ")")
    End If
    
    If Not mbIncludeSubFolders Then Exit Sub
    For lngCount = oFolder.Folders.Count To 1 Step -1
        DoEvents
        If IsOperationAborted Then Exit For
        ProcessFolderRec oFolder.Folders.Item(lngCount)
    Next
    
cleanup:
    PDebug mc_ModuleName & "->ProcessFolderRec" & "->Exit"
    Exit Sub
errHandler:
    RaiseEvent OnError(PError(mc_ModuleName, "ProcessFolderRec", Err.Number, Err.Description))
    Resume Next
    GoTo cleanup
End Sub

Private Function MoveFolderItems( _
        ByRef oSourceFolder As Outlook.folder, _
        Optional ByVal SkipEncrypted As Boolean = True) As Long
    
    PDebug mc_ModuleName & " -> MoveFolderItems -> Start " & oSourceFolder.Name
    
    Dim oArchiveFolder              As Outlook.folder
    Dim oArchiveFolderVar           As Outlook.folder
    Dim objVariant                  As Variant
    Dim blnMoveConditionEncrypted   As Boolean
    Dim blnMoveConditionUnread      As Boolean
    Dim blnMoveConditionAge         As Boolean
    Dim lItemYear                   As Long
    Dim lngCount                    As Long
    Dim lItemsCount                 As Long
    Dim objMail                     As MailItem
    Dim dteToday                    As Date
    Dim lEmailAge                   As Long
    
    On Error GoTo errHandler
    
    If IsOperationAborted Then Exit Function
    
    If (Not IsObject(oSourceFolder) Or oSourceFolder Is Nothing) Then
        RaiseError mc_ModuleName, "MoveFolderItems", "oSourceFolder variable is nothing."
    End If
        
    Set oArchiveFolder = GetArchiveFolder(mlArchiveYear)
    RaiseEvent ArchiveFolderChange(oArchiveFolder.FolderPath)
    
    If (Not IsObject(oArchiveFolder) Or oArchiveFolder Is Nothing) Then
        RaiseError mc_ModuleName, "MoveFolderItems", "oArchiveFolder variable is nothing."
    End If
    
    lItemsCount = oSourceFolder.items.Count
    mlTotalItems = mlTotalItems + lItemsCount
    mlItemsMoved = 0
    
    'PDebug mc_ModuleName & "-> MoveFolderItems -> lItemsCount " & lItemsCount
    'PDebug mc_ModuleName & "-> MoveFolderItems -> mlTotalItems " & mlTotalItems
    'PDebug mc_ModuleName & "-> MoveFolderItems -> SkipEncrypted " & SkipEncrypted
    DoEvents
    
    dteToday = RemoveTimeFromDate(Now)
    
    'On Error Resume Next
    For lngCount = lItemsCount To 1 Step -1
        If IsOperationAborted Then Exit Function
        DoEvents
        Set objVariant = oSourceFolder.items.Item(lngCount)
        
        lEmailAge = DateDiff("d", objVariant.ReceivedTime, dteToday)
        blnMoveConditionAge = (lEmailAge >= mlDaysToKeep)
        If blnMoveConditionAge Or (Me.DaysToKeep = 0) Then
            Select Case objVariant.Class
                Case olMail, _
                    olAppointment, _
                    olMeeting, _
                    olMeetingRequest, _
                    olMeetingCancellation, _
                    olMeetingResponseNegative, _
                    olMeetingResponsePositive, _
                    olMeetingResponseTentative, _
                    olMeetingForwardNotification
                    
                    'Set objMail = objVariant
                    blnMoveConditionEncrypted = (Not (objVariant.MessageClass = mc_MessageClassEncrypted)) Or ((SkipEncrypted = False) And (objVariant.MessageClass = mc_MessageClassEncrypted))
                    If blnMoveConditionEncrypted Or (Me.DaysToKeep = 0) Then
                        lItemYear = CLng(Year(objVariant.ReceivedTime))
                            
                        blnMoveConditionUnread = (Not objVariant.UnRead) Or (objVariant.UnRead And (mbIncludeUnread = True))
                        If blnMoveConditionUnread Or (Me.DaysToKeep = 0) Then
                            If lItemYear = mlArchiveYear Then
                                objVariant.Move oArchiveFolder.Folders.Item(CStr(Month(objVariant.ReceivedTime)))
                                RaiseEvent ItemMovedText("[" & objVariant.Sender & "] " & objVariant.Subject)
                                mlItemsMoved = mlItemsMoved + 1
                            Else
                                Set oArchiveFolderVar = GetArchiveFolder(lItemYear)
                                If (Not IsObject(oArchiveFolderVar) Or oArchiveFolderVar Is Nothing) Then
                                    RaiseEvent OnError(PError(mc_ModuleName, "MoveFolderItems", Err.Number, TypeName(objVariant) & "-" & "oArchiveFolderVar variable is nothing."))
                                End If
                                RaiseEvent ArchiveFolderChange(oArchiveFolderVar.FolderPath)
                                objVariant.Move oArchiveFolderVar.Folders.Item(CStr(Month(objVariant.ReceivedTime)))
                                mlItemsMoved = mlItemsMoved + 1
                            End If
                        Else
                            PDebug mc_ModuleName & "-> MoveFolderItems -> Move unread? " & blnMoveConditionUnread & ";; Email is unread? " & objVariant.UnRead
                        End If
                    Else
                        PDebug mc_ModuleName & "-> MoveFolderItems -> Move encrypted emails? " & blnMoveConditionEncrypted & ";; Email is encrypted? " & objVariant.MessageClass
                    End If
                Case Else
                    PDebug mc_ModuleName & "-> MoveFolderItems -> Mail MessageClass: " & objVariant.MessageClass & ";; Mail Class: " & objVariant.Class
            End Select
        Else
            PDebug mc_ModuleName & "-> MoveFolderItems -> Mail Age (days): " & lEmailAge
        End If
        DoEvents
        RaiseEvent ItemMoved(mlItemsMoved, mlTotalItems - lngCount)
    Next lngCount
    MoveFolderItems = mlItemsMoved
    
cleanup:
    Set oArchiveFolder = Nothing
    Set oArchiveFolderVar = Nothing
    Set objVariant = Nothing
    Set objMail = Nothing
    PDebug mc_ModuleName & "->MoveFolderItems" & "->Exit"
    Exit Function
    
errHandler:
    Select Case Err.Description
        Case "Your Digital ID name cannot be found by the underlying security system." 'ignorable errors
            PDebug mc_ModuleName & "-> MoveFolderItems -> skipped MailItem because of error: Your Digital ID name cannot be found...)"
            If mbSkipEncrypted = False Then
                AbortOperation = True
                Err.Raise 12345, mc_ModuleName & "-> MoveFolderItems", "Please open an encrypted email before unchecking the SKIP ENCRYPTED emails checkbox."
            Else
                Resume Next
            End If
        Case Else
            RaiseEvent OnError(PError(mc_ModuleName, "MoveFolderItems", Err.Number, Err.Description))
            Resume Next
    End Select
    GoTo cleanup
End Function

' public procedures or functions

'Public Sub InitiateProperties(ByVal pArchiveYear As Long, _
'                              Optional ByVal pDaysToKeep As Long = glDefaultDaysToKeep, _
'                              Optional ByVal pIncludeUnread As Boolean = True, _
'                              Optional ByVal pSkipEncrypted As Boolean = True, _
'                              Optional ByVal pIncludeSubFolders As Boolean = True)
'    mlArchiveYear = pArchiveYear
'    mlDaysToKeep = pDaysToKeep
'    mbIncludeUnread = pIncludeUnread
'    mbSkipEncrypted = pSkipEncrypted
'    mbIncludeSubFolders = pIncludeSubFolders
'End Sub

Public Function GetSourceFolder() As Outlook.folder
    Set GetSourceFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
End Function

Public Function GetArchiveFolder(ByVal ArchiveYear As Long) As Outlook.folder
    On Error GoTo errHandler
    If (ArchiveYear = 0) Then
        RaiseError mc_ModuleName, "GetArchiveFolder", "ArchiveYear variable is 0."
    Else
        If mlArchiveYear = ArchiveYear Then
            If (Not IsObject(mobjArchiveFolder) Or mobjArchiveFolder Is Nothing) Then
                Set mobjArchiveFolder = Main.StoreFactory(ArchiveYear).GetStore.GetFolderByCategory(gsINBOX_PT)
            End If
            Set GetArchiveFolder = mobjArchiveFolder
        Else
            Set GetArchiveFolder = Main.StoreFactory(ArchiveYear).GetStore.GetFolderByCategory(gsINBOX_PT)
        End If
    End If
cleanup:
    Exit Function
errHandler:
    RaiseEvent OnError(PError(mc_ModuleName, "GetArchiveFolder", Err.Number, Err.Description))
    Resume Next
    GoTo cleanup
End Function

''NEW 18-06-2018
'Public Sub ArchiveSingleEmail()
'    PDebug mc_ModuleName & "->ArchiveSingleEmail->Start "
'    On Error GoTo errHandler
'
'    Dim oArchiveFolder  As Outlook.folder
'    Dim objVariant           As Variant
'    Set objVariant = Application.ActiveInspector.CurrentItem
'    Set oArchiveFolder = GetArchiveFolder(CLng(Year(objVariant.ReceivedTime)))
'    objVariant.Move oArchiveFolder.Folders.Item(CStr(Month(objVariant.ReceivedTime)))
'
'cleanup:
'    PDebug mc_ModuleName & "->ArchiveSingleEmail->Exit "
'    Exit Sub
'errHandler:
'    RaiseEvent OnError(PError(mc_ModuleName, "ArchiveSingleEmail", Err.Number, Err.Description))
'    Resume Next
'    GoTo cleanup
'End Sub

'-------------------------------------------------------------------------------------------------
'Private Sub AddFolderPathToList(ByVal oFolder As Outlook.MAPIFolder, ByRef arrFolders() As String)
'    Dim lngCount    As Long
'    Dim lTotalMoved As Long
'    Dim lngPos      As Long
'    On Error GoTo errHandler
'    If (Not IsObject(oFolder) Or oFolder Is Nothing) Then
'        RaiseError mc_ModuleName, "AddFolderPathToList", "oFolder variable is nothing."
'    End If
'
'    If oFolder.Description <> mc_SkipString Then
'        If IsInitializedArray(arrFolders) Then
'            lngPos = UBound(arrFolders) + 1
'            ReDim Preserve arrFolders(lngPos)
'        Else
'            lngPos = 0
'            ReDim arrFolders(lngPos)
'        End If
'
'        arrFolders(lngPos) = oFolder.FolderPath & " (" & oFolder.items.Count & ")"
'        For lngCount = oFolder.Folders.Count To 1 Step -1
'            DoEvents
'            AddFolderPathToList oFolder.Folders.Item(lngCount), arrFolders
'        Next
'    End If
'cleanup:
'    Exit Sub
'errHandler:
'    RaiseEvent OnError(PError(mc_ModuleName, "AddFolderPathToList", Err.Number, Err.Description))
'    Resume Next
'    GoTo cleanup
'End Sub

'Public Function GetArrFolders() As String()
'    Dim oFolder       As Outlook.folder
'    Dim lngCount      As Long
'    Dim lItemsCount   As Long
'    Dim objVariant         As Variant
'    Dim arrFolders()  As String
'    On Error GoTo errHandler
'    Set oFolder = GetSourceFolder
'    If (Not IsObject(oFolder) Or oFolder Is Nothing) Then
'        RaiseError mc_ModuleName, "GetArrFolders", "oFolder variable is nothing."
'    End If
'    Erase arrFolders
'    AddFolderPathToList oFolder, arrFolders
'    GetArrFolders = arrFolders
'cleanup:
'    Exit Function
'errHandler:
'    RaiseEvent OnError(PError(mc_ModuleName, "GetArrFolders", Err.Number, Err.Description))
'    Resume Next
'    GoTo cleanup
'End Function



