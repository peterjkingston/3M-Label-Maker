VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProgram_MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim Program As New CProgram_MAIN
Private M_DataStore As Collection
Private WithEvents mManualEntry As WinManualEntry, WithEvents mPrintDialog As WinPrintDialogue, mPreferredNameEntry As WinPreferredEntry
Attribute mManualEntry.VB_VarHelpID = -1
Attribute mPrintDialog.VB_VarHelpID = -1
Private WithEvents mSplash As WinSplash, WithEvents mMainNav As WinMainNav, WithEvents mLogNav As WinLogNav
Attribute mSplash.VB_VarHelpID = -1
Attribute mMainNav.VB_VarHelpID = -1
Attribute mLogNav.VB_VarHelpID = -1
Private WithEvents mSearch As WinUserQuery, WithEvents mSettings As WinSettingsMain, WithEvents mPreferredNames As WinPreferredNames
Attribute mSearch.VB_VarHelpID = -1
Attribute mSettings.VB_VarHelpID = -1
Attribute mPreferredNames.VB_VarHelpID = -1
Private WithEvents mPreferredEntry As WinPreferredEntry, WithEvents mFileManager As WinFileManager, WithEvents mQueryManager As WinQueryManager
Attribute mPreferredEntry.VB_VarHelpID = -1
Attribute mFileManager.VB_VarHelpID = -1
Attribute mQueryManager.VB_VarHelpID = -1
Private WithEvents mPrintPreview As WinPrintPreview
Attribute mPrintPreview.VB_VarHelpID = -1

Public Sub Main()
    
    Set M_DataStore = New Collection

    Excel.Application.Visible = False
    Set mSplash = New WinSplash
    mSplash.Show
    Set mSplash = Nothing
    
    Set mMainNav = New WinMainNav
    mMainNav.Show

End Sub
Public Property Get StoreFound() As Boolean
    StoreFound = Not CBool(M_DataStore.count = 0)
End Property
Public Property Get StoreObject(byName As String) As Object
    On Error GoTo ErrorHandler
    Set StoreObject = M_DataStore(byName)
Exit Property
ErrorHandler:
    Set StoreObject = Nothing
End Property
Public Sub SplashActions()
    Dim time As Double, waitSplash As Double
    waitSplash = 3
    time = Timer
    Initializer.Run M_DataStore
    If (Timer - time) < waitSplash Then: PKLib.Wait waitSplash - (Timer - time)
    mSplash.Hide
End Sub

Private Sub Class_Terminate()
    Excel.Application.Visible = True
End Sub

Private Sub mLogNav_RequestViewManualEntry(EditMode As Boolean, dataArgs() As String, index As Integer)
    If mManualEntry Is Nothing Then: Set mManualEntry = New WinManualEntry
    mManualEntry.EditMode = EditMode
    mManualEntry.TextBoxSO = dataArgs(0)
    mManualEntry.TextBoxCName = dataArgs(1)
    mManualEntry.TextBoxPO = dataArgs(2)
    mManualEntry.TextBoxCSRep = dataArgs(3)
    mManualEntry.EditIndex = index
    mLogNav.ListenTo mManualEntry
    
    mManualEntry.Show
End Sub

Private Sub mLogNav_RequestViewPrintPreview()
    If mPrintPreview Is Nothing Then: Set mPrintPreview = New WinPrintPreview
    mPrintPreview.Show
End Sub

Private Sub mLogNav_Terminated()
    Set mLogNav = Nothing
End Sub

Private Sub mMainNav_Terminated()
    Set mMainNav = Nothing
End Sub

Private Sub mManualEntry_Terminated()
    Set mManualEntry = Nothing
End Sub

Private Sub mPreferredEntry_Terminated()
    Set mPreferredEntry = Nothing
End Sub

Private Sub mPreferredNames_PreferredEntryViewRequested(EditMode As Boolean, dataArgs() As String)
    If mPreferredEntry Is Nothing Then: Set mPreferredEntry = New WinPreferredEntry
    
    mPreferredEntry.EditMode = EditMode
    mPreferredEntry.TextBoxSoldTo = dataArgs(0)
    mPreferredEntry.TextBoxPreferredName = dataArgs(1)
    
    mPreferredEntry.Show
End Sub

Private Sub mPreferredNames_Terminated()
    Set mPreferredNames = Nothing
End Sub


Private Sub mPrintDialog_OnQuit()
    Quit
End Sub
Private Sub Quit()
    If Not mManualEntry Is Nothing Then: Set mManualEntry = Nothing
    If Not mPrintDialog Is Nothing Then: Set mPrintDialog = Nothing
    If Not mPreferredNameEntry Is Nothing Then: Set mPreferredNameEntry = Nothing
    If Not mSplash Is Nothing Then: Set mSplash = Nothing
    If Not mMainNav Is Nothing Then: Set mMainNav = Nothing
    If Not mLogNav Is Nothing Then: Set mLogNav = Nothing
    If Not mSearch Is Nothing Then: Set mSearch = Nothing
    If Not mSettings Is Nothing Then: Set mSettings = Nothing
    If Not mPreferredNames Is Nothing Then: Set mPreferredNames = Nothing
    If Not mPreferredEntry Is Nothing Then: Set mManualEntry = Nothing
    If Not mFileManager Is Nothing Then: Set mFileManager = Nothing
    If Not mQueryManager Is Nothing Then: Set mQueryManager = Nothing
    If Not mPrintPreview Is Nothing Then: Set mPrintPreview = Nothing
End Sub

Private Sub mPrintDialog_OnReturn()
    If DataAccess.LabelCount = 0 Then
        If Not mPrintPreview Is Nothing Then
            mPrintPreview.Hide
            Set mPrintPreview = Nothing
        End If
    Else
    End If
End Sub

Private Sub mPrintDialog_Terminated()
    Set mPrintDialog = Nothing
End Sub

Private Sub mPrintPreview_RequestViewPrintDialog()
    If mPrintDialog Is Nothing Then: Set mPrintDialog = New WinPrintDialogue
    mPrintDialog.Show
End Sub

Private Sub mPrintPreview_Terminated()
    Set mPrintPreview = Nothing
End Sub

Private Sub mSearch_LogViewRequested(dataView As CCursorReader)
    If mLogNav Is Nothing Then: Set mLogNav = New WinLogNav
    DataAccess.SetUserLabelsQuery dataView.GetRecords2DArray
    mLogNav.Show
End Sub

Private Sub mSearch_Terminated()
    Set mSearch = Nothing
End Sub

Private Sub mSettings_FileManagerViewRequested()
    If mFileManager Is Nothing Then: Set mFileManager = New WinFileManager
    mFileManager.Show
End Sub

Private Sub mSettings_PreferedNameViewRequested()
    If mPreferredNames Is Nothing Then: Set mPreferredNames = New WinPreferredNames
    mPreferredNames.Show
End Sub

Private Sub mSettings_QueryManagerViewRequested()
    If mQueryManager Is Nothing Then: Set mQueryManager = New WinQueryManager
    mQueryManager.Show
End Sub

Private Sub mSettings_Terminated()
    Set mSettings = Nothing
End Sub

Private Sub mSplash_OnActivate()
    SplashActions
End Sub
Private Sub mMainNav_SearchViewRequested()
    If mSearch Is Nothing Then: Set mSearch = New WinUserQuery
    mSearch.Show
End Sub
Private Sub mMainNav_ManualEntryViewRequested(MODE As iGenerateLabels.ENTRY_MODE)
    If mLogNav Is Nothing Then: Set mLogNav = New WinLogNav
    mLogNav.FillEntries MODE
    mLogNav.Show
End Sub
Private Sub mMainNav_SettingsViewRequested()
    If mSettings Is Nothing Then: Set mSettings = New WinSettingsMain
    mSettings.Show
End Sub

Private Sub mSplash_Terminated()
    Set mSplash = Nothing
End Sub
