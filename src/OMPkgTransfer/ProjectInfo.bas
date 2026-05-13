Attribute VB_Name = "ProjectInfo"
'@Folder "OMPkgTransfer"
'##############################################################################################################################################################################
'  ## DOCUMENTATION ##
'
'           Title:   ProjectInfo
'          Author:   Damian J Roberts
'          Github:   https://github.com/IAmArudis/RubberDuck-Templates
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  ## REFERENCES ##
'
'         Library:   -
'
'     ProjectName:   CodeSnippets
'         Classes:      CSSystem
'                       CSVersion
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  ## DESCRIPTION ##
'
'         PURPOSE:   Project Identity Module.  Hanles the export and saving of modules and xlam/xlsm versions for source control
'
'    INSTRUCTIONS:   1a)  Copy this module to each project and fill in the constants below.
'                    Alternatively:
'                    1b)  You can save these .template files in %APPDATA%\Roaming\Rubberduck\Templates\
'                    2)  Then, in RubberDuckVBA you can Right-click in the CodeExplorer to add this template to any project (May require a restart of the Office Application).
'                    3)  Once that is done, add these 2 methods to your workbook's code-behind:
'
'                        Private Sub Workbook_AfterSave(ByVal Success As Boolean)
'                           Call ProjectInfo.Event_Workbook_AfterSave(Success)
'                        End Sub
'
'
'                        Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'                           Call ProjectInfo.Event_Workbook_BeforeSave(SaveAsUI, Cancel)
'                        End Sub
'
'        OVERRIDE:   IsBaseProject if the project requires non-standard filename comparison.
'                    (e.g. date-stamped filenames like theOnePackage where the filename contains a
'                    date that may not match exactly but the prefix should still be recognized.)
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'##############################################################################################################################################################################

'##############################################################################################################################################################################
'  PRE-PROCESSORS, IMPLEMENTATION STATEMENTS, SANITY STATEMENTS, CONSTANTS, TYPES, ENUMS
'##############################################################################################################################################################################
Option Private Module
Option Explicit
Option Base 1


'The name of this module.  This shouldn't change from project to project.
Private Const CLASS_NAME            As String = "ProjectInfo"

'The name of the project.  e.g. "VBACodeSnippets"
Private Const PROJECT_NAME          As String = "OMPkgTransfer"

'The current version string.  e.g. "v2.5.24" or "v10.20.1"
Private Const VERSION_              As String = "v0.1.0"    'New TransferData_ classes get a minor bump.

'The expected base filename  e.g. "CodeSnippets.xlam" or "Employee Schedule Master v1.4.7.xlsm"
Private Const MASTER_FILE_NAME      As String = "OMPkgTransfer.xlsm"

'The date of the last revision.  e.g. "2026-03-30"
Private Const REVISION_DATE         As String = "2026-05-12"

'The extension of the live project.  e.g. ".xlsm" or ".xlam"
Private Const PROJECT_LIVE_TARGET_EXTENSION As String = ".xlam"

'The parent path for all repositories.
Private Const DEFAULT_REPO_PATH     As String = "R:\Users\DRoberts\Repositories\"

'The src\ export path for Git tracking.  e.g. "\src or \src\ComponentName"
Private Const SRC_EXPORT_PATH       As String = DEFAULT_REPO_PATH & PROJECT_NAME & "\src"

'Set to True if this project requires versioned xlsm backups.
Private Const NEEDS_VERSIONED_XLSM  As Boolean = False

'The path for versioned xlsm backups.  e.g. "\backups" or "\backups\ComponentName"
Private Const VERSIONED_BACKUP_PATH As String = DEFAULT_REPO_PATH & PROJECT_NAME & "\backups"

'The primary live deployment path.  e.g. "\\dbserver\docs\Upload\TechSupport Shared\Add-Ins\CodeSnippets.xlam"
Private Const LIVE_PATH_PRIMARY     As String = "\\dbserver\docs\Upload\TechSupport Shared\Add-Ins\" & PROJECT_NAME & PROJECT_LIVE_TARGET_EXTENSION

'The secondary live deployment path.  e.g. "R:\.TechSupport_Shared\Excel Add-Ins\CodeSnippets.xlam"
Private Const LIVE_PATH_SECONDARY   As String = "R:\.TechSupport_Shared\Excel Add-Ins\" & PROJECT_NAME & PROJECT_LIVE_TARGET_EXTENSION


Private Type TClass
   oVersion As CSVersion
End Type


'##############################################################################################################################################################################
'  VARIABLES, FUNCTIONS-AS-CONSTANTS
'##############################################################################################################################################################################
Private this As TClass


'##############################################################################################################################################################################
'  PROPERTIES
'##############################################################################################################################################################################
'@Description("Returns True if `ThisWorkbook` is the base project file, ignoring version differences in the filename. Override this property if non-standard filename comparison is required.")
Public Property Get IsBaseProject() As Boolean
   IsBaseProject = (CSSystem.StripVersionToken(ThisWorkbook.Name) = CSSystem.StripVersionToken(MasterFileName))
End Property





'@Description("Returns True if this project deploys as an .xlam add-in.")
Private Property Get IsXlamProject() As Boolean
   IsXlamProject = (PROJECT_LIVE_TARGET_EXTENSION = ".xlam")
End Property





'@Description("Returns the primary live deployment path.")
Public Property Get LivePathPrimary() As String
   LivePathPrimary = LIVE_PATH_PRIMARY
End Property





'@Description("Returns the secondary live deployment path.")
Public Property Get LivePathSecondary() As String
   LivePathSecondary = LIVE_PATH_SECONDARY
End Property





'@Description("Returns the expected base filename of the project.")
Public Property Get MasterFileName() As String
   MasterFileName = MASTER_FILE_NAME
End Property





'@Description("Returns True if this project requires versioned xlsm backups.")
Public Property Get NeedsVersionedXlsm() As Boolean
   NeedsVersionedXlsm = NEEDS_VERSIONED_XLSM
End Property





'@Description("Returns the name of the project.")
Public Property Get ProjectName() As String
   ProjectName = PROJECT_NAME
End Property





'@Description("Returns the date of the last revision as a string.")
Public Property Get RevisionDate() As String
   RevisionDate = REVISION_DATE
End Property





'@Description("Returns the src\ export path for Git tracking.")
Public Property Get SrcExportPath() As String
   SrcExportPath = SRC_EXPORT_PATH
End Property





'@Description("Returns the current version as a CSVersion object. Lazily initialized from VERSION_ and cached for the lifetime of the session.")
Public Property Get Version() As CSVersion
   If this.oVersion Is Nothing Then
      Set this.oVersion = CSVersion.CreateFromString(VERSION_)
   ElseIf Not this.oVersion.IsInitialized Then
      Set this.oVersion = CSVersion.CreateFromString(VERSION_)
   End If
   Set Version = this.oVersion
End Property





'@Description("Returns the path for versioned xlsm backups.")
Public Property Get VersionedBackupPath() As String
   VersionedBackupPath = VERSIONED_BACKUP_PATH
End Property





'##############################################################################################################################################################################
'  EVENTS
'##############################################################################################################################################################################
'@Description("Handles the BeforeSave event. Validates version consistency across VERSION_, MASTER_FILE_NAME, and NextExpectedVersion, blocking the save and reporting all issues at once if any check fails. If the workbook filename needs updating to match VERSION_, cancels and launches a pre-populated Save As dialog.")
Public Sub Event_Workbook_BeforeSave(ByVal bSaveAsUI As Boolean, ByRef bCancel As Boolean)
   If Not IsBaseProject Then Exit Sub
   If Not CSSystem.IsTechSupport Then Exit Sub

   Dim sIssues                         As String
   sIssues = vbNullString

   'Check 1: Required constants are filled in and valid.
   If VERSION_ = vbNullString Then sIssues = sIssues & "  - VERSION_ is empty." & vbNewLine
   If MASTER_FILE_NAME = vbNullString Then sIssues = sIssues & "  - MASTER_FILE_NAME is empty." & vbNewLine
   If REVISION_DATE = vbNullString Then
      sIssues = sIssues & "  - REVISION_DATE is empty." & vbNewLine
   ElseIf Not IsDate(REVISION_DATE) Then
      sIssues = sIssues & "  - REVISION_DATE (" & REVISION_DATE & ") is not a valid date." & vbNewLine
   ElseIf CDate(REVISION_DATE) <> Date Then
      sIssues = sIssues & "  - REVISION_DATE (" & REVISION_DATE & ") is not today's date (" & Format$(Date, "YYYY-MM-DD") & ")." & vbNewLine
   End If

   'Construct a CSVersion object from VERSION_ -- all subsequent version checks use this object.
   'An uninitialized object (bad parse) is treated as a blocker alongside the empty-constant check.
   If VERSION_ <> vbNullString And Not Version.IsInitialized Then
      sIssues = sIssues & "  - VERSION_ (" & VERSION_ & ") could not be parsed." & vbNewLine
   End If

   'Version-string checks only make sense if VERSION_ parsed correctly.
   If Version.IsInitialized Then

      'Check 2: Version token in MASTER_FILE_NAME constant matches VERSION_.
      Dim eMFNState                    As fv__FileVersionState
      eMFNState = CSSystem.GetFilenameVersionState(MASTER_FILE_NAME, Version.ToString())
      Select Case eMFNState
         Case fv_Mismatch:   sIssues = sIssues & "  - MASTER_FILE_NAME version does not match VERSION_ (" & Version.ToString() & ")." & vbNewLine
         Case fv_Ambiguous:  sIssues = sIssues & "  - MASTER_FILE_NAME contains multiple version strings -- resolve the ambiguity." & vbNewLine
      End Select

      'Check 3: VERSION_ meets or exceeds the NextExpectedVersion document property.
      Dim sNextExpected                As String
      sNextExpected = GetNextExpectedVersion()
      If sNextExpected <> vbNullString Then
         Dim oNextExpected             As CSVersion
         Set oNextExpected = CSVersion.CreateFromString(sNextExpected)
         If Not oNextExpected.IsInitialized Then
            sIssues = sIssues & "  - NextExpectedVersion property (" & sNextExpected & ") could not be parsed -- delete it via the document properties dialog to reset." & vbNewLine
         ElseIf Version.IsOlderThan(oNextExpected) Then
            sIssues = sIssues & "  - VERSION_ (" & Version.ToString() & ") is behind expected (" & sNextExpected & "). Bump the version or check that you haven't rolled back." & vbNewLine
         End If
      End If

   End If

   'Block the save and report all issues at once before proceeding to the filename check.
   If sIssues <> vbNullString Then
      MsgBox "Save cancelled. Resolve the following before saving:" & vbNewLine & vbNewLine _
           & sIssues _
           , vbExclamation Or vbOKOnly, PROJECT_NAME & " -- Pre-Save Check"
      bCancel = True
      Exit Sub
   End If

   'Check 4: Workbook filename version matches VERSION_.
   'Only meaningful when VERSION_ parsed correctly and all prior checks passed.
   'Versionless filenames (fv_Missing) are always allowed through.
   Static bOurDialogActive             As Boolean
   If bOurDialogActive Then Exit Sub   'Second pass triggered by our own dialog -- let it through.
   If Version.IsInitialized And ThisWorkbook.Name <> MASTER_FILE_NAME Then
      Dim eWBState                     As fv__FileVersionState
      eWBState = CSSystem.GetFilenameVersionState(ThisWorkbook.Name, Version.ToString())
      Select Case eWBState
         Case fv_Mismatch, fv_Ambiguous
            bCancel = True
            If bSaveAsUI Then
               'User-initiated Save As with wrong filename -- notify and re-launch pre-populated with the correct name.
               MsgBox "Filename does not match VERSION_ (" & Version.ToString() & ")." & vbNewLine & vbNewLine _
                    & "Opening Save As pre-populated with:" & vbNewLine _
                    & "  " & MASTER_FILE_NAME _
                    , vbExclamation Or vbOKOnly, PROJECT_NAME & " -- Pre-Save Check"
            Else
               'Regular save with a stale filename -- notify and launch Save As pre-populated with the correct name.
               MsgBox "The filename needs to be updated to match VERSION_ (" & Version.ToString() & ")." & vbNewLine & vbNewLine _
                    & "Opening Save As pre-populated with:" & vbNewLine _
                    & "  " & MASTER_FILE_NAME _
                    , vbInformation Or vbOKOnly, PROJECT_NAME & " -- Pre-Save Check"
            End If
            bOurDialogActive = True
            Application.Dialogs(XlBuiltInDialog.xlDialogSaveAs).Show MASTER_FILE_NAME
            bOurDialogActive = False
      End Select
   End If
End Sub











'@Description("Handles the AfterSave event. Exports modules to `SrcExportPath` for Git tracking and optionally exports to live locations.")
Public Sub Event_Workbook_AfterSave(ByVal bSuccess As Boolean)
   Const METHOD_NAME                As String = "Event_Workbook_AfterSave"
   If Not bSuccess Then Exit Sub
   If Not IsBaseProject Then Exit Sub
   If Not CSSystem.IsTechSupport Then Exit Sub
   On Error GoTo CleanUp
   
   'Always silently export modules to src\ for Git tracking.
   Call CSSystem.ExportProjectModules(ThisWorkbook.VBProject, SrcExportPath)

   'Save current status of Events and Alerts.
   Dim bEventsEnabled               As Boolean
   Dim bDisplayAlerts               As Boolean
   bEventsEnabled = Application.EnableEvents
   bDisplayAlerts = Application.DisplayAlerts
   Application.EnableEvents = False
   Application.DisplayAlerts = False

   'Advance NextExpectedVersion for the next save.
   'Version consistency was already validated in BeforeSave so we advance unconditionally here.
   Call SetNextExpectedVersion(Version.Clone.BumpPatch().ToString())

   If IsXlamProject Then
      Dim vInput                    As Variant
      vInput = MsgBox("Export to live locations as .xlam?", vbYesNo Or vbQuestion)
      If vInput = vbYes Then
         Dim sOriginalPath          As String
         sOriginalPath = ThisWorkbook.FullName
         Call CSSystem.ExportToLiveLocation(ThisWorkbook, LivePathPrimary)
         Call CSSystem.ExportToLiveLocation(ThisWorkbook, LivePathSecondary)
         ThisWorkbook.SaveAs Filename:=sOriginalPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
      End If
   End If

   'Save versioned backup of the xlsm if required.
   If NeedsVersionedXlsm Then Call CSSystem.SaveVersionedBackup(ThisWorkbook, VersionedBackupPath)
   
   'Restore status of Events and Alerts.
   Application.EnableEvents = bEventsEnabled
   Application.DisplayAlerts = bDisplayAlerts
   Exit Sub

CleanUp:
   Dim lErrNum                      As Long
   Dim sErrDesc                     As String
   lErrNum = Err.Number
   sErrDesc = Err.Description
   Application.EnableEvents = bEventsEnabled
   Application.DisplayAlerts = bDisplayAlerts
   On Error GoTo 0
   Select Case lErrNum
      Case 1004
         If sErrDesc = "Method 'SaveAs' of object '_Workbook' failed" Then
            MsgBox "Cancelled exporting as .xlam.", vbOKOnly Or vbInformation
            Exit Sub
         End If
         Err.Raise lErrNum, ERR_SOURCE(METHOD_NAME), sErrDesc
      Case Else
         Err.Raise lErrNum, ERR_SOURCE(METHOD_NAME), sErrDesc
         Resume
   End Select
End Sub















'##############################################################################################################################################################################
'  EXPOSED METHODS
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  PRIVATE METHODS
'##############################################################################################################################################################################
'@Description("Returns the expected next version string from the custom document property. Returns empty string if the property does not exist.")
Private Function GetNextExpectedVersion() As String
   On Error Resume Next
   GetNextExpectedVersion = ThisWorkbook.CustomDocumentProperties("NextExpectedVersion").Value
   On Error GoTo 0
End Function





'@Description("Sets the expected next version custom document property, creating it if it does not exist.")
Private Sub SetNextExpectedVersion(ByVal sVersion As String)
   On Error Resume Next
   ThisWorkbook.CustomDocumentProperties("NextExpectedVersion").Value = sVersion
   If Err.Number <> 0 Then
      Err.Clear
      On Error GoTo 0
      ThisWorkbook.CustomDocumentProperties.Add _
         Name:="NextExpectedVersion", _
         LinkToContent:=False, _
         Type:=msoPropertyTypeString, _
         Value:=sVersion
   End If
   On Error GoTo 0
End Sub





'##############################################################################################################################################################################
'  GENERIC CLASS BOILERPLATE CODE
'##############################################################################################################################################################################
'@Description ("Returns the error source string to enable generating more descriptive and useful error messages.")
Private Property Get ERR_SOURCE(ByVal methodName As String) As String
   ERR_SOURCE = PROJECT_NAME & "." & CLASS_NAME & "." & methodName & "()"
End Property


'##############################################################################################################################################################################
'  OLD CODE / TEMP CODE / DEBUG CODE
'     Please place retired code here.
'##############################################################################################################################################################################

