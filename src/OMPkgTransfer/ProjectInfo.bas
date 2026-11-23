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
'     ProjectName:   VBACodeSnippets
'         Classes:      CSSystem
'                       CSVersion
'         Modules:      CodeSnippetsDS:   vp__VersionParts
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  ## DESCRIPTION ##
'
'         PURPOSE:   Project Identity Module.  Validates version consistency before each save, exports modules to src\ for source control,
'                    and deploys xlam/xlsm versions to the live locations.
'
'    INSTRUCTIONS:   1)  Add this module to the project, either by:
'                        a)  Copying it in directly, or
'                        b)  Saving these .template files in %APPDATA%\Rubberduck\Templates\, then right-clicking in the Code Explorer
'                            to add it from the templates menu.  (May require a restart of the Office Application.)
'                    2)  Create the repository folder (DEFAULT_REPO_PATH & PROJECT_NAME) before the first save.  Saves are blocked until
'                        it exists.
'                    3)  Fill in the constants:
'                        -  Release:           VERSION_ and REVISION_DATE.  Update these with every release.  Leave DATE_OF_LAST_USAGE
'                                              empty until a breaking release supersedes this version.
'                        -  Project creation:  Set once when the project is created.
'                        -  Environment:       The same for every project.  Every folder constant ends with "\".
'                    4)  Add these 2 methods to your workbook's code-behind:
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
'                    NOTE:  A minor version bump means a BREAKING change in this suite, not the semver meaning.  Patches are non-breaking
'                           and overwrite the live file in place.  See the README for the full versioning scheme.
'
'        OVERRIDE:   IsBaseProject if the project requires non-standard filename comparison.
'                    (e.g. date-stamped filenames like theOnePackage, where the filename contains a date that may not match exactly but
'                    the prefix should still be recognized.)
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'##############################################################################################################################################################################

'##############################################################################################################################################################################
'  PRE-PROCESSORS, IMPLEMENTATION STATEMENTS, SANITY STATEMENTS, CONSTANTS, TYPES, ENUMS
'##############################################################################################################################################################################
Option Private Module
Option Explicit
Option Base 1


'The name of this module.  This shouldn't change from project to project.
Private Const CLASS_NAME                        As String = "ProjectInfo"

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'  These constants are the ones that will need to be updated with each release.
'
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'The current version string.  e.g. "v2.5.24" or "v10.20.1"
Private Const VERSION_                          As String = "v0.1.6"

'The date of the last revision.  e.g. "2026-03-30"
Private Const REVISION_DATE                     As String = "2026-11-20"

'The date this version was superseded by a breaking (minor or major) release.  Leave empty while current.  e.g. "2027-01-15"
Private Const DATE_OF_LAST_USAGE                As String = ""





'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'  These constants shouldn't change often at all.  They are the ones that are set at the beginning of a project's creation and only need updating if the scope of a project
'  changes in some way.
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'The base filename shared by the master and live files, without version token or extension.  e.g. "VBACodeSnippets" or "OMPkgDocs"
Private Const FILE_BASE_NAME                    As String = "OMPkgTransfer"

'How much of VERSION_ appears in the master and live filenames.  vp_None, vp_Major, vp_MajorMinor, or vp_Full.
'   e.g. vp_None       -> "VBACodeSnippets.xlsm"
'        vp_MajorMinor -> "OMPkgDocs v1.1.xlsm"
'Declared As Long because a Const can't be declared as an Enum type.
Private Const FILENAME_VERSION_PARTS            As Long = vp_None

'The name of the project.  e.g. "VBACodeSnippets"
Private Const PROJECT_NAME                      As String = "OMPkgTransfer"

'The extension of every master (development) file.  This shouldn't change from project to project.
Private Const MASTER_FILE_EXTENSION             As String = ".xlsm"

'The extension of the live project.  e.g. ".xlsm" or ".xlam"
Private Const PROJECT_LIVE_TARGET_EXTENSION     As String = ".xlam"

'Set to True if this project requires versioned xlsm backups.
Private Const NEEDS_VERSIONED_XLSM              As Boolean = False





'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'  Environment Constants
'
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'The parent path for all repositories.
Private Const DEFAULT_REPO_PATH                 As String = "R:\Users\DRoberts\Repositories\"

'The src\ export path for Git tracking.  e.g. "\src\" or "\src\ComponentName\"
Private Const SRC_EXPORT_PATH                   As String = DEFAULT_REPO_PATH & PROJECT_NAME & "\src\"

'The path for versioned xlsm backups.  e.g. "\backups\" or "\backups\ComponentName\"
Private Const VERSIONED_BACKUP_PATH             As String = DEFAULT_REPO_PATH & PROJECT_NAME & "\backups\"

'The primary live deployment folder.  The filename is built from FILE_BASE_NAME, the version token, and PROJECT_LIVE_TARGET_EXTENSION.
Private Const LIVE_FOLDER_PRIMARY               As String = "\\dbserver\docs\Upload\TechSupport Shared\Add-Ins\"

'The secondary live deployment folder.  The filename is built the same way as the primary.
Private Const LIVE_FOLDER_SECONDARY             As String = "R:\.TechSupport_Shared\Excel Add-Ins\"





'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'  Module-level Types and Enums
'
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Private Type TClass
   oVersion As CSVersion
End Type


'##############################################################################################################################################################################
'  VARIABLES, FUNCTIONS-AS-CONSTANTS
'##############################################################################################################################################################################
Private this                                    As TClass


'##############################################################################################################################################################################
'  PROPERTIES
'##############################################################################################################################################################################
'@Description("Returns the date this version was superseded by a breaking release, or empty string if it is current.")
Public Property Get DateOfLastUsage() As String
   DateOfLastUsage = DATE_OF_LAST_USAGE
End Property





'@Description("Returns the version token used in this project's master and live filenames (e.g. 'v1.1' for vp_MajorMinor).  Returns empty string for vp_None or if VERSION_ could not be parsed.")
Public Property Get FilenameVersionToken() As String
   FilenameVersionToken = Version.ToString(eParts:=FILENAME_VERSION_PARTS)
End Property





'@Description("Returns True if `ThisWorkbook` is the base project file, ignoring the version token in the filename. Override this property if non-standard filename comparison is required.")
Public Property Get IsBaseProject() As Boolean
   IsBaseProject = (CSSystem.StripVersionToken(ThisWorkbook.Name, FILENAME_VERSION_PARTS) = FILE_BASE_NAME & MASTER_FILE_EXTENSION)
End Property





'@Description("Returns True if this project deploys as an .xlam add-in.")
Private Property Get IsXlamProject() As Boolean
   IsXlamProject = (PROJECT_LIVE_TARGET_EXTENSION = ".xlam")
End Property





'@Description("Returns the primary live deployment path, including the version token if FILENAME_VERSION_PARTS calls for one.")
Public Property Get LivePathPrimary() As String
   LivePathPrimary = LIVE_FOLDER_PRIMARY & BuildVersionedFileName(PROJECT_LIVE_TARGET_EXTENSION)
End Property





'@Description("Returns the secondary live deployment path, including the version token if FILENAME_VERSION_PARTS calls for one.")
Public Property Get LivePathSecondary() As String
   LivePathSecondary = LIVE_FOLDER_SECONDARY & BuildVersionedFileName(PROJECT_LIVE_TARGET_EXTENSION)
End Property





'@Description("Returns the expected master filename, including the version token if FILENAME_VERSION_PARTS calls for one.")
Public Property Get MasterFileName() As String
   MasterFileName = BuildVersionedFileName(MASTER_FILE_EXTENSION)
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
'@Description("Handles the BeforeSave event. Validates the ProjectInfo constants and version consistency, blocking the save and reporting all issues at once if any check fails. If the workbook filename doesn't match MasterFileName, cancels and launches a pre-populated Save As dialog.")
Public Sub Event_Workbook_BeforeSave(ByVal bSaveAsUI As Boolean, ByRef bCancel As Boolean)
   Static bOurDialogActive             As Boolean
   If bOurDialogActive Then Exit Sub   'Second pass triggered by our own Save As dialog -- let it through.
   If Not CSSystem.IsTechSupport Then Exit Sub

   'IsBaseProject compares against FILE_BASE_NAME, so an empty value would make every check below silently skip.
   If FILE_BASE_NAME = vbNullString Then
      MsgBox "Save cancelled. FILE_BASE_NAME is empty, so ProjectInfo can't identify the base project." _
           , vbExclamation Or vbOKOnly, PROJECT_NAME & " -- Pre-Save Check"
      bCancel = True
      Exit Sub
   End If
   If Not IsBaseProject Then Exit Sub

   Dim sIssues                         As String
   sIssues = vbNullString

   'Check 1: Required constants are filled in and valid.  Ordered to match the constant bands.
   'Release constants.
   If VERSION_ = vbNullString Then sIssues = sIssues & "  - VERSION_ is empty." & vbNewLine
   If REVISION_DATE = vbNullString Then
      sIssues = sIssues & "  - REVISION_DATE is empty." & vbNewLine
   ElseIf Not IsDate(REVISION_DATE) Then
      sIssues = sIssues & "  - REVISION_DATE (" & REVISION_DATE & ") is not a valid date." & vbNewLine
   ElseIf Not CDate(REVISION_DATE) = Date Then
      sIssues = sIssues & "  - REVISION_DATE (" & REVISION_DATE & ") is not today's date (" & Format$(Date, "YYYY-MM-DD") & ")." & vbNewLine
   End If
   If Not DATE_OF_LAST_USAGE = vbNullString And Not IsDate(DATE_OF_LAST_USAGE) Then
      sIssues = sIssues & "  - DATE_OF_LAST_USAGE (" & DATE_OF_LAST_USAGE & ") is not a valid date." & vbNewLine
   End If

   'Project-creation constants.
   If FILENAME_VERSION_PARTS < vp_None Or FILENAME_VERSION_PARTS > vp_Full Then
      sIssues = sIssues & "  - FILENAME_VERSION_PARTS (" & FILENAME_VERSION_PARTS & ") is not a valid vp__VersionParts value." & vbNewLine
   End If
   If PROJECT_NAME = vbNullString Then sIssues = sIssues & "  - PROJECT_NAME is empty." & vbNewLine
   If MASTER_FILE_EXTENSION = vbNullString Then sIssues = sIssues & "  - MASTER_FILE_EXTENSION is empty." & vbNewLine

   'Confirm Live Target Extension has been set.
   If Not (PROJECT_LIVE_TARGET_EXTENSION = ".xlam" Or PROJECT_LIVE_TARGET_EXTENSION = ".xlsm") Then
      sIssues = sIssues & "  - PROJECT_LIVE_TARGET_EXTENSION (" & PROJECT_LIVE_TARGET_EXTENSION & ") must be "".xlam"" or "".xlsm""." & vbNewLine
   End If

   'Environment.  The repository must exist before the project is started.
   If Not PROJECT_NAME = vbNullString And Dir(DEFAULT_REPO_PATH & PROJECT_NAME, vbDirectory) = vbNullString Then
      sIssues = sIssues & "  - Repository folder not found: " & DEFAULT_REPO_PATH & PROJECT_NAME & vbNewLine
   End If

   'An uninitialized Version (bad parse) is treated as a blocker alongside the empty-constant check.
   If Not VERSION_ = vbNullString And Not Version.IsInitialized Then
      sIssues = sIssues & "  - VERSION_ (" & VERSION_ & ") could not be parsed." & vbNewLine
   End If

   'Check 2: VERSION_ meets or exceeds the NextExpectedVersion document property.
   'Only meaningful if VERSION_ parsed correctly.
   If Version.IsInitialized Then
      Dim sNextExpected                As String
      sNextExpected = GetNextExpectedVersion()
      If Not sNextExpected = vbNullString Then
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
   If Not sIssues = vbNullString Then
      MsgBox "Save cancelled. Resolve the following before saving:" & vbNewLine & vbNewLine _
           & sIssues _
           , vbExclamation Or vbOKOnly, PROJECT_NAME & " -- Pre-Save Check"
      bCancel = True
      Exit Sub
   End If

   'Check 3: Workbook filename matches MasterFileName.
   'IsBaseProject already confirmed the base name and extension, so any remaining difference is the version token
   '(wrong, missing, or present on a vp_None project).
   If Not ThisWorkbook.Name = MasterFileName Then
      bCancel = True
      If bSaveAsUI Then
         'User-initiated Save As with the wrong filename -- notify and re-launch pre-populated with the correct name.
         MsgBox "Filename does not match the expected name for VERSION_ (" & Version.ToString() & ")." & vbNewLine & vbNewLine _
              & "Opening Save As pre-populated with:" & vbNewLine _
              & "  " & MasterFileName _
              , vbExclamation Or vbOKOnly, PROJECT_NAME & " -- Pre-Save Check"
      Else
         'Regular save with a stale filename -- notify and launch Save As pre-populated with the correct name.
         MsgBox "The filename needs to be updated to match VERSION_ (" & Version.ToString() & ")." & vbNewLine & vbNewLine _
              & "Opening Save As pre-populated with:" & vbNewLine _
              & "  " & MasterFileName _
              , vbInformation Or vbOKOnly, PROJECT_NAME & " -- Pre-Save Check"
      End If
      bOurDialogActive = True
      Application.Dialogs(XlBuiltInDialog.xlDialogSaveAs).Show MasterFileName
      bOurDialogActive = False
   End If
End Sub











'@Description("Handles the AfterSave event. Exports modules to `SrcExportPath` for Git tracking and optionally exports to live locations.")
Public Sub Event_Workbook_AfterSave(ByVal bSuccess As Boolean)
   Const METHOD_NAME                As String = "Event_Workbook_AfterSave"
   If Not bSuccess Then Exit Sub
   If Not IsBaseProject Then Exit Sub
   If Not CSSystem.IsTechSupport Then Exit Sub
   On Error GoTo CleanUp

   'Save current status of Events and Alerts.
   Dim bEventsEnabled               As Boolean
   Dim bDisplayAlerts               As Boolean
   bEventsEnabled = Application.EnableEvents
   bDisplayAlerts = Application.DisplayAlerts

   'Always silently export modules to src\ for Git tracking.
   Call CSSystem.ExportProjectModules(ThisWorkbook.VBProject, SrcExportPath)

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
   End Select
End Sub















'##############################################################################################################################################################################
'  EXPOSED METHODS
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  PRIVATE METHODS
'##############################################################################################################################################################################
'@Description("[HELPER] Returns FILE_BASE_NAME, followed by a space and the filename version token (if any), followed by 'sExt'.")
Private Function BuildVersionedFileName(ByVal sExt As String) As String
   Dim sToken                       As String
   sToken = FilenameVersionToken
   If sToken = vbNullString Then
      BuildVersionedFileName = FILE_BASE_NAME & sExt
   Else
      BuildVersionedFileName = FILE_BASE_NAME & " " & sToken & sExt
   End If
End Function





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
   If Not Err.Number = 0 Then
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


