VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@Folder("OMPkgTransfer.Classes")
'@PredeclaredId
'##############################################################################################################################################################################
'  ## DOCUMENTATION ##
'
'           Title:   TransferEngine
'          Author:   Damian J Roberts
'          Github:   https://github.com/IAmArudis/OMPkgTransfer
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  ## REFERENCES ##
'
'         Library:   -
'
'     ProjectName:   OMLogger
'     ProjectName:   VBACodeSnippets
'         Classes:      CSVersion
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  ## DESCRIPTION ##
'
'         PURPOSE:   Orchestrates the full translator chain on behalf of the consuming project.  Accepts the raw log entry and the version the caller needs, walks the
'                    TranslatorRegistry oldest-to-newest, and returns the fully translated TransferData_vX_XX container as Object.
'
'                    This class is the only entry point consuming code needs.  IProposalTranslator and all Translator_vX_XX classes are internal to OMPkgTransfer and never
'                    need to be referenced directly by the consumer.
'
'    INSTRUCTIONS:   No changes needed here when a new translator version is introduced.  Add the new Translator_vX_XX to TranslatorRegistry.RegisterAll() only.
'
'          NOTES:    Both version arguments must be clones of the caller's CSVersion instances.  GetTransferObject() does not clone them internally - the caller is
'                    responsible.  This prevents TransferEngine from mutating version references that belong to the consuming project.
'
'                    Returns Nothing if no translator is found for oLogVersion.  The caller is responsible for handling a Nothing return gracefully.
'
'        OVERRIDE:   N/A
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'##############################################################################################################################################################################

'##############################################################################################################################################################################
'  PRE-PROCESSORS, IMPLEMENTATION STATEMENTS, SANITY STATEMENTS, CONSTANTS, TYPES, ENUMS
'##############################################################################################################################################################################
Option Explicit
Option Base 1


Private Type TClass
   Value As Variant
End Type


'##############################################################################################################################################################################
'  VARIABLES, FUNCTIONS-AS-CONSTANTS, EVENT DEFINITIONS
'##############################################################################################################################################################################
Private this As TClass


'##############################################################################################################################################################################
'  PROPERTIES
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  EVENTS
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  EXPOSED METHODS
'##############################################################################################################################################################################
'@Description ("Walks the translator chain and returns a fully translated TransferData_vX_XX as Object.  oLogVersion is the version the log data originated from.  oNeededVersion is the version the caller needs the data translated to.  Returns Nothing if no translator is found for oLogVersion.")
Public Function GetTransferObject(ByVal oEntryData As OMLogger.LogEntryData, ByVal oLogVersion As CSVersion, ByVal oNeededVersion As CSVersion) As Object
   Const METHOD_NAME                As String = "GetTransferObject"

   '  Find the starting translator for the log version.
   Dim oStartTranslator             As IProposalTranslator
   Set oStartTranslator = TranslatorRegistry.GetTranslator(oLogVersion.Clone)
   If oStartTranslator Is Nothing Then Exit Function

   '  Process the raw log entry into the initial container.
   Dim oContainer                   As Object
   Set oContainer = oStartTranslator.Process(oEntryData, oLogVersion.Clone)

   '  Walk the registry oldest-to-newest.  On each iteration, if the container
   '  version still hasn't reached the needed version, find the translator that
   '  owns the current container version and call Translate() to advance it.
   '  Stop as soon as the container version meets or beats the needed version.
   Dim i                            As Long
   For i = 1 To TranslatorRegistry.Count
      If Not oContainer.Version.IsOlderThan(oNeededVersion) Then Exit For
      If TranslatorRegistry.Item(i).CanHandle(oContainer.Version) Then
         Set oContainer = TranslatorRegistry.Item(i).Translate(oContainer)
      End If
   Next i

   Set GetTransferObject = oContainer
End Function



'##############################################################################################################################################################################
'  INTERNAL METHODS
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  PRIVATE METHODS
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  GENERIC CLASS BOILERPLATE CODE
'##############################################################################################################################################################################
'@Description ("[CONSTRUCTOR]")
Private Sub Class_Initialize()
End Sub


'@Description ("[DESTRUCTOR]")
Private Sub Class_Terminate()
End Sub


'@Description ("Returns the error source string to enable generating more descriptive and useful error messages.")
Private Property Get ERR_SOURCE(ByVal methodName As String) As String
   ERR_SOURCE = ProjectInfo.ProjectName & "." & TypeName(Me) & "." & methodName & "()"
End Property


'@Description ("This method returns this object.  Allows us to use WITH statements for creation in factories and static classes like this:  With New Class1:  Set Create = .Self:  End with")
Friend Property Get Self() As TransferEngine
   Set Self = Me
End Property


'##############################################################################################################################################################################
'  IClass IMPLEMENTATION
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  IClass IMPLEMENTATION
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  OLD CODE / TEMP CODE / DEBUG CODE
'     Please place retired code here.
'##############################################################################################################################################################################
