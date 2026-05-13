VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TranslatorRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "OMPkgTransfer.Classes"
'@PredeclaredId
'##############################################################################################################################################################################
'  ## DOCUMENTATION ##
'
'           Title:   TranslatorRegistry
'          Author:   Damian J Roberts
'          Github:   https://github.com/IAmArudis/OMPkgTransfer
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  ## REFERENCES ##
'
'         Library:   -
'
'     ProjectName:   VBACodeSnippets
'         Classes:      CSVersion
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  ## DESCRIPTION ##
'
'         PURPOSE:   Holds all known IProposalTranslator implementations in oldest-to-newest order.
'                    Exposes Count and Item so TransferCode can walk the chain by index, which
'                    guarantees ordering in a way that For Each cannot.
'
'                    The last registered translator (Item(Count)) is always the current version.  TransferEngine walks this registry to orchestrate the full translation
'                    chain.  Apply() lives in PkgDocs, which has native access to all worksheet targets.
'
'    INSTRUCTIONS:   When a new Translator_vX_XX is created:
'                       1. Add it to RegisterAll() BELOW all previous entries.
'                       2. Do not change the order of existing entries.
'                    That is the only change required here when a new version is introduced.
'
'        OVERRIDE:   RegisterAll():  Add new Translator_vX_XX entries here as new versions are created.  Always insert at the BOTTOM of the registration list.
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'##############################################################################################################################################################################
Option Explicit
Option Base 1


Private Type TClass
   Translators As Collection     'Ordered oldest-first.  Populated once by RegisterAll().
End Type


'##############################################################################################################################################################################
'  VARIABLES, FUNCTIONS-AS-CONSTANTS, EVENT DEFINITIONS
'##############################################################################################################################################################################
Private this As TClass


'##############################################################################################################################################################################
'  PROPERTIES
'##############################################################################################################################################################################
'@Description ("Returns the number of registered translators.")
Public Property Get Count() As Long
   If this.Translators Is Nothing Then Call RegisterAll
   Count = this.Translators.Count
End Property


'##############################################################################################################################################################################
'  EVENTS
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  EXPOSED METHODS
'##############################################################################################################################################################################
'@Description ("Returns the IProposalTranslator at the given 1-based index.")
Public Function Item(ByVal i As Long) As IProposalTranslator
   If this.Translators Is Nothing Then Call RegisterAll
   Set Item = this.Translators(i)
End Function


'@Description ("Returns the correct IProposalTranslator for the given version, walking oldest-to-newest by index.  Returns Nothing if no match is found.")
Public Function GetTranslator(ByVal oVersion As CSVersion) As IProposalTranslator
   If this.Translators Is Nothing Then Call RegisterAll
   Dim i As Long
   For i = 1 To this.Translators.Count
      Dim oTranslator As IProposalTranslator
      Set oTranslator = this.Translators(i)
      If oTranslator.CanHandle(oVersion) Then
         Set GetTranslator = oTranslator
         Exit Function
      End If
   Next i
   'No match found - caller is responsible for handling a Nothing return gracefully.
End Function


'##############################################################################################################################################################################
'  INTERNAL METHODS
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  PRIVATE METHODS
'##############################################################################################################################################################################
'@Override
'@Description ("Registers all known translators in oldest-to-newest order.  Called once on first use.")
Private Sub RegisterAll()
   Set this.Translators = New Collection
   '## ADD NEW TRANSLATORS HERE - always insert BELOW the previous entry. ##
   this.Translators.Add Translator_v4_86
   '## /ADD ##
End Sub


'##############################################################################################################################################################################
'  GENERIC CLASS BOILERPLATE CODE
'##############################################################################################################################################################################
'@Description ("[CONSTRUCTOR]")
Private Sub Class_Initialize()
End Sub


'@Description ("[DESTRUCTOR]")
Private Sub Class_Terminate()
   Set this.Translators = Nothing
End Sub


'@Description ("Returns the error source string to enable generating more descriptive and useful error messages.")
Private Property Get ERR_SOURCE(ByVal methodName As String) As String
   ERR_SOURCE = ProjectInfo.ProjectName & "." & TypeName(Me) & "." & methodName & "()"
End Property


'@Description ("This method returns this object.  Allows us to use WITH statements for creation in factories and static classes like this:  With New Class1:  Set Create = .Self:  End with")
Friend Property Get Self() As TranslatorRegistry
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
