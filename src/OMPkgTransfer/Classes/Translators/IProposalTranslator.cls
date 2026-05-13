VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProposalTranslator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "OMPkgTransfer.Classes.Translators"
'@Interface
'##############################################################################################################################################################################
'  ## DOCUMENTATION ##
'
'           Title:   IProposalTranslator
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
'     ProjectName:   OMLogger
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  ## DESCRIPTION ##
'
'         PURPOSE:   Interface contract for all versioned ProposalTransfer implementations.  Each concrete Translator_vX_XX class implements this interface and is responsible
'                    for reading a specific version range of ProposalTransfer.log data and translating it forward to the current version's shape.
'
'    INSTRUCTIONS:   Implement this interface in each Translator_vX_XX class.  Each method must be implemented - VBA does not support default implementations.
'
'                    Responsibility of each method:
'
'                      CanHandle()   - Returns True if this translator can handle the given version.  The LATEST translator always has an open upper bound.  When a new
'                                      translator is created, the previous translator's CanHandle() is given a concrete upper bound and the new one takes over as the
'                                      open-ended handler.
'
'                      Process()     - Reads the LogEntryData for the given version and returns a fully populated TransferData_vX_XX container as Object.  Calls each
'                                      Get*Data() method internally to build the container.  Get*Data() methods are NOT on this interface - they are public methods on each
'                                      concrete class, called directly by name when a future translator needs to delegate to a prior version.
'
'                      Translate()   - Receives the previous version's TransferData_vX_XX as Object and converts it to this version's shape.  When MAX_MAJOR is -1, stamps
'                                      version to v999.999.999 to signal end of chain.  When MAX_* are concrete, stamps to MAX_MAJOR.MAX_MINOR.(MAX_PATCH + 1) to hand off to
'                                      the next translator.  Version to MAX_PATCH + 1 before returning so the next translator in the chain knows to pick it up.  Returns the
'                                      converted container as Object back to TransferCode.  Each implementation is standalone - no base calls.
'
'          NOTES:    Apply() is intentionally absent from this interface.  It lives in TransferCode inside PkgDocs, which has native access to all worksheet targets (cells,
'                    ComboBoxes, etc.).  Putting Apply() here would require the interface to carry worksheet addresses, which is not its responsibility.
'
'                    These classes use PredeclaredId - keeping them stateless is intentional and important.
'
'                    Process() and Translate() return Object, and Translate() accepts Object, because VBA interfaces cannot return or accept versioned UDTs by name across
'                    implementations.  Each concrete class casts to its own TransferData_vX_XX internally via Set.
'
'                    CSVersion is used throughout instead of three separate Long arguments - it is already a dependency of this project via VBACodeSnippets.
'
'        OVERRIDE:   All methods - every Translator_vX_XX must provide a full implementation of every method defined here.  No method may be left as a stub in a concrete
'                    class.
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'##############################################################################################################################################################################
Option Explicit
Option Base 1


'##############################################################################################################################################################################
'  VARIABLES, FUNCTIONS-AS-CONSTANTS, EVENT DEFINITIONS
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  PROPERTIES
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  EVENTS
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  EXPOSED METHODS
'##############################################################################################################################################################################
'@Description ("Returns True if this translator can handle the given version of PkgDocs transfer data.")
Public Function CanHandle(ByVal oVersion As CSVersion) As Boolean
End Function


'@Description ("Reads the LogEntryData for the given version and returns a fully populated TransferData_vX_XX container as Object.")
Public Function Process(ByVal oEntryData As OMLogger.LogEntryData, ByVal oVersion As CSVersion) As Object
End Function


'@Description ("Converts the incoming TransferData_vX_XX to this version's shape.  Stamps Version to MAX_PATCH + 1 before returning.  Returns the converted container as Object.")
Public Function Translate(ByVal oData As Object) As Object
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
