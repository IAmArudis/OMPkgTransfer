VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferData_v4_86"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder "OMPkgTransfer.Classes.TransferData"
'##############################################################################################################################################################################
'  ## DOCUMENTATION ##
'
'           Title:   TransferData_v4_86
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
'         PURPOSE:   Thin container class for the v4.86.x ProposalTransfer data structures.  Holds all versioned UDTs from DS_TransferTypes privately in TClass, exposed via
'                    Public Property Get (read) and Friend Property Let/Set (write).
'
'                    Public Get allows PkgDocs (and any external consumer) to read any field.  Friend Let/Set restricts whole-UDT assignment to within PKGTransfer, so only
'                    Translator_vX_XX.Process() and Translate() can stamp data into the container.
'
'                    Exists solely to allow the UDT payload to cross Object boundaries in the IProposalTranslator interface - Process() returns this as Object, Translate()
'                    receives and returns it as Object.  Each concrete translator casts it back to this type internally via Set.
'
'    INSTRUCTIONS:   When a new translator version is needed:
'                      1. Create a new TransferData_vX_XX class (copy this one).
'                      2. Update TClass members and properties to reflect any UDT type changes.  Unchanged members keep their previous version's UDT type.
'                      3. Do NOT modify this class once a newer TransferData_vX_XX exists.  It is frozen history.
'
'          NOTES:    This class has no methods - it is a pure data container.  All logic lives in the concrete Translator_vX_XX implementations.
'
'        OVERRIDE:   N/A - no methods.
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
   Version      As CSVersion
   OfficeUse    As TOfficeUseData_v4_86
   Client1      As TClientData_v4_86
   Client2      As TClientData_v4_86
   Client3      As TClientData_v4_86
   Client4      As TClientData_v4_86
   EventData    As TEventData_v4_86
   FoodData     As TFoodData_v4_86
   Optionals    As TOptionalsData_v4_86
End Type


'##############################################################################################################################################################################
'  VARIABLES, FUNCTIONS-AS-CONSTANTS, EVENT DEFINITIONS
'##############################################################################################################################################################################
Private this As TClass


'##############################################################################################################################################################################
'  PROPERTIES
'##############################################################################################################################################################################
'@Description ("Returns the CSVersion instance stamped onto this container by the translator chain.")
Public Property Get Version() As CSVersion
   Set Version = this.Version
End Property


'@Description ("Sets the CSVersion instance.  Restricted to PKGTransfer internals.")
Friend Property Set Version(Value As CSVersion)
   Set this.Version = Value
End Property





'@Description ("Returns the office-use data block.")
Public Property Get OfficeUse() As TOfficeUseData_v4_86
   OfficeUse = this.OfficeUse
End Property


'@Description ("Sets the entire office-use data block.  Restricted to PKGTransfer internals.")
Friend Property Let OfficeUse(Value As TOfficeUseData_v4_86)
   this.OfficeUse = Value
End Property





'@Description ("Returns the primary client data block.")
Public Property Get Client1() As TClientData_v4_86
   Client1 = this.Client1
End Property


'@Description ("Sets the entire primary client data block.  Restricted to PKGTransfer internals.")
Friend Property Let Client1(Value As TClientData_v4_86)
   this.Client1 = Value
End Property





'@Description ("Returns the second client data block.")
Public Property Get Client2() As TClientData_v4_86
   Client2 = this.Client2
End Property


'@Description ("Sets the entire second client data block.  Restricted to PKGTransfer internals.")
Friend Property Let Client2(Value As TClientData_v4_86)
   this.Client2 = Value
End Property





'@Description ("Returns the third client data block.")
Public Property Get Client3() As TClientData_v4_86
   Client3 = this.Client3
End Property


'@Description ("Sets the entire third client data block.  Restricted to PKGTransfer internals.")
Friend Property Let Client3(Value As TClientData_v4_86)
   this.Client3 = Value
End Property





'@Description ("Returns the fourth client data block.")
Public Property Get Client4() As TClientData_v4_86
   Client4 = this.Client4
End Property


'@Description ("Sets the entire fourth client data block.  Restricted to PKGTransfer internals.")
Friend Property Let Client4(Value As TClientData_v4_86)
   this.Client4 = Value
End Property





'@Description ("Returns the event data block.")
Public Property Get EventData() As TEventData_v4_86
   EventData = this.EventData
End Property


'@Description ("Sets the entire event data block.  Restricted to PKGTransfer internals.")
Friend Property Let EventData(Value As TEventData_v4_86)
   this.EventData = Value
End Property





'@Description ("Returns the food data block.")
Public Property Get FoodData() As TFoodData_v4_86
   FoodData = this.FoodData
End Property


'@Description ("Sets the entire food data block.  Restricted to PKGTransfer internals.")
Friend Property Let FoodData(Value As TFoodData_v4_86)
   this.FoodData = Value
End Property





'@Description ("Returns the optionals data block.")
Public Property Get OptionalsData() As TOptionalsData_v4_86
   OptionalsData = this.Optionals
End Property


'@Description ("Sets the entire optionals data block.  Restricted to PKGTransfer internals.")
Friend Property Let OptionalsData(Value As TOptionalsData_v4_86)
   this.Optionals = Value
End Property


'##############################################################################################################################################################################
'  EVENTS
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  EXPOSED METHODS
'##############################################################################################################################################################################
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
   Set this.Version = CSVersion.CreateFromString(vbNullString)
End Sub


'@Description ("[DESTRUCTOR]")
Private Sub Class_Terminate()
   Set this.Version = Nothing
End Sub


'@Description ("Returns the error source string to enable generating more descriptive and useful error messages.")
Private Property Get ERR_SOURCE(ByVal methodName As String) As String
   ERR_SOURCE = ProjectInfo.ProjectName & "." & TypeName(Me) & "." & methodName & "()"
End Property


'@Description ("This method returns this object.  Allows us to use WITH statements for creation in factories and static classes like this:  With New Class1:  Set Create = .Self:  End with")
Friend Property Get Self() As TransferData_v4_86
   Set Self = Me
End Property


'##############################################################################################################################################################################
'  OLD CODE / TEMP CODE / DEBUG CODE
'     Please place retired code here.
'##############################################################################################################################################################################
