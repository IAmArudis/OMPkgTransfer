Attribute VB_Name = "DS_EnumMappers"
'@Folder "OMPkgTransfer.Data Structures"
'##############################################################################################################################################################################
'  ## DOCUMENTATION ##
'
'           Title:   DS_EnumMappers
'          Author:   Damian J Roberts
'          Github:   https://github.com/IAmArudis/OM_PkgTransfer
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  ## REFERENCES ##
'
'         Library:   -
'
'     ProjectName:   -
'         Classes:      -
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  ## DESCRIPTION ##
'
'         PURPOSE:   Maps raw ProposalTransfer.log string values to their corresponding Long enum values.  Returns Long rather than a named enum type to keep this module
'                    fully decoupled from PkgDocs' enum definitions.
'
'                    0 is reserved as the unset/error sentinel across all enums.  Any mapper that receives an unrecognized string returns 0 so that callers can detect a bad
'                    parse without needing to know the valid range.
'
'                    These values are FROZEN once a version of PkgDocs has written them to a log file.  Never renumber or remove a case.  Add new cases at the bottom of the
'                    relevant Select block only.
'
'    INSTRUCTIONS:   When a new enum value is added in PkgDocs:
'                       1. Add the new Case at the bottom of the relevant function here.
'                       2. Do NOT change the Long value of any existing Case.
'                       3. Do NOT remove any existing Case.
'
'        OVERRIDE:   N/A - pure data module.  No classes.
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


Private Const CLASS_NAME = "DS_EnumMappers"


Private Type TClass
   Value As Boolean
End Type


'##############################################################################################################################################################################
'  VARIABLES, FUNCTIONS-AS-CONSTANTS
'##############################################################################################################################################################################
Private this As TClass


'##############################################################################################################################################################################
'  PROPERTIES
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  EXPOSED METHODS
'##############################################################################################################################################################################
'@Description ("Maps an event type string to its Long value.  Returns 0 if unrecognized.")
Public Function GetEventTypeEnum(ByVal Value As String) As Long
   Select Case Value
      Case "Ceremony Only":           GetEventTypeEnum = 1
      Case "Reception Only":          GetEventTypeEnum = 2
      Case "Ceremony and Reception":  GetEventTypeEnum = 3
      Case Else:                      GetEventTypeEnum = 0
   End Select
End Function





'@Description ("Maps a catering option string to its Long value.  Returns 0 if unrecognized.")
Public Function GetFoodEnum(ByVal Value As String) As Long
   Select Case Value
      Case "None":           GetFoodEnum = 1
      Case "Appetizer":      GetFoodEnum = 2
      Case "Plated Dinner":  GetFoodEnum = 3
      Case "Dinner Plus":    GetFoodEnum = 4
      Case Else:             GetFoodEnum = 0
   End Select
End Function





'@Description ("Maps a ceremony location string to its Long value.  Returns 0 if unrecognized.")
Public Function GetLocationEnum(ByVal Value As String) As Long
   Select Case Value
      Case "No Ceremony":  GetLocationEnum = 1
      Case "Chapel":       GetLocationEnum = 2
      Case "Garden":       GetLocationEnum = 3
      Case "Fire Place":   GetLocationEnum = 4
      Case Else:           GetLocationEnum = 0
   End Select
End Function





'@Description ("Maps a package type string to its Long value.  Returns 0 if unrecognized.")
Public Function GetPackageEnum(ByVal Value As String) As Long
   Select Case Value
      Case "Wedding Package":           GetPackageEnum = 1
      Case "Ceremony Only Package":     GetPackageEnum = 2
      Case "Reception Only Package":    GetPackageEnum = 3
      Case "Special Event Package":     GetPackageEnum = 4
      Case "Memorial Service Package":  GetPackageEnum = 5
      Case Else:                        GetPackageEnum = 0
   End Select
End Function





'@Description ("Maps a rehearsal type string to its Long value.  Returns 0 if unrecognized.")
Public Function GetRehearsalTypeEnum(ByVal Value As String) As Long
   Select Case Value
      Case "None":          GetRehearsalTypeEnum = 1
      Case "30-Mins Prior": GetRehearsalTypeEnum = 2
      Case "Thursday":      GetRehearsalTypeEnum = 3
      Case "Friday":        GetRehearsalTypeEnum = 4
      Case "Other":         GetRehearsalTypeEnum = 5
      Case Else:            GetRehearsalTypeEnum = 0
   End Select
End Function





'##############################################################################################################################################################################
'  PRIVATE METHODS
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  GENERIC CLASS BOILERPLATE CODE
'##############################################################################################################################################################################
'@Description ("Returns the error source string to enable generating more descriptive and useful error messages.")
Private Property Get ERR_SOURCE(ByVal methodName As String) As String
   ERR_SOURCE = ProjectInfo.ProjectName & "." & CLASS_NAME & "." & methodName & "()"
End Property


'##############################################################################################################################################################################
'  OLD CODE / TEMP CODE / DEBUG CODE
'     Please place retired code here.
'##############################################################################################################################################################################
