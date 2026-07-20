Attribute VB_Name = "DS_TransferTypes"
'@Exposed
'@Folder "OMPkgTransfer.Data Structures"
'##############################################################################################################################################################################
'  ## DOCUMENTATION ##
'
'           Title:   DS_TransferTypes
'          Author:   Damian J Roberts
'          Github:   https://github.com/IAmArudis/OMPkgTransfer
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
'         PURPOSE:   Versioned UDT definitions for all ProposalTransfer data structures.  Each version of PkgDocs that introduced a structural change to the transfer log gets
'                    its own set of T*Data_vX_XX types here.
'
'                    These types serve as the intermediate shapes used by each IProposalTranslator concrete implementation's Get*Data methods.  By the time data reaches
'                    Translate(), it has been converted to the CURRENT version's types.
'
'    INSTRUCTIONS:   When a new translator version is needed:
'                       1. Add a new versioned block below (copy the previous version's block).
'                       2. Modify only the types that changed.  Unchanged types are exact copies - intentional.
'                       3. Name each type after the FIRST version that introduced that shape.
'                          e.g. TClientData_v4_86 was born in v4.86.x and keeps that name until a structural change occurs, at which point TClientData_v4_86_5 (or
'                          TClientData_v4_87, etc.) is added.
'                       4. Declare nested UDTs BEFORE the UDTs that contain them.
'                       5. Do NOT remove or modify existing versioned types.  They are frozen history.
'
'           NOTES:   Enum-backed fields (PackageType, EventLocation, FoodOption, Rehearsal, etc.) are stored as Long intentionally.  This decouples the UDTs from enum
'                    definitions that may change in future versions.  Callers are responsible for mapping Long values to their current enum equivalents.
'
'        OVERRIDE:   N/A - this is a pure data module.  No methods.
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'##############################################################################################################################################################################
Option Explicit
Option Base 1


'##############################################################################################################################################################################
'  v4.86.x  -  Initial transfer log structure.
'              Baseline types.  All subsequent versions derive from these.
'##############################################################################################################################################################################

Public Type TOfficeUseData_v4_86
   Agent                   As String
   Minister                As String
   TourInfo                As String
   Budget                  As Currency
   SpecialConsiderations   As String
End Type


Public Type TClientData_v4_86
   Title                   As String
   ClientName              As String
   Phone                   As String
   Email                   As String
   Instagram               As String
   Address1                As String
   Address2                As String
   City                    As String
   State                   As String
   Zip                     As String
End Type


Public Type TEventData_v4_86
   EventDate               As String
   CheckInTime             As String
   CheckOutTime            As String
   CeremonyTime            As String
   Headcount               As Long
   PackageType             As Long        'Long - see NOTES above.  Maps to PackageTypeEnum via DS.GetPackageEnum
   EventLocation           As Long        'Long - see NOTES above.  Maps to LocationEnum via DS.GetLocationEnum
End Type


Public Type TFoodData_v4_86
   FoodOption              As Long        'Long - see NOTES above.  Maps to FoodEnum via DS.GetFoodEnum
   FoodQty                 As Long
   FoodPrice               As Currency
   CakeQty                 As Long
   CakePrice               As Currency
   PunchQty                As Long
   PunchPrice              As Currency
End Type


Public Type TMiscItemData_v4_86
   Qty                     As Long
   Label                   As String
   Price                   As Currency
End Type


'  TOptionalsData_v4_86 - declared after TMiscItemData_v4_86 which it contains.
Public Type TOptionalsData_v4_86
'  All of these are Qty indicators except ColdSparklers, CloudMachine, and Misc items.
   BarRental               As Long
   BeverageTower           As Long
   SilverService           As Long
   CandyStation            As Long
   ChocolateFountain       As Long
   GroomsCake              As Long
   PlatinumPackage         As Long
   ReceptionCenterpieces   As Long
   DanceFloorMonogram      As Long
   VideoMonitorRental      As Long
   Minister                As Long
   WeddingPartyFlowers     As Long
   Rehearsal               As Long        'Long - see NOTES above.  Maps to RehearsalTypeEnum via DS.GetRehearsalTypeEnum
   VideoCeremony           As Long
   ColdSparklers           As String      'These are string because the Qty will always map to one except in one case where it maps to 0.
   CloudMachine            As String      'These are string because the Qty will always map to one except in one case where it maps to 0.
   MiscTax1                As TMiscItemData_v4_86
   MiscTax2                As TMiscItemData_v4_86
   MiscTax3                As TMiscItemData_v4_86
   MiscNoTax1              As TMiscItemData_v4_86
   MiscNoTax2              As TMiscItemData_v4_86
   MiscNoTax3              As TMiscItemData_v4_86
End Type


'##############################################################################################################################################################################
'  vX.XX.x  -  [NEXT VERSION GOES HERE]
'              Copy the previous version's block, increment the suffix on changed types only.
'              Unchanged types from the previous version are referenced directly by name.
'              Declare nested UDTs before the UDTs that contain them.
'              Leave all previous version blocks untouched.
'##############################################################################################################################################################################
