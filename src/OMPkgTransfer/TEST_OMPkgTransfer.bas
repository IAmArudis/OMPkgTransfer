Attribute VB_Name = "TEST_OMPkgTransfer"
'@Folder("OMPkgTransfer")
'##############################################################################################################################################################################
'  ## DOCUMENTATION ##
'
'           Title:   TEST_OMPkgTransfer
'          Author:   Damian J Roberts
'          Github:   https://github.com/IAmArudis/OMPkgTransfer
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  ## REFERENCES ##
'
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  ## DESCRIPTION ##
'
'         PURPOSE:   Test module.  These are NOT unit tests, just random test methods to test certain functions / classes.
'
'    INSTRUCTIONS:   These methods are NOT for production and will not be exposed to the consuming code.
'
'        OVERRIDE:   N/A.
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


Private Const CLASS_NAME = "TEST_OMPkgTransfer"


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
'@Description ("Reads the latest entry from the test log file, runs the translator chain, and prints TransferData fields to the Immediate window for inspection.")
Public Sub TEST_GetTransferData()
   '  --- Load log data ---
   Dim oAllLogData As Collection
   Set oAllLogData = GetAllLogData()
   If oAllLogData Is Nothing Then
      Debug.Print "TEST_GetTransferData: GetAllLogData() returned Nothing."
      Exit Sub
   End If
   If oAllLogData.Count = 0 Then
      Debug.Print "TEST_GetTransferData: Log file is empty."
      Exit Sub
   End If

   '  --- Get the latest entry ---
   Dim oEntryData As OMLogger.LogEntryData
   Set oEntryData = GetLatestLogData(oAllLogData)
   If oEntryData Is Nothing Then
      Debug.Print "TEST_GetTransferData: No matching log entry found."
      Exit Sub
   End If

   '  --- Parse the log version from the entry ---
   Dim oLogVersion As CSVersion
   Set oLogVersion = CSVersion.CreateFromString(oEntryData.Data.GetFirst("version").Value)
   If Not oLogVersion.IsInitialized Then
      Debug.Print "TEST_GetTransferData: Could not parse version from log entry."
      Exit Sub
   End If
   Debug.Print "TEST_GetTransferData: Log version is " & oLogVersion.ToString() & "."

   '  --- Target version is hardcoded for testing purposes ---
   Dim oNeededVersion As CSVersion
   Set oNeededVersion = CSVersion.Create(4, 86, 12)

   '  --- Run the translator chain ---
   Dim oResult As Object
   Set oResult = TransferEngine.GetTransferObject(oEntryData, oLogVersion.Clone, oNeededVersion.Clone)
   If oResult Is Nothing Then
      Debug.Print "TEST_GetTransferData: TransferEngine returned Nothing for version " & oLogVersion.ToString() & "."
      Exit Sub
   End If

   '  --- Cast to the current container type ---
   Dim oData As TransferData_v4_86
   Set oData = oResult
   
   Call PrintTransfer_v4_86(oData)
End Sub





'##############################################################################################################################################################################
'  PRIVATE METHODS
'##############################################################################################################################################################################
'@Description ("Returns the most recent LogEntryData, optionally filtered by message.  Returns nothing if a matching entry isn't found.")
Private Function GetLatestLogData(oLogData As Collection, Optional ByVal sMessage As String = vbNullString) As OMLogger.LogEntryData
   If oLogData Is Nothing Then Exit Function
   If oLogData.Count = 0 Then Exit Function
   Dim oEntryData             As OMLogger.LogEntryData
   Dim oMostRecent            As OMLogger.LogEntryData
   Dim oMostRecentTimeStamp   As CSDateTime
   For Each oEntryData In oLogData
      If sMessage = vbNullString Or oEntryData.Message = sMessage Then
         If oMostRecent Is Nothing Then
            Set oMostRecent = oEntryData
            Set oMostRecentTimeStamp = CSDateTime.CreateFromString(oEntryData.TimeStamp)
         Else
            If CSDateTime.CreateFromString(oEntryData.TimeStamp).AsDateTime > oMostRecentTimeStamp.AsDateTime Then
               Set oMostRecent = oEntryData
               Set oMostRecentTimeStamp = CSDateTime.CreateFromString(oEntryData.TimeStamp)
            End If
         End If
      End If
   Next oEntryData
   Set GetLatestLogData = oMostRecent
End Function





'@Description ("Uses a LogReader to pull data from a ProposalTransfer.log.")
Private Function GetAllLogData() As Collection
   Const LOG_FILE As String = "R:\Users\DRoberts\Repositories\OMPkgTransfer\Testing Folders\.logs\TECHSUPPORT\ProposalTransfer.log"
   Dim oData As Collection
   Set oData = OMLogger.LogReader.ReadFile(LOG_FILE)
   Set GetAllLogData = oData
End Function











'  --------------------------------------------
'  TransferData Printing Methods
'  --------------------------------------------

'@Description ("Prints a specific version of a TransferData_vX_XX object.")
Private Sub PrintTransfer_v4_86(oData As TransferData_v4_86)
   '  ============================================================
   '  INSPECTION BLOCK - add / remove Debug.Print lines as needed.
   '  ============================================================
   Debug.Print "--- Version ---"
   Debug.Print "Container Version:  " & oData.Version.ToString()

   Debug.Print "--- Office Use ---"
   Debug.Print "Agent:              " & oData.OfficeUse.Agent
   Debug.Print "Minister:           " & oData.OfficeUse.Minister
   Debug.Print "Tour Info:          " & oData.OfficeUse.TourInfo
   Debug.Print "Budget:             " & oData.OfficeUse.Budget
   Debug.Print "Special Consid.:    " & oData.OfficeUse.SpecialConsiderations

   Debug.Print "--- Client 1 ---"
   Debug.Print "Name:               " & oData.Client1.ClientName
   Debug.Print "Phone:              " & oData.Client1.Phone
   Debug.Print "Email:              " & oData.Client1.Email
   Debug.Print "Instagram:          " & oData.Client1.Instagram
   Debug.Print "Address:            " & oData.Client1.Address1 & " " & oData.Client1.Address2
   Debug.Print "City/State/Zip:     " & oData.Client1.City & ", " & oData.Client1.State & " " & oData.Client1.Zip

   Debug.Print "--- Client 2 ---"
   Debug.Print "Name:               " & oData.Client2.ClientName
   Debug.Print "Phone:              " & oData.Client2.Phone
   Debug.Print "Email:              " & oData.Client2.Email
   Debug.Print "Instagram:          " & oData.Client2.Instagram
   Debug.Print "Address:            " & oData.Client2.Address1 & " " & oData.Client2.Address2
   Debug.Print "City/State/Zip:     " & oData.Client2.City & ", " & oData.Client2.State & " " & oData.Client2.Zip

   Debug.Print "--- Client 3 ---"
   Debug.Print "Name:               " & oData.Client3.ClientName
   Debug.Print "Phone:              " & oData.Client3.Phone
   Debug.Print "Email:              " & oData.Client3.Email
   Debug.Print "Instagram:          " & oData.Client3.Instagram
   Debug.Print "Address:            " & oData.Client3.Address1 & " " & oData.Client3.Address2
   Debug.Print "City/State/Zip:     " & oData.Client3.City & ", " & oData.Client3.State & " " & oData.Client3.Zip
   
   Debug.Print "--- Client 4 ---"
   Debug.Print "Name:               " & oData.Client4.ClientName
   Debug.Print "Phone:              " & oData.Client4.Phone
   Debug.Print "Email:              " & oData.Client4.Email
   Debug.Print "Instagram:          " & oData.Client4.Instagram
   Debug.Print "Address:            " & oData.Client4.Address1 & " " & oData.Client4.Address2
   Debug.Print "City/State/Zip:     " & oData.Client4.City & ", " & oData.Client4.State & " " & oData.Client4.Zip
   
   Debug.Print "--- Event ---"
   Debug.Print "Date:               " & oData.EventData.EventDate
   Debug.Print "Check In:           " & oData.EventData.CheckInTime
   Debug.Print "Check Out:          " & oData.EventData.CheckOutTime
   Debug.Print "Ceremony Time:      " & oData.EventData.CeremonyTime
   Debug.Print "Headcount:          " & oData.EventData.Headcount
   Debug.Print "Package Type:       " & oData.EventData.PackageType
   Debug.Print "Event Location:     " & oData.EventData.EventLocation

   Debug.Print "--- Food ---"
   Debug.Print "Food Option:        " & oData.FoodData.FoodOption
   Debug.Print "Food Price:         " & oData.FoodData.FoodPrice
   Debug.Print "Cake Price:         " & oData.FoodData.CakePrice
   Debug.Print "Punch Price:        " & oData.FoodData.PunchPrice

   Debug.Print "--- Optionals ---"
   Debug.Print "Bar Rental:         " & oData.OptionalsData.BarRental
   Debug.Print "Beverage Tower:     " & oData.OptionalsData.BeverageTower
   Debug.Print "Silver Service:     " & oData.OptionalsData.SilverService
   Debug.Print "Candy Station:      " & oData.OptionalsData.CandyStation
   Debug.Print "Chocolate Fountain: " & oData.OptionalsData.ChocolateFountain
   Debug.Print "Grooms Cake:        " & oData.OptionalsData.GroomsCake
   Debug.Print "Platinum Package:   " & oData.OptionalsData.PlatinumPackage
   Debug.Print "Rcptn Centerpieces: " & oData.OptionalsData.ReceptionCenterpieces
   Debug.Print "Dance Floor Mono:   " & oData.OptionalsData.DanceFloorMonogram
   Debug.Print "Video Monitor:      " & oData.OptionalsData.VideoMonitorRental
   Debug.Print "Minister:           " & oData.OptionalsData.Minister
   Debug.Print "Wedding Flowers:    " & oData.OptionalsData.WeddingPartyFlowers
   Debug.Print "Rehearsal:          " & oData.OptionalsData.Rehearsal
   Debug.Print "Video Ceremony:     " & oData.OptionalsData.VideoCeremony
   Debug.Print "Cold Sparklers:     " & oData.OptionalsData.ColdSparklers
   Debug.Print "Cloud Machine:      " & oData.OptionalsData.CloudMachine

   Debug.Print "--- Misc (Taxable) ---"
   Debug.Print "Misc Tax 1:         " & oData.OptionalsData.MiscTax1.Qty & "x " & oData.OptionalsData.MiscTax1.Label & " @ " & oData.OptionalsData.MiscTax1.Price
   Debug.Print "Misc Tax 2:         " & oData.OptionalsData.MiscTax2.Qty & "x " & oData.OptionalsData.MiscTax2.Label & " @ " & oData.OptionalsData.MiscTax2.Price
   Debug.Print "Misc Tax 3:         " & oData.OptionalsData.MiscTax3.Qty & "x " & oData.OptionalsData.MiscTax3.Label & " @ " & oData.OptionalsData.MiscTax3.Price

   Debug.Print "--- Misc (Non-Taxable) ---"
   Debug.Print "Misc NoTax 1:       " & oData.OptionalsData.MiscNoTax1.Qty & "x " & oData.OptionalsData.MiscNoTax1.Label & " @ " & oData.OptionalsData.MiscNoTax1.Price
   Debug.Print "Misc NoTax 2:       " & oData.OptionalsData.MiscNoTax2.Qty & "x " & oData.OptionalsData.MiscNoTax2.Label & " @ " & oData.OptionalsData.MiscNoTax2.Price
   Debug.Print "Misc NoTax 3:       " & oData.OptionalsData.MiscNoTax3.Qty & "x " & oData.OptionalsData.MiscNoTax3.Label & " @ " & oData.OptionalsData.MiscNoTax3.Price
   '  ============================================================
   '  /INSPECTION BLOCK
   '  ============================================================
End Sub





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
