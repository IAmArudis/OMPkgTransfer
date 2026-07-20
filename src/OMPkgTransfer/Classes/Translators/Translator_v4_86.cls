VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Translator_v4_86"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "OMPkgTransfer.Classes.Translators"
'@IgnoreModule UnrecognizedAnnotation
'@PredeclaredId
'##############################################################################################################################################################################
'  ## DOCUMENTATION ##
'
'           Title:   Translator_v4_86
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
'         PURPOSE:   Handles ProposalTransfer.log data originating from PkgDocs v4.86.0 and later, until a future version's CanHandle() takes over.
'
'                    This is the baseline translator - all subsequent translators derive their Get*Data() logic from this one, either by direct delegation or by translating a
'                    prior version's UDT into their own.
'
'    INSTRUCTIONS:   This class is consumed exclusively by TransferCode via TranslatorRegistry.  Get*Data() methods are PUBLIC so that future translators can call them
'                    directly by name for delegation purposes without going through the interface.
'
'          NOTES:    CanHandle() is currently open-ended (no upper bound).  When a new translator is created, this method will be given a concrete upper bound and the new
'                    translator will take over as the open-ended handler.  Update MAX_PATCH and add MAX_MINOR / MAX_MAJOR constants as needed at that time.
'
'                    Translate() is effectively a pass-through for this version since it is the baseline - there is no prior version to convert from.  It stamps the version
'                    to MAX_PATCH + 1 so TransferCode knows to move to the next translator.
'
'                    Apply() is NOT on this class or the interface.  It lives in TransferCode inside PkgDocs, which has native access to all worksheet targets.
'
'        OVERRIDE:   CanHandle():   When a new translator is created, give this a concrete upper bound and update MAX_PATCH, MAX_MINOR, and MAX_MAJOR.
'                    Translate():   When a new translator is created, add the code needed to translate the previous version into the latest version (ie. missing members,
'                                   updated enum values, etc.
'                    Get*Data():    When a new translator is created, if the previous version of this method would still work for the new version, it can be invoked directly
'                                   rather than copied by invoking the previous version's translator (Translator_vX_XX) where vX_XX would be the DIRECTLY PREVIOUS translator
'                                   version.  All updates to the information that would need to be retrieved from the specific oKVCol for the new version will need to be
'                                   updated AFTER the previous version's Get*Data() call.  (ie):
'                                      GetClientData = Translator_vX_XX.GetClientData(args)
'                                      With GetClientData
'                                         .SomeNewField = oKVCol.GetFirst("new-key").Value
'                                      End With
'                                   If it isn't specific to the oKVCol data ie. structural changes that don't involve reading new keys (e.g. remapping an enum value that
'                                   changed meaning, splitting a field, computing a derived value) belong in Translate(), not Get*Data().
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'##############################################################################################################################################################################

'##############################################################################################################################################################################
'  PRE-PROCESSORS, IMPLEMENTATION STATEMENTS, SANITY STATEMENTS, CONSTANTS, TYPES, ENUMS
'##############################################################################################################################################################################
Option Explicit
Option Base 1


Implements IProposalTranslator


Private Type TClass
   Value As Variant
End Type


Private Const MIN_MAJOR  As Long = 4
Private Const MIN_MINOR  As Long = 86
Private Const MIN_PATCH  As Long = 0
Private Const MAX_MAJOR  As Long = -1   '@Override - set to the last Major this translator handles when a newer translator takes over.
Private Const MAX_MINOR  As Long = -1   '@Override - set to the last Minor this translator handles when a newer translator takes over.
Private Const MAX_PATCH  As Long = -1   '@Override - set to the last Patch this translator handles when a newer translator takes over.


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
'@Override
'@Description ("Returns True if this translator can handle the given version.  When MAX_MAJOR is -1 this translator is the open-ended latest handler and accepts any version at or above the floor.  When MAX_* are concrete, the range is [MIN, MAX] inclusive.")
Public Function CanHandle(ByVal oVersion As CSVersion) As Boolean
   Dim oFloor   As CSVersion
   Set oFloor = CSVersion.Create(MIN_MAJOR, MIN_MINOR, MIN_PATCH)
   If oVersion.IsOlderThan(oFloor) Then
      CanHandle = False
      Exit Function
   End If
   If MAX_MAJOR = -1 Then
      CanHandle = True
      Exit Function
   End If
   Dim oCeiling As CSVersion
   Set oCeiling = CSVersion.Create(MAX_MAJOR, MAX_MINOR, MAX_PATCH)
   CanHandle = Not oVersion.IsNewerThan(oCeiling)
End Function










'@Description ("Parses a KVCollection into a TClientData_v4_86.")
Public Function GetClientData(ByVal oKVCol As OMLogger.KVCollection) As TClientData_v4_86
   If oKVCol Is Nothing Then Exit Function
   On Error Resume Next
   With GetClientData
      .Title = oKVCol.GetFirst("title").Value
      .ClientName = oKVCol.GetFirst("clientname").Value
      .Phone = oKVCol.GetFirst("phone").Value
      .Email = oKVCol.GetFirst("email").Value
      .Instagram = oKVCol.GetFirst("instagram").Value
      .Address1 = oKVCol.GetFirst("address1").Value
      .Address2 = oKVCol.GetFirst("address2").Value
      .City = oKVCol.GetFirst("city").Value
      .State = oKVCol.GetFirst("state").Value
      .Zip = oKVCol.GetFirst("zip").Value
   End With
   On Error GoTo 0
End Function










'@Description ("Parses a KVCollection into a TEventData_v4_86.")
Public Function GetEventData(ByVal oKVCol As OMLogger.KVCollection) As TEventData_v4_86
   If oKVCol Is Nothing Then Exit Function
   On Error Resume Next
   With GetEventData
      .EventDate = oKVCol.GetFirst("event-date").Value
      .CheckInTime = oKVCol.GetFirst("check-in-time").Value
      .CheckOutTime = oKVCol.GetFirst("check-out-time").Value
      .CeremonyTime = oKVCol.GetFirst("ceremony-time").Value
      .Headcount = oKVCol.GetFirst("headcount").Value
      .PackageType = DS_EnumMappers.GetPackageEnum(oKVCol.GetFirst("package-type").Value)
      .EventLocation = DS_EnumMappers.GetLocationEnum(oKVCol.GetFirst("event-location").Value)
   End With
   On Error GoTo 0
End Function










'@Description ("Parses a KVCollection into a TFoodData_v4_86.")
Public Function GetFoodData(ByVal oKVCol As OMLogger.KVCollection) As TFoodData_v4_86
   If oKVCol Is Nothing Then Exit Function
   On Error Resume Next
   With GetFoodData
      .FoodOption = DS_EnumMappers.GetFoodEnum(oKVCol.GetFirst("food-option").Value)
      .FoodQty = oKVCol.GetFirst("food-qty").Value
      .FoodPrice = oKVCol.GetFirst("food-price").Value
      .CakeQty = oKVCol.GetFirst("cake-qty").Value
      .CakePrice = oKVCol.GetFirst("cake-price").Value
      .PunchQty = oKVCol.GetFirst("punch-qty").Value
      .PunchPrice = oKVCol.GetFirst("punch-price").Value
   End With
   On Error GoTo 0
End Function










'@Description ("Parses a KVCollection into a TMiscItemData_v4_86.")
Public Function GetMiscItemData(ByVal oKVCol As OMLogger.KVCollection) As TMiscItemData_v4_86
   If oKVCol Is Nothing Then Exit Function
   On Error Resume Next
   With GetMiscItemData
      .Qty = oKVCol.GetFirst("qty").Value
      .Label = oKVCol.GetFirst("label").Value
      .Price = oKVCol.GetFirst("price").Value
   End With
   On Error GoTo 0
End Function










'@Description ("Parses a KVCollection into a TOfficeUseData_v4_86.")
Public Function GetOfficeUseData(ByVal oKVCol As OMLogger.KVCollection) As TOfficeUseData_v4_86
   If oKVCol Is Nothing Then Exit Function
   On Error Resume Next
   With GetOfficeUseData
      .Agent = oKVCol.GetFirst("agent").Value
      .Minister = oKVCol.GetFirst("minister").Value
      .Budget = oKVCol.GetFirst("budget").Value
      .TourInfo = oKVCol.GetFirst("tour-info").Value
      .SpecialConsiderations = oKVCol.GetFirst("special-considerations").Value
   End With
   On Error GoTo 0
End Function










'@Description ("Parses a KVCollection into a TOptionalsData_v4_86.")
Public Function GetOptionalsData(ByVal oKVCol As OMLogger.KVCollection) As TOptionalsData_v4_86
   If oKVCol Is Nothing Then Exit Function
   On Error Resume Next
   With GetOptionalsData
      .BarRental = oKVCol.GetFirst("bar-rental").Value
      .BeverageTower = oKVCol.GetFirst("beverage-tower").Value
      .SilverService = oKVCol.GetFirst("silver-service").Value
      .CandyStation = oKVCol.GetFirst("candy-station").Value
      .ChocolateFountain = oKVCol.GetFirst("chocolate-fountain").Value
      .GroomsCake = oKVCol.GetFirst("grooms-cake").Value
      .PlatinumPackage = oKVCol.GetFirst("platinum-package").Value
      .ReceptionCenterpieces = oKVCol.GetFirst("reception-centerpieces").Value
      .DanceFloorMonogram = oKVCol.GetFirst("dance-floor-monogram").Value
      .VideoMonitorRental = oKVCol.GetFirst("video-monitor-rental").Value
      .Minister = oKVCol.GetFirst("minister").Value
      .WeddingPartyFlowers = oKVCol.GetFirst("wedding-party-flowers").Value
      .Rehearsal = DS_EnumMappers.GetRehearsalTypeEnum(oKVCol.GetFirst("rehearsal").Value)
      .VideoCeremony = oKVCol.GetFirst("video-ceremony").Value
      .ColdSparklers = oKVCol.GetFirst("cold-sparklers").Value
      .CloudMachine = oKVCol.GetFirst("cloud-machine").Value
   End With
   On Error GoTo 0
End Function










'@Description ("Reads the LogEntryData and returns a fully populated TransferData_v4_86 as Object.")
Public Function Process(ByVal oEntryData As OMLogger.LogEntryData, ByVal oVersion As CSVersion) As Object
   Const OFC_USE      As String = "office-use"
   Const CLIENT_X     As String = "client"
   Const EVNT_DTLS    As String = "event-details"
   Const FOOD_DTLS    As String = "food-details"
   Const MAIN_OPTS    As String = "main-optionals"
   Const MISC_TAX_1   As String = "misc-tax-1"
   Const MISC_TAX_2   As String = "misc-tax-2"
   Const MISC_TAX_3   As String = "misc-tax-3"
   Const MISC_NOTAX_1 As String = "misc-notax-1"
   Const MISC_NOTAX_2 As String = "misc-notax-2"
   Const MISC_NOTAX_3 As String = "misc-notax-3"

   Dim oData     As TransferData_v4_86
   Dim oRootCol  As OMLogger.KVCollection
   Dim oKVPair   As OMLogger.KVPair
   Dim tOptionals As TOptionalsData_v4_86

   Set oData = New TransferData_v4_86
   Set oData.Version = oVersion.Clone
   Set oRootCol = oEntryData.Data

   Set oKVPair = oRootCol.GetFirst(OFC_USE)
   If Not oKVPair Is Nothing Then oData.OfficeUse = Me.GetOfficeUseData(oKVPair.Value)

   Set oKVPair = oRootCol.GetByOccurrence(CLIENT_X, 1)
   If Not oKVPair Is Nothing Then oData.Client1 = Me.GetClientData(oKVPair.Value)

   Set oKVPair = oRootCol.GetByOccurrence(CLIENT_X, 2)
   If Not oKVPair Is Nothing Then oData.Client2 = Me.GetClientData(oKVPair.Value)

   Set oKVPair = oRootCol.GetByOccurrence(CLIENT_X, 3)
   If Not oKVPair Is Nothing Then oData.Client3 = Me.GetClientData(oKVPair.Value)

   Set oKVPair = oRootCol.GetByOccurrence(CLIENT_X, 4)
   If Not oKVPair Is Nothing Then oData.Client4 = Me.GetClientData(oKVPair.Value)

   Set oKVPair = oRootCol.GetFirst(EVNT_DTLS)
   If Not oKVPair Is Nothing Then oData.EventData = Me.GetEventData(oKVPair.Value)

   Set oKVPair = oRootCol.GetFirst(FOOD_DTLS)
   If Not oKVPair Is Nothing Then oData.FoodData = Me.GetFoodData(oKVPair.Value)

   '  Optionals and its nested Misc items are built into a local UDT first,
   '  then stamped onto the container in one Friend Let call.  This is required
   '  because chained assignment (oData.OptionalsData.MiscTax1 = ...) is not
   '  possible through a Property - the Get returns a copy, not a reference.
   Set oKVPair = oRootCol.GetFirst(MAIN_OPTS)
   If Not oKVPair Is Nothing Then tOptionals = Me.GetOptionalsData(oKVPair.Value)

   Set oKVPair = oRootCol.GetFirst(MISC_TAX_1)
   If Not oKVPair Is Nothing Then tOptionals.MiscTax1 = Me.GetMiscItemData(oKVPair.Value)

   Set oKVPair = oRootCol.GetFirst(MISC_TAX_2)
   If Not oKVPair Is Nothing Then tOptionals.MiscTax2 = Me.GetMiscItemData(oKVPair.Value)

   Set oKVPair = oRootCol.GetFirst(MISC_TAX_3)
   If Not oKVPair Is Nothing Then tOptionals.MiscTax3 = Me.GetMiscItemData(oKVPair.Value)

   Set oKVPair = oRootCol.GetFirst(MISC_NOTAX_1)
   If Not oKVPair Is Nothing Then tOptionals.MiscNoTax1 = Me.GetMiscItemData(oKVPair.Value)

   Set oKVPair = oRootCol.GetFirst(MISC_NOTAX_2)
   If Not oKVPair Is Nothing Then tOptionals.MiscNoTax2 = Me.GetMiscItemData(oKVPair.Value)

   Set oKVPair = oRootCol.GetFirst(MISC_NOTAX_3)
   If Not oKVPair Is Nothing Then tOptionals.MiscNoTax3 = Me.GetMiscItemData(oKVPair.Value)

   oData.OptionalsData = tOptionals

   Set Process = oData
End Function











'@Description ("Baseline translator - no prior version to convert from.  When MAX_MAJOR is -1, stamps version to v999.999.999 to signal end of chain.  When MAX_* are concrete, stamps to MAX_MAJOR.MAX_MINOR.(MAX_PATCH + 1) to hand off to the next translator.")
Public Function Translate(ByVal oData As Object) As Object
   Dim oIn As TransferData_v4_86
   Set oIn = oData
   '  Baseline - no structural conversion needed.  Future translators perform field mapping here.
   If MAX_MAJOR = -1 Then
      Set oIn.Version = CSVersion.Create(999, 999, 999)
   Else
      Set oIn.Version = CSVersion.Create(MAX_MAJOR, MAX_MINOR, MAX_PATCH + 1)
   End If
   Set Translate = oIn
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
Friend Property Get Self() As Translator_v4_86
   Set Self = Me
End Property


'##############################################################################################################################################################################
'  IProposalTranslator IMPLEMENTATION
'##############################################################################################################################################################################
Private Function IProposalTranslator_CanHandle(ByVal oVersion As CSVersion) As Boolean
   IProposalTranslator_CanHandle = Me.CanHandle(oVersion)
End Function

Private Function IProposalTranslator_Process(ByVal oEntryData As OMLogger.LogEntryData, ByVal oVersion As CSVersion) As Object
   Set IProposalTranslator_Process = Me.Process(oEntryData, oVersion)
End Function

Private Function IProposalTranslator_Translate(ByVal oData As Object) As Object
   Set IProposalTranslator_Translate = Me.Translate(oData)
End Function


'##############################################################################################################################################################################
'  IClass IMPLEMENTATION
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'  OLD CODE / TEMP CODE / DEBUG CODE
'     Please place retired code here.
'##############################################################################################################################################################################

