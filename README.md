# OMPkgTransfer

A forward-compatibility translation library for `ProposalTransfer.log` data.  Allows older versions of PkgDocs to have their log data imported into the current version without any changes to the consuming project.

---

## Overview

When PkgDocs exports a proposal to a transfer log, it writes a versioned JSON payload.  If the structure of that payload changes across PkgDocs versions, a naive import would break or silently drop fields.  OMPkgTransfer solves this by maintaining a chain of versioned translators — one per structural change — that progressively reshape older data into the current version's shape.

The consuming project (`PkgDocs`) only ever interacts with two things: `TransferEngine` (the single entry point) and the `TransferData_vX_XX` container that comes back from it.

---

## Object Graph

```
PkgDocs (consumer)
│
└── TransferEngine                          [Entry point.  PredeclaredId.  Exposed.]
    │
    ├── TranslatorRegistry                  [Holds all translators oldest-to-newest.  PredeclaredId.  Internal.]
    │   └── IProposalTranslator             [Interface contract.  Internal.]
    │       └── Translator_v4_86            [Concrete implementation.  PredeclaredId.  Internal.]
    │           ├── DS_EnumMappers          [String → Long mappers.  Private module.  Internal.]
    │           └── DS_TransferTypes        [Versioned UDT definitions.  Exposed.]
    │
    └── TransferData_v4_86                  [Data container.  Returned as Object, cast by consumer.  Exposed.]
        ├── DS_TransferTypes                [UDT definitions used as property types.]
        └── CSVersion (VBACodeSnippets)     [Version stamp.]
```

### Exposure Summary

| Module / Class | Exposed | Notes |
|---|---|---|
| `TransferEngine` | ✅ | Primary entry point for all consumers |
| `TransferData_v4_86` | ✅ | Result container; cast to this type after `GetTransferObject()` |
| `DS_TransferTypes` | ✅ | UDT definitions; required by consumers to read container fields |
| `IProposalTranslator` | ❌ | Internal interface; never referenced by consumers |
| `Translator_v4_86` | ❌ | Internal implementation; never referenced by consumers |
| `TranslatorRegistry` | ❌ | Internal registry; driven entirely by `TransferEngine` |
| `DS_EnumMappers` | ❌ | Private module; internal to `Translator_vX_XX` only |

---

## Modules and Classes

---

### `TransferEngine`
**PredeclaredId. Exposed.**

The single entry point for all consuming code.  Orchestrates the full translator chain internally — the consumer never interacts with any translator, registry, or interface directly.

#### `GetTransferObject(oEntryData, oLogVersion, oNeededVersion) As Object`
Walks the translator chain and returns a fully translated `TransferData_vX_XX` container as `Object`.

| Parameter | Type | Description |
|---|---|---|
| `oEntryData` | `OMLogger.LogEntryData` | The raw log entry to translate |
| `oLogVersion` | `CSVersion` | The version the log data originated from.  Must be a clone. |
| `oNeededVersion` | `CSVersion` | The version the caller needs the data translated to.  Must be a clone. |

**Returns** `Nothing` if no translator is found for `oLogVersion`.  The caller is responsible for handling a `Nothing` return gracefully.

**Typical usage in PkgDocs:**
```vb
Dim oResult As Object
Set oResult = TransferEngine.GetTransferObject(oEntryData, oLogVersion.Clone, oNeededVersion.Clone)
If oResult Is Nothing Then
   '  Handle gracefully.
   Exit Sub
End If
Dim oData As TransferData_v4_86
Set oData = oResult
```

---

### `TransferData_v4_86`
**Not PredeclaredId.  Exposed.**

Thin data container for the v4.86.x `ProposalTransfer` data structures.  Holds all versioned UDTs from `DS_TransferTypes` privately, exposed via `Public Property Get` (readable by all consumers) and `Friend Property Let/Set` (writable only within `OMPkgTransfer`).

Exists solely to allow the UDT payload to cross `Object` boundaries in the `IProposalTranslator` interface.  This class has no methods — all logic lives in the concrete `Translator_vX_XX` implementations.

#### Properties

| Property | Type | Access | Description |
|---|---|---|---|
| `Version` | `CSVersion` | Public Get / Friend Set | Version stamp applied by the translator chain |
| `OfficeUse` | `TOfficeUseData_v4_86` | Public Get / Friend Let | Office-use block (agent, minister, tour info, special considerations) |
| `Client1` | `TClientData_v4_86` | Public Get / Friend Let | Primary client |
| `Client2` | `TClientData_v4_86` | Public Get / Friend Let | Second client |
| `Client3` | `TClientData_v4_86` | Public Get / Friend Let | Third client |
| `Client4` | `TClientData_v4_86` | Public Get / Friend Let | Fourth client |
| `EventData` | `TEventData_v4_86` | Public Get / Friend Let | Event details (date, times, headcount, package type, location) |
| `FoodData` | `TFoodData_v4_86` | Public Get / Friend Let | Catering details (option, prices) |
| `OptionalsData` | `TOptionalsData_v4_86` | Public Get / Friend Let | Optional add-ons and misc line items |

> **When a new translator version is needed:** copy this class, update `TClass` members and properties to reflect any UDT type changes, and freeze this class.  Do NOT modify it once a newer `TransferData_vX_XX` exists.

---

### `DS_TransferTypes`
**Standard module.  Exposed.**

Versioned UDT definitions for all `ProposalTransfer` data structures.  Each structural change to the transfer log gets its own set of `T*Data_vX_XX` types here.  All types are frozen once written.

#### v4.86.x Types

| Type | Fields |
|---|---|
| `TOfficeUseData_v4_86` | `Agent`, `Minister`, `TourInfo`, `SpecialConsiderations` |
| `TClientData_v4_86` | `Title`, `ClientName`, `Phone`, `Email`, `Instagram`, `Address1`, `Address2`, `City`, `State`, `Zip` |
| `TEventData_v4_86` | `EventDate`, `CheckInTime`, `CheckOutTime`, `CeremonyTime`, `Headcount`, `PackageType`*, `EventLocation`* |
| `TFoodData_v4_86` | `FoodOption`*, `FoodPrice`, `CakePrice`, `PunchPrice` |
| `TMiscItemData_v4_86` | `Qty`, `Label`, `Price` |
| `TOptionalsData_v4_86` | All optional add-on fields + 6x `TMiscItemData_v4_86` (3 taxable, 3 non-taxable) |

*\* Stored as `Long`.  Enum-backed fields are decoupled from PkgDocs' enum definitions intentionally — see `DS_EnumMappers` for the mapping.*

> **When adding a new version:** add a new versioned block below the previous one.  Copy only the types that changed and increment their suffix.  Unchanged types keep their previous version's name.  Never remove or modify existing versioned types — they are frozen history.

---

### `IProposalTranslator`
**Interface.  Not PredeclaredId.  Internal.**

Contract for all versioned translator implementations.  Each concrete `Translator_vX_XX` class implements this interface.

#### Methods

| Method | Signature | Description |
|---|---|---|
| `CanHandle` | `(oVersion As CSVersion) As Boolean` | Returns `True` if this translator can handle the given version |
| `Process` | `(oEntryData As LogEntryData, oVersion As CSVersion) As Object` | Reads raw log data and returns a fully populated `TransferData_vX_XX` as `Object` |
| `Translate` | `(oData As Object) As Object` | Converts the incoming container to this version's shape and stamps the version forward |

> `Apply()` is intentionally absent.  It lives in PkgDocs, which has native access to all worksheet targets.

---

### `Translator_v4_86`
**PredeclaredId.  Internal.**

Handles `ProposalTransfer.log` data originating from PkgDocs v4.86.0 and later, until a future version's `CanHandle()` takes over.  This is the baseline translator — all subsequent translators derive their `Get*Data()` logic from this one.

#### Version Constants

| Constant | Default | Description |
|---|---|---|
| `MIN_MAJOR` | `4` | Minimum major version this translator handles |
| `MIN_MINOR` | `86` | Minimum minor version this translator handles |
| `MIN_PATCH` | `0` | Minimum patch version this translator handles |
| `MAX_MAJOR` | `-1` | `-1` = open-ended (latest translator).  Set to concrete value when a newer translator takes over. |
| `MAX_MINOR` | `-1` | See `MAX_MAJOR` |
| `MAX_PATCH` | `-1` | See `MAX_MAJOR` |

#### `CanHandle(oVersion) As Boolean`
Returns `True` if this translator owns the given version.  First rejects anything older than the floor unconditionally.  If `MAX_MAJOR = -1`, accepts any version at or above the floor.  Otherwise checks `[MIN, MAX]` inclusive.

#### `Process(oEntryData, oVersion) As Object`
Reads the `LogEntryData` and returns a fully populated `TransferData_v4_86` as `Object`.  Calls each `Get*Data()` method internally to build the container.

#### `Translate(oData) As Object`
Baseline pass-through — no prior version to convert from.  When `MAX_MAJOR = -1`, stamps `v999.999.999` to signal end of chain.  When `MAX_*` are concrete, stamps `MAX_MAJOR.MAX_MINOR.(MAX_PATCH + 1)` to hand off to the next translator.

#### `Get*Data()` Methods
Public typed parsers — not on the interface.  Future translators can call these directly by class name for delegation rather than duplicating logic.

| Method | Returns | Description |
|---|---|---|
| `GetOfficeUseData(oKVCol)` | `TOfficeUseData_v4_86` | Parses the `office-use` block |
| `GetClientData(oKVCol)` | `TClientData_v4_86` | Parses a `client` block |
| `GetEventData(oKVCol)` | `TEventData_v4_86` | Parses the `event-details` block |
| `GetFoodData(oKVCol)` | `TFoodData_v4_86` | Parses the `food-details` block |
| `GetOptionalsData(oKVCol)` | `TOptionalsData_v4_86` | Parses the `main-optionals` block (excluding Misc items) |
| `GetMiscItemData(oKVCol)` | `TMiscItemData_v4_86` | Parses a single misc line item block |

All `Get*Data()` methods use `On Error Resume Next` to handle missing keys gracefully — any field whose key is absent in the log stays at its default value (`vbNullString` / `0`).

> **Delegation pattern for future translators:**
> ```vb
> GetClientData = Translator_v4_86.GetClientData(oKVCol)
> With GetClientData
>    .SomeNewField = oKVCol.GetFirst("new-key").Value
> End With
> ```
> If a change doesn't involve reading new keys, it belongs in `Translate()` instead.

---

### `DS_EnumMappers`
**Standard module.  Private (`Option Private Module`).  Internal.**

Maps raw `ProposalTransfer.log` string values to their corresponding `Long` enum values.  Returns `Long` rather than named enum types to stay fully decoupled from PkgDocs' enum definitions.  `0` is reserved as the unset/error sentinel across all enums — any unrecognized string returns `0`.

These values are **frozen** once a version of PkgDocs has written them to a log file.  Never renumber or remove a case.

#### Methods

| Method | Description |
|---|---|
| `GetFoodEnum(Value)` | Maps catering option string → `Long` (`None`=1, `Appetizer`=2, `Plated Dinner`=3, `Dinner Plus`=4) |
| `GetLocationEnum(Value)` | Maps ceremony location string → `Long` (`No Ceremony`=1, `Chapel`=2, `Garden`=3, `Fire Place`=4) |
| `GetEventTypeEnum(Value)` | Maps event type string → `Long` (`Ceremony Only`=1, `Reception Only`=2, `Ceremony and Reception`=3) |
| `GetPackageEnum(Value)` | Maps package type string → `Long` (`Wedding Package`=1 … `Memorial Service Package`=5) |
| `GetRehearsalTypeEnum(Value)` | Maps rehearsal type string → `Long` (`None`=1, `30-Mins Prior`=2, `Thursday`=3, `Friday`=4, `Other`=5) |

---

## Adding a New Translator Version

When a structural change to the transfer log requires a new translator:

1. **`DS_TransferTypes.bas`** — Add a new versioned block.  Copy only the types that changed and increment their suffix.  Leave all previous blocks untouched.

2. **`TransferData_vX_XX.cls`** — Copy `TransferData_v4_86.cls`.  Update `TClass` members and properties to reflect changed UDT types.  Unchanged members keep their previous version's type.

3. **`Translator_vX_XX.cls`** — Copy `Translator_v4_86.cls`.  Set `MIN_*` to the first version this translator handles.  Leave `MAX_*` at `-1`.  Update `Get*Data()` methods — delegate to the previous translator where unchanged, rewrite only what changed.  Update `Translate()` to remap any structural differences from the previous version's container.

4. **`Translator_v4_86.cls`** (previous latest) — Set `MAX_MAJOR`, `MAX_MINOR`, `MAX_PATCH` to the last version it handles.

5. **`TranslatorRegistry.cls`** — Add the new `Translator_vX_XX` at the bottom of `RegisterAll()`.  That is the only change needed here.

6. **`DS_EnumMappers.bas`** — Add any new enum cases at the bottom of the relevant `Select` block.  Never renumber or remove existing cases.
