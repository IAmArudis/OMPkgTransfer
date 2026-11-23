# feature/adding-qtys-to-food-details

## DS_TransferTypes
### Updated TOfficeUseData_v4_86
The fields we were using were combining too much information and honestly its better to separate out the information so we don't have to translate into pkg docs.

<br>

Updated fields:
- `Toured`
- `TouringAgent`
- `TourDate`
- `BookingAgent`
- `Budget`
- `MonthlyPayment`
- `Minister`
- `SpecialConsiderations`

<br>

### Updated `TFoodData_v4_86`
Added new fields:
- `FoodQty`
- `CakeQty`
- `PunchQty`
### Updated `TOptionalsData_v4_86`
- Changed all of the data types to `Long` except `CloudSparklers` and `CloudMachine`.
- Added note about all of these being Qty indicators.

### Updated `TEventData_v4_86`
- Added `BookingDate` to the type.

## `Translator_v4_86`
- Updated `GetFoodData()` to also handle the qty values for the above mentioned new fields.
- Updated `GetOfficeUseData()` to also handle the above mentioned new fields.
- Updated `GetEventData()` to also handle the above mentioned new fields.