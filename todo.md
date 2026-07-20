# feature/adding-qtys-to-food-details

## DS_TransferTypes
### Updated TOfficeUseData_v4_86
Added new field:
- `Budget`
### Updated `TFoodData_v4_86`
Added new fields:
- `FoodQty`
- `CakeQty`
- `PunchQty`
### Updated `TOptionalsData_v4_86`
- Changed all of the data types to `Long` except `CloudSparklers` and `CloudMachine`.
- Added note about all of these being Qty indicators.

## `Translator_v4_86`
- Updated `GetFoodData()` to also handle the qty values for the above mentioned new fields.
- Updated `GetOfficeUseData()` to also handle the above mentioned new fields.
