# Sciple's VBA Style Guide

## 1. General
- Always Use Option Explicit
- Two Spaces before Functions including any comments
- Indentation set to four spaces i.e. default

## 2. Module Documentation
- filename:     sampleFilename.vba
 - Purpose:
 - Inputs:
 - Outputs:
 - Dependencies: None
 - By:  T.Sciple, MM/DD/YYYY

## 3. Naming of Subs, Functions, Variables, Constants, Class Items
| Description                                           | Example                               |
|-------------------------------------------------------|---------------------------------------|
| `Sub Naming, Use Verb-Noun Structure                 `| `PascalCase -> GenerateReport        `|
| `Function Naming, Use Verb-Noun Structure            `| `PascalCase -> ExportDataToCSV       `|
| `Function Parameter Names                            `| `camelCase -> productPrice           `|
| `Local Variable Names                                `| `snake_case                          `|
| `Constants                                           `| `ALL_CAP_SNAKE_CASE                  `|
| `Labels                                              `| `PascalCaseReportNoScoreLabel        `|
| `Error Handlers                                      `| `hyphenated Pascal -> Err_CalcTotals `|
| `Class Naming, Use a Noun for Name                   `| `PascalCase-> DataExporter           `|
| `Class Methods                                       `| `PascalCase                          `|
| `Class Private Members                               `| `camelCase -> _internalData          `|
| `Class Properties                                    `| `PascalCase -> TotalAmount           `|
| `Enums -> Name Them with camelCase                   `| `camelCase -> scoreType              `|
| `Enum Elements -> Prefix Them with Short Name        `| `stPrevScoreEmptyExceptFrameOne      `|
| `For event handlers, follow VBA's convention         `| `hyphenated PascalCase Button_Click  `|
 	
## 4. Clear Listing of Sub and Function Parameters
| Description                                           | Example                               |
|-------------------------------------------------------|---------------------------------------|
| `Function Parameters: Use Line Break Per Parameter   `| `ByVal user_input As Variant, _      `|
| `Parameter Type: Always specify the data type        `| `ByVal prev_score As Variant         `|
| `Pass by Value or Reference: Explicitly declare      `| `ByVal prev_score As Variant         `|
| `Function Return Type: Always declare return type    `| `As Variant (for returning  value)   `|

 - Example:
```vba
Function FindColumnByLabel(ByVal label As String, _  
                           ByVal searchRow As Long, _  
                           ByVal shtName As String) As Long
    ' Function body
End Function