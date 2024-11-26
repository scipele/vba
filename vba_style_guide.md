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
| Description                                    | Example                               |
|------------------------------------------------|---------------------------------------|
| Sub Naming, Use Verb-Noun Structure            | `PascalCase` -> GenerateReport        |
| Function Naming, Use Verb-Noun Structure       | `PascalCase` -> ExportDataToCSV       |
| Local Variable Names                           | `snake_case`                          |
| Constants                                      | `ALL_CAP_SNAKE_CASE`                  |
| Labels                                         | `PascalCaseReportNoScoreLabel`        |
| Error Handlers                                 | `PascalCase` -> Err_CalculateTotals   |
| Class Naming, Use a Noun for Name              | `PascalCase` -> DataExporter          |
| Class Methods                                  | `PascalCase`                          |
| Class Private Members                          | `camelCase` -> _internalData          |
| Class Properties                               | `PascalCase` -> TotalAmount           |
| Enums -> Name Them with CamelCase              | `scoreType`                           |
| Enum Elements -> Prefix Them with Short Name   | `stPrevScoreEmptyExceptFrameOne`      |
| For event handlers, follow VBA's convention    | `hyphenated PascalCase Button_Click` |
	
## 4. Clear Listing of Sub and Function Parameters
| Description                                              | Example                                             |
|----------------------------------------------------------|-----------------------------------------------------|
| **Function Parameters**: Use Line Break so that each one can be on a separate line | ` _` |
| **Parameter Type**: Always specify the type, even for `Variant` | `ByVal prev_score As Variant`                      |
| **Pass by Value or Reference**: Explicitly declare `ByVal` or `ByRef` | `ByVal prev_score As Variant` or `ByRef prev_score As Variant` |
| **Return Type**: Always declare the return type | `As Variant` (for returning a value) or `Sub` (for Subs that don’t return) |

- Example:
`code`
Function FindColumnByLabel(ByVal label As String, _
                           ByVal searchRow As Long, _
						   ByVal shtName As String) _
						   As Long'