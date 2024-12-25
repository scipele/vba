# Sciple's VBA Style Guide

## 1. General
- Always Use Option Explicit
- Two Spaces before Functions including any comments
- Indentation set to four spaces i.e. default
- Initialize or Dim variables immediately before use or just above code blocks
- Avoid use of base 1
- Avoid use 'magic numbers' in your programs, but declare them with a descriptive variable or const name 
- Explicitly indicate subs/functions as public or private (make only public if it needs to be accessed outside of module level code)
- Use For each type loops where possible which eliminate the need for an index and make it more efficient
- Avoid use strings in select case statements, but use Enums which makes things more efficient
- Use pass by reference for larger arrays which is more efficient
- Use User Defined Types (UDTs) to help organize related variables when it makes sense

## 2. Module Documentation Example
| Item               |Notes                                                                             |
|--------------------|----------------------------------------------------------------------------------|
|` Filename         `| `sampleFilename.vba                                                             `|
|` EntryPoint       `| `Indicate the main Sub or Function that is the Entry Point into the Program     `|
|` Purpose          `| `compute estimate work hours for various                                        `|
|` Inputs           `| `varies                                                                         `|
|` Outputs          `| `number of work hours                                                           `|
|` Dependencies     `| `Indicate if any libraries are used or none                                     `|
|` By/Name/Date     `| `T.Sciple, MM/DD/YYYY                                                           `|

## 3. Naming of Subs, Functions, Variables, Constants, Class Items
| Description                                           | CaseName               | Example              |
|-------------------------------------------------------|------------------------|----------------------|
| `Sub Naming, Use Verb-Noun Structure                 `| `PascalCase           `|`GenerateReport      `|
| `Function Naming, Use Verb-Noun Structure            `| `PascalCase           `|`ExportDataToCSV     `|
| `Function Parameter Names                            `| `camelCase            `|`productPrice        `|
| `Local Variable Names                                `| `snake_case           `|`current_count       `|
| `Arrays                                              `| `pluralcamelCase      `|`orderNumbers        `|
| `Constants                                           `| `UPPER_SNAKE_CASE     `|`ACCELERATION_GRAVITY`|
| `Labels                                              `| `prefixedPascalCase   `|`Lbl_ReportNoScore   `|
| `Error Handlers                                      `| `hyphenated Pascal    `|`Err_CalcTotals      `|
| `Class Naming, Use a Noun for Name                   `| `PascalCase           `|`DataExporter        `|
| `Class Methods                                       `| `PascalCase           `|`CurrentTime         `|
| `Class Private Members / Member Variables            `| `m_camelCase          `|`m_internalData      `|
| `Class Properties                                    `| `PascalCase           `|`TotalAmount         `|
| `Enums -> Name Them with camelCase                   `| `camelCase            `|`costType            `|
| `Enum Elements -> Prefix Them with Short Name        `| `prefixCamelCase      `|`ctDirectCost        `|
| `For event handlers, follow VBA's convention         `| `hyphenated PascalCase`|`Button_Click        `|
 	
## 4. Clear Listing of Sub and Function Parameters
| Description                                           | Example                                       |
|-------------------------------------------------------|-----------------------------------------------|
| `Function Parameters: Use Line Break Per Parameter   `| `ByVal user_input As Variant, _              `|
| `Parameter Type: Always specify the data type        `| `ByVal prev_score As Variant                 `|
| `Pass by Value or Reference: Explicitly declare      `| `ByVal prev_score As Variant                 `|
| `Function Return Type: Always declare return type    `| `As Variant (for returning  value)           `|

 - Example:
```vba
Function FindColumnByLabel(ByVal label As String, _  
                           ByVal searchRow As Long, _  
                           ByVal shtName As String) As Long
    ' Function body
End Function