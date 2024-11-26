# Sciple's VBA Style Guide
- 

## Working with Arrays, Ranges, Efficiently Outputting Arrays back to Worksheet Ranges
- read_output_array_to_sheet_w_format.vba
- read_sort_unique_remove_elem_output_array_1d.vba

## 1. Always Use Option Explicit
2. Two Spaces before Functions
3. Indentation set to four spaces i.e. default
3. Variable Naming
	+------------------------------------------------------+---------------------------------------+
	| Description                                          | Example                               |
	+------------------------------------------------------+---------------------------------------+
	| Enums -> Name them with CamelCase 'scoreType'        | scoreType                             |
	| Enum Elements -> Prefix them with a shortened name   | stPrevScoreEmptyExceptFrameOne        |
    | Local variable names                                 | snake_case                            |
	| Constants                                            | ALL_CAP_SNAKE_CASE                    |
	| Labels                                               | PascalCaseReportNoScoreLabel          |
	+------------------------------------------------------+---------------------------------------+