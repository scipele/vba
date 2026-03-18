# Plan: Piping BOM Categorizer in MS Access VBA

## Project Goal
Create a Microsoft Access VBA solution to categorize a piping Bill of Material (BOM) into commodity codes.

## Commodity Code Format
Each BOM line will be categorized into a structured commodity code with the following format:
```
AJ|XX|X|0001|AB|AB|11|00|A|X
|  |  |  |   |  |  |  |  | |
|  |  |  |   |  |  |  |  | +---- end_typ_2
|  |  |  |   |  |  |  |  +---- end_typ_1
|  |  |  |   |  |  |  +---- sch_2
|  |  |  |   |  |  +---- sch_1
|  |  |  |   |  +---- matl_grade
|  |  |  |   +--- matl_code
|  |  |  +---- indx_code
|  |  +---- rtg
|  +---- size_2
+---- size_1
```
### Data Tables (prefix: `d_`)
- [ ] d_bom_raw — Imported BOM lines
- [ ] d_bom_parsed — Parsed and categorized BOM data

### Lookup and parsing Tables (prefix: `lkp_`,`p_`)
- [ ] parse_def_com_code
- [ ] parse_def_size
- [ ] parse_def_rtg
- [x] parse_def_indx_code
- [x] parse_def_matl_code_grade
- [ ] parse_def_sch
- [ ] parse_def_end_typ

##  VBA Module Plan
- [x] **mParseSizes**: Parse and convert size fields
- [ ] **mCategBom**: Assign commodity codes using lookups
- [ ] **mUtil**: Helper functions (timing, logging, etc.)

##  Next Steps
- [ ] Review and finalize Access table design
- [ ] Draft VBA modules for each processing step
- [ ] Test with sample BOM data
- [ ] Document usage for coworkers

## References
- VBA style guide: `vba/vba_style_guide.md`