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
- [ ] lkp_com_code          |  [ ] p_com_code
- [ ] lkp_size              |  [ ] p_size
- [ ] lkp_rtg               |  [ ] p_rtg
- [ ] lkp_indx_code         |  [ ] p_indx_code
- [ ] lkp_matl_code         |  [ ] p_matl_code
- [ ] lkp_matl_grade        |  [ ] p_matl_grade
- [ ] lkp_sch               |  [ ] p_sch
- [ ] lkp_end_typ           |  [ ] p_end_typ

##  VBA Module Plan
- [ ] **mParseSizes**: Parse and convert size fields
- [ ] **mCategBom**: Assign commodity codes using lookups
- [ ] **mUtil**: Helper functions (timing, logging, etc.)

##  Next Steps
- [ ] Review and finalize Access table design
- [ ] Draft VBA modules for each processing step
- [ ] Test with sample BOM data
- [ ] Document usage for coworkers

## References
- VBA style guide: `vba/vba_style_guide.md`