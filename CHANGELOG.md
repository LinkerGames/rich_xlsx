# Changelog

## 1.1.0 (unreleased)

- Added cell comments (notes): `Worksheet#add_comment(ref, text, author: nil)` and the `:comment_author` workbook option.
- Added `Worksheet.cell_to_xy`, the reverse of `Worksheet.xy_to_cell`.
- Worksheets now always declare the `r` (relationships) namespace on the root element.

## 1.0.0 (2026-06-03)

- Initial release
