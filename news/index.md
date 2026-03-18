# Changelog

## officerExtras (development version)

### New features for existing functions

- Update
  [`add_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)
  to support `gt_tbl` table input to the value parameter.
- Update
  [`read_officer()`](https://elipousson.github.io/officerExtras/reference/read_officer.md)
  to return empty document (based on included `styles_template.docx`
  file) if `allow_null = TRUE` and filename, path, and x all remain
  `NULL`.
- Update
  [`officer_summary()`](https://elipousson.github.io/officerExtras/reference/officer_summary.md)
  to return a tibble data frame (and add
  [tibble](https://tibble.tidyverse.org/) to Imports).
- Update
  [`cursor_docx()`](https://elipousson.github.io/officerExtras/reference/cursor_docx.md)
  to use a default cursor position function if keyword, id, and index
  are all `NULL`.
- Add `strict` argument to
  [`officer_summary_levels()`](https://elipousson.github.io/officerExtras/reference/officer_summary_levels.md)
  to warn if `levels` contains invalid values.

### New functions

- Add new
  [`vec_add_to_body()`](https://elipousson.github.io/officerExtras/reference/vec_add_to_body.md)
  function with optional `.sep` parameter.
- Add new
  [`make_block_list()`](https://elipousson.github.io/officerExtras/reference/make_block_list.md),
  [`combine_blocks()`](https://elipousson.github.io/officerExtras/reference/make_block_list.md),
  [`officer_add_blocks()`](https://elipousson.github.io/officerExtras/reference/officer_add_blocks.md),
  and
  [`add_blocks_to_body()`](https://elipousson.github.io/officerExtras/reference/officer_add_blocks.md)
  functions.
- Add new
  [`add_list_to_body()`](https://elipousson.github.io/officerExtras/reference/add_list_to_body.md)
  function.
- Add new
  [`combine_docx()`](https://elipousson.github.io/officerExtras/reference/combine_docx.md)
  function.
- Add new
  [`use_doc_version()`](https://elipousson.github.io/officerExtras/reference/use_doc_version.md)
  function.
- Export
  [`is_officer()`](https://elipousson.github.io/officerExtras/reference/is_officer.md)
  and
  [`officer_properties()`](https://elipousson.github.io/officerExtras/reference/officer_properties.md)
  helper functions.
- Add new
  [`officer_open()`](https://elipousson.github.io/officerExtras/reference/officer_open.md)
  function.

### Other changes

- Update included `styles_template.docx` file to use Aptos as base font.

## officerExtras 0.0.1

- Add new
  [`officer_summary_levels()`](https://elipousson.github.io/officerExtras/reference/officer_summary_levels.md)
  and
  [`read_docs_url()`](https://elipousson.github.io/officerExtras/reference/read_docs_url.md)
  functions.
- Add new features to
  [`officer_tables()`](https://elipousson.github.io/officerExtras/reference/officer_tables.md)
  function using the `col`, `stack`, `type_convert`, and `nm` parameters
  partly inspired by [blog post by Matt
  Dray](https://www.rostrum.blog/2023/06/07/rectangular-officer/).

## officerExtras 0.0.0.9004

- Update
  [`officer_media()`](https://elipousson.github.io/officerExtras/reference/officer_media.md)
  to support `rdocx` and `rpptx` objects and fix `overwrite` check so
  the function errors if files already exist.
- Add [ggplot2](https://ggplot2.tidyverse.org) to Suggests (used in a
  test for
  [`add_gg_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)).
- Add new
  [`convert_docx()`](https://elipousson.github.io/officerExtras/reference/convert_docx.md)
  function to support pandoc conversion for `rdocx` objects and docx
  files (requires [rmarkdown](https://github.com/rstudio/rmarkdown) in
  Suggests).
- Update
  [`add_gt_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)
  to incorporate implementation from `{gto}` (with better support to
  table captions) with credit to [Ellis
  Hughes](https://github.com/thebioengineer). Update required adding
  [xml2](https://xml2.r-lib.org) to Imports.

## officerExtras 0.0.0.9003

- Add support for writing custom lastModifiedBy property to
  [`write_officer()`](https://elipousson.github.io/officerExtras/reference/write_officer.md).
- Add new functionality to
  [`officer_summary()`](https://elipousson.github.io/officerExtras/reference/officer_summary.md)
  to support layout and slide summaries.
- Export
  [`check_officer_summary()`](https://elipousson.github.io/officerExtras/reference/check_officer_summary.md)
  helper function.
- Add new
  [`dims_docx_ext()`](https://elipousson.github.io/officerExtras/reference/dims_docx_ext.md)
  function wrapping
  [`officer::docx_dim()`](https://davidgohel.github.io/officer/reference/docx_dim.html)
- Add
  [`add_gg_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)
  function wrapping
  [`officer::body_add_gg()`](https://davidgohel.github.io/officer/reference/body_add_gg.html)

## officerExtras 0.0.0.9002

- Add support for
  [`gt::as_word()`](https://gt.rstudio.com/reference/as_word.html)
  parameters to
  [`add_gt_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md).
- Add support for “dotx” files to
  [`read_officer()`](https://elipousson.github.io/officerExtras/reference/read_officer.md)
- Fix error when `filename = NULL` for
  [`read_officer()`](https://elipousson.github.io/officerExtras/reference/read_officer.md)
- Add
  [`officer_properties()`](https://elipousson.github.io/officerExtras/reference/officer_properties.md),
  [`officer_tables()`](https://elipousson.github.io/officerExtras/reference/officer_tables.md),
  and
  [`officer_media()`](https://elipousson.github.io/officerExtras/reference/officer_media.md)
  functions.

## officerExtras 0.0.0.9001

- The main change in this update is that most functions now support the
  full range of officer class objects to work with pptx and xlsx as well
  as docx files.
  [`read_docx_ext()`](https://elipousson.github.io/officerExtras/reference/read_officer.md)
  is now a wrapper for a more general
  [`read_officer()`](https://elipousson.github.io/officerExtras/reference/read_officer.md)
  function. `write_docx()` is replaced by
  [`write_officer()`](https://elipousson.github.io/officerExtras/reference/write_officer.md)
- This update also includes two new functions that are vectorized
  versions of
  [`add_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md):
  [`add_value_with_keys()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)
  and
  [`add_str_with_keys()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md).
- [`add_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)
  has been updated with a few fixes and improvements making keyword,
  id, + index parameters optional and using
  [`officer::body_add_par()`](https://davidgohel.github.io/officer/reference/body_add_par.html)
  by default not
  [`officer::body_add()`](https://davidgohel.github.io/officer/reference/body_add.html)
  (due to some odd error messages when using the latter and lack of
  support for the pos parameter).
- Finally, test coverage is currently around 64% which is seems like a
  good start for a new package.

## officerExtras 0.0.0.9000

Initial set up for package with functions including:

- [`read_docx_ext()`](https://elipousson.github.io/officerExtras/reference/read_officer.md)
- `write_docx()`
- [`add_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md),
  [`add_text_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md),
  [`add_xml_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md),
  and
  [`add_gt_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)
- [`check_docx()`](https://elipousson.github.io/officerExtras/reference/check_officer.md)
  and
  [`check_docx_fileext()`](https://elipousson.github.io/officerExtras/reference/check_office_fileext.md)
