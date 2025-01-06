# officerExtras (development version)

## New features for existing functions

* Update `add_to_body()` to support `gt_tbl` table input to the value parameter.
* Update `read_officer()` to return empty document (based on included `styles_template.docx` file) if `allow_null = TRUE` and filename, path, and x all remain `NULL`.
* Update `officer_summary()` to return a tibble data frame (and add `{tibble}` to Imports).
* Update `cursor_docx()` to use a default cursor position function if keyword, id, and index are all `NULL`.
* Add `strict` argument to `officer_summary_levels()` to warn if `levels` contains invalid values.

## New functions

* Add new `vec_add_to_body()` function with optional `.sep` parameter.
* Add new `make_block_list()`, `combine_blocks()`, `officer_add_blocks()`, and `add_blocks_to_body()` functions.
* Add new `add_list_to_body()` function.
* Add new `combine_docx()` function.
* Add new `use_doc_version()` function.
* Export `is_officer()` and `officer_properties()` helper functions.

# officerExtras 0.0.1

* Add new `officer_summary_levels()` and `read_docs_url()` functions.
* Add new features to `officer_tables()` function using the `col`, `stack`, `type_convert`, and `nm` parameters partly inspired by [blog post by Matt Dray](https://www.rostrum.blog/2023/06/07/rectangular-officer/).

# officerExtras 0.0.0.9004

* Update `officer_media()` to support `rdocx` and `rpptx` objects and fix `overwrite` check so the function errors if files already exist.
* Add `{ggplot2}` to Suggests (used in a test for `add_gg_to_body()`).
* Add new `convert_docx()` function to support pandoc conversion for `rdocx` objects and docx files (requires `{rmarkdown}` in Suggests).
* Update `add_gt_to_body()` to incorporate implementation from `{gto}` (with better support to table captions) with credit to [Ellis Hughes](https://github.com/thebioengineer). Update required adding `{xml2}` to Imports.

# officerExtras 0.0.0.9003

* Add support for writing custom lastModifiedBy property to `write_officer()`.
* Add new functionality to `officer_summary()` to support layout and slide summaries.
* Export `check_officer_summary()` helper function.
* Add new `dims_docx_ext()` function wrapping `officer::docx_dim()`
* Add `add_gg_to_body()` function wrapping `officer::body_add_gg()`

# officerExtras 0.0.0.9002

* Add support for `gt::as_word()` parameters to `add_gt_to_body()`.
* Add support for "dotx" files to `read_officer()`
* Fix error when `filename = NULL` for `read_officer()`
* Add `officer_properties()`, `officer_tables()`, and `officer_media()` functions.

# officerExtras 0.0.0.9001

* The main change in this update is that most functions now support the full range of officer class objects to work with pptx and xlsx as well as docx files. `read_docx_ext()` is now a wrapper for a more general `read_officer()` function. `write_docx()` is replaced by `write_officer()`
* This update also includes two new functions that are vectorized versions of `add_to_body()`: `add_value_with_keys()` and `add_str_with_keys()`.
* `add_to_body()` has been updated with a few fixes and improvements making keyword, id, + index parameters optional and using `officer::body_add_par()` by default not `officer::body_add()` (due to some odd error messages when using the latter and lack of support for the pos parameter).
* Finally, test coverage is currently around 64% which is seems like a good start for a new package.

# officerExtras 0.0.0.9000

Initial set up for package with functions including:

* `read_docx_ext()`
* `write_docx()`
* `add_to_body()`, `add_text_to_body()`, `add_xml_to_body()`, and `add_gt_to_body()`
* `check_docx()` and `check_docx_fileext()`
