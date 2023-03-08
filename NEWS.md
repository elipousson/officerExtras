<!-- NEWS.md is maintained by https://cynkra.github.io/fledge, do not edit -->

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
