# officerExtras development

* Added support for the `gt::as_word()` parameters to `add_gt_to_body()`.
* Add support for "dotx" files to `read_officer()`
* Fix error when `filename = NULL` for `read_officer()`
* Export `officer_properties()` function

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
