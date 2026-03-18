# Update the version of a document and optionally save a new version of the input document

**\[experimental\]**

`use_doc_version()` is inspired by `usethis::use_version()` and designed
to support semantic versioning for Microsoft Word or PowerPoint
documents. The document version is tracked as a custom file property and
a component of the document filename. If filename is supplied and the
filename contains a version number, the function ignores any version
number stored as a document property and overwrites the property with
the new incremented version number.

`doc_version()` is a helper function that returns the document version
based on a supplied filename or document properties.

The internals for this function are adapted from the internal
`idesc_bump_version()` function authored by Csárdi Gábor for the
`{desc}` package.

## Usage

``` r
use_doc_version(
  filename = NULL,
  x = NULL,
  which = NULL,
  save = TRUE,
  sep = ".",
  property = "version",
  prefix = NULL,
  path = NULL,
  ...,
  call = caller_env()
)

doc_version(
  filename = NULL,
  x = NULL,
  sep = ".",
  property = "version",
  allow_new = TRUE,
  .default = c(0, 1, 0),
  call = caller_env()
)
```

## Arguments

- filename:

  Filename for a Word document or PowerPoint presentation. Optional if x
  is supplied, however, it is required to locally save a file with an
  updated version name.

- x:

  A rdocx or rpptx object. Optional if filename is supplied.

- which:

  A string specifying which level to increment, one of: "major",
  "minor", "patch", "dev".

- save:

  If `TRUE` (default) and filename is supplied, write a new file
  replacing the existing version number.

- sep:

  Character separating version components. Defaults to ".". "-" is also
  supported.

- property:

  Property name optionally containing a version number. Defaults to
  "version".

- prefix:

  Property name to use as prefix for new filename, e.g. "modified" to
  use modified date/time as the new filename prefix. If prefix is
  identical to property and the filename does not already include a
  version number, version is added to the filename as a prefix instead
  of a postfix. Defaults to `NULL`. Note handling of this parameter is
  expected to change in future versions of this function.

- path:

  Passed to
  [`write_officer()`](https://elipousson.github.io/officerExtras/reference/write_officer.md).

- ...:

  Additional parameters passed to
  [`write_officer()`](https://elipousson.github.io/officerExtras/reference/write_officer.md)
  if save is `TRUE`.

- call:

  The execution environment of a currently running function, e.g.
  `caller_env()`. The function will be mentioned in error messages as
  the source of the error. See the `call` argument of
  [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
  information.

- allow_new:

  If `TRUE` (default), return "0.1.0" if version can't be found in the
  filename or the properties of the input rdocx or rpptx object x.

- .default:

  Specification for initial version.

## Value

Invisibly return a rdocx or rpptx object with an updated version
property.
