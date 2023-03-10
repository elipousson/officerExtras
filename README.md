
<!-- README.md is generated from README.Rmd. Please edit that file -->

# officerExtras

<!-- badges: start -->

[![Lifecycle:
experimental](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://lifecycle.r-lib.org/articles/stages.html#experimental)
[![License:
MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Project Status: WIP – Initial development is in progress, but there
has not yet been a stable, usable release suitable for the
public.](https://www.repostatus.org/badges/latest/wip.svg)](https://www.repostatus.org/#wip)
[![Codecov test
coverage](https://codecov.io/gh/elipousson/officerExtras/branch/main/graph/badge.svg)](https://app.codecov.io/gh/elipousson/officerExtras?branch=main)
<!-- badges: end -->

The goal of officerExtras is to provide helper and convenience functions
for the [{officer}](https://github.com/davidgohel/officer) package.

## Installation

You can install the development version of officerExtras like so:

``` r
pak::pkg_install("elipousson/officerExtras")
```

## Usage

The officerExtras package provides a variety of helper functions to
simplify the process of working with officer. For example, a single
`read_officer()` function works with docx, pptx, or xlsx files:

``` r
library(officerExtras)

docx <- read_officer(filename = "example.docx", path = system.file("doc_examples", package = "officer"))

pptx <- read_officer(filename = "example.pptx", path = system.file("doc_examples", package = "officer"))
```

officer uses a print method to save rdocx, rpptx, or rxlsx objects back
to files. officerExtras provides a `write_officer()` with that adds the
option to pass a filename and path separately and pass document
properties to `officer::doc_properties()`.

``` r
withr::with_tempdir({
  write_officer(docx, "write-example.docx", modified_by = "officerExtras", title = "Document Title set by doc_properties", subject = "Microsoft Word, R")
  
  example_docx <- read_officer("write-example.docx")
  
  officer::doc_properties(example_docx)
})
#>               tag                                value
#> 1           title Document Title set by doc_properties
#> 2         subject                    Microsoft Word, R
#> 3         creator                               author
#> 4        keywords                                     
#> 5     description                   these are comments
#> 6  lastModifiedBy                        officerExtras
#> 7        revision                                   11
#> 8         created                 2017-04-26T13:10:00Z
#> 9        modified                 2023-03-13T09:52:30Z
#> 10       category
```

## Related projects

### Extending officer

- [gto](https://github.com/GSK-Biostatistics/gto): allow users to insert
  gt tables into officeverse.
- [onbrand](https://github.com/john-harrold/onbrand): A R package to
  provide a systematic method to script support for different Word of
  PowerPoint templates.
- [officerWinTools](https://github.com/joshmire/officerWinTools): A R
  package to complement the officer package when using Microsoft Office
  in a Windows environment.

### Working with Microsoft Office documents with R

- [docxtractr](https://github.com/hrbrmstr/docxtractr): Extract Tables
  from Microsoft Word Documents with R
- [stylex](https://github.com/niszet/stylex): A R package to help update
  docx style files.
- [deef](https://github.com/prcleary/deef): Data extractor scripts for
  electronic forms, compatible with Microsoft Word files.

### Other tools for working with Microsoft Office documents

- [python-openxml/python-docx](https://github.com/python-openxml/python-docx):Create
  and modify Word documents with Python
