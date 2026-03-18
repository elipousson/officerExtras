# officerExtras

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
[`read_officer()`](https://elipousson.github.io/officerExtras/reference/read_officer.md)
function works with docx, pptx, or xlsx files:

``` r
library(officerExtras)

docx <- read_officer(filename = "example.docx", path = system.file("doc_examples", package = "officer"))

pptx <- read_officer(filename = "example.pptx", path = system.file("doc_examples", package = "officer"))
```

officer uses a print method to save rdocx, rpptx, or rxlsx objects back
to files. officerExtras provides a
[`write_officer()`](https://elipousson.github.io/officerExtras/reference/write_officer.md)
with that adds the option to pass a filename and path separately and
pass document properties to
[`officer::doc_properties()`](https://davidgohel.github.io/officer/reference/doc_properties.html).

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
#> 7        revision                                   12
#> 8         created                 2017-04-26T13:10:00Z
#> 9        modified                 2023-06-22T11:04:39Z
#> 10       category
```

The package also wraps useful functions from a few other packages.
[`convert_docx()`](https://elipousson.github.io/officerExtras/reference/convert_docx.md)
uses
[`rmarkdown::pandoc_convert()`](https://pkgs.rstudio.com/rmarkdown/reference/pandoc_convert.html)
to convert a rdocx object or a Word document to any other pandoc
supported output format:

``` r
convert_docx(docx, to = "markdown")
```

## Related projects

There are a number of related projects for extending officer or working
with [OOXML](https://en.wikipedia.org/wiki/Office_Open_XML) in R and
other languages.

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

- [python-docx](https://github.com/python-openxml/python-docx):Create
  and modify Word documents with Python
- [Docx-templates](https://github.com/guigrpa/docx-templates):
  Template-based docx report creation for both Node and the browser
- [docx](https://github.com/dolanmiu/docx): Easily generate and modify
  .docx files with JS/TS. Works for Node and on the Browser.
- [docxtemplater](https://github.com/open-xml-templating/docxtemplater):
  Generate docx, pptx, and xlsx from templates (Word, Powerpoint and
  Excel documents), from Node.js, the Browser and the command line
