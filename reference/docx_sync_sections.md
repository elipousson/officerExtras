# Copy footer and page settings from the last section to every section

`docx_sync_sections()` copies the header and footer references, page
margins, and page size from the last section of a Word document to the
other sections, filling in what each section is missing.

This fixes documents rendered by Quarto or Pandoc where section breaks
don't carry the reference document's page settings, for example when
using Quarto's `.landscape` divs (which add section breaks with no
footer, no margins, and an A4 landscape page size) or a custom OpenXML
template. See the "Page settings and sections" section of
[`brand_reference_docx()`](https://elipousson.github.io/officerExtras/reference/brand_reference_docx.md).

For each section other than the last:

- If the section has no header or footer references, the last section's
  references are added.

- If the section has no page margins, the last section's margins are
  added.

- If the section has no page size, the last section's page size is
  added. If `page_size = TRUE`, landscape sections also get the last
  section's page size with the width and height swapped (and portrait
  sections get the last section's page size).

To use `docx_sync_sections()` in a Quarto project, add a `post-render`
script to `_quarto.yml`:

    project:
      post-render: post-render.R

With a `post-render.R` file:

    output_files <- strsplit(Sys.getenv("QUARTO_PROJECT_OUTPUT_FILES"), "\n")[[1]]

    for (file in output_files[grepl("\\.docx$", output_files)]) {
      officerExtras::docx_sync_sections(file)
    }

## Usage

``` r
docx_sync_sections(x, path = NULL, page_size = TRUE)
```

## Arguments

- x:

  A rdocx object or a path to a docx file.

- path:

  Output file path. If `NULL` (default) and `x` is a file path, `x` is
  overwritten. If `NULL` and `x` is a rdocx object, the document is not
  written to a file.

- page_size:

  If `TRUE` (default), set the page size of every section to the last
  section's page size (with the width and height swapped for landscape
  sections). If `FALSE`, only add a page size to sections without one.

## Value

A rdocx object (invisibly if a file is written).

## Examples

``` r
docx <- officer::read_docx()
docx <- officer::body_add_par(docx, "Portrait")
docx <- officer::body_end_section_portrait(docx)
docx <- officer::body_add_par(docx, "Landscape")
docx <- officer::body_end_section_landscape(docx)

docx_sync_sections(docx)
#> rdocx document with 5 element(s)
#> 
#> * styles:
#>                 Normal              heading 1              heading 2 
#>            "paragraph"            "paragraph"            "paragraph" 
#>              heading 3 Default Paragraph Font           Normal Table 
#>            "paragraph"            "character"                "table" 
#>                No List                 strong               centered 
#>            "numbering"            "character"            "paragraph" 
#>         table_template    Light List Accent 2            Titre 1 Car 
#>                "table"                "table"            "character" 
#>            Titre 2 Car            Titre 3 Car          Image Caption 
#>            "character"            "character"            "paragraph" 
#>          Table Caption     Table Professional                  toc 1 
#>            "paragraph"                "table"            "paragraph" 
#>                  toc 2           Balloon Text    Texte de bulles Car 
#>            "paragraph"            "paragraph"            "character" 
#>           reference_id          graphic title            table title 
#>            "character"            "paragraph"            "paragraph" 
```
