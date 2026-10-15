# Create a Word reference document from a brand.yml file

`brand_reference_docx()` creates a Word document for use as a Quarto or
Pandoc reference document (the `reference-doc` option for docx output)
using the colors and typography from a
[brand.yml](https://posit-dev.github.io/brand-yml/) file.

By default, the function starts from Pandoc's default `reference.docx`
(so every style Pandoc uses exists) and merges the brand settings into
the existing style definitions. Styles are edited with xml2, rather than
with
[`officer::docx_set_paragraph_style()`](https://davidgohel.github.io/officer/reference/docx_set_paragraph_style.html),
because that function replaces a style wholesale and drops properties
the brand doesn't set (e.g. the heading outline levels Word uses for the
table of contents).

`pandoc_reference_docx()` writes Pandoc's default `reference.docx` to a
file using Quarto's bundled Pandoc (if Quarto is installed) or Pandoc.

## Usage

``` r
brand_reference_docx(
  brand = NULL,
  path = NULL,
  reference = NULL,
  overwrite = TRUE,
  quiet = FALSE
)

pandoc_reference_docx(path = tempfile(fileext = ".docx"))
```

## Arguments

- brand:

  Path to a brand.yml file or a directory containing one, a `brand_yml`
  object from
  [`brand.yml::read_brand_yml()`](https://rdrr.io/pkg/brand.yml/man/read_brand_yml.html),
  or a list with the brand settings (e.g. from
  [`yaml12::read_yaml()`](https://posit-dev.github.io/r-yaml12/reference/parse_yaml.html)).
  If `NULL` (default), the path is taken from the `BRAND_YML_PATH`
  environment variable or, if that is not set, found by searching the
  working directory and its parent directories for `_brand.yml` or
  `_brand.yaml`, either directly or in a `brand` or `_brand`
  subdirectory (the same search used by Quarto and
  [`brand.yml::read_brand_yml()`](https://rdrr.io/pkg/brand.yml/man/read_brand_yml.html)).

- path:

  Optional file path for the output docx file. If `NULL` (default), the
  document is not written to a file.

- reference:

  Optional path to a docx file or a rdocx object to use as the starting
  reference document. If `NULL` (default), Pandoc's default
  `reference.docx` is used (which requires Quarto or Pandoc). Styles
  named in `defaults.docx.styles` must exist in the reference document.
  Note that a rdocx object is modified in place (as with other officer
  functions).

- overwrite:

  If `TRUE` (default), remove file at path if it already exists. If
  `FALSE` and file exists, this function aborts.

- quiet:

  If `TRUE`, suppress the message reporting the path of the written
  file. Defaults to `FALSE`.

## Value

`brand_reference_docx()` returns a rdocx object (invisibly if `path` is
supplied). `pandoc_reference_docx()` returns `path`.

## Brand settings

The following brand.yml settings are used:

- `color`: `palette` and color roles (e.g. `primary` or `foreground`).
  Wherever a color is used, it can be a hex value, a CSS color name, an
  [`rgb()`](https://rdrr.io/r/grDevices/rgb.html) value, a palette name,
  or a color role.

- `typography.base`: `family`, `size`, `weight`, and `line-height` for
  the document defaults and the Normal style.

- `typography.headings`: `family`, `weight`, `style`, `color`, and
  `line-height` for the Heading 1-9, TOC Heading, and Title styles.

- `typography.link`: `color`, `background-color`, `weight`, and
  `decoration` for the Hyperlink style.

- `typography.monospace-inline` (inheriting from
  `typography.monospace`): `family`, `size`, `weight`, `color`, and
  `background-color` for the Verbatim Char style.

- `typography.monospace-block` (inheriting from `typography.monospace`):
  `background-color` (as paragraph shading), `line-height`, `color`,
  `family`, `size`, and `weight` for the Source Code style used for code
  blocks. Pandoc's default reference document has no Source Code style,
  so it is created if `monospace-block` sets `background-color`,
  `line-height`, or `color`; otherwise, Pandoc creates the style when
  rendering (with shading from the highlight style).

Pandoc formats the text in code blocks with the Verbatim Char style (or
syntax highlighting styles based on it), which takes precedence over the
Source Code style. As a result, `monospace-block` values for `family`,
`size`, `weight`, and `color` only apply if `monospace-inline` doesn't
set a different value, and highlighted code keeps the token colors and
shading from the highlight style (set with the Quarto `highlight-style`
option).

brand.yml has no Word-specific options, so additional settings are read
from `defaults.docx` (the brand.yml section for format-specific
options):

- `defaults.docx.styles`: a list of style settings named by style ID
  (e.g. `Heading1` or `BodyText`) or style name (e.g. `Body Text`). Each
  style supports the typography fields `family`, `size`, `weight`,
  `style`, `color`, `background-color`, `decoration`, and `line-height`
  plus the paragraph fields `space-before`, `space-after`,
  `indent-left`, and `keep-with-next`. The `Table` style supports
  `border-color`, `border-width`, `cell-padding` (with `top`, `bottom`,
  `left`, and `right`), and `header` (typography fields for the first
  row).

- `defaults.docx.page`: `size` ("letter", "legal", or "a4"), `margins`
  (with `top`, `bottom`, `left`, `right`, `header`, `footer`, and
  `gutter`), and `footer` (with `text`, `align`, `size`, `color`, and
  `family`). In the footer text, `{page}` and `{pages}` are replaced
  with the page number and page count. Margins that aren't set keep the
  reference document's value (or Word's default of 1in for the page and
  0.5in for the header and footer).

Field names can use hyphens (as in the brand.yml spec) or underscores
(as in objects from
[`brand.yml::read_brand_yml()`](https://rdrr.io/pkg/brand.yml/man/read_brand_yml.html)).

Sizes can use pt, in, cm, mm, px, rem, or em units (bare numbers are
treated as pt). Following CSS, `1rem` and `1em` are 16px (12pt). A
unitless `line-height` is a multiple of the font size; a `line-height`
with units sets an exact line height. Font weights can be numbers
(100-900) or CSS weight names (e.g. "semi-bold"); a weight of 600 or
more is bold.

These brand.yml settings are not used: `color.background` (Word
documents don't have a page color by default), `typography.fonts` (fonts
are not embedded, so the font families must be installed wherever the
document is opened), and `decoration: overline` (Word has no overline).

See `system.file("extdata", "_brand.yml", package = "officerExtras")`
for an example.

Table borders and other details that the brand doesn't cover can be
edited by hand in Word afterwards (but re-creating the reference
document overwrites those edits).

## Page settings and sections

The `defaults.docx.page` settings (page size, margins, and footer) are
stored in the reference document's section properties. Pandoc copies
them to the last section of the rendered document, so they don't always
apply to every page:

- Quarto's `.landscape` divs add section breaks with no footer, no
  margins, and an A4 landscape page size, so the pages before and in a
  landscape section lose the footer and page settings.

- A custom Pandoc OpenXML template (the `template` option for docx
  output, e.g. as used by the better-word Quarto extension) replaces the
  section properties if it doesn't use the `$sectpr$` variable.

Use
[`docx_sync_sections()`](https://elipousson.github.io/officerExtras/reference/docx_sync_sections.md)
on the rendered document (e.g. in a Quarto `post-render` script) to copy
the footer and page settings from the last section to every other
section. The reference document only controls styles and these defaults,
not the document structure (e.g. the title block or table of contents),
which a custom template controls.

## Examples

``` r
if (nzchar(Sys.which("quarto")) || nzchar(Sys.which("pandoc"))) {
  brand_path <- system.file("extdata", "_brand.yml", package = "officerExtras")

  withr::with_tempdir({
    brand_reference_docx(brand_path, path = "reference.docx")
  })
}
#> ✔ Wrote reference.docx
```
