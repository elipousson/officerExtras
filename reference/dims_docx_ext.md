# Get page and block layout dimensions from an rdocx object

This function extends
[`officer::docx_dim()`](https://davidgohel.github.io/officer/reference/docx_dim.html)
by also returning the body text dimensions within the margins, the
aspect ratio of the page and body, and the page orientation as a string
("landscape" or "portrait").

## Usage

``` r
dims_docx_ext(docx)
```

## Arguments

- docx:

  A rdocx object to get dimensions for.

## See also

[`officer::docx_dim()`](https://davidgohel.github.io/officer/reference/docx_dim.html)
