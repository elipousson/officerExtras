#' Create a Word reference document from a brand.yml file
#'
#' @description
#' `brand_reference_docx()` creates a Word document for use as a Quarto or
#' Pandoc reference document (the `reference-doc` option for docx output) using
#' the colors and typography from a
#' [brand.yml](https://posit-dev.github.io/brand-yml/) file.
#'
#' By default, the function starts from Pandoc's default `reference.docx` (so
#' every style Pandoc uses exists) and merges the brand settings into the
#' existing style definitions. Styles are edited with xml2, rather than with
#' [officer::docx_set_paragraph_style()], because that function replaces a
#' style wholesale and drops properties the brand doesn't set (e.g. the heading
#' outline levels Word uses for the table of contents).
#'
#' `pandoc_reference_docx()` writes Pandoc's default `reference.docx` to a file
#' using Quarto's bundled Pandoc (if Quarto is installed) or Pandoc.
#'
#' @section Brand settings:
#'
#' The following brand.yml settings are used:
#'
#' - `color`: `palette` and color roles (e.g. `primary` or `foreground`).
#'   Wherever a color is used, it can be a hex value, a CSS color name, an
#'   `rgb()` value, a palette name, or a color role.
#' - `typography.base`: `family`, `size`, `weight`, and `line-height` for the
#'   document defaults and the Normal style.
#' - `typography.headings`: `family`, `weight`, `style`, `color`, and
#'   `line-height` for the Heading 1-9, TOC Heading, and Title styles.
#' - `typography.link`: `color`, `background-color`, `weight`, and
#'   `decoration` for the Hyperlink style.
#' - `typography.monospace-inline` (inheriting from `typography.monospace`):
#'   `family`, `size`, `weight`, `color`, and `background-color` for the
#'   Verbatim Char style.
#' - `typography.monospace-block` (inheriting from `typography.monospace`):
#'   `background-color` (as paragraph shading), `line-height`, `color`,
#'   `family`, `size`, and `weight` for the Source Code style used for code
#'   blocks. Pandoc's default reference document has no Source Code style, so
#'   it is created if `monospace-block` sets `background-color`,
#'   `line-height`, or `color`; otherwise, Pandoc creates the style when
#'   rendering (with shading from the highlight style).
#'
#' Pandoc formats the text in code blocks with the Verbatim Char style (or
#' syntax highlighting styles based on it), which takes precedence over the
#' Source Code style. As a result, `monospace-block` values for `family`,
#' `size`, `weight`, and `color` only apply if `monospace-inline` doesn't set
#' a different value, and highlighted code keeps the token colors and shading
#' from the highlight style (set with the Quarto `highlight-style` option).
#'
#' brand.yml has no Word-specific options, so additional settings are read
#' from `defaults.docx` (the brand.yml section for format-specific options):
#'
#' - `defaults.docx.styles`: a list of style settings named by style ID (e.g.
#'   `Heading1` or `BodyText`) or style name (e.g. `Body Text`). Each style
#'   supports the typography fields `family`, `size`, `weight`, `style`,
#'   `color`, `background-color`, `decoration`, and `line-height` plus the
#'   paragraph fields `space-before`, `space-after`, `indent-left`, and
#'   `keep-with-next`. The `Table` style supports `border-color`,
#'   `border-width`, `cell-padding` (with `top`, `bottom`, `left`, and `right`),
#'   and `header` (typography fields for the first row).
#' - `defaults.docx.page`: `size` ("letter", "legal", or "a4"), `margins` (with
#'   `top`, `bottom`, `left`, `right`, `header`, `footer`, and `gutter`), and
#'   `footer` (with `text`, `align`, `size`, `color`, and `family`). In the
#'   footer text, `{page}` and `{pages}` are replaced with the page number and
#'   page count. Margins that aren't set keep the reference document's value
#'   (or Word's default of 1in for the page and 0.5in for the header and
#'   footer).
#'
#' Field names can use hyphens (as in the brand.yml spec) or underscores (as in
#' objects from [brand.yml::read_brand_yml()]).
#'
#' Sizes can use pt, in, cm, mm, px, rem, or em units (bare numbers are treated
#' as pt). Following CSS, `1rem` and `1em` are 16px (12pt). A unitless
#' `line-height` is a multiple of the font size; a `line-height` with units sets
#' an exact line height. Font weights can be numbers (100-900) or CSS weight
#' names (e.g. "semi-bold"); a weight of 600 or more is bold.
#'
#' These brand.yml settings are not used: `color.background` (Word documents
#' don't have a page color by default), `typography.fonts` (fonts are not
#' embedded, so the font families
#' must be installed wherever the document is opened), and `decoration:
#' overline` (Word has no overline).
#'
#' See `system.file("extdata", "_brand.yml", package = "officerExtras")` for an
#' example.
#'
#' Table borders and other details that the brand doesn't cover can be edited
#' by hand in Word afterwards (but re-creating the reference document
#' overwrites those edits).
#'
#' @section Page settings and sections:
#'
#' The `defaults.docx.page` settings (page size, margins, and footer) are
#' stored in the reference document's section properties. Pandoc copies them
#' to the last section of the rendered document, so they don't always apply to
#' every page:
#'
#' - Quarto's `.landscape` divs add section breaks with no footer, no margins,
#'   and an A4 landscape page size, so the pages before and in a landscape
#'   section lose the footer and page settings.
#' - A custom Pandoc OpenXML template (the `template` option for docx output,
#'   e.g. as used by the better-word Quarto extension) replaces the section
#'   properties if it doesn't use the `$sectpr$` variable.
#'
#' Use [docx_sync_sections()] on the rendered document (e.g. in a Quarto
#' `post-render` script) to copy the footer and page settings from the last
#' section to every other section. The reference document only controls styles
#' and these defaults, not the document structure (e.g. the title block or
#' table of contents), which a custom template controls.
#'
#' @param brand Path to a brand.yml file or a directory containing one, a
#'   `brand_yml` object from [brand.yml::read_brand_yml()], or a list with the
#'   brand settings (e.g. from [yaml12::read_yaml()]). If `NULL` (default), the
#'   path is taken from the `BRAND_YML_PATH` environment variable or, if that
#'   is not set, found by searching the working directory and its parent
#'   directories for `_brand.yml` or `_brand.yaml`, either directly or in a
#'   `brand` or `_brand` subdirectory (the same search used by Quarto and
#'   [brand.yml::read_brand_yml()]).
#' @param path Optional file path for the output docx file. If `NULL`
#'   (default), the document is not written to a file.
#' @param reference Optional path to a docx file or a rdocx object to use as the
#'   starting reference document. If `NULL` (default), Pandoc's default
#'   `reference.docx` is used (which requires Quarto or Pandoc). Styles named
#'   in `defaults.docx.styles` must exist in the reference document. Note that
#'   a rdocx object is modified in place (as with other officer functions).
#' @inheritParams write_officer
#' @param quiet If `TRUE`, suppress the message reporting the path of the
#'   written file. Defaults to `FALSE`.
#' @returns `brand_reference_docx()` returns a rdocx object (invisibly if
#'   `path` is supplied). `pandoc_reference_docx()` returns `path`.
#' @examples
#' if (nzchar(Sys.which("quarto")) || nzchar(Sys.which("pandoc"))) {
#'   brand_path <- system.file("extdata", "_brand.yml", package = "officerExtras")
#'
#'   withr::with_tempdir({
#'     brand_reference_docx(brand_path, path = "reference.docx")
#'   })
#' }
#' @export
brand_reference_docx <- function(
  brand = NULL,
  path = NULL,
  reference = NULL,
  overwrite = TRUE,
  quiet = FALSE
) {
  brand <- read_brand(brand)

  if (is.null(reference)) {
    reference <- pandoc_reference_docx()
  }

  if (is_officer(reference, "rdocx")) {
    doc <- reference
  } else {
    check_docx_fileext(reference)
    doc <- officer::read_docx(reference)
  }

  # Allow a reference document previously written by officer (e.g. by
  # brand_reference_docx()) to be printed again
  remove_officer_ref_marks(officer::docx_body_xml(doc))

  doc <- set_brand_styles(doc, brand)
  doc <- set_brand_page(doc, brand)

  if (is.null(path)) {
    return(doc)
  }

  write_officer(doc, path = path, overwrite = overwrite)

  if (!isTRUE(quiet)) {
    cli::cli_alert_success("Wrote {.file {path}}")
  }

  invisible(doc)
}

#' @rdname brand_reference_docx
#' @export
pandoc_reference_docx <- function(path = tempfile(fileext = ".docx")) {
  check_docx_fileext(path)

  quarto <- Sys.which("quarto")

  if (nzchar(quarto)) {
    cmd <- quarto
    args <- "pandoc"
  } else {
    cmd <- Sys.which("pandoc")

    if (!nzchar(cmd) && is_installed("rmarkdown") && rmarkdown::pandoc_available()) {
      cmd <- rmarkdown::pandoc_exec()
    }

    if (!nzchar(cmd)) {
      cli_abort("Quarto or Pandoc must be installed.")
    }

    args <- NULL
  }

  status <- system2(
    cmd,
    c(
      args,
      "-o",
      shQuote(path),
      "--print-default-data-file",
      "reference.docx"
    )
  )

  if (status != 0 || !file.exists(path)) {
    cli_abort("Can't write Pandoc's default {.file reference.docx}.")
  }

  path
}

# Read and normalize brand ---------------------------------------------------

#' Read a brand.yml file or validate a brand list
#'
#' Field names are converted to snake case (matching the brand.yml package), so
#' brand lists and `brand_yml` objects can be handled the same way. Palette
#' names and style keys are left as is.
#' @noRd
read_brand <- function(brand = NULL, call = caller_env()) {
  if (is.null(brand) || is_string(brand)) {
    brand <- find_brand_yml(brand, call = call)
    check_installed("yaml12", call = call)
    brand <- yaml12::read_yaml(brand)
  }

  if (!is.list(brand)) {
    cli_abort(
      "{.arg brand} must be a file path or a list, not {.obj_type_friendly {brand}}.",
      call = call
    )
  }

  # brand.yml allows a string in place of a typography element to set the
  # font family
  typography <- lapply(
    snake_names(brand[["typography"]]),
    function(x) if (is.character(x)) list(family = x) else x
  )

  # monospace-inline and monospace-block inherit from monospace
  for (element in c("monospace_inline", "monospace_block")) {
    typography[[element]] <- utils::modifyList(
      typography[["monospace"]] %||% list(),
      typography[[element]] %||% list()
    )
  }

  brand[["typography"]] <- typography

  docx <- brand[["defaults"]][["docx"]]

  if (!is.null(docx)) {
    docx[["styles"]] <- lapply(docx[["styles"]], snake_names)
    docx[["page"]] <- snake_names(docx[["page"]])
    brand[["defaults"]][["docx"]] <- docx
  }

  brand
}

#' Find a brand.yml file
#'
#' Follows the search used by Quarto and `brand.yml::read_brand_yml()`: use
#' `path` (or the `BRAND_YML_PATH` environment variable) if it is a file;
#' otherwise search the directory (default: working directory) and up to 20
#' parent directories for `_brand.yml` or `_brand.yaml`, directly or in a
#' `brand` or `_brand` subdirectory. A `path` to a directory is only searched
#' (not its parents).
#' @noRd
find_brand_yml <- function(path = NULL, call = caller_env()) {
  if (is.null(path)) {
    env_path <- Sys.getenv("BRAND_YML_PATH")
    path <- if (nzchar(env_path)) env_path
  }

  max_parents <- if (is.null(path)) 20 else 1
  dir <- path %||% getwd()

  if (!dir.exists(dir)) {
    if (file.exists(dir)) {
      return(dir)
    }

    cli_abort("{.arg brand} file {.file {dir}} does not exist.", call = call)
  }

  dir <- normalizePath(dir, mustWork = TRUE)
  filenames <- c("_brand.yml", "_brand.yaml")
  candidates <- c(
    filenames,
    file.path("brand", filenames),
    file.path("_brand", filenames)
  )

  for (i in seq_len(max_parents)) {
    found <- file.path(dir, candidates)
    found <- found[file.exists(found) & !dir.exists(found)]

    if (length(found) > 0) {
      return(found[[1]])
    }

    if (dirname(dir) == dir) {
      break
    }

    dir <- dirname(dir)
  }

  cli_abort(
    "Can't find {.file _brand.yml} or {.file _brand.yaml} in
    {.path {path %||% getwd()}}{if (max_parents > 1) ' or its parent directories'}.",
    call = call
  )
}

#' Convert list names to snake case at every level
#' @noRd
snake_names <- function(x) {
  if (!is.list(x)) {
    return(x)
  }

  if (!is.null(names(x))) {
    names(x) <- snake_case(names(x))
  }

  lapply(x, snake_names)
}

snake_case <- function(x) gsub("-", "_", x, fixed = TRUE)

# Unit, weight, and color helpers --------------------------------------------

#' Convert a size (e.g. "11pt", "0.25in", "1rem", or a bare number in pt) to
#' points
#' @noRd
to_pt <- function(x, call = caller_env()) {
  if (is.null(x)) {
    return(NULL)
  }

  if (is.numeric(x) && length(x) == 1) {
    return(x)
  }

  scale <- c(
    pt = 1,
    `in` = 72,
    cm = 72 / 2.54,
    mm = 72 / 25.4,
    px = 0.75,
    rem = 12,
    em = 12
  )

  size <- parse_size(x)

  if (is.null(size) || !size[["unit"]] %in% c("", names(scale))) {
    cli_abort(
      "Can't convert size {.val {x}} to points.
      Use a size in pt, in, cm, mm, px, rem, or em units.",
      call = call
    )
  }

  if (size[["unit"]] == "") {
    return(size[["value"]])
  }

  size[["value"]] * scale[[size[["unit"]]]]
}

#' Split a length 1 size into a numeric value and a (possibly empty) unit, or
#' return `NULL` if `x` is not a size
#' @noRd
parse_size <- function(x) {
  if (is.numeric(x) && length(x) == 1) {
    return(list(value = x, unit = ""))
  }

  if (!is_string(x)) {
    return(NULL)
  }

  x <- tolower(x)
  match <- regmatches(x, regexec("^\\s*([0-9.]+)\\s*([a-z%]*)\\s*$", x))[[1]]
  value <- suppressWarnings(as.numeric(match[2]))

  if (length(match) == 0 || is.na(value)) {
    return(NULL)
  }

  list(value = value, unit = match[3])
}

# Word uses twentieths of a point (twips) for spacing and indents, and
# half-points for font sizes
to_twips <- function(x, call = caller_env()) {
  as.character(round(to_pt(x, call = call) * 20))
}

to_half_pt <- function(x, call = caller_env()) {
  as.character(round(to_pt(x, call = call) * 2))
}

#' Convert a line-height to w:spacing attributes. A unitless value (or a
#' percentage) is a multiple of the font size; a value with units is an exact
#' height.
#' @noRd
line_spacing <- function(x, call = caller_env()) {
  size <- parse_size(x)

  if (is.null(size)) {
    cli_abort("Can't use {.val {x}} as a line height.", call = call)
  }

  if (size[["unit"]] == "") {
    multiple <- size[["value"]]
  } else if (size[["unit"]] == "%") {
    multiple <- size[["value"]] / 100
  } else {
    return(c(line = to_twips(x, call = call), lineRule = "exact"))
  }

  # 240 = single line spacing
  c(line = as.character(round(multiple * 240)), lineRule = "auto")
}

# CSS font weight names
font_weights <- c(
  thin = 100, hairline = 100, extralight = 200, ultralight = 200,
  light = 300, lighter = 300, normal = 400, regular = 400, book = 400,
  medium = 500, semibold = 600, demibold = 600, bold = 700, bolder = 700,
  extrabold = 800, ultrabold = 800, black = 900, heavy = 900
)

#' Convert a font weight (a number or CSS weight name) to a number
#' @noRd
font_weight <- function(x, call = caller_env()) {
  weight <- suppressWarnings(as.numeric(x))

  if (is.na(weight)) {
    weight <- font_weights[gsub("[-_ ]", "", tolower(x))]
  }

  if (length(weight) != 1 || is.na(weight)) {
    cli_abort(
      "Can't use {.val {x}} as a font weight.
      Use a number from 100 to 900 or a weight name like {.val semi-bold}.",
      call = call
    )
  }

  unname(weight)
}

is_bold_weight <- function(x, call = caller_env()) {
  font_weight(x, call = call) >= 600
}

#' Resolve a color (hex value, CSS color name, `rgb()` value, palette name, or
#' color role) to a hex value without "#"
#'
#' Names are matched ignoring hyphens and underscores, because the brand.yml
#' package converts palette names to snake case.
#' @noRd
resolve_color <- function(x, colors, call = caller_env()) {
  if (is.null(x)) {
    return(NULL)
  }

  palette <- colors[["palette"]] %||% list()
  roles <- colors[setdiff(names(colors), "palette")] %||% list()
  names(palette) <- snake_case(names(palette) %||% character(0))
  names(roles) <- snake_case(names(roles) %||% character(0))

  value <- x
  seen <- character(0)

  # Roles and palette names can reference palette names or other roles
  repeat {
    key <- snake_case(value)

    if (key %in% seen) {
      cli_abort(
        "Can't resolve color {.val {x}}: {.val {c(seen, key)}} is a circular reference.",
        call = call
      )
    }

    seen <- c(seen, key)

    if (key %in% names(roles)) {
      value <- roles[[key]]
    } else if (key %in% names(palette)) {
      value <- palette[[key]]
    } else {
      break
    }
  }

  hex <- color_to_hex(value)

  if (is.null(hex)) {
    cli_abort("Can't resolve color {.val {x}} to a hex value.", call = call)
  }

  hex
}

#' Convert a hex value (with or without "#"), CSS color name, or `rgb()` value
#' to a 6 digit hex value without "#", or return `NULL`
#' @noRd
color_to_hex <- function(x) {
  if (!is_string(x)) {
    return(NULL)
  }

  x <- trimws(x)

  if (grepl("^rgba?\\(", x)) {
    channels <- suppressWarnings(as.numeric(
      strsplit(gsub("^rgba?\\(|\\)$", "", x), "[ ,/]+")[[1]]
    ))

    if (length(channels) < 3 || anyNA(channels[1:3])) {
      return(NULL)
    }

    rgb <- matrix(channels[1:3])
  } else {
    # Hex values without "#" are allowed (Word's own format), but other bare
    # numbers are not: col2rgb() treats them as palette indexes
    if (grepl("^([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$", x)) {
      x <- paste0("#", x)
    } else if (grepl("^[0-9]+$", x)) {
      return(NULL)
    }

    rgb <- tryCatch(grDevices::col2rgb(x), error = function(e) NULL)
  }

  if (is.null(rgb)) {
    return(NULL)
  }

  sprintf("%02X%02X%02X", rgb[1], rgb[2], rgb[3])
}

# XML helpers ----------------------------------------------------------------

# OOXML requires child elements in schema order. These vectors list the order
# for the elements set by brand_reference_docx() (from the ECMA-376 schema).
ooxml_child_order <- list(
  style = c(
    "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden",
    "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked",
    "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr",
    "tblPr", "trPr", "tcPr", "tblStylePr"
  ),
  pPr = c(
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
    "pPrChange"
  ),
  rPr = c(
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
    "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
    "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
    "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath"
  ),
  tblPr = c(
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
    "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"
  ),
  tblStylePr = c("pPr", "rPr", "tblPr", "trPr", "tcPr"),
  sectPr = c(
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
    "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols",
    "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi",
    "rtlGutter", "docGrid", "printerSettings", "sectPrChange"
  ),
  tblBorders = c("top", "left", "bottom", "right", "insideH", "insideV"),
  tblCellMar = c("top", "left", "bottom", "right")
)

# Word's default page margins in twips. CT_PageMar requires every attribute.
ooxml_page_margins <- c(
  top = "1440", right = "1440", bottom = "1440", left = "1440",
  header = "720", footer = "720", gutter = "0"
)

w_ns <- function(x) c(w = xml2::xml_ns(x)[["w"]])

#' Find or create the child `w:{name}` of `parent`, inserted in schema order,
#' and replace its attributes with `attrs` (if supplied)
#' @noRd
set_ooxml_child <- function(parent, name, attrs = NULL) {
  ns <- w_ns(parent)
  child <- xml2::xml_find_first(parent, paste0("w:", name), ns)

  if (inherits(child, "xml_missing")) {
    order <- ooxml_child_order[[xml2::xml_name(parent)]]
    siblings <- xml2::xml_name(xml2::xml_children(parent))
    position <- sum(
      match(siblings, order, nomatch = 0) < match(name, order) &
        siblings %in% order
    )
    xml2::xml_add_child(parent, paste0("w:", name), .where = position)
    child <- xml2::xml_find_first(parent, paste0("w:", name), ns)
  }

  if (!is.null(attrs)) {
    # Remove existing attributes first: theme attributes (e.g. w:themeColor,
    # w:asciiTheme) take precedence over explicit values in Word
    for (attr in names(xml2::xml_attrs(child, ns = ns))) {
      xml2::xml_set_attr(child, attr, NULL, ns = ns)
    }

    set_ooxml_attrs(child, attrs)
  }

  invisible(child)
}

#' Set `w:` attributes on a node, keeping any other existing attributes
#' @noRd
set_ooxml_attrs <- function(node, attrs) {
  ns <- w_ns(node)

  for (attr in names(attrs)) {
    xml2::xml_set_attr(node, paste0("w:", attr), attrs[[attr]], ns = ns)
  }

  invisible(node)
}

#' Abort if a list has names other than `allowed`
#' @noRd
check_brand_names <- function(x, allowed, arg, call = caller_env()) {
  unknown <- setdiff(names(x), allowed)

  if (length(unknown) > 0) {
    cli_abort(
      c(
        "{.field {arg}} has unsupported field{?s} {.val {unknown}}.",
        "i" = "Supported fields: {.val {allowed}}."
      ),
      call = call
    )
  }

  invisible(x)
}

#' Apply typography fields (family, size, weight, style, color,
#' background_color, decoration) to a w:rPr node
#' @noRd
set_run_props <- function(rpr, props, colors, call = caller_env()) {
  if (!is.null(props[["family"]])) {
    family <- props[["family"]]
    # Explicit fonts replace the theme font attributes (which take precedence)
    set_ooxml_child(
      rpr,
      "rFonts",
      c(ascii = family, hAnsi = family, eastAsia = family, cs = family)
    )
  }

  if (!is.null(props[["weight"]])) {
    bold <- if (is_bold_weight(props[["weight"]], call = call)) "1" else "0"
    set_ooxml_child(rpr, "b", c(val = bold))
    set_ooxml_child(rpr, "bCs", c(val = bold))
  }

  if (!is.null(props[["style"]])) {
    italic <- if (props[["style"]] %in% c("italic", "oblique")) "1" else "0"
    set_ooxml_child(rpr, "i", c(val = italic))
    set_ooxml_child(rpr, "iCs", c(val = italic))
  }

  if (!is.null(props[["color"]])) {
    color <- resolve_color(props[["color"]], colors, call = call)
    set_ooxml_child(rpr, "color", c(val = color))
  }

  if (!is.null(props[["size"]])) {
    set_ooxml_child(rpr, "sz", c(val = to_half_pt(props[["size"]], call = call)))
    set_ooxml_child(rpr, "szCs", c(val = to_half_pt(props[["size"]], call = call)))
  }

  if (!is.null(props[["decoration"]])) {
    set_decoration(rpr, props[["decoration"]], call = call)
  }

  if (!is.null(props[["background_color"]])) {
    fill <- resolve_color(props[["background_color"]], colors, call = call)
    set_ooxml_child(rpr, "shd", c(val = "clear", color = "auto", fill = fill))
  }

  invisible(rpr)
}

#' Apply a CSS text-decoration value (e.g. "underline", "line-through",
#' "none", or "underline line-through") to a w:rPr node
#' @noRd
set_decoration <- function(rpr, decoration, call = caller_env()) {
  values <- strsplit(trimws(decoration), "\\s+")[[1]]
  unknown <- setdiff(values, c("none", "underline", "line-through", "overline"))

  if (length(unknown) > 0) {
    cli_abort("Can't use {.val {decoration}} as a text decoration.", call = call)
  }

  if ("overline" %in% values) {
    cli::cli_warn("Word has no overline, so {.val overline} is ignored.")
  }

  set_ooxml_child(rpr, "strike", c(val = if ("line-through" %in% values) "1" else "0"))
  set_ooxml_child(rpr, "u", c(val = if ("underline" %in% values) "single" else "none"))

  invisible(rpr)
}

#' Apply paragraph fields (space_before, space_after, line_height, indent_left,
#' keep_with_next) to a w:pPr node
#' @noRd
set_par_props <- function(ppr, props, call = caller_env()) {
  if (isTRUE(props[["keep_with_next"]])) {
    set_ooxml_child(ppr, "keepNext")
  }

  spacing <- c(
    before = if (!is.null(props[["space_before"]])) {
      to_twips(props[["space_before"]], call = call)
    },
    after = if (!is.null(props[["space_after"]])) {
      to_twips(props[["space_after"]], call = call)
    },
    if (!is.null(props[["line_height"]])) {
      line_spacing(props[["line_height"]], call = call)
    }
  )

  if (length(spacing) > 0) {
    # Keep existing spacing attributes the brand doesn't set
    set_ooxml_attrs(set_ooxml_child(ppr, "spacing"), spacing)
  }

  if (!is.null(props[["indent_left"]])) {
    set_ooxml_child(ppr, "ind", c(left = to_twips(props[["indent_left"]], call = call)))
  }

  invisible(ppr)
}

#' Find a style by ID or, failing that, by name (ignoring case, since Word
#' stores built-in style names in lower case, e.g. "heading 1"). If `required =
#' FALSE`, return `NULL` for a missing style.
#' @noRd
find_style <- function(styles, style, required = TRUE, call = caller_env()) {
  ns <- w_ns(styles)
  nodes <- xml2::xml_find_all(styles, "//w:style", ns)
  ids <- xml2::xml_attr(nodes, "styleId")
  style_names <- xml2::xml_attr(xml2::xml_find_first(nodes, "w:name", ns), "val")

  index <- match(style, ids)

  if (is.na(index)) {
    index <- match(tolower(style), tolower(style_names))
  }

  if (!is.na(index)) {
    return(nodes[[index]])
  }

  if (required) {
    cli_abort(
      "Style {.val {style}} is not in the reference document.",
      call = call
    )
  }

  NULL
}

# Build ----------------------------------------------------------------------

#' Apply brand typography and `defaults.docx.styles` to the styles.xml part of
#' a rdocx object
#' @noRd
set_brand_styles <- function(doc, brand, call = caller_env()) {
  styles_path <- file.path(doc[["package_dir"]], "word", "styles.xml")
  styles <- xml2::read_xml(styles_path)
  ns <- w_ns(styles)

  colors <- brand[["color"]]
  typography <- brand[["typography"]]
  docx_styles <- brand[["defaults"]][["docx"]][["styles"]]
  par_fields <- c(
    "space_before", "space_after", "line_height", "indent_left", "keep_with_next"
  )

  # pPr must come before rPr in a style, so it is created first (and only if
  # needed, since character styles don't have paragraph properties)
  set_style_props <- function(style, props) {
    if (is.null(style)) {
      return(invisible(NULL))
    }

    if (any(par_fields %in% names(compact(props)))) {
      set_par_props(set_ooxml_child(style, "pPr"), props, call = call)
    }

    set_run_props(set_ooxml_child(style, "rPr"), props, colors, call = call)
  }

  base <- typography[["base"]]

  # Document defaults (base typography)
  rpr_default <- xml2::xml_find_first(styles, "//w:docDefaults/w:rPrDefault/w:rPr", ns)
  ppr_default <- xml2::xml_find_first(styles, "//w:docDefaults/w:pPrDefault/w:pPr", ns)

  if (!inherits(rpr_default, "xml_missing")) {
    set_run_props(rpr_default, base[c("family", "size", "weight")], colors, call = call)
  }

  if (!inherits(ppr_default, "xml_missing")) {
    set_par_props(ppr_default, base["line_height"], call = call)
  }

  # Normal: base font and foreground color (also set here because some
  # renderers, e.g. macOS Quick Look, ignore the document defaults)
  set_style_props(
    find_style(styles, "Normal", required = FALSE),
    c(
      base[c("family", "size", "weight")],
      list(color = if (!is.null(colors[["foreground"]])) "foreground")
    )
  )

  # Headings 1-9, TOC Heading, and Title (typography.headings); in Pandoc's
  # reference document, Subtitle, Author, and Date are based on Title
  headings <- typography[["headings"]][
    c("family", "weight", "style", "color", "line_height")
  ]

  for (style_id in c(paste0("Heading", 1:9), "TOCHeading", "Title")) {
    set_style_props(find_style(styles, style_id, required = FALSE), headings)
  }

  # Links and code (typography.link, typography.monospace_inline)
  set_style_props(
    find_style(styles, "Hyperlink", required = FALSE),
    typography[["link"]][c("color", "background_color", "weight", "decoration")]
  )

  set_style_props(
    find_style(styles, "VerbatimChar", required = FALSE),
    typography[["monospace_inline"]][
      c("family", "size", "weight", "color", "background_color")
    ]
  )

  set_source_code_style(
    styles,
    typography[["monospace_block"]],
    typography[["monospace_inline"]],
    colors,
    call = call
  )

  # Per-style settings (defaults.docx.styles)
  for (style in setdiff(names(docx_styles), "Table")) {
    set_style_props(find_style(styles, style, call = call), docx_styles[[style]])
  }

  if (!is.null(docx_styles[["Table"]])) {
    set_table_style(
      find_style(styles, "Table", call = call),
      docx_styles[["Table"]],
      colors,
      call = call
    )
  }

  xml2::write_xml(styles, styles_path)

  doc
}

#' Apply typography.monospace_block to the Source Code paragraph style used
#' for code blocks, creating the style if needed
#' @noRd
set_source_code_style <- function(
  styles,
  block,
  inline,
  colors,
  call = caller_env()
) {
  block <- compact(block[c("family", "size", "weight", "color", "background_color", "line_height")])
  style <- find_style(styles, "SourceCode", required = FALSE)

  if (is.null(style)) {
    # Without these settings, the style Pandoc creates when rendering is a
    # better match (it has shading from the highlight style)
    if (!any(c("background_color", "line_height", "color") %in% names(block))) {
      return(invisible(NULL))
    }

    style <- add_source_code_style(styles)
  }

  run_fields <- c("family", "size", "weight", "color")
  overridden <- run_fields[vapply(
    run_fields,
    function(field) {
      !is.null(block[[field]]) && !is.null(inline[[field]]) &&
        !identical(block[[field]], inline[[field]])
    },
    logical(1)
  )]

  if (length(overridden) > 0) {
    cli::cli_warn(c(
      "{.field typography.monospace-block} {.field {overridden}} differ{?s/} from
      {.field typography.monospace-inline}.",
      "i" = "Pandoc formats code block text with the Verbatim Char style, so the
      {.field monospace-inline} {cli::qty(overridden)}value{?s} {?is/are} used."
    ))
  }

  ppr <- set_ooxml_child(style, "pPr")

  if (!is.null(block[["background_color"]])) {
    fill <- resolve_color(block[["background_color"]], colors, call = call)
    set_ooxml_child(ppr, "shd", c(val = "clear", color = "auto", fill = fill))
  }

  set_par_props(ppr, block["line_height"], call = call)

  run_props <- block[intersect(run_fields, names(block))]

  if (length(run_props) > 0) {
    set_run_props(set_ooxml_child(style, "rPr"), run_props, colors, call = call)
  }

  invisible(style)
}

#' Add a Source Code paragraph style matching the one Pandoc creates
#' @noRd
add_source_code_style <- function(styles) {
  ns <- w_ns(styles)
  style <- xml2::xml_add_child(xml2::xml_root(styles), "w:style")
  set_ooxml_attrs(style, c(type = "paragraph", customStyle = "1", styleId = "SourceCode"))
  set_ooxml_child(style, "name", c(val = "Source Code"))
  set_ooxml_child(style, "basedOn", c(val = "Normal"))
  set_ooxml_child(style, "link", c(val = "VerbatimChar"))
  set_ooxml_child(set_ooxml_child(style, "pPr"), "wordWrap", c(val = "off"))

  xml2::xml_find_first(styles, "//w:style[@w:styleId='SourceCode']", ns)
}

#' Apply table settings to a table style node
#' @noRd
set_table_style <- function(table_style, props, colors, call = caller_env()) {
  tbl_pr <- set_ooxml_child(table_style, "tblPr")

  if (!is.null(props[["border_width"]]) && is.null(props[["border_color"]])) {
    cli::cli_warn(
      "{.field border-width} is ignored without {.field border-color} in
      {.field defaults.docx.styles.Table}."
    )
  }

  if (!is.null(props[["border_color"]])) {
    borders <- set_ooxml_child(tbl_pr, "tblBorders")
    # Border widths are in eighths of a point
    border_size <- to_pt(props[["border_width"]] %||% 0.5, call = call) * 8
    border_attrs <- c(
      val = "single",
      sz = as.character(round(border_size)),
      space = "0",
      color = resolve_color(props[["border_color"]], colors, call = call)
    )

    for (side in ooxml_child_order[["tblBorders"]]) {
      set_ooxml_child(borders, side, border_attrs)
    }
  }

  cell_padding <- props[["cell_padding"]]

  if (!is.null(cell_padding)) {
    check_brand_names(
      cell_padding,
      ooxml_child_order[["tblCellMar"]],
      arg = "defaults.docx.styles.Table.cell-padding",
      call = call
    )

    cell_mar <- set_ooxml_child(tbl_pr, "tblCellMar")

    for (side in names(cell_padding)) {
      set_ooxml_child(
        cell_mar,
        side,
        c(w = to_twips(cell_padding[[side]], call = call), type = "dxa")
      )
    }
  }

  if (!is.null(props[["header"]])) {
    ns <- w_ns(table_style)
    first_row <- xml2::xml_find_first(
      table_style,
      "w:tblStylePr[@w:type='firstRow']",
      ns
    )

    # Pandoc marks header rows with w:tblLook w:firstRow="1", so the first row
    # formatting applies once the table style has it. tblStylePr is the last
    # element in a style, so it can be appended.
    if (inherits(first_row, "xml_missing")) {
      xml2::xml_add_child(table_style, "w:tblStylePr")
      first_row <- xml2::xml_find_first(
        table_style,
        "w:tblStylePr[not(@w:type)]",
        ns
      )
      set_ooxml_attrs(first_row, c(type = "firstRow"))
    }

    set_run_props(set_ooxml_child(first_row, "rPr"), props[["header"]], colors, call = call)
  }

  invisible(table_style)
}

#' Apply `defaults.docx.page` (footer, page size, and margins) to the default
#' section of a rdocx object
#' @noRd
set_brand_page <- function(doc, brand, call = caller_env()) {
  page <- brand[["defaults"]][["docx"]][["page"]]

  if (is.null(page)) {
    return(doc)
  }

  colors <- brand[["color"]]
  sect_path <- "//w:body/w:sectPr"

  if (!is.null(page[["footer"]])) {
    doc <- set_brand_footer(doc, page[["footer"]], colors, call = call)
  }

  body <- officer::docx_body_xml(doc)
  ns <- w_ns(body)
  sect_pr <- xml2::xml_find_first(body, sect_path, ns)

  page_size <- switch(page[["size"]] %||% "letter",
    letter = c(w = "12240", h = "15840"),
    legal = c(w = "12240", h = "20160"),
    a4 = c(w = "11906", h = "16838"),
    cli_abort(
      "Unsupported page size {.val {page[['size']]}}.
      Use \"letter\", \"legal\", or \"a4\".",
      call = call
    )
  )

  set_ooxml_child(sect_pr, "pgSz", page_size)

  if (!is.null(page[["margins"]])) {
    set_page_margins(sect_pr, page[["margins"]], call = call)
  }

  # officer writes pgMar before pgSz
  sort_sect_pr(sect_pr)

  doc
}

#' Put section properties in schema order. Header and footer references
#' (which come first) stay in place: officer fills in their content when the
#' document is written.
#' @noRd
sort_sect_pr <- function(sect_pr) {
  children <- xml2::xml_find_all(
    sect_pr,
    "*[not(self::w:headerReference or self::w:footerReference)]",
    w_ns(sect_pr)
  )
  sorted <- children[
    order(match(xml2::xml_name(children), ooxml_child_order[["sectPr"]]))
  ]

  # xml_add_child() appends a copy, so add the sorted copies, then remove the
  # originals
  for (child in sorted) {
    xml2::xml_add_child(sect_pr, child)
  }

  xml2::xml_remove(children)

  invisible(sect_pr)
}

#' Set w:pgMar, filling margins the brand doesn't set from the existing
#' margins or Word's defaults (every attribute is required)
#' @noRd
set_page_margins <- function(sect_pr, margins, call = caller_env()) {
  check_brand_names(
    margins,
    names(ooxml_page_margins),
    arg = "defaults.docx.page.margins",
    call = call
  )

  pg_mar <- xml2::xml_find_first(sect_pr, "w:pgMar", w_ns(sect_pr))
  attrs <- ooxml_page_margins

  if (!inherits(pg_mar, "xml_missing")) {
    existing <- xml2::xml_attrs(pg_mar)
    existing <- existing[names(existing) %in% names(attrs)]
    attrs[names(existing)] <- existing
  }

  attrs[names(margins)] <- vapply(margins, to_twips, character(1), call = call)

  set_ooxml_child(sect_pr, "pgMar", attrs)
}

#' Set the default footer with optional `{page}` and `{pages}` fields
#'
#' [officer::body_set_default_section()] writes the footer part (and its
#' relationship and content type) but replaces the section properties, so the
#' existing footnote settings are copied over.
#'
#' officer leaves a non-standard `w:officer="true"` attribute on the
#' footerReference, which Pandoc copies into rendered documents. Word ignores it
#' (and drops it on save) and opens the file without a repair prompt.
#' @noRd
set_brand_footer <- function(doc, footer, colors, call = caller_env()) {
  check_string(
    footer[["text"]],
    arg = "defaults.docx.page.footer.text",
    call = call
  )

  sect_path <- "//w:body/w:sectPr"
  body <- officer::docx_body_xml(doc)
  ns <- w_ns(body)

  footnote_pr <- xml2::xml_find_first(body, paste0(sect_path, "/w:footnotePr"), ns)

  footer_text <- officer::fp_text_lite(
    font.size = to_pt(footer[["size"]], call = call) %||% NA,
    color = if (!is.null(footer[["color"]])) {
      paste0("#", resolve_color(footer[["color"]], colors, call = call))
    } else {
      NA
    },
    font.family = footer[["family"]] %||% NA
  )

  # Split the footer text into text runs and page number fields, e.g.
  # "Page {page} of {pages}"
  fields <- c(page = "PAGE", pages = "NUMPAGES")
  parts <- regmatches(
    footer[["text"]],
    gregexpr("\\{[a-z]+\\}|[^{]+|\\{", footer[["text"]])
  )[[1]]

  runs <- lapply(parts, function(part) {
    field <- fields[gsub("[{}]", "", part)]

    if (grepl("^\\{[a-z]+\\}$", part) && !is.na(field)) {
      officer::run_word_field(field, prop = footer_text)
    } else {
      officer::ftext(part, prop = footer_text)
    }
  })

  footer_par <- do.call(
    officer::fpar,
    c(runs, list(fp_p = officer::fp_par(text.align = footer[["align"]] %||% "center")))
  )

  doc <- officer::body_set_default_section(
    doc,
    officer::prop_section(footer_default = officer::block_list(footer_par))
  )

  sect_pr <- xml2::xml_find_first(officer::docx_body_xml(doc), sect_path, ns)

  # officer adds a continuous section break type, which isn't needed for the
  # only section
  xml2::xml_remove(xml2::xml_find_all(sect_pr, "w:type", ns))

  # footnotePr goes after the header and footer references
  if (!inherits(footnote_pr, "xml_missing")) {
    n_refs <- length(xml2::xml_find_all(
      sect_pr,
      "w:headerReference|w:footerReference",
      ns
    ))
    xml2::xml_add_child(sect_pr, footnote_pr, .where = n_refs)
  }

  doc
}
