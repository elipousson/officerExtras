#' Set flextable defaults from a brand.yml file
#'
#' `set_flextable_brand_defaults()` sets [flextable::set_flextable_defaults()]
#' from a brand.yml file, so tables created with flextable match the `Table`
#' style in a reference document created with [brand_reference_docx()].
#' flextable formats each table directly (rather than using the reference
#' document's `Table` style), so the defaults are needed for tables to match.
#'
#' The following brand.yml settings are used:
#'
#' - `typography.base`: `family`, `size`, and `line-height` (if unitless) for
#'   the table text.
#' - `color.foreground` for the text color.
#' - `defaults.docx.styles.Table`: `border-color`, `border-width`,
#'   `cell-padding`, and `header` (`family`, `size`, `weight`, `style`,
#'   `color`, and `background-color` for the header row).
#'
#' @inheritParams brand_reference_docx
#' @param theme_fun A flextable theme function applied to each new table,
#'   before the `header` settings. Defaults to [flextable::theme_box()] (borders
#'   on every cell, as in the reference document `Table` style).
#' @param ... Additional arguments passed to
#'   [flextable::set_flextable_defaults()]. These take precedence over the
#'   brand settings.
#' @returns The previous flextable defaults, invisibly. Pass them to
#'   `do.call(flextable::set_flextable_defaults, old)` to restore them.
#' @examples
#' if (rlang::is_installed(c("flextable", "yaml12"))) {
#'   brand_path <- system.file("extdata", "_brand.yml", package = "officerExtras")
#'
#'   old <- set_flextable_brand_defaults(brand_path)
#'   flextable::flextable(head(mtcars, 3))
#'
#'   do.call(flextable::set_flextable_defaults, old)
#' }
#' @export
set_flextable_brand_defaults <- function(
  brand = NULL,
  theme_fun = NULL,
  ...
) {
  check_installed("flextable")
  brand <- read_brand(brand)
  call <- current_env()

  colors <- brand[["color"]]
  base <- brand[["typography"]][["base"]]
  table_style <- brand[["defaults"]][["docx"]][["styles"]][["Table"]]
  padding <- table_style[["cell_padding"]]

  if (!is.null(padding)) {
    check_brand_names(
      padding,
      ooxml_child_order[["tblCellMar"]],
      arg = "defaults.docx.styles.Table.cell-padding"
    )
  }

  hex_color <- function(x) {
    if (!is.null(x)) {
      paste0("#", resolve_color(x, colors, call = call))
    }
  }

  family <- base[["family"]]

  brand_defaults <- list(
    # Set every font slot: flextable's defaults for the others (e.g. hAnsi)
    # are not the base font
    font.family = family,
    hansi.family = family,
    eastasia.family = family,
    cs.family = family,
    font.size = to_pt(base[["size"]]),
    font.color = hex_color(if (!is.null(colors[["foreground"]])) "foreground"),
    line_spacing = brand_line_multiple(base[["line_height"]]),
    border.color = hex_color(table_style[["border_color"]]),
    border.width = to_pt(table_style[["border_width"]]),
    padding.top = to_pt(padding[["top"]]),
    padding.bottom = to_pt(padding[["bottom"]]),
    padding.left = to_pt(padding[["left"]]),
    padding.right = to_pt(padding[["right"]]),
    theme_fun = brand_flextable_theme(
      theme_fun %||% flextable::theme_box,
      table_style[["header"]],
      hex_color
    )
  )

  brand_defaults <- utils::modifyList(compact(brand_defaults), list2(...))

  invisible(do.call(flextable::set_flextable_defaults, brand_defaults))
}

#' Return a unitless (or percentage) line-height as a multiple, or `NULL` for a
#' line-height with units (flextable only supports multiples)
#' @noRd
brand_line_multiple <- function(x) {
  size <- parse_size(x)

  if (is.null(size)) {
    return(NULL)
  }

  if (size[["unit"]] == "") {
    return(size[["value"]])
  }

  if (size[["unit"]] == "%") {
    return(size[["value"]] / 100)
  }

  NULL
}

#' Wrap a flextable theme function to apply the brand Table header settings
#' @noRd
brand_flextable_theme <- function(theme_fun, header, hex_color) {
  theme_fun <- as_function(theme_fun)
  header <- compact(header)

  if (is_empty(header)) {
    return(theme_fun)
  }

  # Convert values up front so invalid values error when the defaults are set,
  # not when a table is created
  header_size <- to_pt(header[["size"]])
  header_bold <- if (!is.null(header[["weight"]])) is_bold_weight(header[["weight"]])
  header_italic <- if (!is.null(header[["style"]])) {
    header[["style"]] %in% c("italic", "oblique")
  }
  header_color <- hex_color(header[["color"]])
  header_bg <- hex_color(header[["background_color"]])

  function(x) {
    x <- theme_fun(x)

    if (!is.null(header[["family"]])) {
      x <- flextable::font(x, fontname = header[["family"]], part = "header")
    }

    if (!is.null(header_size)) {
      x <- flextable::fontsize(x, size = header_size, part = "header")
    }

    if (!is.null(header_bold)) {
      x <- flextable::bold(x, bold = header_bold, part = "header")
    }

    if (!is.null(header_italic)) {
      x <- flextable::italic(x, italic = header_italic, part = "header")
    }

    if (!is.null(header_color)) {
      x <- flextable::color(x, color = header_color, part = "header")
    }

    if (!is.null(header_bg)) {
      x <- flextable::bg(x, bg = header_bg, part = "header")
    }

    x
  }
}
