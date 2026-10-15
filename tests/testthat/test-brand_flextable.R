test_that("set_flextable_brand_defaults sets defaults from a brand", {
  skip_if_not_installed("flextable")

  old <- flextable::get_flextable_defaults()
  withr::defer(do.call(flextable::set_flextable_defaults, old))

  brand <- list(
    color = list(
      palette = list(gray = "#595959", `gray-light` = "#BFBFBF"),
      foreground = "gray"
    ),
    typography = list(base = list(family = "Aptos", size = "11pt", `line-height` = 1.15)),
    defaults = list(docx = list(styles = list(Table = list(
      `border-color` = "gray-light",
      `border-width` = "0.5pt",
      `cell-padding` = list(top = "2pt", left = "5pt"),
      header = list(weight = "normal", color = "gray", `background-color` = "#EEEEEE")
    ))))
  )

  prev <- set_flextable_brand_defaults(brand, digits = 3)
  expect_identical(prev$font.family, old$font.family)

  defaults <- flextable::get_flextable_defaults()
  expect_identical(defaults$font.family, "Aptos")
  expect_identical(defaults$hansi.family, "Aptos")
  expect_identical(defaults$font.size, 11)
  expect_identical(defaults$font.color, "#595959")
  expect_identical(defaults$line_spacing, 1.15)
  expect_identical(defaults$border.color, "#BFBFBF")
  expect_identical(defaults$border.width, 0.5)
  expect_identical(defaults$padding.top, 2)
  expect_identical(defaults$padding.left, 5)
  # Unset padding keeps the flextable default
  expect_identical(defaults$padding.bottom, old$padding.bottom)
  # Arguments in ... take precedence
  expect_identical(defaults$digits, 3)

  ft <- flextable::flextable(head(mtcars, 2))
  header <- flextable::information_data_chunk(ft)
  header <- header[header$.part == "header", ]
  expect_true(all(!header$bold))
  expect_true(all(header$color == "#595959"))
  cells <- flextable::information_data_cell(ft)
  expect_true(all(cells$background.color[cells$.part == "header"] == "#EEEEEE"))
})

test_that("set_flextable_brand_defaults validates brand values up front", {
  skip_if_not_installed("flextable")

  old <- flextable::get_flextable_defaults()
  withr::defer(do.call(flextable::set_flextable_defaults, old))

  table_brand <- function(table) {
    list(defaults = list(docx = list(styles = list(Table = table))))
  }

  expect_error(
    set_flextable_brand_defaults(table_brand(list(`border-color` = "notacolor"))),
    "Can't resolve color"
  )
  expect_error(
    set_flextable_brand_defaults(table_brand(list(header = list(weight = "heavyish")))),
    "font weight"
  )
  expect_error(
    set_flextable_brand_defaults(table_brand(list(`cell-padding` = list(inner = "1pt")))),
    "unsupported field"
  )

  # A line-height with units is skipped (flextable only supports multiples)
  set_flextable_brand_defaults(list(typography = list(base = list(`line-height` = "18pt"))))
  expect_identical(flextable::get_flextable_defaults()$line_spacing, old$line_spacing)
})

test_that("set_flextable_brand_defaults works with the example brand file", {
  skip_if_not_installed("flextable")
  skip_if_not_installed("yaml12")

  old <- flextable::get_flextable_defaults()
  withr::defer(do.call(flextable::set_flextable_defaults, old))

  brand_path <- system.file("extdata", "_brand.yml", package = "officerExtras")
  set_flextable_brand_defaults(brand_path)

  ft <- flextable::flextable(head(mtcars, 2))
  header <- flextable::information_data_chunk(ft)
  expect_true(all(header$bold[header$.part == "header"]))
  expect_identical(flextable::get_flextable_defaults()$border.color, "#BFBFBF")
})
