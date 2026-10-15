style_xml <- function(doc, xpath) {
  styles <- xml2::read_xml(file.path(doc$package_dir, "word", "styles.xml"))
  xml2::xml_find_first(styles, xpath, xml2::xml_ns(styles))
}

style_attr <- function(doc, xpath, attr = "val") {
  xml2::xml_attr(style_xml(doc, xpath), attr)
}

skip_if_no_pandoc <- function() {
  skip_if(!nzchar(Sys.which("quarto")) && !nzchar(Sys.which("pandoc")))
}

test_that("to_pt converts units", {
  expect_identical(to_pt("11pt"), 11)
  expect_identical(to_pt("1in"), 72)
  expect_identical(to_pt("16px"), 12)
  expect_identical(to_pt("1rem"), 12)
  expect_identical(to_pt("0.75em"), 9)
  expect_identical(to_pt("11PT"), 11)
  expect_identical(to_pt(10), 10)
  expect_identical(to_pt("10"), 10)
  expect_null(to_pt(NULL))
  expect_error(to_pt("1vw"), "Can't convert size")
  expect_error(to_pt("large"), "Can't convert size")
  expect_identical(to_twips("0.25in"), "360")
  expect_identical(to_half_pt("11pt"), "22")
})

test_that("line_spacing handles multiples, percentages, and fixed heights", {
  expect_identical(line_spacing(1.15), c(line = "276", lineRule = "auto"))
  expect_identical(line_spacing("1.5"), c(line = "360", lineRule = "auto"))
  expect_identical(line_spacing("150%"), c(line = "360", lineRule = "auto"))
  expect_identical(line_spacing("24px"), c(line = "360", lineRule = "exact"))
  expect_error(line_spacing("normal"), "line height")
})

test_that("resolve_color resolves roles, palette names, and CSS colors", {
  colors <- list(
    palette = list(`blue-dark` = "#00415f", red = "#f00", alias = "red"),
    primary = "blue-dark",
    accent = "primary",
    background = "white"
  )

  expect_identical(resolve_color("primary", colors), "00415F")
  expect_identical(resolve_color("accent", colors), "00415F")
  # brand.yml package objects use snake case palette names
  expect_identical(resolve_color("blue_dark", colors), "00415F")
  expect_identical(resolve_color("alias", colors), "FF0000")
  expect_identical(resolve_color("background", colors), "FFFFFF")
  expect_identical(resolve_color("rgb(0, 65, 95)", colors), "00415F")
  expect_identical(resolve_color("#abcdef", colors), "ABCDEF")
  expect_identical(resolve_color("abcdef", colors), "ABCDEF")
  expect_null(resolve_color(NULL, colors))
  expect_error(resolve_color("notacolor", colors), "Can't resolve color")
  expect_error(resolve_color("1", colors), "Can't resolve color")

  cyclic <- list(palette = list(a = "b", b = "a"))
  expect_error(resolve_color("a", cyclic), "circular reference")
})

test_that("font weights handle names and numbers", {
  expect_true(is_bold_weight("bold"))
  expect_true(is_bold_weight("semi-bold"))
  expect_true(is_bold_weight("SemiBold"))
  expect_true(is_bold_weight("extra-bold"))
  expect_true(is_bold_weight(700))
  expect_true(is_bold_weight("600"))
  expect_false(is_bold_weight("normal"))
  expect_false(is_bold_weight("medium"))
  expect_false(is_bold_weight(400))
  expect_error(is_bold_weight("heavyish"), "font weight")
})

test_that("read_brand validates input and normalizes fields", {
  expect_error(read_brand("missing_brand.yml"), "does not exist")
  expect_error(read_brand(1), "must be a file path or a list")

  brand <- read_brand(list(
    typography = list(
      base = "Aptos",
      monospace = list(family = "Consolas", size = "10pt"),
      `monospace-inline` = list(color = "red")
    ),
    defaults = list(docx = list(
      styles = list(`Body Text` = list(`space-after` = "6pt"))
    ))
  ))

  expect_identical(brand$typography$base, list(family = "Aptos"))
  expect_identical(
    brand$typography$monospace_inline,
    list(family = "Consolas", size = "10pt", color = "red")
  )
  # Style keys are kept, field names are converted
  expect_named(brand$defaults$docx$styles, "Body Text")
  expect_named(brand$defaults$docx$styles$`Body Text`, "space_after")
})

test_that("brand_reference_docx works with a non-Pandoc reference", {
  brand <- list(
    color = list(palette = list(blue = "#00415F"), primary = "blue"),
    typography = list(base = list(family = "Aptos", size = "1rem")),
    defaults = list(docx = list(page = list(size = "a4")))
  )

  doc <- brand_reference_docx(brand, reference = officer::read_docx())
  expect_s3_class(doc, "rdocx")
  expect_identical(
    style_attr(doc, "//w:style[@w:styleId='Normal']/w:rPr/w:sz"),
    "24"
  )

  # Styles named in defaults.docx.styles must exist
  brand$defaults$docx$styles <- list(Compact = list(size = "9pt"))
  expect_error(
    brand_reference_docx(brand, reference = officer::read_docx()),
    "is not in the reference document"
  )
})

test_that("brand_reference_docx applies spec fields", {
  skip_if_no_pandoc()
  reference <- pandoc_reference_docx()

  brand <- list(
    color = list(background = "white", primary = "#00415F"),
    typography = list(
      base = list(`line-height` = "18pt", weight = 300),
      headings = list(weight = "semi-bold", `line-height` = 1.2),
      monospace = list(family = "Consolas"),
      `monospace-inline` = list(`background-color` = "#EEEEEE"),
      link = list(decoration = "line-through")
    ),
    defaults = list(docx = list(styles = list(
      # Style name instead of ID
      `Body Text` = list(`space-after` = "6pt")
    )))
  )

  doc <- brand_reference_docx(brand, reference = reference)

  expect_identical(
    xml2::xml_attrs(style_xml(doc, "//w:pPrDefault//w:spacing"))[
      c("line", "lineRule")
    ],
    c(line = "360", lineRule = "exact")
  )
  expect_identical(style_attr(doc, "//w:rPrDefault//w:b"), "0")
  expect_identical(
    style_attr(doc, "//w:style[@w:styleId='Heading1']/w:rPr/w:b"),
    "1"
  )
  expect_identical(
    style_attr(doc, "//w:style[@w:styleId='Heading1']/w:pPr/w:spacing", "line"),
    "288"
  )
  expect_identical(
    style_attr(doc, "//w:style[@w:styleId='VerbatimChar']/w:rPr/w:shd", "fill"),
    "EEEEEE"
  )
  expect_identical(
    style_attr(doc, "//w:style[@w:styleId='VerbatimChar']/w:rPr/w:rFonts", "ascii"),
    "Consolas"
  )
  # Character styles don't get paragraph properties
  expect_s3_class(
    style_xml(doc, "//w:style[@w:styleId='VerbatimChar']/w:pPr"),
    "xml_missing"
  )
  expect_identical(
    style_attr(doc, "//w:style[@w:styleId='Hyperlink']/w:rPr/w:strike"),
    "1"
  )
  expect_identical(
    style_attr(doc, "//w:style[@w:styleId='BodyText']/w:pPr/w:spacing", "after"),
    "120"
  )

  expect_warning(
    brand_reference_docx(
      list(typography = list(link = list(decoration = "overline"))),
      reference = reference
    ),
    "overline"
  )
})

test_that("brand_reference_docx writes complete page margins", {
  skip_if_no_pandoc()

  doc <- brand_reference_docx(
    list(defaults = list(docx = list(page = list(margins = list(top = "0.5in"))))),
    reference = pandoc_reference_docx()
  )

  pg_mar <- xml2::xml_find_first(officer::docx_body_xml(doc), "//w:pgMar")
  attrs <- xml2::xml_attrs(pg_mar)

  expect_identical(
    attrs[c("top", "right", "bottom", "left", "header", "footer", "gutter")],
    c(
      top = "720", right = "1440", bottom = "1440", left = "1440",
      header = "720", footer = "720", gutter = "0"
    )
  )

  expect_error(
    brand_reference_docx(
      list(defaults = list(docx = list(page = list(margins = list(`margin-top` = "1in"))))),
      reference = pandoc_reference_docx()
    ),
    "unsupported field"
  )
})

test_that("brand_reference_docx checks the footer text", {
  skip_if_no_pandoc()

  expect_error(
    brand_reference_docx(
      list(defaults = list(docx = list(page = list(footer = list(align = "right"))))),
      reference = pandoc_reference_docx()
    ),
    "footer.text"
  )
})

test_that("brand_reference_docx accepts brand.yml objects", {
  skip_if_no_pandoc()
  skip_if_not_installed("brand.yml")

  brand_path <- system.file("extdata", "_brand.yml", package = "officerExtras")
  doc <- brand_reference_docx(
    brand.yml::read_brand_yml(brand_path),
    reference = pandoc_reference_docx()
  )

  expect_identical(
    style_attr(doc, "//w:style[@w:styleId='Title']/w:rPr/w:color"),
    "00415F"
  )
  expect_false(inherits(
    style_xml(doc, "//w:style[@w:styleId='DefinitionTerm']/w:pPr/w:keepNext"),
    "xml_missing"
  ))
})

test_that("brand_reference_docx works with the example brand file", {
  skip_if_no_pandoc()
  skip_if_not_installed("yaml12")

  brand_path <- system.file("extdata", "_brand.yml", package = "officerExtras")
  path <- withr::local_tempfile(fileext = ".docx")

  expect_message(
    brand_reference_docx(brand_path, path = path),
    "Wrote"
  )
  expect_true(file.exists(path))

  doc <- officer::read_docx(path)
  body <- officer::docx_body_xml(doc)
  ns <- xml2::xml_ns(body)

  expect_identical(
    xml2::xml_attr(xml2::xml_find_first(body, "//w:sectPr/w:pgSz", ns), "w"),
    "12240"
  )
  expect_length(xml2::xml_find_all(body, "//w:sectPr/w:footerReference", ns), 1)
  expect_identical(
    style_attr(doc, "//w:style[@w:styleId='Heading1']/w:rPr/w:color"),
    "00415F"
  )
})
