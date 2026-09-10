test_that("read_officer works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_true(
    inherits(docx, "rdocx")
  )

  expect_identical(
    read_docx_ext(docx = docx),
    docx
  )

  expect_identical(
    officer::docx_summary(docx),
    officer::docx_summary(
      officer::read_docx(
        system.file("doc_examples", "example.docx", package = "officer")
      )
    )
  )

  expect_identical(
    read_docx_ext(allow_null = TRUE)[["styles"]],
    officer::read_docx(system.file(
      "template",
      "styles_template.docx",
      package = "officerExtras"
    ))[["styles"]]
  )

  expect_error(
    read_docx_ext()
  )

  expect_message(
    read_docx_ext(docx = docx, quiet = FALSE)
  )

  expect_message(
    read_docx_ext(filename = "test.docx", docx = docx, quiet = FALSE)
  )

  pptx <- read_pptx_ext(
    filename = "example.pptx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_true(
    inherits(pptx, "rpptx")
  )

  xlsx <- read_xlsx_ext(
    filename = "template.xlsx",
    path = system.file("template", package = "officer")
  )

  expect_true(
    inherits(xlsx, "rxlsx")
  )
})

test_that("read_officer generic entrypoint works for all file types", {
  docx <- read_officer(
    system.file("doc_examples/example.docx", package = "officer")
  )
  expect_s3_class(docx, "rdocx")

  expect_s3_class(read_officer(x = docx), "rdocx")

  expect_s3_class(read_officer(fileext = "pptx"), "rpptx")
  expect_s3_class(read_officer(fileext = "xlsx"), "rxlsx")

  expect_error(
    read_officer(x = docx, fileext = "pptx")
  )
})

test_that("officer_properties supports values and keep.null", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  props <- officer_properties(docx)
  expect_true("title" %in% names(props))

  updated <- officer_properties(docx, values = list(title = "New Title"))
  expect_identical(updated[["title"]], "New Title")

  dropped <- officer_properties(docx, values = list(title = NULL))
  expect_false("title" %in% names(dropped))

  kept <- officer_properties(
    docx,
    values = list(title = NULL),
    keep.null = TRUE
  )
  expect_true("title" %in% names(kept))
  expect_null(kept[["title"]])
})
