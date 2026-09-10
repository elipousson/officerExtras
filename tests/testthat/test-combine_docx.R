test_that("combine_docx works", {
  docx_path <- system.file("doc_examples", "example.docx", package = "officer")

  docx1 <- read_officer(docx_path)
  docx2 <- read_officer(docx_path)

  docx_combined <- combine_docx(docx1, docx2)

  expect_true(
    is_rdocx(docx_combined)
  )

  expect_equal(
    nrow(officer_summary(docx_combined)),
    0
  )
})

test_that("combine_docx accepts file paths and defaults docx to a new document", {
  docx_path <- system.file("doc_examples", "example.docx", package = "officer")

  docx_combined <- combine_docx(docx_path, docx_path)

  expect_true(is_rdocx(docx_combined))
})

test_that("combine_docx supports sep as a function", {
  docx_path <- system.file("doc_examples", "example.docx", package = "officer")
  docx1 <- read_officer(docx_path)
  docx2 <- read_officer(docx_path)

  docx_combined <- combine_docx(docx1, docx2, sep = officer::body_add_break)

  expect_true(is_rdocx(docx_combined))
})

test_that("combine_docx supports sep as a value passed to add_to_body", {
  docx_path <- system.file("doc_examples", "example.docx", package = "officer")
  docx1 <- read_officer(docx_path)
  docx2 <- read_officer(docx_path)

  docx_combined <- combine_docx(docx1, docx2, sep = "separator text")

  summary_df <- officer_summary(docx_combined)
  expect_true("separator text" %in% summary_df[["text"]])
})

test_that("combine_docx errors for invalid .list entries", {
  expect_error(
    combine_docx(.list = list(123))
  )
})
