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
    1
  )
})
