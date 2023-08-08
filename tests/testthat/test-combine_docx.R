test_that("combine_docx works", {
  docx1 <- read_officer()
  docx2 <- read_officer()

  docx_combined <- combine_docx(docx1, docx2)

  expect_true(
    is_rdocx(docx_combined)
  )

  expect_equal(
    nrow(officer_summary(docx_combined)),
    nrow(officer_summary(docx1)) * 2
  )

  docx3_path <- system.file("doc_examples", "example.docx", package = "officer")

  # expect_equal(
  #   nrow(officer_summary(combine_docx(docx3_path, docx = docx1))),
  #   nrow(officer_summary(read_officer(docx3_path))) + nrow(officer_summary(docx1))
  # )
})
