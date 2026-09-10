test_that("write_officer works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer"),
    quiet = TRUE
  )

  pptx <- read_pptx_ext(
    filename = "example.pptx",
    path = system.file("doc_examples", package = "officer")
  )

  withr::with_tempdir({
    write_officer(docx, "example.docx")

    expect_error(
      write_officer(docx, "example.docx", overwrite = FALSE)
    )

    write_officer(docx, "example.docx")

    expect_identical(
      docx[["doc_obj"]],
      read_docx_ext(
        filename = "example.docx",
        quiet = TRUE
      )[["doc_obj"]]
    )

    write_officer(docx, "example.docx", modified_by = "test")

    expect_identical(
      "test",
      officer_properties(
        read_docx_ext(
          filename = "example.docx",
          quiet = TRUE
        )
      )[["lastModifiedBy"]]
    )

    write_officer(pptx, "example.pptx")

    expect_identical(
      pptx[["doc_obj"]],
      read_pptx_ext(
        filename = "example.pptx",
        quiet = TRUE
      )[["doc_obj"]]
    )
  })
})

test_that("write_officer writes rxlsx objects", {
  xlsx <- read_xlsx_ext(
    filename = "template.xlsx",
    path = system.file("template", package = "officer")
  )

  withr::with_tempdir({
    write_officer(xlsx, "out.xlsx")

    expect_true(file.exists("out.xlsx"))
  })
})

test_that("write_officer errors for non-officer objects and bad file extensions", {
  expect_error(
    write_officer("not an officer object", "out.docx")
  )

  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_error(
    write_officer(docx, "out.txt")
  )
})
