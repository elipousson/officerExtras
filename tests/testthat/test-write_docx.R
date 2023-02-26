test_that("write_docx works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer"),
    quiet = TRUE
  )

  withr::with_tempdir({
    write_docx(docx, "example.docx")

    expect_error(
      write_docx(docx, "example.docx", overwrite = FALSE)
    )

    write_docx(docx, "example.docx")

    expect_identical(
      docx[["doc_obj"]],
      read_docx_ext(
        filename = "example.docx", quiet = TRUE
      )[["doc_obj"]]
    )
  })
})
