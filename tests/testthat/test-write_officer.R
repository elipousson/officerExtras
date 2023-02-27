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
        filename = "example.docx", quiet = TRUE
      )[["doc_obj"]]
    )

    write_officer(pptx, "example.pptx")

    expect_identical(
      pptx[["doc_obj"]],
      read_pptx_ext(
        filename = "example.pptx", quiet = TRUE
      )[["doc_obj"]]
    )
  })
})
