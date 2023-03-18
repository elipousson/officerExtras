test_that("convert_docx works", {
  skip_on_ci()
  docx <- read_officer(
    system.file("doc_examples/example.docx", package = "officer")
  )

  withr::with_tempdir({
    path <- getwd()

    convert_docx(
      docx,
      output = "test-docx.html",
      path = path
    )

    expect_true(
      file.exists(file.path(path, "test-docx.html"))
    )

    convert_docx(
      system.file("doc_examples/example.docx", package = "officer"),
      output = "test-docx.pdf",
      path = getwd()
    )

    expect_true(
      file.exists(file.path(path, "test-docx.pdf"))
    )
  })
})
