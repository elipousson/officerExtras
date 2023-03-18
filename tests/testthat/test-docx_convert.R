test_that("docx_convert works", {
    docx <- read_officer(
      system.file("doc_examples/example.docx", package = "officer")
    )

    withr::with_tempdir({
      docx_convert(
        docx,
        output = "test-docx.html",
        path = getwd()
      )

      expect_true(
        file.exists("test-docx.html")
      )

      docx_convert(
        system.file("doc_examples/example.docx", package = "officer"),
        output = "test-docx.pdf",
        path = getwd()
      )

      expect_true(
        file.exists("test-docx.pdf")
      )
    })
})
