test_that("read_docs_url works", {

  url <- "https://docs.google.com/document/d/1w49ndk0vCvIA-cf-gVml85roV0FD-Z3cbz-_YWbRklg/edit"

  withr::with_tempdir({
    docx <- suppressWarnings(
      read_docs_url(url, filename = "testfile.docx")
    )
    expect_true(file.exists("testfile.docx"))
    expect_true(is_officer(docx, "rdocx"))
  })
})
