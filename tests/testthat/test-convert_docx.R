test_that("convert_docx converts a rdocx object and a file path", {
  skip_if_not_installed("rmarkdown")
  skip_if_not(rmarkdown::pandoc_available())

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
      output = "test-docx-2.html",
      path = path
    )

    expect_true(
      file.exists(file.path(path, "test-docx-2.html"))
    )
  })
})

test_that("convert_docx errors if output exists and overwrite is FALSE", {
  skip_if_not_installed("rmarkdown")
  skip_if_not(rmarkdown::pandoc_available())

  docx <- read_officer(
    system.file("doc_examples/example.docx", package = "officer")
  )

  withr::with_tempdir({
    convert_docx(docx, output = "test-docx.html", path = getwd())

    expect_error(
      convert_docx(docx, output = "test-docx.html", path = getwd(), overwrite = FALSE)
    )
  })
})

test_that("convert_docx treats a fileext-bearing `to` as the output name", {
  skip_if_not_installed("rmarkdown")
  skip_if_not(rmarkdown::pandoc_available())

  docx <- read_officer(
    system.file("doc_examples/example.docx", package = "officer")
  )

  withr::with_tempdir({
    convert_docx(docx, to = "test-docx.html", path = getwd())

    expect_true(file.exists(file.path(getwd(), "test-docx.html")))
  })
})
