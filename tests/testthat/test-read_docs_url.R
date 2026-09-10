test_that("read_docs_url works", {
  skip_if_offline()

  url <- "https://docs.google.com/document/d/1w49ndk0vCvIA-cf-gVml85roV0FD-Z3cbz-_YWbRklg/edit"

  withr::with_tempdir({
    docx <- suppressWarnings(
      read_docs_url(url, filename = "testfile.docx")
    )
    expect_true(file.exists("testfile.docx"))
    expect_true(is_officer(docx, "rdocx"))
  })
})

test_that("read_docs_url downloads, reads, and optionally cleans up the file", {
  example_path <- system.file("doc_examples/example.docx", package = "officer")
  url <- "https://docs.google.com/document/d/ABC123/edit"

  withr::with_tempdir({
    local_mocked_bindings(
      download.file = function(url, destfile, ...) {
        file.copy(example_path, destfile, overwrite = TRUE)
        invisible(0L)
      },
      .package = "utils"
    )

    docx_named <- read_docs_url(url, filename = "myfile.docx")
    expect_true(file.exists("myfile.docx"))
    expect_true(is_rdocx(docx_named))

    docx_unnamed <- read_docs_url(url)
    expect_true(is_rdocx(docx_unnamed))

    dir.create("subdir")
    docx_path_only <- read_docs_url(url, path = "subdir")
    expect_true(is_rdocx(docx_path_only))
    # file is removed after export since filename wasn't supplied
    expect_length(list.files("subdir"), 0)
  })
})

test_that("is_docs_url family identify Google Docs URL types", {
  doc_url <- "https://docs.google.com/document/d/ABC123/edit"
  slides_url <- "https://docs.google.com/presentation/d/XYZ789/edit"
  sheets_url <- "https://docs.google.com/spreadsheets/d/QRS456/edit"

  expect_true(officerExtras:::is_gdoc_url(doc_url))
  expect_true(officerExtras:::is_docs_url(doc_url))
  expect_false(officerExtras:::is_docs_url(slides_url))

  expect_true(officerExtras:::is_slides_url(slides_url))
  expect_false(officerExtras:::is_slides_url(doc_url))

  expect_true(officerExtras:::is_sheets_url(sheets_url))
  expect_false(officerExtras:::is_sheets_url(doc_url))

  expect_false(officerExtras:::is_gdoc_url("https://example.com"))
})

test_that("extract_docs_id extracts the document id from a URL", {
  doc_url <- "https://docs.google.com/document/d/ABC123/edit"

  expect_identical(
    officerExtras:::extract_docs_id(doc_url, "edit"),
    "ABC123"
  )
})

test_that("prep_docs_export builds export URLs for docs, slides, and sheets", {
  doc_url <- "https://docs.google.com/document/d/ABC123/edit"
  slides_url <- "https://docs.google.com/presentation/d/XYZ789/export"
  sheets_url <- "https://docs.google.com/spreadsheets/d/QRS456/edit"

  doc_export <- officerExtras:::prep_docs_export(doc_url)
  expect_identical(
    doc_export[["url"]],
    "https://docs.google.com/document/d/ABC123/export?format=doc"
  )
  expect_match(doc_export[["filename"]], "\\.docx$")

  slides_export <- officerExtras:::prep_docs_export(slides_url, format = "pptx")
  expect_identical(
    slides_export[["url"]],
    "https://docs.google.com/presentation/d/XYZ789/export/pptx"
  )
  expect_match(slides_export[["filename"]], "\\.pptx$")

  sheets_export <- officerExtras:::prep_docs_export(sheets_url, format = "xlsx")
  expect_identical(
    sheets_export[["url"]],
    "https://docs.google.com/spreadsheets/d/QRS456/export?format=xlsx"
  )
  expect_match(sheets_export[["filename"]], "\\.xlsx$")
})
