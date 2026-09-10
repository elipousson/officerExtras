test_that("officer_open warns and returns x unmodified when not interactive", {
  docx <- read_officer()

  expect_warning(
    result <- officer_open(docx, interactive = FALSE),
    "non-interactive session"
  )

  expect_identical(result, docx)
})

test_that("officer_open writes and opens the file when interactive", {
  docx <- read_officer()

  withr::with_tempdir({
    local_mocked_bindings(open_file = function(...) invisible(NULL), .package = "officer")

    out_path <- file.path(getwd(), "out.docx")
    result <- officer_open(docx, path = out_path, interactive = TRUE)

    expect_true(file.exists(out_path))
    expect_identical(result, docx)
  })
})
