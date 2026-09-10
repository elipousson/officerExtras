test_that("check_officer works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_null(
    check_docx(docx)
  )

  expect_error(
    check_docx(0)
  )

  expect_error(
    check_pptx(docx)
  )

  expect_error(
    check_xlsx(docx)
  )
})

test_that("check_office_fileext works", {
  expect_null(
    check_docx_fileext("test.docx")
  )
  expect_null(
    check_pptx_fileext("test.pptx")
  )
  expect_null(
    check_xlsx_fileext("test.xlsx")
  )
  expect_error(
    check_docx_fileext("test")
  )
  expect_error(
    check_pptx_fileext("test.docx")
  )
})

test_that("check_officer supports multiple what values and custom arg", {
  docx <- read_officer()

  expect_null(check_officer(docx, what = c("rdocx", "rpptx")))

  expect_error(
    check_officer(docx, what = "rpptx"),
    "rpptx"
  )

  expect_error(
    check_officer(docx, what = "rpptx", arg = "my_arg"),
    "my_arg"
  )

  expect_error(
    check_officer(NULL)
  )
})

test_that("check_block_list validates block list objects", {
  expect_null(
    check_block_list(officer::block_list("a"))
  )

  expect_null(
    check_block_list(officer::block_list(), allow_empty = TRUE)
  )

  expect_error(
    check_block_list(officer::block_list())
  )

  expect_error(
    check_block_list("not a block list")
  )

  expect_error(
    check_block_list(NULL)
  )

  expect_null(
    check_block_list(NULL, allow_null = TRUE)
  )
})

test_that("check_office_fileext works with NULL and allow_null", {
  expect_error(
    check_office_fileext(NULL)
  )

  expect_null(
    check_office_fileext(NULL, allow_null = TRUE)
  )

  expect_null(
    check_office_fileext("test.pptx", fileext = c("docx", "pptx"))
  )

  expect_error(
    check_office_fileext("test.txt")
  )
})

test_that("check_officer_summary validates summary data frames", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )
  summary_df <- officer_summary(docx)

  expect_null(check_officer_summary(summary_df))

  expect_error(
    check_officer_summary("not a data frame")
  )

  expect_null(
    check_officer_summary(summary_df, n = c(1, 1000))
  )

  expect_error(
    check_officer_summary(summary_df, n = 1)
  )

  expect_null(
    check_officer_summary(
      summary_df,
      content_type = c("paragraph", "table cell")
    )
  )

  expect_error(
    check_officer_summary(summary_df, content_type = "image")
  )
})
