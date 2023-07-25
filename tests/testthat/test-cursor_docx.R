test_that("cursor_docx works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  bmks <- officer::docx_bookmarks(docx)

  expect_error(
    cursor_docx(docx)
  )

  expect_snapshot(
    cursor_docx(docx, keyword = "Title 1")[["officer_cursor"]]
  )

  expect_message(
    cursor_docx(docx, keyword = "Heading 1"),
    "can't be found in `docx`.",
    perl = TRUE
  )

  expect_snapshot(
    cursor_docx(docx, id = bmks[[1]])[["officer_cursor"]]
  )

  expect_snapshot(
    cursor_docx(docx, index = 10)[["officer_cursor"]]
  )

  expect_error(
    cursor_docx(docx, index = 16)
  )

  expect_error(
    cursor_docx(docx)
  )
})
