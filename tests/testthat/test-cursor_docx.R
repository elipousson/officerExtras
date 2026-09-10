test_that("cursor_docx moves the cursor with keyword, id, or index", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  bmks <- officer::docx_bookmarks(docx)

  cur_keyword <- cursor_docx(docx, keyword = "Title 1")[["officer_cursor"]]
  expect_s3_class(cur_keyword, "officer_cursor")
  expect_identical(cur_keyword[["which"]], 1L)

  cur_id <- cursor_docx(docx, id = bmks[[1]])[["officer_cursor"]]
  expect_identical(cur_id[["which"]], 2L)

  cur_index <- cursor_docx(docx, index = 10)[["officer_cursor"]]
  expect_identical(cur_index[["which"]], 6L)
})

test_that("cursor_docx warns and returns docx unmodified for a missing keyword", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  cur_before <- docx[["officer_cursor"]]

  expect_message(
    docx_after <- cursor_docx(docx, keyword = "not a real keyword"),
    "can't be found"
  )

  expect_identical(docx_after[["officer_cursor"]], cur_before)
})

test_that("cursor_docx suppresses the missing keyword message when quiet is TRUE", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  cur_before <- docx[["officer_cursor"]]

  expect_no_message(
    docx_after <- cursor_docx(docx, keyword = "not a real keyword", quiet = TRUE)
  )

  expect_identical(docx_after[["officer_cursor"]], cur_before)
})

test_that("cursor_docx errors for an invalid bookmark id or out-of-range index", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_error(
    cursor_docx(docx, id = "not_a_bookmark")
  )

  expect_error(
    cursor_docx(docx, index = 1000)
  )
})

test_that("cursor_docx supports the default position options", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  cur_begin <- cursor_docx(docx, default = "begin")[["officer_cursor"]]
  expect_identical(cur_begin[["which"]], 1L)

  cur_end <- cursor_docx(docx, default = "end")[["officer_cursor"]]
  expect_identical(cur_end[["which"]], length(cur_end[["nodes_names"]]))

  docx_begin <- cursor_docx(docx, default = "begin")
  cur_forward <- cursor_docx(docx_begin, default = "forward")[["officer_cursor"]]
  expect_identical(cur_forward[["which"]], 2L)

  docx_end <- cursor_docx(docx, default = "end")
  cur_backward <- cursor_docx(docx_end, default = "backward")[["officer_cursor"]]
  expect_identical(
    cur_backward[["which"]],
    length(cur_backward[["nodes_names"]]) - 1L
  )

  expect_error(
    cursor_docx(docx, default = "not-a-real-option")
  )
})
