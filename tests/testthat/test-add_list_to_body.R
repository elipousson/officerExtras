test_that("add_list_to_body adds a bulleted list", {
  docx <- read_docx_ext(allow_null = TRUE)

  docx_list <- add_list_to_body(
    docx,
    values = c("item 1", "item 2", "item 3")
  )

  summary_df <- officer_summary(docx_list)
  list_rows <- summary_df[summary_df[["style_name"]] == "List Bullet", ]

  expect_identical(
    list_rows[["text"]],
    c("item 1", "item 2", "item 3")
  )
})

test_that("add_list_to_body drops NA values by default", {
  docx <- read_docx_ext(allow_null = TRUE)

  docx_list <- add_list_to_body(
    docx,
    values = c("item 1", NA, "item 2")
  )

  list_rows <- officer_summary(docx_list)
  list_rows <- list_rows[list_rows[["style_name"]] == "List Bullet", ]

  expect_identical(
    list_rows[["text"]],
    c("item 1", "item 2")
  )
})

test_that("add_list_to_body keeps NA values when keep_na is TRUE", {
  docx <- read_docx_ext(allow_null = TRUE)

  docx_list <- add_list_to_body(
    docx,
    values = c("item 1", NA),
    keep_na = TRUE
  )

  list_rows <- officer_summary(docx_list)
  list_rows <- list_rows[list_rows[["style_name"]] == "List Bullet", ]

  expect_equal(nrow(list_rows), 2)
})

test_that("add_list_to_body returns docx unmodified for empty or all-NA values", {
  docx <- read_docx_ext(allow_null = TRUE)
  n_rows <- nrow(officer_summary(docx))

  docx_empty <- add_list_to_body(docx, values = character(0))
  expect_equal(nrow(officer_summary(docx_empty)), n_rows)

  docx_na <- add_list_to_body(docx, values = c(NA_character_, NA_character_))
  expect_equal(nrow(officer_summary(docx_na)), n_rows)
})

test_that("add_list_to_body supports before and after values", {
  docx <- read_docx_ext(allow_null = TRUE)

  docx_list <- add_list_to_body(
    docx,
    values = c("item 1", "item 2"),
    before = "Before text",
    after = "After text"
  )

  summary_df <- officer_summary(docx_list)

  expect_identical(
    summary_df[["text"]],
    c("Before text", "item 1", "item 2", "After text")
  )
})

test_that("add_list_to_body returns the modified docx when after is NULL", {
  docx <- read_docx_ext(allow_null = TRUE)

  docx_list <- add_list_to_body(
    docx,
    values = c("item 1", "item 2"),
    after = NULL
  )

  expect_true(is_rdocx(docx_list))
  expect_identical(
    officer_summary(docx_list)[["text"]],
    c("item 1", "item 2")
  )
})

test_that("add_blocks_to_body and officer_add_blocks add a block list", {
  docx <- read_officer()

  docx_blocks <- add_blocks_to_body(
    docx,
    blocks = officer::block_list("a", "b")
  )

  expect_identical(
    officer_summary(docx_blocks)[["text"]],
    c("a", "b")
  )

  docx2 <- read_officer()

  docx2 <- officer_add_blocks(
    docx2,
    blocks = officer::block_list("text", "text")
  )

  expect_identical(
    officer_summary(docx2)[["text"]],
    c("text", "text")
  )

  expect_error(
    officer_add_blocks(docx2, blocks = "not a block list")
  )

  expect_error(
    officer_add_blocks("not a docx", blocks = officer::block_list("text"))
  )
})

test_that("add_blocks_to_body positions the cursor with keyword, id, or index", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  docx_blocks <- add_blocks_to_body(
    docx,
    blocks = officer::block_list("inserted block"),
    keyword = "Title 1"
  )

  summary_df <- officer_summary(docx_blocks)
  idx <- which(summary_df[["text"]] == "inserted block")

  expect_length(idx, 1)
})
