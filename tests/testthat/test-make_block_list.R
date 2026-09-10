test_that("make_block_list works", {
  expect_error(make_block_list())
  expect_identical(
    make_block_list(allow_empty = TRUE),
    officer::block_list()
  )
  expect_identical(
    make_block_list(blocks = list("text2"), "text"),
    officer::block_list("text2", "text")
  )
})

test_that("make_block_list works with block_list blocks and errors on invalid input", {
  expect_identical(
    make_block_list(blocks = officer::block_list("a")),
    officer::block_list("a")
  )

  expect_error(
    make_block_list(blocks = list())
  )

  expect_identical(
    make_block_list(blocks = list(), allow_empty = TRUE),
    officer::block_list()
  )
})

test_that("combine_blocks combines block_list objects", {
  b1 <- officer::block_list("a", "b")
  b2 <- officer::block_list("c")

  combined <- combine_blocks(b1, b2)

  expect_s3_class(combined, "block_list")
  expect_length(combined, 3)

  expect_error(
    combine_blocks(b1, "not a block list")
  )
})
