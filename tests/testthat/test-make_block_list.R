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
