test_that("is_officer and variants identify object classes", {
  docx <- read_officer()
  pptx <- read_officer(fileext = "pptx")

  expect_true(is_officer(docx))
  expect_true(is_officer(docx, "rdocx"))
  expect_false(is_officer(docx, "rpptx"))
  expect_true(is_officer(pptx, "rpptx"))
  expect_false(is_officer("not an officer object"))

  expect_true(is_rdocx(docx))
  expect_false(is_rdocx(pptx))

  expect_true(is_rpptx(pptx))
  expect_false(is_rpptx(docx))
})

test_that("is_block_list identifies block_list objects", {
  expect_true(is_block_list(officer::block_list("a")))
  expect_true(is_block_list(officer::block_list()))
  expect_false(is_block_list(list("a")))
  expect_false(is_block_list(NULL))
})
