test_that("officer_add_blocks works", {
  docx <- read_officer()

  docx <- officer_add_blocks(
    docx,
    blocks = officer::block_list("text", "text"),
    pos = "on"
  )

  docx_summary <- officer_summary(docx)

  expect_equal(
    nrow(docx_summary),
    2
  )

  expect_equal(
    docx_summary[["text"]],
    c("text", "text")
  )

  pptx <- read_officer(fileext = "pptx")
  pptx <- officer::add_slide(pptx, "Title and Content")

  pptx <- officer_add_blocks(
    pptx,
    blocks = officer::block_list("text"),
    location = officer::ph_location_type()
  )

  expect_identical(
    officer_summary(pptx, "slide", index = 1)[["text"]],
    "text"
  )
})
