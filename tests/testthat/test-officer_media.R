test_that("officer_media works", {
  expect_message(
    officer_media(system.file("doc_examples/example.docx", package = "officer")),
    "No media files found in"
  )

  withr::with_tempdir({
    officer_media(
      system.file("doc_examples/example.pptx", package = "officer"),
      dir = "test-officer_media"
    )

    expect_true(
      file.exists(file.path("test-officer_media", "image1.png"))
    )
  })
})
