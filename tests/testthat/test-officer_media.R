test_that("officer_media works", {
  expect_message(
    officer_media(
      system.file("doc_examples/example.docx", package = "officer")
    ),
    "No media files found in"
  )

  pptx <- read_officer(
    system.file("doc_examples/example.pptx", package = "officer")
  )

  withr::with_tempdir({
    officer_media(
      system.file("doc_examples/example.pptx", package = "officer"),
      target = "test-officer_media"
    )

    expect_true(
      file.exists(file.path("test-officer_media", "image1.png"))
    )

    expect_error(
      officer_media(x = pptx, target = "test-officer_media", overwrite = FALSE)
    )

    file.remove(file.path("test-officer_media", "image1.png"))

    officer_media(x = pptx, target = "test-officer_media", overwrite = FALSE)

    expect_true(
      file.exists(file.path("test-officer_media", "image1.png"))
    )
  })
})
