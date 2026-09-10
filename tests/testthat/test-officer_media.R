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

test_that("officer_media supports list = TRUE", {
  expect_message(
    media_files <- officer_media(
      system.file("doc_examples/example.pptx", package = "officer"),
      list = TRUE
    ),
    "1 media file found"
  )

  expect_true(any(grepl("image1.png", media_files)))
})

test_that("officer_media works with a rpptx object as x", {
  pptx <- read_officer(
    system.file("doc_examples/example.pptx", package = "officer")
  )

  withr::with_tempdir({
    officer_media(x = pptx, target = "media_from_object")

    expect_true(
      file.exists(file.path("media_from_object", "image1.png"))
    )
  })
})
