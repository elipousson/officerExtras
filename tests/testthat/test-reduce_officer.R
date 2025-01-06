test_that("reduce_officer works", {
  example_path <- system.file("doc_examples/example.docx", package = "officer")
  x <- reduce_officer(example_path, value = LETTERS)

  expect_s3_class(
    x,
    "rdocx"
  )

  withr::with_tempdir({
    x <- reduce_officer(
      example_path,
      value = LETTERS,
      .path = "example_file_with_letters.docx"
    )

    expect_true(
      file.exists("example_file_with_letters.docx")
    )
  })
})
