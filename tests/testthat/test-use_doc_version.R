test_that("use_doc_version works", {
  example_path <- system.file("doc_examples/example.docx", package = "officer")

  withr::with_tempdir({
    use_doc_version(
      filename = example_path,
      which = "major",
      path = getwd()
    )

    expect_true(
      file.exists(
        "example_1.0.0.docx"
      )
    )

    use_doc_version(
      filename = "example_1.0.0.docx",
      which = "minor"
    )

    expect_true(
      file.exists(
        "example_1.1.0.docx"
      )
    )
  })
})
