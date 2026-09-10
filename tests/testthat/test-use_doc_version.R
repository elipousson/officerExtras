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

test_that("use_doc_version supports save = FALSE and returns invisibly", {
  example_path <- system.file("doc_examples/example.docx", package = "officer")

  withr::with_tempdir({
    x <- use_doc_version(
      filename = example_path,
      which = "major",
      save = FALSE
    )

    expect_true(is_rdocx(x))
    expect_identical(officer_properties(x)[["version"]], "1.0.0")
    expect_length(list.files(), 0)
  })
})

test_that("use_doc_version supports which as a numeric position", {
  example_path <- system.file("doc_examples/example.docx", package = "officer")

  x <- use_doc_version(
    filename = example_path,
    which = 2,
    save = FALSE
  )

  expect_identical(officer_properties(x)[["version"]], "0.2.0")
})

test_that("use_doc_version supports dev versions and custom separators", {
  example_path <- system.file("doc_examples/example.docx", package = "officer")

  x_dev <- use_doc_version(
    filename = example_path,
    which = "dev",
    save = FALSE
  )
  expect_identical(officer_properties(x_dev)[["version"]], "0.1.0.9000")

  x_sep <- use_doc_version(
    filename = example_path,
    which = "patch",
    sep = "-",
    save = FALSE
  )
  expect_identical(officer_properties(x_sep)[["version"]], "0-1-1")
})

test_that("doc_version extracts a version from a filename", {
  expect_identical(
    doc_version(filename = "myfile_1.2.3.docx"),
    "1.2.3"
  )
})

test_that("doc_version returns the default version when nothing is supplied", {
  expect_identical(
    doc_version(allow_new = TRUE),
    "0.1.0"
  )

  expect_identical(
    doc_version(allow_new = TRUE, .default = c(2, 0, 0)),
    "2.0.0"
  )
})

test_that("doc_version reads the version property from an officer object", {
  example_path <- system.file("doc_examples/example.docx", package = "officer")
  docx <- read_officer(example_path)

  expect_identical(
    doc_version(x = docx),
    "0.1.0"
  )

  docx_versioned <- officer::set_doc_properties(docx, version = "3.4.5")

  expect_identical(
    doc_version(x = docx_versioned),
    "3.4.5"
  )
})

test_that("use_doc_version trims a 4-component version back to 3 when the dev component is zeroed", {
  docx <- read_officer()
  docx <- officer::set_doc_properties(docx, version = "1.2.3.9000")

  x <- use_doc_version(x = docx, which = "patch", save = FALSE)

  expect_identical(officer_properties(x)[["version"]], "1.2.4")
})

test_that("use_doc_version supports a prefix based on a document property", {
  example_path <- system.file("doc_examples/example.docx", package = "officer")

  withr::with_tempdir({
    use_doc_version(
      filename = example_path,
      which = "major",
      path = getwd(),
      prefix = "modified"
    )

    expect_length(list.files(pattern = "example_1.0.0.docx$"), 1)
  })
})

test_that("use_doc_version keeps the version as a filename suffix when prefix equals property", {
  example_path <- system.file("doc_examples/example.docx", package = "officer")

  withr::with_tempdir({
    file.copy(example_path, "example.docx")

    use_doc_version(
      filename = "example.docx",
      which = "major",
      prefix = "version"
    )

    expect_true(file.exists("example_1.0.0.docx"))
  })
})

test_that("doc_version errors when allow_new is FALSE and no version is found", {
  example_path <- system.file("doc_examples/example.docx", package = "officer")
  docx <- read_officer(example_path)

  expect_error(
    doc_version(x = docx, allow_new = FALSE),
    class = "rlang_error"
  )

  expect_error(
    doc_version(filename = example_path, allow_new = FALSE)
  )
})
