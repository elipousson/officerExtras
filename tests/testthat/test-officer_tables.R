test_that("officer_tables works", {
  pptx <-
    read_pptx_ext(
      system.file("doc_examples/example.pptx", package = "officer")
    )

  expect_type(
    officer_tables(pptx),
    "list"
  )

  pptx_tab <- officer_tables(pptx)[[1]]

  expect_s3_class(
    pptx_tab,
    "data.frame"
  )

  expect_named(
    pptx_tab,
    c("Header 1 ", "Header 2", "Header 3")
  )

  expect_named(
    officer_tables(pptx, has_header = FALSE)[[1]],
    c("X1", "X2", "X3")
  )

  expect_error(
    officer_tables("pptx"),
    "must be a"
  )

  withr::with_tempdir({
    skip_if_not_installed("gt")

    docx <- read_docx_ext(
      filename = "example.docx",
      path = system.file("doc_examples", package = "officer")
    )

    tab_1 <-
      gt::gt(
        gt::exibble,
        rowname_col = "row",
        groupname_col = "group"
      )

    write_officer(
      add_gt_to_body(
        docx,
        tab_1,
        keyword = "Sub title 1"
      ),
      "test-officer_tables.docx"
    )

    docx <- read_officer("test-officer_tables.docx")

    type_convert_tbl <- officer_tables(docx, type_convert = TRUE)[[1]]

    expect_s3_class(
      type_convert_tbl,
      "data.frame"
    )
    expect_type(
      c(type_convert_tbl[[2]], type_convert_tbl[[8]]),
      "double"
    )
  })
})
