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

  expect_identical(
    names(officer_tables(pptx, nm = "custom_name")),
    rep("custom_name", length(officer_tables(pptx)))
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

test_that("officer_table extracts a single table as a data frame", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )
  summary_df <- officer_summary(docx)

  tbl <- officer_table(summary_df)

  expect_s3_class(tbl, "data.frame")
  expect_named(tbl, c("Petals", "Internode", "Sepal", "Bract"))

  # has_header only affects tables without an explicit header row (is_header);
  # this table has one, so has_header = FALSE has no effect on column names
  tbl_no_header <- officer_table(summary_df, has_header = FALSE)
  expect_named(tbl_no_header, c("Petals", "Internode", "Sepal", "Bract"))

  pptx <- read_pptx_ext(
    system.file("doc_examples/example.pptx", package = "officer")
  )
  pptx_summary <- officer_summary(pptx)

  expect_named(
    officer_table(pptx_summary, index = "18", has_header = FALSE),
    c("X1", "X2", "X3")
  )
})

test_that("officer_tables errors with stack = TRUE for varying column counts", {
  skip_if_not_installed("gt")

  withr::with_tempdir({
    docx <- read_docx_ext(
      filename = "example.docx",
      path = system.file("doc_examples", package = "officer")
    )

    tab_1 <- gt::gt(gt::exibble[1:2, 1:2])
    docx <- add_gt_to_body(docx, tab_1, keyword = "Sub title 1")

    write_officer(docx, "combo.docx")
    docx <- read_officer("combo.docx")

    expect_error(
      officer_tables(docx, stack = TRUE),
      "varying"
    )
  })
})

test_that("officer_tables supports stack = TRUE", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  stacked <- officer_tables(docx, stack = TRUE)

  expect_s3_class(stacked, "data.frame")
  expect_named(stacked, c("Petals", "Internode", "Sepal", "Bract"))
})
