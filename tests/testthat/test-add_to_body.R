test_that("add_to_body errors when str/value are missing or both supplied", {
  docx_example <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_error(add_to_body(docx_example))
  expect_error(add_to_body(docx_example, value = "a", str = "<w:p/>"))
})

test_that("add_to_body adds paragraph text at a keyword position", {
  docx_example <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  n_rows <- nrow(officer_summary(docx_example))

  docx_text <- add_to_body(
    docx_example,
    keyword = "Title 1",
    value = "ABCDEFG"
  )

  summary_df <- officer_summary(docx_text)

  expect_equal(nrow(summary_df), n_rows + 1)
  expect_true("ABCDEFG" %in% summary_df[["text"]])
})

test_that("add_text_to_body supports glue string interpolation", {
  docx_example <- read_officer()
  value <- "world"

  docx_text <- add_text_to_body(
    docx_example,
    value = "hello {value}"
  )

  expect_identical(
    officer_summary(docx_text)[["text"]],
    "hello world"
  )
})

test_that("add_xml_to_body adds a raw xml string", {
  docx_example <- read_officer()

  xml_str <- "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
    <w:r><w:t>xml text</w:t></w:r>
  </w:p>"

  docx_xml <- add_xml_to_body(docx_example, str = xml_str)

  expect_identical(
    officer_summary(docx_xml)[["text"]],
    "xml text"
  )
})

test_that("add_value_with_keys works with keyword vectors and named vectors", {
  docx_example <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  docx_value_key <- add_value_with_keys(
    docx_example,
    value = c("ABC", "CDE"),
    keyword = c("Title 1", "Title 2")
  )

  summary_df <- officer_summary(docx_value_key)
  expect_true(all(c("ABC", "CDE") %in% summary_df[["text"]]))

  docx_value_named <- add_value_with_keys(
    docx_example,
    value = c("Title 1" = "ABC", "Title 2" = "CDE")
  )

  summary_named_df <- officer_summary(docx_value_named)
  expect_true(all(c("ABC", "CDE") %in% summary_named_df[["text"]]))

  expect_error(
    add_value_with_keys(docx_example, value = c("a", "b"), keyword = "Title 1")
  )
})

test_that("add_str_with_keys works with keyword vectors", {
  docx_example <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  xml_ns <- "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""

  str_1 <- paste0(
    "<w:p ", xml_ns, "><w:r><w:t>str one</w:t></w:r></w:p>"
  )
  str_2 <- paste0(
    "<w:p ", xml_ns, "><w:r><w:t>str two</w:t></w:r></w:p>"
  )

  docx_str_keys <- add_str_with_keys(
    docx_example,
    str = c("Title 1" = str_1, "Title 2" = str_2)
  )

  summary_df <- officer_summary(docx_str_keys)
  expect_true(all(c("str one", "str two") %in% summary_df[["text"]]))

  expect_error(
    add_str_with_keys(docx_example, str = c("a", "b"), keyword = "Title 1")
  )
})

test_that("add_gt_to_body adds a gt table", {
  skip_if_not_installed("gt")

  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  tab_1 <- gt::gt(gt::exibble[1:2, 1:2])

  docx_gt <- add_gt_to_body(
    docx,
    tab_1,
    keyword = "Sub title 1"
  )

  summary_df <- officer_summary(docx_gt)
  table_cells <- summary_df[summary_df[["content_type"]] == "table cell", ]

  expect_true(all(c("apricot", "banana") %in% table_cells[["text"]]))
})

test_that("add_gt_to_body supports tablecontainer = FALSE", {
  skip_if_not_installed("gt")

  docx <- read_officer()
  tab_1 <- gt::gt(gt::exibble[1:2, 1:2])

  docx_gt <- add_gt_to_body(docx, tab_1, tablecontainer = FALSE)

  summary_df <- officer_summary(docx_gt)

  expect_true(all(summary_df[["content_type"]] == "table cell"))
  expect_true(all(c("apricot", "banana") %in% summary_df[["text"]]))
})

test_that("add_gg_to_body adds a plot and caption", {
  skip_if_not_installed("ggplot2")

  docx <- read_officer()

  plot1 <- ggplot2::ggplot(mtcars, ggplot2::aes(mpg, wt)) +
    ggplot2::geom_point() +
    ggplot2::labs(title = "test title")

  docx_gg <- add_gg_to_body(docx, plot1)

  summary_df <- officer_summary(docx_gg)

  expect_true("test title" %in% summary_df[["text"]])
})

test_that("add_gg_to_body supports autonum captions", {
  skip_if_not_installed("ggplot2")

  docx <- read_officer()

  plot1 <- ggplot2::ggplot(mtcars, ggplot2::aes(mpg, wt)) +
    ggplot2::geom_point() +
    ggplot2::labs(title = "cap title")

  docx_gg <- add_gg_to_body(
    docx,
    plot1,
    autonum = officer::run_autonum(seq_id = "fig")
  )

  summary_df <- officer_summary(docx_gg)

  expect_true(any(grepl("cap title", summary_df[["text"]], fixed = TRUE)))
})

test_that("add_gg_to_body skips caption when caption is NULL", {
  skip_if_not_installed("ggplot2")

  docx <- read_officer()

  plot1 <- ggplot2::ggplot(mtcars, ggplot2::aes(mpg, wt)) +
    ggplot2::geom_point() +
    ggplot2::labs(title = "test title")

  docx_gg <- add_gg_to_body(docx, plot1, caption = NULL)

  summary_df <- officer_summary(docx_gg)

  expect_false("test title" %in% summary_df[["text"]])
})
