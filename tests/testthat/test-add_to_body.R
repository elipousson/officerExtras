test_that("add_to_body works", {
  docx_example <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_error(add_to_body(docx_example))

  docx_text <- add_text_to_body(
    docx_example,
    value = "ABCDEFG",
    keyword = "Title 1"
  )

  expect_snapshot(docx_text)

  docx_value_key <- add_value_with_keys(
    docx_example,
    value = c("ABC", "CDE"),
    keyword = c("Title 1", "Title 2")
  )

  expect_snapshot(docx_value_key)

  docx_value_named <- add_value_with_keys(
    docx_example,
    value = c("Title 1" = "ABC", "Title 2" = "CDE")
  )

  expect_snapshot(docx_value_named)
})

test_that("add_gt_to_body works", {
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

  docx_gt <-
    add_gt_to_body(
      docx,
      tab_1,
      keyword = "Sub title 1"
    )

  expect_snapshot(docx_gt)

  tab_str <- gt::as_word(tab_1)

  docx_str_keys <-
    add_str_with_keys(
      docx,
      str = c("Title 1" = tab_str, "Title 2" = tab_str)
    )

  expect_snapshot(docx_str_keys)
})

test_that("add_gg_to_body works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  skip_if_not_installed("ggplot2")

  set.seed(1)
  df <- data.frame(
    gp = factor(rep(letters[1:3], each = 10)),
    y = rnorm(30)
  )
  ds <- do.call(rbind, lapply(split(df, df$gp), function(d) {
    data.frame(mean = mean(d$y), sd = sd(d$y), gp = d$gp)
  }))

  plot1 <- ggplot2::ggplot(df, ggplot2::aes(gp, y)) +
    ggplot2::geom_point() +
    ggplot2::geom_point(data = ds, ggplot2::aes(y = mean), colour = "red", size = 3) +
    ggplot2::labs(
      title = "test title"
    )

  docx_gg <-
    add_gg_to_body(
      docx,
      plot1
    )

  expect_snapshot(docx_gg)
})
