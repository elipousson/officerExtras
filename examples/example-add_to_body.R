if (rlang::is_installed("gt")) {
  docx_example <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  tab_1 <-
    gt::gt(
      gt::exibble,
      rowname_col = "row",
      groupname_col = "group"
    )

  add_gt_to_body(
    docx_example,
    tab_1,
    keyword = "Title 1"
  )

  tab_str <- gt::as_word(tab_1)

  add_str_with_keys(
    docx_example,
    str = c("Title 1" = tab_str, "Title 2" = tab_str)
  )
}
