if (rlang::is_installed("gt")) {
  tab_1 <-
    gt::gt(
      gt::exibble,
      rowname_col = "row",
      groupname_col = "group"
    )

  tab_str <- gt::as_word(tab_1)

  docx_str_keys <-
    add_str_with_keys(
      docx,
      str = c("Title 1" = tab_str, "Title 2" = tab_str)
    )
}
