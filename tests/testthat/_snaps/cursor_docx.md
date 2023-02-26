# cursor_docx works

    Code
      cursor_docx(docx, keyword = "Title 1")[["officer_cursor"]]
    Output
      $nodes_names
       [1] "p"           "bookmarkEnd" "p"           "p"           "p"          
       [6] "p"           "p"           "p"           "p"           "p"          
      [11] "p"           "p"           "p"           "p"           "p"          
      [16] "p"           "tbl"         "p"           "p"          
      
      $which
      [1] 1
      
      attr(,"class")
      [1] "officer_cursor"

---

    Code
      cursor_docx(docx, index = 10)[["officer_cursor"]]
    Output
      $nodes_names
       [1] "p"           "bookmarkEnd" "p"           "p"           "p"          
       [6] "p"           "p"           "p"           "p"           "p"          
      [11] "p"           "p"           "p"           "p"           "p"          
      [16] "p"           "tbl"         "p"           "p"          
      
      $which
      [1] 7
      
      attr(,"class")
      [1] "officer_cursor"

