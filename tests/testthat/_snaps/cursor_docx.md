# cursor_docx works

    Code
      cursor_docx(docx, keyword = "Title 1")[["officer_cursor"]]
    Output
      $nodes_names
       [1] "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"  
      [13] "p"   "p"   "p"   "tbl" "p"   "p"  
      
      $which
      [1] 1
      
      attr(,"class")
      [1] "officer_cursor"

---

    Code
      cursor_docx(docx, id = bmks[[1]])[["officer_cursor"]]
    Output
      $nodes_names
       [1] "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"  
      [13] "p"   "p"   "p"   "tbl" "p"   "p"  
      
      $which
      [1] 2
      
      attr(,"class")
      [1] "officer_cursor"

---

    Code
      cursor_docx(docx, index = 10)[["officer_cursor"]]
    Output
      $nodes_names
       [1] "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"   "p"  
      [13] "p"   "p"   "p"   "tbl" "p"   "p"  
      
      $which
      [1] 6
      
      attr(,"class")
      [1] "officer_cursor"

