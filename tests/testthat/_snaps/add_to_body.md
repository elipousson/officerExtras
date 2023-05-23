# add_to_body works

    Code
      docx_text
    Output
      rdocx document with 21 element(s)
      
      * styles:
                      Normal              heading 1              heading 2 
                 "paragraph"            "paragraph"            "paragraph" 
      Default Paragraph Font           Normal Table                No List 
                 "character"                "table"            "numbering" 
                 Titre 1 Car            Titre 2 Car          Light Shading 
                 "character"            "character"                "table" 
              List Paragraph                 header            En-tête Car 
                 "paragraph"            "paragraph"            "character" 
                      footer       Pied de page Car 
                 "paragraph"            "character" 
      
      * Content at cursor location:
        level num_id    text style_name content_type
      1    NA     NA ABCDEFG     Normal    paragraph

---

    Code
      docx_value_key
    Output
      rdocx document with 23 element(s)
      
      * styles:
                      Normal              heading 1              heading 2 
                 "paragraph"            "paragraph"            "paragraph" 
      Default Paragraph Font           Normal Table                No List 
                 "character"                "table"            "numbering" 
                 Titre 1 Car            Titre 2 Car          Light Shading 
                 "character"            "character"                "table" 
              List Paragraph                 header            En-tête Car 
                 "paragraph"            "paragraph"            "character" 
                      footer       Pied de page Car 
                 "paragraph"            "character" 
      
      * Content at cursor location:
        level num_id text style_name content_type
      1    NA     NA  CDE     Normal    paragraph

---

    Code
      docx_value_named
    Output
      rdocx document with 25 element(s)
      
      * styles:
                      Normal              heading 1              heading 2 
                 "paragraph"            "paragraph"            "paragraph" 
      Default Paragraph Font           Normal Table                No List 
                 "character"                "table"            "numbering" 
                 Titre 1 Car            Titre 2 Car          Light Shading 
                 "character"            "character"                "table" 
              List Paragraph                 header            En-tête Car 
                 "paragraph"            "paragraph"            "character" 
                      footer       Pied de page Car 
                 "paragraph"            "character" 
      
      * Content at cursor location:
        level num_id text style_name content_type
      1    NA     NA  CDE     Normal    paragraph

# add_gt_to_body works

    Code
      docx_gt
    Output
      rdocx document with 21 element(s)
      
      * styles:
                      Normal              heading 1              heading 2 
                 "paragraph"            "paragraph"            "paragraph" 
      Default Paragraph Font           Normal Table                No List 
                 "character"                "table"            "numbering" 
                 Titre 1 Car            Titre 2 Car          Light Shading 
                 "character"            "character"                "table" 
              List Paragraph                 header            En-tête Car 
                 "paragraph"            "paragraph"            "character" 
                      footer       Pied de page Car 
                 "paragraph"            "character" 
      
      * Content at cursor location:
           row_id is_header cell_id             text col_span row_span content_type
      1.1       1      TRUE       1                         1        1   table cell
      1.9       2     FALSE       1            grp_a        8        1   table cell
      1.17      3     FALSE       1            row_1        1        1   table cell
      1.25      4     FALSE       1            row_2        1        1   table cell
      1.33      5     FALSE       1            row_3        1        1   table cell
      1.41      6     FALSE       1            row_4        1        1   table cell
      1.49      7     FALSE       1            grp_b        8        1   table cell
      1.57      8     FALSE       1            row_5        1        1   table cell
      1.65      9     FALSE       1            row_6        1        1   table cell
      1.73     10     FALSE       1            row_7        1        1   table cell
      1.81     11     FALSE       1            row_8        1        1   table cell
      2.2       1      TRUE       2              num        1        1   table cell
      2.10      2     FALSE       2             <NA>        0        1   table cell
      2.18      3     FALSE       2        1.111e-01        1        1   table cell
      2.26      4     FALSE       2        2.222e+00        1        1   table cell
      2.34      5     FALSE       2        3.333e+01        1        1   table cell
      2.42      6     FALSE       2        4.444e+02        1        1   table cell
      2.50      7     FALSE       2             <NA>        0        1   table cell
      2.58      8     FALSE       2        5.550e+03        1        1   table cell
      2.66      9     FALSE       2               NA        1        1   table cell
      2.74     10     FALSE       2        7.770e+05        1        1   table cell
      2.82     11     FALSE       2        8.880e+06        1        1   table cell
      3.3       1      TRUE       3             char        1        1   table cell
      3.11      2     FALSE       3             <NA>        0        1   table cell
      3.19      3     FALSE       3          apricot        1        1   table cell
      3.27      4     FALSE       3           banana        1        1   table cell
      3.35      5     FALSE       3          coconut        1        1   table cell
      3.43      6     FALSE       3           durian        1        1   table cell
      3.51      7     FALSE       3             <NA>        0        1   table cell
      3.59      8     FALSE       3               NA        1        1   table cell
      3.67      9     FALSE       3              fig        1        1   table cell
      3.75     10     FALSE       3       grapefruit        1        1   table cell
      3.83     11     FALSE       3         honeydew        1        1   table cell
      4.4       1      TRUE       4             fctr        1        1   table cell
      4.12      2     FALSE       4             <NA>        0        1   table cell
      4.20      3     FALSE       4              one        1        1   table cell
      4.28      4     FALSE       4              two        1        1   table cell
      4.36      5     FALSE       4            three        1        1   table cell
      4.44      6     FALSE       4             four        1        1   table cell
      4.52      7     FALSE       4             <NA>        0        1   table cell
      4.60      8     FALSE       4             five        1        1   table cell
      4.68      9     FALSE       4              six        1        1   table cell
      4.76     10     FALSE       4            seven        1        1   table cell
      4.84     11     FALSE       4            eight        1        1   table cell
      5.5       1      TRUE       5             date        1        1   table cell
      5.13      2     FALSE       5             <NA>        0        1   table cell
      5.21      3     FALSE       5       2015-01-15        1        1   table cell
      5.29      4     FALSE       5       2015-02-15        1        1   table cell
      5.37      5     FALSE       5       2015-03-15        1        1   table cell
      5.45      6     FALSE       5       2015-04-15        1        1   table cell
      5.53      7     FALSE       5             <NA>        0        1   table cell
      5.61      8     FALSE       5       2015-05-15        1        1   table cell
      5.69      9     FALSE       5       2015-06-15        1        1   table cell
      5.77     10     FALSE       5               NA        1        1   table cell
      5.85     11     FALSE       5       2015-08-15        1        1   table cell
      6.6       1      TRUE       6             time        1        1   table cell
      6.14      2     FALSE       6             <NA>        0        1   table cell
      6.22      3     FALSE       6            13:35        1        1   table cell
      6.30      4     FALSE       6            14:40        1        1   table cell
      6.38      5     FALSE       6            15:45        1        1   table cell
      6.46      6     FALSE       6            16:50        1        1   table cell
      6.54      7     FALSE       6             <NA>        0        1   table cell
      6.62      8     FALSE       6            17:55        1        1   table cell
      6.70      9     FALSE       6               NA        1        1   table cell
      6.78     10     FALSE       6            19:10        1        1   table cell
      6.86     11     FALSE       6            20:20        1        1   table cell
      7.7       1      TRUE       7         datetime        1        1   table cell
      7.15      2     FALSE       7             <NA>        0        1   table cell
      7.23      3     FALSE       7 2018-01-01 02:22        1        1   table cell
      7.31      4     FALSE       7 2018-02-02 14:33        1        1   table cell
      7.39      5     FALSE       7 2018-03-03 03:44        1        1   table cell
      7.47      6     FALSE       7 2018-04-04 15:55        1        1   table cell
      7.55      7     FALSE       7             <NA>        0        1   table cell
      7.63      8     FALSE       7 2018-05-05 04:00        1        1   table cell
      7.71      9     FALSE       7 2018-06-06 16:11        1        1   table cell
      7.79     10     FALSE       7 2018-07-07 05:22        1        1   table cell
      7.87     11     FALSE       7               NA        1        1   table cell
      8.8       1      TRUE       8         currency        1        1   table cell
      8.16      2     FALSE       8             <NA>        0        1   table cell
      8.24      3     FALSE       8           49.950        1        1   table cell
      8.32      4     FALSE       8           17.950        1        1   table cell
      8.40      5     FALSE       8            1.390        1        1   table cell
      8.48      6     FALSE       8        65100.000        1        1   table cell
      8.56      7     FALSE       8             <NA>        0        1   table cell
      8.64      8     FALSE       8         1325.810        1        1   table cell
      8.72      9     FALSE       8           13.255        1        1   table cell
      8.80     10     FALSE       8               NA        1        1   table cell
      8.88     11     FALSE       8            0.440        1        1   table cell

---

    Code
      docx_str_keys
    Output
      rdocx document with 23 element(s)
      
      * styles:
                      Normal              heading 1              heading 2 
                 "paragraph"            "paragraph"            "paragraph" 
      Default Paragraph Font           Normal Table                No List 
                 "character"                "table"            "numbering" 
                 Titre 1 Car            Titre 2 Car          Light Shading 
                 "character"            "character"                "table" 
              List Paragraph                 header            En-tête Car 
                 "paragraph"            "paragraph"            "character" 
                      footer       Pied de page Car 
                 "paragraph"            "character" 
      
      * Content at cursor location:
           row_id is_header cell_id             text col_span row_span content_type
      1.1       1      TRUE       1                         1        1   table cell
      1.9       2     FALSE       1            grp_a        8        1   table cell
      1.17      3     FALSE       1            row_1        1        1   table cell
      1.25      4     FALSE       1            row_2        1        1   table cell
      1.33      5     FALSE       1            row_3        1        1   table cell
      1.41      6     FALSE       1            row_4        1        1   table cell
      1.49      7     FALSE       1            grp_b        8        1   table cell
      1.57      8     FALSE       1            row_5        1        1   table cell
      1.65      9     FALSE       1            row_6        1        1   table cell
      1.73     10     FALSE       1            row_7        1        1   table cell
      1.81     11     FALSE       1            row_8        1        1   table cell
      2.2       1      TRUE       2              num        1        1   table cell
      2.10      2     FALSE       2             <NA>        0        1   table cell
      2.18      3     FALSE       2        1.111e-01        1        1   table cell
      2.26      4     FALSE       2        2.222e+00        1        1   table cell
      2.34      5     FALSE       2        3.333e+01        1        1   table cell
      2.42      6     FALSE       2        4.444e+02        1        1   table cell
      2.50      7     FALSE       2             <NA>        0        1   table cell
      2.58      8     FALSE       2        5.550e+03        1        1   table cell
      2.66      9     FALSE       2               NA        1        1   table cell
      2.74     10     FALSE       2        7.770e+05        1        1   table cell
      2.82     11     FALSE       2        8.880e+06        1        1   table cell
      3.3       1      TRUE       3             char        1        1   table cell
      3.11      2     FALSE       3             <NA>        0        1   table cell
      3.19      3     FALSE       3          apricot        1        1   table cell
      3.27      4     FALSE       3           banana        1        1   table cell
      3.35      5     FALSE       3          coconut        1        1   table cell
      3.43      6     FALSE       3           durian        1        1   table cell
      3.51      7     FALSE       3             <NA>        0        1   table cell
      3.59      8     FALSE       3               NA        1        1   table cell
      3.67      9     FALSE       3              fig        1        1   table cell
      3.75     10     FALSE       3       grapefruit        1        1   table cell
      3.83     11     FALSE       3         honeydew        1        1   table cell
      4.4       1      TRUE       4             fctr        1        1   table cell
      4.12      2     FALSE       4             <NA>        0        1   table cell
      4.20      3     FALSE       4              one        1        1   table cell
      4.28      4     FALSE       4              two        1        1   table cell
      4.36      5     FALSE       4            three        1        1   table cell
      4.44      6     FALSE       4             four        1        1   table cell
      4.52      7     FALSE       4             <NA>        0        1   table cell
      4.60      8     FALSE       4             five        1        1   table cell
      4.68      9     FALSE       4              six        1        1   table cell
      4.76     10     FALSE       4            seven        1        1   table cell
      4.84     11     FALSE       4            eight        1        1   table cell
      5.5       1      TRUE       5             date        1        1   table cell
      5.13      2     FALSE       5             <NA>        0        1   table cell
      5.21      3     FALSE       5       2015-01-15        1        1   table cell
      5.29      4     FALSE       5       2015-02-15        1        1   table cell
      5.37      5     FALSE       5       2015-03-15        1        1   table cell
      5.45      6     FALSE       5       2015-04-15        1        1   table cell
      5.53      7     FALSE       5             <NA>        0        1   table cell
      5.61      8     FALSE       5       2015-05-15        1        1   table cell
      5.69      9     FALSE       5       2015-06-15        1        1   table cell
      5.77     10     FALSE       5               NA        1        1   table cell
      5.85     11     FALSE       5       2015-08-15        1        1   table cell
      6.6       1      TRUE       6             time        1        1   table cell
      6.14      2     FALSE       6             <NA>        0        1   table cell
      6.22      3     FALSE       6            13:35        1        1   table cell
      6.30      4     FALSE       6            14:40        1        1   table cell
      6.38      5     FALSE       6            15:45        1        1   table cell
      6.46      6     FALSE       6            16:50        1        1   table cell
      6.54      7     FALSE       6             <NA>        0        1   table cell
      6.62      8     FALSE       6            17:55        1        1   table cell
      6.70      9     FALSE       6               NA        1        1   table cell
      6.78     10     FALSE       6            19:10        1        1   table cell
      6.86     11     FALSE       6            20:20        1        1   table cell
      7.7       1      TRUE       7         datetime        1        1   table cell
      7.15      2     FALSE       7             <NA>        0        1   table cell
      7.23      3     FALSE       7 2018-01-01 02:22        1        1   table cell
      7.31      4     FALSE       7 2018-02-02 14:33        1        1   table cell
      7.39      5     FALSE       7 2018-03-03 03:44        1        1   table cell
      7.47      6     FALSE       7 2018-04-04 15:55        1        1   table cell
      7.55      7     FALSE       7             <NA>        0        1   table cell
      7.63      8     FALSE       7 2018-05-05 04:00        1        1   table cell
      7.71      9     FALSE       7 2018-06-06 16:11        1        1   table cell
      7.79     10     FALSE       7 2018-07-07 05:22        1        1   table cell
      7.87     11     FALSE       7               NA        1        1   table cell
      8.8       1      TRUE       8         currency        1        1   table cell
      8.16      2     FALSE       8             <NA>        0        1   table cell
      8.24      3     FALSE       8           49.950        1        1   table cell
      8.32      4     FALSE       8           17.950        1        1   table cell
      8.40      5     FALSE       8            1.390        1        1   table cell
      8.48      6     FALSE       8        65100.000        1        1   table cell
      8.56      7     FALSE       8             <NA>        0        1   table cell
      8.64      8     FALSE       8         1325.810        1        1   table cell
      8.72      9     FALSE       8           13.255        1        1   table cell
      8.80     10     FALSE       8               NA        1        1   table cell
      8.88     11     FALSE       8            0.440        1        1   table cell

# add_gg_to_body works

    Code
      docx_gg
    Output
      rdocx document with 22 element(s)
      
      * styles:
                      Normal              heading 1              heading 2 
                 "paragraph"            "paragraph"            "paragraph" 
      Default Paragraph Font           Normal Table                No List 
                 "character"                "table"            "numbering" 
                 Titre 1 Car            Titre 2 Car          Light Shading 
                 "character"            "character"                "table" 
              List Paragraph                 header            En-tête Car 
                 "paragraph"            "paragraph"            "character" 
                      footer       Pied de page Car 
                 "paragraph"            "character" 
      
      * Content at cursor location:
        level num_id       text style_name content_type
      1    NA     NA test title     Normal    paragraph

