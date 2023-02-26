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
      docx
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
      1.9       2     FALSE       1            grp_a        1        1   table cell
      1.10      3     FALSE       1            row_1        1        1   table cell
      1.18      4     FALSE       1            row_2        1        1   table cell
      1.26      5     FALSE       1            row_3        1        1   table cell
      1.34      6     FALSE       1            row_4        1        1   table cell
      1.42      7     FALSE       1            grp_b        1        1   table cell
      1.43      8     FALSE       1            row_5        1        1   table cell
      1.51      9     FALSE       1            row_6        1        1   table cell
      1.59     10     FALSE       1            row_7        1        1   table cell
      1.67     11     FALSE       1            row_8        1        1   table cell
      2.2       1      TRUE       2              num        1        1   table cell
      2.11      3     FALSE       2        1.111e-01        1        1   table cell
      2.19      4     FALSE       2        2.222e+00        1        1   table cell
      2.27      5     FALSE       2        3.333e+01        1        1   table cell
      2.35      6     FALSE       2        4.444e+02        1        1   table cell
      2.44      8     FALSE       2        5.550e+03        1        1   table cell
      2.52      9     FALSE       2               NA        1        1   table cell
      2.60     10     FALSE       2        7.770e+05        1        1   table cell
      2.68     11     FALSE       2        8.880e+06        1        1   table cell
      3.3       1      TRUE       3             char        1        1   table cell
      3.12      3     FALSE       3          apricot        1        1   table cell
      3.20      4     FALSE       3           banana        1        1   table cell
      3.28      5     FALSE       3          coconut        1        1   table cell
      3.36      6     FALSE       3           durian        1        1   table cell
      3.45      8     FALSE       3               NA        1        1   table cell
      3.53      9     FALSE       3              fig        1        1   table cell
      3.61     10     FALSE       3       grapefruit        1        1   table cell
      3.69     11     FALSE       3         honeydew        1        1   table cell
      4.4       1      TRUE       4             fctr        1        1   table cell
      4.13      3     FALSE       4              one        1        1   table cell
      4.21      4     FALSE       4              two        1        1   table cell
      4.29      5     FALSE       4            three        1        1   table cell
      4.37      6     FALSE       4             four        1        1   table cell
      4.46      8     FALSE       4             five        1        1   table cell
      4.54      9     FALSE       4              six        1        1   table cell
      4.62     10     FALSE       4            seven        1        1   table cell
      4.70     11     FALSE       4            eight        1        1   table cell
      5.5       1      TRUE       5             date        1        1   table cell
      5.14      3     FALSE       5       2015-01-15        1        1   table cell
      5.22      4     FALSE       5       2015-02-15        1        1   table cell
      5.30      5     FALSE       5       2015-03-15        1        1   table cell
      5.38      6     FALSE       5       2015-04-15        1        1   table cell
      5.47      8     FALSE       5       2015-05-15        1        1   table cell
      5.55      9     FALSE       5       2015-06-15        1        1   table cell
      5.63     10     FALSE       5               NA        1        1   table cell
      5.71     11     FALSE       5       2015-08-15        1        1   table cell
      6.6       1      TRUE       6             time        1        1   table cell
      6.15      3     FALSE       6            13:35        1        1   table cell
      6.23      4     FALSE       6            14:40        1        1   table cell
      6.31      5     FALSE       6            15:45        1        1   table cell
      6.39      6     FALSE       6            16:50        1        1   table cell
      6.48      8     FALSE       6            17:55        1        1   table cell
      6.56      9     FALSE       6               NA        1        1   table cell
      6.64     10     FALSE       6            19:10        1        1   table cell
      6.72     11     FALSE       6            20:20        1        1   table cell
      7.7       1      TRUE       7         datetime        1        1   table cell
      7.16      3     FALSE       7 2018-01-01 02:22        1        1   table cell
      7.24      4     FALSE       7 2018-02-02 14:33        1        1   table cell
      7.32      5     FALSE       7 2018-03-03 03:44        1        1   table cell
      7.40      6     FALSE       7 2018-04-04 15:55        1        1   table cell
      7.49      8     FALSE       7 2018-05-05 04:00        1        1   table cell
      7.57      9     FALSE       7 2018-06-06 16:11        1        1   table cell
      7.65     10     FALSE       7 2018-07-07 05:22        1        1   table cell
      7.73     11     FALSE       7               NA        1        1   table cell
      8.8       1      TRUE       8         currency        1        1   table cell
      8.17      3     FALSE       8           49.950        1        1   table cell
      8.25      4     FALSE       8           17.950        1        1   table cell
      8.33      5     FALSE       8            1.390        1        1   table cell
      8.41      6     FALSE       8        65100.000        1        1   table cell
      8.50      8     FALSE       8         1325.810        1        1   table cell
      8.58      9     FALSE       8           13.255        1        1   table cell
      8.66     10     FALSE       8               NA        1        1   table cell
      8.74     11     FALSE       8            0.440        1        1   table cell

