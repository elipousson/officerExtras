test_that("dims_docx_ext works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  dims <- dims_docx_ext(docx)
  base_dims <- officer::docx_dim(docx)

  expect_named(
    dims,
    c("page", "landscape", "margins", "orientation", "body")
  )

  # page and margins are passed through from officer::docx_dim() unchanged
  expect_equal(dims[["page"]][c("width", "height")], base_dims[["page"]])
  expect_equal(dims[["margins"]], base_dims[["margins"]])
  expect_identical(dims[["landscape"]], base_dims[["landscape"]])

  # derived values are internally consistent
  expect_equal(
    dims[["page"]][["asp"]],
    dims[["page"]][["width"]] / dims[["page"]][["height"]]
  )

  expect_identical(
    dims[["orientation"]],
    if (isTRUE(dims[["landscape"]])) "landscape" else "portrait"
  )

  margin_w <- dims[["margins"]][["left"]] + dims[["margins"]][["right"]]
  margin_h <- dims[["margins"]][["top"]] + dims[["margins"]][["bottom"]]

  expect_equal(
    dims[["body"]][["width"]],
    dims[["page"]][["width"]] - margin_w
  )
  expect_equal(
    dims[["body"]][["height"]],
    dims[["page"]][["height"]] - margin_h
  )
  expect_equal(
    dims[["body"]][["asp"]],
    dims[["body"]][["width"]] / dims[["body"]][["height"]]
  )
})
