#' Get page and block layout dimensions from an rdocx object
#'
#' This function extends [officer::docx_dim()] by also returning the body text
#' dimensions within the margins, the aspect ratio of the page and body, and the
#' page orientation as a string ("landscape" or "portrait").
#'
#' @param docx A rdocx object to get dimensions for.
#' @seealso [officer::docx_dim()]
#' @export
#' @importFrom officer docx_dim
dims_docx_ext <- function(docx) {
  dims <- officer::docx_dim(docx)
  page_dims <- dims[["page"]]

  dims[["page"]] <- c(
    page_dims,
    "asp" = page_dims[["width"]] / page_dims[["height"]]
  )

  margin_w <- dims[["margins"]][["left"]] + dims[["margins"]][["right"]]
  margin_h <- dims[["margins"]][["top"]] + dims[["margins"]][["bottom"]]

  dims <-
    c(
      dims,
      list(
        "orientation" = ifelse(
          isTRUE(dims[["landscape"]]),
          "landscape", "portrait"
        )
      )
    )

  c(
    dims,
    list(
      "body" = c(
        "width" = page_dims[["width"]] - margin_w,
        "height" = page_dims[["height"]] - margin_h,
        "asp" = (page_dims[["width"]] - margin_w) /
          (page_dims[["height"]] - margin_h)
      )
    )
  )
}
