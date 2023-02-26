#' Summarize a rdocx or rpptx object
#'
#' @param x A rdocx or rpptx object passed to [officer::docx_summary()] or [officer::pptx_summary()].
#' @returns A data.frame object.
#' @export
#' @importFrom officer docx_summary pptx_summary
officer_summary <- function(x) {
  check_officer(x, what = c("rdocx", "rpptx"))
  switch(class(x),
    "rdocx" = officer::docx_summary(x),
    "rpptx" = officer::pptx_summary(x)
  )
}
