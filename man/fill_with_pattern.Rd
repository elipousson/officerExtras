% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/utils.R
\name{fill_with_pattern}
\alias{fill_with_pattern}
\title{Fill a new column based on an existing column depending on a pattern}
\usage{
fill_with_pattern(
  x,
  pattern = "^heading",
  pattern_col = "style_name",
  fill_col = "text",
  col = "heading",
  direction = c("down", "up", "downup", "updown"),
  call = caller_env()
)
}
\arguments{
\item{x}{A input data.frame (assumed to be from \code{\link[=officer_summary]{officer_summary()}} for
default values).}

\item{pattern}{Passed to \code{\link[=grepl]{grepl()}} as the pattern identifying which rows of
the fill_col should have values pulled into the new column named by col.
Defaults to "^heading" which matches the default heading style names.}

\item{pattern_col}{Name of column to use for pattern matching, Defaults to
"style_name".}

\item{fill_col}{Name of column to fill , Defaults to "text".}

\item{col}{Name of new column to fill with values from fill_col, Default:
Defaults to "heading"}

\item{direction}{Direction of fill passed to \code{\link[vctrs:vec_fill_missing]{vctrs::vec_fill_missing()}},
Default: c("down", "up", "downup", "updown")}

\item{call}{The execution environment of a currently
running function, e.g. \code{caller_env()}. The function will be
mentioned in error messages as the source of the error. See the
\code{call} argument of \code{\link[rlang:abort]{abort()}} for more information.}
}
\value{
A data.frame with an additional column taking the name from col and
the values from the column named in fill_col.
}
\description{
This function uses \code{\link[vctrs:vec_fill_missing]{vctrs::vec_fill_missing()}} to convert hierarchically
nested headings and text into a rectangular data.frame. It is an experimental
function that may be modified or removed. At present, it is only used by
\code{\link[=officer_tables]{officer_tables()}}.
}
\seealso{
\code{\link[vctrs:vec_fill_missing]{vctrs::vec_fill_missing()}}
}
