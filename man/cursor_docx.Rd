% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/cursor_docx.R
\name{cursor_docx}
\alias{cursor_docx}
\title{Set cursor position in rdocx object based on keyword, id, or index}
\usage{
cursor_docx(
  docx,
  keyword = NULL,
  id = NULL,
  index = NULL,
  default = "end",
  quiet = FALSE,
  call = caller_env()
)
}
\arguments{
\item{docx}{A rdocx object.}

\item{keyword, id}{A keyword string used to place cursor with
\code{\link[officer:cursor]{officer::cursor_reach()}} or bookmark id with \code{\link[officer:cursor]{officer::cursor_bookmark()}}.
Defaults to \code{NULL}. If keyword or id are not provided, the gt object is
inserted at the front of the document.}

\item{index}{A integer matching a doc_index value appearing in a summary of
the docx object created with \code{\link[officer:docx_summary]{officer::docx_summary()}}. If index is for a
paragraph value, the text of the pargraph is used as a keyword.}

\item{default}{Character string with one of the following options: \code{c("end", "begin", "backward", "forward")} to set cursor position. Only used if
keyword, id, and index are all \code{NULL}.}

\item{quiet}{If \code{FALSE} (default) warn when keyword is not found.}

\item{call}{The execution environment of a currently
running function, e.g. \code{caller_env()}. The function will be
mentioned in error messages as the source of the error. See the
\code{call} argument of \code{\link[rlang:abort]{abort()}} for more information.}
}
\description{
A combined function for setting cursor position with
\code{\link[officer:cursor]{officer::cursor_reach()}}, \code{\link[officer:cursor]{officer::cursor_bookmark()}}, or using a doc_index
value from \code{\link[officer:docx_summary]{officer::docx_summary()}}. Defaults to using
\code{\link[officer:cursor]{officer::cursor_end()}}, \code{\link[officer:cursor]{officer::cursor_begin()}},
\code{\link[officer:cursor]{officer::cursor_backward()}}, or \code{\link[officer:cursor]{officer::cursor_forward()}} if keyword, id,
and index are all \code{NULL}.
}
\seealso{
\code{\link[officer:cursor]{officer::cursor_begin()}}, \code{\link[officer:docx_summary]{officer::docx_summary()}}
}
