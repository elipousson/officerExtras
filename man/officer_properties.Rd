% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/read_officer.R
\name{officer_properties}
\alias{officer_properties}
\title{Get doc properties for a rdocx or rpptx object as a list}
\usage{
officer_properties(x, values = list(), keep.null = FALSE, call = caller_env())
}
\arguments{
\item{x}{A rdocx or rpptx object.}

\item{values}{A named list with new properties to replace existing document
properties before they are returned as a named list.}

\item{keep.null}{Passed to \code{\link[utils:modifyList]{utils::modifyList()}}. If \code{TRUE}, retain
properties in returned list even if they have \code{NULL} values.}

\item{call}{The execution environment of a currently
running function, e.g. \code{caller_env()}. The function will be
mentioned in error messages as the source of the error. See the
\code{call} argument of \code{\link[rlang:abort]{abort()}} for more information.}
}
\value{
A named list of existing document properties or (if values is
supplied) modified document properties.
}
\description{
\code{\link[=officer_properties]{officer_properties()}} is a variant on \code{\link[officer:doc_properties]{officer::doc_properties()}} that will
warn instead of error if document properties can't be found
}
