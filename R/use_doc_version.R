# # MIT License
#
# Copyright (c) 2023 desc authors
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
#   The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#
# https://github.com/r-lib/desc/blob/main/LICENSE.md

#' Update the version of a document and optionally save a new version of the
#' input document
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' [use_doc_version()] is inspired by [usethis::use_version()] and designed to
#' support semantic versioning for Microsoft Word or PowerPoint documents. The
#' document version is tracked as a custom file property and a component of the
#' document filename. If filename is supplied and the filename contains a
#' version number, the function ignores any version number stored as a document
#' property and overwrites the property with the new incremented version number.
#'
#' [doc_version()] is a helper function that returns the document version based
#' on a supplied filename or document properties.
#'
#' The internals for this function are adapted from the
#' internal `idesc_bump_version()` function authored by Csárdi Gábor for the
#' `{desc}` package.
#'
#' @param filename Filename for a Word document or PowerPoint presentation.
#'   Optional if x is supplied, however, it is required to locally save a file
#'   with an updated version name.
#' @param x A rdocx or rpptx object. Optional if filename is supplied.
#' @param which A string specifying which level to increment, one of: "major",
#'   "minor", "patch", "dev".
#' @param save If `TRUE` (default) and filename is supplied, write a new file
#'   replacing the existing version number.
#' @param sep Character separating version components. Defaults to ".". "-" is
#'   also supported.
#' @param property Property name optionally containing a version number.
#'   Defaults to "version".
#' @param prefix Property name to use as prefix for new filename, e.g.
#'   "modified" to use modified date/time as the new filename prefix. If prefix
#'   is identical to property and the filename does not already include a
#'   version number, version is added to the filename as a prefix instead of a
#'   postfix. Defaults to `NULL`. Note handling of this parameter is expected to
#'   change in future versions of this function.
#' @param ... Additional parameters passed to [write_officer()] if save is
#'   `TRUE`.
#' @returns Invisibly return a rdocx or rpptx object with an updated version
#'   property.
#' @export
use_doc_version <- function(filename = NULL,
                            x = NULL,
                            which = NULL,
                            save = TRUE,
                            sep = ".",
                            property = "version",
                            prefix = NULL,
                            ...,
                            call = caller_env()) {
  x <- x %||% read_officer(filename)

  ver_init <- doc_version(x, filename = filename, property = property)

  ver <- get_version_components(ver_init)

  # Adapted from desc::idesc_bump_version()
  # https://github.com/r-lib/desc/blob/c8463062edf4dc5501c66d4e8f6019743fae3fa3/R/version.R
  if (is.character(which)) {
    which <- match(which, c("major", "minor", "patch", "dev"))
  }

  stopifnot(
    is.numeric(which)
  )

  # Special dev version
  inc <- if (which == 4 && length(ver) < 4) 9000 else 1

  # Missing components are equivalent to zero
  if (which > length(ver)) ver <- c(ver, rep(0, which - length(ver)))

  # Bump selected component
  ver[which] <- ver[which] + inc

  # Zero out everything after
  if (which < length(ver)) ver[(which + 1):length(ver)] <- 0

  # Keep at most three components if they are zero
  if (length(ver) > 3 && all(ver[4:length(ver)] == 0)) {
    ver <- ver[1:3]
  }

  new_ver <- paste0(ver, collapse = sep)

  cli_alert_info(
    "Incrementing document version number to {.val {new_ver}}"
  )

  check_name(property, call = call)
  x <- set_doc_properties(x, values = set_names(list(new_ver), property))

  if (is.null(filename) || !save) {
    return(invisible(x))
  }

  if (!is.null(prefix) && !identical(prefix, property)) {
    props <- officer_properties(x)

    check_name(prefix, call = call)
    if (has_name(props, prefix)) {
      prefix <- props[[prefix]]
    }

    prefix <- make.names(gsub("\\s+|[[:punct:]]", "", prefix), allow_ = TRUE)

    filename <- file.path(
      dirname(filename),
      paste0(prefix, "_", basename(filename))
    )
  }

  if (doc_has_ver(filename, sep)) {
    filename <- doc_replace_ver(filename, new_ver)
  } else {
    fileext <- string_extract(filename, "(?<=\\.)docx|pptx$(?!\\.)")
    filename <- sub(paste0("\\.", fileext, "$"), "", filename)

    if (identical(prefix, property)) {
      filename <- paste0(filename, "_", new_ver, ".", fileext)
    } else {
      filename <- file.path(
        dirname(filename),
        paste0(new_ver, "_", basename(filename), ".", fileext)
      )
    }
  }

  cli_alert_info(
    "Saving document as {.file {filename}}"
  )

  write_officer(x, path = filename, ...)
}

# Adapted from desc::get_version_components()
#' @noRd
get_version_components <- function(x, sep = ".") {
  split <- paste0("[-\\", sep, "]")
  as.numeric(strsplit(format(x), split)[[1]])
}

version_sep <- function(sep = ".") {
  paste0(
    c("(\\d+\\", "\\d+\\", "\\d+(?:-\\w+\\", "\\d+)?)"),
    collapse = sep
  )
}

#' @noRd
doc_has_ver <- function(filename, sep = ".") {
  grepl(version_sep(sep), filename)
}

#' @noRd
doc_replace_ver <- function(filename, new_ver, sep = ".") {
  sub(version_sep(sep), new_ver, filename)
}

#' @noRd
doc_str_ver <- function(filename, sep = ".") {
  string_extract(filename, pattern = version_sep(sep))
}

#' Get document version from filename or rdocx or rpptx object
#'
#' @rdname use_doc_version
#' @name doc_version
#' @param allow_new If `TRUE` (default), return "0.1.0" if version can't be
#'   found in the filename or the properties of the input rdocx or rpptx object
#'   x.
#' @inheritParams rlang::args_error_context
#' @export
doc_version <- function(filename = NULL,
                        x = NULL,
                        sep = ".",
                        property = "version",
                        allow_new = TRUE,
                        call = caller_env()) {
  if (!is.null(filename)) {
    ver_str <- string_extract(filename, pattern = version_sep(sep))
    msg <- "Version {.val {ver_str}} found in {.file {filename}}"
  }

  if (allow_new && is.null(ver_str) && is.null(x)) {
    ver_str <- paste0(c(0, 1, 0), collapse = sep)
    msg <- "Using initial version {.val {ver_str}}"
  }

  if (is.null(ver_str)) {
    x <- x %||% read_officer(filename)
    props <- officer_properties(x, call = call)
    check_name(property, call = call)
    if (!allow_new && !has_name(props, property)) {
      cli_abort(
        c(
          "Property {.val {property}} can't be found in {.arg x}",
          "i" = "{.arg filename} must include a version string if the property
          {.val {property}} is missing and {.arg allow_new} is {.code FALSE}"
        ),
        call = call
      )
    } else if (has_name(props, property)) {
      ver_str <- props[[property]]
      msg <- "Version {.val {ver_str}} found in document properties"
    } else if (allow_new) {
      ver_str <- paste0(c(0, 1, 0), collapse = sep)
      msg <- "Using initial version {.val {ver_str}}"
    }
  }

 ver_num <- numeric_version(ver_str)
 cli_alert_info(msg)
 ver_str
}
