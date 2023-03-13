#' @keywords internal
#' @noRd
list_officer_media <- function(path, type = "docx") {
  filenames <- NULL
  if (has_fileext(path)) {
    filenames <- unzip_officer(path, list = TRUE)[["Name"]]
    path <- basename(path)
  }

  list_media_files(path, filenames, type)
}

#' @noRd
#' @keywords internal
#' @importFrom cli cli_alert_warning cli_alert_info
list_media_files <- function(path = NULL,
                             filenames = NULL,
                             type = "docx") {
  pattern <- switch(type,
    "docx" = "word",
    "pptx" = "ppt"
  )

  filenames <- filenames %||% list.files(path, recursive = TRUE)
  pattern <- file.path(pattern, "media", ".+")
  filenames <- filenames[str_detect(filenames, pattern)]
  n_media_files <- length(filenames)

  if (n_media_files == 0) {
    cli::cli_alert_warning("No media files found in {.path {path}}")
    return(NULL)
  }

  cli::cli_alert_info("{n_media_files} media file{?s} found in {.path {path}}")
  filenames
}

#' @keywords internal
#' @noRd
#' @importFrom utils unzip
unzip_officer <- function(path,
                          exdir = ".",
                          list = FALSE,
                          overwrite = TRUE,
                          setTimes = TRUE) {
  utils::unzip(
    zipfile = path,
    exdir = exdir,
    list = list,
    overwrite = overwrite,
    setTimes = setTimes
  )
}

#' Copy media from a docx or pptx file or a rdocx or rpptx object to a target
#' folder
#'
#' Unzip a docx or pptx file to a temporary directory, check if the directory
#' contains a media folder, and copy media files to the directory set by dir. If
#' a rdocx or rpptx object is provided, files are copied from the temporary
#' package_dir associated with the object (accessible via `x[["package_dir"]]`).
#'
#' @param filename,path File name and path for a `docx` or `pptx` file. One of
#'   the two must be provided. Ignored if x is provided. Defaults to `NULL`.
#' @param x A rdocx or rpptx object containing one or more media file.
#' @param target Folder name or path to copy media files. dir is created if no
#'   folder exists at that location.
#' @param list If `TRUE`, display a message listing files contained in the docx
#'   or pptx file but do not copy the files to dir. Defaults to `FALSE`.
#' @param overwrite If `TRUE` (default), overwrite any files with the same names
#'   at target path. If `FALSE`, abort if files with the same names already
#'   exist.
#' @examples
#' officer_media(
#'   system.file("doc_examples/example.pptx", package = "officer"),
#'   list = TRUE
#' )
#' @seealso [officer::media_extract()]
#' @export
#' @importFrom cli cli_bullets cli_alert_success
#' @importFrom rlang set_names
officer_media <- function(filename = NULL,
                          path = NULL,
                          x = NULL,
                          target = "media",
                          list = FALSE,
                          overwrite = TRUE) {
  if (is.null(x)) {
    path <- set_office_path(filename, path, fileext = c("docx", "pptx"))
    type <- str_extract_fileext(path)
  } else {
    check_officer(x)
    path <- x[["package_dir"]]
    type <- str_remove(class(x), "^r")
  }

  media_files <- list_officer_media(path, type = type)

  if (length(media_files) == 0) {
    return(invisible(NULL))
  }

  if (isTRUE(list)) {
    cli::cli_bullets(
      rlang::set_names(
        basename(media_files),
        rep("*", length(media_files))
      )
    )
    return(invisible(media_files))
  }

  exdir <- path

  if (is.null(x)) {
    exdir <- tempdir()
    unzip_officer(path, exdir)
  }

  if (!dir.exists(target)) {
    dir.create(target)
  }

  media_files_from <- file.path(exdir, media_files)
  media_files_to <- file.path(target, basename(media_files))

  if (any(file.exists(media_files_to)) && isFALSE(overwrite)) {
    cli::cli_abort(
      c("One or more media files already exist at {.arg target}: {.path {target}}",
        "i" = "Set {.code overwrite = TRUE} to replace these existing files.")
    )
  }

  file.copy(
    media_files_from,
    media_files_to,
    overwrite = overwrite,
    copy.date = TRUE
  )

  unlink(exdir)

  cli::cli_alert_success("Copied media to {.path {target}}")
}
