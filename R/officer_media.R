#' @keywords internal
#' @noRd
#' @importFrom cli cli_alert_warning cli_alert_info
list_officer_media <- function(path) {
  filenames <- unzip_officer(path, list = TRUE)[["Name"]]

  pattern <- switch(str_extract_fileext(path),
    "docx" = "word",
    "pptx" = "ppt"
  )

  path <- basename(path)

  pattern <- file.path(pattern, "media", ".+")
  filenames <- filenames[str_detect(filenames, pattern)]

  n_images <- length(filenames)

  if (n_images == 0) {
    cli::cli_alert_warning("No media files found in {.path {path}}")
    return(NULL)
  }

  cli::cli_alert_info("{n_images} media file{?s} found in {.path {path}}")
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

#' Copy media from a docx or pptx file to a target folder
#'
#' Unzip a docx or pptx file to a temporary directory, check if the directory
#' contains a media folder, and copy media files to the directory set by dir.
#'
#' @param filename,path File name and path for a `docx` or `pptx` file. One of
#'   the two must be provided. Defaults to `NULL`.
#' @param target Folder name or path to copy media files. dir is created if no
#'   folder exists at that location.
#' @param list If `TRUE`, display a message listing files contained in the docx
#'   or pptx file but do not copy the files to dir. Defaults to `FALSE`.
#' @param overwrite If `TRUE` (default), overwrite any files with the same names
#'   at the dir location.
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
                          target = "media",
                          list = FALSE,
                          overwrite = TRUE) {
  path <- set_office_path(filename, path, fileext = c("docx", "pptx"))
  mediafiles <- list_officer_media(path)

  if (length(mediafiles) == 0) {
    return(invisible(NULL))
  }

  if (isTRUE(list)) {
    cli::cli_bullets(
      rlang::set_names(
        basename(mediafiles),
        rep("*", length(mediafiles))
      )
    )
    return(invisible(mediafiles))
  }

  exdir <- tempdir()

  unzip_officer(path, exdir)

  if (!dir.exists(target)) {
    dir.create(target)
  }

  file.copy(
    file.path(exdir, mediafiles),
    file.path(target, basename(mediafiles)),
    overwrite = overwrite,
    copy.date = TRUE
  )

  unlink(exdir)

  cli::cli_alert_success("Copied media to {.path {target}}")
}
