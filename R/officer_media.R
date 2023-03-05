#' @keywords internal
#' @export
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

  cli::cli_alert_info("{n_images} media files found in {.path {path}}")
  filenames
}

#' @keywords internal
#' @export
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

#' Copy media from a docx or pptx file to a folder
#'
#' @inheritParams set_office_path
#' @param dest Folder name or path to copy media files to. dest is created if no
#'   folder exists at that location.
#' @param overwrite If `TRUE` (default), overwrite any files with the same names
#'   at the dest location.
#' @export
#' @importFrom cli cli_alert_success
#' @importFrom rlang set_names
officer_media <- function(filename = NULL,
                          path = NULL,
                          dest = "media",
                          overwrite = TRUE) {
  path <- set_office_path(filename, path, fileext = c("docx", "pptx"))
  media_files <- list_officer_media(path)

  if (length(media_files) == 0) {
    return(invisible(NULL))
  }

  exdir <- tempdir()

  unzip_officer(path, exdir)

  if (!dir.exists(dest)) {
    dir.create(dest)
  }

  file.copy(
    file.path(exdir, media_files),
    file.path(dest, basename(media_files)),
    overwrite = overwrite,
    copy.date = TRUE
  )

  unlink(exdir)

  cli::cli_alert_success("Copied media to {.path {dest}}")
}
