#' Edit table with Excel
#'
#' Open/edit a file using desktop version of excel  q q
#'
#' @param x table; object to edit
#'
#' @details
#'
#' Works on Windows only.
#'
#' System has to be set up so that Excel reads CSV files. This is generally, the
#' default.
#'
#' @references
#'   \url{http://stackoverflow.com/a/12164899/1504321}
#'
#' @seealso
#'   edit, edit.data.frame
#'
#' @examples
#'
#' \dontrun{
#'   df <- edit_excel(iris)
#' }
#'
#' @import readr
#' @export

edit_excel <- function (x) {

  tempFilePath = paste(tempfile(), ".csv")
  tempPath = dirname(tempFilePath)
  preferredFile = paste(deparse(substitute(x)), ".csv", sep = "")
  preferredFilePath = file.path(tempPath, preferredFile)

  if(length(dim(x))>2){
    stop('Too many dimensions')
  }
  if(is.null(dim(x))){
    x = as.data.frame(x)
  }

  # if (is.null(rownames(x))) {
  #   tmp = 1:nrow(x)
  # }else {
  #   tmp = rownames(x)
  # }
  # rownames(x) = NULL
  # x = data.frame(RowLabels = tmp, x)

  WriteAttempt = try(
    write.table(x, file=preferredFilePath, quote=TRUE, sep=",", na="",
                row.names=FALSE, qmethod="double"),
    silent = TRUE)

  if ("try-error" %in% class(WriteAttempt)) {

    write.table(x, file=tempFilePath, quote=TRUE, sep=",", na=""
                , row.names=FALSE, qmethod="double" )
    shell.exec(tempFilePath)

  } else {
    shell.exec(preferredFilePath)
    file_wait(preferredFilePath)
    ret <- read_csv( preferredFilePath )
  }

  # This is useful for
  rbind(x[0,], ret)
}


file_wait <- function (file, time = 0.2)
{
    mtime = file.mtime(file)
    while (mtime == file %>% file.mtime()) Sys.sleep(time)
}
