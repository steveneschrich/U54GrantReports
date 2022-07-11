#' Update package from github
#'
#' @return Result of [[devtools::install_github()]]
#' @export
#'
#' @examples
#' \dontrun{
#' update_packages()
#' }
update_packages <- function() {
  devtools::install_github("steveneschrich/U54Reports", force = TRUE)
}
