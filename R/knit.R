#' Knit to a docx file, with package-specific config.
#'
#' @description The `U54Reports` package has customizations to the knit command,
#' embedded within this function.
#'
#' @details This code was taken from
#'  https://bookdown.org/yihui/rmarkdown-cookbook/custom-knit.html as a way
#'  to change the output file associated with markdown from within R Studio.
#'
#'  To use these customizations, include the following in the YAML header of
#'  the markdown file:
#'  ```
#'  knit: U54Reports::knit_docx
#'  ```
#'
#'  The knit_docx function is generic in that the defaults to the function operate as
#'  a normal knit would (with the exception of the alternate directory). Note that the
#'  YAML header cannot include a function with parameters as a call, only the reference
#'  to the function.
#'
#'  There are specific functions that handle various combinations of parameters. See, for
#'  instance, [[knit_docx_prompt()]].
#'
#'  The specific customizations of knitr include:
#'
#'  - Output directory is set to `here::here(delivery)` with `delivery` being the default
#'  name (can be overridden with the `delivery` parameter).
#'  - Can ask for parameters to the markdown document via shiny rather than setting them manually.
#'  - Can include `_datestamp` on the output filename when knitting.
#'
#' @param input Input file
#' @param use_output_dir Logical. Should a separate delivery dir be used for output?
#' @param output_dir Output directory (default: delivery)
#' @param use_shiny_prompt Use Shiny to ask for report parameters defined in YAML header
#' @param use_datestamp Add a date stamp to the filename (before the docx).
#' @param datestamp A date stamp to add (_YYYYMMDD)
#' @param use_docx_template Use the [[docx_template()]] function that specifies a docx template for knitting.
#' @param docx_template_file The filename to use as a docx template
#' @param ... Any other parameters
#'
#' @return Nothing, knits and processes file (see [[knitr::knit()]] for details).
#' @export
#'
#'
knit_docx <- function(
    input,
    use_output_dir = TRUE,
    output_dir = "delivery",
    use_shiny_prompt = FALSE,
    use_datestamp = TRUE,
    datestamp = lubridate::today(),
    use_docx_template = TRUE,
    docx_template_file = docx_template(),
    ...
) {
  rmarkdown::render(
    input,
    output_dir = here::here(output_dir),
    output_file = paste0(
      xfun::sans_ext(input),
      ifelse(use_datestamp,paste0('_',datestamp), ""),
      ".docx"
    ),
    params = if (use_shiny_prompt) "ask" else list(),

    output_options = (
      if (use_docx_template)
        list("reference_docx"=docx_template())
      else
        list()
    ),
    envir = globalenv()
  )
}

#' @describeIn knit_docx Knit using the Shiny interface for parameters
#' @export
knit_docx_prompt <- function(...) {
  knit_docx(use_shiny_prompt=TRUE, ...)
}


#' Return path to U54 Word template
#'
#' @description Return the path to a word template for reporting.
#'
#' @details There is a common word document template used for all of the U54 reporting
#'   via officedown. This template has the U54 banner and various settings preconfigured.
#'   Rather than duplicate the word document many times, it is stored in a default directory
#'   within the package installation.
#'
#'   This function returns the path to this file, so that knitr/rmarkdown can refer to the
#'   template indirectly.
#'
#' @seealso [knit_docx_with_datestamp()]
#'
#' @return A path to the Word document template.
#' @export
#'
#' @examples
#' docx_template()
#'
docx_template <- function() {
  system.file("resources", "u54_docx.docx", package = "U54Reports")
}

