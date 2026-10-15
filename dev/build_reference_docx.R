# Build the example Word reference document from inst/extdata/_brand.yml
#
# Usage: Rscript dev/build_reference_docx.R

pkgload::load_all()

brand_reference_docx(
  brand = here::here("inst", "extdata", "_brand.yml"),
  path = here::here("inst", "extdata", "reference.docx")
)
