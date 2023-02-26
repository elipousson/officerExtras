
<!-- README.md is generated from README.Rmd. Please edit that file -->

# officerExtras

<!-- badges: start -->

[![Lifecycle:
experimental](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://lifecycle.r-lib.org/articles/stages.html#experimental)
[![License:
MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Project Status: WIP – Initial development is in progress, but there
has not yet been a stable, usable release suitable for the
public.](https://www.repostatus.org/badges/latest/wip.svg)](https://www.repostatus.org/#wip)
[![Codecov test
coverage](https://codecov.io/gh/elipousson/officerExtras/branch/main/graph/badge.svg)](https://app.codecov.io/gh/elipousson/officerExtras?branch=main)
<!-- badges: end -->

The goal of officerExtras is to provide helper and convenience functions
for the officer package.

## Installation

You can install the development version of officerExtras like so:

``` r
pak::pkg_install("elipousson/officerExtras")
```

## Example

This is a basic example which shows you how to solve a common problem:

``` r
library(officerExtras)

docx <- read_docx_ext(filename = "example.docx", path = system.file("doc_examples", package = "officer"))

check_docx(docx)
```