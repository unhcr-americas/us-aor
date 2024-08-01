library(tidyverse)
library(tidyxl)
library(readxl)
library(rvest)

nms <- c("group", "sub_group", "sub_sub_group", "sub_sub_sub_group")

parse_sheet <- function(f, s) {
  excel <- read_excel(f, s, col_names = FALSE)
  cells <- xlsx_cells(f)
  fmts <- xlsx_formats(f)
  
  data <- 
    excel |> 
    mutate(id = row_number(),
           indent = head(fmts$local$alignment$indent[filter(cells, sheet == s, col == 1)$local_format_id], n()),
           fg = head(fmts$local$fill$patternFill$fgColor$rgb[filter(cells, sheet == s, col == 1)$local_format_id], n()),
           hdr = nms[1+indent],
           .before = 1) |> 
    filter(fg == "FFDCE6F1", !is.na(...1)) |> 
    pivot_wider(names_from = hdr, values_from = ...1)
  
  for (i in 1:nrow(data)) {
    if (data$indent[i] > 0) data$group[i] <- data$group[i-1]
    if (data$indent[i] > 1) data$sub_group[i] <- data$sub_group[i-1]
    if (data$indent[i] > 2) data$sub_sub_group[i] <- data$sub_sub_group[i-1]
  }
  
  data <- 
    data |> 
    select(-c(id, indent, fg)) |> 
    set_names(c(as.character(excel[first(data$id)-1,]), intersect(nms, names(data)))[-1]) |> 
    filter(!is.na(Total))
  
  if (all(names(data) %in% c("Total", nms))) {
    data <- data |> select(-Total, n = Total) |> mutate(n = parse_number(n))
  } else {
    data <- 
      data |> 
      select(-Total) |> 
      type_convert(na = "-") |> 
      pivot_longer(where(is.numeric), names_to = "month", values_to = "n") |> 
      mutate(month = as.Date(as.numeric(month), origin = "1899-12-30"))
  }
  
  data
}

aor <- function() {
  tmpf <- fs::file_temp(ext = "xlsx")
  r <- session("https://ohss.dhs.gov/topics/immigration/refugees-and-asylees/asylum-cohort")
  href <- r |> html_elements("a[href$='xlsx']") |> html_attr("href") |> first()
  r |> session_jump_to(href, httr::write_disk(tmpf))
  
  excel_sheets(tmpf) |> 
    set_names() |> 
    discard(\(x) x %in% c("TOC", "Family AMI Cases")) |> 
    map(\(s) parse_sheet(tmpf, s)) |> 
    list_rbind(names_to = "sheet") |> 
    select(sheet, !!nms, month, n) |> 
    mutate(across(where(is.character), 
                  \(x) str_replace(x, "\\d+$", "")))
}

#* @get /aor.csv
#* @serializer csv
aor

