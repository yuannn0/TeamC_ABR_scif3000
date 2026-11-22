# Full counting: each project counted once per mentioned group, at its FULL award.
# Categories overlap; their sums are NOT the total unique funding.
input_file <- "2020_2025_NHMRC.xlsx"
years <- 2020:2025
include_unclassified <- TRUE
bar_colour <- "#176B87"
unclassified_colour <- "#929292"
text_colour <- "#202020"
main_title <- "Cumulative NHMRC Funding by Research Area, 2020\u20132025"
x_axis_title <- "Cumulative project award totals (AUD millions)"
font_family <- "sans"
label_size <- 0.88
value_size <- 0.95
title_size <- 1.65
label_wrap_width <- 51
value_digits <- 1
figure_width <- 15                 # inches
figure_height <- 11                # inches
png_dpi <- 600
output_folder <- "."  # created beside this script

label_overrides <- c() # Optional named vector of display-label replacements

# Locate this script under RStudio Source, source(), or Rscript.
script_file <- NULL
for (frame in sys.frames()) {
  if (!is.null(frame$ofile)) script_file <- frame$ofile
}
if (is.null(script_file)) {
  script_arg <- grep("^--file=", commandArgs(), value = TRUE)
  if (length(script_arg)) script_file <- sub("^--file=", "", script_arg[1])
}
script_dir <- if (!is.null(script_file)) dirname(normalizePath(script_file, mustWork = TRUE)) else getwd()


if (!requireNamespace("readxl", quietly = TRUE)) {
  stop('Install the Excel reader once with: install.packages("readxl") ; then Source this script again.')
}
input <- if (file.exists(file.path(script_dir, input_file))) file.path(script_dir, input_file) else input_file
if (!file.exists(input)) stop("Excel file not found. Put it beside this R script or set input_file to its full path: ", input_file)
input <- normalizePath(input, winslash = "/")
norm <- function(x) tolower(trimws(gsub("[[:space:]]+", " ", x)))
tokens <- function(x) {
  v <- trimws(strsplit(x, "|", fixed = TRUE)[[1]])
  v[nzchar(v)]
}
# This workbook is ONE flat table, unlike the year-tab reference workbook.
d <- as.data.frame(readxl::read_excel(input, sheet = 1,
     col_types = c("text", "date", "numeric", "text")))
stopifnot(identical(names(d), c("APP ID", "Date Announced", "Total", "Field of Research")))
d$source_excel_row <- seq_len(nrow(d)) + 1L
d$year <- as.integer(format(d[["Date Announced"]], "%Y"))
if (anyNA(d$year)) stop("Missing or invalid announcement dates")
d <- d[d$year %in% years, , drop = FALSE]
if (!all(years %in% d$year)) stop("Missing requested years")
if (anyNA(d$Total) || any(!is.finite(d$Total) | d$Total < 0)) stop("Invalid award amounts")
# Official division names present in this workbook identify 3-level paths.
# Matching the leading division prevents mistaking 3 flat leaf fields for a path.
divisions <- c("biological sciences", "biomedical and clinical sciences", "health sciences",
 "indigenous studies", "psychology", "engineering", "chemical sciences", "physical sciences",
 "mathematical sciences", "information and computing sciences", "environmental sciences",
 "agricultural, veterinary and food sciences", "built environment and design", "education",
 "economics", "commerce, management, tourism and services", "human society", "law and legal studies",
 "creative arts and writing", "language, communication and culture", "history, heritage and archaeology",
 "philosophy and religious studies", "earth sciences")
parse_fields <- function(x, year) {
 if (is.na(x) || !nzchar(trimws(x))) return(data.frame(parsed_path=NA_character_, field="unclassified: missing"))
 # Handle real CR/LF and literal escaped line separators.
 for (z in c("\\r", "\\n", "\r")) x <- gsub(z, "\n", x, fixed=TRUE)
 chunks <- strsplit(x, "\n+")[[1]]
 paths <- character()
 for (chunk in chunks) {
   t <- tokens(norm(chunk))
   if (!length(t)) next
   three_level <- length(t) %% 3L == 0L && all(t[seq(1,length(t),3)] %in% divisions)
   width <- if (three_level && year >= 2022L) 3L else if (year %in% c(2021L,2022L)) 2L else 1L
   if (length(t)==1L) width <- 1L
   if (length(t) %% width != 0L) stop("Unrecognised field path: ", chunk)
   # Every 3-level path must begin with a division. Never silently use a wrong width.
   if (width==3L && any(!t[seq(1,length(t),3)] %in% divisions)) stop("Mixed path formats: ",chunk)
   paths <- c(paths, vapply(split(t,rep(seq_len(length(t)/width),each=width)),
                           paste, character(1), collapse=" | "))
 }
 if (!length(paths)) stop("Empty field after parsing")
 fields <- vapply(paths,function(p) tail(tokens(p),1),character(1))
 fields[fields=="not applicable"] <- "unclassified: not applicable"
 data.frame(parsed_path=paths,field=fields)
}
neuroscience_extra_fields <- c("neurology and neuromuscular diseases", "neurogenetics",
 "sensory systems", "cell neurochemistry", "neurocognitive patterns and neural networks",
 "animal neurobiology", "motor control")
group_research_fields <- function(x) {
 result <- x
 result[grepl("immu",x)] <- "immunology"
 result[grepl("nervous system|neuroscien",x) | x %in% neuroscience_extra_fields] <- "neuroscience"
 result
}
display_fields <- c("neuroscience", "immunology", "unclassified: not applicable",
 "cardiology (incl. cardiovascular diseases)", "infectious diseases",
 "cancer therapy (excl. chemotherapy and radiation therapy)",
 "public health and health services not elsewhere classified", "health promotion",
 "respiratory diseases", "public health not elsewhere classified",
 "psychiatry (incl. psychotherapy)", "preventative health care", "genomics",
 "cell development, proliferation and death", "aboriginal and torres strait islander health")
allocation_list <- mapping_list <- vector("list",nrow(d))
for(i in seq_len(nrow(d))) {
 p <- parse_fields(d[["Field of Research"]][i],d$year[i])
 p$grouped_field <- group_research_fields(p$field)
 p$record_id <- paste0("Sheet1:",d$source_excel_row[i])
 p$app_id <- d[["APP ID"]][i]
 p$year <- d$year[i]
 p$original_text <- d[["Field of Research"]][i]
 mapping_list[[i]] <- p
 f <- unique(p$grouped_field)
 # Match "immu" filter on the entire original cell.
 if (grepl("immu",d[["Field of Research"]][i],ignore.case=TRUE)) f <- union(f,"immunology")
 allocation_list[[i]] <- data.frame(record_id=p$record_id[1],app_id=p$app_id[1],year=d$year[i],
   field=f,award_aud=d$Total[i],counted_aud=d$Total[i],project_count=1L,
   original_text=d[["Field of Research"]][i])
}
allocations <- do.call(rbind,allocation_list)
field_mapping <- do.call(rbind,mapping_list)
funding_data <- do.call(rbind,lapply(split(allocations,allocations$field),function(z)
 data.frame(field=z$field[1],cumulative_known_million_aud=sum(z$counted_aud)/1e6,
            full_award_total_aud=sum(z$counted_aud),involved_records=nrow(z))))
funding_data <- funding_data[order(-funding_data$cumulative_known_million_aud,funding_data$field),]
annual <- aggregate(cbind(counted_aud,project_count)~year+field,allocations,sum)
# Validate full amounts, one row per project/group, and independent raw filter.
stopifnot(!anyDuplicated(allocations[c("record_id","field")]),
 all(allocations$counted_aud==allocations$award_aud))
raw_immu <- grepl("immu",d[["Field of Research"]],ignore.case=TRUE)
stopifnot(abs(sum(d$Total[raw_immu])-sum(allocations$counted_aud[allocations$field=="immunology"]))<0.01)
for(f in unique(allocations$field)) {
 ids <- allocations$record_id[allocations$field==f]
 expected <- sum(d$Total[paste0("Sheet1:",d$source_excel_row) %in% ids])
 stopifnot(abs(expected-sum(allocations$counted_aud[allocations$field==f]))<0.01)
}
output_dir <- file.path(script_dir,output_folder)
audit_dir <- file.path(output_dir,"audit")
dir.create(audit_dir,recursive=TRUE,showWarnings=FALSE)
export <- function(x,name) write.csv(x,file.path(audit_dir,paste0(name,".csv")),
 row.names=FALSE,na="NA",fileEncoding="UTF-8")
export(funding_data,"all_fields_summary")
export(annual,"annual_summary")
export(field_mapping,"field_mapping")
export(unique(field_mapping[c("field","grouped_field")]),"field_grouping")
export(allocations,"project_group_full_amounts")
summary_table <- funding_data[funding_data$field %in% display_fields,]
write.csv(summary_table,file.path(output_dir,"NHMRC_full_count_summary.csv"),row.names=FALSE)
# CSV tables always generated.
if(requireNamespace("writexl",quietly=TRUE)) writexl::write_xlsx(
 list(Summary=summary_table,Annual=annual,Project_groups=allocations),
 file.path(output_dir,"NHMRC_full_count_summary.xlsx"))
message("Validated full counting for ",nrow(d)," projects; immunology AUD ",sum(d$Total[raw_immu]))

needed_columns <- c("field", "cumulative_known_million_aud")
if (!all(needed_columns %in% names(funding_data))) {
  stop("Summary needs columns: field, cumulative_known_million_aud")
}
if (!is.numeric(funding_data$cumulative_known_million_aud) ||
    anyNA(funding_data$cumulative_known_million_aud) ||
    any(!is.finite(funding_data$cumulative_known_million_aud)) ||
    any(funding_data$cumulative_known_million_aud < 0)) {
  stop("Cumulative amounts must be finite non-negative numbers, without missing values.")
}
if (anyDuplicated(funding_data$field)) stop("Duplicate field names in summary data; review before plotting.")
if (!include_unclassified) {
  funding_data <- funding_data[!grepl("^(unclassified:|unresolved:)", funding_data$field), , drop = FALSE]
}
funding_data <- funding_data[order(-funding_data$cumulative_known_million_aud, funding_data$field), , drop = FALSE]
shown <- funding_data[funding_data$field %in% display_fields, , drop = FALSE]
if (nrow(shown) != length(display_fields)) stop("A requested display category is absent; review classification.")
if (!nrow(shown)) stop("No fields remain after filtering.")
shown$display_label <- shown$field
for (field_name in names(label_overrides)) {
  shown$display_label[shown$field == field_name] <- label_overrides[[field_name]]
}
shown$display_label <- paste0(seq_len(nrow(shown)), ". ", shown$display_label)
shown$display_label <- vapply(shown$display_label, function(x) paste(strwrap(x, width = label_wrap_width), collapse = "\n"), character(1))
shown$colour <- ifelse(grepl("^(unclassified:|unresolved:)", shown$field), unclassified_colour, bar_colour)

notes <- 'Grey bar: research field recorded as "Not Applicable"'

draw_nhmrc_bars <- function() {
  old_par <- par(no.readonly = TRUE)
  on.exit(par(old_par), add = TRUE)
  # Margins in inches keep long wrapped names outside the bars.
  view_scale <- min(1, grDevices::dev.size("in")[1] / figure_width,
                    grDevices::dev.size("in")[2] / figure_height)
  par(mai = c(1.60, 4.55, 1.0, 0.65) * view_scale, cex = view_scale, family = font_family,
      fg = text_colour, col.axis = text_colour, col.lab = text_colour,
      xaxs = "i", yaxs = "i", las = 1, bg = "white")
  amounts <- rev(shown$cumulative_known_million_aud)
  n <- length(amounts)
  xmax <- max(c(amounts, 1)) * 1.14
  plot.new()
  plot.window(xlim = c(0, xmax), ylim = c(0.35, n + 0.65))
  ticks <- pretty(c(0, max(c(amounts, 1))), n = 5)
  ticks <- ticks[ticks >= 0 & ticks <= xmax]
  abline(v = ticks, col = "#E7E7E7", lwd = 0.8)
  rect(0, seq_len(n) - 0.33, amounts, seq_len(n) + 0.33,
       col = rev(shown$colour), border = NA)
  axis(1, at = ticks, labels = format(ticks, trim = TRUE, scientific = FALSE),
       lwd = 0, lwd.ticks = 0, cex.axis = 0.9, padj = 0.4)
  text(rep(0, n), seq_len(n), labels = rev(shown$display_label),
       pos = 2, offset = 0.6, cex = label_size, xpd = NA)
  text(amounts + xmax * 0.008, seq_len(n),
       labels = formatC(amounts, format = "f", digits = value_digits),
       adj = 0, cex = value_size)
  mtext(x_axis_title, side = 1, line = 2.3, font = 2, cex = view_scale)
  text(grconvertX(0.5, from = "ndc", to = "user"),
       grconvertY(0.96, from = "ndc", to = "user"),
       labels = main_title, adj = c(0.5, 0.5), font = 2,
       cex = title_size, xpd = NA)
  mtext(subtitle, side = 3, line = 0.85, adj = 0, cex = 0.95 * view_scale)
  for (i in seq_along(notes)) {
    mtext(notes[i], side = 1, line = 4.3 + (i - 1) * 0.85, adj = 0, cex = 0.85 * view_scale)
  }
  invisible(shown)
}

output_dir <- file.path(script_dir, output_folder)
dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)
save_nhmrc_chart <- function(filename, format) {
  if (format == "png") {
    grDevices::png(filename, width = figure_width, height = figure_height,
                  units = "in", res = png_dpi, bg = "white")
  } else {
    grDevices::pdf(filename, width = figure_width, height = figure_height,
                  family = font_family, encoding = "WinAnsi.enc", useDingbats = FALSE, bg = "white")
  }
  on.exit(grDevices::dev.off(), add = TRUE)
  draw_nhmrc_bars()
}
save_nhmrc_chart(file.path(output_dir, "nhmrc_cumulative_bars.png"), "png")
save_nhmrc_chart(file.path(output_dir, "nhmrc_cumulative_bars.pdf"), "pdf")
write.csv(shown[c("field", "cumulative_known_million_aud")],
          file.path(output_dir, "displayed_funding.csv"), row.names = FALSE, fileEncoding = "UTF-8")

if (interactive()) {
  # RStudio draws into its Plots pane. Exported files keep the configured size.
  draw_nhmrc_bars()
}
message("Done. Files saved in: ", normalizePath(output_dir, winslash = "/"))
message("Change settings at the top, then click Source again to update the chart.")
