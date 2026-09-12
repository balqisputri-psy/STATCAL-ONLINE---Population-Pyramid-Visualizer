# ============================================================
# STATCAL ONLINE - Population Pyramid Visualizer
# R Shiny / Shinylive-friendly single-feature application
# ============================================================

# Required packages:
# install.packages(c(
#   "shiny", "shinydashboard", "DT", "readxl", "dplyr", "tidyr",
#   "ggplot2", "shinycssloaders", "scales", "openxlsx",
#   "colourpicker", "officer", "flextable"
# ))

required_packages <- c(
  "shiny", "shinydashboard", "DT", "readxl", "dplyr", "tidyr",
  "ggplot2", "shinycssloaders", "scales", "openxlsx",
  "colourpicker", "officer", "flextable"
)

missing_packages <- required_packages[!vapply(required_packages, requireNamespace, logical(1), quietly = TRUE)]
if (length(missing_packages) > 0) {
  stop(
    "Please install the following R packages first: ",
    paste(missing_packages, collapse = ", "),
    call. = FALSE
  )
}

library(shiny)
library(shinydashboard)
library(DT)
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(shinycssloaders)
library(scales)
library(openxlsx)
library(colourpicker)
library(officer)
library(flextable)

# ============================================================
# SHINYLIVE / CHROMIUM DOWNLOAD WORKAROUND
# ============================================================
# Chromium-based browsers can conflict with the HTML5 download attribute
# when the app is executed through Shinylive/webR. Removing the attribute
# lets Shinylive intercept shiny::downloadButton() correctly.

downloadButton <- function(...) {
  tag <- shiny::downloadButton(...)
  tag$attribs$download <- NULL
  tag
}

# ============================================================
# CONSTANTS
# ============================================================

APP_NAME <- "STATCAL ONLINE"
APP_TITLE <- "Population Pyramid Visualizer"
APP_UPDATED <- "Last updated on September 12, 2026"
WEBSITE_URL <- "https://statcal.com/"
STATCAL_ONLINE_URL <- "https://statcal.com/statcal%20online.html"

# Put the sample Excel file beside app.R when deploying.
# The second name matches the sample file generated during development.
SAMPLE_DATA_CANDIDATES <- c(
  "data_population_pyramid.xlsx",
  "statcal_population_pyramid_sample_200.xlsx"
)

THEMES <- list(
  "White Publication" = list(
    figure_facecolor = "#FFFFFF", axes_facecolor = "#FFFFFF",
    text_color = "#111111", grid_color = "#E5E5E5",
    spine_color = "#222222", strip_fill = "#F2F2F2"
  ),
  "Minimal Scopus Style" = list(
    figure_facecolor = "#FFFFFF", axes_facecolor = "#FFFFFF",
    text_color = "#111111", grid_color = "#ECECEC",
    spine_color = "#111111", strip_fill = "#F5F5F5"
  ),
  "Light Gray Editorial" = list(
    figure_facecolor = "#F7F7F7", axes_facecolor = "#FFFFFF",
    text_color = "#111111", grid_color = "#D8D8D8",
    spine_color = "#333333", strip_fill = "#EAEAEA"
  ),
  "Warm Ivory Journal" = list(
    figure_facecolor = "#FBF7EF", axes_facecolor = "#FFFDF8",
    text_color = "#1F1F1F", grid_color = "#DED4C5",
    spine_color = "#3A3A3A", strip_fill = "#F1E8D9"
  ),
  "Cool Blue Scientific" = list(
    figure_facecolor = "#F3F7FB", axes_facecolor = "#FFFFFF",
    text_color = "#0B1F33", grid_color = "#D2DFEA",
    spine_color = "#1F4E79", strip_fill = "#DDEBF7"
  ),
  "Dark Navy Presentation" = list(
    figure_facecolor = "#0B1320", axes_facecolor = "#111C2E",
    text_color = "#FFFFFF", grid_color = "#3B4A5F",
    spine_color = "#B8C7D9", strip_fill = "#1E2B40"
  )
)

PYRAMID_PALETTES <- list(
  "Red - Green" = c("#D95F59", "#59A96A"),
  "Blue - Orange" = c("#3B82F6", "#F59E0B"),
  "Navy - Teal" = c("#264653", "#2A9D8F"),
  "Blue - Pink" = c("#4E79A7", "#E1578A"),
  "Gray - Blue" = c("#7A7A7A", "#4E79A7"),
  "Publication Red - Blue" = c("#B2182B", "#2166AC"),
  "Manual / Custom" = NULL
)

LEGEND_CHOICES <- c("Right", "Left", "Top", "Bottom", "None / Hide legend")

# ============================================================
# HELPER FUNCTIONS
# ============================================================

clean_dataframe <- function(df) {
  df <- as.data.frame(df)
  names(df) <- trimws(gsub("\\s+", " ", as.character(names(df))))
  if (nrow(df) > 0 && ncol(df) > 0) {
    df <- df[rowSums(is.na(df)) < ncol(df), , drop = FALSE]
  }
  unnamed_cols <- grepl("^unnamed", tolower(names(df)))
  if (any(unnamed_cols)) {
    keep_unnamed <- vapply(df[unnamed_cols], function(x) !all(is.na(x)), logical(1))
    drop_names <- names(df)[unnamed_cols][!keep_unnamed]
    if (length(drop_names) > 0) {
      df <- df[, !names(df) %in% drop_names, drop = FALSE]
    }
  }
  rownames(df) <- NULL
  df
}

make_display_safe <- function(df) {
  df <- as.data.frame(df)
  for (nm in names(df)) {
    if (is.factor(df[[nm]])) df[[nm]] <- as.character(df[[nm]])
  }
  df
}

safe_number <- function(x, default_value, min_value = NULL, max_value = NULL) {
  out <- suppressWarnings(as.numeric(x))
  if (length(out) == 0 || is.na(out) || !is.finite(out)) out <- default_value
  if (!is.null(min_value)) out <- max(out, min_value)
  if (!is.null(max_value)) out <- min(out, max_value)
  out
}

safe_token <- function(x) {
  x <- gsub("[^A-Za-z0-9_]", "_", as.character(x))
  x <- gsub("_+", "_", x)
  x
}

is_hex_color <- function(x) {
  is.character(x) && length(x) == 1 && grepl("^#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$", x)
}

get_theme <- function(theme_name) {
  if (is.null(theme_name) || length(theme_name) == 0 || is.na(theme_name) || !(theme_name %in% names(THEMES))) {
    return(THEMES[["White Publication"]])
  }
  THEMES[[theme_name]]
}

safe_theme_bg <- function(theme_name) {
  tryCatch(get_theme(theme_name)$figure_facecolor, error = function(e) "white")
}

sorted_unique_values <- function(x) {
  vals <- unique(trimws(as.character(x[!is.na(x)])))
  vals <- vals[nzchar(vals)]
  vals[order(vals)]
}

first_numeric_key <- function(x) {
  x <- as.character(x)
  vapply(x, function(z) {
    m <- regexpr("[0-9]+(?:[.,][0-9]+)?", z, perl = TRUE)
    if (length(m) == 0 || m[1] < 0) return(Inf)
    txt <- regmatches(z, m)
    txt <- gsub(",", ".", txt, fixed = TRUE)
    val <- suppressWarnings(as.numeric(txt))
    ifelse(length(val) == 0 || is.na(val), Inf, val)
  }, numeric(1))
}

smart_category_order <- function(x) {
  vals <- unique(trimws(as.character(x[!is.na(x)])))
  vals <- vals[nzchar(vals)]
  if (length(vals) <= 1) return(vals)
  num_key <- first_numeric_key(vals)
  if (any(is.finite(num_key))) {
    return(vals[order(num_key, vals)])
  }
  vals[order(vals)]
}

parse_manual_order <- function(text_value, available_values) {
  available_values <- unique(as.character(available_values))
  available_values <- available_values[!is.na(available_values) & nzchar(trimws(available_values))]
  if (length(available_values) == 0) return(character(0))
  
  if (is.null(text_value) || length(text_value) == 0 || is.na(text_value) || !nzchar(trimws(as.character(text_value)))) {
    return(available_values)
  }
  
  wanted <- trimws(unlist(strsplit(as.character(text_value), ",", fixed = TRUE)))
  wanted <- wanted[nzchar(wanted)]
  ordered <- wanted[wanted %in% available_values]
  remaining <- setdiff(available_values, ordered)
  unique(c(ordered, remaining))
}

preferred_column <- function(columns, candidates, fallback = NULL) {
  if (length(columns) == 0) return(NULL)
  lower_cols <- tolower(columns)
  for (cand in candidates) {
    idx <- which(lower_cols == tolower(cand))
    if (length(idx) > 0) return(columns[idx[1]])
  }
  if (!is.null(fallback) && fallback %in% columns) return(fallback)
  columns[1]
}

preferred_group_value <- function(values, candidates, fallback_index = 1) {
  values <- unique(as.character(values))
  values <- values[!is.na(values) & nzchar(trimws(values))]
  if (length(values) == 0) return(NULL)
  lower_vals <- tolower(values)
  for (cand in candidates) {
    idx <- which(lower_vals == tolower(cand))
    if (length(idx) > 0) return(values[idx[1]])
  }
  values[min(fallback_index, length(values))]
}

find_sample_data <- function() {
  hits <- SAMPLE_DATA_CANDIDATES[file.exists(SAMPLE_DATA_CANDIDATES)]
  if (length(hits) == 0) NULL else hits[1]
}

get_pyramid_colors <- function(input, left_group, right_group) {
  palette_name <- input$chart_palette
  if (is.null(palette_name) || !(palette_name %in% names(PYRAMID_PALETTES))) {
    palette_name <- "Red - Green"
  }
  
  if (identical(palette_name, "Manual / Custom")) {
    left_col <- input$left_color
    right_col <- input$right_color
    if (!is_hex_color(left_col)) left_col <- "#D95F59"
    if (!is_hex_color(right_col)) right_col <- "#59A96A"
    cols <- c(left_col, right_col)
  } else {
    cols <- PYRAMID_PALETTES[[palette_name]]
  }
  
  names(cols) <- c(left_group, right_group)
  cols
}

statcal_pyramid_theme <- function(theme_name,
                                  title_size = 16,
                                  subtitle_size = 11,
                                  axis_title_size = 11,
                                  axis_text_size = 9,
                                  legend_title_size = 10,
                                  legend_text_size = 9,
                                  panel_title_size = 11,
                                  legend_position = "Right",
                                  show_grid = TRUE) {
  th <- get_theme(theme_name)
  legend_pos <- ifelse(
    is.null(legend_position) || legend_position == "None / Hide legend",
    "none",
    tolower(legend_position)
  )
  
  theme_minimal(base_size = axis_text_size) +
    theme(
      plot.background = element_rect(fill = th$figure_facecolor, color = NA),
      panel.background = element_rect(fill = th$axes_facecolor, color = NA),
      panel.grid.major.x = if (isTRUE(show_grid)) element_line(color = th$grid_color, linewidth = 0.35) else element_blank(),
      panel.grid.major.y = element_blank(),
      panel.grid.minor = element_blank(),
      axis.text = element_text(color = th$text_color, size = axis_text_size),
      axis.title = element_text(color = th$text_color, face = "bold", size = axis_title_size),
      plot.title = element_text(color = th$text_color, face = "bold", size = title_size),
      plot.subtitle = element_text(color = th$text_color, size = subtitle_size),
      plot.caption = element_text(color = th$text_color, size = max(6, axis_text_size - 1), hjust = 0),
      legend.text = element_text(color = th$text_color, size = legend_text_size),
      legend.title = element_text(color = th$text_color, face = "bold", size = legend_title_size),
      legend.background = element_rect(fill = th$figure_facecolor, color = NA),
      legend.position = legend_pos,
      strip.background = element_rect(fill = th$strip_fill, color = th$spine_color, linewidth = 0.25),
      strip.text = element_text(color = th$text_color, face = "bold", size = panel_title_size),
      axis.line.x = element_line(color = th$spine_color, linewidth = 0.3),
      axis.ticks.x = element_line(color = th$spine_color, linewidth = 0.3),
      plot.margin = margin(12, 18, 12, 18)
    )
}

# ============================================================
# DATA PREPARATION
# ============================================================

build_pyramid_long_data <- function(df,
                                    data_format,
                                    category_col,
                                    group_col,
                                    left_group,
                                    right_group,
                                    frequency_col = NULL,
                                    category_order = NULL,
                                    panel_col = "None",
                                    selected_panels = NULL,
                                    panel_order = NULL,
                                    percentage_denominator = "Within each group (per panel)",
                                    digits = 2) {
  
  validate(need(category_col %in% names(df), "Please select a valid category variable."))
  validate(need(group_col %in% names(df), "Please select a valid group variable."))
  validate(need(!is.null(left_group) && !is.null(right_group), "Please select left-side and right-side groups."))
  validate(need(left_group != right_group, "Left-side and right-side groups must be different."))
  
  use_panel <- !is.null(panel_col) && panel_col != "None" && panel_col %in% names(df)
  
  keep_cols <- c(category_col, group_col)
  if (use_panel) keep_cols <- c(keep_cols, panel_col)
  if (identical(data_format, "Aggregated frequency data")) {
    validate(need(!is.null(frequency_col) && frequency_col %in% names(df), "Please select a valid frequency variable."))
    keep_cols <- c(keep_cols, frequency_col)
  }
  
  tmp <- df[, unique(keep_cols), drop = FALSE]
  names(tmp)[names(tmp) == category_col] <- ".Category"
  names(tmp)[names(tmp) == group_col] <- ".Group"
  if (use_panel) names(tmp)[names(tmp) == panel_col] <- ".Panel"
  
  tmp$.Category <- trimws(as.character(tmp$.Category))
  tmp$.Group <- trimws(as.character(tmp$.Group))
  if (use_panel) {
    tmp$.Panel <- trimws(as.character(tmp$.Panel))
  } else {
    tmp$.Panel <- "All"
  }
  
  tmp <- tmp %>%
    filter(
      !is.na(.Category), nzchar(.Category),
      !is.na(.Group), .Group %in% c(left_group, right_group),
      !is.na(.Panel), nzchar(.Panel)
    )
  
  if (use_panel && !is.null(selected_panels) && length(selected_panels) > 0) {
    tmp <- tmp %>% filter(.Panel %in% selected_panels)
  }
  
  validate(need(nrow(tmp) > 0, "No observations remain after applying the selected category, groups, and panel filters."))
  
  available_categories <- unique(tmp$.Category)
  if (is.null(category_order) || length(category_order) == 0) {
    category_levels <- smart_category_order(available_categories)
  } else {
    category_levels <- parse_manual_order(paste(category_order, collapse = ","), available_categories)
  }
  
  available_panels <- unique(tmp$.Panel)
  if (use_panel) {
    if (is.null(panel_order) || length(panel_order) == 0) {
      panel_levels <- available_panels
    } else {
      panel_levels <- parse_manual_order(paste(panel_order, collapse = ","), available_panels)
    }
  } else {
    panel_levels <- "All"
  }
  
  if (identical(data_format, "Aggregated frequency data")) {
    tmp$.Frequency <- suppressWarnings(as.numeric(tmp[[frequency_col]]))
    validate(need(!all(is.na(tmp$.Frequency)), "The selected frequency variable does not contain usable numeric values."))
    validate(need(!any(tmp$.Frequency < 0, na.rm = TRUE), "Frequency values must not be negative."))
    tmp$.Frequency[is.na(tmp$.Frequency)] <- 0
    
    counted <- tmp %>%
      group_by(.Panel, .Category, .Group) %>%
      summarise(Frequency = sum(.Frequency, na.rm = TRUE), .groups = "drop")
  } else {
    counted <- tmp %>%
      count(.Panel, .Category, .Group, name = "Frequency")
  }
  
  complete_grid <- tidyr::expand_grid(
    .Panel = panel_levels,
    .Category = category_levels,
    .Group = c(left_group, right_group)
  )
  
  out <- complete_grid %>%
    left_join(counted, by = c(".Panel", ".Category", ".Group")) %>%
    mutate(Frequency = ifelse(is.na(Frequency), 0, Frequency))
  
  overall_total <- sum(out$Frequency, na.rm = TRUE)
  
  if (identical(percentage_denominator, "Total selected sample")) {
    out <- out %>%
      mutate(
        Denominator = overall_total,
        Percentage = ifelse(Denominator > 0, Frequency / Denominator * 100, NA_real_)
      )
  } else if (identical(percentage_denominator, "Within each panel")) {
    out <- out %>%
      group_by(.Panel) %>%
      mutate(
        Denominator = sum(Frequency, na.rm = TRUE),
        Percentage = ifelse(Denominator > 0, Frequency / Denominator * 100, NA_real_)
      ) %>%
      ungroup()
  } else {
    out <- out %>%
      group_by(.Panel, .Group) %>%
      mutate(
        Denominator = sum(Frequency, na.rm = TRUE),
        Percentage = ifelse(Denominator > 0, Frequency / Denominator * 100, NA_real_)
      ) %>%
      ungroup()
  }
  
  out <- out %>%
    mutate(
      Percentage = round(Percentage, digits),
      Category = factor(.Category, levels = category_levels),
      Group = factor(.Group, levels = c(left_group, right_group)),
      Panel = factor(.Panel, levels = panel_levels)
    ) %>%
    select(Panel, Category, Group, Frequency, Percentage, Denominator)
  
  as.data.frame(out)
}

format_distribution_table <- function(long_df, left_group, right_group, include_panel = FALSE, digits = 2) {
  validate(need(nrow(long_df) > 0, "No distribution data are available."))
  
  left_df <- long_df %>%
    filter(as.character(Group) == left_group) %>%
    transmute(
      Panel = as.character(Panel),
      Category = as.character(Category),
      Left_Frequency = Frequency,
      Left_Percentage = Percentage
    )
  
  right_df <- long_df %>%
    filter(as.character(Group) == right_group) %>%
    transmute(
      Panel = as.character(Panel),
      Category = as.character(Category),
      Right_Frequency = Frequency,
      Right_Percentage = Percentage
    )
  
  out <- full_join(left_df, right_df, by = c("Panel", "Category")) %>%
    mutate(
      Left_Frequency = ifelse(is.na(Left_Frequency), 0, Left_Frequency),
      Right_Frequency = ifelse(is.na(Right_Frequency), 0, Right_Frequency),
      `Total f` = Left_Frequency + Right_Frequency
    )
  
  names(out)[names(out) == "Left_Frequency"] <- paste0(left_group, " f")
  names(out)[names(out) == "Left_Percentage"] <- paste0(left_group, " %")
  names(out)[names(out) == "Right_Frequency"] <- paste0(right_group, " f")
  names(out)[names(out) == "Right_Percentage"] <- paste0(right_group, " %")
  names(out)[names(out) == "Category"] <- "Category"
  
  if (!include_panel) out <- out %>% select(-Panel)
  
  # Apply readable percentage rounding after dynamic renaming.
  pct_cols <- grep(" %$", names(out), value = TRUE)
  for (nm in pct_cols) out[[nm]] <- round(as.numeric(out[[nm]]), digits)
  
  as.data.frame(out)
}

# ============================================================
# PYRAMID PLOT
# ============================================================

create_population_pyramid_plot <- function(chart_df,
                                           left_group,
                                           right_group,
                                           colors,
                                           metric = "Frequency",
                                           label_mode = "Frequency and Percentage",
                                           title = "Population Pyramid",
                                           subtitle = "Demographic distribution by category and group",
                                           category_axis_title = "Age Group",
                                           value_axis_title = "",
                                           legend_title = "Group",
                                           theme_name = "Minimal Scopus Style",
                                           bar_width = 0.82,
                                           show_labels = TRUE,
                                           label_size = 3.2,
                                           label_color = "#111111",
                                           show_center_line = TRUE,
                                           center_line_width = 0.45,
                                           center_line_color = "#555555",
                                           show_grid = TRUE,
                                           panel_enabled = FALSE,
                                           panel_cols = 2,
                                           panel_scales = "fixed",
                                           legend_position = "Top",
                                           title_size = 16,
                                           subtitle_size = 11,
                                           axis_title_size = 11,
                                           axis_text_size = 9,
                                           legend_title_size = 10,
                                           legend_text_size = 9,
                                           panel_title_size = 11,
                                           digits = 2) {
  
  validate(need(nrow(chart_df) > 0, "The selected data are empty."))
  validate(need(all(c(left_group, right_group) %in% as.character(unique(chart_df$Group))), "Both selected groups must be present in the chart data."))
  
  th <- get_theme(theme_name)
  plot_df <- chart_df
  plot_df$MetricValue <- if (identical(metric, "Percentage")) plot_df$Percentage else plot_df$Frequency
  plot_df$MetricValue[is.na(plot_df$MetricValue)] <- 0
  plot_df$PlotValue <- ifelse(as.character(plot_df$Group) == left_group, -plot_df$MetricValue, plot_df$MetricValue)
  
  if (identical(label_mode, "Frequency")) {
    plot_df$DisplayLabel <- ifelse(plot_df$Frequency > 0, format(plot_df$Frequency, trim = TRUE, scientific = FALSE), "")
  } else if (identical(label_mode, "Percentage")) {
    plot_df$DisplayLabel <- ifelse(
      plot_df$Frequency > 0,
      paste0(format(round(plot_df$Percentage, digits), nsmall = digits, trim = TRUE), "%"),
      ""
    )
  } else if (identical(label_mode, "Frequency and Percentage")) {
    plot_df$DisplayLabel <- ifelse(
      plot_df$Frequency > 0,
      paste0(
        format(plot_df$Frequency, trim = TRUE, scientific = FALSE),
        " (", format(round(plot_df$Percentage, digits), nsmall = digits, trim = TRUE), "%)"
      ),
      ""
    )
  } else {
    plot_df$DisplayLabel <- ""
  }
  
  default_value_axis <- if (identical(metric, "Percentage")) "Percentage (%)" else "Frequency"
  x_lab <- if (!is.null(value_axis_title) && nzchar(trimws(value_axis_title))) value_axis_title else default_value_axis
  y_lab <- if (!is.null(category_axis_title) && nzchar(trimws(category_axis_title))) category_axis_title else "Category"
  
  p <- ggplot(plot_df, aes(x = PlotValue, y = Category, fill = Group)) +
    geom_col(width = bar_width, color = "white", linewidth = 0.25) +
    scale_fill_manual(
      values = colors,
      breaks = c(left_group, right_group),
      labels = c(paste0(left_group, " (Left)"), paste0(right_group, " (Right)")),
      drop = FALSE
    ) +
    labs(
      title = title,
      subtitle = subtitle,
      x = x_lab,
      y = y_lab,
      fill = legend_title,
      caption = paste0("Left: ", left_group, "   |   Right: ", right_group)
    ) +
    statcal_pyramid_theme(
      theme_name = theme_name,
      title_size = title_size,
      subtitle_size = subtitle_size,
      axis_title_size = axis_title_size,
      axis_text_size = axis_text_size,
      legend_title_size = legend_title_size,
      legend_text_size = legend_text_size,
      panel_title_size = panel_title_size,
      legend_position = legend_position,
      show_grid = show_grid
    )
  
  if (isTRUE(show_center_line)) {
    p <- p + geom_vline(xintercept = 0, linewidth = center_line_width, color = center_line_color)
  }
  
  if (isTRUE(show_labels) && !identical(label_mode, "None")) {
    p <- p + geom_text(
      aes(x = PlotValue / 2, label = DisplayLabel),
      size = label_size,
      color = label_color,
      check_overlap = TRUE
    )
  }
  
  use_panels <- isTRUE(panel_enabled) && length(unique(as.character(plot_df$Panel))) > 1
  
  if (use_panels && identical(panel_scales, "free_x")) {
    axis_df <- plot_df %>%
      group_by(Panel) %>%
      summarise(
        MaxAbs = max(abs(PlotValue), na.rm = TRUE),
        Category = as.character(Category[1]),
        .groups = "drop"
      ) %>%
      mutate(MaxAbs = ifelse(!is.finite(MaxAbs) | MaxAbs <= 0, 1, MaxAbs * 1.08))
    
    axis_blank <- bind_rows(
      axis_df %>% transmute(Panel, Category, PlotValue = -MaxAbs),
      axis_df %>% transmute(Panel, Category, PlotValue = MaxAbs)
    )
    axis_blank$Category <- factor(axis_blank$Category, levels = levels(plot_df$Category))
    axis_blank$Panel <- factor(axis_blank$Panel, levels = levels(plot_df$Panel))
    
    p <- p + geom_blank(
      data = axis_blank,
      aes(x = PlotValue, y = Category),
      inherit.aes = FALSE
    )
  }
  
  if (use_panels) {
    scales_value <- ifelse(identical(panel_scales, "free_x"), "free_x", "fixed")
    p <- p + facet_wrap(~ Panel, ncol = panel_cols, scales = scales_value)
  }
  
  # Symmetric global axis for the non-faceted or fixed-scale case.
  if (!use_panels || !identical(panel_scales, "free_x")) {
    lim <- max(abs(plot_df$PlotValue), na.rm = TRUE)
    if (!is.finite(lim) || lim <= 0) lim <- 1
    lim <- lim * 1.10
    
    if (identical(metric, "Percentage")) {
      p <- p + scale_x_continuous(
        limits = c(-lim, lim),
        labels = function(x) paste0(scales::number(abs(x), accuracy = ifelse(digits == 0, 1, 10^(-digits))), "%"),
        breaks = scales::pretty_breaks(n = 7),
        expand = expansion(mult = c(0.02, 0.02))
      )
    } else {
      p <- p + scale_x_continuous(
        limits = c(-lim, lim),
        labels = function(x) scales::comma(abs(x), accuracy = 1),
        breaks = scales::pretty_breaks(n = 7),
        expand = expansion(mult = c(0.02, 0.02))
      )
    }
  } else {
    if (identical(metric, "Percentage")) {
      p <- p + scale_x_continuous(
        labels = function(x) paste0(scales::number(abs(x), accuracy = ifelse(digits == 0, 1, 10^(-digits))), "%"),
        breaks = scales::pretty_breaks(n = 6),
        expand = expansion(mult = c(0.03, 0.03))
      )
    } else {
      p <- p + scale_x_continuous(
        labels = function(x) scales::comma(abs(x), accuracy = 1),
        breaks = scales::pretty_breaks(n = 6),
        expand = expansion(mult = c(0.03, 0.03))
      )
    }
  }
  
  p + theme(plot.background = element_rect(fill = th$figure_facecolor, color = NA))
}

# ============================================================
# EXPORT HELPERS
# ============================================================

make_export_filename <- function(prefix, ext = "png", dpi = NULL) {
  stamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
  if (!is.null(dpi)) {
    paste0(prefix, "_", dpi, "dpi_", stamp, ".", ext)
  } else {
    paste0(prefix, "_", stamp, ".", ext)
  }
}

create_session_export_dir <- function(session_token) {
  token <- ifelse(
    is.null(session_token) || !nzchar(session_token),
    paste0(Sys.getpid()),
    safe_token(session_token)
  )
  out <- file.path(tempdir(), paste0("statcal_population_pyramid_", token))
  if (!dir.exists(out)) dir.create(out, recursive = TRUE, showWarnings = FALSE)
  out
}

validate_export_file <- function(path, label = "Export file") {
  if (!file.exists(path)) stop(label, " was not created.")
  size <- file.info(path)$size
  if (is.na(size) || size <= 0) stop(label, " is empty.")
  invisible(TRUE)
}

generate_plot_png <- function(plot_object, file, width = 9, height = 6.5, dpi = 1200, bg = "white") {
  width <- safe_number(width, 9, 3, 30)
  height <- safe_number(height, 6.5, 3, 30)
  dpi <- safe_number(dpi, 1200, 72, 1500)
  if (!inherits(plot_object, "ggplot")) stop("The selected chart is not a ggplot object.")
  
  ggplot2::ggsave(
    filename = file,
    plot = plot_object,
    width = width,
    height = height,
    units = "in",
    dpi = dpi,
    bg = bg,
    limitsize = FALSE
  )
  validate_export_file(file, "PNG file")
  invisible(file)
}

new_export_result <- function(path, filename, message) {
  validate_export_file(path)
  list(
    ok = TRUE,
    file = normalizePath(path, winslash = "/", mustWork = FALSE),
    filename = filename,
    message = message,
    size = file.info(path)$size
  )
}

generated_download_ui <- function(result, download_id, button_label, icon_name = "download") {
  if (is.null(result)) {
    return(tags$p(
      class = "small-note",
      "Click Generate first. After the file has been created and verified, the download button will appear."
    ))
  }
  
  if (!isTRUE(result$ok)) {
    return(tags$div(class = "alert alert-danger", result$message))
  }
  
  tagList(
    tags$div(
      class = "alert alert-success",
      tags$b(result$message), tags$br(),
      tags$span(sprintf("File: %s (%.1f KB)", result$filename, result$size / 1024))
    ),
    downloadButton(download_id, tagList(icon(icon_name), button_label))
  )
}

safe_sheet_name <- function(name) {
  name <- as.character(name)
  invalid_chars <- c("\\", "/", "?", "*", "[", "]", ":")
  for (ch in invalid_chars) name <- gsub(ch, "_", name, fixed = TRUE)
  name <- trimws(name)
  name <- substr(name, 1, 31)
  ifelse(nchar(name) == 0, "Sheet", name)
}

write_table_sheet <- function(wb, sheet_name, df, table_title = NULL) {
  sheet_name <- safe_sheet_name(sheet_name)
  openxlsx::addWorksheet(wb, sheet_name)
  df <- make_display_safe(as.data.frame(df))
  start_row <- 1
  
  if (!is.null(table_title) && nzchar(trimws(table_title))) {
    openxlsx::writeData(wb, sheet_name, table_title, startRow = 1, startCol = 1)
    title_style <- openxlsx::createStyle(
      textDecoration = "bold", fontSize = 13, fontColour = "#1F4E79"
    )
    openxlsx::addStyle(wb, sheet_name, title_style, rows = 1, cols = 1, stack = TRUE)
    start_row <- 3
  }
  
  openxlsx::writeData(wb, sheet_name, df, startRow = start_row)
  
  if (ncol(df) > 0) {
    openxlsx::setColWidths(wb, sheet_name, cols = 1:ncol(df), widths = "auto")
    header_style <- openxlsx::createStyle(
      textDecoration = "bold", fgFill = "#D9EAF7", border = "Bottom"
    )
    openxlsx::addStyle(
      wb, sheet_name, header_style,
      rows = start_row, cols = 1:ncol(df), gridExpand = TRUE
    )
    openxlsx::freezePane(wb, sheet_name, firstActiveRow = start_row + 1)
  }
}

export_population_workbook <- function(file,
                                       raw_df,
                                       distribution_df,
                                       chart_df,
                                       metadata_df,
                                       plot_object,
                                       chart_title,
                                       distribution_title,
                                       chart_data_title,
                                       export_width = 9,
                                       export_height = 6.5,
                                       chart_bg = "white") {
  wb <- openxlsx::createWorkbook()
  
  write_table_sheet(wb, "Export Info", metadata_df, "STATCAL ONLINE - Export Information")
  write_table_sheet(wb, "Raw Data", raw_df, "Appendix A. Raw Data")
  write_table_sheet(wb, "Distribution Table", distribution_df, distribution_title)
  write_table_sheet(wb, "Chart Data", chart_df, chart_data_title)
  
  tmp_png <- tempfile(fileext = ".png")
  on.exit(unlink(tmp_png), add = TRUE)
  generate_plot_png(
    plot_object, tmp_png,
    width = export_width,
    height = export_height,
    dpi = 300,
    bg = chart_bg
  )
  
  openxlsx::addWorksheet(wb, "Figure 1")
  openxlsx::writeData(wb, "Figure 1", chart_title, startRow = 1, startCol = 1)
  fig_style <- openxlsx::createStyle(
    textDecoration = "bold", fontSize = 13, fontColour = "#1F4E79"
  )
  openxlsx::addStyle(wb, "Figure 1", fig_style, rows = 1, cols = 1)
  openxlsx::insertImage(
    wb, "Figure 1", tmp_png,
    startRow = 3, startCol = 1,
    width = min(safe_number(export_width, 9, 4, 16), 12),
    height = min(safe_number(export_height, 6.5, 3, 14), 10),
    units = "in"
  )
  
  openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
  validate_export_file(file, "Excel workbook")
}

make_report_flextable <- function(df, font_size = 8) {
  df <- make_display_safe(as.data.frame(df))
  ft <- flextable::flextable(df)
  ft <- flextable::theme_booktabs(ft)
  ft <- flextable::bold(ft, part = "header")
  ft <- flextable::bg(ft, bg = "#D9EAF7", part = "header")
  ft <- flextable::fontsize(ft, size = font_size, part = "all")
  ft <- flextable::autofit(ft)
  ft
}

export_population_word <- function(file,
                                   report_title,
                                   report_subtitle,
                                   metadata_df,
                                   distribution_df,
                                   chart_df,
                                   raw_df,
                                   plot_object,
                                   distribution_title,
                                   chart_data_title,
                                   figure_title,
                                   include_chart_data = TRUE,
                                   include_raw_data = FALSE,
                                   plot_width = 6.4,
                                   plot_height = 4.6,
                                   chart_bg = "white") {
  doc <- officer::read_docx()
  doc <- officer::body_add_par(doc, report_title, style = "heading 1")
  
  if (!is.null(report_subtitle) && nzchar(trimws(report_subtitle))) {
    doc <- officer::body_add_par(doc, report_subtitle, style = "Normal")
  }
  
  doc <- officer::body_add_par(
    doc,
    paste("Generated by STATCAL ONLINE on", format(Sys.time(), "%d %B %Y %H:%M")),
    style = "Normal"
  )
  
  doc <- officer::body_add_par(doc, "Analysis Settings", style = "heading 2")
  doc <- flextable::body_add_flextable(doc, make_report_flextable(metadata_df, 8))
  
  doc <- officer::body_add_par(doc, distribution_title, style = "heading 2")
  doc <- flextable::body_add_flextable(doc, make_report_flextable(distribution_df, 7.5))
  
  tmp_png <- tempfile(fileext = ".png")
  on.exit(unlink(tmp_png), add = TRUE)
  generate_plot_png(
    plot_object, tmp_png,
    width = plot_width,
    height = plot_height,
    dpi = 300,
    bg = chart_bg
  )
  
  doc <- officer::body_add_par(doc, figure_title, style = "heading 2")
  doc <- officer::body_add_img(doc, src = tmp_png, width = plot_width, height = plot_height)
  
  if (isTRUE(include_chart_data)) {
    doc <- officer::body_add_par(doc, chart_data_title, style = "heading 2")
    doc <- flextable::body_add_flextable(doc, make_report_flextable(chart_df, 7.2))
  }
  
  if (isTRUE(include_raw_data)) {
    doc <- officer::body_add_break(doc)
    doc <- officer::body_add_par(doc, "Appendix A. Raw Data", style = "heading 2")
    doc <- flextable::body_add_flextable(doc, make_report_flextable(raw_df, 6.8))
  }
  
  print(doc, target = file)
  validate_export_file(file, "Word report")
}

# ============================================================
# UI
# ============================================================

ui <- dashboardPage(
  dashboardHeader(title = APP_NAME, titleWidth = "100%"),
  dashboardSidebar(disable = TRUE),
  dashboardBody(
    tags$head(
      tags$style(HTML("
        .content-wrapper, .right-side { background-color: #f7f9fb; }
        .box { border-radius: 10px; }
        .statcal-title { font-size: 26px; font-weight: 700; color: #1F4E79; }
        .statcal-subtitle { font-size: 18px; font-weight: 600; color: #333333; }
        .statcal-note { line-height: 1.6; text-align: justify; }
        .small-note { font-size: 12px; color: #666666; line-height: 1.5; }
        .section-note { background: #f2f7fb; border-left: 4px solid #1F4E79; padding: 10px 12px; margin-bottom: 12px; }
      "))
    ),
    
    fluidRow(
      box(
        width = 12, status = "primary", solidHeader = TRUE,
        title = "STATCAL ONLINE - Population Pyramid Visualizer",
        div(class = "statcal-title", APP_TITLE),
        div(class = "statcal-subtitle", APP_UPDATED),
        tags$p(
          class = "statcal-note",
          "This application creates publication-ready population pyramid charts from Excel data. Users can select a category variable, define left-side and right-side groups, manually order categories, optionally create facet panels, display frequency or percentage values, customize chart appearance, and export the results to PNG, Excel, and Word."
        ),
        tags$p(
          tags$b("Website: "), tags$a(href = WEBSITE_URL, target = "_blank", WEBSITE_URL), tags$br(),
          tags$b("STATCAL ONLINE Page: "), tags$a(href = STATCAL_ONLINE_URL, target = "_blank", STATCAL_ONLINE_URL)
        )
      )
    ),
    
    tabsetPanel(
      id = "main_tabs",
      
      # --------------------------------------------------------
      # 1. DATA & SETTINGS
      # --------------------------------------------------------
      tabPanel(
        "1. Data & Settings",
        br(),
        fluidRow(
          box(
            width = 4, title = "Data Input", status = "primary", solidHeader = TRUE,
            fileInput("uploaded_file", "Upload Excel file", accept = c(".xlsx", ".xls")),
            uiOutput("sheet_ui"),
            radioButtons(
              "data_format", "Data format",
              choices = c("Individual / raw data", "Aggregated frequency data"),
              selected = "Individual / raw data"
            ),
            tags$p(
              class = "small-note",
              "If no Excel file is uploaded, STATCAL will use the sample data file when it is available beside app.R."
            )
          ),
          box(
            width = 4, title = "Pyramid Variables", status = "primary", solidHeader = TRUE,
            uiOutput("category_var_ui"),
            uiOutput("group_var_ui"),
            uiOutput("frequency_var_ui"),
            uiOutput("left_group_ui"),
            uiOutput("right_group_ui")
          ),
          box(
            width = 4, title = "Category Order", status = "primary", solidHeader = TRUE,
            uiOutput("category_order_ui"),
            sliderInput("decimal_digits", "Decimal digits for percentages", min = 0, max = 5, value = 2, step = 1),
            tags$p(
              class = "small-note",
              "Enter category labels separated by commas. Categories not listed will automatically be appended after the listed categories."
            )
          )
        ),
        
        fluidRow(
          box(
            width = 12, title = "Optional Panel Settings", status = "info", solidHeader = TRUE,
            collapsible = TRUE,
            fluidRow(
              column(4, uiOutput("panel_var_ui")),
              column(4, uiOutput("panel_values_ui")),
              column(4, uiOutput("panel_order_ui"))
            ),
            tags$p(
              class = "small-note",
              "A panel variable can be Province, Year, Region, Education, or another categorical variable. Select only the panel categories that you want to display."
            )
          )
        ),
        
        fluidRow(
          valueBoxOutput("metric_rows", width = 3),
          valueBoxOutput("metric_categories", width = 3),
          valueBoxOutput("metric_left", width = 3),
          valueBoxOutput("metric_right", width = 3)
        ),
        
        fluidRow(
          box(
            width = 12, title = "Dataset Preview", status = "warning", solidHeader = TRUE,
            shinycssloaders::withSpinner(DTOutput("data_preview"))
          )
        )
      ),
      
      # --------------------------------------------------------
      # 2. DISTRIBUTION TABLE
      # --------------------------------------------------------
      tabPanel(
        "2. Distribution Table",
        br(),
        fluidRow(
          box(
            width = 12, title = "Population Pyramid Distribution Table", status = "warning", solidHeader = TRUE,
            tags$div(
              class = "section-note",
              "The table and the population pyramid are generated from the same reactive distribution dataset. This keeps the frequencies and percentages synchronized across the table, chart, Excel export, and Word export."
            ),
            shinycssloaders::withSpinner(DTOutput("distribution_table"))
          )
        )
      ),
      
      # --------------------------------------------------------
      # 3. POPULATION PYRAMID
      # --------------------------------------------------------
      tabPanel(
        "3. Population Pyramid",
        br(),
        fluidRow(
          box(
            width = 3, title = "Chart Data", status = "primary", solidHeader = TRUE,
            selectInput("chart_metric", "Bar value", choices = c("Frequency", "Percentage"), selected = "Frequency"),
            selectInput(
              "percentage_denominator", "Percentage denominator",
              choices = c(
                "Within each group (per panel)",
                "Within each panel",
                "Total selected sample"
              ),
              selected = "Within each group (per panel)"
            ),
            selectInput(
              "chart_label_mode", "Information on bars",
              choices = c("None", "Frequency", "Percentage", "Frequency and Percentage"),
              selected = "Frequency and Percentage"
            ),
            checkboxInput("chart_show_labels", "Show data labels", value = TRUE),
            sliderInput("chart_label_size", "Data label size", min = 2, max = 8, value = 3.0, step = 0.2),
            colourpicker::colourInput("chart_label_color", "Data label color", value = "#111111", showColour = "both")
          ),
          
          box(
            width = 3, title = "Colors & Bars", status = "primary", solidHeader = TRUE,
            selectInput("chart_palette", "Pyramid color palette", choices = names(PYRAMID_PALETTES), selected = "Red - Green"),
            uiOutput("palette_preview_ui"),
            colourpicker::colourInput("left_color", "Manual left-side color", value = "#D95F59", showColour = "both"),
            colourpicker::colourInput("right_color", "Manual right-side color", value = "#59A96A", showColour = "both"),
            sliderInput("bar_width", "Bar width", min = 0.30, max = 1.00, value = 0.82, step = 0.01),
            checkboxInput("chart_show_center_line", "Show center line", value = TRUE),
            sliderInput("center_line_width", "Center line width", min = 0.1, max = 2.0, value = 0.45, step = 0.05),
            colourpicker::colourInput("center_line_color", "Center line color", value = "#555555", showColour = "both")
          ),
          
          box(
            width = 3, title = "Titles & Theme", status = "primary", solidHeader = TRUE,
            textInput("chart_title", "Chart title", value = "Population Pyramid"),
            textInput("chart_subtitle", "Chart subtitle", value = "Demographic distribution by age group and sex"),
            textInput("category_axis_title", "Category-axis title", value = "Age Group"),
            textInput("value_axis_title", "Value-axis title (blank = automatic)", value = ""),
            textInput("chart_legend_title", "Legend title", value = "Group"),
            selectInput("chart_theme", "Background theme", choices = names(THEMES), selected = "Minimal Scopus Style"),
            selectInput("chart_legend_position", "Legend position", choices = LEGEND_CHOICES, selected = "Top"),
            checkboxInput("chart_show_grid", "Show value-axis grid", value = TRUE)
          ),
          
          box(
            width = 3, title = "Panel Layout", status = "primary", solidHeader = TRUE,
            tags$p(class = "small-note", tags$b("Panel behavior:"), " panels are created automatically whenever a Panel variable is selected in Data & Settings."),
            sliderInput("chart_panel_cols", "Number of panel columns", min = 1, max = 4, value = 2, step = 1),
            selectInput(
              "chart_panel_scales", "Panel value-axis scale",
              choices = c(
                "Same scale across panels" = "fixed",
                "Independent symmetric scale per panel" = "free_x"
              ),
              selected = "fixed"
            ),
            sliderInput("chart_height_px", "Preview chart height (px)", min = 400, max = 1400, value = 700, step = 50),
            tags$p(
              class = "small-note",
              "Use the same scale when panels are intended for direct comparison. Independent scales are useful when panel sizes differ greatly."
            )
          )
        ),
        
        fluidRow(
          box(
            width = 12, title = "Flexible Text Size Settings", status = "info", solidHeader = TRUE,
            collapsible = TRUE, collapsed = TRUE,
            fluidRow(
              column(3, sliderInput("chart_title_size", "Title", 8, 34, 16, 1)),
              column(3, sliderInput("chart_subtitle_size", "Subtitle", 6, 26, 11, 1)),
              column(3, sliderInput("chart_axis_title_size", "Axis title", 6, 24, 11, 1)),
              column(3, sliderInput("chart_axis_text_size", "Axis text", 5, 22, 9, 1)),
              column(3, sliderInput("chart_legend_title_size", "Legend title", 5, 24, 10, 1)),
              column(3, sliderInput("chart_legend_text_size", "Legend text", 5, 22, 9, 1)),
              column(3, sliderInput("chart_panel_title_size", "Panel title", 6, 24, 11, 1))
            )
          )
        ),
        
        fluidRow(
          box(
            width = 12, title = "Publication-Ready Population Pyramid", status = "warning", solidHeader = TRUE,
            shinycssloaders::withSpinner(uiOutput("population_plot_ui"))
          )
        ),
        
        fluidRow(
          box(
            width = 12, title = "Chart Data", status = "info", solidHeader = TRUE,
            collapsible = TRUE, collapsed = TRUE,
            shinycssloaders::withSpinner(DTOutput("chart_data_table"))
          )
        )
      ),
      
      # --------------------------------------------------------
      # 4. EXPORT
      # --------------------------------------------------------
      tabPanel(
        "4. Export",
        br(),
        fluidRow(
          box(
            width = 12, title = "Export Settings", status = "primary", solidHeader = TRUE,
            fluidRow(
              column(4, selectInput("export_dpi", "PNG resolution / DPI", choices = c(300, 600, 900, 1200, 1500), selected = 1200)),
              column(4, numericInput("export_width", "Export width (inches)", value = 9, min = 4, max = 30, step = 0.5)),
              column(4, numericInput("export_height", "Export height (inches)", value = 6.5, min = 3, max = 30, step = 0.5))
            ),
            fluidRow(
              column(6, textInput("word_report_title", "Word report title", value = "STATCAL Population Pyramid Analysis Report")),
              column(6, textInput("word_report_subtitle", "Word report subtitle", value = "Population distribution and demographic pyramid visualization"))
            ),
            fluidRow(
              column(6, textInput("distribution_table_title", "Distribution table title", value = "Table 1. Population Pyramid Frequency and Percentage Distribution")),
              column(6, textInput("chart_data_title", "Chart data title", value = "Table 2. Data Used to Create the Population Pyramid"))
            ),
            textInput("figure_title", "Figure title", value = "Figure 1. Population Pyramid"),
            checkboxInput("word_include_chart_data", "Include Chart Data table in Word", value = TRUE),
            checkboxInput("word_include_raw_data", "Include Raw Data appendix in Word", value = FALSE),
            tags$p(
              class = "small-note",
              "Files are generated first in a writable session temporary folder, verified, and then delivered through Shiny downloadHandler. Excel exports include the pyramid figure as an embedded image."
            )
          )
        ),
        
        fluidRow(
          box(
            width = 4, title = "Population Pyramid PNG", status = "warning", solidHeader = TRUE,
            actionButton("generate_chart_png", "Generate PNG", icon = icon("image")),
            br(), br(), uiOutput("chart_generated_download_ui"),
            br(), downloadButton("download_chart_fallback", "Direct / Fallback PNG")
          ),
          box(
            width = 4, title = "Analysis Excel", status = "success", solidHeader = TRUE,
            actionButton("generate_excel", "Generate Analysis Excel", icon = icon("file-excel")),
            br(), br(), uiOutput("excel_generated_download_ui"),
            br(), downloadButton("download_excel_fallback", "Direct / Fallback Excel")
          ),
          box(
            width = 4, title = "Analysis Word", status = "info", solidHeader = TRUE,
            actionButton("generate_word", "Generate Analysis Word", icon = icon("file-word")),
            br(), br(), uiOutput("word_generated_download_ui"),
            br(), downloadButton("download_word_fallback", "Direct / Fallback Word")
          )
        )
      )
    )
  )
)

# ============================================================
# SERVER
# ============================================================

server <- function(input, output, session) {
  
  chart_export_result <- reactiveVal(NULL)
  excel_export_result <- reactiveVal(NULL)
  word_export_result <- reactiveVal(NULL)
  session_export_dir <- create_session_export_dir(session$token)
  
  session$onSessionEnded(function() {
    try(unlink(session_export_dir, recursive = TRUE, force = TRUE), silent = TRUE)
  })
  
  output$chart_generated_download_ui <- renderUI({
    generated_download_ui(chart_export_result(), "download_chart_generated", "Download Generated PNG", "image")
  })
  
  output$excel_generated_download_ui <- renderUI({
    generated_download_ui(excel_export_result(), "download_excel_generated", "Download Generated Excel", "file-excel")
  })
  
  output$word_generated_download_ui <- renderUI({
    generated_download_ui(word_export_result(), "download_word_generated", "Download Generated Word", "file-word")
  })
  
  # ----------------------------------------------------------
  # DATA SOURCE
  # ----------------------------------------------------------
  
  current_excel_path <- reactive({
    if (!is.null(input$uploaded_file)) {
      input$uploaded_file$datapath
    } else {
      find_sample_data()
    }
  })
  
  output$sheet_ui <- renderUI({
    path <- current_excel_path()
    if (is.null(path)) {
      return(helpText("Please upload an Excel file to start the analysis."))
    }
    sheets <- readxl::excel_sheets(path)
    selectInput("sheet_name", "Worksheet", choices = sheets, selected = sheets[1])
  })
  
  data_raw <- reactive({
    path <- current_excel_path()
    req(path)
    sheets <- readxl::excel_sheets(path)
    sheet <- input$sheet_name
    if (is.null(sheet) || !(sheet %in% sheets)) sheet <- sheets[1]
    clean_dataframe(readxl::read_excel(path, sheet = sheet))
  })
  
  # ----------------------------------------------------------
  # DYNAMIC VARIABLE SELECTION
  # ----------------------------------------------------------
  
  output$category_var_ui <- renderUI({
    df <- data_raw()
    cols <- names(df)
    preferred <- preferred_column(
      cols,
      c("Kelompok_Usia", "Kelompok Usia", "Age_Group", "Age Group", "Category", "Kategori")
    )
    selectInput("category_var", "Category variable", choices = cols, selected = preferred)
  })
  
  output$group_var_ui <- renderUI({
    df <- data_raw()
    cols <- names(df)
    preferred <- preferred_column(
      cols,
      c("Jenis_Kelamin", "Jenis Kelamin", "Gender", "Sex", "Group", "Kelompok")
    )
    selectInput("group_var", "Group variable", choices = cols, selected = preferred)
  })
  
  output$frequency_var_ui <- renderUI({
    if (!identical(input$data_format, "Aggregated frequency data")) return(NULL)
    df <- data_raw()
    numeric_cols <- names(df)[vapply(df, is.numeric, logical(1))]
    choices <- if (length(numeric_cols) > 0) numeric_cols else names(df)
    preferred <- preferred_column(choices, c("Frequency", "Freq", "Frekuensi", "Count", "N"))
    selectInput("frequency_var", "Frequency variable", choices = choices, selected = preferred)
  })
  
  group_values <- reactive({
    df <- data_raw()
    g <- input$group_var
    req(g, g %in% names(df))
    sorted_unique_values(df[[g]])
  })
  
  output$left_group_ui <- renderUI({
    vals <- group_values()
    preferred <- preferred_group_value(vals, c("Laki-laki", "Laki Laki", "Male", "M", "Pria"), 1)
    selectInput("left_group", "Left-side group", choices = vals, selected = preferred)
  })
  
  output$right_group_ui <- renderUI({
    vals <- group_values()
    preferred <- preferred_group_value(vals, c("Perempuan", "Female", "F", "Wanita"), min(2, length(vals)))
    selectInput("right_group", "Right-side group", choices = vals, selected = preferred)
  })
  
  category_values <- reactive({
    df <- data_raw()
    cvar <- input$category_var
    req(cvar, cvar %in% names(df))
    smart_category_order(df[[cvar]])
  })
  
  output$category_order_ui <- renderUI({
    vals <- category_values()
    textInput(
      "category_order_text",
      "Manual category order (comma-separated)",
      value = paste(vals, collapse = ", ")
    )
  })
  
  selected_category_order <- reactive({
    vals <- category_values()
    parse_manual_order(input$category_order_text, vals)
  })
  
  output$panel_var_ui <- renderUI({
    df <- data_raw()
    exclude <- c(input$category_var, input$group_var)
    if (identical(input$data_format, "Aggregated frequency data")) exclude <- c(exclude, input$frequency_var)
    choices <- c("None", setdiff(names(df), exclude))
    
    current <- isolate(input$panel_var)
    preferred <- if (!is.null(current) && current %in% choices) current else "None"
    selectInput("panel_var", "Panel variable (optional)", choices = choices, selected = preferred)
  })
  
  panel_available_values <- reactive({
    df <- data_raw()
    pvar <- input$panel_var
    if (is.null(pvar) || pvar == "None" || !(pvar %in% names(df))) return(character(0))
    sorted_unique_values(df[[pvar]])
  })
  
  output$panel_values_ui <- renderUI({
    vals <- panel_available_values()
    if (length(vals) == 0) {
      return(tags$p(class = "small-note", "No panel categories are needed when Panel variable is None."))
    }
    selectizeInput(
      "panel_values",
      "Panel categories to display",
      choices = vals,
      selected = vals,
      multiple = TRUE,
      options = list(plugins = list("remove_button"))
    )
  })
  
  selected_panel_values <- reactive({
    vals <- panel_available_values()
    if (length(vals) == 0) return(character(0))
    selected <- input$panel_values
    if (is.null(selected) || length(selected) == 0) return(vals)
    selected[selected %in% vals]
  })
  
  output$panel_order_ui <- renderUI({
    vals <- selected_panel_values()
    if (length(vals) == 0) {
      return(tags$p(class = "small-note", "Panel order is not required."))
    }
    textInput(
      "panel_order_text",
      "Panel order (comma-separated)",
      value = paste(vals, collapse = ", ")
    )
  })
  
  selected_panel_order <- reactive({
    vals <- selected_panel_values()
    if (length(vals) == 0) return(character(0))
    parse_manual_order(input$panel_order_text, vals)
  })
  
  # ----------------------------------------------------------
  # DATA OBJECTS
  # ----------------------------------------------------------
  
  pyramid_long_data <- reactive({
    build_pyramid_long_data(
      df = data_raw(),
      data_format = input$data_format,
      category_col = input$category_var,
      group_col = input$group_var,
      left_group = input$left_group,
      right_group = input$right_group,
      frequency_col = input$frequency_var,
      category_order = selected_category_order(),
      panel_col = input$panel_var,
      selected_panels = selected_panel_values(),
      panel_order = selected_panel_order(),
      percentage_denominator = input$percentage_denominator,
      digits = input$decimal_digits
    )
  })
  
  distribution_table_data <- reactive({
    use_panel <- !is.null(input$panel_var) && input$panel_var != "None"
    format_distribution_table(
      pyramid_long_data(),
      left_group = input$left_group,
      right_group = input$right_group,
      include_panel = use_panel,
      digits = input$decimal_digits
    )
  })
  
  chart_data_for_display <- reactive({
    df <- pyramid_long_data()
    out <- make_display_safe(df)
    names(out)[names(out) == "Panel"] <- "Panel"
    names(out)[names(out) == "Category"] <- "Category"
    names(out)[names(out) == "Group"] <- "Group"
    out
  })
  
  # ----------------------------------------------------------
  # VALUE BOXES AND TABLES
  # ----------------------------------------------------------
  
  output$metric_rows <- renderValueBox({
    valueBox(nrow(data_raw()), "Rows in uploaded data", icon = icon("table"), color = "blue")
  })
  
  output$metric_categories <- renderValueBox({
    valueBox(length(selected_category_order()), "Category levels", icon = icon("sort"), color = "yellow")
  })
  
  output$metric_left <- renderValueBox({
    df <- pyramid_long_data()
    value <- sum(df$Frequency[as.character(df$Group) == input$left_group], na.rm = TRUE)
    valueBox(format(value, big.mark = ",", scientific = FALSE), paste0("Left: ", input$left_group), icon = icon("arrow-left"), color = "red")
  })
  
  output$metric_right <- renderValueBox({
    df <- pyramid_long_data()
    value <- sum(df$Frequency[as.character(df$Group) == input$right_group], na.rm = TRUE)
    valueBox(format(value, big.mark = ",", scientific = FALSE), paste0("Right: ", input$right_group), icon = icon("arrow-right"), color = "green")
  })
  
  output$data_preview <- renderDT({
    DT::datatable(
      make_display_safe(data_raw()),
      options = list(scrollX = TRUE, pageLength = 10),
      rownames = FALSE
    )
  })
  
  output$distribution_table <- renderDT({
    DT::datatable(
      make_display_safe(distribution_table_data()),
      options = list(scrollX = TRUE, pageLength = 15),
      rownames = FALSE
    )
  })
  
  output$chart_data_table <- renderDT({
    DT::datatable(
      chart_data_for_display(),
      options = list(scrollX = TRUE, pageLength = 15),
      rownames = FALSE
    )
  })
  
  # ----------------------------------------------------------
  # PALETTE AND PLOT
  # ----------------------------------------------------------
  
  output$palette_preview_ui <- renderUI({
    left <- input$left_group
    right <- input$right_group
    if (is.null(left) || is.null(right)) return(NULL)
    cols <- get_pyramid_colors(input, left, right)
    
    tags$div(
      style = "display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:6px 0 12px 0;",
      tags$div(style = paste0("width:42px;height:22px;border-radius:4px;border:1px solid #aaa;background:", cols[1], ";")),
      tags$span(left),
      tags$div(style = paste0("width:42px;height:22px;border-radius:4px;border:1px solid #aaa;background:", cols[2], ";")),
      tags$span(right)
    )
  })
  
  population_plot_object <- reactive({
    left <- input$left_group
    right <- input$right_group
    req(left, right)
    
    panel_enabled <- !is.null(input$panel_var) && input$panel_var != "None"
    
    create_population_pyramid_plot(
      chart_df = pyramid_long_data(),
      left_group = left,
      right_group = right,
      colors = get_pyramid_colors(input, left, right),
      metric = input$chart_metric,
      label_mode = input$chart_label_mode,
      title = input$chart_title,
      subtitle = input$chart_subtitle,
      category_axis_title = input$category_axis_title,
      value_axis_title = input$value_axis_title,
      legend_title = input$chart_legend_title,
      theme_name = input$chart_theme,
      bar_width = input$bar_width,
      show_labels = input$chart_show_labels,
      label_size = input$chart_label_size,
      label_color = input$chart_label_color,
      show_center_line = input$chart_show_center_line,
      center_line_width = input$center_line_width,
      center_line_color = input$center_line_color,
      show_grid = input$chart_show_grid,
      panel_enabled = panel_enabled,
      panel_cols = input$chart_panel_cols,
      panel_scales = input$chart_panel_scales,
      legend_position = input$chart_legend_position,
      title_size = input$chart_title_size,
      subtitle_size = input$chart_subtitle_size,
      axis_title_size = input$chart_axis_title_size,
      axis_text_size = input$chart_axis_text_size,
      legend_title_size = input$chart_legend_title_size,
      legend_text_size = input$chart_legend_text_size,
      panel_title_size = input$chart_panel_title_size,
      digits = input$decimal_digits
    )
  })
  
  output$population_plot_ui <- renderUI({
    plotOutput(
      "population_pyramid_plot",
      height = paste0(safe_number(input$chart_height_px, 700, 400, 1600), "px")
    )
  })
  
  output$population_pyramid_plot <- renderPlot({
    population_plot_object()
  })
  
  # ----------------------------------------------------------
  # EXPORT METADATA
  # ----------------------------------------------------------
  
  build_export_metadata <- reactive({
    panel_var_value <- if (is.null(input$panel_var)) "None" else as.character(input$panel_var)
    panel_values_value <- if (length(selected_panel_order()) > 0) paste(selected_panel_order(), collapse = ", ") else "None"
    colors <- get_pyramid_colors(input, input$left_group, input$right_group)
    
    data.frame(
      Item = c(
        "Application",
        "Export Time",
        "Data Format",
        "Rows in Uploaded Data",
        "Category Variable",
        "Category Order",
        "Group Variable",
        "Left-side Group",
        "Right-side Group",
        "Frequency Variable",
        "Panel Variable",
        "Selected Panel Categories",
        "Chart Metric",
        "Percentage Denominator",
        "Left Color",
        "Right Color",
        "Theme"
      ),
      Value = c(
        paste(APP_NAME, APP_TITLE, sep = " - "),
        as.character(Sys.time()),
        input$data_format,
        as.character(nrow(data_raw())),
        input$category_var,
        paste(selected_category_order(), collapse = ", "),
        input$group_var,
        input$left_group,
        input$right_group,
        ifelse(identical(input$data_format, "Aggregated frequency data"), input$frequency_var, "Not applicable"),
        panel_var_value,
        panel_values_value,
        input$chart_metric,
        input$percentage_denominator,
        unname(colors[1]),
        unname(colors[2]),
        input$chart_theme
      ),
      stringsAsFactors = FALSE
    )
  })
  
  export_current_excel <- function(file) {
    export_population_workbook(
      file = file,
      raw_df = data_raw(),
      distribution_df = distribution_table_data(),
      chart_df = chart_data_for_display(),
      metadata_df = build_export_metadata(),
      plot_object = population_plot_object(),
      chart_title = input$figure_title,
      distribution_title = input$distribution_table_title,
      chart_data_title = input$chart_data_title,
      export_width = input$export_width,
      export_height = input$export_height,
      chart_bg = safe_theme_bg(input$chart_theme)
    )
  }
  
  export_current_word <- function(file) {
    word_width <- min(safe_number(input$export_width, 9, 4, 12), 6.5)
    ratio <- safe_number(input$export_height, 6.5, 3, 20) / safe_number(input$export_width, 9, 4, 30)
    word_height <- max(3, min(8.5, word_width * ratio))
    
    export_population_word(
      file = file,
      report_title = input$word_report_title,
      report_subtitle = input$word_report_subtitle,
      metadata_df = build_export_metadata(),
      distribution_df = distribution_table_data(),
      chart_df = chart_data_for_display(),
      raw_df = data_raw(),
      plot_object = population_plot_object(),
      distribution_title = input$distribution_table_title,
      chart_data_title = input$chart_data_title,
      figure_title = input$figure_title,
      include_chart_data = isTRUE(input$word_include_chart_data),
      include_raw_data = isTRUE(input$word_include_raw_data),
      plot_width = word_width,
      plot_height = word_height,
      chart_bg = safe_theme_bg(input$chart_theme)
    )
  }
  
  make_result_safe <- function(expr, filename, success_message) {
    tryCatch({
      path <- file.path(session_export_dir, filename)
      expr(path)
      new_export_result(path, filename, success_message)
    }, error = function(e) {
      list(
        ok = FALSE,
        message = paste("Export failed:", conditionMessage(e)),
        file = NULL,
        filename = filename,
        size = 0
      )
    })
  }
  
  # ----------------------------------------------------------
  # GENERATE BUTTONS
  # ----------------------------------------------------------
  
  observeEvent(input$generate_chart_png, {
    filename <- make_export_filename("statcal_population_pyramid", "png", input$export_dpi)
    result <- make_result_safe(
      function(path) {
        generate_plot_png(
          population_plot_object(),
          path,
          input$export_width,
          input$export_height,
          input$export_dpi,
          safe_theme_bg(input$chart_theme)
        )
      },
      filename,
      "PNG figure has been generated and verified successfully."
    )
    chart_export_result(result)
  })
  
  observeEvent(input$generate_excel, {
    filename <- make_export_filename("statcal_population_pyramid_analysis", "xlsx")
    result <- make_result_safe(
      function(path) export_current_excel(path),
      filename,
      "Excel workbook has been generated and verified successfully."
    )
    excel_export_result(result)
  })
  
  observeEvent(input$generate_word, {
    filename <- make_export_filename("statcal_population_pyramid_report", "docx")
    result <- make_result_safe(
      function(path) export_current_word(path),
      filename,
      "Word report has been generated and verified successfully."
    )
    word_export_result(result)
  })
  
  # ----------------------------------------------------------
  # GENERATED DOWNLOAD HANDLERS
  # ----------------------------------------------------------
  
  output$download_chart_generated <- downloadHandler(
    filename = function() {
      req(chart_export_result())
      chart_export_result()$filename
    },
    contentType = "image/png",
    content = function(file) {
      res <- chart_export_result()
      req(res, isTRUE(res$ok), file.exists(res$file))
      file.copy(res$file, file, overwrite = TRUE)
    }
  )
  
  output$download_excel_generated <- downloadHandler(
    filename = function() {
      req(excel_export_result())
      excel_export_result()$filename
    },
    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    content = function(file) {
      res <- excel_export_result()
      req(res, isTRUE(res$ok), file.exists(res$file))
      file.copy(res$file, file, overwrite = TRUE)
    }
  )
  
  output$download_word_generated <- downloadHandler(
    filename = function() {
      req(word_export_result())
      word_export_result()$filename
    },
    contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    content = function(file) {
      res <- word_export_result()
      req(res, isTRUE(res$ok), file.exists(res$file))
      file.copy(res$file, file, overwrite = TRUE)
    }
  )
  
  # ----------------------------------------------------------
  # DIRECT / FALLBACK DOWNLOAD HANDLERS
  # ----------------------------------------------------------
  
  output$download_chart_fallback <- downloadHandler(
    filename = function() make_export_filename("statcal_population_pyramid", "png", input$export_dpi),
    contentType = "image/png",
    content = function(file) {
      generate_plot_png(
        population_plot_object(),
        file,
        input$export_width,
        input$export_height,
        input$export_dpi,
        safe_theme_bg(input$chart_theme)
      )
    }
  )
  
  output$download_excel_fallback <- downloadHandler(
    filename = function() make_export_filename("statcal_population_pyramid_analysis", "xlsx"),
    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    content = function(file) export_current_excel(file)
  )
  
  output$download_word_fallback <- downloadHandler(
    filename = function() make_export_filename("statcal_population_pyramid_report", "docx"),
    contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    content = function(file) export_current_word(file)
  )
}

shinyApp(ui = ui, server = server)
