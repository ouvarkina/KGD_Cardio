# lib/death_green_xml.R
# ---------------------------------------------------------
# Вспомогательные функции для определения death_green
# по цвету шрифта (ARGB) из XLSX/XLSM через XML (styles + sheet)
# Поддерживает цвета через rgb и через theme + tint.
# ---------------------------------------------------------

# зависимости: xml2, stringr
# (подгрузим локально, чтобы main-скрипт был чище)
if (!requireNamespace("xml2", quietly = TRUE)) stop("Package 'xml2' is required")
if (!requireNamespace("stringr", quietly = TRUE)) stop("Package 'stringr' is required")

# Детектор "зелёного" по ARGB (FFRRGGBB)
is_green_rgb <- function(rgb) {
  if (is.na(rgb) || rgb == "") return(FALSE)
  rgb <- toupper(rgb)
  
  # иногда без альфы: "00B050"
  if (stringr::str_detect(rgb, "^[0-9A-F]{6}$")) rgb <- paste0("FF", rgb)
  if (!stringr::str_detect(rgb, "^[0-9A-F]{8}$")) return(FALSE)
  
  r <- strtoi(substr(rgb, 3, 4), 16L)
  g <- strtoi(substr(rgb, 5, 6), 16L)
  b <- strtoi(substr(rgb, 7, 8), 16L)
  
  (g >= 80) && (g - r >= 30) && (g - b >= 30)
}

# 1-based индекс колонки -> Excel буквы (1->A, 27->AA)
excel_col_letters <- function(n) {
  if (is.na(n) || n < 1) return(NA_character_)
  out <- ""
  while (n > 0) {
    r <- (n - 1) %% 26
    out <- paste0(LETTERS[r + 1], out)
    n <- (n - 1) %/% 26
  }
  out
}

zip_list_names <- function(xlsx_path) {
  tryCatch(utils::unzip(xlsx_path, list = TRUE)$Name, error = function(e) character(0))
}

# sheet name -> "xl/worksheets/sheetN.xml"
get_sheet_xml_path <- function(xlsx_path, sheet_name) {
  listed <- zip_list_names(xlsx_path)
  if (!("xl/workbook.xml" %in% listed) || !("xl/_rels/workbook.xml.rels" %in% listed)) return(NA_character_)
  
  wb <- tryCatch(xml2::read_xml(unz(xlsx_path, "xl/workbook.xml")), error = function(e) NULL)
  if (is.null(wb)) return(NA_character_)
  
  sheet_node <- xml2::xml_find_first(wb, sprintf(".//*[local-name()='sheet' and @name='%s']", sheet_name))
  if (inherits(sheet_node, "xml_missing")) return(NA_character_)
  
  attrs <- xml2::xml_attrs(sheet_node)
  
  # relationship id может прийти как "r:id" или "id"
  rid <- NA_character_
  if ("r:id" %in% names(attrs)) rid <- attrs[["r:id"]]
  if ((is.na(rid) || rid == "") && ("id" %in% names(attrs))) rid <- attrs[["id"]]
  
  # последний шанс — любой атрибут, заканчивающийся на id, но не sheetId
  if (is.na(rid) || rid == "") {
    cand <- names(attrs)[grepl("(^id$|:id$|\\}id$)", names(attrs), ignore.case = TRUE)]
    cand <- cand[!tolower(cand) %in% c("sheetid", "sheetId")]
    if (length(cand) > 0) rid <- attrs[[cand[1]]]
  }
  if (is.na(rid) || rid == "") return(NA_character_)
  
  rels <- tryCatch(xml2::read_xml(unz(xlsx_path, "xl/_rels/workbook.xml.rels")), error = function(e) NULL)
  if (is.null(rels)) return(NA_character_)
  
  rel_node <- xml2::xml_find_first(rels, sprintf(".//*[local-name()='Relationship' and @Id='%s']", rid))
  if (inherits(rel_node, "xml_missing")) return(NA_character_)
  
  target <- xml2::xml_attr(rel_node, "Target")
  if (is.na(target) || target == "") return(NA_character_)
  
  sheet_xml <- paste0("xl/", target)
  if (!(sheet_xml %in% listed)) return(NA_character_)
  
  sheet_xml
}

# theme colors (для случаев theme+tint вместо rgb)
read_theme_scheme <- function(xlsx_path) {
  listed <- zip_list_names(xlsx_path)
  theme_path <- listed[grepl("^xl/theme/theme\\d+\\.xml$", listed)][1]
  if (is.na(theme_path)) return(list())
  
  th <- tryCatch(xml2::read_xml(unz(xlsx_path, theme_path)), error = function(e) NULL)
  if (is.null(th)) return(list())
  
  scheme <- c("lt1","dk1","lt2","dk2",
              "accent1","accent2","accent3","accent4","accent5","accent6",
              "hlink","folHlink")
  out <- list()
  
  for (nm in scheme) {
    node <- xml2::xml_find_first(th, sprintf(".//*[local-name()='clrScheme']/*[local-name()='%s']", nm))
    if (inherits(node, "xml_missing")) next
    
    srgb <- xml2::xml_find_first(node, ".//*[local-name()='srgbClr']")
    if (!inherits(srgb, "xml_missing")) {
      val <- xml2::xml_attr(srgb, "val")
      if (!is.na(val) && val != "") out[[nm]] <- paste0("FF", toupper(val))
    } else {
      sys <- xml2::xml_find_first(node, ".//*[local-name()='sysClr']")
      if (!inherits(sys, "xml_missing")) {
        val <- xml2::xml_attr(sys, "lastClr")
        if (!is.na(val) && val != "") out[[nm]] <- paste0("FF", toupper(val))
      }
    }
  }
  
  out
}

apply_tint <- function(argb, tint) {
  if (is.na(argb) || argb == "" || is.na(tint) || tint == 0) return(argb)
  
  if (stringr::str_detect(argb, "^[0-9A-F]{6}$")) argb <- paste0("FF", argb)
  if (!stringr::str_detect(argb, "^[0-9A-F]{8}$")) return(argb)
  
  r <- strtoi(substr(argb, 3, 4), 16L)
  g <- strtoi(substr(argb, 5, 6), 16L)
  b <- strtoi(substr(argb, 7, 8), 16L)
  
  adj <- function(x) {
    if (tint < 0) x <- x * (1 + tint) else x <- x + (255 - x) * tint
    x <- round(x)
    x <- max(0, min(255, x))
    x
  }
  
  sprintf("FF%02X%02X%02X", adj(r), adj(g), adj(b))
}

# styles.xml: xf (0-based) -> цвет шрифта ARGB
read_xf_font_rgb <- function(xlsx_path) {
  listed <- zip_list_names(xlsx_path)
  if (!("xl/styles.xml" %in% listed)) return(character(0))
  
  st <- tryCatch(xml2::read_xml(unz(xlsx_path, "xl/styles.xml")), error = function(e) NULL)
  if (is.null(st)) return(character(0))
  
  theme_map <- read_theme_scheme(xlsx_path)
  scheme_names <- c("lt1","dk1","lt2","dk2",
                    "accent1","accent2","accent3","accent4","accent5","accent6",
                    "hlink","folHlink")
  
  font_nodes <- xml2::xml_find_all(st, ".//*[local-name()='fonts']/*[local-name()='font']")
  font_rgb <- vapply(font_nodes, function(fn) {
    col <- xml2::xml_find_first(fn, ".//*[local-name()='color']")
    if (inherits(col, "xml_missing")) return(NA_character_)
    
    rgb <- xml2::xml_attr(col, "rgb")
    if (!is.na(rgb) && rgb != "") return(toupper(rgb))
    
    th <- suppressWarnings(as.integer(xml2::xml_attr(col, "theme")))
    tint <- suppressWarnings(as.numeric(xml2::xml_attr(col, "tint")))
    
    if (!is.na(th)) {
      nm <- scheme_names[th + 1L]
      base <- theme_map[[nm]]
      if (!is.null(base)) return(apply_tint(base, tint))
    }
    
    NA_character_
  }, character(1))
  
  xf_nodes <- xml2::xml_find_all(st, ".//*[local-name()='cellXfs']/*[local-name()='xf']")
  if (length(xf_nodes) == 0) return(character(0))
  
  vapply(xf_nodes, function(xf) {
    fid <- suppressWarnings(as.integer(xml2::xml_attr(xf, "fontId")))
    if (is.na(fid)) return(NA_character_)
    idx <- fid + 1L # fontId 0-based
    if (idx < 1L || idx > length(font_rgb)) return(NA_character_)
    font_rgb[[idx]]
  }, character(1))
}

# Основная функция: зелёный по стилю ячейки ФИО или стилю строки
calc_death_green_xml <- function(xlsx_path, sheet_name, fio_col_idx, kept_rows_excel) {
  out <- rep(0L, length(kept_rows_excel))
  if (is.na(fio_col_idx) || length(kept_rows_excel) == 0) return(out)
  
  sheet_xml <- get_sheet_xml_path(xlsx_path, sheet_name)
  if (is.na(sheet_xml)) return(out)
  
  xf_rgb <- read_xf_font_rgb(xlsx_path)
  if (length(xf_rgb) == 0) return(out)
  
  col_letters <- excel_col_letters(fio_col_idx)
  if (is.na(col_letters)) return(out)
  
  sh <- tryCatch(xml2::read_xml(unz(xlsx_path, sheet_xml)), error = function(e) NULL)
  if (is.null(sh)) return(out)
  
  row_nodes <- xml2::xml_find_all(sh, ".//*[local-name()='sheetData']/*[local-name()='row']")
  if (length(row_nodes) == 0) return(out)
  
  row_nums <- suppressWarnings(as.integer(xml2::xml_attr(row_nodes, "r")))
  names(row_nodes) <- as.character(row_nums)
  
  rgb_vec <- rep(NA_character_, length(kept_rows_excel))
  
  for (i in seq_along(kept_rows_excel)) {
    rnum <- kept_rows_excel[i]
    rn <- row_nodes[[as.character(rnum)]]
    if (is.null(rn)) next
    
    # стиль строки (xf index, 0-based)
    row_s <- suppressWarnings(as.integer(xml2::xml_attr(rn, "s")))
    
    # стиль ячейки ФИО (xf index, 0-based)
    ref <- paste0(col_letters, rnum)
    cell_node <- xml2::xml_find_first(rn, sprintf(".//*[local-name()='c' and @r='%s']", ref))
    cell_s <- suppressWarnings(as.integer(xml2::xml_attr(cell_node, "s")))
    
    xf <- cell_s
    if (is.na(xf)) xf <- row_s
    if (is.na(xf)) xf <- 0L
    
    idx1 <- xf + 1L
    if (idx1 >= 1L && idx1 <= length(xf_rgb)) rgb_vec[i] <- xf_rgb[[idx1]]
  }
  
  as.integer(vapply(rgb_vec, is_green_rgb, logical(1)))
}
