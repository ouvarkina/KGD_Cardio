local({
  #====НАСТРОЙКА================================================================

  # Проверка, установка при необходимости и загрузка пакета
  install_and_load <- function(package, version = NULL) {
    if (!requireNamespace(package, quietly = TRUE)) {
      if (!is.null(version)) {
        remotes::install_version(package, version = version)
      } else {
        install.packages(package)
      }
    } else if (!is.null(version) && packageVersion(package) != version) {
      remotes::install_version(package, version = version)
    }
    library(package, character.only = TRUE)
  }

  # S3 класс для проектов
  create_project <- function(key, study_id, book_id, name, observations_count) {
    obj <- list(
      key = key,
      study_id = study_id,
      name = name,
      book_id = book_id,
      observations_count = observations_count
    )
    class(obj) <- "project"
    return(obj)
  }

  #====ДАННЫЕ===================================================================

  # Список используемых пакетов конкретных версий
  install_and_load("remotes")
  install_and_load("tidyverse") # для аккуратности ухоженности и наверно свежести зимы
  install_and_load("ggbreak") # Для добавления разрывов на оси
  install_and_load("scales") # for date breaks
  install_and_load("gtsummary") # Для красивой таблицыц
  install_and_load("ggsurvfit") # Для графиков выживаемости
  install_and_load("ggpmisc")
  install_and_load("tidycmprsk") # Для cuminc
  install_and_load("ggthemes")
  install_and_load('npreg')
  install_and_load('Hmisc') # Для HD квантилей
  install_and_load("flextable") # Для таблиц
  install_and_load("ftExtra") # Расширение флекстабле
  install_and_load("kableExtra") # Расширение knit'овских таблиц
  install_and_load("magrittr")
  install_and_load("plotly")
  install_and_load("moments")
  install_and_load("kableExtra")
  install_and_load("writexl")
  install_and_load("readxl")
  install_and_load("metan") # для коррелограммы
  install_and_load("mice") # для вменения
  install_and_load("zoo") #, version = "1.8-12") # если критично, указываем точную версию, и она поставится через Rtools
  install_and_load("ggbreak")
  # Список доступных проектов
  # Параметры:
  # * Просто код исследования;
  # * Код исследования для внешних источников (студентов, например);
  # * Идентификатор книги google из адресной строки;
  # * Красивое человекочитаемое название, обычно на русском;
  # * Количество столбцов на каждый кейс (нужно, чтобы корректно протянуть id пациента, остальные протяжки уже будут зависить от этого)
  projects <<- list(
    P_SHARED = create_project('shared', 'shared', '1hghWIKAxcru8lRinsVFQTmP54GwJ0TtnnsyF4d_C-Bk', 'shared', 1),
    P_21EF05 = create_project('21ef05', "L_mSep_2dC_11", '1q72LTSMFnw-bmzL0uBD6BL8COYON_8j-b2Y-oWGTf4Y', 'Акушерский сепсис', 6),
    P_22EF02 = create_project('22ef02', "P_22EF02", '11bzoeYUQ7C7_oh-ReZcDsWA5ysvMDW3lnn7nvpceSyw', 'ЛАССО НЕО', 6),
    P_22EF03 = create_project('22ef03', "P_22EF03", '1WwR01IPwPyG0lGntmwziOF6RDM6j4cN_vfLCCELg4_M', 'Панкреатит Демихово', 6),
    P_22EF08 = create_project('22ef08', "CL_Sep_3C_6", '1jOolYgMPluI0ikfa_nZUbWidRhhli8ftWfmHtz_-KGE', 'Сепсис Юдина', 2),
    P_23EF02 = create_project('22ef03', "P_23EF02", '13ijyiXgbhe3lzP01-mYpLX_XtArPAWLVfj_iuQsPCvc', 'ОП Краснодар', 6),
    P_COVID = create_project( '22ef03', "P_COVID", '1aswkDS-MVHWJyLDUQJFD77TivM6obRtmGFvODyEk8Z0', 'Ковид-19', 10),
    P_KZ = create_project('22ef03', "L_Sep_2dC_7", '1eTu154W-7csLyXxfvQpJYNjvX5moEQZ-9imAkivCgS8', 'Казахстан', 5),
    P_24EF01 = create_project('24ef01', 'P_24EF01', '1cpp02lmjq0XGDdjUn2rrZZGprY4T1CJUGAQLh4OyviU', 'Псориаз', 7),
    P_24EFv01 = create_project('24EFv01', "P_24EFv01", '1c7iqLLL88cwXB0iJivzwBw5IjKVGTFNqd6253PiJHhk', 'Симутис послеоперационная печень', 8),
    P_23EF02 = create_project('23ef02', "P_23EF02", '13ijyiXgbhe3lzP01-mYpLX_XtArPAWLVfj_iuQsPCvc', 'ОП Краснодар', 6),
    P_YOUNG = create_project('young', "P_YOUNG", '1SlzzxRAUtJxgZp1ISSpur59t-LJAIBNSmWIHWL4wotk', 'Неонатальная серия', 6),
    P_LASSO = create_project("lasso", "L_Sep_2R_4", "10l4ivzej4Se-FnKtabBJ0zg0gqXABvPrWufjdiP40ZY", "ЛАССО", 5),
    P_FIRST = create_project("first", "L_Sep_1C_1", "1mYWXyRed8-9LI27tXWAVtMrbtkfAl5sFn85FNACgqHI", "Первый сепсис", 22),
    P_TEST = create_project("test", "Test", "1EwziqjlZ2yVT1DYIvBYR5firPdxwbQvLcdg5jcY8l1w", "Test", 5),
    P_24EF04 = create_project("24EF04", "P_24EF04", "1rSo6MgDmf4nXU6ujccIjPL658nchmFXgjWfXFJVPquk", "Кардиохирургия с ИК", 7),
    P_24EF03 = create_project("24ef03", "P_24EF03", "1TLoqV_sjAVOfof8wuIpSDuI-AVeDymPto0GucXurmyo", "Клиренс АБ", 6)
  )
})

# ====ФУНКЦИИ ЗАГРУЗКИ ДАННЫХ===================================================

load.sheet <- function(project, sheet_name) {
  read.csv(
    paste0(
      "https://docs.google.com/spreadsheets/d/",
      project$book_id,
      "/gviz/tq?tqx=out:csv&headers=1&sheet=",
      URLencode(sheet_name)
    ),
    na.strings = c("", "-") # Значения, которые будут прочитаны как NA
  ) |> select(-starts_with("X")) # так обрезаются пустые строки
}

load.variables <- function(project) {
  variables <- load.sheet(project, "variables")

  factors <- c("DEPENDENCY_TYPE", "VAR_TYPE", "VALUE_TYPE", "LIST_TYPE")
  numerics <- c("PRECISION", "SORT")

  variables <- variables |>
    mutate(across(all_of(factors), as.factor)) |>
    mutate(across(all_of(numerics), as.numeric))

  # to_rds(variables, "raw.variables")

  return(variables)
}

load.raw <- function(project) {

  # Исходные данные
  raw <- load.sheet(project, "DATA")
  raw <- raw |> rename_with(~ gsub("_\\.", "_%", .x)) # из-за % в названиях переменных

  # Проверяем, что хватает настроек для обработки
  if (isFALSE("system_time" %in% colnames(raw))) stop("В исходных данных не определен служебный параметр: system_time")
  if (isFALSE("system_patient_id" %in% colnames(raw))) stop("В исходных данных не определен служебный параметр: system_patient_id")
  if (isFALSE("system_patient_death" %in% colnames(raw))) stop("В исходных данных не определен служебный параметр: system_patient_death")
  #if (isFALSE("system_patient_disabled" %in% colnames(raw))) stop("В исходных данных не определен служебный параметр: system_patient_disabled")
  if (isFALSE("system_patient_death_time" %in% colnames(raw))) stop("В исходных данных не определен служебный параметр: system_patient_death_time")

  # Выставляем ключи пациентов на все точки времени, обрезаем лишние строки
  raw <- raw |>
    mutate(system_patient_id = fill_na_with_maxgap(system_patient_id, maxgap = project$observations_count - 1)) |>
    filter(!is.na(system_patient_id))

  return(raw)
}

ensure_valid_types <- function(dataset, variables) {

  # Логические переменные
  bool_vars <- variables |> filter(VALUE_TYPE == "boolean") |> pull(VAR)
  dataset <- dataset |>
    mutate(across(
      .cols = all_of(bool_vars),
      .fns = ~ case_when(
        is.na(.) ~ NA,
        . %in% c(1, "1", "TRUE", TRUE, "да", "yes") ~ TRUE,
        TRUE ~ FALSE
      )
    ))

  # Факторы
  factor_vars <- variables |> filter(VALUE_TYPE == "factor") |> pull(VAR)
  dataset <- dataset |>
    mutate(across(
      .cols = all_of(factor_vars),
      .fns = factor
    ))

  # Даты
  date_vars <- variables |> filter(VALUE_TYPE == "date") |> pull(VAR)
  dataset <- dataset |>
    mutate(across(
      .cols = all_of(date_vars),
      .fns = ~ case_when(
        grepl(":", .) ~ as.POSIXct(., format = "%d.%m.%Y %H:%M", optional = TRUE),
        grepl("\\.", .) ~ as.POSIXct(., format = "%d.%m.%Y", optional = TRUE),
        TRUE ~ as.POSIXct(NA_character_)
      )
    ))

  # Числа
  numeric_vars <- variables |> filter(VALUE_TYPE == "numeric") |> pull(VAR)
  dataset <- dataset |>
    mutate(across(
      .cols = all_of(numeric_vars),
      .fns = ~ case_when(
        grepl("%", .) ~ as.numeric(gsub("%", "", .)),
        grepl("^>", .) ~ as.numeric(gsub(">", "", .)) + 0.1,
        grepl("^<", .) ~ as.numeric(gsub("<", "", .)) - 0.1,
        TRUE ~ as.numeric(.)
      )
    ))

  # Строки
  string_vars <- variables |> filter(VALUE_TYPE == "string") |> pull(VAR)
  dataset <- dataset |>
    mutate(across(
      .cols = all_of(string_vars),
      .fns = as.character
    ))

  return(dataset)
}

get_static_vars <- function(variables) {
  variables |>
    filter(VAR_TYPE == "property", VAR != "system_patient_id") |>
    pull(VAR)
}

get_dynamic_vars <- function(variables) {
  variables |>
    filter(VAR_TYPE == "dynamic", VAR != "system_time", VAR != "real_time") |>
    filter(VALUE_TYPE %in% c("numeric")) |>
    pull(VAR)
}

find_na_points <- function(dataset, variables) {

  # Получаем список динамических переменных
  dynamics <- get_dynamic_vars(variables)

  # Находим временные точки, где все динамические переменные = NA
  na_points <- dataset |>
    group_by(system_time) |>
    summarise(
      all_na = all(across(all_of(dynamics), is.na), na.rm = TRUE),
      .groups = "drop"
    ) |>
    filter(all_na) |>
    select(system_time)

  return(na_points)
}

fill_static_vars <- function(dataset, static_vars) {
  dataset |>
    group_by(system_patient_id) |>
    fill(all_of(static_vars), .direction = "updown") |>
    ungroup()
}

fill_values_real_time <- function(dataset, dynamic_vars, no_imputation_locf = FALSE, target_times = NULL) {

  # определяем временные точки
  if (is.null(target_times)) {
    # если целевые точки не заданы, используем все уникальные из данных
    all_times <- unique(dataset$real_time) |> sort()
    filter_times <- all_times
  } else {
    # объединяем нужные точки с существующими
    existing_times <- unique(dataset$real_time)
    all_times <- union(existing_times, target_times) |> sort()
    filter_times <- target_times
  }

  # тянем
  dataset |>
    group_by(system_patient_id) |>
    tidyr::complete(real_time = all_times) |>
    mutate(system_time = real_time) |>
    mutate(across(
      all_of(dynamic_vars),
      ~ {
        x <- na.approx(., x = system_time, na.rm = FALSE)
        if (!no_imputation_locf) {
          x <- na.locf(x, na.rm = FALSE, fromLast = TRUE)
          x <- na.locf(x, na.rm = FALSE, fromLast = FALSE)
        }
        x
      }
    )) |>
    filter(system_time %in% filter_times) |>
    ungroup()
}

fill_values_system_time <- function(dataset, dynamic_vars, no_imputation_locf = FALSE) {
  dataset |>
    group_by(system_patient_id) |>
    mutate(across(
      all_of(dynamic_vars),
      ~ {
        x <- na.approx(., x = system_time, na.rm = FALSE)
        if (!no_imputation_locf) {
          x <- na.locf(x, na.rm = FALSE, fromLast = TRUE)
          x <- na.locf(x, na.rm = FALSE, fromLast = FALSE)
        }
        x
      }
    )) |>
    ungroup()
}

drop_disabled <- function(dataset) {

  if (!("system_patient_disabled_why" %in% names(dataset))) {
    dataset[["system_patient_disabled_why"]] <- NA
  }

  disabled <- unique(
    dataset |>
      filter(system_patient_disabled == TRUE) |>
      select(
        system_patient_id,
        system_patient_disabled_why,
      )
  )

  dataset <- dataset |> filter(!(system_patient_id %in% disabled$system_patient_id))

  if (nrow(disabled) > 0) {
    print(paste0("Исключены из обработки: ", paste(disabled$system_patient_id, disabled$system_patient_disabled_why, sep = " - ", collapse = ", ")))
  }

  return(dataset)
}

drop_values_if_dead <- function (dataset, variables) {
  dynamics <- get_dynamic_vars(variables)
  dataset |>
    mutate(
      system_patient_death = na_recode(system_patient_death, FALSE),
      system_patient_death_time = na_recode(system_patient_death_time, 9999999)
    ) |>
    mutate(is_alive = system_patient_death != TRUE | system_time < system_patient_death_time) |>
    mutate(across(
      all_of(dynamics),
      ~ if_else(is_alive, .x, NA_real_)
    ))
}

load.dynamic <- function(
    project,
    include_dead = FALSE, # LOCF для умерших после времени смерти
    keep_disabled = FALSE, # сохранить исключенных пациентов в датасете
    keep_NA_if_no_value_at_time = FALSE, # оставлять NA, если нет ни одного исходного значения на эту точку времени
    no_imputation = FALSE, # не выполнять операций по протяжке или интерполяции,
    no_imputation_dynamic = FALSE, # не выполнять операций по протяжке или интерполяции (только динамические переменные)
    no_imputation_locf = FALSE, # не делать протяжку первого и последнего значений (заполнять только промежутки)
    use_real_time = FALSE, # использовать для расчета состояния смерти и интерполяции переменную real_time
    target_times = NA # заранее определенные точки. Их количество должно совпадать с количеством оригинальных
) {
  
  # Загружаем исходные данные, выставляем ключи по пациентам и проверяем наличие нужных переменных
  raw <- load.raw(project)
  variables <- load.variables(project)

  # Убираем лишние переменные из обработки
  variables <- variables |> filter(VAR %in% names(raw))
  static_vars <- get_static_vars(variables)
  dynamic_vars <- get_dynamic_vars(variables)
  V <- namedList(variables$VAR)

  # Обрабатываем различные типы данных
  raw <- ensure_valid_types(raw, variables)

  # Заполняем пропуски
  if (!no_imputation) {

    with_na <- find_na_points(raw, variables) # специально до заполнения

    # растягиваем статику
    raw <- fill_static_vars(raw, static_vars)

    if (!no_imputation_dynamic) {
      if (use_real_time) {
        raw <- fill_values_real_time(raw, dynamic_vars, no_imputation_locf, target_times)
        raw <- fill_static_vars(raw, static_vars) # из-за изменения шкалы времени, дополнительно перетягиваем статику
      }
      else {
        raw <- fill_values_system_time(raw, dynamic_vars, no_imputation_locf)
      }
    }

    if (keep_NA_if_no_value_at_time & nrow(with_na) > 0) {
      print(paste0("Точки без реальных данных: ", paste(with_na, collapse = ", ")))
      raw[raw$system_time %in% with_na, dynamic_vars] <- NA
    }

    if (!include_dead) {
      raw <- drop_values_if_dead(raw, variables)
    }
  }

  # Исключаем исключенных, если не попросили оставить
  if (!keep_disabled) {
    raw <- drop_disabled(raw)
  }

  # Добавляем название исследования как колонку
  raw <- raw |> mutate(study_id = project$study_id)

  return(raw)
}

# Новая функция для интерполяции
interpolate_dynamic_values <- function(
    dataset,
    id_var,
    time_var,
    dynamic_vars,
    death_time_var = NULL,
    no_imputation_locf = FALSE
) {
  # Сохраняем оригинальные системные колонки если они существуют
  orig_system_id <- NULL
  orig_system_time <- NULL
  orig_system_death <- NULL

  if ("system_patient_id" %in% names(dataset)) {
    orig_system_id <- dataset$system_patient_id
  }
  if ("system_time" %in% names(dataset)) {
    orig_system_time <- dataset$system_time
  }
  if (!is.null(death_time_var) && "system_patient_death_time" %in% names(dataset)) {
    orig_system_death <- dataset$system_patient_death_time
  }

  # Копируем переданные значения в системные колонки
  dataset$system_patient_id <- dataset[[id_var]]
  dataset$system_time <- dataset[[time_var]]
  if (!is.null(death_time_var)) {
    dataset$system_patient_death_time <- dataset[[death_time_var]]
  }

  # Выполняем интерполяцию
  dataset <- fill_values_system_time(dataset, dynamic_vars, no_imputation_locf)

  # Восстанавливаем оригинальные системные колонки
  if (!is.null(orig_system_id)) {
    dataset$system_patient_id <- orig_system_id
  } else {
    dataset$system_patient_id <- NULL
  }

  if (!is.null(orig_system_time)) {
    dataset$system_time <- orig_system_time
  } else {
    dataset$system_time <- NULL
  }

  if (!is.null(orig_system_death)) {
    dataset$system_patient_death_time <- orig_system_death
  } else if (!is.null(death_time_var)) {
    dataset$system_patient_death_time <- NULL
  }

  return(dataset)
}

# ====ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ===================================================

# Модификация базовой функции is.nan, чтоб она работала с датафрэймом
# https://stackoverflow.com/questions/18142117/how-to-replace-nan-value-with-zero-in-a-huge-data-frame
is.nan.data.frame <- function(x) {
  do.call(cbind, lapply(x, is.nan))
}

# Функция locf по определенному количеству значений
fill_na_with_maxgap <- function(vec, maxgap) {
  # Заполняем пропуски вперед
  filled_forward <- na.locf(vec, na.rm = FALSE)

  # Заполняем пропуски назад
  filled_backward <- na.locf(vec, fromLast = TRUE, na.rm = FALSE)

  # Объединяем результаты, учитывая maxgap
  result <- vec
  for (i in seq_along(vec)) {
    if (is.na(vec[i])) {
      forward_gap <- sum(is.na(vec[max(1, i - maxgap):i]))
      backward_gap <- sum(is.na(vec[i:min(length(vec), i + maxgap)]))

      if (forward_gap <= maxgap) {
        result[i] <- filled_forward[i]
      } else if (backward_gap <= maxgap) {
        result[i] <- filled_backward[i]
      }
    }
  }

  return(result)
}

# Обертка, загружающая объект в .rds файл
to_rds <- function (obj, name = "", raw = FALSE) {
  folder <- if_else(raw, "data/raw", "data")
  dir.exists(folder) || dir.create(folder)
  saveRDS(obj, file = paste0(folder, "/", name, ".rds"))
#  return(obj)
}

# Чтение объекта из .rds файл
rds <- function (name, raw = FALSE) {
  obj <- NULL
  folder <- if_else(raw, "data/raw", "data")
  tryCatch(expr = {
    obj <- readRDS(paste0(folder, "/", name, ".rds"))
  }, error = function (e) {
    print(paste0("Набор данных ", name, " не прочитан"))
  })
  return(obj)
}

namedList <- function (column) {
  obj <- as.vector(column)
  obj <- setNames(as.list(obj), obj)
  return(obj)
}

# Выделение выбранных переменных из общего списка, сохряняя порядок
get_variables <- function (variables, names) {
  selected_variables <- data.frame(VAR = names) %>%
    inner_join(variables, by = 'VAR')
  return(selected_variables)
}

# Функция замены NA на какое-то значение
na_recode <- function(x, na_label = '-') {
  # Векторизированная версия if_else сохраняет тип данных вектора. if_else - функция из tidyverse
  x <- if_else(is.na(x), na_label, x)
}

# Функция, принимающая dataframe или gt_summary и возвращающая kable
to_kable <- function(table, font_size = 14, save_as = NA, ...) {

  class_name <- class(table)[1]
  is_gt <- class_name == "tbl_summary" | class_name == "gtsummary"
  is_stack <- class_name == "tbl_stack"

  save_table <- function (table) {
    if (!is.na(save_as)) {
      dir.exists("tables") || dir.create("tables")
      write_xlsx(table, paste0("tables/", save_as, ".xlsx"),
                 col_names = TRUE,
                 format_headers = TRUE)
    }
  }

  # Вот без этой опции QUATTRO перехватывает сырой вывод kable html и сам форматирует, как считает нужным
  # https://examples.quarto.pub/opt-out-table-processing-kable/
  # Глобальное задание в YAML или настройке библиотеки не работает
  if (is_gt) {
    save_table(as.data.frame(table))
    table <- as_kable_extra(table, table.attr = "quarto-disable-processing=true", ...)
  } else if (is_stack) {
    save_table(as.data.frame(table))
    table <- as_kable_extra(table, table.attr = "quarto-disable-processing=true", ...)
  } else {
    save_table(table)
    table <- kbl(table, table.attr = "quarto-disable-processing=true", ...)
  }

  table %>%
    kable_styling(
      bootstrap_options = c("striped", "hover", "condensed"),
      fixed_thead = TRUE,
      font_size = font_size,
      full_width = FALSE) %>%
    kable_classic() %>%
    row_spec(0, bold = TRUE)
}



get_var <- function(variables, var) {
  settings <- variables |> filter(VAR == var)
  obj <- setNames(settings |> first(), names(settings))
  return(obj)
}


var_units_glue <- function(var, units) {
  result <- list()
  for (i in seq_along(var)) {
    result[[var[i]]] <- ifelse(is.na(units[i]), var[i], paste0(var[i], ", ", units[i]))
  }
  return(result)
}

# Вычисление разницы между датами в сутках
days <- function (date_end, date_start) {
  as.numeric(difftime(date_end, date_start), units = "days")
}


# Функция расчета pSOFA
add_pSOFA_score <- function(df) {
  temp <- df |>
    mutate(
      # Создание возрастной категории
      age_based = case_when(
        is.na(age_m) ~ NA_real_,
        age_m <  1   ~ 1,
        age_m <  12  ~ 2,
        age_m <  24  ~ 3,
        age_m <  60  ~ 4,
        age_m <  144 ~ 5,
        age_m <  216 ~ 6,
        age_m >= 216 ~ 7,
        TRUE         ~ NA_real_
      ),

      respiratory_PF = case_when(
        is.na(PF_ratio) ~ NA_real_,
        PF_ratio >= 400 ~ 0,
        PF_ratio >= 300 ~ 1,
        PF_ratio >= 200 ~ 2,
        PF_ratio >= 100 ~ 3,
        PF_ratio < 100  ~ 4,
        TRUE            ~ NA_real_
      ),
      respiratory_SF = case_when(
        is.na(SF_ratio) ~ NA_real_,
        SF_ratio > 292  ~ 0,
        SF_ratio >= 264 ~ 1,
        SF_ratio >= 221 ~ 2,
        SF_ratio >= 148 ~ 3,
        SF_ratio < 148  ~ 4,
        TRUE            ~ NA_real_
      ),
      respiratory = case_when(
        !is.na(PF_ratio) ~ respiratory_PF,
        !is.na(SF_ratio) ~ respiratory_SF,
        TRUE ~ NA_real_
      ),

      coagulation = case_when(
        is.na(PLT) ~ NA_real_,
        PLT >= 150 ~ 0,
        PLT >= 100 ~ 1,
        PLT >= 50  ~ 2,
        PLT >= 20  ~ 3,
        PLT <  20  ~ 4,
        TRUE       ~ NA_real_
      ),

      bilirubin_DL = if_else(is.na(bilirubin_total), NA_real_, round(bilirubin_total / 17.1, 1)),
      liver = case_when(
        is.na(bilirubin_DL) ~ NA_real_,
        bilirubin_DL >= 12  ~ 4,
        bilirubin_DL >= 6   ~ 3,
        bilirubin_DL >= 2   ~ 2,
        bilirubin_DL >= 1.2 ~ 1,
        bilirubin_DL < 1.2  ~ 0,
        TRUE                ~ NA_real_
      ),

      cardiovascular_MAP = case_when(
        is.na(age_m) | is.na(MAP) ~ NA_real_,
        VIS2020 != 0             ~ 0,
        MAP >= c(46, 55, 60, 62, 65, 67, 70)[age_based] ~ 0,
        MAP <  c(46, 55, 60, 62, 65, 67, 70)[age_based] ~ 1,
        TRUE                     ~ NA_real_
      ),
      cardiovascular_VIS = case_when(
        is.na(VIS2020) ~ NA_real_,
        VIS2020 == 0   ~ 0,
        VIS2020 <= 5   ~ 2,
        VIS2020 <= 15  ~ 3,
        VIS2020 >  15  ~ 4,
        TRUE           ~ NA_real_
      ),
      cardiovascular = pmax(cardiovascular_MAP, cardiovascular_VIS, na.rm = TRUE),

      neurologic = case_when(
        is.na(GCS) ~ NA_real_,
        GCS >= 15  ~ 0,
        GCS >= 13  ~ 1,
        GCS >= 10  ~ 2,
        GCS >= 6   ~ 3,
        GCS < 6    ~ 4,
        TRUE       ~ NA_real_
      ),


      creatinine_DL = if_else(is.na(creatinine), NA_real_, round(creatinine / 88.4, 1)),
      renal = case_when(
        is.na(creatinine_DL) ~ NA_real_,
        creatinine_DL < c(0.8, 0.3, 0.4, 0.6, 0.7, 1.0, 1.2)[age_based] ~ 0,
        creatinine_DL < c(1.0, 0.5, 0.6, 0.9, 1.1, 1.7, 2.0)[age_based] ~ 1,
        creatinine_DL < c(1.2, 0.8, 1.1, 1.6, 1.8, 2.9, 3.5)[age_based] ~ 2,
        creatinine_DL < c(1.6, 1.2, 1.5, 2.3, 2.6, 4.2, 5.0)[age_based] ~ 3,
        creatinine_DL >= c(1.6, 1.2, 1.5, 2.3, 2.6, 4.2, 5.0)[age_based] ~ 4,
        TRUE ~ NA_real_
      )
    ) |>
    # Заполнение NA значений средними по группам
    #TODO: что делать если в группе все NA и расчет среднего даст NA
    group_by(system_time) |>
    mutate(
      across(
        c(respiratory, cardiovascular, renal, coagulation, liver, neurologic),
        ~ ifelse(is.na(.), round(mean(., na.rm = TRUE)), .)
      )
    ) |>
    ungroup() |>
    mutate(
      # Расчитываем pSOFA_calc
      pSOFA_calc = rowSums(across(c(respiratory, cardiovascular, renal, coagulation, liver, neurologic), ~ replace_na(., 0)))
    )

  df <- df |>
    mutate(respiratory = temp$respiratory,
           cardiovascular = temp$cardiovascular,
           renal = temp$renal,
           coagulation = temp$coagulation,
           liver = temp$liver,
           neurologic = temp$neurologic,
           pSOFA_calc = temp$pSOFA_calc)

  return(df)
}


# # Пример использования функции
# dataset <- add_pSOFA_score(dataset)


# Функция для расчета ОЦК у детей до 18 лет, Источник: https://sci-hub.ru/10.1097/MAT.0b013e3181d0c28d
blood_volume_pediatric <- function(dataset) {
  dataset <- dataset %>%
    mutate(
      coefficient = case_when(
        age_m <= (1/31) ~ 87.7,   # Доношенные новорожденные (0–24 ч) ≈ 1/30 месяца
        age_m > (1/31) & age_m <= (7/31) ~ 91.7,  # Доношенные (1–7 дней) ≈ 7/30 месяца
        age_m > (7/31) & age_m <= 1 ~ 86.0,  # Доношенные (7–30 дней) = 1 месяц
        age_m > 1 & age_m <= 6 ~ 84.0,  # Дети 1–6 месяцев
        age_m > 6 & age_m <= 24 ~ 75.4,  # Дети 6–24 месяцев (0.5–2 года)
        age_m > 24 & age_m <= 60 ~ 75.7,  # Дети 2–5 лет (24–60 месяцев)
        age_m > 60 & age_m <= 144 ~ 74.9,  # Дети 5–12 лет (60–144 месяцев)
        age_m > 144 & age_m <= 216 ~ 71.4,  # Дети 12–18 лет (144–216 месяцев)
        TRUE ~ 80.0  # Среднее значение для всех остальных
      ),
      pediatric_blood_volume = weight * coefficient
    ) %>%
    select(-coefficient)

  return(dataset)
}

# # Пример использования функции
# dataset <- blood_volume_pediatric(dataset)

extract_list <- function(df, df_variables, list_name) {
  tryCatch(
    expr = {
      index_name <- ifelse(list_name == "shocks", "shock", list_name)

      # Сначала получаем переменные листа, которые существуют в этом датасете
      exist_vars <- intersect(
        df_variables |>
          filter(LIST_TYPE == list_name) |>
          pull(VAR),
        colnames(df)
      )

      # Теперь по ним выбираем нужные данные, используя индекс листа как основной ключ
      df_list <- df |>
        select(
          system_patient_id,
          all_of(exist_vars)
          # Cюда добавлять те переменные из общего датасета, которые нужны для расчетов
        ) |>
        filter(!is.na(!!sym(paste0(index_name, "_index"))))

      to_rds(df_list, list_name)
    },
    error = function(e) {
      print(paste0("Данные по ", list_name, " не прочитаны:", e$message))
    }
  )}

extract_list2 <- function(df, df_variables, list_name) {
  tryCatch(
    expr = {
      index_name <- ifelse(list_name == "shocks", "shock", list_name)

      # Сначала получаем переменные листа, которые существуют в этом датасете
      exist_vars <- intersect(
        df_variables |>
          filter(LIST_TYPE == list_name) |>
          pull(VAR),
        colnames(df)
      )

      # Теперь по ним выбираем нужные данные, используя индекс листа как основной ключ
      df_list <- df |>
        select(
          system_patient_id,
          all_of(exist_vars)
          # Cюда добавлять те переменные из общего датасета, которые нужны для расчетов
        ) |>
        filter(!is.na(!!sym(paste0(index_name, "_index"))))

      return(df_list)
    },
    error = function(e) {
      print(paste0("Данные по ", list_name, " не прочитаны:", e$message))
    }
  )}

# Функция для получения минимальной даты
get_min_date <- function(dates) {
  # Преобразуем данные в POSIXct
  parsed_dates <- as.POSIXct(dates, format = "%Y-%m-%d %H:%M:%S %z", tz = "UTC")

  # Фильтруем некорректные даты (например, до 1900 года)
  valid_dates <- parsed_dates[!is.na(parsed_dates) & parsed_dates > as.Date("1900-01-01")]

  if (length(valid_dates) == 0) {
    warning("Нет корректных дат для расчета.")
    return(NA)
  }
  format(min(valid_dates), "%Y.%m.%d")
}

# Функция для получения максимальной даты
get_max_date <- function(dates) {
  # Преобразуем данные в POSIXct
  parsed_dates <- as.POSIXct(dates, format = "%Y-%m-%d %H:%M:%S %z", tz = "UTC")

  # Фильтруем некорректные даты (например, до 1900 года)
  valid_dates <- parsed_dates[!is.na(parsed_dates) & parsed_dates > as.Date("1900-01-01")]

  if (length(valid_dates) == 0) {
    warning("Нет корректных дат для расчета.")
    return(NA)
  }
  format(max(valid_dates), "%Y.%m.%d")
}


# Функция для округления значений (надо перенести в source!)
round_values <- function(value, var_name, variables) {
  if (var_name %in% variables$VAR) {
    precision <- variables$PRECISION[variables$VAR == var_name]
    if (!is.na(precision)) {
      # Раз мы округляем значение, значит мы уже получаем не число форматированную строку
      return(as.character(round(value, digits = precision))) # signif для значащих разрядов
    }
  }
  return(as.character(round(value, digits = 3)))
}

#' Функция для рисования инцидентных графиков
#' @export
incidence_plot <- function(cuminc,
                           failcode = 'recovery',
                           state_lbl = 'ICU',
                           time_lbl = 'Время дни',
                           two_figure = TRUE) {

  if (nrow(cuminc$tidy) <= 1) {
    # R сессия тупо падает если тут будет 0 событий
    return(print('ERR: nrow <= 1'))
  }

  # Проверка, что такой failcode есть в переданном объекте. Если его не будет, дальше возникнет ошибка
  if (!any(cuminc$tidy$outcome == failcode)) {
    return(print(paste0("ERR: failcode \"", failcode, "\" not found in outcome")))
  }

  if (n_distinct(cuminc$tidy$outcome) > 1) {
    plot <- cuminc %>%
      ggcuminc(outcome = failcode) +
      labs(
        title = paste(state_lbl, ":", failcode, sep = " "),
        x = time_lbl
      ) +
      ylim(c(0, 1)) +
      add_censor_mark() +
      theme(legend.position=c(0.7, 0.2))
  } else {
    plot <- cuminc %>%
      ggcuminc(outcome = failcode) +
      labs(
        title = paste(state_lbl, ":", failcode, sep = " "),
        x = time_lbl
      ) +
      ylim(c(0, 1)) +
      add_censor_mark()
  }

  return(plot)
}
# Функция для округления значений
round_values <- function(value, var_name, variables) {
  if (var_name %in% variables$VAR) {
    precision <- variables$PRECISION[variables$VAR == var_name]
    if (!is.na(precision)) {
      # Раз мы округляем значение, значит мы уже получаем не число форматированную строку
      return(as.character(round(value, digits = precision))) # signif для значащих разрядов
    }
  }
  return(as.character(round(value, digits = 3)))
}