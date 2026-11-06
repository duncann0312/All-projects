#!/usr/bin/env Rscript
# =============================================================================
# Policy renewal - IMPORT MODULE
# -----------------------------------------------------------------------------
# Reads the raw EXPIRING and CURRENT Excel exports (tab "risk data"), works out
# which columns hold the fields policy_renewal_analysis.R needs, recodes the
# policy-type values to the four standard types, and writes standardised CSVs
# ready for the analysis script.
#
#   1. Checks both files exist and finds the "risk data" tab (case-insensitive).
#   2. Auto-detects the header row (handles title rows above the table).
#   3. Scores every column against each field using its NAME (synonyms) and
#      CONTENT (reference pattern, uniqueness, numeric values) and picks the
#      best match. Required: policy_ref, master_ref, limit, gn_premium.
#      Optional: stack_code, deductible, participation, class_of_business.
#   4. Picks the best policy-type column (header + values) and suggests which
#      standard type each distinct value maps to.
#   5. Writes:
#        input/expiring.csv, input/current.csv    -> analysis inputs
#        input/import_field_mapping.csv           -> chosen column + runners-up
#        input/import_type_mapping.csv            -> type value suggestions
#
# Review the two mapping reports. If anything is wrong, fix it with
# imp_cfg$field_overrides / imp_cfg$type_col_override / imp_cfg$type_overrides
# and re-run. Then run policy_renewal_analysis.R with cfg$use_dummy <- FALSE.
#
# Needs package {readxl}. Demo mode also needs {openxlsx} or {writexl}.
# =============================================================================

# ---- 0. CONFIGURATION -------------------------------------------------------
imp_cfg <- list(
  # TRUE  = build two messy demo workbooks in input_dir and import those
  # FALSE = import your real files named below
  demo          = TRUE,

  input_dir     = "input",
  expiring_file = "expiring_export.xlsx",
  current_file  = "current_export.xlsx",
  sheet         = "risk data",          # matched ignoring case / spaces
  header_row    = NA,                   # NA = auto-detect, or the row number
  output_dir    = "input",              # where the standardised CSVs go

  min_score     = 0.6,                  # minimum match score to accept a column

  # Force a column for a field, e.g. list(policy_ref = "Risk Ref", limit = "Limit 100%")
  field_overrides   = list(),
  # Force the policy-type column, e.g. "Placement Basis"
  type_col_override = NA,
  # Force a category for raw type values, e.g. c("Facility Dec" = "binder", "Fac" = "open_market")
  # Categories: open_market, lineslip, block, binder
  type_overrides    = c(),

  # Standard type labels written to the CSVs - keep the same as
  # cfg$type_labels in policy_renewal_analysis.R
  type_labels = list(
    open_market = "Open Market",
    lineslip    = "Lineslip Declaration",
    block       = "Block Entry",
    binder      = "Binder Declaration"
  )
)

# ---- 1. FIELD DEFINITIONS ---------------------------------------------------
# syn  = strong header synonyms   weak = generic synonyms (scored at 80%)
# neg  = regex on the normalised header that halves the score
# kind = content test used to confirm the column
FIELD_SPEC <- list(
  policy_ref = list(required = TRUE, kind = "ref_unique",
    syn  = c("policyreference", "policyref", "policyno", "policynumber", "policyid",
             "riskreference", "riskref", "riskid", "risknumber",
             "certificatereference", "certificatenumber", "certificateref", "certref",
             "declarationreference", "declarationref", "decref", "contractreference",
             "uniquereference"),
    weak = c("reference", "ref", "umr"),
    neg  = "master|binder|parent|facility|group|broker|claim"),
  master_ref = list(required = TRUE, kind = "ref_master",
    syn  = c("masterreference", "masterref", "masterpolicy", "masterpolicyreference",
             "masterpolicynumber", "masterumr", "binderreference", "binderref",
             "binderumr", "bindernumber", "facilityreference", "facilityref",
             "parentreference", "parentref", "bindingauthorityreference"),
    weak = c("master", "binder", "facility", "parent", "umr"),
    neg  = "claim|broker"),
  limit = list(required = TRUE, kind = "amount",
    syn  = c("limit", "limit100", "policylimit", "limitofliability", "limitofindemnity",
             "suminsured", "limitamount", "limitusd", "aggregatelimit",
             "occurrencelimit", "eachlosslimit"),
    weak = c("exposure", "si"),
    neg  = "deductible|excess|attachment|retention|sublimit|currency|ccy|share|line|pct|rate"),
  gn_premium = list(required = TRUE, kind = "amount",
    syn  = c("gnpremium", "gnp", "gnwp", "gnpusd", "grossnetpremium",
             "grossnetwrittenpremium", "netpremium", "gnpremium100"),
    weak = c("premium", "writtenpremium", "grosspremium"),
    neg  = "tax|ipt|comm|brokerage|fee|levy|rate|pct|adjust"),
  stack_code = list(required = FALSE, kind = "text",
    syn  = c("stackcode", "stack", "stackref", "stackreference", "stackid", "stacknumber"),
    weak = character(), neg = NULL),
  deductible = list(required = FALSE, kind = "amount",
    syn  = c("deductible", "deductibleamount", "excess", "attachment", "attachmentpoint",
             "retention", "xs"),
    weak = character(), neg = "limit|premium"),
  participation = list(required = FALSE, kind = "share",
    syn  = c("participation", "participationpct", "signedline", "signedlinepct",
             "writtenline", "writtenlinepct", "ourshare", "ourline", "signedshare",
             "share", "line", "linesize", "linepct"),
    weak = character(), neg = "slip|limit|premium"),
  class_of_business = list(required = FALSE, kind = "text",
    syn  = c("classofbusiness", "class", "cob", "businessclass", "lineofbusiness", "lob",
             "classname", "majorclass", "subclass", "classcode"),
    weak = character(), neg = "type|basis")
)

TYPE_SYN  <- c("policytype", "risktype", "contracttype", "placementtype", "placementbasis",
               "placingbasis", "placementmethod", "methodofplacement", "businesstype",
               "typeofbusiness", "entrytype", "recordtype", "declarationtype")
TYPE_WEAK <- c("type", "basis", "method")

# Value patterns for each standard type (checked in this order, lower-case)
TYPE_PATTERNS <- list(
  lineslip    = "line\\s*-?\\s*slip|\\bls\\b|\\blsd\\b|\\bl/s\\b",
  binder      = "binder|binding|\\bbdx\\b|bordereau|coverholder|delegated|\\bdua\\b|\\bb/a\\b|\\bba\\b",
  block       = "\\bblock\\b",
  open_market = "open\\s*market|\\bom\\b|\\bo/m\\b|individual|direct|facultative|\\bfac\\b|stand\\s*-?\\s*alone"
)

# ---- 2. MATCHING HELPERS ----------------------------------------------------
norm <- function(x) gsub("[^a-z0-9]", "", tolower(as.character(x)))

best_syn <- function(n, syn) {
  if (!length(syn) || !nzchar(n)) return(0)
  max(vapply(syn, function(s) {
    if (n == s) 1
    else if (nchar(s) >= 3 && grepl(s, n, fixed = TRUE)) 0.75 + 0.2 * nchar(s) / nchar(n)
    else if (nchar(n) >= 3 && grepl(n, s, fixed = TRUE)) 0.6 + 0.2 * nchar(n) / nchar(s)
    else 0.7 * (1 - adist(n, s)[1, 1] / max(nchar(n), nchar(s)))   # typo tolerance
  }, numeric(1)))
}

name_score <- function(col, syn, weak = character(), neg = NULL) {
  n  <- norm(col)
  sc <- max(best_syn(n, syn), 0.8 * best_syn(n, weak))
  if (!is.null(neg) && grepl(neg, n)) sc <- sc * 0.5
  sc
}

NUM_RE <- "^[-+(]?[$£€]?[0-9,]*\\.?[0-9]+([eE][-+]?[0-9]+)?%?\\)?$"

content_score <- function(x, kind) {
  v <- trimws(as.character(x))
  v <- v[!is.na(v) & nzchar(v)]
  if (!length(v)) return(0)
  is_num <- grepl(NUM_RE, v)
  pnum   <- mean(is_num)
  num    <- suppressWarnings(as.numeric(gsub("[^0-9.eE-]", "", v[is_num])))
  pyoa   <- mean(nchar(v) >= 5 & grepl("^[0-9]{2}$", substr(v, 4, 5)) & !is_num)
  uniq   <- length(unique(v)) / length(v)
  switch(kind,
    ref_unique = 0.6 * pyoa + 0.4 * uniq,
    ref_master = 0.6 * pyoa + 0.4 * (1 - uniq),
    amount     = if (pnum == 0) 0 else pnum * mean(num >= 0, na.rm = TRUE),
    share      = if (pnum == 0) 0 else pnum * mean(num >= 0 & num <= 100, na.rm = TRUE),
    text       = 0.8,
    0)
}

classify_type <- function(values, cfg) {
  lv <- tolower(trimws(as.character(values)))
  n  <- length(lv)
  if (!n) return(data.frame(value = character(), category = character(),
                            note = character(), stringsAsFactors = FALSE))
  hits <- matrix(vapply(TYPE_PATTERNS, function(p) grepl(p, lv, perl = TRUE),
                        logical(n)), nrow = n)
  cats <- names(TYPE_PATTERNS)
  category <- apply(hits, 1, function(h) if (any(h)) cats[which(h)[1]] else NA_character_)
  note <- ifelse(rowSums(hits) > 1,
                 paste("matches", apply(hits, 1, function(h) paste(cats[h], collapse = "/")),
                       "- first used"),
          ifelse(rowSums(hits) == 0, "NO MATCH - add to imp_cfg$type_overrides", "pattern"))
  ov <- cfg$type_overrides
  if (length(ov)) {
    bad <- setdiff(unname(ov), names(cfg$type_labels))
    if (length(bad)) stop("type_overrides use unknown categories: ", paste(bad, collapse = ", "))
    k <- match(lv, tolower(trimws(names(ov))))
    category[!is.na(k)] <- unname(ov[k[!is.na(k)]])
    note[!is.na(k)]     <- "override"
  }
  data.frame(value = values, category = category, note = note, stringsAsFactors = FALSE)
}

# ---- 3. FILE / SHEET / HEADER -----------------------------------------------
resolve_file <- function(file, dir) {
  path <- if (file.exists(file)) file else file.path(dir, file)
  if (!file.exists(path)) {
    found <- list.files(c(dir, "."), pattern = "\\.xls[xm]?$", ignore.case = TRUE)
    stop("File not found: '", path, "' (working directory: ", getwd(), ")\n",
         "Excel files available: ",
         if (length(found)) paste(unique(found), collapse = ", ") else "none", call. = FALSE)
  }
  path
}

find_sheet <- function(path, wanted) {
  sheets <- readxl::excel_sheets(path)
  hit <- which(norm(sheets) == norm(wanted))
  if (!length(hit)) hit <- which(grepl("risk", sheets, ignore.case = TRUE))
  if (length(hit) != 1)
    stop("Could not identify the '", wanted, "' tab in ", basename(path),
         ". Tabs found: ", paste(sheets, collapse = ", "), call. = FALSE)
  if (norm(sheets[hit]) != norm(wanted))
    message("  Note: using tab '", sheets[hit], "' (closest to '", wanted, "')")
  sheets[hit]
}

detect_header <- function(raw, max_rows = 25) {
  all_syn <- unique(c(unlist(lapply(FIELD_SPEC, `[[`, "syn")), TYPE_SYN))
  rows <- seq_len(min(nrow(raw), max_rows))
  hits <- vapply(rows, function(i) {
    cells <- unlist(raw[i, ], use.names = FALSE)
    cells <- cells[!is.na(cells) & nzchar(trimws(cells))]
    sum(vapply(cells, function(cl) best_syn(norm(cl), all_syn) >= 0.85, logical(1)))
  }, numeric(1))
  if (!length(hits) || max(hits) == 0)
    stop("Could not find a header row - set imp_cfg$header_row", call. = FALSE)
  which.max(hits)
}

read_risk_data <- function(path, cfg) {
  sheet <- find_sheet(path, cfg$sheet)
  raw <- as.data.frame(readxl::read_excel(path, sheet = sheet, col_names = FALSE,
                                          col_types = "text", .name_repair = "minimal"),
                       stringsAsFactors = FALSE)
  if (!nrow(raw)) stop("Tab '", sheet, "' in ", basename(path), " is empty", call. = FALSE)
  hdr <- if (is.na(cfg$header_row)) detect_header(raw) else cfg$header_row
  h <- trimws(unlist(raw[hdr, ], use.names = FALSE))
  h[is.na(h) | !nzchar(h)] <- paste0("col_", which(is.na(h) | !nzchar(h)))
  df <- raw[-seq_len(hdr), , drop = FALSE]
  names(df) <- make.unique(h)
  df <- df[rowSums(!is.na(df)) > 0, , drop = FALSE]          # drop empty rows
  rownames(df) <- NULL
  message(sprintf("  Tab '%s': header on row %d, %d data rows, %d columns",
                  sheet, hdr, nrow(df), ncol(df)))
  df
}

# ---- 4. COLUMN SELECTION ----------------------------------------------------
score_type_cols <- function(df, cfg) {
  do.call(rbind, lapply(names(df), function(cn) {
    v  <- trimws(df[[cn]]); v <- v[!is.na(v) & nzchar(v)]
    tb <- table(v)
    vs <- if (length(tb) >= 1 && length(tb) <= 50) {
      cl <- classify_type(names(tb), cfg)
      sum(tb[!is.na(cl$category)]) / sum(tb)
    } else 0
    ns <- name_score(cn, TYPE_SYN, TYPE_WEAK)
    data.frame(column = cn, n_distinct = length(tb), name_score = round(ns, 3),
               value_score = round(vs, 3), score = round(0.4 * ns + 0.6 * vs, 3),
               stringsAsFactors = FALSE)
  }))
}

pick_type_col <- function(df, cfg) {
  if (!is.na(cfg$type_col_override)) {
    if (!cfg$type_col_override %in% names(df))
      stop("type_col_override '", cfg$type_col_override, "' not in file", call. = FALSE)
    return(list(column = cfg$type_col_override, how = "override", scores = NULL))
  }
  sc <- score_type_cols(df, cfg)
  sc <- sc[order(-sc$score), ]
  if (sc$score[1] < 0.4)
    stop("No convincing policy-type column found - set imp_cfg$type_col_override.\n",
         "Best candidates: ", paste0(head(sc$column, 3), " (", head(sc$score, 3), ")",
                                     collapse = ", "), call. = FALSE)
  list(column = sc$column[1], how = "auto", scores = sc)
}

score_fields <- function(df, cols) {
  do.call(rbind, lapply(names(FIELD_SPEC), function(f) {
    sp <- FIELD_SPEC[[f]]
    do.call(rbind, lapply(cols, function(cn) {
      ns <- name_score(cn, sp$syn, sp$weak, sp$neg)
      cs <- content_score(df[[cn]], sp$kind)
      data.frame(field = f, column = cn, name_score = round(ns, 3),
                 content_score = round(cs, 3), score = round(0.7 * ns + 0.3 * cs, 3),
                 stringsAsFactors = FALSE)
    }))
  }))
}

assign_fields <- function(sc, cols, cfg) {
  fields <- names(FIELD_SPEC)
  map <- setNames(rep(NA_character_, length(fields)), fields)
  how <- setNames(rep("not found", length(fields)), fields)
  for (f in names(cfg$field_overrides)) {
    if (!f %in% fields) stop("field_overrides: unknown field '", f, "'", call. = FALSE)
    col <- cfg$field_overrides[[f]]
    if (!col %in% cols) stop("field_overrides: column '", col, "' not in file", call. = FALSE)
    map[f] <- col; how[f] <- "override"
  }
  req <- fields[vapply(FIELD_SPEC, `[[`, logical(1), "required")]
  # Greedy: best remaining (field, column) pair first; required fields before optional
  for (grp in list(req, setdiff(fields, req))) {
    repeat {
      open <- names(map)[is.na(map)]
      cand <- sc[sc$field %in% intersect(grp, open) & !sc$column %in% map &
                   sc$score >= cfg$min_score & sc$name_score >= 0.5, ]
      if (!nrow(cand)) break
      b <- cand[which.max(cand$score), ]
      map[b$field] <- b$column; how[b$field] <- "auto"
    }
  }
  list(map = map, how = how)
}

# ---- 5. IMPORT ONE FILE -----------------------------------------------------
import_one <- function(path, label, cfg) {
  message("\n", label, ": ", path)
  df <- read_risk_data(path, cfg)

  tp <- pick_type_col(df, cfg)
  message("  Policy-type column: '", tp$column, "' (", tp$how, ")")

  cols <- setdiff(names(df), tp$column)
  sc   <- score_fields(df, cols)
  asg  <- assign_fields(sc, cols, cfg)
  map  <- asg$map

  field_report <- do.call(rbind, lapply(names(FIELD_SPEC), function(f) {
    s   <- sc[sc$field == f, ]
    s   <- s[order(-s$score), ]
    hit <- s[s$column %in% map[f], ]
    alt <- head(s[!s$column %in% map[f], ], 3)
    data.frame(file = label, field = f, required = FIELD_SPEC[[f]]$required,
               matched_column = unname(map[f]), how = unname(asg$how[f]),
               name_score = if (nrow(hit)) hit$name_score else NA,
               content_score = if (nrow(hit)) hit$content_score else NA,
               score = if (nrow(hit)) hit$score else NA,
               runners_up = paste0(alt$column, " (", alt$score, ")", collapse = "; "),
               stringsAsFactors = FALSE)
  }))
  field_report <- rbind(field_report, data.frame(
    file = label, field = "policy_type", required = TRUE, matched_column = tp$column,
    how = tp$how,
    name_score    = if (is.null(tp$scores)) NA else tp$scores$name_score[1],
    content_score = if (is.null(tp$scores)) NA else tp$scores$value_score[1],
    score         = if (is.null(tp$scores)) NA else tp$scores$score[1],
    runners_up    = if (is.null(tp$scores)) "" else
      paste0(tp$scores$column[2:4], " (", tp$scores$score[2:4], ")", collapse = "; "),
    stringsAsFactors = FALSE))

  # Policy-type value suggestions
  raw_type <- trimws(df[[tp$column]])
  raw_type[!is.na(raw_type) & !nzchar(raw_type)] <- NA
  tb <- table(raw_type, useNA = "ifany")
  vals <- names(tb); vals[is.na(vals)] <- NA
  cl <- classify_type(vals, cfg)
  type_report <- data.frame(
    file = label, raw_value = ifelse(is.na(vals), "(blank)", vals), n_rows = as.integer(tb),
    suggested_category = cl$category,
    standard_label = unname(unlist(cfg$type_labels)[cl$category]),
    basis = cl$note, stringsAsFactors = FALSE)
  type_report <- type_report[order(type_report$suggested_category, -type_report$n_rows), ]

  # Standardised data
  n   <- nrow(df)
  get <- function(f) if (is.na(map[f])) rep(NA_character_, n) else trimws(df[[map[f]]])
  cat_row <- cl$category[match(raw_type, vals)]
  std <- data.frame(
    source_row      = seq_len(n),
    policy_ref      = get("policy_ref"),
    master_ref      = get("master_ref"),
    policy_type     = ifelse(is.na(cat_row), raw_type,
                             unname(unlist(cfg$type_labels)[cat_row])),
    policy_type_raw = raw_type,
    limit           = get("limit"),
    gn_premium      = get("gn_premium"),
    stringsAsFactors = FALSE)
  for (f in names(FIELD_SPEC)[!vapply(FIELD_SPEC, `[[`, logical(1), "required")])
    if (!is.na(map[f])) std[[f]] <- get(f)

  keep <- !is.na(std$policy_ref) | !is.na(std$master_ref)
  if (any(!keep)) message("  Dropped ", sum(!keep),
                          " row(s) with no policy or master reference (e.g. totals)")
  std <- std[keep, , drop = FALSE]

  list(data = std, field_report = field_report, type_report = type_report,
       missing_required = names(map)[is.na(map) & vapply(FIELD_SPEC, `[[`, logical(1), "required")])
}

# ---- 6. DEMO WORKBOOKS ------------------------------------------------------
make_demo_workbooks <- function(cfg) {
  has_ox <- requireNamespace("openxlsx", quietly = TRUE)
  has_wx <- requireNamespace("writexl",  quietly = TRUE)
  if (!has_ox && !has_wx)
    stop("Demo mode needs {openxlsx} or {writexl}: install.packages('openxlsx')", call. = FALSE)
  dir.create(cfg$input_dir, showWarnings = FALSE, recursive = TRUE)
  set.seed(11)
  ref  <- function(p, y, i) sprintf("%s%s%05d", p, y, i)
  roll <- function(x) { substr(x, 4, 5) <- sprintf("%02d", as.integer(substr(x, 4, 5)) + 1); x }
  classes <- c("Property D&F", "Marine Cargo", "Energy Onshore", "Terrorism", "Casualty")
  rec  <- function(pref, master, type, lim, ded, line, stack, cob) {
    n <- length(pref)
    data.frame("Risk Reference" = pref, "Master UMR" = master, "Placement Basis" = type,
               "Insured Name" = paste("Insured", pref), "Underwriter" = sample(c("AB", "CD", "EF"), n, TRUE),
               "Limit 100%" = lim, "GNP (USD)" = round(lim * runif(n, .002, .012)),
               "Premium Tax" = round(runif(n, 100, 5000)), "Stack Code" = stack,
               "Excess" = ded, "Signed Line %" = line, "Class of Business" = cob,
               check.names = FALSE, stringsAsFactors = FALSE)
  }
  ind <- function(p, y, i, type) {
    n <- length(i)
    rec(ref(p, y, i), NA_character_, type, sample(c(1e6, 2e6, 5e6, 1e7, 2.5e7), n, TRUE),
        sample(c(1e4, 2.5e4, 5e4), n, TRUE), sample(c(10, 25, 50, 100), n, TRUE),
        sprintf("STK%03d", sample(1:20, n, TRUE)), sample(classes, n, TRUE))
  }
  decs <- function(master, y, i, type) {
    n <- length(i)
    rec(ref("DEC", y, i), master, type, sample(c(2.5e5, 5e5, 1e6, 2e6), n, TRUE),
        sample(c(1e3, 2.5e3, 5e3), n, TRUE), rep(sample(c(25, 50, 100), 1), n), NA_character_,
        sample(c("Property Binder", "Property D&F"), n, TRUE, prob = c(.8, .2)))
  }
  # Expiring: labels as one system writes them; block entries are YOA 24
  e_ind <- rbind(ind("OMK", "25", 1:20, "Open Market"), ind("LSD", "25", 1:8, "Line Slip Dec"),
                 ind("BLK", "24", 1:6, "Block Entry"))
  expiring <- rbind(e_ind,
                    decs("BDM2500001", "25", 1:20,  "Binder Dec"),
                    decs("BDM2400002", "24", 21:35, "Binder Dec"),
                    decs("BDM2500003", "25", 36:47, "Binder Dec"))   # lapses
  # Current: different labels, ~75% renew, some new, one unmapped value
  ren <- e_ind[runif(nrow(e_ind)) < 0.75, ]
  ren$`Risk Reference` <- roll(ren$`Risk Reference`)
  ren$`Placement Basis` <- c("Open Market" = "OM", "Line Slip Dec" = "LS Declaration",
                             "Block Entry" = "Block")[ren$`Placement Basis`]
  current <- rbind(ren, ind("OMK", "26", 101:105, "OM"),
                   decs("BDM2600001", "26", 1:22,  "Coverholder Bdx"),
                   decs("BDM2500002", "25", 23:40, "Coverholder Bdx"),
                   ind("TTY", "26", 1, "Treaty"))
  files <- c(expiring = "demo_expiring_export.xlsx", current = "demo_current_export.xlsx")
  for (k in names(files)) {
    df   <- if (k == "expiring") expiring else current
    path <- file.path(cfg$input_dir, files[[k]])
    if (has_ox) {
      wb <- openxlsx::createWorkbook()
      openxlsx::addWorksheet(wb, "Notes")
      openxlsx::writeData(wb, "Notes", "Demo export")
      openxlsx::addWorksheet(wb, "Risk Data")
      openxlsx::writeData(wb, "Risk Data", paste(k, "risk data export (demo)"), startRow = 1)
      openxlsx::writeData(wb, "Risk Data", df, startRow = 3)       # title row above table
      openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
    } else {
      writexl::write_xlsx(list(Notes = data.frame(note = "Demo export"), "Risk Data" = df), path)
    }
  }
  # Demo curves file (only if you don't already have one)
  cpath <- file.path(cfg$input_dir, "curves.csv")
  if (!file.exists(cpath))
    write.csv(data.frame(curve_id = c("CRV_PROP", "CRV_MAR", "CRV_ENG", "CRV_TERR"),
                         classes  = c("Property D&F; Property Binder", "Marine Cargo",
                                      "Energy Onshore, Energy Offshore", "Terrorism | Property D&F")),
              cpath, row.names = FALSE)
  message("Demo workbooks written to ", normalizePath(cfg$input_dir))
  files
}

# ---- 7. RUN -----------------------------------------------------------------
run_import <- function(cfg) {
  if (!requireNamespace("readxl", quietly = TRUE))
    stop("Package {readxl} is required: install.packages('readxl')", call. = FALSE)

  if (cfg$demo) {
    f <- make_demo_workbooks(cfg)
    cfg$expiring_file <- f[["expiring"]]; cfg$current_file <- f[["current"]]
  }
  paths <- c(Expiring = resolve_file(cfg$expiring_file, cfg$input_dir),
             Current  = resolve_file(cfg$current_file,  cfg$input_dir))

  res <- lapply(names(paths), function(lb) import_one(paths[[lb]], lb, cfg))
  names(res) <- names(paths)

  field_report <- do.call(rbind, lapply(res, `[[`, "field_report"))
  type_report  <- do.call(rbind, lapply(res, `[[`, "type_report"))
  rownames(field_report) <- rownames(type_report) <- NULL

  dir.create(cfg$output_dir, showWarnings = FALSE, recursive = TRUE)
  write.csv(field_report, file.path(cfg$output_dir, "import_field_mapping.csv"), row.names = FALSE, na = "")
  write.csv(type_report,  file.path(cfg$output_dir, "import_type_mapping.csv"),  row.names = FALSE, na = "")

  cat("\n==== FIELD MAPPING ====\n")
  print(field_report[, c("file", "field", "required", "matched_column", "how", "score", "runners_up")],
        row.names = FALSE)
  cat("\n==== POLICY-TYPE VALUES ====\n")
  print(type_report, row.names = FALSE)

  # Consistency between files
  fe <- field_report[field_report$file == "Expiring", c("field", "matched_column")]
  fc <- field_report[field_report$file == "Current",  c("field", "matched_column")]
  differ <- fe$field[!mapply(identical, fe$matched_column,
                             fc$matched_column[match(fe$field, fc$field)])]
  if (length(differ))
    warning("Expiring and current files mapped different columns for: ",
            paste(differ, collapse = ", "), " - check import_field_mapping.csv", call. = FALSE)

  unmapped <- type_report[is.na(type_report$suggested_category), ]
  if (nrow(unmapped))
    warning(nrow(unmapped), " policy-type value(s) not mapped (",
            paste(unique(unmapped$raw_value), collapse = ", "),
            "). They will be flagged as 'Unknown policy type' by the analysis - ",
            "add them to imp_cfg$type_overrides if they belong to a category.", call. = FALSE)

  missing <- unlist(lapply(names(res), function(lb)
    if (length(res[[lb]]$missing_required))
      paste0(lb, ": ", paste(res[[lb]]$missing_required, collapse = ", "))))
  if (length(missing))
    stop("Required field(s) not matched - ", paste(missing, collapse = " | "),
         "\nSee runners_up in import_field_mapping.csv and set imp_cfg$field_overrides.",
         call. = FALSE)

  write.csv(res$Expiring$data, file.path(cfg$output_dir, "expiring.csv"), row.names = FALSE, na = "")
  write.csv(res$Current$data,  file.path(cfg$output_dir, "current.csv"),  row.names = FALSE, na = "")
  message("\nReady for analysis: ",
          file.path(cfg$output_dir, "expiring.csv"), " (", nrow(res$Expiring$data), " rows), ",
          file.path(cfg$output_dir, "current.csv"),  " (", nrow(res$Current$data), " rows)",
          "\nNext: set cfg$use_dummy <- FALSE in policy_renewal_analysis.R and run it.")
  invisible(res)
}

imported <- run_import(imp_cfg)
