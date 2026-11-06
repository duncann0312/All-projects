#!/usr/bin/env Rscript
# =============================================================================
# Policy consolidation & renewal status analysis
# -----------------------------------------------------------------------------
# Stage 1 - Consolidation & checks
#   * 4 policy types: Open Market, Lineslip Declaration, Block Entry,
#     Binder Declaration.
#   * Binder declarations are grouped by their master reference and replaced
#     by a "Dec profile": declarations banded on LIMIT with GN premium summed.
#     Each binder master then becomes ONE record (key = master_ref).
#   * All other types are individual records keyed on policy_ref, which must
#     be unique -> checked and reported.
#   * Optional fields (cfg$optional_cols) are carried through to the outputs:
#     stack_code (blank for binder masters), deductible (average of decs for
#     binders) and participation (average of decs for binders).
#
# Stage 2 - Renewal status
#   * Year of account (YOA) is held in characters 4-5 of the reference.
#     A renewal carries the same reference with the YOA rolled forward by one
#     (e.g. "24" -> "25", "25" -> "26"); the rest is unchanged. Each reference
#     is rolled from its own YOA, so the expiring set may mix years.
#   * Expiring records -> Renewed / Lapsed
#   * Current records  -> Renewal / New
#
# Base R only (no packages needed). If {openxlsx} is installed, a single
# Excel workbook is also written alongside the CSVs.
#
# To use real data: run policy_renewal_import.R first. It reads the Excel
# exports, matches the columns and policy types, and writes input/expiring.csv
# and input/current.csv in the standard layout below. Then set
# cfg$use_dummy <- FALSE here and run this script.
# =============================================================================

# ---- 0. CONFIGURATION -------------------------------------------------------
cfg <- list(
  use_dummy     = TRUE,
  expiring_file = "input/expiring.csv",   # written by policy_renewal_import.R
  current_file  = "input/current.csv",
  output_dir    = "output",

  # Standard name used in this script = column name in YOUR data
  col_map = c(
    policy_ref  = "policy_ref",
    master_ref  = "master_ref",
    policy_type = "policy_type",
    limit       = "limit",
    gn_premium  = "gn_premium"
  ),

  # OPTIONAL fields carried through to the outputs.
  #   col    = column name in YOUR data (if absent, the field is left blank)
  #   type   = "text" or "numeric"
  #   binder = how binder declarations roll up to the master / Dec profile band:
  #            "blank", "mean", "sum", "max" or "first"
  # Delete an entry to drop the field from the outputs; add entries for more.
  optional_cols = list(
    stack_code    = list(col = "stack_code",    type = "text",    binder = "blank"),
    deductible    = list(col = "deductible",    type = "numeric", binder = "mean"),
    participation = list(col = "participation", type = "numeric", binder = "mean")
  ),

  # Policy-type values (keep the same as imp_cfg$type_labels in the import)
  type_labels = list(
    open_market = "Open Market",
    lineslip    = "Lineslip Declaration",
    block       = "Block Entry",
    binder      = "Binder Declaration"
  ),

  # Year-of-account position in the reference. Renewal = YOA + 1.
  yoa_pos      = c(4L, 5L),
  # YOAs expected in each set - used only for the data check (references
  # outside these are flagged but still matched on YOA + 1)
  expiring_yoa = c("24", "25"),
  current_yoa  = c("25", "26"),

  # Limit bands for the binder Dec profile (upper bound inclusive)
  limit_breaks = c(0, 5e5, 1e6, 2.5e6, 5e6, 1e7, Inf)
)

STATUS <- list(
  renewed  = "Renewed",
  lapsed   = "Lapsed",
  renewal  = "Renewal",
  new      = "New",
  unclass  = "Unclassified (reference / YOA check failed)"
)

# ---- 1. HELPERS -------------------------------------------------------------
clean_ref <- function(x) {
  x <- toupper(trimws(as.character(x)))
  x[x %in% c("", "NA", "N/A", "NULL")] <- NA_character_
  x
}

clean_num <- function(x) {
  if (is.numeric(x)) return(as.numeric(x))
  as.numeric(gsub("[^0-9.eE-]", "", as.character(x)))   # strips commas, currency
}

# Roll the 2-digit YOA in a reference by 'by' years (+1 = renewal reference,
# -1 = prior-year reference). Returns NA where the reference is missing, too
# short, or the YOA characters are not two digits.
shift_yoa <- function(ref, by, pos) {
  out <- rep(NA_character_, length(ref))
  yoa <- substr(ref, pos[1], pos[2])
  ok  <- !is.na(ref) & nchar(ref) >= pos[2] & grepl("^[0-9]{2}$", yoa)
  x <- ref[ok]
  substr(x, pos[1], pos[2]) <- sprintf("%02d", (as.integer(yoa[ok]) + by) %% 100)
  out[ok] <- x
  out
}

fmt_m <- function(x) {
  ifelse(is.infinite(x), "Inf",
         paste0(formatC(x / 1e6, format = "fg", digits = 4), "m"))
}

band_labels <- function(breaks) {
  lo <- head(breaks, -1); hi <- tail(breaks, -1)
  ifelse(is.infinite(hi), paste0("> ", fmt_m(lo)),
         paste0(fmt_m(lo), " - ", fmt_m(hi)))
}

band_limit <- function(limit, breaks) {
  lv <- band_labels(breaks)
  b  <- as.character(cut(limit, breaks = breaks, labels = lv,
                         right = TRUE, include.lowest = TRUE))
  b[is.na(b)] <- "Unknown"
  factor(b, levels = c(lv, "Unknown"))
}

sum_na <- function(x) sum(x, na.rm = TRUE)

# Roll up an optional field across binder declarations
rollup <- function(x, rule) {
  if (all(is.na(x))) return(x[NA_integer_])
  switch(rule,
         blank = x[NA_integer_],
         mean  = mean(x, na.rm = TRUE),
         sum   = sum(x, na.rm = TRUE),
         max   = max(x, na.rm = TRUE),
         first = x[!is.na(x)][1],
         stop("Unknown binder roll-up rule: ", rule))
}

# ---- 2. STANDARDISE INPUT ---------------------------------------------------
standardise <- function(df, cfg, label = "") {
  src <- cfg$col_map
  missing_cols <- setdiff(src, names(df))
  if (length(missing_cols))
    stop("Missing column(s) in input: ", paste(missing_cols, collapse = ", "))
  names(df)[match(src, names(df))] <- names(src)
  df$policy_ref  <- clean_ref(df$policy_ref)
  df$master_ref  <- clean_ref(df$master_ref)
  df$policy_type <- trimws(as.character(df$policy_type))
  df$limit       <- clean_num(df$limit)
  df$gn_premium  <- clean_num(df$gn_premium)

  for (nm in names(cfg$optional_cols)) {
    o <- cfg$optional_cols[[nm]]
    if (o$col %in% names(df)) {
      v <- df[[o$col]]
    } else {
      message(label, ": optional column '", o$col, "' not found - left blank")
      v <- rep(NA, nrow(df))
    }
    if (o$type == "numeric") {
      v <- clean_num(v)
    } else {
      v <- trimws(as.character(v)); v[v %in% ""] <- NA_character_
    }
    df[[nm]] <- v
  }
  df$row_id <- seq_len(nrow(df))
  df
}

# ---- 3. STAGE 1 CHECKS ------------------------------------------------------
issue_template <- data.frame(dataset = character(), check = character(),
                             row_id = integer(), policy_ref = character(),
                             master_ref = character(), policy_type = character(),
                             detail = character(), stringsAsFactors = FALSE)

validate_dataset <- function(df, label, expected_yoa, cfg) {
  tl <- cfg$type_labels
  issues <- list()
  add <- function(check, idx, detail = "") {
    if (!length(idx)) return(invisible(NULL))
    issues[[length(issues) + 1]] <<- data.frame(
      dataset = label, check = check, row_id = df$row_id[idx],
      policy_ref = df$policy_ref[idx], master_ref = df$master_ref[idx],
      policy_type = df$policy_type[idx], detail = detail,
      stringsAsFactors = FALSE)
  }
  is_b <- df$policy_type %in% tl$binder

  # Unknown policy type
  add("Unknown policy type", which(!df$policy_type %in% unlist(tl)),
      "Not in cfg$type_labels")

  # Individual policies: reference present and unique
  add("Missing policy reference", which(!is_b & is.na(df$policy_ref)))
  ind <- which(!is_b & !is.na(df$policy_ref))
  cnt <- table(df$policy_ref[ind])
  dup <- ind[df$policy_ref[ind] %in% names(cnt)[cnt > 1]]
  add("Duplicate policy reference", dup,
      paste("Reference appears", as.integer(cnt[df$policy_ref[dup]]), "times"))

  # Binder declarations: master reference present
  add("Binder declaration missing master reference",
      which(is_b & is.na(df$master_ref)), "Excluded from Dec profile")

  # Binder master refs must not clash with an individual policy reference
  b_key  <- ifelse(is_b, df$master_ref, NA)
  first  <- is_b & !is.na(b_key) & !duplicated(b_key)
  add("Binder master reference clashes with individual policy reference",
      which(first & b_key %in% df$policy_ref[!is_b]))

  # YOA characters on the record key (policy_ref, or master_ref for binders)
  key <- ifelse(is_b, df$master_ref, df$policy_ref)
  yoa <- substr(key, cfg$yoa_pos[1], cfg$yoa_pos[2])
  bad <- which((!is_b | first) & !is.na(key) & !yoa %in% expected_yoa)
  add("Unexpected year of account in reference", bad,
      paste0("Chars ", cfg$yoa_pos[1], "-", cfg$yoa_pos[2], " = '", yoa[bad],
             "', expected ", paste(expected_yoa, collapse = " or ")))

  # Numeric fields
  add("Missing or non-positive limit", which(is.na(df$limit) | df$limit <= 0))
  add("Missing GN premium", which(is.na(df$gn_premium)))

  if (!length(issues)) return(issue_template)
  do.call(rbind, issues)
}

# ---- 4. BINDER DEC PROFILE --------------------------------------------------
build_dec_profile <- function(df, cfg, label) {
  b <- df[df$policy_type %in% cfg$type_labels$binder & !is.na(df$master_ref), ,
          drop = FALSE]
  if (!nrow(b)) return(NULL)
  b$limit_band  <- band_limit(b$limit, cfg$limit_breaks)
  b$n_decs      <- 1L
  b$total_limit <- b$limit
  prof <- aggregate(cbind(n_decs, total_limit, gn_premium) ~ master_ref + limit_band,
                    data = b, FUN = sum_na, na.action = na.pass)
  # Optional numeric fields rolled up per band, e.g. mean_deductible
  for (nm in names(cfg$optional_cols)) {
    o <- cfg$optional_cols[[nm]]
    if (o$type != "numeric" || o$binder == "blank") next
    a <- aggregate(list(val = b[[nm]]),
                   by = list(master_ref = b$master_ref, limit_band = b$limit_band),
                   FUN = function(x) rollup(x, o$binder))
    names(a)[names(a) == "val"] <- paste0(o$binder, "_", nm)
    prof <- merge(prof, a, by = c("master_ref", "limit_band"), all.x = TRUE)
  }
  prof$limit_band <- factor(as.character(prof$limit_band), levels = levels(b$limit_band))
  prof <- prof[order(prof$master_ref, prof$limit_band), ]
  prof$avg_premium_per_dec <- round(prof$gn_premium / prof$n_decs, 2)
  prof$pct_of_master_premium <-
    round(100 * prof$gn_premium / ave(prof$gn_premium, prof$master_ref, FUN = sum), 1)
  rownames(prof) <- NULL
  cbind(dataset = label, prof, stringsAsFactors = FALSE)
}

# ---- 5. CONSOLIDATED POLICY SET (one row per renewable unit) ----------------
consolidate <- function(df, cfg) {
  tl   <- cfg$type_labels
  is_b <- df$policy_type %in% tl$binder

  ind <- df[!is_b, , drop = FALSE]
  out <- data.frame(policy_ref = ind$policy_ref, policy_type = ind$policy_type,
                    record_basis = rep("Individual", nrow(ind)),
                    n_underlying = rep(1L, nrow(ind)),
                    limit = ind$limit, gn_premium = ind$gn_premium,
                    stringsAsFactors = FALSE)
  for (nm in names(cfg$optional_cols)) out[[nm]] <- ind[[nm]]

  b <- df[is_b & !is.na(df$master_ref), , drop = FALSE]
  if (nrow(b)) {
    bout <- do.call(rbind, lapply(split(b, b$master_ref), function(x) {
      mx <- suppressWarnings(max(x$limit, na.rm = TRUE))
      r <- data.frame(policy_ref = x$master_ref[1], policy_type = tl$binder,
                      record_basis = "Dec profile (binder master)",
                      n_underlying = nrow(x),
                      limit = if (is.finite(mx)) mx else NA_real_,   # max dec limit
                      gn_premium = sum_na(x$gn_premium),
                      stringsAsFactors = FALSE)
      for (nm in names(cfg$optional_cols))
        r[[nm]] <- rollup(x[[nm]], cfg$optional_cols[[nm]]$binder)
      r
    }))
    out <- rbind(out, bout)
  }
  rownames(out) <- NULL
  out
}

# ---- 6. STAGE 2 - RENEWAL STATUS --------------------------------------------
classify <- function(con_exp, con_cur, cfg) {
  p <- cfg$yoa_pos

  # Expiring -> Renewed / Lapsed
  con_exp$expected_renewal_ref <- shift_yoa(con_exp$policy_ref, 1, p)
  con_exp$status <- ifelse(is.na(con_exp$expected_renewal_ref), STATUS$unclass,
                    ifelse(con_exp$expected_renewal_ref %in% con_cur$policy_ref,
                           STATUS$renewed, STATUS$lapsed))
  mc <- match(con_exp$expected_renewal_ref, con_cur$policy_ref)
  con_exp$renewal_gn_premium <- con_cur$gn_premium[mc]

  # Current -> Renewal / New
  con_cur$prior_ref <- shift_yoa(con_cur$policy_ref, -1, p)
  con_cur$status <- ifelse(is.na(con_cur$prior_ref), STATUS$unclass,
                    ifelse(con_cur$prior_ref %in% con_exp$policy_ref,
                           STATUS$renewal, STATUS$new))
  me <- match(con_cur$prior_ref, con_exp$policy_ref)
  con_cur$expiring_policy_type <- con_exp$policy_type[me]
  con_cur$expiring_limit       <- con_exp$limit[me]
  con_cur$expiring_gn_premium  <- con_exp$gn_premium[me]
  for (nm in names(cfg$optional_cols))
    con_cur[[paste0("expiring_", nm)]] <- con_exp[[nm]][me]
  con_cur$premium_change_pct   <- ifelse(!is.na(me) & con_cur$expiring_gn_premium > 0,
    round(100 * (con_cur$gn_premium / con_cur$expiring_gn_premium - 1), 1), NA)
  con_cur$type_changed <- !is.na(me) &
    con_cur$policy_type != con_cur$expiring_policy_type

  list(expiring = con_exp, current = con_cur)
}

summarise_status <- function(df, label) {
  df$policies <- 1L
  s <- aggregate(cbind(policies, n_underlying, gn_premium) ~ policy_type + status,
                 data = df, FUN = sum_na, na.action = na.pass)
  tot <- aggregate(cbind(policies, n_underlying, gn_premium) ~ status,
                   data = df, FUN = sum_na, na.action = na.pass)
  tot$policy_type <- "ALL"
  s <- rbind(s, tot[, names(s)])
  s$pct_of_gn_premium <- round(100 * s$gn_premium /
                                 sum_na(df$gn_premium), 1)
  s <- s[order(s$policy_type == "ALL", s$policy_type, s$status), ]
  rownames(s) <- NULL
  cbind(dataset = label, s, stringsAsFactors = FALSE)
}

reconcile <- function(cls) {
  e <- cls$expiring; cu <- cls$current
  sp <- function(d, st) sum_na(d$gn_premium[d$status == st])
  sn <- function(d, st) sum(d$status == st)
  data.frame(
    measure = c("Expiring - records", "  Renewed", "  Lapsed", "  Unclassified",
                "Expiring - GN premium", "  Renewed", "  Lapsed", "  Unclassified",
                "Current - records", "  Renewal", "  New", "  Unclassified",
                "Current - GN premium", "  Renewal", "  New", "  Unclassified",
                "Retention rate by count (%)", "Retention rate by premium (%)",
                "Premium change on renewals (%)"),
    value = c(nrow(e), sn(e, STATUS$renewed), sn(e, STATUS$lapsed), sn(e, STATUS$unclass),
              sum_na(e$gn_premium), sp(e, STATUS$renewed), sp(e, STATUS$lapsed), sp(e, STATUS$unclass),
              nrow(cu), sn(cu, STATUS$renewal), sn(cu, STATUS$new), sn(cu, STATUS$unclass),
              sum_na(cu$gn_premium), sp(cu, STATUS$renewal), sp(cu, STATUS$new), sp(cu, STATUS$unclass),
              round(100 * sn(e, STATUS$renewed) / max(nrow(e), 1), 1),
              round(100 * sp(e, STATUS$renewed) / max(sum_na(e$gn_premium), 1), 1),
              round(100 * (sp(cu, STATUS$renewal) / max(sp(e, STATUS$renewed), 1) - 1), 1)),
    stringsAsFactors = FALSE)
}

# Band-by-band comparison of Dec profiles for binders that renewed
compare_profiles <- function(prof_exp, prof_cur, cfg) {
  if (is.null(prof_exp) || is.null(prof_cur)) return(NULL)
  pe <- prof_exp
  pe$current_master_ref <- shift_yoa(pe$master_ref, 1, cfg$yoa_pos)
  pe <- pe[pe$current_master_ref %in% prof_cur$master_ref, ]
  pc <- prof_cur[prof_cur$master_ref %in% pe$current_master_ref, ]
  if (!nrow(pe)) return(NULL)
  pe <- data.frame(current_master_ref = pe$current_master_ref, limit_band = pe$limit_band,
                   n_decs_expiring = pe$n_decs, gn_premium_expiring = pe$gn_premium)
  pc <- data.frame(current_master_ref = pc$master_ref, limit_band = pc$limit_band,
                   n_decs_current = pc$n_decs, gn_premium_current = pc$gn_premium)
  m <- merge(pe, pc, by = c("current_master_ref", "limit_band"), all = TRUE)
  for (v in c("n_decs_expiring", "gn_premium_expiring", "n_decs_current", "gn_premium_current"))
    m[[v]][is.na(m[[v]])] <- 0
  m$expiring_master_ref <- shift_yoa(m$current_master_ref, -1, cfg$yoa_pos)
  m$premium_change <- m$gn_premium_current - m$gn_premium_expiring
  m$limit_band <- factor(as.character(m$limit_band), levels = levels(prof_exp$limit_band))
  m <- m[order(m$current_master_ref, m$limit_band),
         c("expiring_master_ref", "current_master_ref", "limit_band",
           "n_decs_expiring", "n_decs_current",
           "gn_premium_expiring", "gn_premium_current", "premium_change")]
  rownames(m) <- NULL
  m
}

# ---- 7. PIPELINE ------------------------------------------------------------
run_analysis <- function(expiring_raw, current_raw, cfg) {
  exp <- standardise(expiring_raw, cfg, "Expiring")
  cur <- standardise(current_raw,  cfg, "Current")

  issues <- rbind(validate_dataset(exp, "Expiring", cfg$expiring_yoa, cfg),
                  validate_dataset(cur, "Current",  cfg$current_yoa,  cfg))

  prof_exp <- build_dec_profile(exp, cfg, "Expiring")
  prof_cur <- build_dec_profile(cur, cfg, "Current")

  cls <- classify(consolidate(exp, cfg), consolidate(cur, cfg), cfg)

  # Post-classification checks
  tc <- cls$current[cls$current$type_changed, ]
  if (nrow(tc))
    issues <- rbind(issues, data.frame(
      dataset = "Current", check = "Policy type changed on renewal", row_id = NA_integer_,
      policy_ref = tc$policy_ref, master_ref = NA_character_, policy_type = tc$policy_type,
      detail = paste("Expiring type:", tc$expiring_policy_type), stringsAsFactors = FALSE))
  for (nm in c("expiring", "current")) {
    d <- cls[[nm]]$policy_ref[!is.na(cls[[nm]]$policy_ref)]
    if (anyDuplicated(d))
      warning(nm, " consolidated set still has duplicate keys - see check_issues")
  }

  list(
    check_issues          = issues,
    check_issue_summary   = if (nrow(issues))
                              aggregate(list(n_rows = rep(1L, nrow(issues))),
                                        by = list(dataset = issues$dataset, check = issues$check),
                                        FUN = sum) else NULL,
    dec_profile_expiring  = prof_exp,
    dec_profile_current   = prof_cur,
    dec_profile_movement  = compare_profiles(prof_exp, prof_cur, cfg),
    expiring_status       = cls$expiring,
    current_status        = cls$current,
    status_summary        = rbind(summarise_status(cls$expiring, "Expiring"),
                                  summarise_status(cls$current,  "Current")),
    reconciliation        = reconcile(cls)
  )
}

write_outputs <- function(res, dir) {
  dir.create(dir, showWarnings = FALSE, recursive = TRUE)
  res <- Filter(function(x) !is.null(x) && nrow(x) > 0, res)
  for (nm in names(res))
    write.csv(res[[nm]], file.path(dir, paste0(nm, ".csv")), row.names = FALSE, na = "")
  if (requireNamespace("openxlsx", quietly = TRUE)) {
    names(res) <- substr(names(res), 1, 31)
    openxlsx::write.xlsx(res, file.path(dir, "policy_renewal_analysis.xlsx"), overwrite = TRUE)
  }
  message("Outputs written to: ", normalizePath(dir))
}

# ---- 8. DUMMY DATA ----------------------------------------------------------
make_dummy_data <- function(cfg, seed = 2026) {
  set.seed(seed)
  tl  <- cfg$type_labels
  lims <- c(2.5e5, 5e5, 1e6, 2e6, 5e6, 1e7, 2.5e7, 5e7)
  # Reference format: 3-char prefix + 2-char YOA (chars 4-5) + 5-digit number
  mk_ref <- function(prefix, yoa, idx) sprintf("%s%s%05d", prefix, yoa, idx)

  mk_ind <- function(prefix, type, yoa, idx) {
    n <- length(idx)
    data.frame(policy_ref = mk_ref(prefix, yoa, idx), master_ref = NA_character_,
               policy_type = type, insured = paste("Insured", prefix, idx),
               limit = sample(lims[3:8], n, TRUE),
               gn_premium = round(runif(n, 1e4, 4e5)),
               stack_code = sprintf("STK%03d", sample(1:25, n, TRUE)),
               deductible = sample(c(1e4, 2.5e4, 5e4, 1e5), n, TRUE),
               participation = sample(c(.1, .2, .25, .5, 1), n, TRUE),
               stringsAsFactors = FALSE)
  }
  mk_binder <- function(master, yoa, n, start) {
    idx <- start + seq_len(n) - 1
    lim <- sample(lims[1:6], n, TRUE, prob = c(.25, .3, .2, .12, .08, .05))
    data.frame(policy_ref = mk_ref("DEC", yoa, idx), master_ref = master,
               policy_type = tl$binder, insured = paste("Dec insured", idx),
               limit = lim, gn_premium = round(lim * runif(n, .002, .012)),
               stack_code = NA_character_,                    # blank for decs
               deductible = sample(c(1e3, 2.5e3, 5e3, 1e4), n, TRUE),
               participation = sample(c(.25, .5, 1), 1),      # same line on all decs
               stringsAsFactors = FALSE)
  }
  mk_binders <- function(masters, yoa, start = 1) {
    out <- list()
    for (m in masters) {
      n <- sample(15:40, 1)
      out[[m]] <- mk_binder(m, yoa, n, start)
      start <- start + n
    }
    do.call(rbind, out)
  }

  # --- Expiring: mostly YOA 25; block entries and binder master 2 are YOA 24
  exp_ind <- rbind(mk_ind("OMK", tl$open_market, "25", 1:30),
                   mk_ind("LSD", tl$lineslip,    "25", 1:15),
                   mk_ind("BLK", tl$block,       "24", 1:10))
  exp_masters <- c(mk_ref("BDM", "25", 1), mk_ref("BDM", "24", 2),
                   mk_ref("BDM", "25", 3:4))
  expiring <- rbind(exp_ind, mk_binders(exp_masters, "25"))

  # --- Current: ~75% of individual policies renew at YOA + 1 (25->26, 24->25)
  ren <- exp_ind[runif(nrow(exp_ind)) < 0.75, ]
  ren$policy_ref <- shift_yoa(ren$policy_ref, 1, cfg$yoa_pos)
  ren$gn_premium <- round(ren$gn_premium * runif(nrow(ren), 0.9, 1.15))
  new_ind <- rbind(mk_ind("OMK", tl$open_market, "26", 101:108),
                   mk_ind("LSD", tl$lineslip,    "26", 101:104),
                   mk_ind("BLK", tl$block,       "25", 101:103))
  # Binders: masters 1-3 renew (incl. 24 -> 25), master 4 lapses, one new master
  cur_masters <- c(shift_yoa(exp_masters[1:3], 1, cfg$yoa_pos), mk_ref("BDM", "26", 10))
  current <- rbind(ren, new_ind, mk_binders(cur_masters, "26", start = 1))

  # --- Deliberate data issues so the checks have something to find
  expiring$policy_ref[5] <- mk_ref("OMK", "23", 5)                  # unexpected YOA
  current <- rbind(current, current[1, ])                           # duplicate ref
  current$master_ref[which(current$policy_type == tl$binder)[1]] <- NA  # no master
  ls_ren <- which(current$policy_type == tl$lineslip &
                  as.integer(substr(current$policy_ref, 6, 10)) < 100)[1]
  current$policy_type[ls_ren] <- tl$open_market                     # type change

  list(expiring = expiring, current = current)
}

# ---- 9. MAIN ----------------------------------------------------------------
if (cfg$use_dummy) {
  d <- make_dummy_data(cfg)
  expiring_raw <- d$expiring
  current_raw  <- d$current
  dir.create(cfg$output_dir, showWarnings = FALSE, recursive = TRUE)
  write.csv(expiring_raw, file.path(cfg$output_dir, "dummy_input_expiring.csv"), row.names = FALSE, na = "")
  write.csv(current_raw,  file.path(cfg$output_dir, "dummy_input_current.csv"),  row.names = FALSE, na = "")
} else {
  # Standardised CSVs from policy_renewal_import.R (read as text so references
  # keep leading zeros; numbers are parsed in standardise())
  for (f in c(cfg$expiring_file, cfg$current_file))
    if (!file.exists(f)) stop("Input not found: ", f, " - run policy_renewal_import.R first")
  expiring_raw <- read.csv(cfg$expiring_file, colClasses = "character", na.strings = c("", "NA"),
                           check.names = FALSE)
  current_raw  <- read.csv(cfg$current_file,  colClasses = "character", na.strings = c("", "NA"),
                           check.names = FALSE)
}

res <- run_analysis(expiring_raw, current_raw, cfg)

cat("\n==== STAGE 1: CHECK ISSUES ====\n")
if (is.null(res$check_issue_summary)) cat("No issues found.\n") else
  print(res$check_issue_summary, row.names = FALSE)

cat("\n==== STAGE 1: BINDER DEC PROFILE (current) ====\n")
print(res$dec_profile_current, row.names = FALSE)

cat("\n==== STAGE 2: STATUS SUMMARY ====\n")
print(res$status_summary, row.names = FALSE)

cat("\n==== STAGE 2: RECONCILIATION ====\n")
print(format(res$reconciliation, big.mark = ",", scientific = FALSE), row.names = FALSE)

write_outputs(res, cfg$output_dir)
