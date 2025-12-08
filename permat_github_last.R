# ============================================
# PERMAT – Modèle déterministe & Monte Carlo
# Export Excel (avec ajustement démographique)
# Objectifs TR spécifiques pour cas 3
# + Cibles "moitié du gap de TR" vs gen 1960
# ============================================

# --- Packages ---
suppressPackageStartupMessages({
  library(dplyr)
  library(readr)
  library(stringr)
  library(purrr)
  library(tidyr)
  library(tibble)
  library(openxlsx)
  library(readxl)
  library(ggplot2)
})

# ============================================
# Paramètres globaux
# ============================================
FRAIS_ARERAGE <- 0.0118
RVTG_PAR_CAS <- c("1" = 0.30, "2" = 0.11, "3" = 0.00)
PRORATA_RVTG <- 0.90
PRORATA_PS <- 0.40
TAUX_PS <- 0.172
ANNEE_DEBUT <- 2025
AGE_LIQ <- 64
PLAFOND_DEFISC_PCT <- 0.10
TMI_THRESH_BASE <- c(11497, 29315, 83823, 180294)
TMI_INDEX_RATE <- 0.0175

# Ajustements démographiques des rendements
GAMMA_ACTIONS <- -0.005
GAMMA_OBLIGATIONS <- -0.010

# Objectifs de taux de remplacement pour le cas 3 par génération (plein comblement)
TR_OBJECTIFS_CAS3 <- c(
  "1980" = 0.6492339, "1981" = 0.6457423, "1982" = 0.6428391, "1983" = 0.6397566,
  "1984" = 0.6372416, "1985" = 0.6346842, "1986" = 0.6325324, "1987" = 0.6306867,
  "1988" = 0.6292978, "1989" = 0.6281771, "1990" = 0.6268778, "1991" = 0.6259457,
  "1992" = 0.6249482, "1993" = 0.6239760, "1994" = 0.6229540, "1995" = 0.6220974,
  "1996" = 0.6213564, "1997" = 0.6205621, "1998" = 0.6198159, "1999" = 0.6195105,
  "2000" = 0.6187678
)

# ============================================
# Utilitaires généraux
# ============================================
normalize_id <- function(x) {
  x %>% as.character() %>% str_replace_all("\\s+", " ") %>% str_trim() %>% str_to_lower()
}

parse_number <- function(x) {
  suppressWarnings(as.numeric(str_replace_all(as.character(x), ",", ".")))
}

choose_column <- function(df, candidates, required = FALSE, label = "colonne") {
  match <- candidates[candidates %in% names(df)]
  if (length(match)) {
    match[[1]]
  } else if (required) {
    stop(sprintf("Impossible de trouver %s. Essais: %s. Présentes: %s",
                 label, paste(candidates, collapse = ", "), paste(names(df), collapse = ", ")))
  } else {
    NA_character_
  }
}

read_dataset <- function(path) {
  ext <- tolower(tools::file_ext(path))
  if (ext %in% c("xlsx", "xls")) {
    df <- readxl::read_xlsx(path)
  } else {
    df <- try(
      readr::read_csv(
        path,
        locale = readr::locale(decimal_mark = ","),
        show_col_types = FALSE,
        guess_max = 1e5
      ),
      silent = TRUE
    )
    if (inherits(df, "try-error") || ncol(df) == 1) {
      df <- readr::read_delim(
        path,
        delim = ";",
        locale = readr::locale(decimal_mark = ","),
        show_col_types = FALSE,
        guess_max = 1e5
      )
    }
  }
  if (!is.data.frame(df)) stop(sprintf("Lecture impossible pour '%s'", path))
  nm <- names(df)
  nm <- trimws(nm)
  dup_blank <- nm == "" | stringr::str_detect(nm, "^\\.\\.\\d+$")
  if (any(dup_blank)) {
    df <- df[, !dup_blank, drop = FALSE]
    nm <- nm[!dup_blank]
  }
  names(df) <- nm
  df
}

charger_donnees <- function(path_remu = "data_permat1.csv",
                            path_perf = "data_permat2.csv",
                            path_conv = "data_permat3.csv",
                            path_ref = "data_permat4.csv",
                            path_demo = "data_permat5.csv") {
  list(
    remu = read_dataset(path_remu),
    perf = read_dataset(path_perf),
    conv = read_dataset(path_conv),
    ref = read_dataset(path_ref),
    demo = if (file.exists(path_demo)) read_dataset(path_demo) else NULL
  )
}

# ============================================
# Préparation des tables d'entrée
# ============================================
preparer_tables <- function(remu, perf, conv, ref, demo = NULL) {
  # --- rémunérations
  id_col <- choose_column(remu, c("id", "ID", "\ufeffid"), required = TRUE, label = "la colonne id")
  year_col <- choose_column(remu, c("annee", "Année", "Annee"), required = TRUE, label = "la colonne année")
  net_candidates <- unique(c(
    "Remu.nette.deflatee", "Remu_nette_deflatee", "Remu.nette", "remu_nette",
    "remuneration", "Remuneration", "rémunération"
  ))
  gross_candidates <- unique(c(
    "Remu.brute.deflatee", "Remu_brute_deflatee", "Remu.brute", "remu_brute",
    "Remu.deflatee", "salaire"
  ))
  fallback_candidates <- c(net_candidates, gross_candidates)
  pay_col <- choose_column(remu, fallback_candidates, required = TRUE, label = "la colonne de rémunération")
  net_col <- choose_column(remu, net_candidates, required = FALSE)
  gross_col <- choose_column(remu, gross_candidates, required = FALSE)
  
  remu_tbl <- remu %>%
    transmute(
      id__ = .data[[id_col]],
      id_key = normalize_id(.data[[id_col]]),
      annee__ = as.integer(.data[[year_col]]),
      pay_net_raw = if (!is.na(net_col)) parse_number(.data[[net_col]]) else NA_real_,
      pay_gross_raw = if (!is.na(gross_col)) parse_number(.data[[gross_col]]) else NA_real_,
      pay_base = parse_number(.data[[pay_col]])
    ) %>%
    mutate(
      pay__ = coalesce(pay_net_raw, pay_base, pay_gross_raw),
      net_pay__ = coalesce(pay_net_raw, pay_gross_raw, pay_base),
      brut_pay__ = coalesce(pay_gross_raw, pay_base, pay_net_raw)
    ) %>%
    select(id__, id_key, annee__, pay__, net_pay__, brut_pay__) %>%
    filter(!is.na(id_key) & id_key != "")
  
  # --- performances
  perf_year_col <- choose_column(perf, c("Annees.eloignant.retraite", "Années.éloignant.retraite", "Annees.eloignant.retraite."),
                                 required = TRUE, label = "la colonne années éloignant retraite")
  perf_val_col <- choose_column(perf, c("Performances.nettes", "Performance.nettes", "Perf.nettes", "perf"),
                                required = TRUE, label = "la colonne performances")
  perf_tbl <- perf %>%
    transmute(
      annees_eloign = as.integer(.data[[perf_year_col]]),
      perf_fac = parse_number(.data[[perf_val_col]])
    ) %>%
    arrange(desc(annees_eloign))
  
  # --- taux de conversion
  conv_gen_col <- choose_column(conv, c("Generation", "Génération", "generation"), required = TRUE, label = "la colonne génération")
  conv_val_col <- choose_column(conv, c("Taux.conversion", "Taux.de.conversion", "Taux.de.conversion.gen.2000", "taux"),
                                required = TRUE, label = "la colonne taux de conversion")
  conv_tbl <- conv %>%
    transmute(
      Generation = as.integer(str_extract(as.character(.data[[conv_gen_col]]), "\\d{4}")),
      Taux_conversion = parse_number(str_replace_all(.data[[conv_val_col]], "%", "")) /
        ifelse(str_detect(as.character(.data[[conv_val_col]]), "%"), 100, 1)
    ) %>%
    distinct(Generation, .keep_all = TRUE)
  
  # --- références
  ref_id_col <- choose_column(ref, c("id", "ID", "\ufeffid"), required = TRUE, label = "la colonne id (références)")
  ref_tr_col <- choose_column(ref, c("txRemplacementNetEuroConstant", "tauxRemplacementNet", "Tx.remplacement.net", "TR.net"),
                              required = TRUE, label = "la colonne taux de remplacement")
  ref_sal_col <- choose_column(ref, c(
    "avantDerniereRemuNetteTotaleAnnuellePositiveEurosConstants",
    "Avant.derniere.remu.nette", "Remu.net.avant.derniere", "net_avant_derniere"
  ), required = TRUE, label = "la colonne référence salaire net")
  ref_tbl <- ref %>%
    transmute(
      id__ = .data[[ref_id_col]],
      id_key = normalize_id(.data[[ref_id_col]]),
      tr_net = parse_number(.data[[ref_tr_col]]),
      net_avder = parse_number(.data[[ref_sal_col]]),
      generation = as.integer(str_extract(as.character(.data[[ref_id_col]]), "\\d{4}")),
      cas_type = suppressWarnings(as.integer(str_match(str_to_lower(as.character(.data[[ref_id_col]])), "cas[^0-9]*([123])")[, 2]))
    )
  
  # --- démographie
  demo_tbl <- NULL
  if (!is.null(demo)) {
    demo_year_col <- choose_column(demo, c("Annee", "Année", "annee", "year"), required = TRUE, label = "la colonne année (demo)")
    demo_pct_col <- choose_column(demo, c("Part_65ans", "Part_65", "Pct_65"), required = TRUE, label = "la colonne part 65+")
    demo_tbl <- demo %>%
      transmute(
        year = as.integer(.data[[demo_year_col]]),
        old_share = parse_number(str_replace_all(.data[[demo_pct_col]], "%", "")) /
          ifelse(str_detect(as.character(.data[[demo_pct_col]]), "%"), 100, 1) * 100
      ) %>%
      arrange(year)
  }
  
  list(remu = remu_tbl, perf = perf_tbl, conv = conv_tbl, ref = ref_tbl, demo = demo_tbl)
}

# ============================================
# Fonctions coeur du modèle
# ============================================
TMI_from_income_dyn <- function(rev, annee, base_year = ANNEE_DEBUT,
                                thresh_base = TMI_THRESH_BASE,
                                index_rate = TMI_INDEX_RATE) {
  if (is.na(rev) || is.na(annee)) return(0)
  k <- max(0L, as.integer(annee - base_year))
  idx <- (1 + index_rate)^k
  b <- thresh_base * idx
  dplyr::case_when(
    rev <= b[1] ~ 0.00,
    rev <= b[2] ~ 0.11,
    rev <= b[3] ~ 0.30,
    rev <= b[4] ~ 0.41,
    TRUE ~ 0.45
  )
}

nb_annees_cotise <- function(gen) {
  as.integer(AGE_LIQ - (ANNEE_DEBUT - gen))
}

simuler_capital <- function(remu_vec, perf_table) {
  N <- length(remu_vec)
  perf_seq <- perf_table$perf_fac[match(N:1, perf_table$annees_eloign)]
  if (any(is.na(perf_seq))) {
    stop(sprintf("La table des performances ne couvre pas l'horizon N=%d (dispo: %s)",
                 N, paste(perf_table$annees_eloign, collapse = ", ")))
  }
  cap <- 0
  for (k in seq_len(N)) {
    cap <- (cap + remu_vec[k]) * perf_seq[k]
  }
  cap
}

.simuler_unitaire_net <- function(pay_vec, perf_table, taux_conv, rvtg) {
  cap1 <- simuler_capital(pay_vec, perf_table)
  rb <- cap1 * taux_conv
  (rb - rb * FRAIS_ARERAGE) - (rb * PRORATA_RVTG * rvtg) - (rb * PRORATA_PS * TAUX_PS)
}

# ============================================
# Ajustement démographique
# ============================================
ajuster_mu_par_classe <- function(mu_fac, years, demo_series = NULL,
                                  gamma_eq = GAMMA_ACTIONS,
                                  gamma_bd = GAMMA_OBLIGATIONS) {
  if (is.null(demo_series)) return(list(mu_eq = mu_fac, mu_bd = mu_fac))
  ref_idx <- which.min(abs(years - ANNEE_DEBUT))
  ref_share <- demo_series$old_share[match(years[ref_idx], demo_series$year)]
  if (is.na(ref_share)) return(list(mu_eq = mu_fac, mu_bd = mu_fac))
  
  adjust <- function(gamma_coef) {
    sapply(years, function(y) {
      d <- demo_series$old_share[match(y, demo_series$year)]
      if (is.na(d)) return(1)
      1 + gamma_coef * (d - ref_share)
    })
  }
  
  adj_eq <- adjust(gamma_eq)
  adj_bd <- adjust(gamma_bd)
  mu_eq <- pmax(1 + (mu_fac - 1) * adj_eq, 1.000)
  mu_bd <- pmax(1 + (mu_fac - 1) * adj_bd, 1.000)
  list(mu_eq = mu_eq, mu_bd = mu_bd)
}

# ============================================
# Monte-Carlo
# ============================================
simulate_unit_samples <- function(pay_vec, perf_table, years, Mq = 50000,
                                  sigma_eq = 0.18, sigma_bd = 0.06,
                                  w_start = 0.70, w_end = 0.20,
                                  demo_series = NULL, rho = 0.10) {
  N <- length(pay_vec)
  mu_fac <- perf_table$perf_fac[match(N:1, perf_table$annees_eloign)]
  if (any(is.na(mu_fac))) stop(sprintf("perma2 ne couvre pas N=%d", N))
  
  adj <- ajuster_mu_par_classe(mu_fac, years, demo_series, GAMMA_ACTIONS, GAMMA_OBLIGATIONS)
  mu_eq <- adj$mu_eq
  mu_bd <- adj$mu_bd
  w_t <- if (N > 1) w_start + (w_end - w_start) * ((1:N - 1) / (N - 1)) else w_end
  mu_t <- w_t * mu_eq + (1 - w_t) * mu_bd
  sigma_p_t <- sqrt((w_t * sigma_eq)^2 + ((1 - w_t) * sigma_bd)^2 +
                      2 * w_t * (1 - w_t) * rho * sigma_eq * sigma_bd)
  
  res <- numeric(Mq)
  for (m in seq_len(Mq)) {
    shocks <- exp(-0.5 * sigma_p_t^2 + sigma_p_t * rnorm(N))
    fac <- mu_t * shocks
    suf <- rev(cumprod(rev(fac)))
    res[m] <- sum(pay_vec * suf)
  }
  res
}

rente_nette_from_capital <- function(cap, taux_conv, rvtg) {
  rb <- cap * taux_conv
  (rb - rb * FRAIS_ARERAGE) - (rb * PRORATA_RVTG * rvtg) - (rb * PRORATA_PS * TAUX_PS)
}

# ============================================
# Montant à combler (plein & moitié du gap)
# ============================================
safe_tr <- function(x) {
  x <- suppressWarnings(as.numeric(x))
  ifelse(is.finite(x) & x > 0 & x <= 1, x, NA_real_)
}

extract_last_net <- function(remu_tbl, mykey, years = NULL) {
  if (is.null(remu_tbl)) return(NA_real_)
  rows <- remu_tbl %>% filter(id_key == !!mykey)
  if (!is.null(years)) rows <- rows %>% filter(annee__ %in% years)
  rows <- rows %>% arrange(annee__)
  if (!nrow(rows)) return(NA_real_)
  
  take_last_positive <- function(x) {
    vals <- suppressWarnings(as.numeric(x))
    vals <- vals[is.finite(vals)]
    pos <- vals[vals > 0]
    if (length(pos)) return(tail(pos, 1))
    if (length(vals)) return(tail(vals, 1))
    numeric(0)
  }
  
  val <- take_last_positive(rows$net_pay__)
  if (!length(val)) val <- take_last_positive(rows$pay__)
  if (!length(val)) val <- take_last_positive(rows$brut_pay__)
  if (!length(val)) return(NA_real_)
  val
}

compute_montant_a_combler <- function(mykey, cas_type, generation, ref_tbl,
                                      remu_tbl = NULL, years = NULL) {
  ref_row <- ref_tbl %>% filter(id_key == !!mykey) %>% slice(1)
  tr_gen_net <- if (nrow(ref_row)) safe_tr(ref_row$tr_net[1]) else NA_real_
  net_avder_ref <- if (nrow(ref_row)) suppressWarnings(as.numeric(ref_row$net_avder[1])) else NA_real_
  
  net_avder <- extract_last_net(remu_tbl, mykey, years)
  if (!is.finite(net_avder)) net_avder <- net_avder_ref
  
  # Cible plein comblement
  if (cas_type == 3) {
    tr_objectif <- TR_OBJECTIFS_CAS3[as.character(generation)]
    if (is.na(tr_objectif)) {
      warning(sprintf("Pas d'objectif TR défini pour cas 3, génération %d. Utilisation de la logique standard.", generation))
      tr_ref_1960 <- ref_tbl %>%
        filter(cas_type == !!cas_type, generation == 1960L) %>%
        pull(tr_net) %>%
        safe_tr() %>%
        median(na.rm = TRUE)
      if (!is.finite(tr_ref_1960)) {
        tr_ref_1960 <- ref_tbl %>%
          filter(generation == 1960L) %>%
          pull(tr_net) %>%
          safe_tr() %>%
          median(na.rm = TRUE)
      }
      tr_objectif <- tr_ref_1960
    }
  } else {
    tr_objectif <- ref_tbl %>%
      filter(cas_type == !!cas_type, generation == 1960L) %>%
      pull(tr_net) %>%
      safe_tr() %>%
      median(na.rm = TRUE)
    if (!is.finite(tr_objectif)) {
      tr_objectif <- ref_tbl %>%
        filter(generation == 1960L) %>%
        pull(tr_net) %>%
        safe_tr() %>%
        median(na.rm = TRUE)
    }
    if (!is.finite(tr_objectif)) tr_objectif <- NA_real_
  }
  
  pension_actuelle <- if (!is.na(tr_gen_net) && !is.na(net_avder)) tr_gen_net * net_avder else NA_real_
  pension_cible <- if (!is.na(tr_objectif) && !is.na(net_avder)) tr_objectif * net_avder else NA_real_
  montant <- if (!is.na(pension_cible) && !is.na(pension_actuelle)) pension_cible - pension_actuelle else NA_real_
  montant <- if (!is.na(montant)) max(0, montant) else NA_real_
  
  # Cible "moitié du gap"
  tr_cible_half <- if (!is.na(tr_gen_net) && !is.na(tr_objectif)) {
    tr_gen_net + 0.5 * (tr_objectif - tr_gen_net)
  } else NA_real_
  pension_cible_half <- if (!is.na(tr_cible_half) && !is.na(net_avder)) tr_cible_half * net_avder else NA_real_
  montant_half <- if (!is.na(pension_cible_half) && !is.na(pension_actuelle)) pension_cible_half - pension_actuelle else NA_real_
  montant_half <- if (!is.na(montant_half)) max(0, montant_half) else NA_real_
  
  list(
    tr_gen_net = tr_gen_net,
    tr_objectif = tr_objectif,
    tr_cible_half = tr_cible_half,
    net_avder = net_avder,
    pension_actuelle_nette = pension_actuelle,
    pension_cible_nette = pension_cible,
    pension_cible_nette_half = pension_cible_half,
    montant_a_combler = montant,
    montant_a_combler_half = montant_half
  )
}

# ============================================
# A) Déterministe
# ============================================
simuler_id <- function(id, generation, taux_cotisation, tables) {
  N <- nb_annees_cotise(generation)
  if (N <= 0) return(tibble(id = id, generation = generation, N = N, everything = NA))
  
  cas_str <- str_match(str_to_lower(id), "cas[^0-9]*([123])")[, 2]
  if (is.na(cas_str)) cas_str <- str_extract(id, "(?<!\\d)([123])(?!\\d)")
  if (is.na(cas_str)) stop(sprintf("Cas-type non reconnu pour id='%s'.", id))
  cas_type <- as.integer(cas_str)
  rvtg <- RVTG_PAR_CAS[[as.character(cas_type)]]
  
  mykey <- normalize_id(id)
  years <- ANNEE_DEBUT + 0:(N - 1)
  pay_vec <- tables$remu %>%
    filter(id_key == mykey, annee__ %in% years) %>%
    arrange(annee__) %>%
    pull(pay__)
  if (length(pay_vec) != N) {
    stop(sprintf("Rémunérations manquantes pour id='%s' (attendu N=%d, trouvé %d).", id, N, length(pay_vec)))
  }
  
  remu_cum <- sum(pay_vec, na.rm = TRUE)
  cot_vec <- taux_cotisation * pay_vec
  cot_tot <- sum(cot_vec, na.rm = TRUE)
  taux_conv <- tables$conv$Taux_conversion[tables$conv$Generation == generation]
  if (!length(taux_conv)) stop(sprintf("Taux de conversion manquant pour génération %d", generation))
  
  cap <- simuler_capital(cot_vec, tables$perf)
  rb <- cap * taux_conv
  rnet <- (rb - rb * FRAIS_ARERAGE) - (rb * PRORATA_RVTG * rvtg) - (rb * PRORATA_PS * TAUX_PS)
  rnet_mens <- rnet / 12
  
  tmi_vec <- mapply(TMI_from_income_dyn, rev = pay_vec, annee = years)
  base_ded <- pmin(cot_vec, PLAFOND_DEFISC_PCT * pay_vec)
  defisc <- sum(base_ded * tmi_vec, na.rm = TRUE)
  taux_eff <- if (remu_cum > 0) (cot_tot - defisc) / remu_cum else NA_real_
  
  mc <- compute_montant_a_combler(mykey, cas_type, generation, tables$ref, remu_tbl = tables$remu, years = years)
  rnet_unit <- .simuler_unitaire_net(pay_vec, tables$perf, taux_conv, rvtg)
  
  t_combler <- if (!is.na(mc$montant_a_combler) && is.finite(rnet_unit) && rnet_unit > 0) mc$montant_a_combler / rnet_unit else NA_real_
  t_eff_combler <- if (!is.na(t_combler) && remu_cum > 0) {
    cot_req <- t_combler * pay_vec
    def_req <- sum(pmin(cot_req, PLAFOND_DEFISC_PCT * pay_vec) * tmi_vec, na.rm = TRUE)
    (sum(cot_req, na.rm = TRUE) - def_req) / remu_cum
  } else NA_real_
  
  t_combler_half <- if (!is.na(mc$montant_a_combler_half) && is.finite(rnet_unit) && rnet_unit > 0) mc$montant_a_combler_half / rnet_unit else NA_real_
  t_eff_combler_half <- if (!is.na(t_combler_half) && remu_cum > 0) {
    cot_req <- t_combler_half * pay_vec
    def_req <- sum(pmin(cot_req, PLAFOND_DEFISC_PCT * pay_vec) * tmi_vec, na.rm = TRUE)
    (sum(cot_req, na.rm = TRUE) - def_req) / remu_cum
  } else NA_real_
  
  tibble(
    id = id,
    cas_type = cas_type,
    generation = generation,
    N = N,
    remu_brute_deflatee_cumulee = remu_cum,
    cotisations_totales = cot_tot,
    defisc_totale = defisc,
    taux_cotisation = taux_cotisation,
    taux_cotisation_effectif = taux_eff,
    taux_conversion = taux_conv,
    rente_nette_recue = rnet,
    rente_nette_mensuelle = rnet_mens,
    tr_gen_net = mc$tr_gen_net,
    tr_objectif = mc$tr_objectif,
    tr_cible_half = mc$tr_cible_half,
    net_avant_derniere = mc$net_avder,
    pension_actuelle_nette = mc$pension_actuelle_nette,
    pension_cible_nette = mc$pension_cible_nette,
    pension_cible_nette_half = mc$pension_cible_nette_half,
    montant_a_combler = mc$montant_a_combler,
    montant_a_combler_half = mc$montant_a_combler_half,
    taux_pour_combler_ecart = t_combler,
    taux_effectif_pour_combler_ecart = t_eff_combler,
    taux_pour_moitie_ecart = t_combler_half,
    taux_effectif_pour_moitie_ecart = t_eff_combler_half
  )
}

# ============================================
# B) Monte-Carlo
# ============================================
simuler_id_MC <- function(id, generation, taux_cotisation, tables, M = 50000,
                          sigma_eq = 0.18, sigma_bd = 0.06, w_start = 0.70, w_end = 0.20,
                          alphas = c(0.01, 0.05, 0.10, 0.50), rho = 0.10) {
  N <- nb_annees_cotise(generation)
  if (N <= 0) {
    out <- tibble(id = id, cas_type = NA_integer_, generation = generation, N = N)
    cols_to_init <- c(
      "remu_brute_deflatee_cumulee", "cotisations_totales", "defisc_totale",
      "taux_cotisation", "taux_cotisation_effectif", "tr_gen_net", "tr_objectif",
      "tr_cible_half", "net_avant_derniere", "pension_actuelle_nette", "pension_cible_nette",
      "pension_cible_nette_half", "montant_a_combler", "montant_a_combler_half",
      "capital_P1", "capital_P5", "capital_P10", "capital_P50", "capital_P90", "capital_P95", "capital_P99",
      "rente_nette_P1", "rente_nette_P5", "rente_nette_P10", "rente_nette_P50", "rente_nette_P90", "rente_nette_P95", "rente_nette_P99",
      "tr_PER_P1", "tr_PER_P5", "tr_PER_P10", "tr_PER_P50", "tr_PER_P90", "tr_PER_P95", "tr_PER_P99",
      "tr_total_P1", "tr_total_P5", "tr_total_P10", "tr_total_P50", "tr_total_P90", "tr_total_P95", "tr_total_P99",
      "gain_vs_objectif_P1", "gain_vs_objectif_P5", "gain_vs_objectif_P10", "gain_vs_objectif_P50",
      "gain_vs_objectif_P90", "gain_vs_objectif_P95", "gain_vs_objectif_P99",
      "prob_succes_au_taux", "prob_succes_au_taux_half",
      "taux_pour_combler_ecart", "taux_pour_moitie_ecart"
    )
    for (nm in cols_to_init) out[[nm]] <- NA_real_
    for (a in alphas) {
      lab <- sprintf("%02d", round((1 - a) * 100))
      out[[paste0("tau_star_", lab)]] <- NA_real_
      out[[paste0("tau_effectif_star_", lab)]] <- NA_real_
      out[[paste0("tau_star_half_", lab)]] <- NA_real_
      out[[paste0("tau_effectif_star_half_", lab)]] <- NA_real_
    }
    return(out)
  }
  
  cas_str <- str_match(str_to_lower(id), "cas[^0-9]*([123])")[, 2]
  if (is.na(cas_str)) cas_str <- str_extract(id, "(?<!\\d)([123])(?!\\d)")
  if (is.na(cas_str)) stop(sprintf("Cas-type non reconnu pour id='%s'.", id))
  cas_type <- as.integer(cas_str)
  rvtg <- RVTG_PAR_CAS[[as.character(cas_type)]]
  
  mykey <- normalize_id(id)
  years <- ANNEE_DEBUT + 0:(N - 1)
  pay_vec <- tables$remu %>%
    filter(id_key == mykey, annee__ %in% years) %>%
    arrange(annee__) %>%
    pull(pay__)
  if (length(pay_vec) != N) {
    stop(sprintf("Rémunérations manquantes pour id='%s' (attendu N=%d, trouvé %d).", id, N, length(pay_vec)))
  }
  
  remu_cum <- sum(pay_vec, na.rm = TRUE)
  cot_vec <- taux_cotisation * pay_vec
  cot_tot <- sum(cot_vec, na.rm = TRUE)
  taux_conv <- tables$conv$Taux_conversion[tables$conv$Generation == generation]
  if (!length(taux_conv)) stop(sprintf("Taux de conversion manquant pour génération %d", generation))
  
  tmi_vec <- mapply(TMI_from_income_dyn, rev = pay_vec, annee = years)
  base_ded <- pmin(cot_vec, PLAFOND_DEFISC_PCT * pay_vec)
  defisc <- sum(base_ded * tmi_vec, na.rm = TRUE)
  taux_eff <- if (remu_cum > 0) (cot_tot - defisc) / remu_cum else NA_real_
  
  mc <- compute_montant_a_combler(mykey, cas_type, generation, tables$ref, remu_tbl = tables$remu, years = years)
  K_unit <- simulate_unit_samples(pay_vec, tables$perf, years, M, sigma_eq, sigma_bd, w_start, w_end, tables$demo, rho)
  rnet_u <- rente_nette_from_capital(K_unit, taux_conv, rvtg)
  
  # Distribution pour le taux courant
  caps_tau <- taux_cotisation * K_unit
  rnet_tau <- taux_cotisation * rnet_u
  quants <- c(0.01, 0.05, 0.10, 0.50, 0.90, 0.95, 0.99)
  q_cap <- quantile(caps_tau, quants, na.rm = TRUE, names = FALSE)
  q_rnt <- quantile(rnet_tau, quants, na.rm = TRUE, names = FALSE)
  
  prob_full <- if (is.na(mc$montant_a_combler)) NA_real_ else mean(rnet_tau >= mc$montant_a_combler)
  prob_half <- if (is.na(mc$montant_a_combler_half)) NA_real_ else mean(rnet_tau >= mc$montant_a_combler_half)
  
  tr_per <- if (!is.na(mc$net_avder) && mc$net_avder > 0) q_rnt / mc$net_avder else rep(NA_real_, length(q_rnt))
  names(tr_per) <- paste0("tr_PER_P", c(1, 5, 10, 50, 90, 95, 99))
  tr_total <- if (!is.na(mc$tr_gen_net)) mc$tr_gen_net + tr_per else rep(NA_real_, length(tr_per))
  names(tr_total) <- paste0("tr_total_P", c(1, 5, 10, 50, 90, 95, 99))
  gain_vs_objectif <- if (!is.na(mc$tr_objectif)) tr_total - mc$tr_objectif else rep(NA_real_, length(tr_total))
  names(gain_vs_objectif) <- paste0("gain_vs_objectif_P", c(1, 5, 10, 50, 90, 95, 99))
  
  rnet_unit_det <- .simuler_unitaire_net(pay_vec, tables$perf, taux_conv, rvtg)
  t_combler <- if (!is.na(mc$montant_a_combler) && is.finite(rnet_unit_det) && rnet_unit_det > 0) mc$montant_a_combler / rnet_unit_det else NA_real_
  
  res <- tibble(
    id = id,
    cas_type = cas_type,
    generation = generation,
    N = N,
    remu_brute_deflatee_cumulee = remu_cum,
    cotisations_totales = cot_tot,
    defisc_totale = defisc,
    taux_cotisation = taux_cotisation,
    taux_cotisation_effectif = taux_eff,
    tr_gen_net = mc$tr_gen_net,
    tr_objectif = mc$tr_objectif,
    tr_cible_half = mc$tr_cible_half,
    net_avant_derniere = mc$net_avder,
    pension_actuelle_nette = mc$pension_actuelle_nette,
    pension_cible_nette = mc$pension_cible_nette,
    pension_cible_nette_half = mc$pension_cible_nette_half,
    montant_a_combler = mc$montant_a_combler,
    montant_a_combler_half = mc$montant_a_combler_half,
    capital_P1 = q_cap[1], capital_P5 = q_cap[2], capital_P10 = q_cap[3],
    capital_P50 = q_cap[4], capital_P90 = q_cap[5], capital_P95 = q_cap[6], capital_P99 = q_cap[7],
    rente_nette_P1 = q_rnt[1], rente_nette_P5 = q_rnt[2], rente_nette_P10 = q_rnt[3],
    rente_nette_P50 = q_rnt[4], rente_nette_P90 = q_rnt[5], rente_nette_P95 = q_rnt[6], rente_nette_P99 = q_rnt[7],
    prob_succes_au_taux = prob_full,
    prob_succes_au_taux_half = prob_half,
    taux_pour_combler_ecart = t_combler
  )
  
  for (nm in names(tr_per)) res[[nm]] <- tr_per[[nm]]
  for (nm in names(tr_total)) res[[nm]] <- tr_total[[nm]]
  for (nm in names(gain_vs_objectif)) res[[nm]] <- gain_vs_objectif[[nm]]
  
  for (a in alphas) {
    lab <- sprintf("%02d", round((1 - a) * 100))
    q <- as.numeric(quantile(rnet_u, probs = a, na.rm = TRUE, names = FALSE))
    
    # Plein comblement
    tau_star <- if (!is.na(mc$montant_a_combler) && is.finite(q) && q > 0) mc$montant_a_combler / q else NA_real_
    tau_eff <- if (!is.na(tau_star) && remu_cum > 0) {
      cot_req <- tau_star * pay_vec
      def_req <- sum(pmin(cot_req, PLAFOND_DEFISC_PCT * pay_vec) * tmi_vec, na.rm = TRUE)
      (sum(cot_req, na.rm = TRUE) - def_req) / remu_cum
    } else NA_real_
    res[[paste0("tau_star_", lab)]] <- tau_star
    res[[paste0("tau_effectif_star_", lab)]] <- tau_eff
    
    # Moitié du gap
    tau_star_half <- if (!is.na(mc$montant_a_combler_half) && is.finite(q) && q > 0) mc$montant_a_combler_half / q else NA_real_
    tau_eff_half <- if (!is.na(tau_star_half) && remu_cum > 0) {
      cot_req <- tau_star_half * pay_vec
      def_req <- sum(pmin(cot_req, PLAFOND_DEFISC_PCT * pay_vec) * tmi_vec, na.rm = TRUE)
      (sum(cot_req, na.rm = TRUE) - def_req) / remu_cum
    } else NA_real_
    res[[paste0("tau_star_half_", lab)]] <- tau_star_half
    res[[paste0("tau_effectif_star_half_", lab)]] <- tau_eff_half
  }
  
  res
}

# ============================================
# Enveloppes
# ============================================
simuler_tout_MC <- function(paths = list(
  remu = "data_permat1.csv",
  perf = "data_permat2.csv",
  conv = "data_permat3.csv",
  ref = "data_permat4.csv",
  demo = "data_permat5.csv"
),
taux = c(0.1),
M = 50000,
sigma_eq = 0.18,
sigma_bd = 0.06,
w_start = 0.70,
w_end = 0.20,
alphas = c(0.01, 0.05, 0.10, 0.50),
rho = 0.10) {
  raw <- charger_donnees(paths$remu, paths$perf, paths$conv, paths$ref, paths$demo)
  tables <- preparer_tables(raw$remu, raw$perf, raw$conv, raw$ref, raw$demo)
  id_gen <- tables$remu %>%
    filter(!is.na(id__) & str_detect(id__, "\\d{4}")) %>%
    distinct(id__, id_key) %>%
    mutate(generation = suppressWarnings(as.integer(str_extract(id__, "\\d{4}"))))
  if (!nrow(id_gen)) {
    stop("Impossible de déduire la génération depuis 'id'.")
  }
  if (any(is.na(id_gen$generation))) {
    ids_fail <- id_gen$id__[is.na(id_gen$generation)]
    stop(sprintf(
      "Impossible de déduire la génération depuis 'id' (exemples: %s).",
      paste(head(ids_fail, 5), collapse = ", ")
    ))
  }
  
  grid <- tidyr::crossing(id_gen, taux_cotisation = taux)
  pmap_dfr(
    list(grid$id__, grid$generation, grid$taux_cotisation),
    ~ simuler_id_MC(..1, ..2, ..3, tables = tables, M = M, sigma_eq = sigma_eq,
                    sigma_bd = sigma_bd, w_start = w_start, w_end = w_end,
                    alphas = alphas, rho = rho)
  )
}

# ============================================
# Export Excel
# ============================================
.per_default_outfile <- function(prefix, taux) {
  base <- getwd()
  dir.create(base, recursive = TRUE, showWarnings = FALSE)
  stamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
  file.path(base, sprintf("%s_%02dpct_%s.xlsx", prefix, round(taux * 100), stamp))
}

exporter_xlsx_taux_MC <- function(taux = 0.1,
                                  paths = list(
                                    remu = "data_permat1.csv",
                                    perf = "data_permat2.csv",
                                    conv = "data_permat3.csv",
                                    ref = "data_permat4.csv",
                                    demo = "data_permat5.csv"
                                  ),
                                  outfile = NULL,
                                  M = 50000,
                                  sigma_eq = 0.18,
                                  sigma_bd = 0.06,
                                  w_start = 0.70,
                                  w_end = 0.20,
                                  alphas = c(0.01, 0.05, 0.10, 0.50),
                                  rho = 0.10) {
  res <- simuler_tout_MC(paths = paths, taux = taux, M = M, sigma_eq = sigma_eq,
                         sigma_bd = sigma_bd, w_start = w_start, w_end = w_end,
                         alphas = alphas, rho = rho)
  labs <- sprintf("%02d", round((1 - alphas) * 100))
  
  detail <- res %>% arrange(cas_type, generation, id)
  
  resume <- res %>%
    group_by(cas_type, generation) %>%
    summarise(
      remu_tot = sum(remu_brute_deflatee_cumulee, na.rm = TRUE),
      taux_cotisation = unique(taux_cotisation)[1],
      taux_cotisation_effectif_moy = mean(taux_cotisation_effectif, na.rm = TRUE),
      montant_a_combler_total = sum(montant_a_combler, na.rm = TRUE),
      montant_a_combler_half_total = sum(montant_a_combler_half, na.rm = TRUE),
      prob_succes_au_taux_moy = mean(prob_succes_au_taux, na.rm = TRUE),
      prob_succes_au_taux_half_moy = mean(prob_succes_au_taux_half, na.rm = TRUE),
      !!!setNames(
        as.list(colMeans(res[paste0("rente_nette_P", c(1, 5, 10, 50, 90, 95, 99))], na.rm = TRUE)),
        paste0("rente_P", c(1, 5, 10, 50, 90, 95, 99), "_moy")
      ),
      !!!setNames(
        as.list(colMeans(res[paste0("tr_total_P", c(1, 5, 10, 50, 90, 95, 99))], na.rm = TRUE)),
        paste0("tr_total_P", c(1, 5, 10, 50, 90, 95, 99), "_moy")
      ),
      !!!setNames(
        as.list(colMeans(res[paste0("gain_vs_objectif_P", c(1, 5, 10, 50, 90, 95, 99))], na.rm = TRUE)),
        paste0("gain_vs_objectif_P", c(1, 5, 10, 50, 90, 95, 99), "_moy")
      ),
      !!!setNames(
        as.list(colMeans(res[paste0("tau_effectif_star_", labs)], na.rm = TRUE)),
        paste0("tau_effectif_star_", labs, "_moy")
      ),
      !!!setNames(
        as.list(colMeans(res[paste0("tau_effectif_star_half_", labs)], na.rm = TRUE)),
        paste0("tau_effectif_star_half_", labs, "_moy")
      ),
      nb_ids = n(),
      .groups = "drop"
    ) %>%
    arrange(cas_type, generation)
  
  if (is.null(outfile)) {
    outfile <- .per_default_outfile("per_MC", taux)
  } else {
    dir.create(dirname(outfile), recursive = TRUE, showWarnings = FALSE)
  }
  
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "detail")
  openxlsx::writeData(wb, "detail", detail, keepNA = FALSE)
  openxlsx::addWorksheet(wb, "resume")
  openxlsx::writeData(wb, "resume", resume, keepNA = FALSE)
  
  pct <- openxlsx::createStyle(numFmt = "0.00%")
  eur <- openxlsx::createStyle(numFmt = "#,##0")
  
  pct_cols_detail <- which(names(detail) %in% c(
    "taux_cotisation", "taux_cotisation_effectif", "prob_succes_au_taux", "prob_succes_au_taux_half",
    paste0("tr_PER_P", c(1, 5, 10, 50, 90, 95, 99)),
    paste0("tr_total_P", c(1, 5, 10, 50, 90, 95, 99)),
    paste0("tau_star_", labs), paste0("tau_effectif_star_", labs),
    paste0("tau_star_half_", labs), paste0("tau_effectif_star_half_", labs)
  ))
  if (length(pct_cols_detail) && nrow(detail)) {
    openxlsx::addStyle(wb, "detail", pct, rows = 2:(nrow(detail) + 1), cols = pct_cols_detail, gridExpand = TRUE)
  }
  
  eur_cols_detail <- which(names(detail) %in% c(
    "montant_a_combler", "montant_a_combler_half",
    paste0("rente_nette_P", c(1, 5, 10, 50, 90, 95, 99))
  ))
  if (length(eur_cols_detail) && nrow(detail)) {
    openxlsx::addStyle(wb, "detail", eur, rows = 2:(nrow(detail) + 1), cols = eur_cols_detail, gridExpand = TRUE)
  }
  
  pct_cols_resume <- which(names(resume) %in% c(
    "taux_cotisation", "taux_cotisation_effectif_moy",
    "prob_succes_au_taux_moy", "prob_succes_au_taux_half_moy",
    paste0("tr_total_P", c(1, 5, 10, 50, 90, 95, 99), "_moy"),
    paste0("gain_vs_objectif_P", c(1, 5, 10, 50, 90, 95, 99), "_moy"),
    paste0("tau_effectif_star_", labs, "_moy"),
    paste0("tau_effectif_star_half_", labs, "_moy")
  ))
  if (length(pct_cols_resume) && nrow(resume)) {
    openxlsx::addStyle(wb, "resume", pct, rows = 2:(nrow(resume) + 1), cols = pct_cols_resume, gridExpand = TRUE)
  }
  
  eur_cols_resume <- which(names(resume) %in% paste0("rente_P", c(1, 5, 10, 50, 90, 95, 99), "_moy"))
  if (length(eur_cols_resume) && nrow(resume)) {
    openxlsx::addStyle(wb, "resume", eur, rows = 2:(nrow(resume) + 1), cols = eur_cols_resume, gridExpand = TRUE)
  }
  
  openxlsx::setColWidths(wb, "detail", 1:ncol(detail), "auto")
  openxlsx::setColWidths(wb, "resume", 1:ncol(resume), "auto")
  openxlsx::saveWorkbook(wb, outfile, overwrite = TRUE)
  message("Écrit : ", normalizePath(outfile, winslash = "\\", mustWork = FALSE))
  invisible(outfile)
}

# ============================================
# Exemple d'appel
# ============================================
exporter_xlsx_taux_MC(taux = 0.1)


