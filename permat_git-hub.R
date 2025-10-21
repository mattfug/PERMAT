# ============================================
# PERMAT – Modèle déterministe & Monte Carlo
# Exports Excel dans le dossier du projet
# (avec ajustement démographique par classe d'actif)
# ============================================
# --- Packages ---
suppressPackageStartupMessages({
  library(dplyr); library(readr); library(stringr); library(purrr)
  library(tidyr); library(tibble); library(openxlsx); library(readxl)
})
# ============================================
# Hypothèses fixes
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
# Coefficients d'ajustement démographique par classe d'actif
# (réduction de rendement par point de % de pop 65+)
GAMMA_ACTIONS <- -0.005 # –0,5% par point de % d'augmentation
GAMMA_OBLIGATIONS <- -0.010 # –1% par point de % d'augmentation
TMI_from_income_dyn <- function(rev, annee, base_year = ANNEE_DEBUT,
                                thresh_base = TMI_THRESH_BASE,
                                index_rate = TMI_INDEX_RATE) {
  if (is.na(rev) || is.na(annee)) return(0)
  k <- max(0L, as.integer(annee - base_year))
  idx <- (1 + index_rate)^k
  b <- thresh_base * idx
  dplyr::case_when(
    rev <= b[1] ~ 0.00, rev <= b[2] ~ 0.11, rev <= b[3] ~ 0.30,
    rev <= b[4] ~ 0.41, TRUE ~ 0.45
  )
}
normalize_id <- function(x){
  x %>% as.character() %>% str_replace_all("\\s+", " ") %>%
    str_trim() %>% str_to_lower()
}
.read_any <- function(path){
  ext <- tolower(tools::file_ext(path))
  if (ext %in% c("xlsx","xls")) {
    df <- readxl::read_xlsx(path)
  } else {
    df <- try(readr::read_csv(path, locale = readr::locale(decimal_mark=","),
                              show_col_types=FALSE, guess_max=1e5), silent = TRUE)
    if (inherits(df,"try-error") || ncol(df) == 1) {
      df <- readr::read_delim(path, delim=";",
                              locale = readr::locale(decimal_mark=","),
                              show_col_types=FALSE, guess_max=1e5)
    }
  }
  df
}
charger_donnees <- function(path_remu = "data_permat1.csv",
                            path_perf = "data_permat2.csv",
                            path_conv = "data_permat3.csv",
                            path_ref = "data_permat4.csv",
                            path_demo = "data_permat5.csv") {
  list(
    remu = .read_any(path_remu),
    perf = .read_any(path_perf),
    conv = .read_any(path_conv),
    ref = .read_any(path_ref),
    demo = if (file.exists(path_demo)) .read_any(path_demo) else NULL
  )
}
col_or_stop <- function(df, candidates){
  hit <- candidates[candidates %in% names(df)]
  if (!length(hit)) stop(sprintf(
    "Colonnes attendues non trouvées. Essais: %s. Présentes: %s",
    paste(candidates, collapse=", "), paste(names(df), collapse=", ")))
  hit[[1]]
}
preparer_tables <- function(remu, perf, conv, ref, demo=NULL){
  # --- perma1 (rémus) ---
  idc <- col_or_stop(remu, c("id","ID","\ufeffid"))
  anc <- col_or_stop(remu, c("annee","Année","Annee"))
  payc <- col_or_stop(remu, c("Remu.brute.deflatee","Remu.deflatee","remuneration",
                              "rémunération","Remuneration","salaire"))
  remu2 <- remu %>% mutate(
    id__ = .data[[idc]], id_key = normalize_id(.data[[idc]]),
    annee__ = as.integer(.data[[anc]]),
    pay__ = suppressWarnings(as.numeric(str_replace_all(as.character(.data[[payc]]), ",", ".")))
  ) %>% select(id__, id_key, annee__, pay__)
  # --- perma2 (perf nettes multiplicatives) ---
  a_er <- col_or_stop(perf, c("Annees.eloignant.retraite","Années.éloignant.retraite",
                              "Annees.eloignant.retraite."))
  pcol <- col_or_stop(perf, c("Performances.nettes","Performance.nettes","Perf.nettes","perf"))
  perf2 <- perf %>% transmute(
    annees_eloign = as.integer(.data[[a_er]]),
    perf_fac = suppressWarnings(as.numeric(str_replace_all(as.character(.data[[pcol]]), ",",".")))
  ) %>% arrange(desc(annees_eloign))
  # --- perma3 (taux de conversion) ---
  gcol <- col_or_stop(conv, c("Generation","Génération","generation"))
  tcol <- col_or_stop(conv, c("Taux.conversion","Taux.de.conversion",
                              "Taux.de.conversion.gen.2000","taux"))
  conv2 <- conv %>% transmute(
    Generation = as.integer(str_extract(as.character(.data[[gcol]]), "\\d{4}")),
    Taux_conversion = suppressWarnings(as.numeric(
      str_replace_all(str_replace_all(as.character(.data[[tcol]]), "%",""), ",",".")
    )) / ifelse(str_detect(as.character(.data[[tcol]]), "%"), 100, 1)
  ) %>% distinct(Generation, .keep_all = TRUE)
  # --- perma4 (réfs TR & net avant-dernière) ---
  id4 <- col_or_stop(ref, c("id","ID","\ufeffid"))
  tr4 <- col_or_stop(ref, c("txRemplacementNetEuroConstant","tauxRemplacementNet",
                            "Tx.remplacement.net","TR.net"))
  na4 <- col_or_stop(ref, c("avantDerniereRemuNetteTotaleAnnuellePositiveEurosConstants",
                            "Avant.derniere.remu.nette","Remu.net.avant.derniere",
                            "net_avant_derniere"))
  ref2 <- ref %>% transmute(
    id__ = .data[[id4]], id_key = normalize_id(.data[[id4]]),
    tr_net = suppressWarnings(as.numeric(str_replace_all(as.character(.data[[tr4]]), ",","."))),
    net_avder = suppressWarnings(as.numeric(str_replace_all(as.character(.data[[na4]]), ",","."))),
    generation = as.integer(str_extract(as.character(.data[[id4]]), "\\d{4}")),
    cas_type = suppressWarnings(as.integer(str_match(tolower(as.character(.data[[id4]])),
                                                     "cas[^0-9]*([123])")[,2]))
  )
  # --- perma5 (démographie) ---
  demo2 <- NULL
  if (!is.null(demo)) {
    anc_d <- col_or_stop(demo, c("Annee","Année","annee","year"))
    prc_d <- col_or_stop(demo, c("Part_65ans","Part_65","Pct_65"))
    demo2 <- demo %>% transmute(
      year = as.integer(.data[[anc_d]]),
      old_share = suppressWarnings(as.numeric(
        str_replace_all(str_replace_all(as.character(.data[[prc_d]]), "%",""), ",",".")
      )) / ifelse(str_detect(as.character(.data[[prc_d]]), "%"), 100, 1) * 100
    ) %>% arrange(year)
  }
  list(remu=remu2, perf=perf2, conv=conv2, ref=ref2, demo=demo2)
}
# ============================================
# Mécanique retraite
# ============================================
nb_annees_cotise <- function(gen) as.integer(65 - (ANNEE_DEBUT - gen))
simuler_capital <- function(remu_vec, perf_table){
  N <- length(remu_vec)
  perf_seq <- perf_table$perf_fac[match(N:1, perf_table$annees_eloign)]
  if (any(is.na(perf_seq))) stop(sprintf(
    "La table des performances ne couvre pas l'horizon N=%d (dispo: %s)",
    N, paste(perf_table$annees_eloign, collapse=", ")))
  cap <- 0
  for (k in seq_len(N)) cap <- (cap + remu_vec[k]) * perf_seq[k]
  cap
}
.simuler_unitaire_net <- function(pay_vec, perf_table, taux_conv, rvtg){
  cap1 <- simuler_capital(pay_vec, perf_table)
  rb <- cap1 * taux_conv
  (rb - rb * FRAIS_ARERAGE) - (rb * PRORATA_RVTG * rvtg) - (rb * PRORATA_PS * TAUX_PS)
}
# ============================================
# Ajustement démographique par classe d'actif
# ============================================
ajuster_mu_par_classe <- function(mu_fac, years, demo_series=NULL,
                                  gamma_eq=GAMMA_ACTIONS,
                                  gamma_bd=GAMMA_OBLIGATIONS){
  if (is.null(demo_series)) return(list(mu_eq=mu_fac, mu_bd=mu_fac))
  ref_idx <- which.min(abs(years - ANNEE_DEBUT))
  ref_share <- demo_series$old_share[match(years[ref_idx], demo_series$year)]
  if (is.na(ref_share)) return(list(mu_eq=mu_fac, mu_bd=mu_fac))
  adj_eq <- sapply(years, function(y){
    d <- demo_series$old_share[match(y, demo_series$year)]
    if (is.na(d)) return(1)
    delta_share <- (d - ref_share)
    1 + gamma_eq * delta_share
  })
  adj_bd <- sapply(years, function(y){
    d <- demo_series$old_share[match(y, demo_series$year)]
    if (is.na(d)) return(1)
    delta_share <- (d - ref_share)
    1 + gamma_bd * delta_share
  })
  mu_eq <- pmax(1 + (mu_fac - 1) * adj_eq, 1.000)
  mu_bd <- pmax(1 + (mu_fac - 1) * adj_bd, 1.000)
  list(mu_eq=mu_eq, mu_bd=mu_bd)
}
# ============================================
# Monte-Carlo : séparation actions/obligations
# ============================================
simulate_unit_samples <- function(pay_vec, perf_table, years, Mq=100000,
                                  sigma_eq=0.18, sigma_bd=0.06,
                                  w_start=0.70, w_end=0.20,
                                  demo_series=NULL, rho=0.10){
  N <- length(pay_vec)
  mu_fac <- perf_table$perf_fac[match(N:1, perf_table$annees_eloign)]
  if (any(is.na(mu_fac))) stop(sprintf("perma2 ne couvre pas N=%d", N))
  adj_list <- ajuster_mu_par_classe(mu_fac, years, demo_series,
                                    GAMMA_ACTIONS, GAMMA_OBLIGATIONS)
  mu_eq <- adj_list$mu_eq
  mu_bd <- adj_list$mu_bd
  w_t <- if (N>1) w_start + (w_end - w_start) * ((1:N - 1)/(N - 1)) else w_end
  mu_t <- w_t * mu_eq + (1 - w_t) * mu_bd
  sigma_p_t <- sqrt((w_t * sigma_eq)^2 + ((1 - w_t) * sigma_bd)^2 +
                      2 * w_t * (1 - w_t) * rho * sigma_eq * sigma_bd)
  res <- numeric(Mq)
  for (m in seq_len(Mq)){
    shocks <- exp(-0.5 * sigma_p_t^2 + sigma_p_t * rnorm(N))
    fac <- mu_t * shocks
    suf <- rev(cumprod(rev(fac)))
    res[m] <- sum(pay_vec * suf)
  }
  res
}
rente_nette_from_capital <- function(cap, taux_conv, rvtg){
  rb <- cap * taux_conv
  (rb - rb * FRAIS_ARERAGE) - (rb * PRORATA_RVTG * rvtg) - (rb * PRORATA_PS * TAUX_PS)
}
# ============================================
# Montant à combler
# ============================================
safe_tr <- function(x){
  x <- suppressWarnings(as.numeric(x))
  ifelse(is.finite(x) & x > 0 & x <= 1, x, NA_real_)
}
compute_montant_a_combler <- function(mykey, cas_type, generation, ref_tbl){
  ref_id <- ref_tbl %>% filter(id_key == !!mykey) %>% slice(1)
  tr_gen_net <- if (nrow(ref_id)) safe_tr(ref_id$tr_net[1]) else NA_real_
  net_avder <- if (nrow(ref_id)) suppressWarnings(as.numeric(ref_id$net_avder[1])) else NA_real_
  tr_ref_1960 <- ref_tbl %>% filter(cas_type == !!cas_type, generation == 1960L) %>%
    pull(tr_net) %>% safe_tr() %>% median(na.rm=TRUE)
  if (!is.finite(tr_ref_1960)) {
    tr_ref_1960 <- ref_tbl %>% filter(generation == 1960L) %>%
      pull(tr_net) %>% safe_tr() %>% median(na.rm=TRUE)
  }
  if (!is.finite(tr_ref_1960)) tr_ref_1960 <- NA_real_
  pension_actuelle <- if (!is.na(tr_gen_net) && !is.na(net_avder))
    tr_gen_net * net_avder else NA_real_
  pension_cible <- if (!is.na(tr_ref_1960) && !is.na(net_avder))
    tr_ref_1960 * net_avder else NA_real_
  montant <- if (!is.na(pension_cible) && !is.na(pension_actuelle))
    pension_cible - pension_actuelle else NA_real_
  montant <- if (!is.na(montant)) max(0, montant) else NA_real_
  list(tr_gen_net = tr_gen_net, tr_ref_1960 = tr_ref_1960, net_avder = net_avder,
       pension_actuelle_nette = pension_actuelle, pension_cible_nette = pension_cible,
       montant_a_combler = montant)
}
# ============================================
# A) Déterministe
# ============================================
simuler_id <- function(id, generation, taux_cotisation, tables){
  N <- nb_annees_cotise(generation)
  if (N <= 0) return(tibble(id=id, generation=generation, N=N, everything=NA))
  cas_str <- str_match(tolower(id), "cas[^0-9]*([123])")[,2]
  if (is.na(cas_str)) cas_str <- str_extract(id, "(?<!\\d)([123])(?!\\d)")
  if (is.na(cas_str)) stop(sprintf("Cas-type non reconnu pour id='%s'.", id))
  cas_type <- as.integer(cas_str)
  rvtg <- RVTG_PAR_CAS[[as.character(cas_type)]]
  mykey <- normalize_id(id)
  years <- ANNEE_DEBUT + 0:(N-1)
  pay_vec <- tables$remu %>% filter(id_key == mykey, annee__ %in% years) %>%
    arrange(annee__) %>% pull(pay__)
  if (length(pay_vec) != N) stop(sprintf(
    "Rémunérations manquantes pour id='%s' (attendu N=%d, trouvé %d).",
    id, N, length(pay_vec)))
  remu_cum <- sum(pay_vec, na.rm=TRUE)
  cot_vec <- taux_cotisation * pay_vec
  cot_tot <- sum(cot_vec, na.rm=TRUE)
  taux_conv <- tables$conv$Taux_conversion[tables$conv$Generation == generation]
  if (!length(taux_conv)) stop(sprintf("Taux de conversion manquant pour génération %d", generation))
  cap <- simuler_capital(cot_vec, tables$perf)
  rb <- cap * taux_conv
  rnet <- (rb - rb * FRAIS_ARERAGE) - (rb * PRORATA_RVTG * rvtg) - (rb * PRORATA_PS * TAUX_PS)
  rnet_mens <- rnet/12
  tmi_vec <- if (cas_type %in% c(1L,2L))
    mapply(TMI_from_income_dyn, rev=pay_vec, annee=years) else rep(0,N)
  base_ded <- pmin(cot_vec, PLAFOND_DEFISC_PCT * pay_vec)
  defisc <- sum(base_ded * tmi_vec, na.rm=TRUE)
  taux_eff <- if (remu_cum>0) (cot_tot - defisc)/remu_cum else NA_real_
  mc <- compute_montant_a_combler(mykey, cas_type, generation, tables$ref)
  rnet_unit <- .simuler_unitaire_net(pay_vec, tables$perf, taux_conv, rvtg)
  t_combler <- if (!is.na(mc$montant_a_combler) && is.finite(rnet_unit) && rnet_unit>0)
    mc$montant_a_combler / rnet_unit else NA_real_
  t_eff_combler <- if (!is.na(t_combler) && remu_cum>0){
    cot_req <- t_combler * pay_vec
    def_req <- sum(pmin(cot_req, PLAFOND_DEFISC_PCT*pay_vec) * tmi_vec, na.rm=TRUE)
    (sum(cot_req, na.rm=TRUE) - def_req) / remu_cum
  } else NA_real_
  tibble(
    id=id, cas_type=cas_type, generation=generation, N=N,
    remu_brute_deflatee_cumulee=remu_cum, cotisations_totales=cot_tot,
    defisc_totale=defisc, taux_cotisation=taux_cotisation,
    taux_cotisation_effectif=taux_eff, taux_conversion=taux_conv,
    rente_nette_recue=rnet, rente_nette_mensuelle=rnet_mens,
    tr_gen_net=mc$tr_gen_net, tr_ref_1960=mc$tr_ref_1960,
    net_avant_derniere=mc$net_avder, pension_actuelle_nette=mc$pension_actuelle_nette,
    pension_cible_nette=mc$pension_cible_nette, montant_a_combler=mc$montant_a_combler,
    taux_pour_combler_ecart=t_combler, taux_effectif_pour_combler_ecart=t_eff_combler
  )
}
# ============================================
# B) Monte-Carlo (avec démo)
# ============================================
simuler_id_MC <- function(id, generation, taux_cotisation, tables, M=100000,
                          sigma_eq=0.18, sigma_bd=0.06, w_start=0.70, w_end=0.20,
                          alphas=c(0.01,0.05,0.10), rho=0.10){
  N <- nb_annees_cotise(generation)
  if (N <= 0) {
    out <- tibble(id=id, cas_type=NA_integer_, generation=generation, N=N)
    for (nm in c("remu_brute_deflatee_cumulee","cotisations_totales","defisc_totale",
                 "taux_cotisation","taux_cotisation_effectif","tr_gen_net","tr_ref_1960",
                 "net_avant_derniere","pension_actuelle_nette","pension_cible_nette",
                 "montant_a_combler",
                 "capital_P1","capital_P5","capital_P10","capital_P50","capital_P90","capital_P95","capital_P99",
                 "rente_nette_P1","rente_nette_P5","rente_nette_P10","rente_nette_P50","rente_nette_P90","rente_nette_P95","rente_nette_P99",
                 "tr_PER_P1","tr_PER_P5","tr_PER_P10","tr_PER_P50","tr_PER_P90","tr_PER_P95","tr_PER_P99",
                 "tr_total_P1","tr_total_P5","tr_total_P10","tr_total_P50","tr_total_P90","tr_total_P95","tr_total_P99",
                 "gain_vs_ref_1960_P1","gain_vs_ref_1960_P5","gain_vs_ref_1960_P10","gain_vs_ref_1960_P50",
                 "gain_vs_ref_1960_P90","gain_vs_ref_1960_P95","gain_vs_ref_1960_P99",
                 "prob_succes_au_taux","taux_pour_combler_ecart")) out[[nm]] <- NA_real_
    for (a in alphas){
      lab <- sprintf("%02d", round((1-a)*100))
      out[[paste0("tau_star_", lab)]] <- NA_real_
      out[[paste0("tau_effectif_star_", lab)]] <- NA_real_
    }
    return(out)
  }
  cas_str <- str_match(tolower(id), "cas[^0-9]*([123])")[,2]
  if (is.na(cas_str)) cas_str <- str_extract(id, "(?<!\\d)([123])(?!\\d)")
  if (is.na(cas_str)) stop(sprintf("Cas-type non reconnu pour id='%s'.", id))
  cas_type <- as.integer(cas_str)
  rvtg <- RVTG_PAR_CAS[[as.character(cas_type)]]
  mykey <- normalize_id(id)
  years <- ANNEE_DEBUT + 0:(N-1)
  pay_vec <- tables$remu %>% filter(id_key==mykey, annee__ %in% years) %>%
    arrange(annee__) %>% pull(pay__)
  if (length(pay_vec) != N) stop(sprintf(
    "Rémunérations manquantes pour id='%s' (attendu N=%d, trouvé %d).",
    id, N, length(pay_vec)))
  remu_cum <- sum(pay_vec, na.rm=TRUE)
  cot_vec <- taux_cotisation * pay_vec
  cot_tot <- sum(cot_vec, na.rm=TRUE)
  taux_conv <- tables$conv$Taux_conversion[tables$conv$Generation == generation]
  if (!length(taux_conv)) stop(sprintf("Taux de conversion manquant pour génération %d", generation))
  tmi_vec <- if (cas_type %in% c(1L,2L))
    mapply(TMI_from_income_dyn, rev=pay_vec, annee=years) else rep(0,N)
  base_ded <- pmin(cot_vec, PLAFOND_DEFISC_PCT*pay_vec)
  defisc <- sum(base_ded * tmi_vec, na.rm=TRUE)
  taux_eff <- if (remu_cum>0) (cot_tot - defisc)/remu_cum else NA_real_
  mc <- compute_montant_a_combler(mykey, cas_type, generation, tables$ref)
  K_unit <- simulate_unit_samples(pay_vec, tables$perf, years, M, sigma_eq, sigma_bd,
                                  w_start, w_end, tables$demo, rho)
  rnet_u <- rente_nette_from_capital(K_unit, taux_conv, rvtg)
  caps_tau <- taux_cotisation * K_unit
  rnet_tau <- taux_cotisation * rnet_u
  quants <- c(.01, .05, .10, .50, .90, .95, .99)
  q_cap <- quantile(caps_tau, quants, na.rm=TRUE, names=FALSE)
  q_rnt <- quantile(rnet_tau, quants, na.rm=TRUE, names=FALSE)
  prob <- if (is.na(mc$montant_a_combler)) NA_real_ else mean(rnet_tau >= mc$montant_a_combler)
  tr_per <- if (!is.na(mc$net_avder) && mc$net_avder > 0)
    q_rnt / mc$net_avder else rep(NA_real_, length(q_rnt))
  names(tr_per) <- paste0("tr_PER_P", c(1,5,10,50,90,95,99))
  tr_total <- if (!is.na(mc$tr_gen_net)) mc$tr_gen_net + tr_per else rep(NA_real_, length(tr_per))
  names(tr_total) <- paste0("tr_total_P", c(1,5,10,50,90,95,99))
  gain_vs_1960 <- if (!is.na(mc$tr_ref_1960)) tr_total - mc$tr_ref_1960 else rep(NA_real_, length(tr_total))
  names(gain_vs_1960) <- paste0("gain_vs_ref_1960_P", c(1,5,10,50,90,95,99))
  rnet_unit_det <- .simuler_unitaire_net(pay_vec, tables$perf, taux_conv, rvtg)
  t_combler <- if (!is.na(mc$montant_a_combler) && is.finite(rnet_unit_det) && rnet_unit_det>0)
    mc$montant_a_combler / rnet_unit_det else NA_real_
  res <- tibble(
    id=id, cas_type=cas_type, generation=generation, N=N,
    remu_brute_deflatee_cumulee=remu_cum, cotisations_totales=cot_tot,
    defisc_totale=defisc, taux_cotisation=taux_cotisation,
    taux_cotisation_effectif=taux_eff, tr_gen_net=mc$tr_gen_net,
    tr_ref_1960=mc$tr_ref_1960, net_avant_derniere=mc$net_avder,
    pension_actuelle_nette=mc$pension_actuelle_nette,
    pension_cible_nette=mc$pension_cible_nette,
    montant_a_combler=mc$montant_a_combler,
    capital_P1=q_cap[1], capital_P5=q_cap[2], capital_P10=q_cap[3],
    capital_P50=q_cap[4], capital_P90=q_cap[5], capital_P95=q_cap[6],
    capital_P99=q_cap[7],
    rente_nette_P1=q_rnt[1], rente_nette_P5=q_rnt[2], rente_nette_P10=q_rnt[3],
    rente_nette_P50=q_rnt[4], rente_nette_P90=q_rnt[5], rente_nette_P95=q_rnt[6],
    rente_nette_P99=q_rnt[7],
    prob_succes_au_taux=prob, taux_pour_combler_ecart=t_combler
  )
  for (nm in names(tr_per)) res[[nm]] <- tr_per[[nm]]
  for (nm in names(tr_total)) res[[nm]] <- tr_total[[nm]]
  for (nm in names(gain_vs_1960)) res[[nm]] <- gain_vs_1960[[nm]]
  for (a in alphas){
    lab <- sprintf("%02d", round((1-a)*100))
    q <- as.numeric(quantile(rnet_u, probs=a, na.rm=TRUE, names=FALSE))
    tau_star <- if (!is.na(mc$montant_a_combler) && is.finite(q) && q>0)
      mc$montant_a_combler/q else NA_real_
    tau_eff <- if (!is.na(tau_star) && remu_cum>0){
      cot_req <- tau_star * pay_vec
      def_req <- sum(pmin(cot_req, PLAFOND_DEFISC_PCT*pay_vec) * tmi_vec, na.rm=TRUE)
      (sum(cot_req, na.rm=TRUE) - def_req) / remu_cum
    } else NA_real_
    res[[paste0("tau_star_", lab)]] <- tau_star
    res[[paste0("tau_effectif_star_", lab)]] <- tau_eff
  }
  res
}
# ============================================
# Enveloppes
# ============================================
simuler_tout_MC <- function(paths=list(remu="data_permat1.csv", perf="data_permat2.csv",
                                       conv="data_permat3.csv", ref="data_permat4.csv",
                                       demo="data_permat5.csv"),
                            taux=c(0.105), M=100000, sigma_eq=0.18, sigma_bd=0.06,
                            w_start=0.70, w_end=0.20, alphas=c(0.01,0.05,0.10), rho=0.10){
  raw <- charger_donnees(paths$remu, paths$perf, paths$conv, paths$ref, paths$demo)
  tab <- preparer_tables(raw$remu, raw$perf, raw$conv, raw$ref, raw$demo)
  id_gen <- tab$remu %>% distinct(id__, id_key) %>%
    mutate(generation=as.integer(str_extract(id__, "\\d{4}")))
  if (any(is.na(id_gen$generation))) stop("Impossible de déduire la génération depuis 'id'.")
  grid <- tidyr::crossing(id_gen, taux_cotisation=taux)
  pmap_dfr(list(grid$id__, grid$generation, grid$taux_cotisation),
           ~ simuler_id_MC(..1, ..2, ..3, tables=tab, M=M, sigma_eq=sigma_eq,
                           sigma_bd=sigma_bd, w_start=w_start, w_end=w_end,
                           alphas=alphas, rho=rho))
}
# ============================================
# Export Excel
# ============================================
.per_default_outfile <- function(prefix, taux){
  base <- getwd()
  dir.create(base, recursive = TRUE, showWarnings = FALSE)
  stamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
  file.path(base, sprintf("%s_%02dpct_%s.xlsx", prefix, round(taux*100), stamp))
}
exporter_xlsx_taux_MC <- function(taux=0.105,
                                  paths=list(remu="data_permat1.csv", perf="data_permat2.csv",
                                             conv="data_permat3.csv", ref="data_permat4.csv",
                                             demo="data_permat5.csv"),
                                  outfile=NULL, M=100000, sigma_eq=0.18, sigma_bd=0.06,
                                  w_start=0.70, w_end=0.20, alphas=c(0.01,0.05,0.10), rho=0.10){
  res <- simuler_tout_MC(paths=paths, taux=taux, M=M, sigma_eq=sigma_eq,
                         sigma_bd=sigma_bd, w_start=w_start, w_end=w_end,
                         alphas=alphas, rho=rho)
  labs <- sprintf("%02d", round((1-alphas)*100))
  tau_cols <- c(paste0("tau_star_", labs), paste0("tau_effectif_star_", labs))
  detail <- res %>% arrange(cas_type, generation, id)
  resume <- res %>% group_by(cas_type, generation) %>%
    summarise(
      remu_tot = sum(remu_brute_deflatee_cumulee, na.rm=TRUE),
      taux_cotisation = unique(taux_cotisation)[1],
      taux_cotisation_effectif_moy = mean(taux_cotisation_effectif, na.rm=TRUE),
      montant_a_combler_total = sum(montant_a_combler, na.rm=TRUE),
      prob_succes_au_taux_moy = mean(prob_succes_au_taux, na.rm=TRUE),
      !!!setNames(as.list(colMeans(res[paste0("rente_nette_P", c(1,5,10,50,90,95,99))], na.rm=TRUE)),
                  paste0("rente_P", c(1,5,10,50,90,95,99), "_moy")),
      !!!setNames(as.list(colMeans(res[paste0("tr_total_P", c(1,5,10,50,90,95,99))], na.rm=TRUE)),
                  paste0("tr_total_P", c(1,5,10,50,90,95,99), "_moy")),
      !!!setNames(as.list(colMeans(res[paste0("gain_vs_ref_1960_P", c(1,5,10,50,90,95,99))], na.rm=TRUE)),
                  paste0("gain_vs_ref_1960_P", c(1,5,10,50,90,95,99), "_moy")),
      !!!setNames(as.list(colMeans(res[paste0("tau_effectif_star_", labs)], na.rm=TRUE)),
                  paste0("tau_effectif_star_", labs, "_moy")),
      nb_ids = n(), .groups="drop"
    ) %>% arrange(cas_type, generation)
  if (is.null(outfile)) outfile <- .per_default_outfile("per_MC", taux)
  else dir.create(dirname(outfile), recursive = TRUE, showWarnings = FALSE)
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb,"detail")
  openxlsx::writeData(wb,"detail",detail,keepNA=FALSE)
  openxlsx::addWorksheet(wb,"resume")
  openxlsx::writeData(wb,"resume",resume,keepNA=FALSE)
  pct <- openxlsx::createStyle(numFmt="0.00%")
  eur <- openxlsx::createStyle(numFmt="#,##0")
  col_pct_d <- which(names(detail) %in% c(
    "taux_cotisation","taux_cotisation_effectif","prob_succes_au_taux",
    paste0("tr_PER_P", c(1,5,10,50,90,95,99)),
    paste0("tr_total_P", c(1,5,10,50,90,95,99)),
    paste0("tau_star_", labs), paste0("tau_effectif_star_", labs)
  ))
  if (length(col_pct_d) && nrow(detail))
    openxlsx::addStyle(wb,"detail",pct,rows=2:(nrow(detail)+1),cols=col_pct_d,gridExpand=TRUE)
  col_eur_d <- which(names(detail) %in% c(
    "montant_a_combler", paste0("rente_nette_P", c(1,5,10,50,90,95,99))
  ))
  if (length(col_eur_d) && nrow(detail))
    openxlsx::addStyle(wb,"detail",eur,rows=2:(nrow(detail)+1),cols=col_eur_d,gridExpand=TRUE)
  col_pct_r <- which(names(resume) %in% c(
    "taux_cotisation","taux_cotisation_effectif_moy",
    paste0("tr_total_P", c(1,5,10,50,90,95,99), "_moy"),
    paste0("gain_vs_ref_1960_P", c(1,5,10,50,90,95,99), "_moy"),
    paste0("tau_effectif_star_", labs, "_moy")
  ))
  if (length(col_pct_r) && nrow(resume))
    openxlsx::addStyle(wb,"resume",pct,rows=2:(nrow(resume)+1),cols=col_pct_r,gridExpand=TRUE)
  col_eur_r <- which(names(resume) %in% paste0("rente_P", c(1,5,10,50,90,95,99), "_moy"))
  if (length(col_eur_r) && nrow(resume))
    openxlsx::addStyle(wb,"resume",eur,rows=2:(nrow(resume)+1),cols=col_eur_r,gridExpand=TRUE)
  openxlsx::setColWidths(wb,"detail",1:ncol(detail),"auto")
  openxlsx::setColWidths(wb,"resume",1:ncol(resume),"auto")
  openxlsx::saveWorkbook(wb,outfile,overwrite=TRUE)
  message("Écrit : ", normalizePath(outfile, winslash="\\", mustWork = FALSE))
  invisible(outfile)
}
# ============================================
# Exemples d'appel
# ============================================
# -> ./per_MC_10pct_YYYYMMDD_HHMMSS.xlsx
exporter_xlsx_taux_MC(taux = 0.105)