import math
import re
from dataclasses import dataclass
from datetime import date
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd


@dataclass
class ReorderConfig:
    coverage_days: int = 30
    max_coverage_days: int = 180
    as_of_date: Optional[date] = None

    # Se valorizzato: prova a raggiungere questo valore ordine (€)
    target_value_eur: Optional[float] = None

    # Regole costo
    cheap_cost_threshold: float = 5.0
    cheap_min_stock: int = 2              # stock minimo da tenere
    cheap_min_annual_moves: float = 2.0   # almeno 2-3 pezzi/anno per tenerlo "di base"

    expensive_cost_threshold: float = 50.0
    expensive_min_annual_moves: float = 4.0

    # Parsing GIACENZA stile "-1,000"
    # Se True interpreta "-1,000" come -1.0 (virgola=decimali, 3 decimali fissi)
    giacenza_three_decimals_style: bool = True

    # Mapping "logico" colonne (nomi standard interni)
    columns: Dict[str, str] = None

    def __post_init__(self):
        if self.as_of_date is None:
            self.as_of_date = date.today()
        if self.columns is None:
            # Nomi "standard" che useremo nel DF prodotto dal parser blindato
            self.columns = {
                "brand": "MARCA",
                "sku": "CODICE ARTICOLO",
                "desc": "DESCRIZIONE",
                "grp": "GRP. MER.",
                "scar_ac": "SCAR. AC",
                "upa": "U.P.A.",
                "listino": "LISTINO",     # opzionale (non presente nel parser blindato)
                "netto10": "NETTO 10",     # opzionale (non presente nel parser blindato)
                "scar_ap": "SCAR. AP",
                "giacenza": "GIACENZA",
                # opzionali futuri:
                "venduto_n_gg": "VENDUTO_N_GG",
                "giorni_n": "GIORNI_N",
            }
        # clamp coverage
        self.coverage_days = int(max(30, min(self.coverage_days, self.max_coverage_days)))


def _days_in_year(y: int) -> int:
    return 366 if (y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)) else 365


def _days_elapsed_in_year(d: date) -> int:
    start = date(d.year, 1, 1)
    return (d - start).days + 1  # includo il giorno corrente


def _to_float_general(value: Any) -> Optional[float]:
    """Parsa numeri in formato IT/EU in modo robusto."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if s == "" or s.lower() in {"nan", "none"}:
        return None

    s = s.replace(" ", "")

    # EU: 1.234,56
    if "," in s and "." in s:
        s2 = s.replace(".", "").replace(",", ".")
        try:
            return float(s2)
        except Exception:
            pass

    # IT semplice: 5,00
    if "," in s and "." not in s:
        try:
            return float(s.replace(",", "."))
        except Exception:
            return None

    try:
        return float(s)
    except Exception:
        m = re.search(r"-?\d+(?:[.,]\d+)?", s)
        if not m:
            return None
        return _to_float_general(m.group(0))


def _parse_giacenza(value: Any, three_decimals_style: bool) -> Optional[float]:
    """
    Gestisce il caso specifico GIACENZA stile '-1,000'.
    Se three_decimals_style=True: interpreta '-1,000' come -1.0 (3 decimali fissi).
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(" ", "")
    if s == "":
        return None

    if three_decimals_style:
        # pattern: -?\d+,\d{3} (es: -1,000)
        if re.fullmatch(r"-?\d+,\d{3}", s):
            try:
                return float(s.replace(",", "."))
            except Exception:
                return None
        # pattern: -?\d+\.\d{3}
        if re.fullmatch(r"-?\d+\.\d{3}", s):
            try:
                return float(s)
            except Exception:
                return None

    return _to_float_general(s)


def parse_input_excel_fixed_columns(file_like, config: ReorderConfig) -> pd.DataFrame:
    """
    Parser 'blindato' per export con colonne fisse:
    A=MARCA, D=CODICE ARTICOLO, F=DESCRIZIONE, G=GRP. MER., H=SCAR. AC, I=U.P.A., L=SCAR. AP, M=GIACENZA.
    Ignora i nomi colonna e prende i dati per posizione.

    Indici 0-based:
      A=0, D=3, F=5, G=6, H=7, I=8, L=11, M=12
    """
    raw = pd.read_excel(file_like, sheet_name=0, header=None, dtype=str)

    def norm(x: Any) -> str:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        return str(x).strip().lower().replace("\n", " ").replace("\r", " ")

    # Trova riga header cercando "codice" e "articolo" (robusto)
    header_row_idx = None
    for i in range(min(len(raw), 150)):
        row = [norm(v) for v in raw.iloc[i].tolist()]
        joined = " | ".join(row)
        if ("codice" in joined and "articolo" in joined):
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError("Header non trovato: non vedo una riga con 'CODICE ARTICOLO' (o simile).")

    data = raw.iloc[header_row_idx + 1 :].copy()

    # Verifica colonne minime (fino a M=12)
    if data.shape[1] <= 12:
        raise ValueError(
            f"Il foglio ha {data.shape[1]} colonne, ma ne servono almeno 13 (fino a colonna M)."
        )

    c = config.columns
    df = pd.DataFrame({
        c["brand"]: data.iloc[:, 0],    # A
        c["sku"]: data.iloc[:, 3],      # D
        c["desc"]: data.iloc[:, 5],     # F
        c["grp"]: data.iloc[:, 6],      # G
        c["scar_ac"]: data.iloc[:, 7],  # H
        c["upa"]: data.iloc[:, 8],      # I
        c["scar_ap"]: data.iloc[:, 11], # L
        c["giacenza"]: data.iloc[:, 12] # M
    })

    # Pulizia stringhe
    for k in ["brand", "sku", "desc", "grp"]:
        colname = c[k]
        df[colname] = df[colname].astype(str).fillna("").str.strip()

    # Parse numerici
    df[c["scar_ac"]] = df[c["scar_ac"]].apply(_to_float_general)
    df[c["upa"]] = df[c["upa"]].apply(_to_float_general)
    df[c["scar_ap"]] = df[c["scar_ap"]].apply(_to_float_general)
    df[c["giacenza"]] = df[c["giacenza"]].apply(
        lambda x: _parse_giacenza(x, config.giacenza_three_decimals_style)
    )

    # Drop righe senza codice articolo
    df = df[df[c["sku"]].astype(str).str.strip().ne("")].copy()

    return df


def compute_reorders(df: pd.DataFrame, config: ReorderConfig):

    warnings = []
    col = config.columns

    brand = col["brand"]
    sku = col["sku"]
    desc = col["desc"]
    grp = col["grp"]
    scar_ac = col["scar_ac"]
    scar_ap = col["scar_ap"]
    giacenza = col["giacenza"]
    upa = col["upa"]

    as_of = config.as_of_date
    days_elapsed = _days_elapsed_in_year(as_of)
    days_total = _days_in_year(as_of.year)

    peso_corrente = days_elapsed / days_total
    peso_precedente = 1 - peso_corrente

    out = df.copy()

    # ===============================
    # STIMA VENDUTO ULTIMI 12 MESI
    # ===============================

    scar_ac_val = out[scar_ac].fillna(0.0)
    scar_ap_val = out[scar_ap].fillna(0.0)

    out["venduto_12m"] = scar_ac_val + (scar_ap_val * peso_precedente)

    # domanda giornaliera su base 365
    out["daily_demand"] = out["venduto_12m"] / 365.0
    out["annual_movement_est"] = out["venduto_12m"]

    # ===============================
    # COSTO UNITARIO
    # ===============================
    out["costo_unitario"] = out[upa].fillna(0.0)

    # ===============================
    # GIACENZA
    # ===============================
    out["giacenza_input"] = out[giacenza].fillna(0.0)
    out["giacenza_effettiva"] = out["giacenza_input"].clip(lower=0.0)

    # ===============================
    # STOCK TARGET
    # ===============================
    coverage = config.coverage_days
    out["stock_target"] = out["daily_demand"] * coverage

    # Fabbisogno reale
    out["needed_float"] = out["stock_target"] - out["giacenza_effettiva"]

    # Non ordinare se < 1 pezzo
    out["qty_to_order"] = np.where(
        out["needed_float"] >= 1,
        np.ceil(out["needed_float"]),
        0,
    ).astype(int)

    # ===============================
    # REGOLE COSTO
    # ===============================

    out["categoria_regola"] = "standard"
    out["motivazione"] = ""

    cheap = out["costo_unitario"] < config.cheap_cost_threshold
    expensive = out["costo_unitario"] > config.expensive_cost_threshold

    # <5€ → tienilo in casa se almeno 2 movimenti anno
    cheap_keep = cheap & (out["annual_movement_est"] >= config.cheap_min_annual_moves)

    out.loc[cheap_keep & (out["qty_to_order"] == 0), "qty_to_order"] = config.cheap_min_stock
    out.loc[cheap_keep, "categoria_regola"] = "<5€"
    out.loc[cheap_keep, "motivazione"] = "Articolo economico: mantieni stock minimo."

    # >50€ → deve muovere almeno 4/anno
    low_moves_expensive = expensive & (out["annual_movement_est"] < config.expensive_min_annual_moves)

    out.loc[low_moves_expensive, "qty_to_order"] = 0
    out.loc[low_moves_expensive, "categoria_regola"] = ">50€"
    out.loc[low_moves_expensive, "motivazione"] = "Articolo costoso con rotazione insufficiente."

    # ===============================
    # OVERRIDE MARCA ASPL (minimo 1 pezzo a stock)
    # ===============================
    aspl_mask = out[brand].astype(str).str.strip().str.upper() == "ASPL"

    need_aspl = aspl_mask & (out["giacenza_effettiva"] < 1)

    out.loc[need_aspl, "qty_to_order"] = (
        1 - out.loc[need_aspl, "giacenza_effettiva"]
    ).apply(lambda x: int(math.ceil(float(x))) if float(x) > 0 else 0)

    out.loc[need_aspl, "categoria_regola"] = "ASPL-override"
    out.loc[need_aspl, "motivazione"] = "ASPL: mantieni sempre almeno 1 pezzo a stock."

    # ===============================
    # VALORE RIGA
    # ===============================
    out["valore_riga"] = out["qty_to_order"] * out["costo_unitario"]

    # TARGET VALORE ORDINE (€) - FAST
    # ===============================
    if config.target_value_eur is not None and float(config.target_value_eur) > 0:
        target = float(config.target_value_eur)

        total_value = float(out["valore_riga"].sum())

        if total_value < target:
            remaining = target - total_value

            cheap = out["costo_unitario"] < config.cheap_cost_threshold
            expensive = out["costo_unitario"] > config.expensive_cost_threshold
            low_moves_expensive = expensive & (out["annual_movement_est"] < config.expensive_min_annual_moves)

            dd = out["daily_demand"].fillna(0.0)
            cost = out["costo_unitario"].fillna(0.0)

            # stock attuale (giacenza + ordine base)
            stock_now = out["giacenza_effettiva"].fillna(0.0) + out["qty_to_order"].fillna(0.0)

            # massimo stock consentito per 180gg (se dd=0 -> 0)
            max_stock_180 = dd * float(config.max_coverage_days)

            # extra pezzi possibili senza superare 180gg
            max_extra_180 = (max_stock_180 - stock_now).clip(lower=0.0).fillna(0.0)
            max_extra_180_int = np.floor(max_extra_180).astype(int)

            # priorità: sotto copertura -> alta domanda -> costo economico
            coverage_now = np.where(dd > 0, stock_now / dd, np.inf)
            under_cov = (dd > 0) & (coverage_now < float(config.coverage_days))

            # candidati: escludi gonfiaggi su costosi lenti, cost=0 o dd=0
            candidates_mask = (~low_moves_expensive) & (cost > 0) & ((dd > 0) | (cheap.astype(bool)))
            candidates = out.loc[candidates_mask, ["qty_to_order", "costo_unitario"]].copy()

            if not candidates.empty:
                candidates["_under_cov"] = under_cov[candidates_mask].astype(int)
                candidates["_daily_demand"] = dd[candidates_mask]
                candidates["_cheap"] = cheap[candidates_mask].astype(int)
                candidates["_max_extra_180"] = max_extra_180_int[candidates_mask]

                # priorità: sotto copertura (desc), domanda (desc), economici (desc), costo (asc)
                candidates = candidates.sort_values(
                    by=["_under_cov", "_daily_demand", "_cheap", "costo_unitario"],
                    ascending=[False, False, False, True],
                )

                candidate_idx = candidates.index.to_numpy()
                candidate_cost = candidates["costo_unitario"].to_numpy(dtype=float)
                candidate_cap_180 = candidates["_max_extra_180"].to_numpy(dtype=int)
                qty_updates = np.zeros(len(candidates), dtype=int)
                reason_180_mask = np.zeros(len(candidates), dtype=bool)
                reason_over_mask = np.zeros(len(candidates), dtype=bool)

                # PASSO 1: riempi fino a target restando entro 180gg
                for pos in range(len(candidates)):
                    if remaining <= 0:
                        break
                    c = candidate_cost[pos]
                    if c <= 0:
                        continue
                    max_extra = candidate_cap_180[pos]
                    if max_extra <= 0:
                        continue
                    need_qty = int(math.ceil(remaining / c))
                    add_qty = min(need_qty, max_extra)
                    if add_qty <= 0:
                        continue
                    qty_updates[pos] += add_qty
                    reason_180_mask[pos] = True
                    remaining -= add_qty * c

                # PASSO 2: se ancora sotto, sfora 180gg e avvisa
                if remaining > 0:
                    for pos in range(len(candidates)):
                        if remaining <= 0:
                            break
                        c = candidate_cost[pos]
                        if c <= 0:
                            continue
                        need_qty = int(math.ceil(remaining / c))
                        if need_qty <= 0:
                            continue
                        qty_updates[pos] += need_qty
                        reason_over_mask[pos] = True
                        remaining -= need_qty * c

                    warnings.append(
                        "Target € raggiunto sforando 180 giorni su alcune righe (controlla flag_over_180)."
                    )

                if np.any(qty_updates > 0):
                    inc_series = pd.Series(qty_updates, index=candidate_idx)
                    out.loc[candidate_idx, "qty_to_order"] = (
                        out.loc[candidate_idx, "qty_to_order"].astype(int) + inc_series
                    )

                    reason_180_series = pd.Series(reason_180_mask, index=candidate_idx)
                    reason_over_series = pd.Series(reason_over_mask, index=candidate_idx)

                    idx_180 = reason_180_series[reason_180_series].index
                    idx_over = reason_over_series[reason_over_series].index

                    if len(idx_180) > 0:
                        out.loc[idx_180, "motivazione"] = (
                            out.loc[idx_180, "motivazione"].astype(str).str.strip()
                            + " + incremento per target € (<=180gg)"
                        ).str.strip(" +")
                    if len(idx_over) > 0:
                        out.loc[idx_over, "motivazione"] = (
                            out.loc[idx_over, "motivazione"].astype(str).str.strip()
                            + " + incremento per target € (sforando 180gg)"
                        ).str.strip(" +")

                out["valore_riga"] = out["qty_to_order"] * out["costo_unitario"]

    # coverage finale e flag over 180 sempre valorizzati
    dd2 = out["daily_demand"].fillna(0.0)
    stock2 = out["giacenza_effettiva"].fillna(0.0) + out["qty_to_order"].fillna(0.0)
    out["coverage_post_ordine"] = np.where(dd2 > 0, stock2 / dd2, np.nan)
    out["flag_over_180"] = (out["coverage_post_ordine"] > float(config.max_coverage_days)).fillna(False)
       # ===============================
    # MOTIVO SCARTO + OUTPUTS
    # ===============================
    out["motivo_scarto"] = ""

    # Regole di scarto (ordine di priorità)
    out.loc[out["qty_to_order"] <= 0, "motivo_scarto"] = "Nessun fabbisogno (copertura già sufficiente o needed < 1)."

    # Se esiste motivazione (es. costoso bassa rotazione) e qty=0, usala come motivo scarto
    has_motiv = out["motivazione"].astype(str).str.strip().ne("")
    out.loc[(out["qty_to_order"] <= 0) & has_motiv, "motivo_scarto"] = out.loc[
        (out["qty_to_order"] <= 0) & has_motiv, "motivazione"
    ]

    # Se daily_demand ~ 0 e non è cheap_keep / ASPL
    out.loc[(out["qty_to_order"] <= 0) & (out["daily_demand"] <= 0) & (~cheap_keep) & (~aspl_mask),
            "motivo_scarto"] = "Domanda stimata nulla negli ultimi 12 mesi."

    # Riordino e scartati
    riordino = out[out["qty_to_order"] > 0].copy()
    scartati = out[out["qty_to_order"] <= 0].copy()

    # Colonne in output (riordino)
    riordino_cols = [brand, sku, desc, grp,
                     "costo_unitario",
                     "giacenza_input", "giacenza_effettiva",
                     "venduto_12m",
                     "daily_demand",
                     "stock_target",
                     "needed_float",
                     "qty_to_order",
                     "valore_riga",
                     "coverage_post_ordine",
                     "flag_over_180",
                     "categoria_regola",
                     "motivazione"]
    riordino_cols = [c for c in riordino_cols if c in riordino.columns]
    riordino = riordino[riordino_cols]

    # Colonne in output (scartati)
    scartati_cols = [brand, sku, desc, grp,
                     "costo_unitario",
                     "giacenza_input", "giacenza_effettiva",
                     "venduto_12m",
                     "daily_demand",
                     "stock_target",
                     "needed_float",
                     "qty_to_order",
                     "coverage_post_ordine",
                     "flag_over_180",
                     "categoria_regola",
                     "motivo_scarto"]
    scartati_cols = [c for c in scartati_cols if c in scartati.columns]
    scartati = scartati[scartati_cols]

    summary = {
        "totale_righe_ordinate": int(len(riordino)),
        "totale_righe_scartate": int(len(scartati)),
        "totale_pezzi": int(riordino["qty_to_order"].sum()) if len(riordino) else 0,
        "totale_valore": float(riordino["valore_riga"].sum()) if len(riordino) else 0.0,
        "coverage_days_usato": int(config.coverage_days),
        "as_of_date": str(as_of),
    }

    return riordino, scartati, summary, warnings

def export_to_excel(df_riordino: pd.DataFrame,
                    df_scartati: pd.DataFrame,
                    summary: Dict[str, Any],
                    warnings: List[str],
                    output_path: str) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_riordino.to_excel(writer, index=False, sheet_name="riordino")
        df_scartati.to_excel(writer, index=False, sheet_name="scartati")

        summary_rows = [{"campo": k, "valore": v} for k, v in summary.items()]
        if warnings:
            summary_rows.append({"campo": "warnings", "valore": " | ".join(warnings)})

        pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name="summary")
