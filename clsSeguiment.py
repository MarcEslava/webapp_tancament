import os
import datetime
import pandas as pd


RAPPEL_DIR = r"Z:\Compres\INDÚSTRIA FARMACÈUTICA\01. CARPETES LABORATORIS\102. Informes i Rappel"
OUTPUT_CSV = os.path.join(RAPPEL_DIR, "Seguiment_labs.csv")

VALID_MONTHS = {f"{m:02d}" for m in range(1, 13)}


class clsSeguiment:
    # Initializes empty DataFrames for accumulated lab data and fee conditions.
    # year: campaign year for the SEGUIMENT conditions file. Defaults to the
    # current year; a future frontend can pass it explicitly.
    def __init__(self, min_year: int = 2024, year: int = None):
        self.min_year = min_year
        self.year = year or datetime.date.today().year
        self.seguiment_file = os.path.join(RAPPEL_DIR, f"a.SEGUIMENT LABS {self.year}.xlsx")
        self.seguiment_labs = pd.DataFrame()
        self.condicio_seguiment = pd.DataFrame()

    # Main entry point. Traverses year/month folders, reads all lab files,
    # calculates fees, and saves the result to CSV.
    def run(self):
        years = self._get_years()
        for year_path in years:
            for mes_path, month_label in self._get_months(year_path):
                self._read_files(mes_path, month_label)
        self._calc_fee()
        self.seguiment_labs.to_csv(OUTPUT_CSV, index=False, sep=";")
        print(f"Saved to {OUTPUT_CSV}")

    # Returns sorted list of year folder paths where the folder name
    # is a number >= min_year (e.g. 2024, 2025).
    def _get_years(self):
        years = []
        for name in os.listdir(RAPPEL_DIR):
            full = os.path.join(RAPPEL_DIR, name)
            if os.path.isdir(full) and name.isdigit() and int(name) >= self.min_year:
                years.append(full)
        return sorted(years)

    # Returns list of (path, month_label) tuples for each valid month folder
    # inside a year directory. If a "mes" subfolder exists, uses that path.
    # month_label is the folder name like "01.2025".
    def _get_months(self, year_path):
        results = []
        for name in sorted(os.listdir(year_path)):
            if name[:2] not in VALID_MONTHS:
                continue
            month_path = os.path.join(year_path, name)
            if not os.path.isdir(month_path):
                continue
            mes_path = os.path.join(month_path, "mes")
            if os.path.isdir(mes_path):
                results.append((mes_path, name))
            else:
                results.append((month_path, name))
        return results

    # Reads all PARAFARMACIA xlsx files in a folder, adds a "date" column
    # with the month label, and accumulates into self.seguiment_labs.
    # Skips YTD and (NC) files.
    def _read_files(self, folder_path, month_label):
        for name in os.listdir(folder_path):
            if not name.endswith(".xlsx") or "PARAFARMACIA" not in name:
                continue
            if "YTD" in name or "(NC)" in name:
                continue
            print(name)
            file_path = os.path.join(folder_path, name)
            df = pd.read_excel(file_path, sheet_name="Acuerdo book")
            df["date"] = month_label
            self.seguiment_labs = pd.concat([self.seguiment_labs, df], ignore_index=True)

    # Reads the SEGUIMENT LABS Excel file, extracts the fee conditions
    # from the "Parafarmacia" sheet, and returns a cleaned DataFrame.
    def _get_seguiment_conditions(self):
        seg = pd.read_excel(self.seguiment_file, sheet_name="Parafarmacia")
        seg.columns = seg.iloc[0]
        seg = seg.drop(seg.index[0])
        if seg.empty:
            return pd.DataFrame()
        # The monthly-fee column is named "fee mes-YY" (year suffix varies),
        # so match it by prefix to stay valid across campaigns.
        fee_col = next((c for c in seg.columns
                        if str(c).strip().lower().startswith("fee mes")), None)
        keep = ['Etiquetas de fila', 'Neto', 'Fijo', 'Variable']
        rename = {"Etiquetas de fila": "Laboratorio"}
        if fee_col is not None:
            keep.append(fee_col)
            rename[fee_col] = "fee_mes"
        seg = seg[keep].rename(columns=rename)
        seg = seg.drop_duplicates().reset_index(drop=True)
        return seg

    # Merges fee conditions into the accumulated lab data based on
    # matching "Laboratorio Categorizado" with "Laboratorio".
    def _calc_fee(self):
        self.condicio_seguiment = self._get_seguiment_conditions()
        if self.condicio_seguiment.empty:
            print("No data in the condition file")
            return
        fee_cols = [c for c in ["Neto", "Fijo", "Variable", "fee_mes"]
                    if c in self.condicio_seguiment.columns]
        # One row per lab: with duplicate labs in SEGUIMENT the left-merge would
        # grow and the positional column assignment below would misalign rows.
        cond = self.condicio_seguiment.drop_duplicates(subset=["Laboratorio"], keep="first")
        merged = self.seguiment_labs.merge(
            cond[["Laboratorio"] + fee_cols],
            left_on="Laboratorio Categorizado",
            right_on="Laboratorio",
            how="left",
            suffixes=("", "_cond")
        )
        for col in fee_cols:
            self.seguiment_labs[col] = merged[col + "_cond"] if col + "_cond" in merged.columns else merged[col]
        if "Laboratorio_cond" in merged.columns:
            self.seguiment_labs.drop(columns=["Laboratorio_cond"], errors="ignore", inplace=True)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Acumula l'històric i calcula fees (Seguiment_labs.csv).")
    parser.add_argument("--min-year", type=int, default=2024, help="Any mínim de carpetes històriques")
    parser.add_argument("--year", type=int, default=None, help="Any de la campanya SEGUIMENT (opcional)")
    args = parser.parse_args()
    s = clsSeguiment(args.min_year, args.year)
    s.run()
