import pandas as pd
file_path = "Mappe1.xlsx"

# Load Excel file into a DataFrame
df = pd.read_excel(file_path)
df_cleaned = df.dropna(subset=["Technischer Name"])
result = df_cleaned.melt(id_vars=["Beschreibung", "Technischer Name", "Metakette", "FB/Verantw.", "DSCR fuer Text"], var_name="Datum", value_name="Geplant")
result = result.dropna(subset=["Geplant"])
result['Geplant'] = "X"
column = df[["Beschreibung", "Technischer Name"]]


def find_App_Period(id):
    period = ""
    applik = ""
    for i, r in column.iterrows():
        if pd.isna(r["Technischer Name"]):
            if period == "":
                period = r["Beschreibung"]
            elif applik == "":
                applik = r["Beschreibung"][3:]
            else:
                if r["Beschreibung"][0].isdigit():
                    applik = r["Beschreibung"][3:]
                else:
                    period = r["Beschreibung"]
                    applik = ""
        else:
            if r["Beschreibung"] == id:
                return period, applik


for index, row in result.iterrows():
    period, applik = find_App_Period(row["Beschreibung"])
    if applik == "ertrieb":
        applik = "Vertrieb"
    result.loc[index, "Period"] = period
    result.loc[index, "Applikation"] = applik

order = ["Period", "Applikation", "Technischer Name", "Beschreibung", "Metakette", "FB/Verantw.", "DSCR fuer Text", "Datum", "Geplant"]
result = result.reindex(columns=order)    

result.to_excel("test.xlsx", index=False) # Save DataFrame to Excel file