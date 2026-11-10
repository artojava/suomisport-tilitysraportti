"""Local Streamlit interface for settlement analysis."""

import csv
from decimal import Decimal
from io import StringIO

import pandas as pd
import streamlit as st

from tiliraportti_analyysi.analysis import analyze


def csv_value(value):
    if isinstance(value, Decimal):
        return str(value).replace(".", ",")
    if isinstance(value, str) and (
        value.startswith(("\t", "\r", "\n"))
        or value.lstrip().lstrip("\ufeff").lstrip().startswith(("=", "+", "-", "@"))
    ):
        return "'" + value
    return value


def csv_bytes(rows):
    stream = StringIO(newline="")
    writer = csv.DictWriter(stream, fieldnames=list(rows[0]), delimiter=";")
    writer.writeheader()
    for row in rows:
        writer.writerow({key: csv_value(value) for key, value in row.items()})
    return stream.getvalue().encode("utf-8-sig")


def euro(value):
    return f"{value:,.2f} €".replace(",", " ").replace(".", ",")


def frame(rows):
    return pd.DataFrame([{key: float(value) if isinstance(value, Decimal) else value
                          for key, value in row.items()} for row in rows])


def main():
    st.set_page_config(page_title="Tilitysanalyysi", page_icon="📊", layout="wide")
    st.title("Tilitysanalyysi")
    st.write("Tarkastele joukkueiden palvelumaksuja, komissioita ja nettotilityksiä.")
    st.caption("Tiedosto lähetetään sovelluksen palvelimelle ja käsitellään muistissa. "
               "Sovellus ei tallenna sitä levylle. Enimmäiskoko on 5 MiB. "
               "Lataa vain tietoja, joita sinulla on lupa käsitellä tässä palvelussa.")
    uploaded = st.file_uploader("Raahaa Excel tähän tai valitse tiedosto", type=["xlsx"])
    if uploaded is None:
        st.info("Valitse Suomisportin tilitysraportti, jossa on Tilitysrivit-välilehti.")
        return
    try:
        result = analyze(uploaded.getvalue())
    except ValueError as exc:
        st.error("Tiedostoa ei voitu käsitellä.")
        st.text(str(exc))
        return
    if result.warnings:
        st.warning(f"{len(result.warnings)} riviä jäi kohdistamatta. Ne sisältyvät Kohdistamatta-ryhmään.")
        with st.expander("Näytä kohdistusongelmat"):
            for warning in result.warnings:
                st.text(warning)
    selected = st.multiselect("Joukkueet", sorted({row["Joukkue"] for row in result.summary}),
                              placeholder="Kaikki joukkueet")
    rows = [row for row in result.summary if not selected or row["Joukkue"] in selected]
    details = [row for row in result.details if not selected or row["Joukkue"] in selected]
    columns = st.columns(4)
    for column, label in zip(columns, ["Maksut (€)", "Palvelumaksut (€)", "Komissiot (€)", "Nettotilitys (€)"]):
        column.metric(label.removesuffix(" (€)"), euro(sum(row[label] for row in rows)))
    has_settlement_date = "Tilityspäivä" in rows[0]
    st.subheader("Yhteenveto joukkueittain ja tilityspäivittäin" if has_settlement_date
                 else "Joukkuekohtainen yhteenveto")
    config = {key: st.column_config.NumberColumn(key, format="%.2f")
              for key in rows[0] if key.endswith(("(€)", "(%)"))}
    if has_settlement_date:
        config["Tilityspäivä"] = st.column_config.DateColumn("Tilityspäivä", format="DD.MM.YYYY")
    summary = frame(rows).style
    if has_settlement_date:
        palette = ["#dbeafe", "#dcfce7", "#fce7f3", "#ede9fe",
                   "#ffedd5", "#cffafe", "#fef9c3", "#e2e8f0"]
        days = sorted({row["Tilityspäivä"] for row in result.summary
                       if row["Tilityspäivä"] is not None})
        day_colors = {day: palette[index % len(palette)] for index, day in enumerate(days)}
        summary = summary.apply(
            lambda row: [f"background-color: {day_colors.get(row['Tilityspäivä'], '#f3f4f6')}; "
                         "color: #1f2937"] * len(row), axis=1,
        )
    else:
        summary = summary.set_properties(
            subset=["Kulut yhteensä (€)"],
            **{"background-color": "#fef3c7"},
        )
    summary = summary.set_properties(
        subset=["Kulut yhteensä (€)"], **{"color": "#b91c1c", "font-weight": "bold"},
    )
    st.dataframe(summary, hide_index=True, width="stretch", column_config=config)
    st.caption("Kulut esitetään positiivisina vähennyksinä. Prosentit lasketaan yhteenvetorivin "
               "maksujen summasta. Nollasumman prosentti jätetään tyhjäksi. "
               "Nettotilitys = maksut − kulut.")
    st.download_button("Lataa yhteenveto (CSV)", csv_bytes(rows), "joukkueyhteenveto.csv", "text/csv")
    st.subheader("Kulut joukkueittain")
    chart = frame(rows).groupby("Joukkue")[["Palvelumaksut (€)", "Komissiot (€)"]].sum()
    st.bar_chart(chart, stack=True, color=["#2563eb", "#f59e0b"])
    with st.expander("Tarkista lähderivit"):
        st.dataframe(frame(details), hide_index=True, width="stretch")
        st.download_button("Lataa kohdistetut rivit (CSV)", csv_bytes(details),
                           "kohdistetut-rivit.csv", "text/csv")
    with st.expander("Miten kulut kohdistetaan?"):
        st.write("Summat luetaan Maksuosuus-sarakkeesta. Rivin omaa joukkuetta käytetään ensin. "
                 "Jos se puuttuu, joukkue haetaan saman maksun Maksu-riveiltä. Maksu erotellaan "
                 "organisaation, tilityksen viitenumeron, maksun tunnisteen ja tilityspäivän perusteella. "
                 "Jos samalle maksulle löytyy useita joukkueita tai joukkuetta ei löydy, rivi jää "
                 "kohdistamatta. Yhteenveto-välilehteä ei summata uudelleen. "
                 "Maksujen lukumäärä tarkoittaa erillisiä maksun tunnisteita joukkueessa.")


if __name__ == "__main__":
    main()
