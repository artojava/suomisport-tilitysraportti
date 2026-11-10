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
    selected = st.multiselect("Joukkueet", [row["Joukkue"] for row in result.summary],
                              placeholder="Kaikki joukkueet")
    rows = [row for row in result.summary if not selected or row["Joukkue"] in selected]
    details = [row for row in result.details if not selected or row["Joukkue"] in selected]
    columns = st.columns(4)
    for column, label in zip(columns, ["Maksut (€)", "Palvelumaksut (€)", "Komissiot (€)", "Nettotilitys (€)"]):
        column.metric(label.removesuffix(" (€)"), euro(sum(row[label] for row in rows)))
    st.subheader("Joukkuekohtainen yhteenveto")
    config = {key: st.column_config.NumberColumn(key, format="%.2f")
              for key in rows[0] if key.endswith(("(€)", "(%)"))}
    summary = frame(rows).style.set_properties(
        subset=["Kulut yhteensä (€)"],
        **{"background-color": "#fef3c7", "color": "#78350f", "font-weight": "bold"},
    )
    st.dataframe(summary, hide_index=True, width="stretch", column_config=config)
    st.caption("Kulut esitetään positiivisina vähennyksinä. Prosentit lasketaan joukkueen maksujen "
               "summasta. Nollasumman prosentti jätetään tyhjäksi. Nettotilitys = maksut − kulut.")
    st.download_button("Lataa yhteenveto (CSV)", csv_bytes(rows), "joukkueyhteenveto.csv", "text/csv")
    st.subheader("Kulut joukkueittain")
    chart = frame(rows).set_index("Joukkue")[["Palvelumaksut (€)", "Komissiot (€)"]]
    st.bar_chart(chart, stack=True, color=["#2563eb", "#f59e0b"])
    with st.expander("Tarkista lähderivit"):
        st.dataframe(frame(details), hide_index=True, width="stretch")
        st.download_button("Lataa kohdistetut rivit (CSV)", csv_bytes(details),
                           "kohdistetut-rivit.csv", "text/csv")
    with st.expander("Miten kulut kohdistetaan?"):
        st.write("Summat luetaan Maksuosuus-sarakkeesta. Rivin omaa joukkuetta käytetään ensin. "
                 "Jos se puuttuu, joukkue haetaan saman maksun Maksu-riveiltä. Maksu erotellaan "
                 "organisaation, tilityksen viitenumeron ja maksun tunnisteen perusteella. "
                 "Jos samalle maksulle löytyy useita joukkueita tai joukkuetta ei löydy, rivi jää "
                 "kohdistamatta. Yhteenveto-välilehteä ei summata uudelleen. "
                 "Maksujen lukumäärä tarkoittaa erillisiä maksun tunnisteita joukkueessa.")


if __name__ == "__main__":
    main()
