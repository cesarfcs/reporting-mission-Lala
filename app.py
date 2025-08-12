# -*- coding: utf-8 -*-
import datetime as dt
import streamlit as st
import pandas as pd
from offers import OFFERS
from kpis import compute_metrics, filter_period
from ppt_builder import build_ppt

st.set_page_config(page_title="Lalaleads – Reporting", layout="wide")
st.title("Lalaleads – Génération de rapport")
st.caption("Téléversez l'export HubSpot, choisissez l'offre et les canaux, puis générez le PowerPoint.")

# 1) Config rapport
colA, colB, colC = st.columns([1,1,1])
with colA:
    client_name = st.text_input("Nom du client", value="Client Exemple")
    report_type = st.selectbox("Type de rapport", ["Hebdomadaire", "Mensuel"])
    report_cycle = st.text_input("Semaine/Cycle", value="Semaine 1")
with colB:
    offer = st.selectbox("Offre souscrite", list(OFFERS.keys()))
    custom_contacts = st.number_input("Contacts par cycle (si Offre personnalisée)", min_value=100, step=50, value=1600,
                                      disabled=(offer != "Offre personnalisée"))
    channels_default = OFFERS[offer]["channels"][:]
    if OFFERS[offer]["linkedin_optional"]:
        channels_default = channels_default + ["LinkedIn"]
with colC:
    channels = st.multiselect("Canaux de prospection", ["Téléphone", "E-mail", "LinkedIn"], default=channels_default)
    today = dt.date.today()
    default_start = today.replace(day=1)
    date_range = st.date_input("Période (début / fin)", value=(default_start, today))
    logo_file = st.file_uploader("Logo client (PNG/JPG)", type=["png", "jpg", "jpeg"])

st.markdown("---")

# 2) CSV
csv = st.file_uploader("Export HubSpot (CSV)", type=["csv"])
generate = st.button("Générer le rapport")

if generate:
    if csv is None:
        st.error("Merci d’importer un fichier CSV HubSpot.")
        st.stop()

    df = pd.read_csv(csv)
    start = pd.to_datetime(date_range[0])
    end = pd.to_datetime(date_range[1])
    dfp = filter_period(df, start, end)
    m = compute_metrics(dfp)

    # Aperçu KPI
    st.subheader("Aperçu des KPI (période sélectionnée)")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Contacts téléphone", m["contacts_phone"])
    k2.metric("Contacts e-mail", m["contacts_email"])
    k3.metric("RDV téléphone", m["rdv_phone"], f"{m['calls_conv_rate']*100:.1f}% conv.")
    k4.metric("RDV e-mail", m["rdv_email"], f"{m['email_conv_rate']*100:.1f}% conv.")

    # Génération PPT
    offer_target = OFFERS[offer]["contacts_target"] if offer != "Offre personnalisée" else int(custom_contacts)
    date_label = f"{date_range[0].strftime('%d/%m/%Y')} – {date_range[1].strftime('%d/%m/%Y')}"
    logo_path = None
    if logo_file is not None:
        logo_path = "uploaded_logo.png"
        with open(logo_path, "wb") as f:
            f.write(logo_file.read())

    out = "rapport_lalaleads.pptx"
    build_ppt(
        out_path=out,
        client_name=client_name,
        logo_path=logo_path,
        report_type=report_type,
        report_cycle=report_cycle,
        offer_name=offer,
        offer_contacts_target=offer_target,
        channels_list=channels,
        metrics=m,
        date_range_label=date_label,
    )

    with open(out, "rb") as f:
        st.success("Rapport généré avec succès ✅")
        st.download_button("Télécharger le PowerPoint", f.read(), file_name=out, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
