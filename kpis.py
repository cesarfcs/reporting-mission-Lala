# -*- coding: utf-8 -*-
import pandas as pd

PITCH_TAGS = {"Meeting", "Pitch"}
CONNECTED_TAGS = {"Meeting", "Pitch", "Sans Suite", "Standard"}
CALLED_TAGS = {"Meeting", "No answer", "Numéro Faux", "Pitch", "Sans Suite", "Standard"}

def _norm(s):
    return (s or "").strip()

def filter_period(df, start, end):
    df = df.copy()
    for col in ["Last Aircall call timestamp", "Date de la dernière activité"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    if start:
        df = df[(df["Last Aircall call timestamp"].isna() | (df["Last Aircall call timestamp"] >= start))
                | (df["Date de la dernière activité"].isna() | (df["Date de la dernière activité"] >= start))]
    if end:
        df = df[(df["Last Aircall call timestamp"].isna() | (df["Last Aircall call timestamp"] <= end))
                | (df["Date de la dernière activité"].isna() | (df["Date de la dernière activité"] <= end))]
    return df

def compute_metrics(df):
    df = df.copy()
    df["Last used Aircall tags"] = df.get("Last used Aircall tags", "").fillna("").map(_norm)
    df["Last Aircall call outcome"] = df.get("Last Aircall call outcome", "").fillna("").map(_norm)
    df["lemlist lead status"] = df.get("lemlist lead status", "").fillna("").map(_norm)
    df["Phase du cycle de vie"] = df.get("Phase du cycle de vie", "").fillna("").map(_norm)

    contacts_phone = int(df["Last Aircall call timestamp"].notna().sum())
    contacts_email = int(df["lemlist lead status"].replace("", pd.NA).notna().sum())
    calls_contacted = int(df["Last used Aircall tags"].isin(CALLED_TAGS).sum())
    calls_connected = int(df["Last used Aircall tags"].isin(CONNECTED_TAGS).sum())
    calls_pitched = int(df["Last used Aircall tags"].isin(PITCH_TAGS).sum())
    pitch_rate = (calls_pitched / calls_connected) if calls_connected else 0.0
    contacts_mailed = int((df["lemlist lead status"] != "").sum())
    contacts_opened = int((df["lemlist lead status"] == "Email opened").sum())
    contacts_replied = int((df["lemlist lead status"] == "Email replied").sum())
    rdv_phone = int((df["Last used Aircall tags"] == "Meeting").sum())
    rdv_email = int(((df["Phase du cycle de vie"] == "RDV - Bon contact") & (df["lemlist lead status"] == "Email replied")).sum())
    calls_conv_rate = (rdv_phone / calls_contacted) if calls_contacted else 0.0
    email_conv_rate = (rdv_email / contacts_mailed) if contacts_mailed else 0.0

    return {
        "contacts_phone": contacts_phone,
        "contacts_email": contacts_email,
        "calls_contacted": calls_contacted,
        "calls_connected": calls_connected,
        "calls_pitched": calls_pitched,
        "pitch_rate": pitch_rate,
        "contacts_mailed": contacts_mailed,
        "contacts_opened": contacts_opened,
        "contacts_replied": contacts_replied,
        "rdv_phone": rdv_phone,
        "rdv_email": rdv_email,
        "calls_conv_rate": calls_conv_rate,
        "email_conv_rate": email_conv_rate,
    }
