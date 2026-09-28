from __future__ import annotations

import base64
import datetime as dt
import hashlib
import hmac
import io
import json
import math
import os
import sqlite3
import zipfile
from pathlib import Path
from zoneinfo import ZoneInfo

import joblib
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.graph_objects as go
import requests
import streamlit as st
from PIL import Image

try:
    import folium
    from folium.plugins import Draw
    from streamlit_folium import st_folium
except Exception:
    folium = None
    Draw = None
    st_folium = None

try:
    import rasterio
    from rasterio.io import MemoryFile
except Exception:
    rasterio = None
    MemoryFile = None

# ================================================================
# QUOTORA PRO — final production-oriented Streamlit application
# ================================================================
def _load_page_icon():
    """Logo Quotora comme favicon (onglet du navigateur). Repli sur l'emoji si le fichier est absent."""
    base = Path(__file__).resolve().parent
    for name in ("quotora_icon.png", "quotora_logo.png"):  # icône carrée (symbole Q) en priorité
        try:
            return Image.open(base / name)
        except Exception:
            continue
    return "🏗️"

st.set_page_config(
    page_title="Quotora Pro — Construction Intelligence",
    page_icon=_load_page_icon(),
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE = Path(__file__).resolve().parent
MODEL_PATH = BASE / "quotora_model_final.joblib"
LOGO_PATH = BASE / "quotora_logo.png"
BUILDING_PATH = BASE / "quotora_building.png"
DB_PATH = BASE / "quotora_suivi.db"

APP_NAME = "Quotora Pro"
TAGLINE = "Estimation & Pilotage Construction"
REPORT_TITLE = "Construction Estimation & Technical Report"

# -------------------------------
# Catalogues / technical references
# -------------------------------
REGIONS = ["Adamaoua", "Centre", "Est", "Extrême-Nord", "Littoral", "Nord", "Nord-Ouest", "Ouest", "Sud", "Sud-Ouest"]
STANDINGS = ["Économique", "Standard", "Moyen standing", "Haut standing", "Luxe"]
FOUNDATIONS = ["Semelles isolées", "Semelles filantes", "Semelles isolées + longrines", "Radier général", "Radier nervuré", "Fondation sur pieux", "Fondation profonde", "Autre"]
SOILS = ["Latéritique", "Argileux", "Sableux", "Marécageux", "Rocheux", "Remblai", "Sol mixte", "Autre"]
ROOFS = ["Artisanale simple", "Artisanale améliorée", "Charpente métallique + tôle", "Charpente métallique + bac acier", "Couverture en tuile", "Tuile terre cuite", "Polycarbonate", "Polycarbonate alvéolaire", "Toiture-terrasse non accessible", "Toiture-terrasse accessible", "Toiture-terrasse haut standing"]
RDC_TYPES = ["Dallage sur terre-plein", "Dallage terre-plein renforcé", "Dallage industriel", "Dallage industriel lourd", "Radier"]
SLABS = ["Dalle pleine en béton armé", "Dalle à hourdis / corps creux"]
OPENINGS = ["Fenêtre bois", "Fenêtre bois améliorée", "Fenêtre aluminium vitrée", "Fenêtre aluminium renforcée", "Fenêtre aluminium double vitrage", "Baie vitrée aluminium", "Baie vitrée aluminium double vitrage", "Porte bois", "Porte aluminium vitrée"]
CEILINGS = ["Sans plafond", "Staff standard", "Staff décoratif", "Staff haut standing", "Lambris PVC", "Lambris bois", "Lambris bois haut standing", "Contreplaqué", "Contreplaqué décoratif", "BA13 / plaque de plâtre", "PVC décoratif", "Plafond métallique", "Plafond acoustique", "Plafond aluminium", "Plafond bois architectural", "Béton apparent", "Autre"]

# Référentiel utilisateur pour le lot 3 — Élévations.
# Pour les fourchettes, le prix de référence par défaut est le milieu de la fourchette et reste modifiable dans le formulaire.
WALL_CATALOG = {
    "Parpaing 10": {"thickness": "10 cm", "min": 7500, "max": 7500, "default": 7500},
    "Parpaing 15": {"thickness": "15 cm", "min": 10000, "max": 10000, "default": 10000},
    "Parpaing 20": {"thickness": "20 cm", "min": 12000, "max": 12000, "default": 12000},
    "Parpaing plein": {"thickness": "20 cm", "min": 14000, "max": 14000, "default": 14000},
    "Parpaing bourré": {"thickness": "20 cm", "min": 16000, "max": 20000, "default": 18000},
    "Brique terre/adobe": {"thickness": "—", "min": 10000, "max": 14000, "default": 12000},
    "Brique cuite": {"thickness": "10–15 cm", "min": 15000, "max": 20000, "default": 17500},
    "Brique cuite pleine": {"thickness": "20–30 cm", "min": 18000, "max": 25000, "default": 21500},
    "BTC": {"thickness": "15 cm", "min": 14000, "max": 18000, "default": 16000},
    "BTC stabilisé": {"thickness": "15 cm", "min": 16000, "max": 21000, "default": 18500},
    "Pisé": {"thickness": "—", "min": 15000, "max": 22000, "default": 18500},
    "Pierre / moellons": {"thickness": "—", "min": 18000, "max": 25000, "default": 21500},
    "Bloc latéritique taillé": {"thickness": "15–20 cm", "min": 14000, "max": 18000, "default": 16000},
    "Carabot": {"thickness": "—", "min": 12000, "max": 16000, "default": 14000},
    "Béton armé coulé": {"thickness": "15 cm", "min": 45000, "max": 55000, "default": 50000},
    "Béton armé coulé 20 cm": {"thickness": "20 cm", "min": 55000, "max": 70000, "default": 62500},
    "Mur mixte parpaing + BA": {"thickness": "—", "min": 20000, "max": 28000, "default": 24000},
    "Mur mixte BTC + BA": {"thickness": "—", "min": 21000, "max": 30000, "default": 25500},
    "Mur bois": {"thickness": "—", "min": 25000, "max": 35000, "default": 30000},
    "Aluminium vitré": {"thickness": "—", "min": 65000, "max": 100000, "default": 82500},
    "Mur-rideau": {"thickness": "—", "min": 120000, "max": 160000, "default": 140000},
}
WALL_PRICE = {k: v["default"] for k, v in WALL_CATALOG.items()}
# Compatibilité avec l'ancien catalogue de l'application.
WALLS = list(WALL_CATALOG.keys())
WALL_THICKNESS = {k: v["thickness"] for k, v in WALL_CATALOG.items()}
OPEN_PRICE = {"Fenêtre bois": 55000, "Fenêtre bois améliorée": 75000, "Fenêtre aluminium vitrée": 95000, "Fenêtre aluminium renforcée": 120000, "Fenêtre aluminium double vitrage": 140000, "Baie vitrée aluminium": 160000, "Baie vitrée aluminium double vitrage": 220000, "Porte bois": 85000, "Porte aluminium vitrée": 120000}
CEIL_PRICE = {"Sans plafond": 0, "Staff standard": 25000, "Staff décoratif": 25000, "Staff haut standing": 25000, "Lambris PVC": 7500, "Lambris bois": 12000, "Lambris bois haut standing": 20000, "Contreplaqué": 7000, "Contreplaqué décoratif": 10000, "BA13 / plaque de plâtre": 12000, "PVC décoratif": 9000, "Plafond métallique": 18000, "Plafond acoustique": 20000, "Plafond aluminium": 22000, "Plafond bois architectural": 25000, "Béton apparent": 6000, "Autre": 0}
ROOF_PRICE = {"Artisanale simple": 18000, "Artisanale améliorée": 24000, "Charpente métallique + tôle": 28000, "Charpente métallique + bac acier": 32000, "Couverture en tuile": 35000, "Tuile terre cuite": 40000, "Polycarbonate": 38000, "Polycarbonate alvéolaire": 45000, "Toiture-terrasse non accessible": 30000, "Toiture-terrasse accessible": 45000, "Toiture-terrasse haut standing": 65000}
# Options de toiture détaillées : la sélection est combinée en interne dans type_toiture pour le modèle ML.
ROOF_STRUCTURES = [
    "Charpente bois",
    "Charpente métallique",
    "Charpente artisanale",
    "Toiture-terrasse",
]
ROOF_SHEETS = [
    "Tôle ondulée simple",
    "Tôle bac acier 4/10",
    "Tôle bac acier 5/10",
    "Tôle bac acier 6/10",
    "Tôle bac acier 7/10",
    "Tôle bac acier 8/10",
]
RDC_PRICES = {"Dallage sur terre-plein": 6000, "Dallage terre-plein renforcé": 14000, "Dallage industriel": 45000, "Dallage industriel lourd": 0, "Radier": 0}
# Terrassement: optional. For plain-pied (R+0), use a lighter editable budgetary reference instead of the ML amount.
TERRASSEMENT_PLAIN_PIED_DEFAULT = 3500  # FCFA/m², provisional and editable in the form
FOUNDATION_CONCRETE_PRICE_MIN = 190000
FOUNDATION_CONCRETE_PRICE_MAX = 200000
FOUNDATION_CONCRETE_PRICE_DEFAULT = 195000
FOUNDATION_WALL_TYPES = ["Parpaing 15 cm", "Parpaing 20 cm", "Voile béton armé"]
FOUNDATION_WALL_PRICE = {"Parpaing 15 cm": 15000, "Parpaing 20 cm": 18000}
FOUNDATION_VOILE_PRICE_MIN = 210000
FOUNDATION_VOILE_PRICE_MAX = 210000
FOUNDATION_VOILE_PRICE_DEFAULT = 210000
SLAB_PRICES = {"Dalle pleine en béton armé": 45000, "Dalle à hourdis / corps creux": 32000}
RDC_FIN = {"Économique": .05, "Standard": .08, "Moyen standing": .10, "Haut standing": .12, "Luxe": .15}
SLAB_FIN = {"Économique": .10, "Standard": .15, "Moyen standing": .17, "Haut standing": .20, "Luxe": .25}
REGION_MULT = {"Littoral": 1.00, "Centre": 1.05, "Ouest": 1.10, "Nord-Ouest": 1.15, "Sud-Ouest": 1.10, "Sud": 1.10, "Est": 1.15, "Adamaoua": 1.15, "Nord": 1.15, "Extrême-Nord": 1.20}
PRIMARY_WALL = {"Adamaoua": "Brique terre/adobe", "Centre": "Parpaing 15", "Est": "Brique terre/adobe", "Extrême-Nord": "Brique terre/adobe", "Littoral": "Parpaing 15", "Nord": "Brique terre/adobe", "Nord-Ouest": "Brique terre/adobe", "Ouest": "Brique terre/adobe", "Sud": "Parpaing 15", "Sud-Ouest": "Parpaing 15"}

CURRENCIES = {"XAF — Franc CFA BEAC": "XAF", "EUR — Euro": "EUR", "USD — US Dollar": "USD", "CHF — Swiss Franc": "CHF", "GBP — Pound Sterling": "GBP", "CAD — Canadian Dollar": "CAD", "AUD — Australian Dollar": "AUD", "JPY — Japanese Yen": "JPY", "CNY — Chinese Yuan": "CNY", "AED — UAE Dirham": "AED", "MAD — Moroccan Dirham": "MAD", "ZAR — South African Rand": "ZAR", "NGN — Nigerian Naira": "NGN", "INR — Indian Rupee": "INR", "SGD — Singapore Dollar": "SGD"}
FALLBACK_PER_XAF = {"XAF": 1.0, "EUR": 1/655.957, "USD": 1/600.0, "CHF": 1/700.0, "GBP": 1/760.0, "CAD": 1/435.0, "AUD": 1/390.0, "JPY": 1/4.0, "CNY": 1/82.0, "AED": 1/163.0, "MAD": 1/65.0, "ZAR": 1/33.0, "NGN": 1/1.7, "INR": 1/7.1, "SGD": 1/460.0}

LOTS = ["1 — Installation de chantier", "2 — Terrassement", "3 — Fondations", "4 — Élévation", "5 — Charpente & couverture", "6 — Enduits, chape & étanchéité", "7 — Menuiseries", "8 — Plomberie", "9 — Électricité", "10 — Carrelage, faux plafond & peinture", "11 — VRD & extérieurs", "12 — Planchers, dalles & dallages"]

# -------------------------------
# Translation dictionaries
# -------------------------------
T = {
    "fr": {
        "home": "Accueil", "new": "Nouveau projet", "projects": "Mes projets", "terrain": "Terrain & Carte", "estimate": "Estimation", "tracking": "Suivi de chantier", "reports": "Rapports", "settings": "Paramètres", "help": "Aide & Support",
        "quick": "Estimation rapide", "quick_sub": "Obtenez une estimation en quelques clics.", "professional": "Estimation technique professionnelle", "professional_sub": "Configuration complète avec les détails techniques.", "start": "Commencer", "calculate": "Obtenir l’estimation", "save": "Enregistrer", "reset": "Réinitialiser",
        "language": "Langue", "currency": "Devise", "region": "Région", "building": "Type de bâtiment", "surface": "Surface du RDC (m²)", "levels": "Niveaux au-dessus du RDC", "standing": "Standing", "soil": "Type de sol", "foundation": "Fondation", "roof": "Toiture", "wall": "Mur principal", "opening": "Ouvertures", "ceiling": "Plafond", "floor": "Dallage / plancher", "season": "Saison", "access": "Accès chantier", "distance": "Distance transport (km)", "earthwork": "Terrassement", "earthwork_yes": "Oui", "earthwork_no": "Non", "earthwork_pp_price": "Prix terrassement plain-pied (FCFA/m²)",
        "project_info": "Informations du projet", "building_config": "Bâtiment", "walls": "Murs & façades", "terrain_map": "Terrain & cartographie", "advanced": "Options avancées", "results": "Résultats", "costs": "Estimation des coûts", "gross": "Gros œuvre", "finishes": "Finitions", "total": "Total projet", "cost_m2": "Coût moyen / m²", "print": "Imprimer", "report": "Générer le rapport", "founder": "Fondé par Tsala Christian", "quick_hint": "Pour un client qui veut simplement connaître le budget.", "pro_hint": "Pour une étude budgétaire détaillée et un pré-diagnostic technique.",
        "map": "Carte satellite", "terrain3d": "Terrain 3D", "coordinates": "Coordonnées GPS", "draw": "Dessiner sur la carte", "import": "Importer KML / GPX / GeoJSON / CSV", "relief": "Dénivelée", "retaining": "Mur de soutènement", "recommended": "Étude de mur recommandée", "no_wall": "Aucun côté ne dépasse le seuil de 1,50 m.", "length": "Longueur estimée", "photos": "Photos & rapports", "progress": "Avancement", "status": "Statut", "observation": "Observation", "update": "Enregistrer la mise à jour", "history": "Historique", "currency_note": "Les taux sont indicatifs et peuvent être actualisés en ligne.",
    },
    "en": {
        "home": "Home", "new": "New project", "projects": "My projects", "terrain": "Site & Map", "estimate": "Estimate", "tracking": "Site progress", "reports": "Reports", "settings": "Settings", "help": "Help & Support",
        "quick": "Quick estimate", "quick_sub": "Get a budget estimate in a few clicks.", "professional": "Professional technical estimate", "professional_sub": "Full configuration with technical details.", "start": "Start", "calculate": "Get estimate", "save": "Save", "reset": "Reset",
        "language": "Language", "currency": "Currency", "region": "Region", "building": "Building type", "surface": "Ground-floor area (m²)", "levels": "Levels above ground floor", "standing": "Quality level", "soil": "Soil type", "foundation": "Foundation", "roof": "Roof", "wall": "Main wall", "opening": "Openings", "ceiling": "Ceiling", "floor": "Ground floor / slab", "season": "Season", "access": "Site access", "distance": "Transport distance (km)", "earthwork": "Earthworks", "earthwork_yes": "Yes", "earthwork_no": "No", "earthwork_pp_price": "Plain-pied earthworks price (FCFA/m²)",
        "project_info": "Project information", "building_config": "Building", "walls": "Walls & façades", "terrain_map": "Site & mapping", "advanced": "Advanced options", "results": "Results", "costs": "Cost estimate", "gross": "Structural works", "finishes": "Finishes", "total": "Total project", "cost_m2": "Average cost / m²", "print": "Print", "report": "Generate report", "founder": "Founded by Tsala Christian", "quick_hint": "For clients who simply need a budget range.", "pro_hint": "For detailed budget configuration and technical pre-diagnosis.",
        "map": "Satellite map", "terrain3d": "3D terrain", "coordinates": "GPS coordinates", "draw": "Draw on map", "import": "Import KML / GPX / GeoJSON / CSV", "relief": "Relief", "retaining": "Retaining wall", "recommended": "Retaining wall study recommended", "no_wall": "No side exceeds the 1.50 m threshold.", "length": "Estimated length", "photos": "Photos & reports", "progress": "Progress", "status": "Status", "observation": "Observation", "update": "Save progress update", "history": "History", "currency_note": "Rates are indicative and may be refreshed online.",
    },
}

TRANSLATE = {
    "Économique": "Economy", "Standard": "Standard", "Moyen standing": "Mid-range", "Haut standing": "High-end", "Luxe": "Luxury",
    "Semelles isolées": "Isolated footings", "Semelles filantes": "Strip footings", "Semelles isolées + longrines": "Isolated footings + tie beams", "Radier général": "Raft foundation", "Radier nervuré": "Ribbed raft", "Fondation sur pieux": "Pile foundation", "Fondation profonde": "Deep foundation", "Autre": "Other",
    "Latéritique": "Lateritic", "Argileux": "Clayey", "Sableux": "Sandy", "Marécageux": "Swampy", "Rocheux": "Rocky", "Remblai": "Fill soil", "Sol mixte": "Mixed soil",
    "Artisanale simple": "Basic artisan roof", "Artisanale améliorée": "Improved artisan roof", "Charpente métallique + tôle": "Steel frame + sheets", "Charpente métallique + bac acier": "Steel frame + steel deck", "Couverture en tuile": "Tile roofing", "Tuile terre cuite": "Terracotta tiles", "Polycarbonate": "Polycarbonate", "Polycarbonate alvéolaire": "Multiwall polycarbonate", "Toiture-terrasse non accessible": "Non-accessible roof terrace", "Toiture-terrasse accessible": "Accessible roof terrace", "Toiture-terrasse haut standing": "High-end roof terrace",
    "Dallage sur terre-plein": "Ground-bearing slab", "Dallage terre-plein renforcé": "Reinforced ground-bearing slab", "Dallage industriel": "Industrial slab", "Dallage industriel lourd": "Heavy-duty industrial slab", "Radier": "Raft slab", "Dalle pleine en béton armé": "Reinforced concrete solid slab", "Dalle à hourdis / corps creux": "Hollow-block slab",
    "Parpaing 10": "10 cm concrete block", "Parpaing 15": "15 cm concrete block", "Parpaing 20": "20 cm concrete block", "Parpaing plein": "Solid concrete block", "Brique terre/adobe": "Earth/adobe brick", "Brique cuite": "Fired brick", "BTC": "Compressed earth block", "BTC stabilisé": "Stabilized earth block", "Pisé": "Rammed earth", "Pierre/moellons": "Stone/rubble masonry", "Béton armé coulé": "Cast reinforced concrete", "Aluminium vitré": "Glazed aluminium", "Mur-rideau": "Curtain wall", "Bois": "Timber wall", "Mur mixte parpaing + béton armé": "Concrete block + RC hybrid wall", "Mur mixte BTC + béton armé": "Earth block + RC hybrid wall", "Carabot": "Carabot block",
    "Fenêtre bois": "Wood window", "Fenêtre bois améliorée": "Improved wood window", "Fenêtre aluminium vitrée": "Glazed aluminium window", "Fenêtre aluminium renforcée": "Reinforced aluminium window", "Fenêtre aluminium double vitrage": "Double-glazed aluminium window", "Baie vitrée aluminium": "Aluminium glazed door", "Baie vitrée aluminium double vitrage": "Double-glazed aluminium glazed door", "Porte bois": "Wood door", "Porte aluminium vitrée": "Glazed aluminium door",
    "Sans plafond": "No ceiling", "Staff standard": "Standard plaster ceiling", "Staff décoratif": "Decorative plaster ceiling", "Staff haut standing": "High-end plaster ceiling", "Lambris PVC": "PVC cladding", "Lambris bois": "Wood cladding", "Lambris bois haut standing": "High-end wood cladding", "Contreplaqué": "Plywood", "Contreplaqué décoratif": "Decorative plywood", "BA13 / plaque de plâtre": "Drywall / plasterboard", "PVC décoratif": "Decorative PVC", "Plafond métallique": "Metal ceiling", "Plafond acoustique": "Acoustic ceiling", "Plafond aluminium": "Aluminium ceiling", "Plafond bois architectural": "Architectural wood ceiling", "Béton apparent": "Exposed concrete ceiling",
    "Saison sèche": "Dry season", "Saison des pluies": "Rainy season", "Accès normal": "Normal access", "Accès difficile": "Difficult access",
}

BUILDING_TYPES = {"Maison individuelle": "Single-family house", "Immeuble résidentiel": "Residential building", "Immeuble commercial": "Commercial building", "Bureaux": "Office building", "Commerce": "Retail building", "Équipement public": "Public facility"}

# -------------------------------
# UI / CSS
# -------------------------------
def tr(key: str, lang: str) -> str:
    return T.get(lang, T["fr"]).get(key, key)

def label(value: str, lang: str) -> str:
    if lang == "fr":
        return value
    return TRANSLATE.get(value, value)

def fcfa(x: float) -> str:
    return f"{x:,.0f}".replace(",", " ") + " FCFA"

def money(value: float, code: str) -> str:
    symbols = {"XAF": "FCFA", "EUR": "€", "USD": "$", "CHF": "CHF", "GBP": "£", "CAD": "C$", "AUD": "A$", "JPY": "¥", "CNY": "¥", "AED": "AED", "MAD": "MAD", "ZAR": "R", "NGN": "₦", "INR": "₹", "SGD": "S$"}
    if code == "XAF":
        return fcfa(value)
    return f"{value:,.2f} {symbols.get(code, code)}"

def inject_css():
    st.markdown("""
    <style>
    :root { --q-blue:#2563eb; --q-navy:#071b2a; --q-soft:#f5f8fc; --q-border:#e3e9f2; --q-text:#142235; }
    .stApp { background:#f7f9fc; color:var(--q-text); }
    [data-testid="stSidebar"] { background:linear-gradient(180deg,#071b2a 0%,#0b2940 100%); }
    [data-testid="stSidebar"] * { color:#e9f2fb !important; }
    /* Sidebar: keep the dark background, but make navigation button text readable */
    [data-testid="stSidebar"] div.stButton > button,
    [data-testid="stSidebar"] div.stButton > button p,
    [data-testid="stSidebar"] div.stButton > button span {
        color:#142235 !important;
        background:#ffffff !important;
        opacity:1 !important;
        font-weight:700 !important;
    }
    [data-testid="stSidebar"] div.stButton > button:hover,
    [data-testid="stSidebar"] div.stButton > button:hover p,
    [data-testid="stSidebar"] div.stButton > button:hover span {
        color:#1d4ed8 !important;
    }
    /* Language selector remains readable on the dark sidebar */
    [data-testid="stSidebar"] div[data-testid="stRadio"] label,
    [data-testid="stSidebar"] div[data-testid="stRadio"] label p,
    [data-testid="stSidebar"] div[data-testid="stRadio"] label span {
        color:#e9f2fb !important;
        opacity:1 !important;
    }
    .q-hero { min-height:190px; border-radius:22px; padding:28px 34px; background:linear-gradient(90deg,rgba(4,20,33,.94) 0%,rgba(4,20,33,.70) 45%,rgba(4,20,33,.15) 100%), url('data:image/png;base64,BUILDING') center/cover; color:white; margin-bottom:22px; }
    .q-hero h1 { font-size:38px; margin:0; letter-spacing:-1px; }
    .q-hero p { font-size:16px; margin:8px 0 0; opacity:.9; }
    .q-card { background:white; border:1px solid var(--q-border); border-radius:18px; padding:22px; box-shadow:0 5px 20px rgba(18,40,70,.05); }
    .q-card h3 { margin-top:0; }
    .q-mode { min-height:170px; }
    .q-small { color:#667085; font-size:13px; }
    .q-footer { text-align:center; color:#7a8798; font-size:12px; padding:30px 0 8px; }
    .q-badge { display:inline-block; padding:5px 10px; border-radius:99px; background:#eaf2ff; color:#1d4ed8; font-size:12px; font-weight:700; }
    .q-result { border-radius:16px; padding:18px; background:linear-gradient(135deg,#0f4fd8,#2563eb); color:white; }
    .q-result-green { background:linear-gradient(135deg,#058b68,#10b981); }
    .q-result-purple { background:linear-gradient(135deg,#6d28d9,#8b5cf6); }
    .q-divider { height:1px; background:#e6ebf2; margin:14px 0; }
    button[kind=secondary], div.stButton > button { color:#142235 !important; background:#ffffff !important; border:1px solid #d5deea !important; }
    div.stButton > button:hover { border-color:#2563eb !important; color:#1d4ed8 !important; }
    div[data-testid=stExpander] { border:1px solid #dbe3ec !important; border-radius:12px !important; margin-bottom:8px !important; }
    div[data-testid=stExpander] details summary { color:#142235 !important; font-weight:700 !important; }
    @media (max-width:900px) { .q-hero h1{font-size:28px}.q-hero{padding:22px}.q-card{padding:16px} }
    /* Field labels: always visible above every input/select */
    div[data-testid="stNumberInput"] label,
    div[data-testid="stSelectbox"] label,
    div[data-testid="stTextInput"] label,
    div[data-testid="stRadio"] > label,
    div[data-testid="stSlider"] > label,
    div[data-testid="stDateInput"] label,
    div[data-testid="stFileUploader"] label {
        color:#142235 !important;
        font-weight:700 !important;
        font-size:14px !important;
        margin-bottom:5px !important;
        opacity:1 !important;
    }
    div[data-testid="stNumberInput"] label p,
    div[data-testid="stSelectbox"] label p,
    div[data-testid="stTextInput"] label p,
    /* Radio options: explicitly display Oui / Non (and all choices) in dark text */
    div[data-testid="stRadio"] [role="radiogroup"] label,
    div[data-testid="stRadio"] [role="radiogroup"] label p,
    div[data-testid="stRadio"] [role="radiogroup"] label span {
        color:#142235 !important;
        opacity:1 !important;
        visibility:visible !important;
        font-weight:600 !important;
    }
    div[data-testid="stRadio"] > label p,
    div[data-testid="stSlider"] > label p,
    div[data-testid="stDateInput"] label p,
    div[data-testid="stFileUploader"] label p {
        color:#142235 !important;
        opacity:1 !important;
    }
    /* Small explanatory text directly under/around fields */
    .q-field-help {
        color:#667085 !important;
        font-size:12px !important;
        margin-top:-4px !important;
        margin-bottom:10px !important;
    }
    /* Full-width GIS map component */
    [data-testid="stCustomComponentV1"] { width: 100% !important; }
    </style>
    """.replace("BUILDING", building_data_uri()), unsafe_allow_html=True)

def building_data_uri():
    if not BUILDING_PATH.exists():
        return ""
    return base64.b64encode(BUILDING_PATH.read_bytes()).decode("ascii")

def logo_uri():
    if not LOGO_PATH.exists():
        return ""
    return "data:image/png;base64," + base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")

# -------------------------------
# Model
# -------------------------------
@st.cache_resource(show_spinner=False)
def load_model():
    art = joblib.load(MODEL_PATH)
    return art

ART = load_model()
PIPE = ART["pipeline"]
FEATURES = ART["features"]
LOT_COLS = ART["lot_cols"]
MODEL_METRICS = ART.get("metrics", {})

# -------------------------------
# Persistence
# -------------------------------
def init_db():
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""CREATE TABLE IF NOT EXISTS projects(project_id TEXT PRIMARY KEY, project_name TEXT, created_at TEXT, payload TEXT)""")
        con.execute("""CREATE TABLE IF NOT EXISTS progress(project_id TEXT, updated_at TEXT, progress REAL, status TEXT, note TEXT)""")
        con.execute("""CREATE TABLE IF NOT EXISTS photos(project_id TEXT, updated_at TEXT, filename TEXT, data BLOB)""")
        con.commit()

init_db()

# -------------------------------
# Geometry / GIS
# -------------------------------
def parse_coords(text):
    pts = []
    for raw in text.replace(";", "\n").splitlines():
        s = raw.strip().replace("(", "").replace(")", "")
        if not s:
            continue
        parts = [p.strip() for p in s.split(",")]
        if len(parts) < 2:
            parts = s.split()
        if len(parts) >= 2:
            try:
                a, b = float(parts[0]), float(parts[1])
                if -90 <= a <= 90 and -180 <= b <= 180:
                    pts.append((a, b))
            except ValueError:
                pass
    return pts

def parse_uploaded_geometry(uploaded):
    if not uploaded:
        return []
    name = uploaded.name.lower()
    raw = uploaded.getvalue()
    if name.endswith((".geojson", ".json")):
        obj = json.loads(raw.decode("utf-8")); coords = []
        def collect(g):
            typ = g.get("type")
            if typ == "FeatureCollection":
                for f in g.get("features", []): collect(f.get("geometry", {}))
            elif typ == "Feature": collect(g.get("geometry", {}))
            elif typ == "Polygon":
                for c in g.get("coordinates", [[]])[0]: coords.append((float(c[1]), float(c[0])))
            elif typ == "MultiPolygon":
                for ring in g.get("coordinates", []):
                    for c in ring[0]: coords.append((float(c[1]), float(c[0])))
        collect(obj); return coords
    if name.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(raw)); cols = {c.lower().strip(): c for c in df.columns}
        latc = next((cols[k] for k in cols if k in ("lat", "latitude", "y")), None)
        lonc = next((cols[k] for k in cols if k in ("lon", "longitude", "lng", "x")), None)
        if not latc or not lonc:
            raise ValueError("CSV must contain latitude/lat and longitude/lon columns.")
        return [(float(a), float(b)) for a,b in zip(df[latc], df[lonc]) if -90 <= float(a) <= 90 and -180 <= float(b) <= 180]
    import xml.etree.ElementTree as ET
    root = ET.fromstring(raw.decode("utf-8", "ignore")); coords = []
    for el in root.iter():
        tag = el.tag.lower()
        if tag.endswith("coordinates") and el.text:
            for item in el.text.replace("\n", " ").split():
                p = item.split(",")
                if len(p) >= 2: coords.append((float(p[1]), float(p[0])))
        elif tag.endswith("trkpt") or tag.endswith("wpt"):
            if el.attrib.get("lat") and el.attrib.get("lon"):
                coords.append((float(el.attrib["lat"]), float(el.attrib["lon"])))
    return coords

def distance_m(a,b):
    R=6371000.0; p1=math.radians(a[0]); dp=math.radians(b[0]-a[0]); dl=math.radians(b[1]-a[1])
    h=math.sin(dp/2)**2+math.cos(p1)*math.cos(math.radians(b[0]))*math.sin(dl/2)**2
    return 2*R*math.asin(math.sqrt(h))

def bearing(a,b):
    y=math.sin(math.radians(b[1]-a[1]))*math.cos(math.radians(b[0]))
    x=math.cos(math.radians(a[0]))*math.sin(math.radians(b[0]))-math.sin(math.radians(a[0]))*math.cos(math.radians(b[0]))*math.cos(math.radians(b[1]-a[1]))
    return (math.degrees(math.atan2(y,x))+360)%360

def local_xy(lat,lon,lat0,lon0):
    return (lon-lon0)*111320*math.cos(math.radians(lat0)), (lat-lat0)*110540

def polygon_area_m2(poly):
    if len(poly)<3: return 0
    lat0=sum(p[0] for p in poly)/len(poly); lon0=sum(p[1] for p in poly)/len(poly)
    xy=[local_xy(a,b,lat0,lon0) for a,b in poly]
    return abs(sum(xy[i][0]*xy[(i+1)%len(xy)][1]-xy[(i+1)%len(xy)][0]*xy[i][1] for i in range(len(xy)))/2)

@st.cache_data(ttl=3600, show_spinner=False)
def fetch_elevations(points):
    """Fetch elevations from Open-Meteo in small batches.

    Open-Meteo accepts coordinate lists, but sending a whole terrain grid in one
    request can exceed URL/request limits. Batch the points and preserve order.
    """
    if not points:
        return []
    points=[(float(a),float(b)) for a,b in points]
    batch_size=50
    elevations=[]
    for start in range(0,len(points),batch_size):
        batch=points[start:start+batch_size]
        params={
            "latitude":",".join(f"{a:.8f}" for a,_ in batch),
            "longitude":",".join(f"{b:.8f}" for _,b in batch),
        }
        r=requests.get("https://api.open-meteo.com/v1/elevation", params=params, timeout=20)
        r.raise_for_status()
        values=r.json().get("elevation",[])
        if len(values)!=len(batch):
            raise RuntimeError(f"Elevation service returned {len(values)} values for {len(batch)} requested points.")
        elevations.extend(values)
    return elevations

def interpolate_elev(point,sample_points,sample_z):
    if not sample_points: return 0.0
    d=np.array([distance_m(point,p) for p in sample_points]); idx=np.argsort(d)[:min(4,len(d))]; w=1/(d[idx]+0.5)
    return float(np.sum(np.array(sample_z)[idx]*w)/np.sum(w))

def terrain_analysis(poly, dem_bytes=None):
    if len(poly)<3: raise ValueError("At least three GPS points are required.")
    lat0=sum(p[0] for p in poly)/len(poly); lon0=sum(p[1] for p in poly)/len(poly)
    if dem_bytes and rasterio and MemoryFile:
        with MemoryFile(dem_bytes) as mem:
            with mem.open() as src:
                src_crs = src.crs.to_string() if src.crs else "EPSG:4326"
                if src_crs not in ("EPSG:4326", "OGC:CRS84"):
                    from rasterio.warp import transform
                    xs,ys=transform("EPSG:4326",src_crs,[p[1] for p in poly],[p[0] for p in poly])
                    sample_xy=list(zip(xs,ys))
                else:
                    sample_xy=[(p[1],p[0]) for p in poly]
                parcel_z=[]
                for xy in sample_xy:
                    v=list(src.sample([xy]))[0][0]
                    if not np.isfinite(v): raise ValueError("The GeoTIFF does not contain a valid elevation at a parcel vertex.")
                    parcel_z.append(float(v))
                # Build a small geographic grid from the parcel and interpolate the imported DEM.
                glat=np.linspace(min(p[0] for p in poly)-.0005,max(p[0] for p in poly)+.0005,18)
                glon=np.linspace(min(p[1] for p in poly)-.0005,max(p[1] for p in poly)+.0005,18)
                if src_crs not in ("EPSG:4326", "OGC:CRS84"):
                    xs,ys=transform("EPSG:4326",src_crs,glon.tolist(),glat.tolist())
                    # transform() on paired lists is not a Cartesian grid, so sample row by row.
                    Z=np.zeros((len(glat),len(glon)))
                    for i,lat in enumerate(glat):
                        xx,yy=transform("EPSG:4326",src_crs,glon.tolist(),[float(lat)]*len(glon))
                        vals=[list(src.sample([(x,y)]))[0][0] for x,y in zip(xx,yy)]
                        Z[i,:]=np.asarray(vals,dtype=float)
                else:
                    Z=np.zeros((len(glat),len(glon)))
                    for i,lat in enumerate(glat):
                        vals=[list(src.sample([(float(lon),float(lat))]))[0][0] for lon in glon]
                        Z[i,:]=np.asarray(vals,dtype=float)
                source="Imported GeoTIFF DEM"
    else:
        glat=np.linspace(min(p[0] for p in poly)-.0015,max(p[0] for p in poly)+.0015,12)
        glon=np.linspace(min(p[1] for p in poly)-.0015,max(p[1] for p in poly)+.0015,12)
        grid=[(float(a),float(b)) for a in glat for b in glon]; z=fetch_elevations(grid)
        if len(z)!=len(grid): raise RuntimeError("Elevation service returned incomplete data.")
        Z=np.array(z).reshape(len(glat),len(glon)); parcel_z=[interpolate_elev(p,grid,z) for p in poly]
        source="Global DEM / Open-Meteo"
    # Use the parcel vertices as interpolation anchors for side screening.
    if dem_bytes and rasterio and MemoryFile:
        sample_points=poly; sample_z=parcel_z
    else:
        sample_points=grid; sample_z=z
    centroid=(lat0,lon0); relief=max(parcel_z)-min(parcel_z); rows=[]
    for i in range(len(poly)):
        a,b=poly[i],poly[(i+1)%len(poly)]; mid=((a[0]+b[0])/2,(a[1]+b[1])/2)
        inner=(mid[0]+.60*(lat0-mid[0]),mid[1]+.60*(lon0-mid[1]))
        edge=(interpolate_elev(a,sample_points,sample_z)+interpolate_elev(b,sample_points,sample_z))/2
        inner_z=interpolate_elev(inner,sample_points,sample_z)
        delta=abs(inner_z-edge)
        rows.append({"side":f"Côté {i+1}","point_a":a,"point_b":b,"length_m":distance_m(a,b),"edge_elev_m":edge,"inner_elev_m":inner_z,"delta_m":delta,"bearing":bearing(a,b),"retaining":delta>1.5})
    return {"grid_lat":glat,"grid_lon":glon,"Z":Z,"parcel_z":parcel_z,"relief":relief,"rows":rows,"area_m2":polygon_area_m2(poly),"centroid":centroid,"source":source}

def render_map(poly, rows=None):
    center=[sum(p[0] for p in poly)/len(poly),sum(p[1] for p in poly)/len(poly)]
    m=folium.Map(location=center,zoom_start=18,control_scale=True,tiles=None)
    folium.TileLayer(tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",attr="Esri World Imagery",name="Satellite",overlay=False).add_to(m)
    folium.TileLayer(tiles="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",attr="© OpenStreetMap contributors",name="Plan",overlay=False).add_to(m)
    folium.Polygon(poly,color="#2563eb",weight=4,fill=True,fill_opacity=.15,tooltip="Quotora site").add_to(m)
    for r in rows or []:
        if r["retaining"]:
            folium.PolyLine([r["point_a"],r["point_b"]],color="#dc2626",weight=7,tooltip=f"Retaining wall study — {r['length_m']:.1f} m").add_to(m)
    folium.LayerControl().add_to(m)
    return m

def terrain_3d(poly,analysis):
    glat,glon,Z=analysis["grid_lat"],analysis["grid_lon"],analysis["Z"]; lat0,lon0=analysis["centroid"]
    X=np.zeros_like(Z,dtype=float); Y=np.zeros_like(Z,dtype=float)
    for i,a in enumerate(glat):
        for j,b in enumerate(glon): X[i,j],Y[i,j]=local_xy(a,b,lat0,lon0)
    fig=go.Figure(go.Surface(x=X,y=Y,z=Z,colorscale="Earth",showscale=True,colorbar=dict(title="m")))
    pxy=[local_xy(a,b,lat0,lon0) for a,b in poly]; pz=[analysis["parcel_z"][i]+.25 for i in range(len(poly))]
    fig.add_trace(go.Scatter3d(x=[p[0] for p in pxy]+[pxy[0][0]],y=[p[1] for p in pxy]+[pxy[0][1]],z=pz+[pz[0]],mode="lines+markers",line=dict(width=7),name="Boundary"))
    for r in analysis["rows"]:
        if r["retaining"]:
            xa,ya=local_xy(*r["point_a"],lat0,lon0); xb,yb=local_xy(*r["point_b"],lat0,lon0); zz=max(r["edge_elev_m"],r["inner_elev_m"])+.5
            fig.add_trace(go.Scatter3d(x=[xa,xb],y=[ya,yb],z=[zz,zz],mode="lines",line=dict(width=12),name=r["side"]))
    fig.update_layout(height=500,margin=dict(l=0,r=0,t=25,b=0),scene=dict(xaxis_title="E (m)",yaxis_title="N (m)",zaxis_title="Altitude (m)",aspectmode="data"))
    return fig

# -------------------------------
# Estimation engine
# -------------------------------
def allowed_walls(n):
    # Catalogue utilisateur : les règles de compatibilité restent indicatives.
    if n<=1: return WALLS
    if n<=3:
        return [x for x in WALLS if x in ["Parpaing 15","Parpaing 20","Parpaing plein","Parpaing bourré","BTC","BTC stabilisé","Brique cuite","Brique cuite pleine","Béton armé coulé","Béton armé coulé 20 cm","Aluminium vitré","Mur-rideau","Mur mixte parpaing + BA","Mur mixte BTC + BA","Mur bois","Carabot"]]
    if n<=6:
        return [x for x in WALLS if x in ["Parpaing 15","Parpaing 20","Parpaing plein","Béton armé coulé","Béton armé coulé 20 cm","Aluminium vitré","Mur-rideau","Mur mixte parpaing + BA","Mur mixte BTC + BA","Carabot"]]
    return [x for x in WALLS if x in ["Parpaing 15","Parpaing 20","Parpaing plein","Béton armé coulé","Béton armé coulé 20 cm","Aluminium vitré","Mur-rideau","Mur mixte parpaing + BA","Carabot"]]


def build_feature_row(region,rdc,n,surfaces,rdc_type,slabs,wall,opening,roof,ceiling,foundation,soil,standing,season,access,distance,glass,wall_area_actual=None):
    inter=sum(surfaces[1:]); full=sum(surfaces); slab_full=sum(a for a,t in zip(surfaces[1:],slabs) if t==SLABS[0]); slab_hour=sum(a for a,t in zip(surfaces[1:],slabs) if t==SLABS[1])
    wall_area=(float(wall_area_actual) if wall_area_actual is not None else sum(4.2*math.sqrt(max(s,1))*3.0*.82 for s in surfaces))
    wall_regional=WALL_PRICE.get(wall, 0)*REGION_MULT.get(region,1.0)
    return pd.DataFrame([{ "region":region,"surface_rdc_m2":rdc,"niveaux_au_dessus_rdc":n,"surface_totale_construite_m2":full,"surface_planchers_intermediaires_m2":inter,"surface_dallage_rdc_m2":rdc,"surface_dalle_pleine_m2":slab_full,"surface_dalle_hourdis_m2":slab_hour,"type_plancher_rdc":rdc_type,"type_mur":wall,"prix_mur_region_fcfa_m2":wall_regional,"surface_mur_exterieur_m2":wall_area,"part_facade_vitree_pct":glass,"type_ouverture":opening,"type_toiture":roof,"type_plafond":ceiling,"prix_plafond_reference_fcfa_m2":CEIL_PRICE[ceiling],"type_fondation":foundation,"type_sol":soil,"standing":standing,"saison":season,"acces_chantier":access,"distance_transport_km":distance }])

def calculate_elevation_rows(rows):
    """Build the physical elevation inputs and the user's reference cost.
    The resulting area/reference cost are then supplied to the ML estimator;
    they are not themselves the final Lot 3 prediction.
    """
    clean=[]
    total_area=0.0
    total_cost=0.0
    for i,r in enumerate(rows or [],1):
        length=float(r.get("length_m",0) or 0)
        height=float(r.get("height_m",0) or 0)
        openings=float(r.get("openings_m2",0) or 0)
        price=float(r.get("unit_price",0) or 0)
        gross_area=max(length*height,0.0)
        net_area=max(gross_area-openings,0.0)
        amount=net_area*price
        item=dict(r)
        item.update({"gross_area_m2":gross_area,"net_area_m2":net_area,"amount_fcfa":amount})
        clean.append(item)
        total_area+=net_area
        total_cost+=amount
    return clean,total_area,total_cost

def build_internal_elevation_rows(surfaces, wall):
    """Create internal elevation geometry without exposing it in the UI.
    The application uses an indicative perimeter/height/opening ratio to feed
    the Lot 3 ML prediction and its price anchor.
    """
    rows=[]
    price=float(WALL_PRICE.get(wall,10000))
    for idx, surface in enumerate(surfaces):
        surface=max(float(surface or 0),1.0)
        perimeter=4.0*math.sqrt(surface)
        height=3.0
        gross=perimeter*height
        openings=gross*0.15
        rows.append({
            "level": "RDC" if idx==0 else f"R+{idx}",
            "designation": "Élévations internes",
            "wall_type": wall,
            "thickness": WALL_THICKNESS.get(wall,"—"),
            "length_m": perimeter,
            "height_m": height,
            "openings_m2": openings,
            "unit_price": price,
        })
    return rows

def roof_selection_ui(prefix, lang, default_structure="Charpente bois"):
    """Return a single roof descriptor while exposing structure and sheet type separately."""
    structure = st.selectbox(
        "Type de charpente / toiture" if lang == "fr" else "Roof / frame type",
        ROOF_STRUCTURES,
        index=ROOF_STRUCTURES.index(default_structure) if default_structure in ROOF_STRUCTURES else 0,
        format_func=lambda x: label(x, lang),
        key=f"{prefix}_roof_structure"
    )
    sheet = None
    if structure in ("Charpente bois", "Charpente métallique", "Charpente artisanale"):
        sheet = st.selectbox(
            "Type de tôle / couverture" if lang == "fr" else "Sheet / covering type",
            ROOF_SHEETS,
            index=2,
            format_func=lambda x: label(x, lang),
            key=f"{prefix}_roof_sheet"
        )
    if structure == "Toiture-terrasse":
        roof = "Toiture-terrasse accessible"
    else:
        roof = f"{structure} + {sheet}"
    return roof, structure, sheet


def calculate_installation_chantier_cost(surface_m2):
    """Calculate Lot 01 installation de chantier from the site footprint rules supplied by the user."""
    s=max(float(surface_m2 or 0),0.0)
    if s <= 50:
        return 500000.0
    if s <= 100:
        return 650000.0
    if s <= 150:
        return 800000.0
    if s <= 200:
        return 1000000.0
    if s <= 300:
        return 1250000.0
    if s <= 400:
        return 1500000.0
    if s <= 500:
        return 1800000.0
    if s <= 750:
        return 2200000.0
    if s <= 1000:
        return 2700000.0
    return 2700000.0 + (s-1000.0)*2500.0


def estimate(region,rdc,n,standing,wall,opening,roof,rdc_type,slabs,ceiling,foundation,soil,season,access,distance,glass,retaining_cost=0,elevation_rows=None,terrassement=True,terrassement_plain_pied_price=TERRASSEMENT_PLAIN_PIED_DEFAULT,foundation_data=None,vertical_structure_data=None,lot9_data=None,installation_cost=0.0):
    surfaces=[rdc]+[round(rdc*.95,1) for _ in range(n)]
    # User dimensions create the physical elevation area and a reference cost.
    # The final Lot 3 remains an ML prediction, with the reference acting as a
    # construction-price anchor so an unrealistic model output cannot dominate.
    if elevation_rows is None:
        elevation_rows = build_internal_elevation_rows(surfaces, wall)
    elevation_rows,elevation_area,reference_elevation_cost=calculate_elevation_rows(elevation_rows)

    # Use the dominant wall and weighted user price to inform the ML model.
    weights=np.array([max(float(r.get("net_area_m2",0) or 0),0.0) for r in elevation_rows],dtype=float)
    if weights.sum()>0:
        weighted_price=float(sum(max(float(r.get("unit_price",0) or 0),0.0)*w for r,w in zip(elevation_rows,weights))/weights.sum())
        dominant_row=elevation_rows[int(np.argmax(weights))]
        model_wall=str(dominant_row.get("wall_type",wall))
    else:
        weighted_price=float(WALL_PRICE.get(wall,10000)); model_wall=wall
    wall_area_for_model=elevation_area
    X=build_feature_row(region,rdc,n,surfaces,rdc_type,slabs,model_wall,opening,roof,ceiling,foundation,soil,standing,season,access,distance,glass,wall_area_actual=wall_area_for_model)
    # Make the ML model see the actual user-selected weighted wall price, rather than
    # a generic catalogue price. This keeps the prediction tied to the entered walls.
    X.loc[:,"prix_mur_region_fcfa_m2"]=weighted_price
    pred=np.maximum(PIPE.predict(X)[0].astype(float),0)

    # Lot 1 — Terrassement is explicitly optional.
    if len(pred)>0:
        if not terrassement:
            pred[0]=0.0
        elif n==0:
            pred[0]=max(float(rdc),0.0)*max(float(terrassement_plain_pied_price),0.0)
    if len(pred)>1: pred[1]+=retaining_cost

    # Lot 2 — volumes de béton saisis par l'utilisateur. Le même prix de référence
    # au m³ est appliqué aux radiers, semelles, longrines et poteaux. La valeur
    # reste un ancrage de contrôle autour de la prédiction ML.
    foundation_data = foundation_data or {}
    foundation_semelles_volume_m3 = max(float(foundation_data.get("semelles_concrete_volume_m3", 0) or 0), 0.0)
    foundation_longrines_volume_m3 = max(float(foundation_data.get("longrines_concrete_volume_m3", 0) or 0), 0.0)
    foundation_amorces_volume_m3 = max(float(foundation_data.get("amorces_poteaux_concrete_volume_m3", 0) or 0), 0.0)
    foundation_volume_m3 = foundation_semelles_volume_m3 + foundation_longrines_volume_m3 + foundation_amorces_volume_m3
    foundation_unit_price_m3 = float(foundation_data.get("concrete_unit_price_fcfa_m3", FOUNDATION_CONCRETE_PRICE_DEFAULT) or FOUNDATION_CONCRETE_PRICE_DEFAULT)
    foundation_wall_area_m2 = max(float(foundation_data.get("foundation_wall_area_m2", 0) or 0), 0.0)
    foundation_wall_type = str(foundation_data.get("foundation_wall_type", "Parpaing 15 cm"))
    foundation_wall_unit_price = float(foundation_data.get("foundation_wall_unit_price_fcfa_m2", FOUNDATION_WALL_PRICE.get(foundation_wall_type, 0)) or 0)
    foundation_wall_cost = max(float(foundation_data.get("foundation_wall_cost_fcfa", foundation_wall_area_m2 * foundation_wall_unit_price) or 0), 0.0)
    foundation_voile_volume_m3 = max(float(foundation_data.get("foundation_voile_volume_m3", 0) or 0), 0.0)
    foundation_voile_unit_price_m3 = float(foundation_data.get("foundation_voile_unit_price_fcfa_m3", FOUNDATION_VOILE_PRICE_DEFAULT) or FOUNDATION_VOILE_PRICE_DEFAULT)
    foundation_reference_cost = foundation_volume_m3 * foundation_unit_price_m3 + foundation_wall_cost + foundation_voile_volume_m3 * foundation_voile_unit_price_m3
    foundation_prediction_ml = float(pred[1]) if len(pred)>1 else 0.0
    if len(pred)>1 and foundation_reference_cost > 0:
        pred[1] = float(np.clip(foundation_prediction_ml, foundation_reference_cost*0.80, foundation_reference_cost*1.20))

    # Structure verticale intégrée au Lot 3 — Élévations.
    vertical_structure_data = vertical_structure_data or []
    vertical_structure_cost = 0.0
    for item in vertical_structure_data:
        pv = max(float(item.get("poteau_volume_m3", 0) or 0), 0.0)
        pp = max(float(item.get("poteau_price_fcfa_m3", 0) or 0), 0.0)
        lv = max(float(item.get("linteau_volume_m3", 0) or 0), 0.0)
        lp = max(float(item.get("linteau_price_fcfa_m3", 0) or 0), 0.0)
        vertical_structure_cost += pv * pp + lv * lp

    # Lot 3 — ML prediction calibrated by the physical wall area and the user's
    # unit-price reference. The model remains the predictor; the reference provides
    # a reasonable construction-cost envelope and prevents extreme legacy outputs.
    if len(pred)>2:
        if reference_elevation_cost>0:
            reference_elevation_cost += vertical_structure_cost
            lower=reference_elevation_cost*0.80
            upper=reference_elevation_cost*1.20
            lot3_prediction=float(np.clip(pred[2],lower,upper))
        else:
            lot3_prediction=float(pred[2])
        pred[2]=lot3_prediction
    else:
        lot3_prediction=0.0

    # Lot 9 — detailed finishes (professional mode).
    lot9_data = lot9_data or {}
    ceiling_total = 0.0
    if lot9_data.get("ceiling_include", False) and str(lot9_data.get("ceiling_type", "Sans plafond")) != "Sans plafond":
        ceiling_total = max(float(lot9_data.get("ceiling_surface_m2", 0) or 0), 0.0) * max(float(lot9_data.get("ceiling_price_fcfa_m2", 0) or 0), 0.0)
    tile_total = 0.0
    for key in ("tile_interior", "tile_exterior"):
        d = lot9_data.get(key, {}) or {}
        if d.get("include", False):
            tile_total += max(float(d.get("surface_m2", 0) or 0), 0.0) * max(float(d.get("price_fcfa_m2", 0) or 0), 0.0)
    paint_total = 0.0
    for key in ("paint_interior", "paint_exterior"):
        d = lot9_data.get(key, {}) or {}
        if d.get("include", False):
            paint_total += max(float(d.get("surface_m2", 0) or 0), 0.0) * max(float(d.get("price_fcfa_m2", 0) or 0), 0.0)
    lot9_detail_total = ceiling_total + tile_total + paint_total
    if lot9_data:
        pred[8] = lot9_detail_total

    rdc_structure=rdc*RDC_PRICES[rdc_type]; rdc_finish=rdc_structure*RDC_FIN[standing]
    slab_structure=sum(a*SLAB_PRICES[t] for a,t in zip(surfaces[1:],slabs)); slab_finish=sum(a*SLAB_PRICES[t]*SLAB_FIN[standing] for a,t in zip(surfaces[1:],slabs))
    lot11=rdc_structure+rdc_finish+slab_structure+slab_finish
    installation_cost=max(float(installation_cost or 0),0.0)
    gross=float(installation_cost+pred[:4].sum()+rdc_structure+slab_structure); finishes=float(pred[4:].sum()+rdc_finish+slab_finish); total=gross+finishes
    amounts=[installation_cost]+list(pred)+[lot11]
    return {"X":X,"surfaces":surfaces,"pred":pred,"lot11":lot11,"gross":gross,"finishes":finishes,"total":total,"cost_m2":total/sum(surfaces),"rdc_structure":rdc_structure,"rdc_finish":rdc_finish,"slab_structure":slab_structure,"slab_finish":slab_finish,"amounts":amounts,"elevation_rows":elevation_rows,"elevation_area_m2":elevation_area,"elevation_reference_cost_fcfa":reference_elevation_cost,"elevation_cost_fcfa":lot3_prediction,"elevation_prediction_ml":lot3_prediction,"foundation_volume_m3":foundation_volume_m3,"foundation_concrete_unit_price_fcfa_m3":foundation_unit_price_m3,"foundation_reference_cost_fcfa":foundation_reference_cost,"foundation_prediction_ml":foundation_prediction_ml,"vertical_structure_data":vertical_structure_data,"vertical_structure_cost_fcfa":vertical_structure_cost,"lot9_data":lot9_data,"lot9_detail_total_fcfa":lot9_detail_total,"foundation_wall_type":foundation_wall_type,"foundation_wall_area_m2":foundation_wall_area_m2,"foundation_wall_unit_price_fcfa_m2":foundation_wall_unit_price,"foundation_wall_cost_fcfa":foundation_wall_cost,"foundation_voile_volume_m3":foundation_voile_volume_m3,"foundation_voile_unit_price_fcfa_m3":foundation_voile_unit_price_m3,"terrassement_choice":"Oui" if terrassement else "Non","installation_cost_fcfa":installation_cost}


# -------------------------------
# Currency / report
# -------------------------------
@st.cache_data(ttl=3600,show_spinner=False)
def rates_xaf():
    try:
        r=requests.get("https://open.er-api.com/v6/latest/XAF",timeout=8); r.raise_for_status(); data=r.json(); rates=data.get("rates",{})
        return {c:float(rates.get(c,FALLBACK_PER_XAF[c])) for c in CURRENCIES.values()}, "Live exchange-rate service", data.get("time_last_update_utc","")
    except Exception:
        return FALLBACK_PER_XAF, "Indicative fallback rates — offline/service unavailable", dt.datetime.now().strftime("%Y-%m-%d %H:%M")

def print_report_html(project,lang,region,n,standing,soil,foundation,wall,roof,total_surface,gross,finishes,total,cost_m2,df_lots,terrain=None,elevation_rows=None):
    logo=logo_uri(); proposed=[]
    terrain_html=""
    if terrain:
        proposed=[r for r in terrain["rows"] if r["retaining"]]
        terrain_html=f"<h2>Terrain analysis</h2><p>Area: {terrain['area_m2']:,.0f} m² · Relief: {terrain['relief']:.2f} m · Source: {terrain['source']}</p>"
        terrain_html += f"<p><b>Retaining wall study:</b> {sum(r['length_m'] for r in proposed):.1f} linear metres on {len(proposed)} side(s).</p>" if proposed else "<p>No side exceeds the automatic 1.50 m screening threshold.</p>"
    elevation_rows=elevation_rows or []
    rows="".join(f"<tr><td>{r['Lot']}</td><td>{r['Montant FCFA']:,.0f} FCFA</td><td>{r['Part (%)']:.1f}%</td></tr>" for _,r in df_lots.iterrows())
    building=building_data_uri()
    return f"""<!doctype html><html><head><meta charset='utf-8'><title>{REPORT_TITLE}</title><style>@page{{size:A4;margin:14mm}}body{{font-family:Arial;color:#172033;font-size:11px}}.head{{display:flex;gap:15px;align-items:center;border-bottom:3px solid #2563eb;padding-bottom:12px}}.head img{{width:80px;height:80px;object-fit:contain}}h1{{margin:0;font-size:22px}}h2{{color:#183b68;font-size:15px;margin-top:20px}}.meta,.card{{background:#f5f8fc;border:1px solid #e1e8f1;border-radius:8px;padding:10px}}.cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-top:12px}}.label{{font-size:9px;color:#667085;text-transform:uppercase}}.value{{font-size:14px;font-weight:700;margin-top:4px}}.presentation{{margin-top:14px;text-align:center;border:1px solid #e1e8f1;border-radius:10px;padding:8px;background:#fff}}.presentation img{{width:100%;max-height:330px;object-fit:cover;border-radius:7px}}table{{width:100%;border-collapse:collapse;margin-top:8px}}th,td{{border:1px solid #dbe3ec;padding:7px}}th{{background:#eef3f8}}td:nth-child(n+2){{text-align:right}}.print{{position:fixed;right:20px;top:20px;background:#2563eb;color:white;border:0;border-radius:7px;padding:10px 14px}}@media print{{.print{{display:none}}}}</style></head><body><button class='print' onclick='window.print()'>Print / PDF</button><div class='head'><img src='{logo}'><div><h1>{APP_NAME}</h1><b>{REPORT_TITLE}</b><div>Budgetary report — not a contractual quotation</div></div></div><div class='meta'><b>Project:</b> {project} · <b>Region:</b> {region} · <b>Level:</b> R+{n} · <b>Quality:</b> {standing}<br><b>Soil:</b> {soil} · <b>Foundation:</b> {foundation} · <b>Roof:</b> {roof}</div><div class='presentation'><img src='{building}' alt='Project presentation'></div><div class='cards'><div class='card'><div class='label'>Built area</div><div class='value'>{total_surface:,.1f} m²</div></div><div class='card'><div class='label'>Structural works</div><div class='value'>{gross:,.0f} FCFA</div></div><div class='card'><div class='label'>Finishes</div><div class='value'>{finishes:,.0f} FCFA</div></div><div class='card'><div class='label'>Total</div><div class='value'>{total:,.0f} FCFA</div></div></div>{terrain_html}<h2>12 work packages</h2><table><tr><th>Package</th><th>Amount</th><th>Share</th></tr>{rows}</table><h2>Summary</h2><table><tr><th>Structural works</th><td>{gross:,.0f} FCFA</td></tr><tr><th>Finishes</th><td>{finishes:,.0f} FCFA</td></tr><tr><th>Total</th><td><b>{total:,.0f} FCFA</b></td></tr><tr><th>Average cost</th><td>{cost_m2:,.0f} FCFA/m²</td></tr></table><p style='font-size:9px;color:#667085;margin-top:20px'>Quotora is a budgetary estimation and site pre-diagnostic tool. Prices, exchange rates and terrain screening must be verified before contractual use. Retaining-wall screening is not structural design approval.</p></body></html>"""

# -------------------------------
# Export — Excel détaillé / HTML détaillé
# -------------------------------
def build_detailed_rows(res):
    """Build an editable BPU/DQE-style list with Lot 01 dedicated to site installation."""
    rows=[]
    amounts=list(res.get("amounts",[]))
    surfaces=list(res.get("surfaces",[]))
    area=float(res.get("area",0) or 0)
    n=int(res.get("n",0) or 0)

    # Lot 1 — Installation de chantier: user-entered amount, independent from project information.
    installation=float(res.get("installation_cost_fcfa", amounts[0] if amounts else 0) or 0)
    installation_enabled=bool(res.get("installation_enabled", installation>0))
    installation_surface=float(res.get("installation_surface_m2",0) or 0)
    installation_obs=(f"Calcul automatique selon l'emprise au sol : {installation_surface:,.2f} m²." if installation_enabled and installation_surface>0 else "Installation de chantier non retenue : 0 FCFA.")
    rows.append(["1 — Installation de chantier","Installation de chantier","Forfait",1,installation,installation,installation_obs])

    # Lot 2 — Terrassement (old Lot 1)
    lot2=float(amounts[1]) if len(amounts)>1 else 0.0
    if res.get("terrassement_choice") == "Oui":
        if n==0 and area>0:
            pu=float(res.get("terrassement_plain_pied_price",0) or 0)
            rows.append(["2 — Terrassement","Terrassement général plain-pied","m²",area,pu,area*pu,"Calcul automatique : surface RDC × prix terrassement"])
        else:
            rows.append(["2 — Terrassement","Terrassement — montant estimé ML","Forfait",1,lot2,lot2,"Quantités détaillées de fouilles non saisies; montant ML conservé comme base"])
        rows.append(["2 — Terrassement","Fouille / décapage — à compléter","m³",0,0,0,"Ligne éditable dans Excel"])
    else:
        rows.append(["2 — Terrassement","Terrassement non retenu","Forfait",1,0,0,"Option Terrassement = Non"])

    # Lot 3 — Fondations (old Lot 2)
    lot3=float(amounts[2]) if len(amounts)>2 else 0.0
    fw_area=float(res.get("foundation_wall_area_m2",0) or 0)
    fw_pu=float(res.get("foundation_wall_unit_price_fcfa_m2",0) or 0)
    fw_type=res.get("foundation_wall_type","")
    foundation_items=[]
    if fw_area>0:
        foundation_items.append(["3 — Fondations",f"Mur de fondation — {fw_type}","m²",fw_area,fw_pu,fw_area*fw_pu,"Surface = longueur développée × profondeur"])
    fpu=float(res.get("foundation_concrete_unit_price_fcfa_m3",0) or 0)
    fsem=float(res.get("foundation_semelles_volume_m3",0) or 0)
    flon=float(res.get("foundation_longrines_volume_m3",0) or 0)
    famo=float(res.get("foundation_amorces_volume_m3",0) or 0)
    if fsem>0:
        foundation_items.append(["3 — Fondations","Béton des semelles","m³",fsem,fpu,fsem*fpu,"Volume saisi pour les semelles"])
    if flon>0:
        foundation_items.append(["3 — Fondations","Béton des longrines","m³",flon,fpu,flon*fpu,"Volume saisi pour les longrines"])
    if famo>0:
        foundation_items.append(["3 — Fondations","Béton des amorces de poteaux","m³",famo,fpu,famo*fpu,"Volume saisi pour les amorces de poteaux"])
    if res.get("foundation_voile_volume_m3",0):
        vv=float(res.get("foundation_voile_volume_m3",0) or 0); vpu=float(res.get("foundation_voile_unit_price_fcfa_m3",0) or 0)
        foundation_items.append(["3 — Fondations","Voile BA","m³",vv,vpu,vv*vpu,"Prix de référence voile BA"])
    rows.extend(foundation_items)
    used=sum(float(r[5]) for r in foundation_items)
    rows.append(["3 — Fondations","Autres travaux de fondation — ajustement ML","Forfait",1,max(0,lot3-used),max(0,lot3-used),"Ajustement permettant de conserver le montant du lot prédit"])

    # Lot 4 — Élévation (old Lot 3)
    lot4=float(amounts[3]) if len(amounts)>3 else 0.0
    elev=res.get("elevation_rows") or []
    elev_sum=0.0
    for r in elev:
        q=float(r.get("net_area_m2",0) or (float(r.get("length_m",0) or 0)*float(r.get("height_m",0) or 0)))
        pu=float(r.get("unit_price",0) or 0); total=q*pu; elev_sum+=total
        rows.append(["4 — Élévation",r.get("designation","Groupe de mur"),"m²",q,pu,total,"Surface = longueur développée × hauteur"])
    vertical_sum=0.0
    for v in (res.get("vertical_structure_data") or []):
        level=v.get("level","Étage")
        pv=float(v.get("poteau_volume_m3",0) or 0); pp=float(v.get("poteau_price_fcfa_m3",0) or 0)
        lv=float(v.get("linteau_volume_m3",0) or 0); lp=float(v.get("linteau_price_fcfa_m3",0) or 0)
        vertical_sum += pv*pp + lv*lp
        rows.append(["4 — Élévation",f"{level} — Structure verticale — béton des poteaux","m³",pv,pp,pv*pp,"Structure verticale intégrée à l'élévation"])
        rows.append(["4 — Élévation",f"{level} — Linteaux — béton","m³",lv,lp,lv*lp,"Linteaux intégrés à l'élévation"])
    rows.append(["4 — Élévation","Ajustement prédiction ML du Lot 4","Forfait",1,max(0,lot4-elev_sum-vertical_sum),max(0,lot4-elev_sum-vertical_sum),"Ajustement de cohérence du montant du lot"])

    # Lots 5-10 — correspond aux anciens lots 4-9.
    for idx in range(4,10):
        amount_index=idx+1  # amounts[5..10]
        amt=float(amounts[amount_index]) if amount_index<len(amounts) else 0.0
        lot_number=idx+1
        lot_label=LOTS[lot_number-1]
        if idx == 8 and res.get("lot9_data"):
            d=res.get("lot9_data") or {}
            rows.append(["10 — Carrelage, faux plafond & peinture",f"Faux plafond — {d.get('ceiling_type','Sans plafond')}","m²",float(d.get("ceiling_surface_m2",0) or 0),float(d.get("ceiling_price_fcfa_m2",0) or 0),(float(d.get("ceiling_surface_m2",0) or 0)*float(d.get("ceiling_price_fcfa_m2",0) or 0) if d.get("ceiling_include",False) and d.get("ceiling_type")!="Sans plafond" else 0),"0 FCFA si Sans plafond / non intégré"])
            if d.get("tile_include",False):
                rows.append(["10 — Carrelage, faux plafond & peinture","Carrelage — dallage RDC + dalles/planchers","m²",float(d.get("tile_surface_m2",0) or 0),float(d.get("tile_price_fcfa_m2",0) or 0),float(d.get("tile_surface_m2",0) or 0)*float(d.get("tile_price_fcfa_m2",0) or 0),"Base automatique : dallage RDC + tous les niveaux supérieurs"])
            pd=d.get("paint_detail",{}) or {}
            if pd.get("detail",False):
                for key,labeltxt in (("paint_interior","Peinture intérieure"),("paint_exterior","Peinture extérieure")):
                    q=d.get(key,{}) or {}
                    if q.get("include",False):
                        rows.append(["10 — Carrelage, faux plafond & peinture",labeltxt,"m²",float(q.get("surface_m2",0) or 0),float(q.get("price_fcfa_m2",0) or 0),float(q.get("surface_m2",0) or 0)*float(q.get("price_fcfa_m2",0) or 0),"Peinture détaillée"])
            elif pd and pd.get("include",False):
                rows.append(["10 — Carrelage, faux plafond & peinture","Peinture — toutes surfaces de murs","m²",float(pd.get("surface_m2",0) or 0),float(pd.get("price_fcfa_m2",0) or 0),float(pd.get("surface_m2",0) or 0)*float(pd.get("price_fcfa_m2",0) or 0),"Mode peinture non détaillé"])
        else:
            rows.append([lot_label,"Estimation ML du lot","Forfait",1,amt,amt,"Montant ML — quantité détaillée non saisie"])

    # Lot 12 — floor/slabs: old Lot 11.
    rdc_struct=float(res.get("rdc_structure",0) or 0); rdc_finish=float(res.get("rdc_finish",0) or 0)
    slab_struct=float(res.get("slab_structure",0) or 0); slab_finish=float(res.get("slab_finish",0) or 0)
    if area>0 and rdc_struct:
        rows.append(["12 — Planchers, dalles & dallages","RDC — dallage / plancher","m²",area,rdc_struct/area,rdc_struct,"Prix moyen calculé"])
        rows.append(["12 — Planchers, dalles & dallages","RDC — finitions","m²",area,rdc_finish/area,rdc_finish,"Selon standing"])
    inter_area=max(sum(surfaces[1:]) if len(surfaces)>1 else 0,0)
    if inter_area>0 and slab_struct:
        rows.append(["12 — Planchers, dalles & dallages","Dalles intermédiaires","m²",inter_area,slab_struct/inter_area,slab_struct,"Selon type de dalle sélectionné"])
        rows.append(["12 — Planchers, dalles & dallages","Finitions planchers intermédiaires","m²",inter_area,slab_finish/inter_area,slab_finish,"Selon standing"])
    return rows


def detailed_excel_bytes(res, lang="fr"):
    wb=Workbook()
    ws=wb.active; ws.title="Présentation"
    thin=Side(style="thin",color="D9E2EC")
    header_fill=PatternFill("solid",fgColor="163A63")
    sub_fill=PatternFill("solid",fgColor="EAF1F8")
    title_font=Font(size=18,bold=True,color="FFFFFF")
    header_font=Font(bold=True,color="FFFFFF")
    ws.merge_cells("A1:G1"); ws["A1"]="QUOTORA PRO — DEVIS DÉTAILLÉ / BORDEREAU QUANTITATIF"; ws["A1"].font=title_font; ws["A1"].fill=header_fill; ws["A1"].alignment=Alignment(horizontal="center")
    meta=[("Projet",res.get("project","")),("Région",res.get("region","")),("Niveau",f"R+{res.get('n',0)}"),("Surface RDC (m²)",res.get("area",0)),("Standing",res.get("standing","")),("Terrassement",res.get("terrassement_choice","Non"))]
    for i,(k,v) in enumerate(meta,3): ws.cell(i,1,k).font=Font(bold=True); ws.cell(i,2,v)
    ws["A10"]="DÉTAIL QUANTITATIF"; ws["A10"].font=Font(size=14,bold=True,color="163A63")
    headers=["Lot","Désignation","Unité","Quantité","Prix unitaire FCFA","Montant FCFA","Observation / source"]
    for j,h in enumerate(headers,1): ws.cell(11,j,h).fill=header_fill; ws.cell(11,j).font=header_font; ws.cell(11,j).alignment=Alignment(horizontal="center")
    rows=build_detailed_rows(res)
    for i,row in enumerate(rows,12):
        for j,val in enumerate(row,1): ws.cell(i,j,val); ws.cell(i,j).border=Border(bottom=thin)
    for col in range(1,8): ws.column_dimensions[get_column_letter(col)].width=[26,42,12,14,20,20,48][col-1]
    for row in range(12,12+len(rows)): ws.cell(row,4).number_format='#,##0.00'; ws.cell(row,5).number_format='#,##0'; ws.cell(row,6).number_format='#,##0'
    ws.freeze_panes="A12"; ws.auto_filter.ref=f"A11:G{11+len(rows)}"
    # Summary sheet
    s=wb.create_sheet("Synthèse")
    s.append(["QUOTORA PRO — SYNTHÈSE"]); s["A1"].font=Font(size=16,bold=True,color="163A63")
    s.append([]); s.append(["Lot","Montant FCFA","Part (%)"])
    for j in range(1,4): s.cell(3,j).fill=header_fill; s.cell(3,j).font=header_font
    for idx,(lot,amt) in enumerate(zip(LOTS,res.get("amounts",[])),4): s.append([lot,float(amt),float(amt)/float(res.get("total",1))*100 if res.get("total") else 0])
    s.append(["Gros œuvre",float(res.get("gross",0)),None]); s.append(["Finitions",float(res.get("finishes",0)),None]); s.append(["TOTAL",float(res.get("total",0)),100])
    for r in range(4,s.max_row+1): s.cell(r,2).number_format='#,##0'; s.cell(r,3).number_format='0.0%'
    s.column_dimensions['A'].width=40; s.column_dimensions['B'].width=22; s.column_dimensions['C'].width=14
    # Input data sheet for traceability/editing
    d=wb.create_sheet("Données projet")
    for r,(k,v) in enumerate(sorted(json_safe_result(res).items()),1):
        if isinstance(v,(dict,list)): v=json.dumps(v,ensure_ascii=False,default=str)
        d.cell(r,1,k); d.cell(r,2,v); d.cell(r,1).font=Font(bold=True)
    d.column_dimensions['A'].width=42; d.column_dimensions['B'].width=80
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()


def summary_excel_bytes(res, lang="fr"):
    """Export Excel autonome contenant uniquement la synthèse des 12 lots
    (distinct du devis détaillé — corrige un bug où les deux boutons de
    téléchargement renvoyaient le même fichier)."""
    wb=Workbook()
    s=wb.active; s.title="Synthèse"
    header_fill=PatternFill("solid",fgColor="163A63")
    header_font=Font(bold=True,color="FFFFFF")
    s.merge_cells("A1:C1")
    s["A1"]="QUOTORA PRO — SYNTHÈSE 12 LOTS" if lang=="fr" else "QUOTORA PRO — 12-PACKAGE SUMMARY"
    s["A1"].font=Font(size=16,bold=True,color="FFFFFF"); s["A1"].fill=header_fill; s["A1"].alignment=Alignment(horizontal="center")
    meta=[("Projet" if lang=="fr" else "Project",res.get("project","")),("Région" if lang=="fr" else "Region",res.get("region","")),("Niveau" if lang=="fr" else "Level",f"R+{res.get('n',0)}"),("Standing" if lang=="fr" else "Quality level",res.get("standing",""))]
    for i,(k,v) in enumerate(meta,3): s.cell(i,1,k).font=Font(bold=True); s.cell(i,2,v)
    header_row=8
    headers=["Lot","Montant FCFA","Part (%)"]
    for j,h in enumerate(headers,1): s.cell(header_row,j,h).fill=header_fill; s.cell(header_row,j).font=header_font; s.cell(header_row,j).alignment=Alignment(horizontal="center")
    total=float(res.get("total",0) or 0)
    for i,(lot,amt) in enumerate(zip(LOTS,res.get("amounts",[])),header_row+1):
        s.cell(i,1,lot); s.cell(i,2,float(amt)); s.cell(i,2).number_format='#,##0'
        s.cell(i,3,float(amt)/total if total else 0); s.cell(i,3).number_format='0.0%'
    r=header_row+1+len(res.get("amounts",[]))
    s.cell(r,1,"Gros œuvre" if lang=="fr" else "Structural works").font=Font(bold=True); s.cell(r,2,float(res.get("gross",0))).number_format='#,##0'
    s.cell(r+1,1,"Finitions" if lang=="fr" else "Finishes").font=Font(bold=True); s.cell(r+1,2,float(res.get("finishes",0))).number_format='#,##0'
    s.cell(r+2,1,"TOTAL").font=Font(bold=True,color="163A63"); s.cell(r+2,2,total).number_format='#,##0'
    s.column_dimensions['A'].width=42; s.column_dimensions['B'].width=22; s.column_dimensions['C'].width=14
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()


def detailed_html(res, lang="fr"):
    rows=build_detailed_rows(res)
    tr_rows="".join(f"<tr><td>{r[0]}</td><td contenteditable='true'>{r[1]}</td><td>{r[2]}</td><td contenteditable='true'>{r[3]:,.2f}</td><td contenteditable='true'>{r[4]:,.0f}</td><td>{r[5]:,.0f}</td><td>{r[6]}</td></tr>" for r in rows)
    return f"""<!doctype html><html><head><meta charset='utf-8'><title>Devis détaillé — {res.get('project','')}</title><style>body{{font-family:Arial;color:#172033;margin:30px}}h1{{color:#163A63}}.note{{background:#f5f8fc;padding:12px;border:1px solid #dbe3ec;border-radius:8px}}table{{width:100%;border-collapse:collapse;margin-top:18px}}th,td{{border:1px solid #dbe3ec;padding:7px}}th{{background:#163A63;color:#fff}}td:nth-child(4),td:nth-child(5),td:nth-child(6){{text-align:right}}td[contenteditable='true']{{background:#fffdf2;outline:none}}button{{padding:10px 16px;border:0;border-radius:7px;background:#2563eb;color:#fff;font-weight:700}}@media print{{button,.note{{display:none}}}}</style></head><body><button onclick='window.print()'>Imprimer / PDF</button><h1>QUOTORA PRO — Devis détaillé / Bordereau quantitatif</h1><div class='note'>Document budgétaire éditable. Les cellules jaunes peuvent être modifiées directement dans le navigateur. Les quantités d'ingénierie non saisies dans QUOTORA sont laissées à compléter plutôt que d'être inventées.</div><p><b>Projet :</b> {res.get('project','')} · <b>Région :</b> {res.get('region','')} · <b>Niveau :</b> R+{res.get('n',0)} · <b>Surface RDC :</b> {res.get('area',0):,.2f} m²</p><table><tr><th>Lot</th><th>Désignation</th><th>Unité</th><th>Quantité</th><th>Prix unitaire FCFA</th><th>Montant FCFA</th><th>Observation</th></tr>{tr_rows}</table><h2>Synthèse</h2><p><b>Gros œuvre :</b> {res.get('gross',0):,.0f} FCFA</p><p><b>Finitions :</b> {res.get('finishes',0):,.0f} FCFA</p><p><b>Total :</b> {res.get('total',0):,.0f} FCFA</p><p style='color:#667085;font-size:11px'>Rapport budgétaire — sans valeur de devis contractuel. Vérifier les métrés et prix avant utilisation contractuelle.</p></body></html>"""

# -------------------------------
# Project storage helpers
# -------------------------------
def capture_quick_input_state(region,btype,area,n,standing,currency,wall,terrassement_choice,terrassement_pp_price,roof_structure,roof_sheet,foundation="Semelles isolées + longrines",foundation_wall_length=0.0,foundation_depth=0.80,foundation_wall_type="Parpaing 15 cm",foundation_wall_unit_price=15000.0,foundation_voile_thickness=0.20,foundation_voile_volume=0.0,foundation_voile_unit_price=210000.0,foundation_volume_total=0.0,foundation_concrete_unit_price=195000.0,elevation_rows=None):
    state={"mode":"quick","q_installation_cost":float(st.session_state.get("q_installation_cost",0.0) or 0.0),"q_installation_enabled":st.session_state.get("q_installation_enabled","Non"),"q_installation_surface":float(st.session_state.get("q_installation_surface",0.0) or 0.0),"q_region":region,"q_btype":btype,"q_area":float(area),"q_levels":int(n),"q_standing":standing,"q_currency":currency,"q_wall":wall,"q_earthwork":terrassement_choice,"q_earthwork_pp_price":float(terrassement_pp_price),"q_roof_structure":roof_structure,"q_roof_sheet":roof_sheet,
           "q_found":foundation,"q_found_wall_len":float(foundation_wall_length),"q_found_depth":float(foundation_depth),"q_found_wall_type":foundation_wall_type,"q_found_wall_price":float(foundation_wall_unit_price),"q_found_voile_thickness":float(foundation_voile_thickness),"q_found_voile_volume":float(foundation_voile_volume),"q_found_voile_price":float(foundation_voile_unit_price),"q_foundation_total_volume":float(foundation_volume_total),"q_foundation_m3_price":float(foundation_concrete_unit_price),"q_elev_nb":int(len(elevation_rows or []))}
    for i,row in enumerate(elevation_rows or []):
        state[f"q_el_length_{i}"]=float(row.get("length_m",0) or 0)
        state[f"q_el_height_{i}"]=float(row.get("height_m",0) or 0)
    # Conserver toutes les valeurs des widgets rapides pour une restauration
    # complète depuis le fichier JSON.
    for _k, _v in st.session_state.items():
        if _k.startswith("q_") and _k not in state:
            try:
                json.dumps(_v, ensure_ascii=False, default=str)
                state[_k] = _v
            except Exception:
                pass
    return state

def capture_professional_input_state(project,region,btype,n,area,standing,wall,terrassement_choice,terrassement_pp_price,roof_structure,roof_sheet,opening,rdc_type,ceiling,foundation,soil,season,access,distance,foundation_wall_length,foundation_depth,foundation_wall_type,foundation_wall_unit_price,foundation_voile_thickness,foundation_voile_volume,foundation_voile_unit_price,foundation_volume_total,foundation_concrete_unit_price,elevation_rows,vertical_structure_data=None,lot9_data=None,foundation_semelles_volume=0.0,foundation_longrines_volume=0.0,foundation_amorces_volume=0.0):
    state={"mode":"professional","p_project":project,"p_installation_cost":float(st.session_state.get("p_installation_cost",0.0) or 0.0),"p_installation_enabled":st.session_state.get("p_installation_enabled","Non"),"p_installation_surface":float(st.session_state.get("p_installation_surface",0.0) or 0.0),"p_region":region,"p_btype":btype,"p_n":int(n),"p_area":float(area),"p_standing":standing,"p_wall":wall,"p_earthwork":terrassement_choice,"p_earthwork_pp_price":float(terrassement_pp_price),"p_roof_structure":roof_structure,"p_roof_sheet":roof_sheet,"p_open":opening,"p_rdc":rdc_type,"p_ceil":ceiling,"p_found":foundation,"p_soil":soil,"p_season":season,"p_access":access,"p_dist":float(distance),"p_id":project,"p_found_wall_len":float(foundation_wall_length),"p_found_depth":float(foundation_depth),"p_found_wall_type":foundation_wall_type,"p_found_wall_price":float(foundation_wall_unit_price),"p_found_voile_thickness":float(foundation_voile_thickness),"p_found_voile_volume":float(foundation_voile_volume),"p_found_voile_price":float(foundation_voile_unit_price),"p_foundation_total_volume":float(foundation_volume_total),"p_foundation_semelles_volume":float(foundation_semelles_volume),"p_foundation_longrines_volume":float(foundation_longrines_volume),"p_foundation_amorces_volume":float(foundation_amorces_volume),"p_foundation_m3_price":float(foundation_concrete_unit_price),"p_elev_nb":int(len(elevation_rows))}
    for i,row in enumerate(elevation_rows): state[f"p_el_length_{i}"]=float(row.get("length_m",0) or 0); state[f"p_el_height_{i}"]=float(row.get("height_m",0) or 0)
    state["p_vertical_structure_data"]=vertical_structure_data or []; state["p_lot9_data"]=lot9_data or {}
    # Conserver également toutes les valeurs des widgets professionnels (y compris
    # les options de finitions) afin qu'un JSON exporté puisse être réimporté
    # puis modifié sans perdre de données.
    for _k, _v in st.session_state.items():
        if _k.startswith("p_") and _k not in state:
            try:
                json.dumps(_v, ensure_ascii=False, default=str)
                state[_k] = _v
            except Exception:
                pass
    return state

def restore_saved_state(saved):
    state=saved.get("input_state") if isinstance(saved,dict) else None
    if not state:
        mode=saved.get("mode","quick") if isinstance(saved,dict) else "quick"; prefix="p_" if mode=="professional" else "q_"; state={"mode":mode}
        mapping={f"{prefix}region":saved.get("region"),f"{prefix}btype":saved.get("btype"),f"{prefix}area":saved.get("area"),f"{prefix}levels":saved.get("n"),f"{prefix}standing":saved.get("standing"),f"{prefix}wall":saved.get("wall"),f"{prefix}earthwork":saved.get("terrassement_choice","Non")}
        state.update({k:v for k,v in mapping.items() if v is not None})
        if mode=="professional": state.update({"p_project":saved.get("project","PROJET-001"),"p_id":saved.get("project","PROJET-001")})
    for key,value in state.items():
        if key!="mode" and value is not None: st.session_state[key]=value
    if saved.get("terrain") is not None: st.session_state.terrain=saved.get("terrain")
    st.session_state.loaded_mode=state.get("mode","professional"); st.session_state.estimate_result=saved

def json_safe_result(res):
    payload={k:v for k,v in res.items() if k not in ("X","pred")}
    payload["quotora_file_type"]="QUOTORA_PROJECT"
    payload["quotora_schema_version"]="2.0"
    payload["exported_at"]=dt.datetime.now().isoformat(timespec="seconds")
    return payload

def save_project_result(res):
    payload={k:v for k,v in res.items() if k not in ("X","pred")}
    payload["amounts"]=list(map(float,res.get("amounts",[])))
    payload["surfaces"]=list(map(float,res.get("surfaces",[])))
    with sqlite3.connect(DB_PATH) as con:
        con.execute("INSERT OR REPLACE INTO projects(project_id,project_name,created_at,payload) VALUES (?,?,?,?)",(res["project"],res["project"],dt.datetime.now().isoformat(timespec="seconds"),json.dumps(payload,ensure_ascii=False,default=str)))
        con.commit()

# -------------------------------
# Access control — codes Q14 (14 jours, 2 000 FCFA)
# -------------------------------
# Format du code : Q14-AAMMJJ-XXXX-SSSSSSSSSS
#   AAMMJJ     = date d'émission (fuseau Africa/Douala)
#   XXXX       = identifiant aléatoire (rend chaque code unique)
#   SSSSSSSSSS = signature HMAC-SHA256 (clé secrète ACCESS_SECRET) — impossible à forger sans la clé
# L'accès expire à la fin du 14e jour suivant la date d'émission (heure de Douala).
ACCESS_PREFIX = "Q14"
ACCESS_DAYS = 14
ACCESS_PRICE_FCFA = 2000
try:
    DOUALA_TZ = ZoneInfo("Africa/Douala")
except Exception:  # tzdata absent : Douala = UTC+1 toute l'année, sans heure d'été
    DOUALA_TZ = dt.timezone(dt.timedelta(hours=1), "Africa/Douala")

def _access_secret():
    try:
        sec = st.secrets.get("ACCESS_SECRET", None)
    except Exception:
        sec = None
    sec = sec or os.environ.get("ACCESS_SECRET")
    return str(sec).strip() if sec else None

def _access_signature(secret, body):
    digest = hmac.new(secret.encode("utf-8"), body.encode("utf-8"), hashlib.sha256).digest()
    return base64.b32encode(digest).decode("ascii")[:10]

def verify_access_code(code, secret, now=None):
    """Retourne (statut, date_expiration). statut ∈ {'ok','invalid','expired'}."""
    now = now or dt.datetime.now(DOUALA_TZ)
    parts = (code or "").strip().upper().split("-")
    if len(parts) != 4 or parts[0] != ACCESS_PREFIX or len(parts[1]) != 6 or not parts[1].isdigit() or len(parts[2]) != 4 or len(parts[3]) != 10:
        return "invalid", None
    body = "-".join(parts[:3])
    if not hmac.compare_digest(_access_signature(secret, body), parts[3]):
        return "invalid", None
    try:
        issued = dt.datetime.strptime(parts[1], "%y%m%d").date()
    except ValueError:
        return "invalid", None
    if issued > now.date() + dt.timedelta(days=1):  # émission dans le futur : rejeté
        return "invalid", None
    expiry = dt.datetime.combine(issued + dt.timedelta(days=ACCESS_DAYS), dt.time(23, 59, 59), tzinfo=DOUALA_TZ)
    if now > expiry:
        return "expired", expiry
    return "ok", expiry

def access_gate():
    secret = _access_secret()
    if not secret:
        # Aucun ACCESS_SECRET : mode libre (développement local).
        st.session_state.access_gate_configured = False
        return
    st.session_state.access_gate_configured = True
    gate_lang = "fr" if st.session_state.get("lang", "fr") == "fr" else "en"
    notice = None

    # Revalidation à CHAQUE exécution : si la date d'expiration est dépassée, l'application se bloque.
    code = st.session_state.get("access_code") or st.query_params.get("code")
    if code:
        status, expiry = verify_access_code(code, secret)
        if status == "ok":
            st.session_state.access_code = code.strip().upper()
            st.session_state.access_expiry = expiry
            return
        st.session_state.pop("access_code", None)
        st.session_state.pop("access_expiry", None)
        if "code" in st.query_params:
            del st.query_params["code"]
        if status == "expired":
            d = expiry.strftime("%d/%m/%Y")
            notice = (f"⏳ Votre accès a expiré le {d}. Obtenez un nouveau code pour continuer." if gate_lang == "fr"
                      else f"⏳ Your access expired on {d}. Get a new code to continue.")

    st.markdown("<div style='max-width:420px;margin:80px auto;text-align:center'>", unsafe_allow_html=True)
    if LOGO_PATH.exists(): st.image(str(LOGO_PATH), width=140)
    st.markdown(f"<h2>{APP_NAME}</h2><p style='color:#667085'>{'Accès protégé — 14 jours' if gate_lang=='fr' else 'Protected access — 14 days'}</p>", unsafe_allow_html=True)
    if notice: st.warning(notice)
    entered = st.text_input("Code d'accès (Q14-…)" if gate_lang == "fr" else "Access code (Q14-…)", key="gate_code_input", placeholder="Q14-XXXXXX-XXXX-XXXXXXXXXX")
    if st.button("Valider" if gate_lang == "fr" else "Submit", type="primary", use_container_width=True, key="gate_submit"):
        status, expiry = verify_access_code(entered, secret)
        if status == "ok":
            st.session_state.access_code = entered.strip().upper()
            st.session_state.access_expiry = expiry
            st.query_params["code"] = entered.strip().upper()  # survit à l'actualisation de la page
            st.rerun()
        elif status == "expired":
            st.error(("Ce code a expiré le " if gate_lang == "fr" else "This code expired on ") + expiry.strftime("%d/%m/%Y") + ".")
        else:
            st.error("Code invalide." if gate_lang == "fr" else "Invalid code.")
    st.caption(f"Accès {ACCESS_DAYS} jours — {ACCESS_PRICE_FCFA:,} FCFA".replace(",", " ") + " via Mobile Money / Orange Money. Contactez le fondateur pour obtenir un code." if gate_lang == "fr"
               else f"{ACCESS_DAYS}-day access — {ACCESS_PRICE_FCFA:,} FCFA via Mobile Money. Contact the founder for a code.")
    lang_toggle = st.radio("Language / Langue", ["🇫🇷 Français", "🇬🇧 English"], index=0 if gate_lang == "fr" else 1, horizontal=True, label_visibility="collapsed", key="gate_lang_toggle")
    st.session_state.lang = "fr" if lang_toggle.startswith("🇫🇷") else "en"
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

# -------------------------------
# State / navigation
# -------------------------------
if "lang" not in st.session_state: st.session_state.lang="fr"
if "page" not in st.session_state: st.session_state.page="home"
if "estimate_result" not in st.session_state: st.session_state.estimate_result=None
if "terrain" not in st.session_state: st.session_state.terrain=None

inject_css()
access_gate()

# Sidebar
with st.sidebar:
    if LOGO_PATH.exists(): st.image(str(LOGO_PATH),width=125)
    st.markdown(f"### {APP_NAME}")
    _exp = st.session_state.get("access_expiry")
    if _exp:
        _left = (_exp.date() - dt.datetime.now(DOUALA_TZ).date()).days
        st.caption((f"🔑 Accès valide jusqu'au {_exp.strftime('%d/%m/%Y')} ({_left} j restants)") if st.session_state.lang == "fr"
                   else f"🔑 Access valid until {_exp.strftime('%d/%m/%Y')} ({_left} days left)")
    lang_label=st.radio("Language / Langue",["🇫🇷 Français","🇬🇧 English"],index=0 if st.session_state.lang=="fr" else 1,horizontal=True,label_visibility="collapsed")
    st.session_state.lang="fr" if lang_label.startswith("🇫🇷") else "en"
    lang=st.session_state.lang
    st.divider()
    nav=[("home","🏠",tr("home",lang)),("quick","⚡",tr("quick",lang)),("professional","⚙️",tr("professional",lang)),("projects","📁",tr("projects",lang)),("terrain","🗺️",tr("terrain",lang)),("tracking","📈",tr("tracking",lang)),("reports","📄",tr("reports",lang)),("settings","⚙️",tr("settings",lang))]
    for key,icon,title in nav:
        if st.button(f"{icon}  {title}",key=f"nav_{key}",use_container_width=True,type="primary" if st.session_state.page==key else "secondary"):
            st.session_state.page=key; st.rerun()
    st.divider()
    st.caption("💾 Enregistrez puis réimportez un projet pour reprendre et modifier les données." if lang=="fr" else "💾 Save then re-import a project to resume and edit its data.")
    st.caption(tr("founder",lang))
    st.caption("© 2026 Quotora")

# Hero on home only
if st.session_state.page=="home":
    st.markdown(f"<div class='q-hero'><span class='q-badge'>QUOTORA PRO</span><h1>{'Build smarter' if lang=='en' else 'Construisez plus intelligemment'}</h1><p>{'Estimation · Terrain · Reports · Site progress' if lang=='en' else 'Estimation · Terrain · Rapports · Suivi de chantier'}</p></div>",unsafe_allow_html=True)

# -------------------------------
# Home
# -------------------------------
def home_page():
    st.markdown(f"<h2 style='text-align:center'>{'What would you like to do?' if lang=='en' else 'Que souhaitez-vous faire ?'}</h2><p style='text-align:center;color:#667085'>{'Choose only the workflow you need.' if lang=='en' else 'Choisissez uniquement le parcours dont vous avez besoin.'}</p>",unsafe_allow_html=True)
    a,b=st.columns(2)
    with a:
        st.markdown(f"<div class='q-card q-mode'><h3>⚡ {tr('quick',lang)}</h3><p>{tr('quick_sub',lang)}</p><p class='q-small'>{tr('quick_hint',lang)}</p></div>",unsafe_allow_html=True)
        if st.button(tr("start",lang)+" — ⚡",key="home_quick",use_container_width=True): st.session_state.page="quick"; st.rerun()
    with b:
        st.markdown(f"<div class='q-card q-mode'><h3>⚙️ {tr('professional',lang)}</h3><p>{tr('professional_sub',lang)}</p><p class='q-small'>{tr('pro_hint',lang)}</p></div>",unsafe_allow_html=True)
        if st.button(tr("start",lang)+" — ⚙️",key="home_pro",use_container_width=True): st.session_state.page="professional"; st.rerun()
    st.write("")
    c1,c2,c3,c4=st.columns(4)
    cards=[("🗺️",tr("terrain",lang),"terrain"),("📈",tr("tracking",lang),"tracking"),("📄",tr("reports",lang),"reports"),("💱",tr("currency",lang),"results")]
    for col,(ico,title,key) in zip([c1,c2,c3,c4],cards):
        with col:
            st.markdown(f"<div class='q-card'><h3>{ico} {title}</h3><p class='q-small'>{'Open the dedicated workspace.' if lang=='en' else 'Ouvrir l’espace dédié.'}</p></div>",unsafe_allow_html=True)
            if st.button(title,key=f"home_{key}",use_container_width=True): st.session_state.page=key; st.rerun()

# -------------------------------
# Shared project estimation form
# -------------------------------
def quick_form():
    b1,b2,b3=st.columns([1,1,5])
    with b1:
        if st.button("← Accueil" if lang=="fr" else "← Home",key="quick_back_home"): st.session_state.page="home"; st.rerun()
    with b2:
        if st.button("⚙️ Estimation technique" if lang=="fr" else "⚙️ Professional",key="quick_to_pro"): st.session_state.page="professional"; st.rerun()
    st.subheader(f"⚡ {tr('quick',lang)}")
    st.caption(tr('quick_hint',lang))
    installation_enabled=st.radio(
        "Activer l'installation de chantier ?" if lang=="fr" else "Include site installation?",
        ["Oui","Non"], index=1, horizontal=True, key="q_installation_enabled"
    )
    installation_surface=st.number_input(
        "Superficie de l'emprise au sol pour l'installation de chantier (m²)" if lang=="fr" else "Site footprint for site installation (m²)",
        min_value=0.0, value=float(st.session_state.get("q_installation_surface", 0.0) or 0.0), step=1.0, key="q_installation_surface",
        disabled=(installation_enabled!="Oui")
    )
    installation_cost=calculate_installation_chantier_cost(installation_surface) if installation_enabled=="Oui" and installation_surface>0 else 0.0
    if installation_enabled=="Oui" and installation_surface>0:
        st.metric("Montant Lot 01 — Installation de chantier",f"{installation_cost:,.0f} FCFA")
        st.caption("Montant calculé automatiquement selon l'emprise au sol. Il n'est pas saisi manuellement.")
    else:
        st.info("Installation de chantier non retenue : Lot 01 = 0 FCFA.")
    c1,c2,c3=st.columns(3)
    with c1: region=st.selectbox(tr("region",lang),REGIONS,format_func=lambda x:label(x,lang),key="q_region")
    with c2: btype=st.selectbox(tr("building",lang),list(BUILDING_TYPES),format_func=lambda x: x if lang=='fr' else BUILDING_TYPES[x],key="q_btype")
    with c3: area=st.number_input(tr("surface",lang),40.,10000.,150.,10.,key="q_area")
    c1,c2,c3=st.columns(3)
    with c1: n=st.selectbox(tr("levels",lang),list(range(0,11)),format_func=lambda x:f"R+{x}",key="q_levels")
    with c2: standing=st.selectbox(tr("standing",lang),STANDINGS,index=1,format_func=lambda x:label(x,lang),key="q_standing")
    with c3: currency=st.selectbox(tr("currency",lang),list(CURRENCIES),format_func=lambda x: x if lang=='fr' else x,key="q_currency")

    candidates=allowed_walls(n)
    default_wall=PRIMARY_WALL.get(region,candidates[0])
    if default_wall not in candidates: default_wall=candidates[0]
    wall=st.selectbox(
        "Type de construction / mur principal" if lang=="fr" else "Construction type / main wall",
        candidates,
        index=candidates.index(default_wall),
        format_func=lambda x:label(x,lang),
        key="q_wall"
    )

    st.markdown("### 🏗️ Terrassement" if lang=="fr" else "### 🏗️ Earthworks")
    ew1,ew2=st.columns(2)
    with ew1:
        st.caption("Terrassement : Oui / Non" if lang=="fr" else "Earthworks: Yes / No")
        terrassement_choice=st.radio("Terrassement :" if lang=="fr" else "Earthworks:",["Oui","Non"],index=1,horizontal=True,key="q_earthwork")
    with ew2:
        terrassement_pp_price=st.number_input(tr("earthwork_pp_price",lang),min_value=0.0,value=float(TERRASSEMENT_PLAIN_PIED_DEFAULT),step=500.0,key="q_earthwork_pp_price",disabled=(n!=0 or terrassement_choice!="Oui"))

    st.markdown("### 🏠 Toiture" if lang=="fr" else "### 🏠 Roof")
    roof, roof_structure, roof_sheet = roof_selection_ui("q",lang)

    # Le mode rapide reste volontairement simple : pas de détails de fondations ni d'élévations.
    # Les paramètres techniques détaillés sont réservés au mode professionnel.
    foundation=FOUNDATIONS[1]
    installation_cost=0.0
    foundation_wall_length=0.0; foundation_depth=0.80; foundation_wall_type="Parpaing 15 cm"; foundation_wall_unit_price=15000.0
    foundation_voile_thickness=0.0; foundation_voile_volume=0.0; foundation_voile_unit_price=FOUNDATION_VOILE_PRICE_DEFAULT
    foundation_volume_total=0.0; foundation_concrete_unit_price=FOUNDATION_CONCRETE_PRICE_DEFAULT; foundation_wall_area=0.0; foundation_wall_cost=0.0
    elevation_rows=[]

    if st.button(tr("calculate",lang),type="primary",use_container_width=True,key="q_calc"):
        opening="Fenêtre aluminium vitrée"; rdc_type="Dallage sur terre-plein"; slabs=["Dalle à hourdis / corps creux"]*n; ceiling="Staff standard"; soil="Latéritique"
        result=estimate(region,area,n,standing,wall,opening,roof,rdc_type,slabs,ceiling,foundation,soil,"Saison sèche","Accès normal",15,20,0,elevation_rows=None,terrassement=(terrassement_choice=="Oui"),terrassement_plain_pied_price=terrassement_pp_price,foundation_data=None,installation_cost=installation_cost)
        result.update({"project":"QUICK-001","installation_cost_fcfa":installation_cost,"installation_enabled":installation_enabled=="Oui","installation_surface_m2":installation_surface,"region":region,"standing":standing,"soil":soil,"foundation":foundation,"wall":wall,"roof":roof,"roof_structure":roof_structure,"roof_sheet":roof_sheet,"opening":opening,"ceiling":ceiling,"area":area,"n":n,"language":lang,"btype":btype,"foundation_wall_length":0.0,"foundation_depth":0.80,"longrine_length":0.0,"isolated_count":0,"strip_length":0.0,"foundation_wall_area_m2":0.0,"foundation_wall_type":"Parpaing 15 cm","foundation_wall_unit_price_fcfa_m2":15000.0,"foundation_wall_cost_fcfa":0.0,"foundation_voile_thickness_m":0.0,"foundation_voile_volume_m3":0.0,"foundation_voile_unit_price_fcfa_m3":FOUNDATION_VOILE_PRICE_DEFAULT,"foundation_linear_length_m":0.0,"foundation_volume_total_m3":0.0,"foundation_concrete_unit_price_fcfa_m3":FOUNDATION_CONCRETE_PRICE_DEFAULT,"foundation_reference_cost_fcfa":0.0})
        result["input_state"]=capture_quick_input_state(region,btype,area,n,standing,currency,wall,terrassement_choice,terrassement_pp_price,roof_structure,roof_sheet,foundation,0.0,0.80,"Parpaing 15 cm",15000.0,0.0,0.0,FOUNDATION_VOILE_PRICE_DEFAULT,0.0,FOUNDATION_CONCRETE_PRICE_DEFAULT,[])
        result["mode"]="quick"
        if st.session_state.get("terrain") is not None: result["terrain"]=st.session_state.terrain
        st.session_state.estimate_result=result; st.session_state.page="results"; st.rerun()

def professional_form():
    b1,b2,b3=st.columns([1,1,5])
    with b1:
        if st.button("← Accueil" if lang=="fr" else "← Home",key="pro_back_home"): st.session_state.page="home"; st.rerun()
    with b2:
        if st.button("⚡ Estimation rapide" if lang=="fr" else "⚡ Quick estimate",key="pro_to_quick"): st.session_state.page="quick"; st.rerun()
    st.subheader(f"⚙️ {tr('professional',lang)}")
    st.info(
        "Ordre de construction affiché : 01 Installation de chantier → Terrain & implantation → Sol & environnement → 02 Terrassement → 03 Fondations → 04 Élévations → 12 Planchers, dalles & dallages → 05 Charpente & couverture → 07 Menuiseries → 10 Finitions.",
        icon="🏗️"
    )
    st.caption(tr("pro_hint",lang))

    # Ordre visuel des rubriques selon la séquence de construction.
    # Les numéros de lots restent inchangés dans les calculs/DQE.
    ph_project = st.empty()
    ph_install = st.empty()
    ph_terrain = st.empty()
    ph_soil = st.empty()
    ph_earth = st.empty()
    ph_found = st.empty()
    ph_elev = st.empty()
    ph_slab = st.empty()
    ph_roof = st.empty()
    ph_menu = st.empty()
    ph_lot9 = st.empty()
    ph_options = st.empty()
    # Les informations du projet sont séparées des lots. Le Lot 01 est exclusivement
    # réservé à l'installation de chantier et son montant est librement saisi.
    with ph_project.container():
        with st.expander("Informations du projet",expanded=True):
            project=st.text_input("Project name / Nom du projet",value="PROJET-001",key="p_project")
            region=st.selectbox(tr("region",lang),REGIONS,format_func=lambda x:label(x,lang),key="p_region")
            btype=st.selectbox(tr("building",lang),list(BUILDING_TYPES),format_func=lambda x:x if lang=='fr' else BUILDING_TYPES[x],key="p_btype")
    with ph_install.container():
        with st.expander("1. Installation de chantier",expanded=True):
            installation_enabled=st.radio(
                "Activer l'installation de chantier ?" if lang=="fr" else "Include site installation?",
                ["Oui","Non"], index=1, horizontal=True, key="p_installation_enabled"
            )
            installation_surface=st.number_input(
                "Superficie de l'emprise au sol pour l'installation de chantier (m²)" if lang=="fr" else "Site footprint for site installation (m²)",
                min_value=0.0, value=float(st.session_state.get("p_installation_surface", 0.0) or 0.0), step=1.0, key="p_installation_surface",
                disabled=(installation_enabled!="Oui")
            )
            installation_cost=calculate_installation_chantier_cost(installation_surface) if installation_enabled=="Oui" and installation_surface>0 else 0.0
            if installation_enabled=="Oui" and installation_surface>0:
                st.metric("Montant Lot 01 — Installation de chantier",f"{installation_cost:,.0f} FCFA")
                st.caption("Montant calculé automatiquement selon l'emprise au sol. Il n'est pas saisi manuellement.")
            else:
                st.info("Installation de chantier non retenue : Lot 01 = 0 FCFA.")
            st.caption("Barème : ≤50 m² = 500 000 · 51–100 = 650 000 · 101–150 = 800 000 · 151–200 = 1 000 000 · 201–300 = 1 250 000 · 301–400 = 1 500 000 · 401–500 = 1 800 000 · 501–750 = 2 200 000 · 751–1 000 = 2 700 000 · >1 000 = 2 700 000 + 2 500 FCFA/m² supplémentaire.")
    with ph_terrain.container():
        with st.expander("Terrain & implantation",expanded=False):
            st.write("GPS, satellite, 3D and retaining-wall screening are available in the dedicated Terrain workspace." if lang=='en' else "GPS, satellite, 3D et pré-diagnostic du mur de soutènement sont disponibles dans l’espace Terrain.")
            if st.button(tr("terrain",lang),key="go_terrain"): st.session_state.page="terrain"; st.rerun()
    with ph_earth.container():
        with st.expander("2. Terrassement & paramètres du bâtiment",expanded=True):
            c1,c2=st.columns(2)
            with c1: n=st.selectbox(tr("levels",lang),list(range(0,101)),format_func=lambda x:f"R+{x}",key="p_n")
            with c2: area=st.number_input(tr("surface",lang),40.,10000.,180.,10.,key="p_area")
            standing=st.selectbox(tr("standing",lang),STANDINGS,index=1,format_func=lambda x:label(x,lang),key="p_standing")
            candidates=allowed_walls(n)
            default_wall=PRIMARY_WALL.get(region,candidates[0])
            if default_wall not in candidates: default_wall=candidates[0]
            wall=st.selectbox(
                "Type de construction / mur principal" if lang=="fr" else "Construction type / main wall",
                candidates,
                index=candidates.index(default_wall),
                format_func=lambda x:label(x,lang),
                key="p_wall"
            )
            surfaces=[area]+[st.number_input(f"R+{i} — surface (m²)",1.,10000.,round(area*.95,1),5.,key=f"surf_{i}") for i in range(1,n+1)]
            st.markdown("### 🏗️ Terrassement" if lang=="fr" else "### 🏗️ Earthworks")
            ew1,ew2=st.columns(2)
            with ew1:
                st.caption("Terrassement : Oui / Non" if lang=="fr" else "Earthworks: Yes / No")
                terrassement_choice=st.radio("Terrassement :" if lang=="fr" else "Earthworks:",["Oui","Non"],index=1,horizontal=True,key="p_earthwork")
            st.markdown(("**Terrassement sélectionné :** " + terrassement_choice) if lang=="fr" else ("**Selected earthworks:** " + terrassement_choice))
            with ew2:
                st.caption("Prix plain-pied" if lang=="fr" else "Plain-pied rate")
                terrassement_pp_price=st.number_input(tr("earthwork_pp_price",lang),min_value=0.0,value=float(TERRASSEMENT_PLAIN_PIED_DEFAULT),step=500.0,key="p_earthwork_pp_price",disabled=(n!=0 or terrassement_choice!="Oui"))
            if terrassement_choice=="Non": st.info("Le lot 02 — Terrassement sera exclu du calcul (0 FCFA)." if lang=="fr" else "Lot 02 — Earthworks will be excluded from the estimate (0 XAF).")
            elif n==0: st.info(f"Plain-pied : surface × {terrassement_pp_price:,.0f} FCFA/m². Prix modifiable." if lang=="fr" else f"Plain-pied: area × {terrassement_pp_price:,.0f} XAF/m². Editable rate.")
        # Les murs & façades ne constituent plus une section séparée : ces informations
        # seraient répétées avec les élévations. Elles restent toutefois disponibles
        # en interne pour alimenter la prédiction du Lot 4.
    glass=20

    # ------------------------------------------------------------
    # Élévations — murs + structure verticale intégrés dans la même rubrique
    # ------------------------------------------------------------
    elevation_rows = []
    vertical_structure_data = []
    with ph_elev.container():
        with st.expander("4. Élévations — murs & structure verticale", expanded=False):
            st.caption("La structure verticale fait partie intégrante de l'élévation : murs, poteaux/amorce de poteaux et linteaux sont saisis niveau par niveau.")
            for i in range(n+1):
                level = "RDC" if i == 0 else f"R+{i}"
                level_surface = float(surfaces[i])
                st.markdown(f"### {level} — Surface : {level_surface:,.2f} m²")
                e1,e2,e3=st.columns(3)
                with e1:
                    length_m=st.number_input(f"{level} — Longueur développée des murs (m)",min_value=0.0,value=float(4*math.sqrt(max(level_surface,1))),step=0.5,key=f"p_el_length_{i}")
                with e2:
                    height_m=st.number_input(f"{level} — Hauteur des murs (m)",min_value=0.0,value=3.0,step=0.1,key=f"p_el_height_{i}")
                with e3:
                    level_wall=st.selectbox(f"{level} — Type de mur",allowed_walls(n),index=allowed_walls(n).index(wall) if wall in allowed_walls(n) else 0,format_func=lambda x:label(x,lang),key=f"p_el_wall_{i}")
                group_area=max(length_m*height_m,0.0)
                st.metric(f"{level} — Surface totale des murs",f"{group_area:,.2f} m²")
                elevation_rows.append({"level":level,"designation":f"Murs / élévation {level}","wall_type":level_wall,"thickness":WALL_THICKNESS.get(level_wall,"—"),"length_m":length_m,"height_m":height_m,"openings_m2":0.0,"unit_price":float(WALL_PRICE.get(level_wall,10000))})

                st.markdown(f"**{level} — Structure verticale**")
                v1,v2,v3,v4=st.columns(4)
                with v1:
                    poteau_volume=st.number_input(f"{level} — Volume béton poteaux / amorces (m³)",min_value=0.0,value=0.0,step=0.1,key=f"p_poteau_vol_{i}")
                with v2:
                    poteau_price=st.number_input(f"{level} — Prix béton poteaux / amorces (FCFA/m³)",min_value=0.0,value=float(FOUNDATION_CONCRETE_PRICE_DEFAULT),step=1000.0,key=f"p_poteau_price_{i}")
                with v3:
                    linteau_volume=st.number_input(f"{level} — Volume béton linteaux (m³)",min_value=0.0,value=0.0,step=0.1,key=f"p_linteau_vol_{i}")
                with v4:
                    linteau_price=st.number_input(f"{level} — Prix béton linteaux (FCFA/m³)",min_value=0.0,value=float(FOUNDATION_CONCRETE_PRICE_DEFAULT),step=1000.0,key=f"p_linteau_price_{i}")
                vertical_structure_data.append({"level":level,"poteau_volume_m3":poteau_volume,"poteau_price_fcfa_m3":poteau_price,"linteau_volume_m3":linteau_volume,"linteau_price_fcfa_m3":linteau_price})

            elevation_total_area=sum(float(r.get("length_m",0) or 0)*float(r.get("height_m",0) or 0) for r in elevation_rows)
            st.markdown(f"<div class='q-card' style='padding:12px;margin-top:10px'><b>Superficie totale des murs</b><div style='font-size:24px;font-weight:800;color:#0f2742'>{elevation_total_area:,.2f} m²</div></div>",unsafe_allow_html=True)

        # Technical defaults are defined before the final sections.
    opening="Fenêtre aluminium vitrée"; roof="Charpente métallique + bac acier"; rdc_type="Dallage sur terre-plein"; slabs=["Dalle à hourdis / corps creux"]*n; ceiling="Staff standard"; foundation="Semelles isolées + longrines"; soil="Latéritique"
    st.markdown(
        "<div class='q-field-help'>💡 Les intitulés en gras indiquent toujours ce que vous devez saisir ou sélectionner.</div>",
        unsafe_allow_html=True
    )
    with ph_menu.container():
        with st.expander("7. Menuiseries & ouvertures",expanded=False):
            opening=st.selectbox(tr("opening",lang),OPENINGS,format_func=lambda x:label(x,lang),key="p_open")
    with ph_roof.container():
        with st.expander("5. Charpente & couverture",expanded=False):
            roof, roof_structure, roof_sheet = roof_selection_ui("p",lang)
            if lang=="fr":
                st.caption(f"Choix enregistré : {roof}")
            else:
                st.caption(f"Selected: {roof}")
    with ph_slab.container():
        with st.expander("12. Planchers, dalles & dallages",expanded=False):
            rdc_type=st.selectbox("RDC",RDC_TYPES,format_func=lambda x:label(x,lang),key="p_rdc")
            slabs=[st.selectbox(f"R+{i}",SLABS,format_func=lambda x:label(x,lang),key=f"slab_{i}") for i in range(1,n+1)]
        # Lot 9 — finitions : calculs automatiques + détail optionnel.
    total_ceiling_surface=float(sum(surfaces))
    total_wall_surface=float(sum(float(r.get("length_m",0) or 0)*float(r.get("height_m",0) or 0) for r in elevation_rows))
    total_slab_surface=float(sum(surfaces[1:]))
    total_rdc_floor_surface=float(surfaces[0])
    total_tiling_base_surface=float(total_slab_surface + total_rdc_floor_surface)
    with ph_lot9.container():
        with st.expander("10. Carrelage, faux plafond & peinture",expanded=False):
            st.markdown("### A. Faux plafond")
            ceiling_include=st.radio("Voulez-vous intégrer le faux plafond ?",["Oui","Non"],index=0,key="p_ceil_include",horizontal=True)
            c1,c2=st.columns(2)
            with c1:
                ceiling=st.selectbox("Type de plafond",CEILINGS,format_func=lambda x:label(x,lang),key="p_ceil")
            with c2:
                ceiling_price=st.number_input("Prix plafond (FCFA/m²)",min_value=0.0,value=float(CEIL_PRICE.get(ceiling,0)),step=500.0,key="p_ceil_price")
            if ceiling_include=="Non" or ceiling=="Sans plafond":
                ceiling_price=0.0
                ceiling_amount=0.0
            else:
                ceiling_amount=total_ceiling_surface*ceiling_price
            st.metric("Surface plafond totale",f"{total_ceiling_surface:,.2f} m²")
            st.metric("Montant plafond",f"{ceiling_amount:,.0f} FCFA")

            st.markdown("### B. Peinture")
            paint_detail=st.radio("Voulez-vous détailler la peinture ?",["Non","Oui"],index=0,key="p_paint_detail",horizontal=True)
            if paint_detail=="Oui":
                p1,p2=st.columns(2)
                with p1:
                    paint_int_include=st.radio("Intégrer la peinture intérieure ?",["Oui","Non"],index=0,key="p_paint_int_include",horizontal=True)
                    paint_int_surface=st.number_input("Surface peinture intérieure (m²)",min_value=0.0,value=float(total_wall_surface),step=1.0,key="p_paint_int_surface")
                    paint_int_price=st.number_input("Prix peinture intérieure (FCFA/m²)",min_value=0.0,value=5000.0,step=250.0,key="p_paint_int_price")
                with p2:
                    paint_ext_include=st.radio("Intégrer la peinture extérieure ?",["Oui","Non"],index=0,key="p_paint_ext_include",horizontal=True)
                    paint_ext_surface=st.number_input("Surface peinture extérieure (m²)",min_value=0.0,value=float(total_wall_surface),step=1.0,key="p_paint_ext_surface")
                    paint_ext_price=st.number_input("Prix peinture extérieure (FCFA/m²)",min_value=0.0,value=6000.0,step=250.0,key="p_paint_ext_price")
                paint_data={
                    "detail":True,
                    "interior":{"include":paint_int_include=="Oui","surface_m2":paint_int_surface,"price_fcfa_m2":paint_int_price},
                    "exterior":{"include":paint_ext_include=="Oui","surface_m2":paint_ext_surface,"price_fcfa_m2":paint_ext_price},
                }
                paint_total=(paint_int_surface*paint_int_price if paint_int_include=="Oui" else 0)+(paint_ext_surface*paint_ext_price if paint_ext_include=="Oui" else 0)
            else:
                paint_include=st.radio("Intégrer la peinture ?",["Oui","Non"],index=0,key="p_paint_include",horizontal=True)
                paint_surface=total_wall_surface
                paint_price=st.number_input("Prix peinture (FCFA/m²)",min_value=0.0,value=5500.0,step=250.0,key="p_paint_price")
                paint_data={"detail":False,"include":paint_include=="Oui","surface_m2":paint_surface,"price_fcfa_m2":paint_price}
                paint_total=paint_surface*paint_price if paint_include=="Oui" else 0.0
                st.metric("Surface totale des murs prise en compte",f"{paint_surface:,.2f} m²")
            st.metric("Montant peinture",f"{paint_total:,.0f} FCFA")

            st.markdown("### C. Carrelage")
            st.caption("Base automatique : surface du dallage au sol du RDC + somme des surfaces de toutes les dalles/planchers des niveaux supérieurs.")
            tile_include=st.radio("Intégrer le carrelage ?",["Oui","Non"],index=0,key="p_tile_include",horizontal=True)
            tile_price=st.number_input("Prix carrelage (FCFA/m²)",min_value=0.0,value=10000.0,step=500.0,key="p_tile_price")
            tile_surface=total_tiling_base_surface
            tile_total=tile_surface*tile_price if tile_include=="Oui" else 0.0
            st.metric("Surface totale carrelage",f"{tile_surface:,.2f} m²")
            st.metric("Montant carrelage",f"{tile_total:,.0f} FCFA")
            lot9_data={
                "ceiling_include": ceiling_include=="Oui",
                "ceiling_type": ceiling,
                "ceiling_surface_m2": total_ceiling_surface,
                "ceiling_price_fcfa_m2": ceiling_price,
                "paint_detail": paint_data,
                "paint_interior": paint_data.get("interior",{}) if paint_data.get("detail") else {"include":paint_data.get("include",False),"surface_m2":paint_surface,"price_fcfa_m2":paint_price},
                "paint_exterior": paint_data.get("exterior",{}) if paint_data.get("detail") else {"include":False,"surface_m2":0.0,"price_fcfa_m2":0.0},
                "tile_include": tile_include=="Oui",
                "tile_surface_m2": tile_surface,
                "tile_price_fcfa_m2": tile_price,
                "tile_interior":{"include":tile_include=="Oui","surface_m2":tile_surface,"price_fcfa_m2":tile_price},
                "tile_exterior":{"include":False,"surface_m2":0.0,"price_fcfa_m2":0.0},
            }
            lot9_preview=ceiling_amount+paint_total+tile_total
            st.metric("Total détaillé Lot 9",f"{lot9_preview:,.0f} FCFA")
    with ph_found.container():
        with st.expander("3. Fondations",expanded=False):
            foundation=st.selectbox(tr("foundation",lang),FOUNDATIONS,format_func=lambda x:label(x,lang),key="p_found")
            st.markdown("#### 📐 Dimensions de la fondation" if lang=="fr" else "#### 📐 Foundation dimensions")
            f1,f2=st.columns(2)
            with f1:
                foundation_wall_length=st.number_input("Longueur développée du mur de fondation (m)" if lang=="fr" else "Developed foundation-wall length (m)",min_value=0.0,value=float(max(4*math.sqrt(area),0)),step=0.5,key="p_found_wall_len")
            with f2:
                foundation_depth=st.number_input("Profondeur du mur de fondation (m)" if lang=="fr" else "Foundation-wall depth (m)",min_value=0.0,value=0.80,step=0.05,key="p_found_depth")

            foundation_wall_area=foundation_wall_length*foundation_depth
            st.markdown("#### 🧱 Mur de fondation" if lang=="fr" else "#### 🧱 Foundation wall")
            w1,w2=st.columns(2)
            with w1:
                foundation_wall_type=st.selectbox("Type de mur de fondation" if lang=="fr" else "Foundation wall type",FOUNDATION_WALL_TYPES,key="p_found_wall_type")
            with w2:
                if foundation_wall_type=="Voile béton armé":
                    foundation_wall_unit_price=0.0
                    foundation_voile_thickness=st.number_input("Épaisseur du voile (m)" if lang=="fr" else "Shear wall thickness (m)",min_value=0.10,max_value=1.00,value=0.20,step=0.05,key="p_found_voile_thickness")
                else:
                    foundation_voile_thickness=0.0
                    foundation_wall_unit_price=st.number_input("Prix du mur (FCFA/m²)" if lang=="fr" else "Wall price (FCFA/m²)",min_value=float(FOUNDATION_WALL_PRICE[foundation_wall_type]),value=float(FOUNDATION_WALL_PRICE[foundation_wall_type]),step=500.0,key="p_found_wall_price")

            if foundation_wall_type=="Voile béton armé":
                vwall1,vwall2=st.columns(2)
                with vwall1:
                    foundation_voile_unit_price=st.number_input("Prix du voile (FCFA/m³)" if lang=="fr" else "Shear-wall price (FCFA/m³)",min_value=float(FOUNDATION_VOILE_PRICE_MIN),max_value=float(FOUNDATION_VOILE_PRICE_MAX),value=float(FOUNDATION_VOILE_PRICE_DEFAULT),step=1000.0,key="p_found_voile_price")
                with vwall2:
                    foundation_voile_volume=foundation_wall_area*foundation_voile_thickness
                    st.metric("Volume des voiles" if lang=="fr" else "Shear-wall volume",f"{foundation_voile_volume:,.2f} m³")
                foundation_wall_cost=foundation_voile_volume*foundation_voile_unit_price
            else:
                foundation_voile_unit_price=FOUNDATION_VOILE_PRICE_DEFAULT
                foundation_voile_volume=0.0
                foundation_wall_cost=foundation_wall_area*foundation_wall_unit_price

            st.metric("Surface calculée des murs de fondation (m²)" if lang=="fr" else "Calculated foundation-wall area (m²)",f"{foundation_wall_area:,.2f} m²")
            st.metric("Coût calculé des murs de fondation (FCFA)" if lang=="fr" else "Calculated foundation-wall cost (FCFA)",f"{foundation_wall_cost:,.0f} FCFA")

            # Champs retirés volontairement : nombre de semelles isolées,
            # longueur des semelles filantes et longueur développée des longrines.
            # Les volumes restent saisis directement ci-dessous.
            longrine_length=0.0
            isolated_count=0
            strip_length=0.0

            st.markdown("#### 🧱 Niveau de détail — volumes de béton" if lang=="fr" else "#### 🧱 Detail level — concrete volumes")
            v1,v2=st.columns(2)
            with v1:
                foundation_semelles_volume=st.number_input(
                    "Volume de béton des semelles (m³)" if lang=="fr" else "Concrete volume for footings (m³)",
                    min_value=0.0,value=0.0,step=0.10,key="p_foundation_semelles_volume"
                )
            with v2:
                foundation_longrines_volume=st.number_input(
                    "Volume de béton des longrines (m³)" if lang=="fr" else "Concrete volume for tie beams (m³)",
                    min_value=0.0,value=0.0,step=0.10,key="p_foundation_longrines_volume"
                )
            v3,v4=st.columns(2)
            with v3:
                foundation_amorces_volume=st.number_input(
                    "Volume de béton des amorces de poteaux (m³)" if lang=="fr" else "Concrete volume for column starters (m³)",
                    min_value=0.0,value=0.0,step=0.10,key="p_foundation_amorces_volume"
                )
            foundation_volume_total=foundation_semelles_volume+foundation_longrines_volume+foundation_amorces_volume

            # Compatibilité interne avec les anciennes variables.
            radier_thickness=0.0
            radier_volume=foundation_semelles_volume if foundation in ("Radier général","Radier nervuré") else 0.0
            isolated_volume=foundation_semelles_volume if foundation in ("Semelles isolées","Semelles isolées + longrines") else 0.0
            strip_volume=foundation_semelles_volume if foundation=="Semelles filantes" else 0.0
            longrine_volume=foundation_longrines_volume
            poteau_volume=foundation_amorces_volume

            foundation_concrete_unit_price=st.number_input("Prix de référence du béton (FCFA/m³)" if lang=="fr" else "Reference concrete price (FCFA/m³)",min_value=float(FOUNDATION_CONCRETE_PRICE_MIN),max_value=float(FOUNDATION_CONCRETE_PRICE_MAX),value=float(FOUNDATION_CONCRETE_PRICE_DEFAULT),step=1000.0,key="p_foundation_m3_price")
            foundation_reference_cost=foundation_volume_total*foundation_concrete_unit_price+foundation_wall_cost
            if foundation_wall_type=="Voile béton armé":
                foundation_reference_cost += foundation_voile_volume*foundation_voile_unit_price
            foundation_linear_length=0.0
            st.metric("Volume total de béton calculé" if lang=="fr" else "Calculated total concrete volume",f"{foundation_volume_total:,.2f} m³")
            st.metric("Référence financière des fondations (FCFA)" if lang=="fr" else "Foundation financial reference (FCFA)",f"{foundation_reference_cost:,.0f} FCFA")
            st.caption((f"Prix béton commun : {foundation_concrete_unit_price:,.0f} FCFA/m³ · parpaing 15 cm : 15 000 FCFA/m² · parpaing 20 cm : 18 000 FCFA/m² · voile BA : 210 000 FCFA/m³") if lang=="fr" else (f"Common concrete rate: {foundation_concrete_unit_price:,.0f} FCFA/m³ · 15 cm block: 15,000 FCFA/m² · 20 cm block: 18,000 FCFA/m² · reinforced-concrete wall: 210,000 FCFA/m³"))
    with ph_soil.container():
        with st.expander("Sol & environnement",expanded=False):
            soil=st.selectbox(tr("soil",lang),SOILS,format_func=lambda x:label(x,lang),key="p_soil")
    with ph_options.container():
        with st.expander("Options avancées",expanded=False):
            season=st.selectbox(tr("season",lang),["Saison sèche","Saison des pluies"],format_func=lambda x:label(x,lang),key="p_season")
            access=st.selectbox(tr("access",lang),["Accès normal","Accès difficile"],format_func=lambda x:label(x,lang),key="p_access")
            distance=st.number_input(tr("distance",lang),0.,500.,15.,1.,key="p_dist")
            project_id=st.text_input("Project ID",value=project,key="p_id")
    if st.button(tr("calculate",lang),type="primary",use_container_width=True,key="p_calc"):
        retained=0.0
        if st.session_state.get("terrain"):
            retained=sum(r["length_m"] for r in st.session_state.terrain["rows"] if r["retaining"])*55000
        result=estimate(region,area,n,standing,wall,opening,roof,rdc_type,slabs,ceiling,foundation,soil,season,access,distance,glass,retained,elevation_rows=elevation_rows,terrassement=(terrassement_choice=="Oui"),terrassement_plain_pied_price=terrassement_pp_price,foundation_data={"total_concrete_volume_m3":foundation_volume_total,"semelles_concrete_volume_m3":foundation_semelles_volume,"longrines_concrete_volume_m3":foundation_longrines_volume,"amorces_poteaux_concrete_volume_m3":foundation_amorces_volume,"concrete_unit_price_fcfa_m3":foundation_concrete_unit_price,"foundation_wall_area_m2":foundation_wall_area,"foundation_wall_type":foundation_wall_type,"foundation_wall_unit_price_fcfa_m2":foundation_wall_unit_price,"foundation_wall_cost_fcfa":foundation_wall_cost,"foundation_voile_volume_m3":foundation_voile_volume,"foundation_voile_unit_price_fcfa_m3":foundation_voile_unit_price},vertical_structure_data=vertical_structure_data,lot9_data=lot9_data,installation_cost=installation_cost)
        result.update({"project":project_id,"installation_cost_fcfa":installation_cost,"installation_enabled":installation_enabled=="Oui","installation_surface_m2":installation_surface,"region":region,"standing":standing,"soil":soil,"foundation":foundation,"wall":wall,"roof":roof,"roof_structure":roof_structure,"roof_sheet":roof_sheet,"opening":opening,"ceiling":ceiling,"area":area,"n":n,"language":lang,"btype":btype,"glass":glass,"season":season,"access":access,"distance":distance,"foundation_wall_length":foundation_wall_length,"foundation_depth":foundation_depth,"longrine_length":0.0,"isolated_count":0,"strip_length":0.0,"foundation_wall_area_m2":foundation_wall_area,"foundation_wall_type":foundation_wall_type,"foundation_wall_unit_price_fcfa_m2":foundation_wall_unit_price,"foundation_wall_cost_fcfa":foundation_wall_cost,"foundation_voile_thickness_m":foundation_voile_thickness,"foundation_voile_volume_m3":foundation_voile_volume,"foundation_voile_unit_price_fcfa_m3":foundation_voile_unit_price,"foundation_linear_length_m":foundation_linear_length,"radier_thickness_m":radier_thickness,"radier_volume_m3":radier_volume,"isolated_volume_m3":isolated_volume,"strip_volume_m3":strip_volume,"longrine_volume_m3":longrine_volume,"poteau_volume_m3":poteau_volume,"foundation_volume_total_m3":foundation_volume_total,"foundation_semelles_volume_m3":foundation_semelles_volume,"foundation_longrines_volume_m3":foundation_longrines_volume,"foundation_amorces_volume_m3":foundation_amorces_volume,"foundation_concrete_unit_price_fcfa_m3":foundation_concrete_unit_price,"foundation_reference_cost_fcfa":foundation_reference_cost,"vertical_structure_data":vertical_structure_data,"lot9_data":lot9_data,"lot9_detail_total_fcfa":lot9_preview})
        result["input_state"]=capture_professional_input_state(project_id,region,btype,n,area,standing,wall,terrassement_choice,terrassement_pp_price,roof_structure,roof_sheet,opening,rdc_type,ceiling,foundation,soil,season,access,distance,foundation_wall_length,foundation_depth,foundation_wall_type,foundation_wall_unit_price,foundation_voile_thickness,foundation_voile_volume,foundation_voile_unit_price,foundation_volume_total,foundation_concrete_unit_price,elevation_rows,vertical_structure_data,lot9_data,foundation_semelles_volume,foundation_longrines_volume,foundation_amorces_volume)
        result["mode"]="professional"
        if st.session_state.get("terrain") is not None: result["terrain"]=st.session_state.terrain
        st.session_state.estimate_result=result; st.session_state.page="results"; st.rerun()

# -------------------------------
# Pages
# -------------------------------
if st.session_state.page=="home": home_page()
elif st.session_state.page=="quick": quick_form()
elif st.session_state.page=="professional": professional_form()
elif st.session_state.page=="projects":
    st.subheader("📁 "+tr("projects",lang))
    st.markdown("### 📥 Importer / reprendre un projet" if lang=="fr" else "### 📥 Import / resume a project")
    uploaded_project=st.file_uploader("Insérer un fichier Quotora (.json) pour récupérer toutes les données et les modifier" if lang=="fr" else "Upload a Quotora (.json) file to restore all data and edit it",type=["json"],key="project_import_json")
    if uploaded_project is not None:
        try:
            imported=json.loads(uploaded_project.getvalue().decode("utf-8"))
            if not isinstance(imported,dict): raise ValueError("Format JSON invalide.")
            imported_state=imported.get("input_state", {})
            imported_mode=imported_state.get("mode", imported.get("mode", "professional")) if isinstance(imported_state,dict) else imported.get("mode", "professional")
            imported_project=imported.get("project", imported.get("project_id", "Projet importé"))
            st.success((f"✅ Fichier valide — Projet : **{imported_project}** · Mode : **{imported_mode}**. Les données sont prêtes à être chargées et modifiées." if lang=="fr" else f"✅ Valid file — Project: **{imported_project}** · Mode: **{imported_mode}**. Data are ready to load and edit."))
            if st.button("↩️ Charger les données dans le formulaire et modifier" if lang=="fr" else "↩️ Load data into the form and edit",type="primary",key="load_imported_project"):
                restore_saved_state(imported); st.session_state.page="professional" if st.session_state.get("loaded_mode")=="professional" else "quick"; st.rerun()
        except Exception as e: st.error(f"Impossible de lire le fichier JSON : {e}")
    st.divider()
    with sqlite3.connect(DB_PATH) as con: projects=pd.read_sql_query("SELECT project_id AS ID, project_name AS Projet, created_at AS Créé FROM projects ORDER BY created_at DESC",con)
    if projects.empty: st.info("No saved projects yet." if lang=="en" else "Aucun projet enregistré.")
    else:
        st.dataframe(projects,use_container_width=True,hide_index=True); selected=st.selectbox("Project",projects["ID"].tolist()); b1,b2=st.columns(2)
        with b1:
            if st.button("Ouvrir / modifier" if lang=="fr" else "Open / edit",type="primary",use_container_width=True,key="open_edit_project"):
                with sqlite3.connect(DB_PATH) as con: row=con.execute("SELECT payload FROM projects WHERE project_id=?",(selected,)).fetchone()
                if row:
                    loaded=json.loads(row[0]); restore_saved_state(loaded); st.session_state.page="professional" if st.session_state.get("loaded_mode")=="professional" else "quick"; st.rerun()
        with b2:
            if st.button("Voir les résultats" if lang=="fr" else "View results",use_container_width=True,key="view_saved_project"):
                with sqlite3.connect(DB_PATH) as con: row=con.execute("SELECT payload FROM projects WHERE project_id=?",(selected,)).fetchone()
                if row: st.session_state.estimate_result=json.loads(row[0]); st.session_state.page="results"; st.rerun()

elif st.session_state.page=="terrain":
    # Navigation shortcut: return directly to the professional technical estimate
    back1, back2 = st.columns([1, 5])
    with back1:
        if st.button("← Estimation technique" if lang=="fr" else "← Professional estimate", key="terrain_to_professional", use_container_width=True):
            st.session_state.page="professional"
            st.rerun()
    st.subheader(f"🗺️ {tr('terrain',lang)}")
    st.caption("Use GPS points, draw the parcel, or import a GIS file." if lang=="en" else "Saisissez les points GPS, dessinez la parcelle ou importez un fichier GIS.")
    tab1,tab2=st.tabs([tr("coordinates",lang),tr("map",lang)])
    with tab1:
        txt=st.text_area("Latitude, longitude",height=130,placeholder="3.848000, 11.502000\n3.848120, 11.502150\n3.847950, 11.502320",key="gps_text")
        uploaded=st.file_uploader(tr("import",lang),type=["geojson","json","kml","gpx","csv"],key="gis_file")
        dem_file=st.file_uploader("MNT GeoTIFF (optional)" if lang=="en" else "MNT GeoTIFF (optionnel)",type=["tif","tiff"],key="dem_file")
        if uploaded:
            try: st.session_state.terrain_points=parse_uploaded_geometry(uploaded); st.success(f"{len(st.session_state.terrain_points)} points imported.")
            except Exception as e: st.error(str(e))
        if st.button("Analyser le terrain" if lang=="fr" else "Analyze site",type="primary"):
            pts=parse_coords(txt) if txt.strip() else st.session_state.get("terrain_points",[])
            if len(pts)>=3:
                try: st.session_state.terrain={"points":pts,**terrain_analysis(pts, dem_file.getvalue() if dem_file else None)}; st.success("Terrain analysed." if lang=="en" else "Terrain analysé.")
                except Exception as e: st.error(str(e))
            else: st.error("At least 3 GPS points are required." if lang=="en" else "Au moins 3 points GPS sont nécessaires.")
    with tab2:
        if st_folium and folium:
            dm=folium.Map(location=[3.85,11.50],zoom_start=6,tiles="OpenStreetMap")
            Draw(export=True,draw_options={"polyline":False,"rectangle":True,"circle":False,"circlemarker":False,"marker":False,"polygon":True},edit_options={"edit":True,"remove":True}).add_to(dm)
            drawn=st_folium(dm,height=450,width=1400,key="qdraw",returned_objects=["all_drawings"])
            drawings=(drawn or {}).get("all_drawings") or []
            if drawings:
                g=drawings[-1].get("geometry",{})
                if g.get("type")=="Polygon":
                    pts=[(float(c[1]),float(c[0])) for c in g.get("coordinates", [[]])[0][:-1]]
                    if len(pts)>=3 and st.button("Use drawn parcel",key="use_drawn"):
                        st.session_state.terrain={"points":pts,**terrain_analysis(pts, dem_file.getvalue() if dem_file else None)}; st.rerun()
        else: st.warning("Install streamlit-folium to use interactive maps.")
    if st.session_state.get("terrain"):
        a=st.session_state.terrain
        st.divider(); m1,m2,m3=st.columns(3); m1.metric("Area",f"{a['area_m2']:,.0f} m²"); m2.metric("Relief",f"{a['relief']:.2f} m"); proposed=[r for r in a['rows'] if r['retaining']]; m3.metric("Retaining wall",f"{sum(r['length_m'] for r in proposed):.1f} ml" if proposed else "—")
        if st_folium and folium: st_folium(render_map(a["points"],a["rows"]),height=480,width=1400,key="result_map")
        st.plotly_chart(terrain_3d(a["points"],a),use_container_width=True)
        if proposed:
            st.success(f"🧱 {tr('recommended',lang)} — {sum(r['length_m'] for r in proposed):.1f} ml")
            df=pd.DataFrame([{ "Côté":r["side"],"Longueur (m)":r["length_m"],"Dénivelée (m)":r["delta_m"],"Azimut (°)":r["bearing"]} for r in proposed])
            st.dataframe(df,use_container_width=True,hide_index=True)
            rw_cost=sum(r["length_m"] for r in proposed)*55000
            st.metric("Pré-chiffrage mur de soutènement",fcfa(rw_cost))
        else: st.info(tr("no_wall",lang))

elif st.session_state.page=="results":
    res=st.session_state.get("estimate_result")
    if not res:
        st.info("No estimate yet." if lang=="en" else "Aucune estimation disponible.")
        if st.button(tr("quick",lang)): st.session_state.page="quick"; st.rerun()
    else:
        st.subheader(f"📊 {tr('results',lang)}")
        rb1,rb2,rb3=st.columns([1,1,4])
        with rb1:
            if st.button("← Modifier" if lang=="fr" else "← Edit",key="results_edit"):
                mode=res.get("mode",(res.get("input_state") or {}).get("mode","professional")); restore_saved_state(res); st.session_state.page="professional" if mode=="professional" else "quick"; st.rerun()
        with rb2:
            if st.button("⚡ Rapide" if lang=="fr" else "⚡ Quick",key="results_quick"): st.session_state.page="quick"; st.rerun()
        s1=st.container(); s1.metric("Terrassement" if lang=="fr" else "Earthworks", res.get("terrassement_choice","Non"))
        a,b,c,d=st.columns(4)
        a.metric(tr("gross",lang),fcfa(res["gross"])); b.metric(tr("finishes",lang),fcfa(res["finishes"])); c.metric(tr("total",lang),fcfa(res["total"])); d.metric(tr("cost_m2",lang),fcfa(res["cost_m2"]))
        st.divider()
        st.subheader("💱 "+tr("currency",lang))
        rates,status,date=rates_xaf(); target=st.selectbox(tr("currency",lang),list(CURRENCIES),format_func=lambda x:x,key="result_currency"); code=CURRENCIES[target]; rate=rates.get(code,FALLBACK_PER_XAF[code])
        cc=st.columns(4); cc[0].metric(tr("gross",lang),money(res["gross"]*rate,code)); cc[1].metric(tr("finishes",lang),money(res["finishes"]*rate,code)); cc[2].metric(tr("total",lang),money(res["total"]*rate,code)); cc[3].metric(tr("cost_m2",lang),money(res["cost_m2"]*rate,code))
        st.caption(f"1 XAF = {rate:.8g} {code} · {status} · {date}")
        st.divider(); st.subheader(tr("costs",lang))
        amounts=res["amounts"]; df=pd.DataFrame({"Lot":LOTS,"Montant FCFA":np.round(amounts).astype(int)}); df["Part (%)"]=df["Montant FCFA"]/res["total"]*100 if res["total"] else 0
        st.dataframe(df.style.format({"Montant FCFA":"{:,.0f}","Part (%)":"{:.1f}"}),use_container_width=True,hide_index=True)
        b1,b2=st.columns(2)
        with b1:
            if st.button("💾 "+tr("save",lang),use_container_width=True):
                save_project_result(res); st.success("Project saved." if lang=="en" else "Projet enregistré.")
        with b2:
            st.download_button("⬇️ JSON — Projet rééditable" if lang=="fr" else "⬇️ JSON — Editable project",json.dumps(json_safe_result(res),ensure_ascii=False,default=str,indent=2).encode(),f"{res['project']}_quotora.json","application/json",use_container_width=True)
        if st.button("📄 "+tr("report",lang),type="primary"):
            st.session_state.page="reports"; st.rerun()

elif st.session_state.page=="reports":
    res=st.session_state.get("estimate_result")
    st.subheader("📄 "+tr("reports",lang))
    if st.button("← Résultats" if lang=="fr" else "← Results",key="report_back_results"): st.session_state.page="results"; st.rerun()
    if not res:
        st.info("Run an estimate first." if lang=="en" else "Lancez d’abord une estimation.")
    else:
        amounts=res["amounts"]; df=pd.DataFrame({"Lot":LOTS,"Montant FCFA":np.round(amounts).astype(int)}); df["Part (%)"]=df["Montant FCFA"]/res["total"]*100 if res["total"] else 0
        html=print_report_html(res["project"],lang,res["region"],res["n"],res["standing"],res["soil"],res["foundation"],res["wall"],res["roof"],sum(res["surfaces"]),res["gross"],res["finishes"],res["total"],res["cost_m2"],df,st.session_state.get("terrain"),res.get("elevation_rows"))
        excel_bytes=detailed_excel_bytes(res,lang)
        summary_bytes=summary_excel_bytes(res,lang)
        detailed_html_bytes=detailed_html(res,lang).encode("utf-8")
        st.markdown("### 📑 Sortie du devis / bordereau" if lang=="fr" else "### 📑 Estimate / bill of quantities output")
        ex1,ex2,ex3=st.columns(3)
        with ex1:
            st.download_button("📊 Excel — devis détaillé" if lang=="fr" else "📊 Excel — detailed BOQ",excel_bytes,f"{res['project']}_devis_detaille.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        with ex2:
            st.download_button("🌐 HTML — devis détaillé" if lang=="fr" else "🌐 HTML — detailed BOQ",detailed_html_bytes,f"{res['project']}_devis_detaille.html","text/html",use_container_width=True)
        with ex3:
            st.download_button("📋 Excel — synthèse 12 lots" if lang=="fr" else "📋 Excel — 12-package summary",summary_bytes,f"{res['project']}_quotora_synthese.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        st.caption("Excel est entièrement modifiable. HTML contient des cellules éditables dans le navigateur. Les quantités non saisies restent à compléter." if lang=="fr" else "Excel is fully editable. HTML contains editable cells in the browser. Quantities not entered remain to be completed.")
        st.components.v1.html("<button onclick='window.parent.print()' style='padding:10px 16px;background:#2563eb;color:#fff;border:0;border-radius:8px;font-weight:700'>🖨 Print / PDF</button>",height=50)
        st.download_button("⬇️ Download printable HTML",html.encode("utf-8"),f"{res['project']}_quotora_report.html","text/html",use_container_width=True)
        st.markdown("The printable report is a **budgetary report, not a contractual quotation**." if lang=="en" else "Le rapport imprimable est un **rapport budgétaire, sans valeur de devis contractuel**.")

elif st.session_state.page=="tracking":
    st.subheader("📈 "+tr("tracking",lang))
    project=st.text_input("Project ID",value=(st.session_state.get("estimate_result") or {}).get("project","PROJET-001"))
    progress=st.slider(tr("progress",lang),0,100,0,1); status=st.selectbox(tr("status",lang),["Non démarré","En préparation","En cours","En pause","Terminé"],format_func=lambda x:label(x,lang)); note=st.text_area(tr("observation",lang)); photos=st.file_uploader(tr("photos",lang),type=["jpg","jpeg","png"],accept_multiple_files=True)
    if st.button(tr("update",lang),type="primary"):
        now=dt.datetime.now().isoformat(timespec="seconds")
        with sqlite3.connect(DB_PATH) as con:
            con.execute("INSERT INTO progress VALUES (?,?,?,?,?)",(project,now,progress,status,note))
            for ph in photos or []: con.execute("INSERT INTO photos VALUES (?,?,?,?)",(project,now,ph.name,ph.getvalue()))
            con.commit()
        st.success("Saved." if lang=="en" else "Mise à jour enregistrée.")
    with sqlite3.connect(DB_PATH) as con:
        hist=pd.read_sql_query("SELECT updated_at AS Date, progress AS Avancement, status AS Statut, note AS Observation FROM progress WHERE project_id=? ORDER BY updated_at",con,params=(project,))
    if not hist.empty:
        st.metric(tr("progress",lang),f"{hist.iloc[-1]['Avancement']:.0f}%")
        fig=go.Figure(go.Scatter(x=hist["Date"],y=hist["Avancement"],mode="lines+markers")); fig.update_layout(height=320,yaxis=dict(range=[0,100],title="%"),xaxis_title="Date"); st.plotly_chart(fig,use_container_width=True)
        st.dataframe(hist,use_container_width=True,hide_index=True)
        st.download_button("⬇️ CSV",hist.to_csv(index=False).encode(),f"{project}_progress.csv","text/csv")
    else: st.info("No history yet." if lang=="en" else "Aucun historique pour ce projet.")

elif st.session_state.page=="settings":
    st.subheader("⚙️ "+tr("settings",lang))
    st.write("### Language / Langue")
    st.write("Use the language selector in the sidebar. All main workflows, labels and report headings adapt to the selected language." if lang=="en" else "Utilisez le sélecteur de langue dans la barre latérale. Les principaux parcours, libellés et titres de rapport s’adaptent à la langue choisie.")
    st.write("### Model information")
    st.info("Le Lot 4 — Élévations reste une prédiction ML. La géométrie des murs et le référentiel de prix sont calculés automatiquement en interne pour alimenter et calibrer la prédiction." if lang=="fr" else "Lot 4 — Elevations is predicted by the ML model from the real wall area and weighted entered unit price. The user reference acts as a budgetary anchor.")
    if MODEL_METRICS:
        total_metric=MODEL_METRICS.get("10_lots_total",{})
        st.dataframe(pd.DataFrame(total_metric,index=["Value"]),use_container_width=True)
    st.write("### Data / technical note")
    st.info("The current ML model was trained on a synthetic/budgetary dataset. It should be recalibrated with real project cost data before contractual deployment." if lang=="en" else "Le modèle ML actuel a été entraîné sur un dataset synthétique/budgétaire. Il doit être recalibré avec des données réelles de chantiers avant un usage contractuel.")
    st.write("### Access control / Contrôle d'accès")
    if st.session_state.get("access_gate_configured"):
        st.success("🔒 " + ("Accès protégé par code Q14 (14 jours) — actif." if lang=="fr" else "Q14 code-gated access (14 days) is active."))
    else:
        st.warning("🔓 " + ("Mode libre — aucun ACCESS_SECRET configuré. Ajoutez-le dans les Secrets Streamlit pour activer l'accès payant." if lang=="fr" else "Open mode — no ACCESS_SECRET configured. Add it in Streamlit Secrets to enable paid access."))

st.markdown(f"<div class='q-footer'>{APP_NAME} · Estimation & Pilotage Construction · {tr('founder',lang)} · Budgetary decision-support tool</div>",unsafe_allow_html=True)
