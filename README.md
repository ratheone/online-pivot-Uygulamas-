# -*- coding: utf-8 -*-
"""
Created on Thu Jan  1 12:53:27 2026
python -m streamlit run pvt_excl.py
@author: Muhammed KOÇ
"""

import streamlit as st
import sqlite3
import hashlib
import json
import time
import pandas as pd
import numpy as np
import io
import os
import base64
from datetime import datetime
import plotly.express as px

# --- TKINTER ENTEGRASYONU (Klasör Seçimi) ---
import tkinter as tk
from tkinter import filedialog

# ==============================================================================
# 🛠️ GEREKLİ KÜTÜPHANELER VE AYARLAR
# ==============================================================================
DB_NAME = "sistem_v8.db"
ANALYSIS_PRESET_FILE = "analysis_presets.json"

SIDEBAR_MENUS = {
    "dinamik_tablo": "📊 Dinamik Tablo (Analiz)",
    "dinamik_yonetim": "📂 Dinamik Tablolar (Dosya Yükle/Sil)"
}

EXCEL_CHART_TYPES = [
    "Sütun (Column)", "Küme Sütun (Clustered Column)", "Yığılmış Sütun (Stacked Column)", 
    "Yüzdelik Yığılmış Sütun (100% Stacked Column)", "3-B Sütun (3-D Column)", 
    "Çubuk (Bar)", "Çizgi (Line)", "Pasta (Pie)", "Halka (Doughnut)", 
    "Isı Haritası (Heatmap)", "Ağaç Haritası (Treemap)", "Sunburst (Güneş Patlaması)", 
    "Huni (Funnel)", "Kutucuk (Box & Whisker)", "Histogram"
]

# ==============================================================================
# 🛠️ YARDIMCI FONKSİYONLAR (TKINTER VE DB)
# ==============================================================================

def select_folder():
    """Yerel klasör seçme penceresini açar."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_selected = filedialog.askdirectory()
    root.destroy()
    return folder_selected

def get_db(): return sqlite3.connect(DB_NAME)

def init_db():
    with get_db() as conn:
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT UNIQUE NOT NULL, phone TEXT UNIQUE NOT NULL, address TEXT NOT NULL, password TEXT NOT NULL, is_admin BOOLEAN DEFAULT 0, access_granted BOOLEAN DEFAULT 0, verification_code TEXT, is_blocked BOOLEAN DEFAULT 0, allowed_menus TEXT DEFAULT '')''')
        c.execute('''CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS excel_tables (id INTEGER PRIMARY KEY AUTOINCREMENT, menu_id TEXT, name TEXT, file_data BLOB, description TEXT, created_at DATETIME DEFAULT CURRENT_TIMESTAMP)''')
        c.execute('''CREATE TABLE IF NOT EXISTS analysis_configs (table_id INTEGER PRIMARY KEY, config_json TEXT)''')
        c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('bg_active', '1')")
        c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('bg_url', 'https://images.unsplash.com/photo-1451187580459-43490279c0fa?q=80&w=2072&auto=format&fit=crop')")
        admin_pass = hashlib.sha256("admin123".encode()).hexdigest()
        c.execute("INSERT OR IGNORE INTO users (email, phone, address, password, is_admin, access_granted, is_blocked, allowed_menus) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                  ("admin@sistem.com", "000", "Merkez Ofis", admin_pass, 1, 1, 0, "all"))
        conn.commit()

# --- VERİ TABANI İŞLEMLERİ ---
def get_excel_tables(menu_id):
    with get_db() as conn:
        return conn.execute("SELECT id, name, created_at, description FROM excel_tables WHERE menu_id=?", (menu_id,)).fetchall()

def get_excel_table_details(tid):
    with get_db() as conn:
        row = conn.execute("SELECT id, name, file_data, description FROM excel_tables WHERE id=?", (tid,)).fetchone()
        if row:
            df = pd.read_excel(io.BytesIO(row[2]))
            return row[0], row[1], df, row[3]
    return None, None, pd.DataFrame(), None

def save_analysis_config(tid, config):
    with get_db() as conn:
        conn.execute("INSERT OR REPLACE INTO analysis_configs (table_id, config_json) VALUES (?, ?)", (tid, json.dumps(config)))
        conn.commit()

def delete_excel_table(tid):
    with get_db() as conn:
        conn.execute("DELETE FROM excel_tables WHERE id=?", (tid,))
        conn.execute("DELETE FROM analysis_configs WHERE table_id=?", (tid,))
        conn.commit()

# ==============================================================================
# 🎨 GÖRSEL TASARIM (CSS) - ÖZGÜN YAPI KORUNDU
# ==============================================================================
def inject_custom_css():
    bg_active = "1"
    bg_url = "https://images.unsplash.com/photo-1451187580459-43490279c0fa?q=80&w=2072&auto=format&fit=crop"
    bg_styles = f'content: ""; position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: linear-gradient(rgba(0,0,0,0.5), rgba(0,0,0,0.5)), url("{bg_url}") no-repeat center center fixed; background-size: cover; filter: blur(8px); z-index: -9999;'
    
    st.markdown(f'''<style>
    .stApp {{ background: transparent !important; }}
    .stApp::before {{ {bg_styles} }}
    .programlar-header-box {{ background: rgba(255, 255, 255, 0.2); border: 1px solid rgba(255, 255, 255, 0.3); border-radius: 30px; padding: 10px 50px; margin: 0 auto 30px auto; width: fit-content; text-align: center; backdrop-filter: blur(10px); }}
    .programlar-header-text {{ color: white; font-size: 38px; font-weight: 900; margin: 0; text-shadow: 2px 2px 4px rgba(0,0,0,0.5); }}
    [data-testid="stSidebar"] {{ background-color: rgba(0,0,0,0.7) !important; }}
    .stButton>button {{ width: 100%; border-radius: 10px; }}
    </style>''', unsafe_allow_html=True)

# ==============================================================================
# 🚀 ANA UYGULAMA
# ==============================================================================
def main():
    st.set_page_config(page_title="Dinamik Analiz Sistemi", layout="wide")
    init_db()
    inject_custom_css()

    if 'logged_in' not in st.session_state: st.session_state.logged_in = False
    if 'current_view' not in st.session_state: st.session_state.current_view = "dashboard"

    # --- SIDEBAR (SADECE İSTEDİĞİNİZ 2 ÖZELLİK) ---
    if st.session_state.logged_in:
        with st.sidebar:
            st.markdown('<div style="text-align:center; padding:10px;"><h2 style="color:white;">🚀 MENÜ</h2></div>', unsafe_allow_html=True)
            st.divider()
            if st.button("🏠 Ana Sayfa"): st.session_state.current_view = "dashboard"; st.rerun()
            for key, label in SIDEBAR_MENUS.items():
                if st.button(label): st.session_state.current_view = key; st.rerun()
            st.divider()
            if st.button("🚪 Çıkış Yap"): st.session_state.logged_in = False; st.rerun()

    # --- GİRİŞ EKRANI ---
    if not st.session_state.logged_in:
        _, center_col, _ = st.columns([1, 2, 1])
        with center_col:
            st.markdown('<div class="programlar-header-box"><h1 class="programlar-header-text">Programlar</h1></div>', unsafe_allow_html=True)
            with st.form("login_form"):
                st.subheader("🔐 Sistem Girişi")
                u_email = st.text_input("Email")
                u_pass = st.text_input("Şifre", type="password")
                if st.form_submit_button("Giriş Yap"):
                    with get_db() as conn:
                        user = conn.execute("SELECT * FROM users WHERE email=?", (u_email,)).fetchone()
                        if user and hashlib.sha256(u_pass.encode()).hexdigest() == user[4]:
                            st.session_state.logged_in = True
                            st.session_state.user_email = u_email
                            st.session_state.is_admin = user[5]
                            st.rerun()
                        else: st.error("Hatalı bilgiler.")

    # --- İÇERİK YÖNETİMİ ---
    else:
        view = st.session_state.current_view

        if view == "dashboard":
            st.markdown('<div class="programlar-header-box"><h1 class="programlar-header-text">Hoş Geldiniz</h1></div>', unsafe_allow_html=True)
            st.info("Lütfen işlem yapmak için soldaki menüyü kullanın.")

        # ==============================================================================
        # 📊 DİNAMİK TABLO (ANALİZ)
        # ==============================================================================
        elif view == "dinamik_tablo":
            st.header("📊 Dinamik Tablo Analiz ve Grafik Motoru")
            tbls = get_excel_tables("genel_analiz")
            
            if not tbls:
                st.warning("Henüz tablo yüklenmemiş. Lütfen 'Dosya Yükle' menüsünü kullanın.")
            else:
                ts = {f"{t[1]} ({t[2]})": t[0] for t in tbls}
                sel_name = st.selectbox("Analiz Edilecek Tabloyu Seçin", list(ts.keys()))
                tid = ts[sel_name]
                
                id_, adi, df, ack = get_excel_table_details(tid)
                
                if 'pivot_configs' not in st.session_state: st.session_state.pivot_configs = {}
                if tid not in st.session_state.pivot_configs: st.session_state.pivot_configs[tid] = {}

                # Analiz Ekleme
                if st.button("➕ Yeni Analiz ve Grafik Ekle"):
                    new_id = max(st.session_state.pivot_configs[tid].keys(), default=0) + 1
                    st.session_state.pivot_configs[tid][new_id] = {'p_idx': df.columns[0], 'p_col': "Yok", 'p_val': df.columns[0], 'agg': 'Count', 'chart': EXCEL_CHART_TYPES[0]}
                    st.rerun()

                for pid, conf in list(st.session_state.pivot_configs[tid].items()):
                    with st.expander(f"📉 Analiz Paneli #{pid}", expanded=True):
                        c1, c2 = st.columns([1, 2])
                        with c1:
                            cols = df.columns.tolist()
                            p_idx = st.selectbox("Satır (X)", cols, index=cols.index(conf['p_idx']) if conf['p_idx'] in cols else 0, key=f"idx_{tid}_{pid}")
                            p_col = st.selectbox("Sütun (Seri)", ["Yok"] + cols, index=(cols.index(conf['p_col'])+1) if conf['p_col'] in cols else 0, key=f"col_{tid}_{pid}")
                            p_val = st.selectbox("Değer", cols, index=cols.index(conf['p_val']) if conf['p_val'] in cols else 0, key=f"val_{tid}_{pid}")
                            agg = st.selectbox("İşlem", ["Count", "Sum", "Mean", "Max", "Min"], index=["Count", "Sum", "Mean", "Max", "Min"].index(conf['agg']), key=f"agg_{tid}_{pid}")
                            chart = st.selectbox("Excel Grafik Türü", EXCEL_CHART_TYPES, index=EXCEL_CHART_TYPES.index(conf['chart']), key=f"cht_{tid}_{pid}")
                            
                            st.session_state.pivot_configs[tid][pid].update({'p_idx':p_idx, 'p_col':p_col, 'p_val':p_val, 'agg':agg, 'chart':chart})
                            
                            if st.button("🗑️ Analizi Kaldır", key=f"del_{tid}_{pid}"):
                                del st.session_state.pivot_configs[tid][pid]; st.rerun()

                        with c2:
                            try:
                                agg_map = {"Count":"count", "Sum":"sum", "Mean":"mean", "Max":"max", "Min":"min"}
                                p_df = df.pivot_table(index=p_idx, columns=None if p_col=="Yok" else p_col, values=p_val, aggfunc=agg_map[agg]).fillna(0)
                                st.write("**Önizleme:**")
                                st.dataframe(p_df, use_container_width=True)

                                # Plotly Önizleme
                                fig = px.bar(p_df.reset_index(), x=p_idx, y=p_df.columns, barmode="group", title=f"{chart} Önizleme")
                                st.plotly_chart(fig, use_container_width=True)

                                # EXCEL İNDİRME VE KAYDETME
                                buf = io.BytesIO()
                                with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                                    p_df.to_excel(writer, sheet_name='Analiz')
                                    workbook = writer.book
                                    worksheet = writer.sheets['Analiz']
                                    
                                    # XlsxWriter Grafik Eşleştirme
                                    chart_type_map = {
                                        "Sütun (Column)": 'column', "Çubuk (Bar)": 'bar', "Çizgi (Line)": 'line',
                                        "Pasta (Pie)": 'pie', "Halka (Doughnut)": 'doughnut', "Küme Sütun (Clustered Column)": 'column',
                                        "Yığılmış Sütun (Stacked Column)": 'column', "Yüzdelik Yığılmış Sütun (100% Stacked Column)": 'column',
                                        "3-B Sütun (3-D Column)": 'column', "Histogram": 'column'
                                    }
                                    
                                    xlsx_chart = workbook.add_chart({'type': chart_type_map.get(chart, 'column')})
                                    
                                    for i in range(len(p_df.columns)):
                                        xlsx_chart.add_series({
                                            'name':       ['Analiz', 0, i + 1],
                                            'categories': ['Analiz', 1, 0, len(p_df), 0],
                                            'values':     ['Analiz', 1, i + 1, len(p_df), i + 1],
                                        })
                                    
                                    worksheet.insert_chart('G2', xlsx_chart)

                                st.divider()
                                ca, cb = st.columns(2)
                                with ca:
                                    st.download_button("📥 Excel İndir (İndirilenler)", buf.getvalue(), f"{adi}_analiz.xlsx", "application/vnd.ms-excel")
                                with cb:
                                    if st.button("📁 Yerel Klasöre Kaydet", key=f"loc_{tid}_{pid}"):
                                        folder = select_folder()
                                        if folder:
                                            path = os.path.join(folder, f"{adi}_analiz.xlsx")
                                            with open(path, "wb") as f: f.write(buf.getvalue())
                                            st.success(f"Kaydedildi: {path}")

                            except Exception as e: st.error(f"Hata: {e}")

        # ==============================================================================
        # 📂 DİNAMİK TABLO (DOSYA YÜKLE/SİL)
        # ==============================================================================
        elif view == "dinamik_yonetim":
            st.header("📂 Dinamik Tablo Dosya Yönetimi")
            
            with st.form("yukleme_form"):
                st.write("### 📤 Yeni Excel Tablosu Yükle")
                u_name = st.text_input("Tablo Adı")
                u_desc = st.text_area("Açıklama (Opsiyonel)")
                u_file = st.file_uploader("Excel Seç (.xlsx)", type=["xlsx"])
                if st.form_submit_button("Sisteme Kaydet"):
                    if u_file and u_name:
                        with get_db() as conn:
                            conn.execute("INSERT INTO excel_tables (menu_id, name, file_data, description) VALUES (?,?,?,?)",
                                         ("genel_analiz", u_name, u_file.read(), u_desc))
                        st.success("Tablo yüklendi!"); st.rerun()

            st.divider()
            st.write("### 🗑️ Mevcut Tablolar")
            tbls = get_excel_tables("genel_analiz")
            if not tbls:
                st.info("Sistemde yüklü tablo bulunmuyor.")
            else:
                for tid, tname, tdate, tdesc in tbls:
                    with st.expander(f"📄 {tname} (Yükleme: {tdate})"):
                        st.write(f"**Açıklama:** {tdesc}")
                        if st.button("Tabloyu Kalıcı Olarak Sil", key=f"del_{tid}"):
                            delete_excel_table(tid)
                            st.warning("Silindi!"); st.rerun()

if __name__ == "__main__":
    main()
