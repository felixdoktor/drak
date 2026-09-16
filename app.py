import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import os
import difflib
import urllib.request
from bs4 import BeautifulSoup
import requests

st.set_page_config(page_title="Startovní rozpis dračích lodí - Vícedenní", layout="wide")

# ===================================================================
# PŘIHLÁŠENÍ HESLEM
# ===================================================================
def over_heslo():
    """Zobrazí přihlašovací formulář a zastaví běh aplikace, dokud není zadáno správné heslo."""
    VYZADOVAT_HESLO = False   # 👈 Nastavte na False pro vypnutí, nebo True pro zapnutí
    SPRAVNE_HESLO = "draci2026"

    # Pokud je heslo vypnuté, rovnou pustíme uživatele dál
    if not VYZADOVAT_HESLO:
        return True

    if "prihlasen" not in st.session_state:
        st.session_state["prihlasen"] = False

    if st.session_state["prihlasen"]:
        return True

    # Přihlašovací okno
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("🔒 Přihlášení do generátoru")
        heslo_vstup = st.text_input("Zadejte přístupové heslo:", type="password")
        
        if st.button("Přihlásit se", type="primary", use_container_width=True):
            if heslo_vstup == SPRAVNE_HESLO:
                st.session_state["prihlasen"] = True
                st.rerun()
            else:
                st.error("❌ Nesprávné heslo!")

    st.stop()

# Volání kontroly zůstává beze změny:
over_heslo()



# ===================================================================
# STAHOVÁNÍ ŽEBŘÍČKU Z DRAGONBOAT.CZ
# ===================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def nacti_zebricek_cpo():
    """Stáhne oficiální pořadí klubů a body z webu dragonboat.cz."""
    url = "https://www.dragonboat.cz/poradi/"
    
    # Kompletní hlavičky, které simulují reálný prohlížeč z mobilu/PC
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "cs-CZ,cs;q=0.9,en;q=0.8",
        "Referer": "https://www.dragonboat.cz/",
        "Connection": "keep-alive"
    }

    try:
        session = requests.Session()
        resp = session.get(url, headers=headers, timeout=10)
        
        if resp.status_code != 200:
            return {}

        soup = BeautifulSoup(resp.text, "html.parser")
        
        # Hledání tabulky výsledků
        tabulky = soup.find_all("table")
        if not tabulky:
            return {}

        zebricek = {}
        for row in tabulky[0].find_all("tr"):
            bunky = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
            
            # Formát tabulky na dragonboat.cz:
            # bunky obvykle obsahují: [Pořadí, Název klubu / posádky, Celkem bodů, ...]
            if len(bunky) >= 3:
                nazev = bunky[1].strip()
                # Vyfiltrování pouze číslic pro celkové body
                body_cislice = "".join(ch for ch in bunky[2] if ch.isdigit())
                if nazev and body_cislice:
                    zebricek[nazev] = int(body_cislice)

        return zebricek

    except Exception:
        return {}

# ===================================================================
# 1. ZÁKLADNÍ PARAMETRY ZÁVODU
# ===================================================================
    
with st.sidebar:
    st.header("⚙️ Nastavení mistrovství")
    zavod_titul = st.text_input("Název akce:", value="20. MISTROVSTVÍ ČR DRAČÍCH LODÍ")
    zavod_misto_datum = st.text_input("Místo konání:", value="LABE ARÉNA RAČICE")
    zavod_podtitul = st.text_input("Podtitul / Seriál:", value="Euro Grand Prix Race")

    st.divider()

    st.subheader("🏁 Pravidla nasazování")
    oddelit_stejne_kluby = st.checkbox(
        "🚫 Oddělit posádky ze stejného oddílu v rozjížďkách", 
        value=True,
        help="Zajistí, aby se posádky jednoho klubu nepotkaly hned v základní rozjížďce."
    )

    st.divider()

    # --- SEKCE ČESKÉHO POHÁRU (ČP) ---
    st.subheader("🏆 Český pohár (ČP)")
    pouzit_cpo = st.checkbox(
        "Nasazovat podle žebříčku ČP", 
        value=False,
        help="Seřadí posádky podle bodů v ČP a rozdělí je do rozjížděk serpentýnou s výhodnými dráhami."
    )

    zebricek_cp = {}

    if pouzit_cpo:
        with st.spinner("Stahuji aktuální žebříček z dragonboat.cz..."):
            zebricek_cp = nacti_zebricek_cpo()

        if zebricek_cp:
            st.success(f"✅ Žebříček ČP načten online ({len(zebricek_cp)} klubů)")
            with st.expander("👀 Zobrazit stažené pořadí a body"):
                df_nahled_cp = pd.DataFrame(
                    list(zebricek_cp.items()), 
                    columns=["Klub / Posádka", "Body ČP"]
                ).sort_values(by="Body ČP", ascending=False).reset_index(drop=True)
                st.dataframe(df_nahled_cp, use_container_width=True, height=250)
        else:
            st.warning("⚠️ Web ČADL dočasně zablokoval automatické stažení (ochrana serveru).")
            st.caption("Nahrajte žebříček ručně (Excel nebo CSV se sloupci 'klub' a 'body'):")
            zebricek_file = st.file_uploader(
                "Vybrat soubor žebříčku:", 
                type=["xlsx", "xls", "csv"], 
                key="cpo_backup"
            )
            if zebricek_file:
                try:
                    if zebricek_file.name.endswith(".csv"):
                        df_z = pd.read_csv(zebricek_file)
                    else:
                        df_z = pd.read_excel(zebricek_file)

                    # Flexibilní nalezení sloupců pro klub a body
                    col_t = next(c for c in df_z.columns if any(k in c.lower() for k in ["oddíl", "oddil", "klub", "tým", "tym", "posádk", "posadk"]))
                    col_b = next(c for c in df_z.columns if any(k in c.lower() for k in ["bod", "skóre", "skore", "celkem"]))

                    df_z[col_t] = df_z[col_t].astype(str).str.strip()
                    df_z[col_b] = pd.to_numeric(df_z[col_b].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(0).astype(int)

                    zebricek_cp = dict(zip(df_z[col_t], df_z[col_b]))
                    st.success(f"✅ Ruční žebříček načten ({len(zebricek_cp)} klubů)!")
                except Exception as e:
                    st.error(f"Chyba při načítání souboru žebříčku: {e}")

    st.divider()
    st.caption("Verze aplikace: 2.1 (Online Cloud Edition)")




# ===================================================================
# 2. NAHRÁNÍ PŘIHLÁŠEK (EXCEL / CSV / TXT)
# ===================================================================
st.header("1. Nahrání přihlášek ze systému")

uploaded_file = st.file_uploader(
    "Nahrajte export přihlášek (Excel .xlsx / .xls nebo CSV / TXT):", 
    type=["xlsx", "xls", "csv", "txt"]
)

df_raw = None

if uploaded_file is not None:
    file_name = uploaded_file.name.lower()
    
    try:
        # A) Zpracování Excel souborů (.xlsx, .xls)
        if file_name.endswith((".xlsx", ".xls")):
            df_raw = pd.read_excel(uploaded_file)
            st.success(f"✅ Excel soubor `{uploaded_file.name}` byl úspěšně načten.")

        # B) Zpracování CSV / textových souborů
        else:
            bytes_data = uploaded_file.getvalue()
            # Automatická detekce kódování (české znaky v Excel CSV bývají v cp1250)
            nacteno = False
            for kodovani in ["utf-8-sig", "utf-8", "cp1250", "iso-8859-2", "latin2"]:
                try:
                    text_content = bytes_data.decode(kodovani)
                    # Detekce oddělovače (středník vs. čárka vs. tabulátor)
                    prvni_radky = text_content[:2000]
                    sep = ";" if prvni_radky.count(";") >= prvni_radky.count(",") else ","
                    if prvni_radky.count("\t") > prvni_radky.count(sep):
                        sep = "\t"
                        
                    df_raw = pd.read_csv(io.StringIO(text_content), sep=sep)
                    nacteno = True
                    st.success(f"✅ CSV soubor `{uploaded_file.name}` načten (kódování: {kodovani}).")
                    break
                except (UnicodeDecodeError, Exception):
                    continue

            if not nacteno:
                st.error("❌ Nepodařilo se rozpoznat kódování CSV souboru. Zkuste soubor uložit jako Excel (.xlsx).")
                st.stop()

    except Exception as e:
        st.error(f"❌ Chyba při otevírání souboru: {e}")
        st.stop()

    # Očištění názvů sloupců (odstranění mezer)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    # Vyhledání klíčových sloupců bez ohledu na velikost písmen a diakritiku
    def najdi_sloupec(kandidati):
        for col in df_raw.columns:
            col_low = col.lower()
            if any(k in col_low for k in kandidati):
                return col
        return None

    col_trat = najdi_sloupec(["trať", "trat", "vzdalenost", "distance", "délka", "delka"])
    col_kat = najdi_sloupec(["kategorie", "kat", "category"])
    col_oddil = najdi_sloupec(["oddíl", "oddil", "klub", "tým", "tym", "posádk", "posadk", "team", "club"])

    # Kontrola povinných sloupců
    chybejici = []
    if not col_trat: chybejici.append("trať (např. 200m, 500m, 1000m)")
    if not col_kat: chybejici.append("kategorie (např. MIX, OPEN, ŽENY)")
    if not col_oddil: chybejici.append("oddíl / tým / posádka")

    if chybejici:
        st.error(f"⚠️ V souboru chybí následující povinné sloupce: **{', '.join(chybejici)}**.")
        st.write("Nalezené sloupce v souboru:", list(df_raw.columns))
        df_raw = None
    else:
        # Převedení hodnot na čistý text
        df_raw[col_trat] = df_raw[col_trat].astype(str).str.strip()
        df_raw[col_kat] = df_raw[col_kat].astype(str).str.strip()
        df_raw[col_oddil] = df_raw[col_oddil].astype(str).str.strip()

        with st.expander("🔍 Náhled a ruční úprava přihlášek (pokud je potřeba něco opravit)", expanded=False):
            st.info("Zde můžete přímo v tabulce přepsat překlepy v názvu týmu nebo upravit kategorii.")
            df_raw = st.data_editor(
                df_raw,
                column_order=[col_trat, col_kat, col_oddil],
                use_container_width=True,
                num_rows="dynamic"
            )

        # Unikátní identifikátor disciplíny
        df_raw["_Disciplina_ID"] = df_raw[col_trat] + "m | " + df_raw[col_kat]
        vsechny_discipliny = sorted(df_raw["_Disciplina_ID"].unique().tolist())
        st.write(f"Celkem přihlášeno **{len(df_raw)} posádek** do **{len(vsechny_discipliny)}** různých disciplín.")

    # ===================================================================
    # 3. NASTAVENÍ JEDNOTLIVÝCH DNŮ
    # ===================================================================
    st.divider()
    st.header("2. Nastavení jednotlivých závodních dnů")

    pocet_dnu = st.slider("Počet závodních dnů:", min_value=1, max_value=4, value=3)

    konfigurace_dnu = []
    vychozi_nazvy = ["Pátek", "Sobota", "Neděle", "Pondělí"]
    vychozi_casy = ["14:00", "09:00", "09:00", "09:00"]

    for den_i in range(pocet_dnu):
        with st.expander(f"📅 Nastavení dne {den_i + 1}: {vychozi_nazvy[den_i]}", expanded=True):
            c_d1, c_d2, c_d3 = st.columns(3)
            with c_d1:
                den_nazev = st.text_input(f"Název dne {den_i + 1}:", value=vychozi_nazvy[den_i], key=f"d_nazev_{den_i}")
                den_start_cas = st.text_input(f"Čas startu dne {den_i + 1}:", value=vychozi_casy[den_i], key=f"d_cas_{den_i}")
            with c_d2:
                den_drah = st.selectbox(f"Počet drah pro {den_nazev}:", options=[2, 3, 4, 5, 6, 8], index=4, key=f"d_drah_{den_i}")
                den_interval = st.number_input(f"Interval mezi jízdami (min) pro {den_nazev}:", min_value=5, max_value=60, value=15, key=f"d_int_{den_i}")
            with c_d3:
                den_system = st.selectbox(
                    f"Systém rozjížděk pro {den_nazev}:",
                    options=[
                        "AUTO", 
                        "DLOUHE_TRATE", 
                        "POUZE_ROZJIZDKY_A_FINALE",
                        "PEVNE_2_ROZJIZDKY",
                        "PEVNE_3_ROZJIZDKY"
                    ],
                    format_func=lambda x: {
                        "AUTO": "Automatický pavouk (podle počtu posádek)",
                        "DLOUHE_TRATE": "Dlouhé tratě: Pouze finálové jízdy D1, D2... (na čas)",
                        "POUZE_ROZJIZDKY_A_FINALE": "Pouze rozjížďky a Finále A/B (bez semifinále)",
                        "PEVNE_2_ROZJIZDKY": "Pevně 2 rozjížďky + Finále A/B",
                        "PEVNE_3_ROZJIZDKY": "Pevně 3 rozjížďky + SF + Finále A/B"
                    }.get(x, x),
                    key=f"d_sys_{den_i}"
                )

            vybrane_disc = st.multiselect(
                f"Disciplíny jedoucí se v den '{den_nazev}':",
                options=vsechny_discipliny,
                key=f"d_disc_{den_i}"
            )

            konfigurace_dnu.append({
                "nazev": den_nazev,
                "start_cas": den_start_cas,
                "drah": den_drah,
                "interval": den_interval,
                "system": den_system,
                "discipliny": vybrane_disc
            })

    # ===================================================================
    # 4. SOLVER ENGINE S GARANCÍ IZOLACE KATEGORIÍ
    # ===================================================================
    def rozrad_posadky_do_jizd_solver(posadky_list, drah, zadany_pocet_jizd=None, sep_oddily=True):
        n = len(posadky_list)
        if zadany_pocet_jizd is not None:
            pocet_jizd = zadany_pocet_jizd
        else:
            pocet_jizd = (n + drah - 1) // drah

        if pocet_jizd <= 0 or n == 0:
            return []

        model = cp_model.CpModel()
        jizdy = [model.NewIntVar(0, pocet_jizd - 1, f"p_{i}") for i in range(n)]

        min_v_jizde = n // pocet_jizd
        for j in range(pocet_jizd):
            in_heat = [model.NewBoolVar(f"ih_{i}_{j}") for i in range(n)]
            for i in range(n):
                model.Add(jizdy[i] == j).OnlyEnforceIf(in_heat[i])
                model.Add(jizdy[i] != j).OnlyEnforceIf(in_heat[i].Not())
            model.Add(sum(in_heat) <= drah)
            model.Add(sum(in_heat) >= min_v_jizde)

        # Zákaz potkání stejných oddílů ve stejné rozjížďce
        if sep_oddily:
            for t in set(posadky_list):
                t_ids = [i for i, x in enumerate(posadky_list) if x == t]
                if 1 < len(t_ids) <= pocet_jizd:
                    model.AddAllDifferent([jizdy[i] for i in t_ids])

        solver = cp_model.CpSolver()
        status = solver.Solve(model)

        # Nasazení drah od středu (např. 3, 4, 2, 5, 1, 6)
        stred = (drah + 1) / 2
        poradi_drah = sorted(range(1, drah + 1), key=lambda x: (abs(x - stred), x))

        jizdy_data = []
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            pos_prirazeni = [(posadky_list[i], solver.Value(jizdy[i])) for i in range(n)]
            for j_idx in range(pocet_jizd):
                tymy_v_jizde = [t for t, j in pos_prirazeni if j == j_idx]
                drahy_dict = {}
                for d_num, tym in zip(poradi_drah, tymy_v_jizde):
                    drahy_dict[d_num] = tym
                jizdy_data.append(drahy_dict)
        return jizdy_data

    def vytvor_jizdy_pro_disciplinu(kat_nazev, trat_nazev, posadky, den_cfg, sep_oddily, cislo_dlouhe_start=1):
        n = len(posadky)
        drah = den_cfg["drah"]
        sys_typ = den_cfg["system"]
        trat_cislo = int(''.join(filter(str.isdigit, str(trat_nazev))) or 0)
        
        je_dlouha = (sys_typ == "DLOUHE_TRATE") or (trat_cislo >= 1000)
        nazev_discipliny = f"{kat_nazev} - {trat_nazev}m"
        jizdy = []

        if je_dlouha:
            rozrazene = rozrad_posadky_do_jizd_solver(posadky, drah, None, sep_oddily)
            for i, drahy in enumerate(rozrazene):
                kod = f"D{cislo_dlouhe_start + i}"
                jizdy.append({
                    "kod": kod,
                    "nazev": f"{nazev_discipliny} - jízda {kod}",
                    "klic": "O umístění rozhoduje dosažený čas!",
                    "posadky": drahy,
                    "je_dlouha": True
                })
        elif sys_typ == "PEVNE_2_ROZJIZDKY":
            rozrazene = rozrad_posadky_do_jizd_solver(posadky, drah, zadany_pocet_jizd=2, sep_oddily=sep_oddily)
            jizdy.append({"kod": "R1", "nazev": f"{nazev_discipliny} - Rozjížďka 1", "klic": "1.-3. -> FA, 4.-6. -> FB", "posadky": rozrazene[0] if len(rozrazene) > 0 else {}, "je_dlouha": False})
            jizdy.append({"kod": "R2", "nazev": f"{nazev_discipliny} - Rozjížďka 2", "klic": "1.-3. -> FA, 4.-6. -> FB", "posadky": rozrazene[1] if len(rozrazene) > 1 else {}, "je_dlouha": False})
            jizdy.append({"kod": "FB", "nazev": f"{nazev_discipliny} - FINÁLE B", "klic": "O konečné 7.-12. místo", "posadky": {}, "je_dlouha": False})
            jizdy.append({"kod": "FA", "nazev": f"{nazev_discipliny} - FINÁLE A", "klic": "O 1.-6. místo a medaile", "posadky": {}, "je_dlouha": False})

        elif sys_typ == "PEVNE_3_ROZJIZDKY":
            rozrazene = rozrad_posadky_do_jizd_solver(posadky, drah, zadany_pocet_jizd=3, sep_oddily=sep_oddily)
            for i in range(3):
                drahy = rozrazene[i] if i < len(rozrazene) else {}
                jizdy.append({"kod": f"R{i+1}", "nazev": f"{nazev_discipliny} - Rozjížďka {i+1}", "klic": "1.-2. -> FA, 3.-4. -> Semifinále", "posadky": drahy, "je_dlouha": False})
            jizdy.append({"kod": "SF1", "nazev": f"{nazev_discipliny} - Semifinále 1", "klic": "1.-3. -> FA, ostatní -> FB", "posadky": {}, "je_dlouha": False})
            jizdy.append({"kod": "FB", "nazev": f"{nazev_discipliny} - FINÁLE B", "klic": "Finále B", "posadky": {}, "je_dlouha": False})
            jizdy.append({"kod": "FA", "nazev": f"{nazev_discipliny} - FINÁLE A", "klic": "Finále A o medaile", "posadky": {}, "je_dlouha": False})

        elif sys_typ == "POUZE_ROZJIZDKY_A_FINALE":
            pocet_r = max(1, (n + drah - 1) // drah)
            rozrazene = rozrad_posadky_do_jizd_solver(posadky, drah, zadany_pocet_jizd=pocet_r, sep_oddily=sep_oddily)
            if pocet_r == 1:
                jizdy.append({"kod": "FA", "nazev": f"{nazev_discipliny} - FINÁLE A", "klic": "Přímá finálová jízda", "posadky": rozrazene[0] if rozrazene else {}, "je_dlouha": False})
            else:
                for i, drahy in enumerate(rozrazene, 1):
                    jizdy.append({"kod": f"R{i}", "nazev": f"{nazev_discipliny} - Rozjížďka {i}", "klic": "Postup do Finále A/B podle klíče", "posadky": drahy, "je_dlouha": False})
                jizdy.append({"kod": "FB", "nazev": f"{nazev_discipliny} - FINÁLE B", "klic": "Finále B", "posadky": {}, "je_dlouha": False})
                jizdy.append({"kod": "FA", "nazev": f"{nazev_discipliny} - FINÁLE A", "klic": "Finále A o medaile", "posadky": {}, "je_dlouha": False})

        else: # AUTO
            if n <= drah:
                rozrazene = rozrad_posadky_do_jizd_solver(posadky, drah, None, sep_oddily)
                jizdy.append({"kod": "FA", "nazev": f"{nazev_discipliny} - FINÁLE A", "klic": "Finálová jízda o medaile", "posadky": rozrazene[0] if rozrazene else {}, "je_dlouha": False})
            elif n <= 2 * drah:
                rozrazene = rozrad_posadky_do_jizd_solver(posadky, drah, 2, sep_oddily)
                jizdy.append({"kod": "R1", "nazev": f"{nazev_discipliny} - Rozjížďka 1", "klic": "1.-3. -> FA, 4.-6. -> FB", "posadky": rozrazene[0], "je_dlouha": False})
                jizdy.append({"kod": "R2", "nazev": f"{nazev_discipliny} - Rozjížďka 2", "klic": "1.-3. -> FA, 4.-6. -> FB", "posadky": rozrazene[1], "je_dlouha": False})
                jizdy.append({"kod": "FB", "nazev": f"{nazev_discipliny} - FINÁLE B", "klic": "O konečné 7.-12. místo", "posadky": {}, "je_dlouha": False})
                jizdy.append({"kod": "FA", "nazev": f"{nazev_discipliny} - FINÁLE A", "klic": "O 1.-6. místo a medaile", "posadky": {}, "je_dlouha": False})
            else:
                rozrazene = rozrad_posadky_do_jizd_solver(posadky, drah, None, sep_oddily)
                for i, drahy in enumerate(rozrazene, 1):
                    jizdy.append({"kod": f"R{i}", "nazev": f"{nazev_discipliny} - Rozjížďka {i}", "klic": "1.-2. -> FA, 3.-4. -> Semifinále, ostatní končí", "posadky": drahy, "je_dlouha": False})
                jizdy.append({"kod": "SF1", "nazev": f"{nazev_discipliny} - Semifinále 1", "klic": "1.-3. -> FA, ostatní -> FB", "posadky": {}, "je_dlouha": False})
                jizdy.append({"kod": "FB", "nazev": f"{nazev_discipliny} - FINÁLE B", "klic": "Finále B", "posadky": {}, "je_dlouha": False})
                jizdy.append({"kod": "FA", "nazev": f"{nazev_discipliny} - FINÁLE A", "klic": "Finále A o medaile", "posadky": {}, "je_dlouha": False})

        return jizdy

    # Přísná hierarchie barev (Junioři mají přednost před MIXEM)
    def ziskej_barvu_kategorie(nazev):
        n = str(nazev).lower()
        if "jun" in n or "u15" in n or "u18" in n:
            return "FFFF00"   # Žlutá pro všechny juniorské posádky
        elif "žen" in n or "women" in n:
            return "FFCCFF"   # Růžová pro ženy
        elif "mix" in n:
            return "91CF50"   # Zelená pro dospělý MIX
        elif "open" in n or "muž" in n:
            return "93DBF7"   # Modrá pro OPEN
        return "CAEDFB"

    # ===================================================================
    # 5. GENERÁTOR EXCELU
    # ===================================================================
    def vytvor_vicedenni_excel(data_dnu, info_zavod):
        wb = Workbook()
        wb.remove(wb.active)

        f_title = Font(name="Arial", size=14, bold=True)
        f_head_white = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        f_bold = Font(name="Arial", size=10, bold=True)
        f_sub = Font(name="Arial", size=10, bold=True)
        f_row = Font(name="Arial", size=10)
        f_klic = Font(name="Arial", size=9, italic=True)
        fill_head = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        fill_gray_h = PatternFill(start_color="D8D8D8", end_color="D8D8D8", fill_type="solid")
        border_all = Border(left=Side(style="thin", color="000000"), right=Side(style="thin", color="000000"),
                            top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"))

        for den_data in data_dnu:
            den_nazev = den_data["nazev"]
            den_drah = den_data["drah"]
            jizdy_dne = den_data["jizdy"]
            if not jizdy_dne:
                continue

            # LIST 1: Program dne
            ws_prog = wb.create_sheet(title=f"{den_nazev} - Program")
            ws_prog.page_setup.orientation = ws_prog.ORIENTATION_LANDSCAPE
            ws_prog.page_setup.paperSize = ws_prog.PAPERSIZE_A4

            ws_prog.merge_cells("A1:E1")
            ws_prog["A1"] = f"{info_zavod['titul']} - ČASOVÝ PROGRAM ({den_nazev.upper()}) [Drah: {den_drah}]"
            ws_prog["A1"].font = f_title
            ws_prog["A1"].alignment = Alignment(horizontal="center", vertical="center")

            hlavicky = ["Závod č.", "Čas startu", "Jízda / Disciplína", "Počet lodí", "Postupový klíč"]
            for c_i, h in enumerate(hlavicky, 1):
                cell = ws_prog.cell(row=3, column=c_i, value=h)
                cell.font = f_head_white
                cell.fill = fill_head
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border_all

            for j in jizdy_dne:
                r = 3 + (j["cislo_zavodu"] - jizdy_dne[0]["cislo_zavodu"] + 1)
                ws_prog.cell(row=r, column=1, value=j["cislo_zavodu"]).alignment = Alignment(horizontal="center", vertical="center")
                ws_prog.cell(row=r, column=2, value=j["cas"]).alignment = Alignment(horizontal="center", vertical="center")
                
                c_nazev = ws_prog.cell(row=r, column=3, value=j["nazev"])
                c_nazev.alignment = Alignment(horizontal="left", vertical="center")
                c_nazev.fill = PatternFill(start_color=ziskej_barvu_kategorie(j["nazev"]), end_color=ziskej_barvu_kategorie(j["nazev"]), fill_type="solid")

                ws_prog.cell(row=r, column=4, value=len(j["posadky"]) if j["posadky"] else den_drah).alignment = Alignment(horizontal="center", vertical="center")
                ws_prog.cell(row=r, column=5, value=j["klic"]).alignment = Alignment(horizontal="left", vertical="center")

                for c_i in range(1, 6):
                    ws_prog.cell(row=r, column=c_i).border = border_all

            ws_prog.column_dimensions["A"].width = 10
            ws_prog.column_dimensions["B"].width = 12
            ws_prog.column_dimensions["C"].width = 48
            ws_prog.column_dimensions["D"].width = 12
            ws_prog.column_dimensions["E"].width = 38

            # LIST 2: Rozpis dne
            ws_rozpis = wb.create_sheet(title=f"{den_nazev} - Rozpis")
            ws_rozpis.page_setup.orientation = ws_rozpis.ORIENTATION_PORTRAIT
            ws_rozpis.page_setup.paperSize = ws_rozpis.PAPERSIZE_A4

            curr = 1
            for t in [info_zavod["titul"], f"{info_zavod['misto_datum']} - {den_nazev.upper()}", info_zavod["podtitul"]]:
                ws_rozpis.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=5)
                c = ws_rozpis.cell(row=curr, column=1, value=t)
                c.font = f_sub
                c.alignment = Alignment(horizontal="center", vertical="center")
                curr += 1
            curr += 1

            for j in jizdy_dne:
                fill_kat = PatternFill(start_color=ziskej_barvu_kategorie(j["nazev"]), end_color=ziskej_barvu_kategorie(j["nazev"]), fill_type="solid")

                c_num = ws_rozpis.cell(row=curr, column=1, value=j["cislo_zavodu"])
                c_num.font = f_bold
                c_num.alignment = Alignment(horizontal="center", vertical="center")
                c_num.fill = fill_kat
                c_num.border = border_all

                ws_rozpis.merge_cells(start_row=curr, start_column=2, end_row=curr, end_column=3)
                c_tit = ws_rozpis.cell(row=curr, column=2, value=j["nazev"])
                c_tit.font = f_bold
                c_tit.fill = fill_kat
                c_tit.alignment = Alignment(horizontal="left", vertical="center")

                for c_i, h in [(4, "Čas"), (5, "Pořadí")]:
                    cell = ws_rozpis.cell(row=curr, column=c_i, value=h)
                    cell.font = f_bold
                    cell.fill = fill_gray_h
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = border_all

                for c_i in range(1, 6):
                    ws_rozpis.cell(row=curr, column=c_i).border = border_all

                start_drah = curr + 1

                for d_num in range(1, den_drah + 1):
                    r = start_drah + d_num - 1
                    ws_rozpis.cell(row=r, column=2, value=d_num).alignment = Alignment(horizontal="center", vertical="center")
                    ws_rozpis.cell(row=r, column=2).font = f_bold
                    ws_rozpis.cell(row=r, column=2).border = border_all

                    tym = j["posadky"].get(d_num, "")
                    ws_rozpis.cell(row=r, column=3, value=tym).alignment = Alignment(horizontal="left", vertical="center")
                    ws_rozpis.cell(row=r, column=3).font = f_row
                    ws_rozpis.cell(row=r, column=3).border = border_all

                    ws_rozpis.cell(row=r, column=4, value="").border = border_all
                    ws_rozpis.cell(row=r, column=5, value="").border = border_all

                polovina = den_drah // 2
                ws_rozpis.merge_cells(start_row=start_drah, start_column=1, end_row=start_drah + polovina - 1, end_column=1)
                c_k = ws_rozpis.cell(row=start_drah, column=1, value=j["kod"])
                c_k.font = f_bold
                c_k.alignment = Alignment(horizontal="center", vertical="center")
                c_k.fill = fill_kat

                ws_rozpis.merge_cells(start_row=start_drah + polovina, start_column=1, end_row=start_drah + den_drah - 1, end_column=1)
                c_t = ws_rozpis.cell(row=start_drah + polovina, column=1, value=j["cas"])
                c_t.font = f_bold
                c_t.alignment = Alignment(horizontal="center", vertical="center")

                for r_x in range(start_drah, start_drah + den_drah):
                    ws_rozpis.cell(row=r_x, column=1).border = border_all

                curr = start_drah + den_drah

                ws_rozpis.cell(row=curr, column=1, value="Klíč:").font = f_bold
                ws_rozpis.merge_cells(start_row=curr, start_column=2, end_row=curr, end_column=5)
                ws_rozpis.cell(row=curr, column=2, value=j["klic"]).font = f_klic
                curr += 2

            ws_rozpis.column_dimensions["A"].width = 10
            ws_rozpis.column_dimensions["B"].width = 6
            ws_rozpis.column_dimensions["C"].width = 38
            ws_rozpis.column_dimensions["D"].width = 12
            ws_rozpis.column_dimensions["E"].width = 10

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    # ===================================================================
    # 6. SPUŠTĚNÍ VÝPOČTU A ULOŽENÍ VÝSLEDKŮ
    # ===================================================================
    st.divider()

    # Nativní dialog Windows "Uložit jako..." přes Win32 API (ctypes)
    def vyber_kam_ulozit_windows(data_bytes, vychozi_nazev="vicedenni_rozpis_zavodu.xlsx"):
        import ctypes
        from ctypes import wintypes

        class OPENFILENAME(ctypes.Structure):
            _fields_ = [
                ("lStructSize", wintypes.DWORD),
                ("hwndOwner", wintypes.HWND),
                ("hInstance", wintypes.HINSTANCE),
                ("lpstrFilter", wintypes.LPCWSTR),
                ("lpstrCustomFilter", wintypes.LPWSTR),
                ("nMaxCustFilter", wintypes.DWORD),
                ("nFilterIndex", wintypes.DWORD),
                ("lpstrFile", wintypes.LPWSTR),
                ("nMaxFile", wintypes.DWORD),
                ("lpstrFileTitle", wintypes.LPWSTR),
                ("nMaxFileTitle", wintypes.DWORD),
                ("lpstrInitialDir", wintypes.LPCWSTR),
                ("lpstrTitle", wintypes.LPCWSTR),
                ("Flags", wintypes.DWORD),
                ("nFileOffset", wintypes.WORD),
                ("nFileExtension", wintypes.WORD),
                ("lpstrDefExt", wintypes.LPCWSTR),
                ("lCustData", wintypes.LPARAM),
                ("lpfnHook", ctypes.c_void_p),
                ("lpTemplateName", wintypes.LPCWSTR),
                ("pvReserved", ctypes.c_void_p),
                ("dwReserved", wintypes.DWORD),
                ("FlagsEx", wintypes.DWORD)
            ]

        buffer = ctypes.create_unicode_buffer(vychozi_nazev, 1024)
        ofn = OPENFILENAME()
        ofn.lStructSize = ctypes.sizeof(OPENFILENAME)
        ofn.lpstrFilter = "Excel sešit (*.xlsx)\0*.xlsx\0Všechny soubory (*.*)\0*.*\0\0"
        ofn.lpstrFile = ctypes.cast(buffer, wintypes.LPWSTR)
        ofn.nMaxFile = 1024
        ofn.lpstrTitle = "Zvolte kam uložit rozpis závodu"
        ofn.lpstrDefExt = "xlsx"
        ofn.Flags = 0x00080000 | 0x00000002

        if ctypes.windll.comdlg32.GetSaveFileNameW(ctypes.byref(ofn)):
            vybrana_cesta = buffer.value
            if vybrana_cesta:
                with open(vybrana_cesta, "wb") as f:
                    f.write(data_bytes.getbuffer())
                return vybrana_cesta
        return None

    if st.button("🚀 Vygenerovat vícedenní rozpis", type="primary", use_container_width=True):
        data_vsech_dnu = []
        globalni_cislo_zavodu = 1
        pocitadlo_dlouhych_global = 1

        for den in konfigurace_dnu:
            if not den["discipliny"]:
                continue

            jizdy_tohoto_dne = []
            sh, sm = map(int, den["start_cas"].split(":"))

            def ziskej_metry(disc_str):
                return int(''.join(filter(str.isdigit, disc_str.split('|')[0])) or 0)

            kratke = [d for d in den["discipliny"] if ziskej_metry(d) < 1000]
            dlouhe = [d for d in den["discipliny"] if ziskej_metry(d) >= 1000]
            serazene = kratke + dlouhe

            for disc_id in serazene:
                subset = df_raw[df_raw["_Disciplina_ID"] == disc_id].copy().reset_index(drop=True)
                if subset.empty:
                    continue

                t_val = str(subset[col_trat].iloc[0]).strip()
                k_val = str(subset[col_kat].iloc[0]).strip()
                posadky = subset[col_oddil].dropna().astype(str).str.strip().tolist()

                jizdy_bloku = vytvor_jizdy_pro_disciplinu(k_val, t_val, posadky, den, oddelit_stejne_kluby, pocitadlo_dlouhych_global)

                for j in jizdy_bloku:
                    j["cas"] = f"{sh:02d}:{sm:02d}"
                    j["cislo_zavodu"] = globalni_cislo_zavodu
                    jizdy_tohoto_dne.append(j)

                    globalni_cislo_zavodu += 1
                    if j["je_dlouha"]:
                        pocitadlo_dlouhych_global += 1

                    sm += den["interval"]
                    sh += sm // 60
                    sm %= 60

            data_vsech_dnu.append({
                "nazev": den["nazev"],
                "drah": den["drah"],
                "jizdy": jizdy_tohoto_dne
            })

        if not any(d["jizdy"] for d in data_vsech_dnu):
            st.warning("Přiřaďte prosím alespoň do jednoho dne nějaké disciplíny.")
        else:
            info_zavodu = {
                "titul": zavod_titul,
                "misto_datum": zavod_misto_datum,
                "podtitul": zavod_podtitul
            }
            st.session_state["vysledna_data_dnu"] = data_vsech_dnu
            st.session_state["excel_bytes"] = vytvor_vicedenni_excel(data_vsech_dnu, info_zavodu)

    # --- ZOBRAZENÍ TABULEK A TLAČÍTKA MIMO GENERUJÍCÍ BLOK ---
    # Zobrazení výsledků a tlačítka pro uložení
    if "vysledna_data_dnu" in st.session_state and st.session_state["vysledna_data_dnu"]:
        data_dnu = st.session_state["vysledna_data_dnu"]
        celkem_jizd = sum(len(d["jizdy"]) for d in data_dnu)
        
        st.success(f"✅ Rozpis vygenerován: celkem **{celkem_jizd} jízd** bez jakéhokoliv křížení kategorií!")

        for d in data_dnu:
            st.subheader(f"Přehled: {d['nazev']} ({d['drah']} drah)")
            nahled_df = pd.DataFrame([
                {"Závod č.": j["cislo_zavodu"], "Čas": j["cas"], "Jízda": j["nazev"], "Klíč": j["klic"]}
                for j in d["jizdy"]
            ])
            st.dataframe(nahled_df, use_container_width=True)

        st.divider()

        # ZDE NAHRAĎTE PŮVODNÍ WINDOWS TLAČÍTKO TÍMTO ŘÁDKEM PRO WEB:
        st.download_button(
            label="📥 STÁHNOUT VÝSLEDNÝ EXCEL ROZPIS (.xlsx)",
            data=st.session_state["excel_bytes"].getvalue(),
            file_name="rozpis_zavodu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )