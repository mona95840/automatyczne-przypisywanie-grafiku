import streamlit as st
import pandas as pd
import io

# Konfiguracja strony
st.set_page_config(page_title="Generator Grafiku", layout="centered")

# Stylizacja i nagłówek
st.title("System Planowania Grafiku")
st.info("""
**Instrukcja:**
1. Wgraj plik Excel z dostępnością pracowników.
2. W panelu bocznym wybierz osobę z limitem zmian (np. pomocnik) i ustaw jej limit.
3. Kliknij **Generuj**, sprawdź podsumowanie i pobierz gotowy plik.
""")

# Panel boczny
st.sidebar.header("Ustawienia limitów")
uploaded_file = st.file_uploader("Wgraj plik .xlsx", type=['xlsx'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    pracownicy = df.columns[2:].tolist()

    prac_ostatni = st.sidebar.selectbox("Pracownik z limitem zmian:", pracownicy, index=len(pracownicy) - 1)
    limit_max = st.sidebar.slider(f"Maksymalna liczba zmian dla: {prac_ostatni}", 1, 20, 6)

    if st.button("Generuj grafik ✨"):
        liczniki = {p: 0 for p in pracownicy}
        ostatnia_zmiana_pracownika = {p: None for p in pracownicy}
        wynikowa_kolumna = []

        for index, row in df.iterrows():
            typ_zmiany = row['ZMIANA']
            dostepni = [p for p in pracownicy if row[p] == 1]

            idealni = []
            do_ratunku = []

            for p in dostepni:
                if p == prac_ostatni and liczniki[p] >= limit_max:
                    continue

                if ostatnia_zmiana_pracownika[p] == 2 and typ_zmiany == 1:
                    do_ratunku.append(p)
                else:
                    idealni.append(p)

            wybrani = []
            idealni_posortowani = sorted(idealni, key=lambda x: liczniki[x])
            wybrani.extend(idealni_posortowani[:2])

            if len(wybrani) < 2:
                brakujacych = 2 - len(wybrani)
                ratownicy_posortowani = sorted(do_ratunku, key=lambda x: liczniki[x])
                wybrani.extend(ratownicy_posortowani[:brakujacych])

            if wybrani:
                for w in wybrani:
                    liczniki[w] += 1
                    ostatnia_zmiana_pracownika[w] = typ_zmiany

                final_names = [f"{w} (!)" if w in do_ratunku else w for w in wybrani]
                wynikowa_kolumna.append(", ".join(final_names))
            else:
                wynikowa_kolumna.append("BRAK OBSADY")

            for p in pracownicy:
                if p not in wybrani:
                    ostatnia_zmiana_pracownika[p] = None

        df['PRZYPISANI_PRACOWNICY'] = wynikowa_kolumna

        # Prezentacja wyników
        st.divider()
        st.subheader("Podgląd wygenerowanego grafiku")
        st.write("(*) Oznacza osobę dobraną awaryjnie mimo braku pełnego odpoczynku.")
        st.dataframe(df[['DATA', 'ZMIANA', 'PRZYPISANI_PRACOWNICY']], use_container_width=True)

        st.subheader("Podsumowanie liczby zmian")
        st.bar_chart(pd.Series(liczniki))

        # Przygotowanie do pobrania
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Grafik_Wynik')

        st.download_button(
            label="📥 Pobierz grafik jako Excel",
            data=output.getvalue(),
            file_name="gotowy_grafik.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
