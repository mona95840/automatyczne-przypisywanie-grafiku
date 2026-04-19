import streamlit as st
import pandas as pd
import io

# Konfiguracja strony
st.set_page_config(page_title="Generator Grafiku", layout="centered")

# Nagłówek i instrukcja
st.title("System Planowania Grafiku")
st.info("""
**Instrukcja:**
1. Wgraj plik Excel z dostępnością (1 = dostępny, 0 lub puste = niedostępny).
2. W panelu bocznym wybierz pracownika z limitem zmian i ustaw jego maksimum.
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
            # Sprawdzamy dostępność: tylko wartość 1 jest brana pod uwagę
            dostepni = [p for p in pracownicy if row[p] == 1]

            idealni = []
            do_ratunku = [] 

            for p in dostepni:
                # Blokada limitu dla wybranego pracownika
                if p == prac_ostatni and liczniki[p] >= limit_max:
                    continue

                # Podział na idealnych i tych, którzy potrzebują odpoczynku (po 2. zmianie)
                if ostatnia_zmiana_pracownika[p] == 2 and typ_zmiany == 1:
                    do_ratunku.append(p)
                else:
                    idealni.append(p)

            wybrani = []
            # Sortowanie po liczbie zmian zapewnia sprawiedliwy podział
            idealni_posortowani = sorted(idealni, key=lambda x: liczniki[x])
            wybrani.extend(idealni_posortowani[:2])

            # Jeśli brakuje osób do pary, dobieramy z grupy "do ratunku"
            if len(wybrani) < 2:
                brakujacych = 2 - len(wybrani)
                ratownicy_posortowani = sorted(do_ratunku, key=lambda x: liczniki[x])
                wybrani.extend(ratownicy_posortowani[:brakujacych])

            if wybrani:
                for w in wybrani:
                    liczniki[w] += 1
                    ostatnia_zmiana_pracownika[w] = typ_zmiany
                
                # Dodanie symbolu (!) dla osób pracujących bez pełnego odpoczynku
                final_names = [f"{w} (!)" if w in do_ratunku else w for w in wybrani]
                wynikowa_kolumna.append(", ".join(final_names))
            else:
                wynikowa_kolumna.append("BRAK OBSADY")

            # Resetujemy status odpoczynku dla osób, które w tej turze nie pracowały
            for p in pracownicy:
                if p not in wybrani:
                    ostatnia_zmiana_pracownika[p] = None

        df['PRZYPISANI_PRACOWNICY'] = wynikowa_kolumna

        # Prezentacja wyników
        st.divider()
        st.subheader("Podgląd wygenerowanego grafiku")
        st.caption("(!) Oznacza osobę dobraną awaryjnie (krótki odstęp między zmianami).")
        st.dataframe(df[['DATA', 'ZMIANA', 'PRZYPISANI_PRACOWNICY']], use_container_width=True)
        
        st.subheader("Suma zmian w miesiącu")
        st.bar_chart(pd.Series(liczniki))

        # Przygotowanie Excela do pobrania
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Grafik_Wynik')

        st.download_button(
            label="📥 Pobierz grafik jako Excel",
            data=output.getvalue(),
            file_name="gotowy_grafik.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
