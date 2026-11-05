import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, colorchooser, filedialog
import tkinter.font as tkfont
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import shutil
import os
import json
import re
import urllib.parse
import urllib.request
import urllib.error
import ssl
import base64
import uuid
import webbrowser
import sys
import subprocess
import tempfile
from datetime import datetime

# --- Dawny, wlasny backend Freyonatora (Supabase) - uzywany TYLKO do
# sprawdzania nowych wersji programu (patrz _sprawdz_aktualizacje). Program
# NIE jest juz aktywowany kodem licencyjnym przez ten projekt - patrz
# "Aktywacja programu przez konto KOS" nizej.
FREYA_SUPABASE_URL = "https://dlffxncncnnyufdsgnxs.supabase.co"
FREYA_SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRsZmZ4bmNuY25ueXVmZHNnbnhzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODg5NzM0NjIsImV4cCI6MjEwNDU0OTQ2Mn0.mZVeHOQJvIpgpesqcmYqjBPrz6X9FEQf3sfJBMj1Rxo"
APP_WERSJA = "1.0.0"

# --- Aktywacja programu przez konto KOS (fv-strona-www) ---
# Program nie uzywa wlasnego kodu licencyjnego - przy starcie loguje sie
# (login/e-mail + haslo) na to samo konto, ktorym dana osoba loguje sie na
# stronie KOS, i sprawdza TAM plan firmy. Dziala tylko dla firm na planie
# Ultra. Uzywa publicznego "publishable" klucza Supabase (bezpieczny do
# wpisania w kodzie klienckim - to samo robi strona KOS w przegladarce) wraz
# z sesja (access/refresh token) zwrocona po zalogowaniu - wiec odczyt
# company_users/companies dziala dzieki tym samym politykom RLS, co na
# stronie, a NIE dzieki temu kluczowi.
KOS_SUPABASE_URL = "https://utsbtupupxwrbtdbukot.supabase.co"
KOS_SUPABASE_ANON_KEY = "sb_publishable_Kqm0H2XrcBQV2hE77yYSjw_gwLv7QtK"
KOS_API_URL = "https://kos-kawiarniaos.vercel.app"

# Minimalna dlugosc komentarza/raportu zmianowego (zakladka "Rozliczenie
# Dnia") - wymagana zarowno przy zapisie dnia do KOS (zapisz_excel), jak i
# przy wysylce raportu mailem (wyslij_raport_email). Komentarz jest teraz
# zapisywany w KOS (freyonator_dzien_weryfikacja.komentarz) - jest wiec
# widoczny takze na stronie, nie tylko w mailu.
MIN_DLUGOSC_KOMENTARZA = 67

class KawiarniaOSApp:
    def __init__(self, root):
        self.root = root
        
        self.plik_ustawien = "ustawienia_freya.json"
        self.szablon_plik = "szablon.xlsx"
        self.format_nazwy_pliku = "Rozliczenie_{miesiac}.xlsx"
        self.format_nazwy_obecnosci = "Obecnosci_{rok}.xlsx"
        self.wykluczeni_napiwki = "Kuba"
        self.pracownicy_wykluczeni_kos = set()  # imiona (lowercase) z flaga wykluczony_z_napiwkow=True w KOS
        self.nazwa_programu = "KawiarniaOS"
        self.haslo_dostepu = ""
        self.szerokosc_okna = 1280
        self.wysokosc_okna = 850
        self.klucz_kuchnia = "kuchnia"
        self.klucz_bar = "bar"
        self.klucz_admin = "admin"
        self.pracownicy_kuchnia = ""
        self.pracownicy_bar = ""
        self.pracownicy_kelner = ""
        self.pracownicy_admin = ""
        self.wiersze_na_dzien = 12
        self.adresy_email = ""
        # Aktywacja przez konto KOS (patrz _wymagaj_aktywacji_kos) - zastapila
        # dawny kod licencyjny.
        self.kos_login = ""
        self.kos_refresh_token = ""
        self.kos_access_token = ""  # tylko w pamieci, NIE zapisywane do pliku ustawien
        self.kos_company_id = ""
        self.kos_company_nazwa = ""
        self.kos_plan = ""
        self.ostatnia_udana_aktywacja = ""
        self.folder_zapisu = ""
        self.folder_zapisu_obecnosci = ""
        
        self.kolor_tla = "#10182B"
        self.kolor_czcionki = "#F5F7FA"
        self.karta_bg = "#161F36"
        self.karta_border = "#283355"
        self.tekst_muted = "#9AA5C0"
        self.kolor_sukces = "#4ADE80"
        self.kolor_blad = "#F87171"

        # --- Pola wejsciowe na kartach: wczesniej czesc pol miala jaskrawe,
        # jasne tla (bialy/pastelowy) z czarnym tekstem - kontrastowalo to
        # ostro z ciemnym motywem calej aplikacji. Teraz kazde pole to
        # stonowany, zabarwiony odcien tla karty + jasny tekst, zachowujac to
        # samo znaczenie kolorow (czerwony=recznie wpisywane/ostrzezenie,
        # niebieski=wyliczane automatycznie, zielony=potwierdzone/sukces,
        # bursztynowy=dane zewnetrzne z kasy fiskalnej).
        self.pole_bg = self._rozjasnij_kolor(self.karta_bg, 7)
        self.pole_fg = self.kolor_czcionki
        self.pole_ostrzezenie_bg = self._zabarw(self.karta_bg, "#F87171", 0.24)
        self.pole_ostrzezenie_fg = "#FCA5A5"
        self.pole_info_bg = self._zabarw(self.karta_bg, "#60A5FA", 0.24)
        self.pole_info_fg = "#93C5FD"
        self.pole_sukces_bg = self._zabarw(self.karta_bg, "#4ADE80", 0.22)
        self.pole_sukces_fg = "#86EFAC"
        self.pole_uwaga_bg = self._zabarw(self.karta_bg, "#FBBF24", 0.22)
        self.pole_uwaga_fg = "#FCD34D"

        # --- Nowoczesny font: najlepszy dostepny bezszeryfowy font systemowy,
        # z fallbackiem na Helvetica. Ustawiane raz tutaj, uzywane wszedzie w GUI
        # (zamiast rozsianych po calym pliku sztywnych literalow self.czcionka_nazwa).
        try:
            _dostepne_fonty = set(tkfont.families())
        except Exception:
            _dostepne_fonty = set()
        self.czcionka_nazwa = next(
            (f for f in ("SF Pro Text", "Inter", "Helvetica Neue", "Segoe UI", "Helvetica", "Arial")
             if f in _dostepne_fonty),
            "Helvetica",
        )

        # --- Jeden spojny kolor akcentujacy dla glownych akcji (zapisz / dodaj),
        # jeden kolor dla akcji niszczacych (usun) i jeden stonowany styl dla
        # reszty przyciskow pomocniczych - zamiast kilkunastu konkurujacych barw.
        self.kolor_akcent = "#16a34a"
        self.kolor_akcent_fg = "#f0fdf4"
        self.kolor_niebezpieczny = "#dc2626"
        self.kolor_niebezpieczny_fg = "#fef2f2"
        self.kolor_wygaszony_bg = self._rozjasnij_kolor(self.karta_bg, 14)
        self.kolor_wygaszony_fg = "#C7D0E6"

        self.kol_vip = "VIP/STAFF\n/STRATY"
        self.kol_kmob = "Karta Mob."
        self.kol_kkasa = "Karta Kasa"
        self.kol_przelewy = "Przelewy"
        self.kol_foodsi = "Foodsi"
        self.kol_serwis = "Serwis / Tip"
        self.kol_razem = "Suma Płatności"
        self.kol_gotowka = "Gotówka (Oczek.)"
        self.kol_sprawdz = "Zdano (Wpłata)"
        self.kol_roznica = "Różnica"

        self.txt_wybor_daty = "Wybór Daty (Edycja / Nowy Dzień)"
        self.txt_wprowadzanie = "Raporty Osobowe"
        self.txt_kalkulator = "Kalkulator Napiwków (Serwis dzieli się na wszystkich)"
        self.txt_weryfikacja = "Weryfikacja z Kasą i Terminalami"
        self.txt_zapisz_btn = "ZAPISZ DZIEŃ W KOS"
        self.txt_wyslij_btn = "WYŚLIJ RAPORT (GMAIL)"
        self.txt_wczytaj_btn = "WCZYTAJ Z KOS (Edycja)"
        self.txt_wyczysc_btn = "WYCZYŚĆ FORMULARZ"
        self.txt_kto_rozlicza = "Rozlicza (WYMAGANE):"
        self.txt_komentarz = f"Komentarz / Raport zmianowy (min. {MIN_DLUGOSC_KOMENTARZA} znaków):"
        self.txt_podziel_btn = "Podziel i Przypisz"
        self.txt_dodaj_kelnera_btn = "+ Dodaj Kelnera"
        self.txt_usun_wiersz_btn = "- Usuń"
        self.txt_oblicz_gotowka_btn = "Oblicz podział"
        self.txt_wczytaj_obc_btn = "WCZYTAJ DZIEŃ"
        self.txt_wyczysc_obc_btn = "WYCZYŚĆ"
        self.txt_dodaj_obc_btn = "+ Dodaj Pracownika"
        self.txt_zapisz_obc_btn = "ZAPISZ OBECNOŚCI DNIA"
        self.txt_kalkulator_gotowka_naglowek = "Podział gotówki (tylko podgląd, NIE zapisuje do Excela)"
        self.txt_wybor_daty_obc = "Wybór Daty (Obecności)"
        self.txt_godziny_naglowek = "Godziny Pracy (godziny,minuty - np. 19:45 albo 19,45)"
        self.txt_podsumowanie_obc_naglowek = "Godziny pracy i należny serwis (netto) — pracownicy"
        
        self._wczytaj_wskaznik_folderu()
        pierwsze_uruchomienie = not os.path.exists(self.sciezka(self.plik_ustawien))
        self.wczytaj_ustawienia()
        if pierwsze_uruchomienie:
            self._pierwsze_uruchomienie_wybierz_folder()
        self._przygotuj_dane_startowe()

        if not self._wymagaj_aktywacji_kos():
            self.root.destroy()
            return
        self._sprawdz_aktualizacje()
        self._sprawdz_i_zainstaluj_aktualizacje_windows()
        
        self.root.title(self.nazwa_programu)
        _ekran_szer = self.root.winfo_screenwidth()
        _ekran_wys = self.root.winfo_screenheight()
        self.szerokosc_okna = max(self.szerokosc_okna, min(_ekran_szer - 50, 1680))
        self.wysokosc_okna = max(self.wysokosc_okna, min(_ekran_wys - 90, 1040))
        _poz_x = max(0, (_ekran_szer - self.szerokosc_okna) // 2)
        _poz_y = max(0, (_ekran_wys - self.wysokosc_okna) // 2 - 10)
        self.root.geometry(f"{self.szerokosc_okna}x{self.wysokosc_okna}+{_poz_x}+{_poz_y}")
        self.root.configure(bg=self.kolor_tla)
        self._ustaw_ikone_okna()
        
        self.slownik_miesiecy = {
            "01": "STYCZEŃ", "02": "LUTY", "03": "MARZEC", "04": "KWIECIEŃ",
            "05": "MAJ", "06": "CZERWIEC", "07": "LIPIEC", "08": "SIERPIEŃ",
            "09": "WRZESIEŃ", "10": "PAŹDZIERNIK", "11": "LISTOPAD", "12": "GRUDZIEŃ"
        }
        
        self.wiersze_danych = []
        self.max_wierszy = self.wiersze_na_dzien
        self.lista_kelnerow = []

        # Kasa: Wydatki i Wplaty (wydatki gotowkowe z NIP + wplaty do banku,
        # patrz buduj_zakladke_kasa / app/api/freyonator/kasa-operacje)
        self.wiersze_kasa = []
        self._kasa_typy_opis = ["Wydatek gotówkowy (paragon z NIP)", "Wpłata do banku"]
        self._kasa_typ_backend = {
            "Wydatek gotówkowy (paragon z NIP)": "wydatek",
            "Wpłata do banku": "wplata_bank",
        }

        self.buduj_gui()
        self.odswiez_liste_pracownikow()
        self._kos_wczytaj_kasa()

    # ---------- Nowoczesny wyglad: zaokraglone karty i przyciski (liquid glass) ----------
    def _zaokraglony_ksztalt(self, canvas, x1, y1, x2, y2, promien, **kwargs):
        promien = max(0, min(promien, (x2 - x1) / 2, (y2 - y1) / 2))
        pts = [
            x1 + promien, y1, x2 - promien, y1, x2, y1, x2, y1 + promien,
            x2, y2 - promien, x2, y2, x2 - promien, y2, x1 + promien, y2,
            x1, y2, x1, y2 - promien, x1, y1 + promien, x1, y1,
        ]
        return canvas.create_polygon(pts, smooth=True, **kwargs)

    def _rozjasnij_kolor(self, kolor_hex, ile=18):
        try:
            k = kolor_hex.lstrip("#")
            r, g, b = (int(k[i:i + 2], 16) for i in (0, 2, 4))
            r, g, b = (min(255, c + ile) for c in (r, g, b))
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return kolor_hex

    def _przyciemnij_kolor(self, kolor_hex, ile=18):
        try:
            k = kolor_hex.lstrip("#")
            r, g, b = (int(k[i:i + 2], 16) for i in (0, 2, 4))
            r, g, b = (max(0, c - ile) for c in (r, g, b))
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return kolor_hex

    def _zabarw(self, baza_hex, kolor_hex, sila=0.22):
        """Miesza kolor bazowy (tlo karty) z kolorem akcentu w podanej proporcji -
        daje przyciemniony, stonowany odcien pasujacy do ciemnego motywu (np. pole
        VIP/Straty: tlo karty lekko zabarwione na czerwono), zamiast jaskrawego
        jasnego tla kontrastujacego z reszta ciemnej aplikacji."""
        try:
            b = baza_hex.lstrip("#"); k = kolor_hex.lstrip("#")
            br, bgc, bb = (int(b[i:i + 2], 16) for i in (0, 2, 4))
            kr, kgc, kb = (int(k[i:i + 2], 16) for i in (0, 2, 4))
            r = int(br + (kr - br) * sila)
            g = int(bgc + (kgc - bgc) * sila)
            bl = int(bb + (kb - bb) * sila)
            return f"#{r:02x}{g:02x}{bl:02x}"
        except Exception:
            return baza_hex

    def _toast(self, tekst, typ="sukces", czas_ms=2600):
        """Nienachalny komunikat znikajacy samoczynnie (zamiast blokujacego
        messagebox z przyciskiem OK) - uzywany np. po udanym zapisie dnia,
        zeby nie przerywac pracy dodatkowym klikiem. typ: 'sukces' | 'blad'."""
        tlo = self.kolor_sukces if typ == "sukces" else self.kolor_blad
        fg = "#052e16" if typ == "sukces" else "#450a0a"
        try:
            ramka = tk.Frame(self.root, bg=tlo, highlightthickness=0)
            tk.Label(ramka, text=tekst, font=(self.czcionka_nazwa, 11, "bold"), bg=tlo, fg=fg, padx=18, pady=10).pack()
            ramka.place(relx=0.5, rely=0.96, anchor="s")
            ramka.lift()
            self.root.after(czas_ms, ramka.destroy)
        except Exception:
            pass

    def zrob_karte(self, rodzic, tytul, promien=16, tlo_karty=None, kolor_ramki=None):
        """Nowoczesna karta z zaokraglonymi rogami (Canvas) - zamiennik tk.LabelFrame.
        Tytul "plywa" nad karta wprost na tle strony (bez wlasnego pudelka/tla za
        slowami) - plaski, wspolczesny uklad zamiast starego paska naglowka
        wtopionego w kolor karty. Pod karta dorysowywany jest delikatny,
        rozmyty cien (2 przesuniete warstwy coraz ciemniejsze) dla wrazenia
        glebi, a sama karta ma wiecej "oddechu" (wieksze marginesy wewnetrzne).
        Zwraca (zewnetrzna_ramka_do_spakowania, wewnetrzna_ramka_na_tresc)."""
        tlo_karty = tlo_karty or self.karta_bg
        kolor_ramki = kolor_ramki or self.karta_border
        tlo_strony = self.kolor_tla
        tytul_h = 28
        cien_1 = self._przyciemnij_kolor(tlo_strony, 5)
        cien_2 = self._przyciemnij_kolor(tlo_strony, 9)
        zewn = tk.Frame(rodzic, bg=tlo_strony)
        platno = tk.Canvas(zewn, highlightthickness=0, bd=0, bg=tlo_strony)
        platno.place(x=0, y=0, relwidth=1, relheight=1)
        wewn = tk.Frame(zewn, bg=tlo_karty)
        wewn.pack(fill="both", expand=True, padx=22, pady=(tytul_h + 20, 20))

        def przerysuj(event=None):
            platno.delete("all")
            w, h = zewn.winfo_width(), zewn.winfo_height()
            if w > 6 and h > 6 + tytul_h:
                # Cien: dwie coraz ciemniejsze, przesuniete w dol-prawo warstwy
                # pod kartą - tani (bez alpha-blendingu) ale czytelny substytut
                # miekkiego box-shadow.
                self._zaokraglony_ksztalt(platno, 4, tytul_h + 4, w + 2, h + 2, promien,
                                           fill=cien_1, outline="")
                self._zaokraglony_ksztalt(platno, 2, tytul_h + 2, w, h, promien,
                                           fill=cien_2, outline="")
                # Tytul rysowany NAD kartą, na tle strony - bez tla za slowami.
                platno.create_text(22, tytul_h - 10, text=tytul, anchor="w",
                                    fill=self.tekst_muted, font=(self.czcionka_nazwa, 11, "bold"))
                self._zaokraglony_ksztalt(platno, 1, tytul_h, w - 1, h - 1, promien,
                                           fill=tlo_karty, outline=kolor_ramki, width=1.3)
                platno.create_line(promien, tytul_h + 1, w - promien, tytul_h + 1,
                                    fill=self._rozjasnij_kolor(tlo_karty, 30), width=1)
            platno.tk.call("lower", platno._w)

        zewn.bind("<Configure>", przerysuj)
        return zewn, wewn

    def pill(self, rodzic, text="", command=None, bg="#2563eb", fg="#ffffff",
             font=None, height=None, width=None, outline=False, disabled=False, **_ign):
        """Przycisk w stylu 'pigulki' (zaokraglone rogi, polysk liquid-glass) - zamiennik tk.Button.
        outline=True -> styl drugorzedny: sam obrys w kolorze `bg` bez wypelnienia,
        zeby nie konkurowal wzrokowo z glownym kolorem akcentu (self.kolor_akcent).
        disabled=True -> przycisk od razu wygaszony i nieklikalny; stan mozna
        zmieniac pozniej przez zwrocony obiekt: przycisk.ustaw_aktywny(True/False).
        Widoczne sa 4 stany: hover, wcisniecie (active), focus (obwodka po Tab)
        i disabled (wygaszony, szary tekst, brak reakcji na klik/hover)."""
        czcionka = font or (self.czcionka_nazwa, 10, "bold")
        try:
            tf = tkfont.Font(family=czcionka[0], size=czcionka[1] if len(czcionka) > 1 else 10,
                              weight=czcionka[2] if len(czcionka) > 2 else "normal")
            szer_tekstu = tf.measure(text)
        except Exception:
            szer_tekstu = max(60, len(text) * 8)
        wys = 34 if not height else int(34 + (height - 1) * 20)
        szer = width if width else max(92, szer_tekstu + 42)
        try:
            tlo_rodzica = rodzic["bg"]
        except Exception:
            tlo_rodzica = self.kolor_tla

        platno = tk.Canvas(rodzic, width=szer, height=wys, highlightthickness=0, bd=0,
                            bg=tlo_rodzica, cursor="arrow" if disabled else "hand2", takefocus=1)
        stan = {"hover": False, "wcisniety": False, "focus": False, "aktywny": not disabled}

        def kolor_wypelnienia():
            if not stan["aktywny"]:
                return self.kolor_wygaszony_bg
            if outline:
                return tlo_rodzica
            if stan["wcisniety"]:
                return self._przyciemnij_kolor(bg, 28)
            if stan["hover"]:
                return self._rozjasnij_kolor(bg, 16)
            return bg

        def kolor_tekstu():
            if not stan["aktywny"]:
                return self.kolor_wygaszony_fg
            return bg if outline else fg

        def rysuj(event=None):
            platno.delete("all")
            w = platno.winfo_width() or szer
            h = platno.winfo_height() or wys
            r = h / 2
            k = kolor_wypelnienia()
            ma_obrys = outline and stan["aktywny"]
            self._zaokraglony_ksztalt(platno, 1, 1, w - 1, h - 1, r, fill=k,
                                       outline=bg if ma_obrys else "", width=1.6 if ma_obrys else 0)
            if not outline and stan["aktywny"]:
                platno.create_line(r * 0.6, 3, max(r * 0.6 + 1, w - r * 0.6), 3,
                                    fill=self._rozjasnij_kolor(k, 45), width=1)
            if stan["focus"] and stan["aktywny"]:
                self._zaokraglony_ksztalt(platno, 0, 0, w, h, r + 1, fill="",
                                           outline=self._rozjasnij_kolor(bg, 60), width=2)
            platno.create_text(w / 2, h / 2, text=text, fill=kolor_tekstu(), font=czcionka)

        def wywolaj():
            if stan["aktywny"] and command:
                command()

        def wcisniecie(e):
            if not stan["aktywny"]:
                return
            stan["wcisniety"] = True
            rysuj()

        def puszczenie(e):
            if not stan["aktywny"]:
                return
            byl = stan["wcisniety"]
            stan["wcisniety"] = False
            rysuj()
            if byl:
                wywolaj()

        def wejscie(e):
            if not stan["aktywny"]:
                return
            stan["hover"] = True
            rysuj()

        def wyjscie(e):
            stan["hover"] = False
            rysuj()

        def fokus_wszedl(e):
            stan["focus"] = True
            rysuj()

        def fokus_wyszedl(e):
            stan["focus"] = False
            rysuj()

        def klawisz(e):
            wywolaj()

        def ustaw_aktywny(aktywny):
            stan["aktywny"] = bool(aktywny)
            platno.config(cursor="hand2" if stan["aktywny"] else "arrow")
            rysuj()

        platno.bind("<Configure>", rysuj)
        platno.bind("<ButtonPress-1>", wcisniecie)
        platno.bind("<ButtonRelease-1>", puszczenie)
        platno.bind("<Enter>", wejscie)
        platno.bind("<Leave>", wyjscie)
        platno.bind("<FocusIn>", fokus_wszedl)
        platno.bind("<FocusOut>", fokus_wyszedl)
        platno.bind("<Return>", klawisz)
        platno.bind("<space>", klawisz)
        platno.ustaw_aktywny = ustaw_aktywny
        return platno

    def _sciezka_wskaznika_folderu(self):
        """Stała lokalizacja (niezależna od wybranego folderu zapisu), w której
        program zapisuje INFORMACJĘ o tym, jaki folder zapisu wybrał użytkownik.
        Dzięki temu program potrafi odnaleźć własny folder zapisu przy starcie,
        zanim jeszcze wczyta z niego plik ustawień."""
        baza = os.path.join(os.path.expanduser("~"), "Documents", "KawiarniaOS")
        try:
            os.makedirs(baza, exist_ok=True)
        except Exception:
            pass
        return os.path.join(baza, "folder_wskaznik.txt")

    def _wczytaj_wskaznik_folderu(self):
        try:
            p = self._sciezka_wskaznika_folderu()
            if os.path.exists(p):
                with open(p, "r", encoding="utf-8") as f:
                    wartosc = f.read().strip()
                if wartosc and os.path.isdir(wartosc):
                    self.folder_zapisu = wartosc
        except Exception:
            pass

    def _zapisz_wskaznik_folderu(self, folder):
        try:
            p = self._sciezka_wskaznika_folderu()
            folder = (folder or "").strip()
            if folder:
                with open(p, "w", encoding="utf-8") as f:
                    f.write(folder)
            elif os.path.exists(p):
                os.remove(p)
        except Exception:
            pass

    def _pierwsze_uruchomienie_wybierz_folder(self):
        """Przy pierwszym uruchomieniu na nowym komputerze pyta, gdzie mają być
        zapisywane pliki programu (Excel, ustawienia) - domyślnie ~/Documents/KawiarniaOS,
        ale można od razu wskazać inny folder (np. Dropbox/iCloud jako kopia zapasowa)."""
        domyslny = self._katalog_danych()
        try:
            wlasny = messagebox.askyesno(
                "Folder zapisu plików",
                "Witaj w KawiarniaOS!\n\n"
                f"Domyślnie pliki programu (Excel, ustawienia) będą zapisywane w:\n{domyslny}\n\n"
                "Czy chcesz teraz wskazać inny folder (np. w Dropbox lub iCloud, żeby mieć automatyczną kopię zapasową)?",
                parent=self.root,
            )
        except Exception:
            wlasny = False
        if wlasny:
            wybrany = filedialog.askdirectory(title="Wybierz folder zapisu plików KawiarniaOS", parent=self.root)
            if wybrany:
                self.folder_zapisu = wybrany
                self._zapisz_wskaznik_folderu(wybrany)
        # Pytanie o folder obecności
        try:
            wlasny_obc = messagebox.askyesno(
                "Folder pliku Obecności",
                "Czy plik Obecności (lista godzin pracy) ma być zapisywany w INNYM folderze niż pliki rozliczeń?\n\n"
                "(np. współdzielony folder na dysku sieciowym lub inny komputer)",
                parent=self.root,
            )
        except Exception:
            wlasny_obc = False
        if wlasny_obc:
            wybrany_obc = filedialog.askdirectory(title="Wybierz folder zapisu pliku Obecności", parent=self.root)
            if wybrany_obc:
                self.folder_zapisu_obecnosci = wybrany_obc

    def _katalog_danych(self):
        """Folder z danymi programu (Excel, ustawienia) na nowym komputerze:
        domyślnie ~/Documents/KawiarniaOS. Jeśli program działa w folderze,
        w którym pliki już są (dotychczasowa instalacja), ten folder ma priorytet
        - patrz sciezka().
        """
        baza = os.path.join(os.path.expanduser("~"), "Documents", "KawiarniaOS")
        try: os.makedirs(baza, exist_ok=True)
        except Exception: pass
        return baza

    def sciezka(self, nazwa_pliku):
        """Zwraca pełną ścieżkę do pliku danych (Excel/ustawienia). Jeśli plik
        istnieje w bieżącym katalogu (dotychczasowy sposób działania programu),
        używa go bez zmian. W przeciwnym razie, jeśli w Ustawieniach wybrano własny
        folder zapisu, używa go. Inaczej używa domyślnego folderu z danymi programu
        (~/Documents/KawiarniaOS) - ważne przy pierwszej instalacji na nowym
        komputerze, gdzie taki plik jeszcze nie istnieje w folderze programu.
        """
        if os.path.exists(nazwa_pliku):
            return nazwa_pliku
        if getattr(self, "folder_zapisu", ""):
            try:
                os.makedirs(self.folder_zapisu, exist_ok=True)
            except Exception:
                pass
            return os.path.join(self.folder_zapisu, nazwa_pliku)
        return os.path.join(self._katalog_danych(), nazwa_pliku)

    def _sciezka_wbudowanego_szablonu(self):
        """Ścieżka do szablon.xlsx dołączonego do aplikacji (przy budowaniu .app/.dmg
        ten plik jest częścią programu i służy jako wzorzec przy pierwszym uruchomieniu)."""
        if getattr(sys, "frozen", False):
            return os.path.join(getattr(sys, "_MEIPASS", os.path.dirname(sys.executable)), "szablon.xlsx")
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "szablon.xlsx")

    def _przygotuj_dane_startowe(self):
        """Przy pierwszym uruchomieniu na nowym komputerze: jeśli nigdzie nie ma
        jeszcze szablonu Excela, kopiuje wzorcowy szablon.xlsx dołączony do programu
        do folderu z danymi (~/Documents/KawiarniaOS)."""
        try:
            docelowy = self.sciezka(self.szablon_plik)
            if not os.path.exists(docelowy):
                wbudowany = self._sciezka_wbudowanego_szablonu()
                if os.path.exists(wbudowany):
                    os.makedirs(os.path.dirname(docelowy), exist_ok=True)
                    shutil.copy(wbudowany, docelowy)
        except Exception:
            pass

    def _id_komputera(self):
        """Stabilne, losowe ID tego komputera/instalacji (nie sprzetowe) - zapisane
        raz w folderze z danymi programu. Uzywane do przypisania licencji do 1 komputera."""
        plik_id = self.sciezka("id_komputera.txt")
        try:
            if os.path.exists(plik_id):
                with open(plik_id, "r", encoding="utf-8") as f:
                    zawartosc = f.read().strip()
                    if zawartosc:
                        return zawartosc
        except Exception:
            pass
        nowy_id = str(uuid.uuid4())
        try:
            with open(plik_id, "w", encoding="utf-8") as f:
                f.write(nowy_id)
        except Exception:
            pass
        return nowy_id

    def _kontekst_ssl(self):
        """Kontekst SSL do polaczen HTTPS z Supabase. Doklada certyfikaty z pakietu
        certifi (jesli jest dostepny) do domyslnego magazynu systemowego - to
        naprawia typowy problem na macOS, gdzie Python (np. z instalatora python.org)
        nie ma wgranych certyfikatow root i polaczenia HTTPS zawodza z bledem SSL."""
        if getattr(self, "_ssl_ctx_cache", None) is not None:
            return self._ssl_ctx_cache
        ctx = None
        try:
            ctx = ssl.create_default_context()
        except Exception:
            ctx = None
        if ctx is not None:
            try:
                import certifi
                ctx.load_verify_locations(certifi.where())
            except Exception:
                pass
        self._ssl_ctx_cache = ctx
        return ctx

    def _log_siec(self, tekst):
        """Zapisuje krotka linie diagnostyczna o bledach polaczenia z serwerem do
        pliku log_sieci.txt w folderze danych programu - przydatne przy zdalnej
        pomocy, zeby nie zgadywac co poszlo nie tak (np. blad certyfikatu SSL vs
        brak internetu)."""
        try:
            with open(self.sciezka("log_sieci.txt"), "a", encoding="utf-8") as f:
                f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  {tekst}\n")
        except Exception:
            pass

    def _kos_http(self, method, url, body=None, headers=None):
        """Pomocnicze wywolanie HTTP do KOS/Supabase. Zwraca (status, dane) - status
        to kod HTTP (nawet dla bledow typu 400/401/403, bo serwer wtedy ODPOWIEDZIAL),
        albo (None, None) gdy polaczenie sie nie udalo wgl (np. brak internetu) -
        to rozroznienie jest wazne dla tolerancji offline w _wymagaj_aktywacji_kos."""
        try:
            dane_wyj = json.dumps(body).encode("utf-8") if body is not None else None
            req = urllib.request.Request(url, data=dane_wyj, headers=headers or {}, method=method)
            with urllib.request.urlopen(req, timeout=8, context=self._kontekst_ssl()) as resp:
                surowe = resp.read().decode("utf-8")
                return resp.status, (json.loads(surowe) if surowe else None)
        except urllib.error.HTTPError as e:
            try:
                surowe = e.read().decode("utf-8")
                return e.code, (json.loads(surowe) if surowe else None)
            except Exception:
                return e.code, None
        except Exception as e:
            self._ostatni_blad_aktywacji = f"{type(e).__name__}: {e}"
            self._log_siec(f"kos_http: {self._ostatni_blad_aktywacji}")
            return None, None

    def _kos_resolve_login_email(self, login):
        """Zamienia login/nazwe uzytkownika KOS na jego e-mail (tak jak robi to
        strona KOS przy logowaniu) - zwraca e-mail albo None."""
        status, dane = self._kos_http(
            "POST",
            f"{KOS_SUPABASE_URL}/rest/v1/rpc/resolve_login_email",
            body={"p_login": login},
            headers={
                "Content-Type": "application/json",
                "apikey": KOS_SUPABASE_ANON_KEY,
                "Authorization": f"Bearer {KOS_SUPABASE_ANON_KEY}",
            },
        )
        if status == 200 and isinstance(dane, str) and dane.strip():
            return dane.strip()
        return None

    def _kos_zaloguj(self, email, haslo):
        """Loguje na konto KOS (Supabase Auth, email+haslo - identycznie jak na
        stronie). Zwraca dict {access_token, refresh_token, user_id} albo None."""
        status, dane = self._kos_http(
            "POST",
            f"{KOS_SUPABASE_URL}/auth/v1/token?grant_type=password",
            body={"email": email, "password": haslo},
            headers={"Content-Type": "application/json", "apikey": KOS_SUPABASE_ANON_KEY},
        )
        self._ostatni_status_aktywacji = status
        if status == 200 and isinstance(dane, dict) and dane.get("access_token"):
            return {
                "access_token": dane.get("access_token"),
                "refresh_token": dane.get("refresh_token", ""),
                "user_id": (dane.get("user") or {}).get("id", ""),
            }
        if isinstance(dane, dict):
            self._ostatni_blad_aktywacji = dane.get("error_description") or dane.get("msg") or dane.get("error") or ""
        return None

    def _kos_odswiez_sesje(self, refresh_token):
        """Odswieza sesje KOS bez pytania o haslo (uzywane przy kazdym starcie
        programu, gdy mamy juz zapisany refresh_token). Zwraca dict jak
        _kos_zaloguj albo None."""
        status, dane = self._kos_http(
            "POST",
            f"{KOS_SUPABASE_URL}/auth/v1/token?grant_type=refresh_token",
            body={"refresh_token": refresh_token},
            headers={"Content-Type": "application/json", "apikey": KOS_SUPABASE_ANON_KEY},
        )
        self._ostatni_status_aktywacji = status
        if status == 200 and isinstance(dane, dict) and dane.get("access_token"):
            return {
                "access_token": dane.get("access_token"),
                "refresh_token": dane.get("refresh_token", "") or refresh_token,
                "user_id": (dane.get("user") or {}).get("id", ""),
            }
        return None

    def _kos_pobierz_firme(self, access_token, user_id):
        """Odczytuje firme (company_id/nazwa/plan) zalogowanego uzytkownika KOS -
        identycznie jak robi to strona KOS po zalogowaniu (company_users ->
        companies). Zwraca dict {company_id, nazwa, plan, status_konta} albo None."""
        naglowki = {"apikey": KOS_SUPABASE_ANON_KEY, "Authorization": f"Bearer {access_token}"}
        status, wiersze = self._kos_http(
            "GET",
            f"{KOS_SUPABASE_URL}/rest/v1/company_users?select=company_id,status&user_id=eq.{user_id}",
            headers=naglowki,
        )
        self._ostatni_status_aktywacji = status
        if status != 200 or not isinstance(wiersze, list) or not wiersze:
            return None
        wiersz = wiersze[0]
        company_id = wiersz.get("company_id")
        status_konta = wiersz.get("status")
        if not company_id:
            return None
        status2, firmy = self._kos_http(
            "GET",
            f"{KOS_SUPABASE_URL}/rest/v1/companies?select=name,plan&id=eq.{company_id}",
            headers=naglowki,
        )
        if status2 != 200 or not isinstance(firmy, list) or not firmy:
            return None
        firma = firmy[0]
        return {
            "company_id": company_id,
            "nazwa": firma.get("name") or "",
            "plan": firma.get("plan") or "",
            "status_konta": status_konta or "",
        }

    def _wymagaj_aktywacji_kos(self):
        """Aktywuje program logowaniem na konto KOS (strona kos-kawiarniaos.vercel.app)
        i sprawdzeniem planu firmy TAM - zamiast dawnego kodu licencyjnego. Program
        dziala tylko dla firm na planie Ultra. Zwraca True, gdy program moze dzialac
        dalej. Przy braku internetu (nie: przy odrzuceniu przez serwer) daje 7 dni
        tolerancji od ostatniej udanej aktywacji, zeby brak sieci nie blokowal pracy
        w kawiarni."""
        self._ostatni_blad_aktywacji = ""
        self._ostatni_status_aktywacji = None

        # 1) Mamy juz zapisana sesje - sprobuj odswiezyc w tle, bez pytania o
        #    login/haslo ponownie.
        if self.kos_refresh_token:
            sesja = self._kos_odswiez_sesje(self.kos_refresh_token)
            status_odswiezania = self._ostatni_status_aktywacji
            if sesja and sesja.get("access_token"):
                firma = self._kos_pobierz_firme(sesja["access_token"], sesja["user_id"])
                if firma and firma.get("status_konta") == "active" and firma.get("plan") == "ultra":
                    self.kos_access_token = sesja["access_token"]
                    self.kos_refresh_token = sesja.get("refresh_token") or self.kos_refresh_token
                    self.kos_company_id = firma["company_id"]
                    self.kos_company_nazwa = firma["nazwa"]
                    self.kos_plan = firma["plan"]
                    self.ostatnia_udana_aktywacja = datetime.now().strftime("%Y-%m-%d")
                    self._zapisz_json_ustawien()
                    return True
                if firma is not None:
                    # Serwer odpowiedzial - to realna odmowa (plan nie-Ultra,
                    # konto wylaczone...), NIE problem z siecia - bez tolerancji.
                    messagebox.showerror(
                        "Aktywacja KOS",
                        f"Firma „{firma.get('nazwa') or '—'}” ma plan „{firma.get('plan') or '—'}”.\n"
                        "Program wymaga planu Ultra w KOS. Skontaktuj się z administratorem KOS.",
                    )
                    return False
            elif status_odswiezania is not None:
                # Serwer odrzucil sesje (np. haslo zmienione gdzie indziej) -
                # przejdz nizej do ponownego logowania, bez tolerancji offline.
                pass
            else:
                # Brak polaczenia z serwerem wgl - tolerancja offline.
                if self.ostatnia_udana_aktywacja:
                    try:
                        dni_od = (datetime.now() - datetime.strptime(self.ostatnia_udana_aktywacja, "%Y-%m-%d")).days
                    except Exception:
                        dni_od = 9999
                    if dni_od <= 7:
                        return True
                szczegol = getattr(self, "_ostatni_blad_aktywacji", "")
                tresc = (
                    "Nie można połączyć się z KOS (sprawdź internet).\n"
                    "Program działa maksymalnie 7 dni bez połączenia od ostatniej udanej aktywacji.\n"
                    "Połącz się z internetem i uruchom program ponownie."
                )
                if szczegol:
                    tresc += f"\n\nSzczegóły techniczne: {szczegol}"
                messagebox.showerror("Aktywacja", tresc)
                return False

        # 2) Brak zapisanej / dzialajacej sesji - poproś o zalogowanie na konto KOS.
        login = simpledialog.askstring("Aktywacja KawiarniaOS", "Login lub e-mail konta KOS:", parent=self.root)
        if not login or not login.strip():
            messagebox.showerror("Aktywacja", "Bez zalogowania do konta KOS program nie może działać.")
            return False
        haslo = simpledialog.askstring("Aktywacja KawiarniaOS", "Hasło konta KOS:", parent=self.root, show="*")
        if not haslo:
            messagebox.showerror("Aktywacja", "Bez hasła program nie może działać.")
            return False

        email = self._kos_resolve_login_email(login.strip().lower())
        if not email:
            szczegol = getattr(self, "_ostatni_blad_aktywacji", "")
            tresc = "Nie znaleziono konta KOS dla podanego loginu/e-maila (albo brak internetu)."
            if szczegol:
                tresc += f"\n\nSzczegóły techniczne: {szczegol}"
            messagebox.showerror("Aktywacja", tresc)
            return False

        sesja = self._kos_zaloguj(email, haslo)
        if not sesja:
            szczegol = getattr(self, "_ostatni_blad_aktywacji", "")
            tresc = f"Nie udało się zalogować: {szczegol}" if szczegol else "Nieprawidłowy login lub hasło (albo brak internetu)."
            messagebox.showerror("Aktywacja", tresc)
            return False

        firma = self._kos_pobierz_firme(sesja["access_token"], sesja["user_id"])
        if not firma:
            messagebox.showerror("Aktywacja", "Zalogowano, ale nie udało się odczytać danych firmy z KOS (sprawdź internet).")
            return False
        if firma.get("status_konta") != "active":
            messagebox.showerror("Aktywacja", "To konto KOS nie jest aktywne.")
            return False
        if firma.get("plan") != "ultra":
            messagebox.showerror(
                "Aktywacja KOS",
                f"Firma „{firma.get('nazwa') or '—'}” ma plan „{firma.get('plan') or '—'}”.\n"
                "Program wymaga planu Ultra w KOS. Skontaktuj się z administratorem KOS.",
            )
            return False

        self.kos_login = email
        self.kos_access_token = sesja["access_token"]
        self.kos_refresh_token = sesja.get("refresh_token", "")
        self.kos_company_id = firma["company_id"]
        self.kos_company_nazwa = firma["nazwa"]
        self.kos_plan = firma["plan"]
        self.ostatnia_udana_aktywacja = datetime.now().strftime("%Y-%m-%d")
        self._zapisz_json_ustawien()
        return True

    def wyloguj_kos(self):
        """Usuwa zapisana sesje konta KOS. Program trzeba ponownie aktywowac
        (zalogowac) przy nastepnym uruchomieniu - przydatne przy zmianie
        firmy/konta na tym komputerze."""
        if not messagebox.askyesno(
            "Wyloguj z KOS",
            "Program trzeba będzie ponownie aktywować (zalogować) przy następnym uruchomieniu. Kontynuować?",
        ):
            return
        self.kos_login = ""
        self.kos_refresh_token = ""
        self.kos_access_token = ""
        self.kos_company_id = ""
        self.kos_company_nazwa = ""
        self.kos_plan = ""
        self.ostatnia_udana_aktywacja = ""
        self._zapisz_json_ustawien()
        if hasattr(self, "lbl_status_kos"):
            self.lbl_status_kos.config(text="Nie zalogowano.")
        messagebox.showinfo("Wylogowano", "Uruchom program ponownie, aby zalogować się na inne konto KOS.")

    def _sprawdz_aktualizacje(self):
        """Sprawdza online, czy jest nowsza wersja programu (nie blokuje pracy -
        tylko informuje, bez pobrania automatycznego). Cicho pomija przy braku internetu."""
        try:
            req = urllib.request.Request(
                f"{FREYA_SUPABASE_URL}/rest/v1/wersje?select=wersja,url_pobierania,opis&order=utworzono.desc&limit=1",
                headers={"apikey": FREYA_SUPABASE_ANON_KEY, "Authorization": f"Bearer {FREYA_SUPABASE_ANON_KEY}"},
                method="GET",
            )
            with urllib.request.urlopen(req, timeout=6, context=self._kontekst_ssl()) as resp:
                wynik = json.loads(resp.read().decode("utf-8"))
            if not (isinstance(wynik, list) and wynik):
                return
            najnowsza = wynik[0]
            def _jako_tuple(w):
                try:
                    return tuple(int(x) for x in str(w).strip().split("."))
                except Exception:
                    return (0,)
            if _jako_tuple(najnowsza.get("wersja")) > _jako_tuple(APP_WERSJA):
                url = najnowsza.get("url_pobierania") or ""
                opis = najnowsza.get("opis") or ""
                tresc = f"Dostępna jest nowsza wersja programu: {najnowsza.get('wersja')} (masz {APP_WERSJA}).\n\n{opis}"
                if url:
                    tresc += f"\n\nPobierz tutaj:\n{url}"
                messagebox.showinfo("Nowa wersja KawiarniaOS", tresc)
        except Exception as e:
            self._log_siec(f"sprawdz_aktualizacje: {type(e).__name__}: {e}")

    def _numer_buildu_lokalny(self):
        """Numer builda TEJ zainstalowanej kopii (tylko spakowana wersja Windows,
        patrz .github/workflows/build-windows-installer.yml - krok "Zapisz numer
        builda" zapisuje build_info.txt dolaczany do .exe). Brak pliku (np.
        uruchomienie z kodu zrodlowego) = brak automatycznej aktualizacji."""
        try:
            if getattr(sys, "frozen", False):
                sciezka_pliku = os.path.join(getattr(sys, "_MEIPASS", os.path.dirname(sys.executable)), "build_info.txt")
            else:
                sciezka_pliku = os.path.join(os.path.dirname(os.path.abspath(__file__)), "build_info.txt")
            with open(sciezka_pliku, "r", encoding="utf-8") as f:
                return int(f.read().strip())
        except Exception:
            return None

    def _sprawdz_i_zainstaluj_aktualizacje_windows(self):
        """Na Windows (spakowana wersja .exe) sprawdza najnowszy release na
        GitHub (uba-byte/kawiarniaos-desktop) i - za zgoda uzytkownika - sam
        pobiera najnowszy instalator i go uruchamia, po czym zamyka ten program
        (zeby instalator mogl nadpisac plik .exe). Dzieki temu pracownicy kawiarni
        NIE musza sami szukac strony/pobierac pliku za kazdym razem - wystarczy
        kliknac "Tak" w oknie. Nie blokuje pracy programu - cicho pomija przy
        braku internetu, braku numeru builda (np. kopia z kodu zrodlowego) albo
        na innym systemie niz Windows."""
        if sys.platform != "win32" or not getattr(sys, "frozen", False):
            return
        numer_lokalny = self._numer_buildu_lokalny()
        if numer_lokalny is None:
            return
        try:
            req = urllib.request.Request(
                "https://api.github.com/repos/uba-byte/kawiarniaos-desktop/releases/latest",
                headers={"Accept": "application/vnd.github+json"},
                method="GET",
            )
            with urllib.request.urlopen(req, timeout=6, context=self._kontekst_ssl()) as resp:
                dane = json.loads(resp.read().decode("utf-8"))
            tag = str(dane.get("tag_name") or "")
            if not tag.startswith("build-"):
                return
            numer_zdalny = int(tag.split("build-", 1)[1])
            if numer_zdalny <= numer_lokalny:
                return
            if not messagebox.askyesno(
                "Aktualizacja KawiarniaOS",
                f"Dostępna jest nowa wersja programu (build {numer_zdalny}, masz build {numer_lokalny}).\n\n"
                "Pobrać i zainstalować teraz? Program się zamknie, a instalator otworzy się automatycznie - "
                "wystarczy kliknąć „Dalej” kilka razy.",
            ):
                return
            url_pobierania = "https://github.com/uba-byte/kawiarniaos-desktop/releases/latest/download/KawiarniaOS-Setup.exe"
            tmp_sciezka = os.path.join(tempfile.gettempdir(), "KawiarniaOS-Setup.exe")
            req2 = urllib.request.Request(url_pobierania, method="GET")
            with urllib.request.urlopen(req2, timeout=90, context=self._kontekst_ssl()) as resp2:
                with open(tmp_sciezka, "wb") as f:
                    shutil.copyfileobj(resp2, f)
            subprocess.Popen([tmp_sciezka], shell=False)
            self.root.destroy()
            sys.exit(0)
        except Exception as e:
            self._log_siec(f"aktualizacja_windows: {type(e).__name__}: {e}")

    def wczytaj_ustawienia(self):
        if os.path.exists(self.sciezka(self.plik_ustawien)):
            try:
                with open(self.sciezka(self.plik_ustawien), "r", encoding="utf-8") as f:
                    dane = json.load(f)
                    for klucz, domyslna in vars(self).items():
                        if klucz in dane: setattr(self, klucz, dane[klucz])
            except Exception: pass

    def zapisz_ustawienia(self):
        self.szablon_plik = self.szablon_var.get()
        self.format_nazwy_pliku = self.format_nazwy_var.get()
        self.format_nazwy_obecnosci = self.format_nazwy_obecnosci_var.get()
        self.folder_zapisu = self.folder_zapisu_var.get()
        self._zapisz_wskaznik_folderu(self.folder_zapisu)
        self.wykluczeni_napiwki = self.wykluczeni_napiwki_var.get()
        self.nazwa_programu = self.nazwa_programu_var.get()
        self.haslo_dostepu = self.haslo_var.get()
        self.klucz_kuchnia = self.klucz_kuchnia_var.get()
        self.klucz_bar = self.klucz_bar_var.get()
        self.klucz_admin = self.klucz_admin_var.get()
        self.pracownicy_kuchnia = self.pracownicy_kuchnia_var.get()
        self.pracownicy_bar = self.pracownicy_bar_var.get()
        self.pracownicy_kelner = self.pracownicy_kelner_var.get()
        self.pracownicy_admin = self.pracownicy_admin_var.get()
        self.kolor_tla = self.kolor_tla_var.get()
        self.kolor_czcionki = self.kolor_czcionki_var.get()
        self.adresy_email = self.adresy_email_var.get()
        self.kol_vip = self.kol_vip_var.get()
        self.kol_kmob = self.kol_kmob_var.get()
        self.kol_kkasa = self.kol_kkasa_var.get()
        self.kol_przelewy = self.kol_przelewy_var.get()
        self.kol_foodsi = self.kol_foodsi_var.get()
        self.kol_serwis = self.kol_serwis_var.get()
        self.kol_razem = self.kol_razem_var.get()
        self.kol_gotowka = self.kol_gotowka_var.get()
        self.kol_sprawdz = self.kol_sprawdz_var.get()
        self.kol_roznica = self.kol_roznica_var.get()
        self.txt_wybor_daty = self.txt_wybor_daty_var.get()
        self.txt_wprowadzanie = self.txt_wprowadzanie_var.get()
        self.txt_kalkulator = self.txt_kalkulator_var.get()
        self.txt_weryfikacja = self.txt_weryfikacja_var.get()
        self.txt_zapisz_btn = self.txt_zapisz_btn_var.get()
        self.txt_wyslij_btn = self.txt_wyslij_btn_var.get()
        self.txt_wczytaj_btn = self.txt_wczytaj_btn_var.get()
        self.txt_wyczysc_btn = self.txt_wyczysc_btn_var.get()
        self.txt_kto_rozlicza = self.txt_kto_rozlicza_var.get()
        self.txt_komentarz = self.txt_komentarz_var.get()
        self.txt_podziel_btn = self.txt_podziel_btn_var.get()
        self.txt_dodaj_kelnera_btn = self.txt_dodaj_kelnera_btn_var.get()
        self.txt_usun_wiersz_btn = self.txt_usun_wiersz_btn_var.get()
        self.txt_oblicz_gotowka_btn = self.txt_oblicz_gotowka_btn_var.get()
        self.txt_wczytaj_obc_btn = self.txt_wczytaj_obc_btn_var.get()
        self.txt_wyczysc_obc_btn = self.txt_wyczysc_obc_btn_var.get()
        self.txt_dodaj_obc_btn = self.txt_dodaj_obc_btn_var.get()
        self.txt_zapisz_obc_btn = self.txt_zapisz_obc_btn_var.get()
        self.txt_kalkulator_gotowka_naglowek = self.txt_kalkulator_gotowka_naglowek_var.get()
        self.txt_wybor_daty_obc = self.txt_wybor_daty_obc_var.get()
        self.txt_godziny_naglowek = self.txt_godziny_naglowek_var.get()
        self.txt_podsumowanie_obc_naglowek = self.txt_podsumowanie_obc_naglowek_var.get()
        
        try:
            szer = int(self.szerokosc_var.get())
            wys = int(self.wysokosc_var.get())
            self.szerokosc_okna = szer
            self.wysokosc_okna = wys
            self.root.geometry(f"{szer}x{wys}")
        except ValueError:
            messagebox.showerror("Błąd", "Szerokość i wysokość okna muszą być liczbami!")
            return

        self.root.title(self.nazwa_programu)
        self.root.configure(bg=self.kolor_tla)

        if not self._zapisz_json_ustawien():
            return
        messagebox.showinfo("Zapisano", "Ustawienia zapisane!\nZmiany tekstów zostaną natychmiast odświeżone.")
        for widget in self.root.winfo_children(): widget.destroy()
        self.buduj_gui()
        self.odswiez_liste_pracownikow()
        self._kos_wczytaj_kasa()

    def _zapisz_json_ustawien(self):
        # BIAŁA LISTA KLUCZY DO ZAPISU (zapobiega błędom ComboBox is not JSON)
        klucze_bazowe = [
            "szablon_plik", "format_nazwy_pliku", "nazwa_programu", "haslo_dostepu",
            "szerokosc_okna", "wysokosc_okna", "klucz_kuchnia", "klucz_bar", "klucz_admin",
            "pracownicy_kuchnia", "pracownicy_bar", "pracownicy_kelner", "pracownicy_admin",
            "wiersze_na_dzien", "kolor_tla", "kolor_czcionki", "adresy_email",
            "format_nazwy_obecnosci", "wykluczeni_napiwki",
            "kos_login", "kos_refresh_token", "kos_company_id", "kos_company_nazwa", "kos_plan",
            "ostatnia_udana_aktywacja", "folder_zapisu", "folder_zapisu_obecnosci"
        ]
        dane_do_zapisu = {k: getattr(self, k) for k in klucze_bazowe if hasattr(self, k)}
        for k, v in vars(self).items():
            if (k.startswith("kol_") or k.startswith("txt_")) and isinstance(v, str):
                dane_do_zapisu[k] = v
        try:
            with open(self.sciezka(self.plik_ustawien), "w", encoding="utf-8") as f:
                json.dump(dane_do_zapisu, f, ensure_ascii=False, indent=4)
            return True
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się zapisać ustawień:\n{e}")
            return False

    def sprawdz_dostep(self, event):
        wybrana = self.notebook.index("current")
        # Jedyna zakladka za haslem to Ustawienia (ostatnia w notebooku) -
        # pozostale sa teraz otwarte, dane finansowe i tak zyja w KOS.
        idx_ustawien = self.notebook.index("end") - 1
        if wybrana == idx_ustawien and self.haslo_dostepu and not getattr(self, 'odblokowane', False):
            podane = simpledialog.askstring("Blokada dostępu", "Podaj hasło administratora:", show='*')
            if podane == self.haslo_dostepu: self.odblokowane = True 
            else:
                messagebox.showerror("Odmowa", "Nieprawidłowe hasło lub anulowano.")
                self.notebook.select(0) 

    def _on_focus_in(self, event):
        val = event.widget.get().strip()
        if val in ("0", "0.0", "0.00", "0,0", "0,00"): event.widget.delete(0, tk.END)

    def _on_focus_out(self, event):
        val = event.widget.get().strip()
        if not val: event.widget.insert(0, "0.0")

    def przypnij_czyszczenie_zer(self, pole):
        pole.bind("<FocusIn>", self._on_focus_in)
        pole.bind("<FocusOut>", self._on_focus_out)

    def _odswiez_licznik_komentarza(self, event=None):
        if not hasattr(self, "lbl_komentarz_licznik"):
            return
        dlugosc = len(self.komentarz_txt.get("1.0", tk.END).strip())
        wystarczy = dlugosc >= MIN_DLUGOSC_KOMENTARZA
        self.lbl_komentarz_licznik.config(
            text=f"{dlugosc} / {MIN_DLUGOSC_KOMENTARZA} znaków" + (" ✓" if wystarczy else ""),
            fg="#4ADE80" if wystarczy else "#F87171",
        )

    def _on_mousewheel(self, event):
        try:
            aktywna = self.notebook.index("current")
        except Exception:
            aktywna = 0
        mapa_canvas = {0: getattr(self, "canvas_obc", None), 1: getattr(self, "canvas", None), 2: getattr(self, "canvas_kasa", None), 3: getattr(self, "canvas_ust", None)}
        docelowy = mapa_canvas.get(aktywna)
        if docelowy is None: return
        if event.num == 4: docelowy.yview_scroll(-1, "units")
        elif event.num == 5: docelowy.yview_scroll(1, "units")
        else:
            # Krok NIEZALEZNY od wielkosci event.delta - liczy sie tylko
            # kierunek (+/-). Rzeczywista skala delty z gladzika Maca jest
            # nieprzewidywalna (zalezy od wersji Tk/Pythona i bywa zupelnie
            # inna niz na Windows, gdzie to zawsze wielokrotnosc 120), wiec
            # jakiekolwiek dzielenie/mnozenie przez stala dawalo albo brak
            # ruchu, albo szarpane, zbyt duze skoki. Przy przewijaniu dwoma
            # palcami te zdarzenia i tak przychodza bardzo czesto, wiec stały
            # krok 2 jednostek daje plynny, przewidywalny scroll.
            if event.delta > 0:
                docelowy.yview_scroll(-2, "units")
            elif event.delta < 0:
                docelowy.yview_scroll(2, "units")

    def wybierz_kolor_tla(self):
        kolor = colorchooser.askcolor(initialcolor=self.kolor_tla)[1]
        if kolor: self.kolor_tla_var.set(kolor); self.lbl_wybrany_kolor_tla.config(text=kolor, bg=kolor)
            
    def wybierz_kolor_czcionki(self):
        kolor = colorchooser.askcolor(initialcolor=self.kolor_czcionki)[1]
        if kolor: self.kolor_czcionki_var.set(kolor); self.lbl_wybrany_kolor_czcionki.config(text=kolor, fg=kolor)

    def _rozbij_nr_nazwe(self, n):
        """Rozdziela tekst pola pracownika na (nr, nazwa). Pole bywa w dwoch
        formatach: "17 - Jan Kowalski" (stara lista z lokalnego Excela, z
        numerem) albo po prostu "Jan Kowalski" (lista z KOS, bez numeru) - w
        tym drugim przypadku nazwisko MOZE zawierac wlasny myslnik (np.
        "Kowalska-Nowak"), wiec nie mozna po prostu dzielic po pierwszym "-".
        Numer rozpoznajemy tylko gdy tekst zaczyna sie od cyfr i myslnika -
        w kazdym innym przypadku caly tekst to nazwa."""
        n = (n or "").strip()
        m = re.match(r"^(\d+)\s*-\s*(.+)$", n)
        if m:
            return m.group(1).strip(), m.group(2).strip()
        return "", n

    def _lista_pracownikow_funkcja(self, tekst):
        return [n.strip().lower() for n in (tekst or "").split(",") if n.strip()]

    def jest_kuchnia(self, nazwa_pracownika):
        nazwa = (nazwa_pracownika or "").lower()
        lista = self._lista_pracownikow_funkcja(self.pracownicy_kuchnia)
        if lista:
            return any(n in nazwa for n in lista)
        return self.klucz_kuchnia.lower() in nazwa

    def jest_bar(self, nazwa_pracownika):
        nazwa = (nazwa_pracownika or "").lower()
        lista = self._lista_pracownikow_funkcja(self.pracownicy_bar)
        if lista:
            return any(n in nazwa for n in lista)
        return self.klucz_bar.lower() in nazwa

    def jest_admin(self, nazwa_pracownika):
        """Rola "Admin" - osoby ktore nie biora udzialu w podziale serwisu i
        nie pojawiaja sie na liscie raportow zmianowych (Rozliczenie Dnia).
        Patrz uzycie w podziel_serwis / oblicz_napiwki_gotowka /
        uzupelnij_kelnerow_z_obecnosci / _lista_kelnerow_bez_kuchni."""
        nazwa = (nazwa_pracownika or "").lower()
        lista = self._lista_pracownikow_funkcja(self.pracownicy_admin)
        if lista:
            return any(n in nazwa for n in lista)
        return self.klucz_admin.lower() in nazwa

    def _wiersz_to_kuchnia(self, nazwa_pracownika):
        """Rozpoznaje TEN KONKRETNY zbiorczy wiersz stanowiska "Kuchnia" w
        Rozliczeniu Dnia (np. "1 - Kuchnia") - PO ROWNOSCI ze slowem-kluczem
        (self.klucz_kuchnia), NIEZALEZNIE od listy konkretnych pracownikow
        przypisanych do roli Kuchnia w Ustawieniach (tabelka "Funkcje
        pracownikow"). To rozroznienie jest kluczowe: rola_kuchnia(nazwa)
        odpowiada na "czy TA OSOBA ma przypisana role Kuchnia" (np. do
        filtrowania listy w Rozliczeniu Dnia albo auto-sekcji w
        Obecnosciach), a to tutaj odpowiada na "czy TO JEST ten jeden,
        specjalny zbiorczy wiersz stanowiska" - to dwie rozne rzeczy. Gdy
        jest_kuchnia/jest_bar byly uzywane tu zamiennie, po wypelnieniu
        tabelki "Funkcje pracownikow" realnymi nazwiskami (np. Bar: "Anna
        Kowalska") przestawaly rozpoznawac zbiorczy wiersz "Kuchnia"/"Bar"
        (bo te funkcje zaczynaly sprawdzac TYLKO liste przypisanych osob, a
        nie slowo-klucz) - przy wczytywaniu dnia tworzyly osobny, nowy
        wiersz pracownika (np. "1 - Kuchnia") zamiast wpisac serwis do
        stalego wiersza "Kuchnia", a kazdy pracownik z realnie przypisana
        rola Bar/Kuchnia mial blokowane pole Serwis we WLASNYM wierszu,
        tak jakby byl zbiorczym wierszem stanowiska. Rownosc (nie samo
        "zawiera sie") - zeby pracownik o imieniu zawierajacym slowo-klucz
        jako fragment (np. "Barbara", "Bartosz" przy kluczu "bar") nie byl
        mylnie rozpoznany jako ten zbiorczy wiersz."""
        _, nazwa = self._rozbij_nr_nazwe(nazwa_pracownika or "")
        return nazwa.strip().lower() == self.klucz_kuchnia.strip().lower()

    def _wiersz_to_bar(self, nazwa_pracownika):
        """Patrz _wiersz_to_kuchnia - analogicznie dla zbiorczego wiersza "Bar"."""
        _, nazwa = self._rozbij_nr_nazwe(nazwa_pracownika or "")
        return nazwa.strip().lower() == self.klucz_bar.strip().lower()

    def _lista_kelnerow_bez_kuchni(self):
        """Lista pracownikow do wprowadzania utargow w 'Rozliczenie Dnia' -
        bez REALNYCH OSOB przypisanych do roli Kuchnia (ktore tam nie maja
        czego wprowadzac - nie obsluguja gosci przy kasie/stoliku) i bez
        Admina (ktory nie bierze udzialu w rozliczeniach zmianowych w
        ogole). Zawiera natomiast zawsze same zbiorcze wpisy "Kuchnia" i
        "Bar" (self.str_kuchnia / self.str_bar) - oba sa normalnie
        dodawane automatycznie (patrz wyczysc_formularz), ale musza byc
        tez na tej liscie, zeby mozna je bylo recznie dopisac z listy
        rozwijanej (np. gdy ktos przypadkiem usunie sztywny wiersz
        stanowiska i trzeba go przywrocic) - Kuchnia dostaje tu zbiorczy
        serwis do dalszego recznego rozdzielenia miedzy osoby, Bar to
        zbiorczy raport z kasy barowej, na ktorej pracuje kilka osob
        naraz."""
        wynik = []
        for wpis in (self.lista_kelnerow or []):
            _, nazwa = self._rozbij_nr_nazwe(wpis)
            if not self.jest_kuchnia(nazwa) and not self.jest_admin(nazwa):
                wynik.append(wpis)
        for zbiorczy in (self.str_kuchnia, self.str_bar):
            if zbiorczy and zbiorczy not in wynik:
                wynik.append(zbiorczy)
        return wynik

    def jest_kelner(self, nazwa_pracownika):
        """Rola "Kelner" jest wprost zaznaczana w tabelce funkcji — pracownik
        może mieć kilka ról naraz, więc to NIE jest zwykłe "ani kuchnia, ani bar"."""
        nazwa = (nazwa_pracownika or "").lower()
        lista = self._lista_pracownikow_funkcja(self.pracownicy_kelner)
        return any(n in nazwa for n in lista)

    def _rebuduj_tabele_funkcji_pracownikow(self):
        """Buduje tabelkę przypisania ról (Kelner / Kuchnia / Bar / Admin) do
        każdego pracownika z listy — zamiennik ręcznego wpisywania imion po
        przecinku. Każda rola to osobny checkbox, więc jeden pracownik może
        mieć kilka ról naraz (np. Kuchnia + Bar). Na bieżąco aktualizuje
        self.pracownicy_kuchnia_var / _bar_var / _kelner_var / _admin_var,
        które zapisz_ustawienia() odczytuje tak samo jak wcześniej pole
        tekstowe. Admin nie dostaje serwisu i nie pojawia się na liście
        raportów zmianowych (patrz jest_admin / podziel_serwis /
        uzupelnij_kelnerow_z_obecnosci)."""
        if not hasattr(self, "ramka_tabela_funkcji") or not self.ramka_tabela_funkcji.winfo_exists():
            return
        for widget in self.ramka_tabela_funkcji.winfo_children():
            widget.destroy()

        KARTA_BG = self.karta_bg
        nazwiska = []
        for wpis in getattr(self, "lista_kelnerow_raport", []) or []:
            nazwa = wpis.split(" - ", 1)[1].strip() if " - " in wpis else wpis.strip()
            if nazwa and nazwa not in nazwiska:
                nazwiska.append(nazwa)

        if not nazwiska:
            tk.Label(self.ramka_tabela_funkcji, text="Brak pracowników na liście — dodaj ich w sekcji powyżej.",
                     font=(self.czcionka_nazwa, 9, "italic"), bg=KARTA_BG, fg="#9AA5C0").grid(row=0, column=0, sticky="w", pady=4)
            self.pracownicy_kuchnia_var.set("")
            self.pracownicy_bar_var.set("")
            self.pracownicy_kelner_var.set("")
            self.pracownicy_admin_var.set("")
            return

        role = ("Kelner", "Kuchnia", "Bar", "Admin")
        tk.Label(self.ramka_tabela_funkcji, text="Pracownik", font=(self.czcionka_nazwa, 10, "bold"),
                 bg=KARTA_BG, fg=self.kolor_czcionki).grid(row=0, column=0, sticky="w", padx=(0, 16), pady=(0, 6))
        for col, rola in enumerate(role, start=1):
            tk.Label(self.ramka_tabela_funkcji, text=rola, font=(self.czcionka_nazwa, 10, "bold"),
                     bg=KARTA_BG, fg=self.kolor_czcionki).grid(row=0, column=col, padx=(0, 14), pady=(0, 6))

        # pracownik -> {"Kelner": BooleanVar, "Kuchnia": BooleanVar, "Bar": BooleanVar, "Admin": BooleanVar}
        self.funkcja_pracownika_vars = {}
        kelner_jawnie = set(self._lista_pracownikow_funkcja(self.pracownicy_kelner))

        def przelicz(*_args):
            kuchnia, bar, kelner, admin = [], [], [], []
            for nazwa, vary in self.funkcja_pracownika_vars.items():
                if vary["Kuchnia"].get(): kuchnia.append(nazwa)
                if vary["Bar"].get(): bar.append(nazwa)
                if vary["Kelner"].get(): kelner.append(nazwa)
                if vary["Admin"].get(): admin.append(nazwa)
            self.pracownicy_kuchnia_var.set(", ".join(kuchnia))
            self.pracownicy_bar_var.set(", ".join(bar))
            self.pracownicy_kelner_var.set(", ".join(kelner))
            self.pracownicy_admin_var.set(", ".join(admin))

        for idx, nazwa in enumerate(nazwiska, start=1):
            w_kuchni = self.jest_kuchnia(nazwa)
            w_barze = self.jest_bar(nazwa)
            w_admin = self.jest_admin(nazwa)
            if kelner_jawnie or self.pracownicy_kelner.strip():
                # Rola Kelner była już kiedyś zapisana tą tabelką - ufamy jej wprost.
                w_kelner = nazwa.lower() in kelner_jawnie
            else:
                # Pierwsze otwarcie po aktualizacji: zachowaj stare domyślne
                # zachowanie ("nikt przypisany do Kuchni/Bar/Admina = Kelner").
                w_kelner = not w_kuchni and not w_barze and not w_admin

            tk.Label(self.ramka_tabela_funkcji, text=nazwa, font=(self.czcionka_nazwa, 10),
                     bg=KARTA_BG, fg=self.kolor_czcionki).grid(row=idx, column=0, sticky="w", padx=(0, 16), pady=2)

            vary = {
                "Kelner": tk.BooleanVar(value=w_kelner),
                "Kuchnia": tk.BooleanVar(value=w_kuchni),
                "Bar": tk.BooleanVar(value=w_barze),
                "Admin": tk.BooleanVar(value=w_admin),
            }
            self.funkcja_pracownika_vars[nazwa] = vary
            for col, rola in enumerate(role, start=1):
                tk.Checkbutton(self.ramka_tabela_funkcji, variable=vary[rola], command=przelicz,
                                bg=KARTA_BG, fg=self.kolor_czcionki, selectcolor=self.kolor_tla,
                                activebackground=KARTA_BG, activeforeground=self.kolor_czcionki,
                                highlightthickness=0, bd=0).grid(row=idx, column=col, pady=2)

        przelicz()

    def _ikona_sciezka(self, nazwa_pliku):
        """Sciezka do pliku w folderze assets/ - dziala tez po spakowaniu do .app
        (PyInstaller --add-data), patrz build_dmg.sh."""
        if getattr(sys, "frozen", False):
            baza = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
        else:
            baza = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(baza, "assets", nazwa_pliku)

    def _ustaw_ikone_okna(self):
        """Ikona okna/paska zadan - logo KOS na granatowym tle (assets/app_icon_128.png).
        Nie przerywa startu programu, jesli sie nie uda (np. brak pliku)."""
        try:
            self._ikona_okna_img = tk.PhotoImage(file=self._ikona_sciezka("app_icon_128.png"))
            self.root.iconphoto(True, self._ikona_okna_img)
        except Exception:
            pass

    def _buduj_naglowek_marki(self):
        """Pasek na gorze okna z logo KOS i nazwa programu - nowoczesny, granatowy
        styl spojny ze strona kos-kawiarniaos.vercel.app."""
        naglowek = tk.Frame(self.root, bg="#0A0F1D", height=64)
        naglowek.pack(fill="x", side="top")
        naglowek.pack_propagate(False)

        wnetrze = tk.Frame(naglowek, bg="#0A0F1D")
        wnetrze.pack(side="left", padx=18, pady=10)

        try:
            self._logo_naglowka_img = tk.PhotoImage(file=self._ikona_sciezka("kos-logo-header.png"))
            tk.Label(wnetrze, image=self._logo_naglowka_img, bg="#0A0F1D").pack(side="left", padx=(0, 12))
        except Exception:
            pass

        tekst = tk.Frame(wnetrze, bg="#0A0F1D")
        tekst.pack(side="left")
        tk.Label(tekst, text=self.nazwa_programu, font=(self.czcionka_nazwa, 15, "bold"),
                 bg="#0A0F1D", fg="#FFFFFF").pack(anchor="w")
        tk.Label(tekst, text="Rozliczenia dnia i grafik pracy", font=(self.czcionka_nazwa, 9),
                 bg="#0A0F1D", fg="#8B96AD").pack(anchor="w")

    def buduj_gui(self):
        # Kolejnosc zakladek = kolejnosc, w jakiej realnie potrzebne sa dane:
        # 1) Obecnosci (godziny pracy - wejsciowe dla kalkulatorow napiwkow)
        # 2) Rozliczenie Dnia (utargi kelnerow, korzysta z godzin z Obecnosci)
        # 3) Kalkulator Napiwkow Gotowki (pomocniczy, korzysta z Obecnosci + dnia z Rozliczenia)
        # 4) Podsumowanie Miesiaca (odczyt juz zapisanych danych)
        # 5) Ustawienia (konfiguracja, zmieniana najrzadziej)
        self._buduj_naglowek_marki()

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook", background=self.kolor_tla, borderwidth=0)

        # Oble (zaokraglone) zakladki: ttk nie ma natywnego border-radius, wiec
        # rysujemy 3 malutkie kafelki PNG (gora zaokraglona, dol prosty) i uzywamy
        # ich jako 9-slice "image" elementu stylu - rogi zostaja ostre px-w-px,
        # a srodek rozciaga sie do szerokosci tekstu zakladki. Zadnej zaleznosci
        # od Pillow w runtime - PNG wpisane na sztywno jako base64.
        _tab_png_normal = (
            "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAsCAYAAADVX77/AAAAvElEQVR4nO2UsRGFIBAFD4owtQv7"
            "j+3C1CY00mEQQaIdubehEuxb+T9YB9O8HD3nKfZtDV/PNg/+ZfQbrRix9vLv483aG4p1RhheonQb"
            "Hjdg1PFm5W2xdWA08o3V/wAP3AE8fP2LdGvMH3jh2qyfAC1AEzxe/xT3N0ABaAEaBaAFaBSAFqBR"
            "AFqARgFoARoFoAVoFIAWoFEAWoBGAWgBGgWgBWgUgBagUQBagEYBaAEaBaAFaBSAFqBxH+AEmKsb"
            "47ZlBPcAAAAASUVORK5CYII="
        )
        _tab_png_active = (
            "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAsCAYAAADVX77/AAAAu0lEQVR4nO2UsRGFIBAFD1oxtxrL"
            "tBpza9FIh0EEiXbk3oZKsG/l/2AdTPNy9Jyn2Lc1fD3bPPiX0W+0YsTay7+PN2tvKNYZYXiJ0m14"
            "3IBRx5uVt8XWgdHIN1b/AzxwB/Dw9S/SrTF/4IVrs34CtABN8Hj9U9zfAAWgBWgUgBagUQBagEYB"
            "aAEaBaAFaBSAFqBRAFqARgFoARoFoAVoFIAWoFEAWoBGAWgBGgWgBWgUgBagUQBagMZ9gBO+7x5h"
            "lKz5CAAAAABJRU5ErkJggg=="
        )
        _tab_png_selected = (
            "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAsCAYAAADVX77/AAAAuElEQVR4nO2UwQ2FIBAFF/qwPKuy"
            "PAvRi0aCCHKawL45Gn4yb+QbrId9O7rOUyxr+Hu0fXCU0V80YsTqj0cfb9bcUK4zw/AShdvwvgGz"
            "jjcrboutA9ORbax/AxzwBPDw9m+SrTF/4IZrs/4CtABNcHn9E9zfAAWgBWgUgBagUQBagEYBaAEa"
            "BaAFaBSAFqBRAFqARgFoARoFoAVoFIAWoFEAWoBGAWgBGgWgBWgUgBagUQBagMZ9gBPWRiS0Gz1A"
            "ZAAAAABJRU5ErkJggg=="
        )
        try:
            self._tab_img_normal = tk.PhotoImage(data=_tab_png_normal, format="png")
            self._tab_img_active = tk.PhotoImage(data=_tab_png_active, format="png")
            self._tab_img_selected = tk.PhotoImage(data=_tab_png_selected, format="png")
            style.element_create(
                "Obly.tab", "image", self._tab_img_normal,
                ("selected", self._tab_img_selected),
                ("active", "!selected", self._tab_img_active),
                border=14, sticky="nsew",
            )
            style.layout("TNotebook.Tab", [
                ("Obly.tab", {"sticky": "nsew", "children": [
                    ("Notebook.padding", {"side": "top", "sticky": "nsew", "children": [
                        ("Notebook.label", {"side": "top", "sticky": ""}),
                    ]}),
                ]}),
            ])
        except tk.TclError:
            pass
        style.configure("TNotebook.Tab", padding=(22, 13), font=(self.czcionka_nazwa, 10, "bold"),
                         foreground=self.tekst_muted, focuscolor=self.kolor_tla, borderwidth=0)
        style.map("TNotebook.Tab",
                  foreground=[("selected", "#0A0F1D"), ("active", "#F5F7FA")])
        style.configure("Vertical.TScrollbar", background="#3A4A6B", troughcolor=self.kolor_tla, borderwidth=0, arrowsize=14)
        style.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF", foreground="#10182B",
                         padding=6, font=(self.czcionka_nazwa, 10), bordercolor=self.karta_border, relief="flat")
        style.map("TCombobox",
                  fieldbackground=[("readonly", "#FFFFFF")],
                  foreground=[("readonly", "#10182B")],
                  bordercolor=[("focus", self.kolor_akcent)])
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=20, pady=16)
        self.notebook.bind("<<NotebookTabChanged>>", self.sprawdz_dostep)

        # ---- 1) Obecnosci ----
        self.tab_obecnosci_outer = tk.Frame(self.notebook, bg=self.kolor_tla)
        self.notebook.add(self.tab_obecnosci_outer, text="1. Obecności")

        self.canvas_obc = tk.Canvas(self.tab_obecnosci_outer, bg=self.kolor_tla, highlightthickness=0)
        self.scrollbar_obc = ttk.Scrollbar(self.tab_obecnosci_outer, orient="vertical", command=self.canvas_obc.yview)

        self.tab_obecnosci = tk.Frame(self.canvas_obc, bg=self.kolor_tla)
        self.tab_obecnosci.bind("<Configure>", lambda e: self.canvas_obc.configure(scrollregion=self.canvas_obc.bbox("all")))

        self.canvas_obc.create_window((0, 0), window=self.tab_obecnosci, anchor="nw")
        self.canvas_obc.configure(yscrollcommand=self.scrollbar_obc.set)

        self.canvas_obc.pack(side="left", fill="both", expand=True)
        self.scrollbar_obc.pack(side="right", fill="y")

        self.tab_obecnosci_outer.bind_all("<MouseWheel>", self._on_mousewheel)

        # ---- 2) Rozliczenie Dnia ----
        self.tab_rozliczenie_outer = tk.Frame(self.notebook, bg=self.kolor_tla)
        self.notebook.add(self.tab_rozliczenie_outer, text="2. Rozliczenie Dnia")
        
        self.canvas = tk.Canvas(self.tab_rozliczenie_outer, bg=self.kolor_tla, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.tab_rozliczenie_outer, orient="vertical", command=self.canvas.yview)
        
        self.tab_rozliczenie = tk.Frame(self.canvas, bg=self.kolor_tla)
        self.tab_rozliczenie.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.tab_rozliczenie, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Kalkulator Napiwkow Gotowki byl osobna zakladka - przy zbyt wielu
        # zakladkach nazwy nie mieszczily sie w oknie i byly obcinane, wiec
        # dolaczony jest teraz jako kolejna karta na dole zakladki "2.
        # Rozliczenie Dnia" (patrz buduj_zakladke_kalkulator_gotowka) -
        # tab_kalkulator_gotowka = tab_rozliczenie, bez osobnej zakladki/pozycji w notebooku.
        self.tab_kalkulator_gotowka = self.tab_rozliczenie

        # ---- 3) Kasa: Wydatki i Wplaty ----
        self.tab_kasa_outer = tk.Frame(self.notebook, bg=self.kolor_tla)
        self.notebook.add(self.tab_kasa_outer, text="3. Paragony z nipem")

        self.canvas_kasa = tk.Canvas(self.tab_kasa_outer, bg=self.kolor_tla, highlightthickness=0)
        self.scrollbar_kasa = ttk.Scrollbar(self.tab_kasa_outer, orient="vertical", command=self.canvas_kasa.yview)

        self.tab_kasa = tk.Frame(self.canvas_kasa, bg=self.kolor_tla)
        self.tab_kasa.bind("<Configure>", lambda e: self.canvas_kasa.configure(scrollregion=self.canvas_kasa.bbox("all")))

        self.canvas_kasa.create_window((0, 0), window=self.tab_kasa, anchor="nw")
        self.canvas_kasa.configure(yscrollcommand=self.scrollbar_kasa.set)

        self.canvas_kasa.pack(side="left", fill="both", expand=True)
        self.scrollbar_kasa.pack(side="right", fill="y")

        # ---- 4) Ustawienia ----
        self.tab_ustawienia_outer = tk.Frame(self.notebook, bg=self.kolor_tla)
        self.notebook.add(self.tab_ustawienia_outer, text="4. Ustawienia")

        self.canvas_ust = tk.Canvas(self.tab_ustawienia_outer, bg=self.kolor_tla, highlightthickness=0)
        self.scrollbar_ust = ttk.Scrollbar(self.tab_ustawienia_outer, orient="vertical", command=self.canvas_ust.yview)

        self.tab_ustawienia = tk.Frame(self.canvas_ust, bg=self.kolor_tla)
        self.tab_ustawienia.bind("<Configure>", lambda e: self.canvas_ust.configure(scrollregion=self.canvas_ust.bbox("all")))

        self.canvas_ust.create_window((0, 0), window=self.tab_ustawienia, anchor="nw")
        self.canvas_ust.configure(yscrollcommand=self.scrollbar_ust.set)

        self.canvas_ust.pack(side="left", fill="both", expand=True)
        self.scrollbar_ust.pack(side="right", fill="y")

        self.buduj_zakladke_obecnosci()
        self.buduj_zakladke_rozliczenie()
        self.buduj_zakladke_kalkulator_gotowka()
        self.buduj_zakladke_kasa()
        self.buduj_zakladke_ustawienia()

    def _kos_access_token_swiezy(self):
        """Zwraca aktualny (odswiezony) access token do KOS, albo None gdy nie
        udalo sie polaczyc / sesja juz niewazna."""
        if not self.kos_refresh_token:
            return None
        sesja = self._kos_odswiez_sesje(self.kos_refresh_token)
        if not sesja or not sesja.get("access_token"):
            return None
        self.kos_access_token = sesja["access_token"]
        nowy_refresh = sesja.get("refresh_token")
        if nowy_refresh and nowy_refresh != self.kos_refresh_token:
            # Supabase rotuje refresh_token przy kazdym odswiezeniu sesji - stary
            # token przestaje dzialac od razu po wydaniu nowego. Trzeba go zapisac
            # na dysk NATYCHMIAST (nie tylko trzymac w pamieci), inaczej po
            # zamknieciu programu w pliku ustawien zostaje juz "zuzyty" (stary)
            # refresh_token, serwer go odrzuca przy nastepnym starcie i program
            # za kazdym razem prosi o ponowne logowanie.
            self.kos_refresh_token = nowy_refresh
            self._zapisz_json_ustawien()
        return self.kos_access_token

    def buduj_zakladke_ustawienia(self):
        KARTA_BG = self.karta_bg

        tk.Label(self.tab_ustawienia, text="Ustawienia", font=(self.czcionka_nazwa, 20, "bold"),
                 bg=self.kolor_tla, fg=self.kolor_czcionki).pack(pady=(16, 2), padx=16, anchor="w")
        tk.Label(self.tab_ustawienia, text="Zmiany zaczną obowiązywać po kliknięciu „Zapisz i Odśwież” na samym dole.",
                 font=(self.czcionka_nazwa, 10), bg=self.kolor_tla, fg="#9AA5C0").pack(pady=(0, 14), padx=16, anchor="w")

        def karta(tytul, opis=None):
            _zw, f = self.zrob_karte(self.tab_ustawienia, tytul)
            _zw.pack(fill="x", padx=16, pady=8)
            if opis:
                tk.Label(f, text=opis, font=(self.czcionka_nazwa, 9), bg=KARTA_BG, fg="#9AA5C0",
                         wraplength=820, justify="left").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))
            return f

        def pole(parent, wiersz, etykieta, wartosc_pocz, kolumna=0, szer=28, haslo=False):
            tk.Label(parent, text=etykieta, font=(self.czcionka_nazwa, 10), bg=KARTA_BG, fg=self.kolor_czcionki)\
                .grid(row=wiersz, column=kolumna, padx=(0, 8), pady=5, sticky="e")
            var = tk.StringVar(value=wartosc_pocz)
            e = tk.Entry(parent, textvariable=var, width=szer, relief="solid", bd=1, show="*" if haslo else "", bg=self.kolor_tla, fg=self.kolor_czcionki, insertbackground=self.kolor_czcionki)
            e.grid(row=wiersz, column=kolumna + 1, padx=(0, 20), pady=5, sticky="w")
            return var

        # ---- Podstawowe ----
        k1 = karta("Podstawowe")
        self.nazwa_programu_var = pole(k1, 1, "Nazwa programu:", self.nazwa_programu, szer=32)
        self.haslo_var = pole(k1, 2, "Hasło do zakładek zablokowanych:", self.haslo_dostepu, szer=32, haslo=True)
        self.adresy_email_var = pole(k1, 3, "E-maile do raportów:", self.adresy_email, szer=32)
        self.klucz_kuchnia_var = pole(k1, 1, "Słowo-klucz „Kuchnia”:", self.klucz_kuchnia, kolumna=2, szer=20)
        self.klucz_bar_var = pole(k1, 2, "Słowo-klucz „Bar”:", self.klucz_bar, kolumna=2, szer=20)
        self.pracownicy_kuchnia_var = tk.StringVar(value=self.pracownicy_kuchnia)
        self.pracownicy_bar_var = tk.StringVar(value=self.pracownicy_bar)
        self.pracownicy_kelner_var = tk.StringVar(value=self.pracownicy_kelner)
        self.pracownicy_admin_var = tk.StringVar(value=self.pracownicy_admin)
        self.wykluczeni_napiwki_var = pole(k1, 3, "Wykluczeni z napiwków (po przecinku):", self.wykluczeni_napiwki, kolumna=2, szer=20)
        self.klucz_admin_var = pole(k1, 4, "Słowo-klucz „Admin”:", self.klucz_admin, kolumna=2, szer=20)

        # ---- Konto KOS (aktywacja programu) ----
        k1k = karta(
            "Konto KOS",
            opis="Program jest aktywowany logowaniem na konto KOS (kos-kawiarniaos.vercel.app) - wymaga planu Ultra. Bez kodu licencyjnego.",
        )
        status_kos = (
            f"Zalogowano jako: {self.kos_login}  •  firma: {self.kos_company_nazwa or '—'}  •  plan: {self.kos_plan or '—'}"
            if self.kos_login else "Nie zalogowano."
        )
        self.lbl_status_kos = tk.Label(
            k1k, text=status_kos, font=(self.czcionka_nazwa, 10), bg=KARTA_BG, fg=self.kolor_czcionki, justify="left",
        )
        self.lbl_status_kos.grid(row=1, column=0, columnspan=4, sticky="w", pady=(0, 8))
        self.pill(
            k1k, text="Zmień konto KOS / wyloguj", command=self.wyloguj_kos,
            bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg, relief="flat", font=(self.czcionka_nazwa, 10, "bold"),
        ).grid(row=2, column=0, columnspan=4, sticky="we")

        # ---- Pracownicy (lista) ----
        # Od migracji na KOS lista pracownikow (self.lista_kelnerow) jest
        # domyslnie czytana z shift_employees przez API - ta sama lista co w
        # zakladce "Grafik pracy" na stronie. Dodawanie/usuwanie pracownikow
        # tutaj (stary zapis do Excela) przestalo miec widoczny efekt, bo
        # odswiez_liste_pracownikow() i tak nadpisuje liste danymi z KOS -
        # zamiast zostawiac pozornie dzialajace, a w praktyce martwe
        # przyciski, odsylamy do strony, gdzie pracownikow faktycznie sie
        # teraz zarzadza.
        k1b = karta(
            "Pracownicy",
            opis="Lista pracowników (imiona widoczne w rozliczeniu i obecnościach) pochodzi teraz z zakładki "
                 "\"Grafik pracy\" na stronie KOS. Dodawaj/usuwaj pracowników tam - tutaj zmiana się nie zapisze.",
        )
        self.pill(
            k1b, text="Zarządzaj pracownikami na stronie", command=lambda: webbrowser.open(KOS_API_URL),
            bg=self.kolor_akcent, fg=self.kolor_akcent_fg, relief="flat", font=(self.czcionka_nazwa, 10, "bold"),
        ).grid(row=1, column=0, columnspan=4, pady=(0, 4), sticky="we")

        # ---- Funkcje / role pracownikow (tabela zamiast pola "po przecinku") ----
        k1d = karta("Funkcje pracowników",
                    opis="Zaznacz role każdego pracownika (można wybrać kilka naraz, np. Kuchnia + Bar) — używane do podziału napiwków i przypisań w kalkulatorze.")
        self.ramka_tabela_funkcji = tk.Frame(k1d, bg=KARTA_BG)
        self.ramka_tabela_funkcji.grid(row=1, column=0, columnspan=4, sticky="we")
        self._rebuduj_tabele_funkcji_pracownikow()

        # ---- Pliki ----
        k2 = karta("Pliki Excel", opis="Nazwy plików, z których korzysta program.")
        self.szablon_var = pole(k2, 1, "Plik szablonu:", self.szablon_plik, szer=32)
        self.format_nazwy_var = pole(k2, 2, "Plik rozliczeń (użyj {miesiac}):", self.format_nazwy_pliku, szer=32)
        self.format_nazwy_obecnosci_var = pole(k2, 3, "Plik obecności (użyj {rok}):", self.format_nazwy_obecnosci, szer=32)

        tk.Label(k2, text="Folder zapisu plików (pusty = domyślny):", font=(self.czcionka_nazwa, 10), bg=KARTA_BG, fg=self.kolor_czcionki)\
            .grid(row=4, column=0, padx=(0, 8), pady=5, sticky="e")
        self.folder_zapisu_var = tk.StringVar(value=self.folder_zapisu)
        ffz = tk.Frame(k2, bg=KARTA_BG); ffz.grid(row=4, column=1, columnspan=3, sticky="w", pady=5)
        tk.Entry(ffz, textvariable=self.folder_zapisu_var, width=48, relief="solid", bd=1, bg=self.kolor_tla, fg=self.kolor_czcionki, insertbackground=self.kolor_czcionki).pack(side=tk.LEFT)
        def _wybierz_folder_zapisu():
            wybrany = filedialog.askdirectory(title="Wybierz folder zapisu plików KawiarniaOS")
            if wybrany:
                self.folder_zapisu_var.set(wybrany)
        self.pill(ffz, text="Wybierz folder...", command=_wybierz_folder_zapisu, bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg, relief="flat").pack(side=tk.LEFT, padx=6)
        self.pill(ffz, text="Domyślny", command=lambda: self.folder_zapisu_var.set(""), bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg, relief="flat").pack(side=tk.LEFT, padx=2)
        tk.Label(k2, text="Domyślnie: folder programu, a jeśli plik tam nie istnieje - ~/Documents/KawiarniaOS.", font=(self.czcionka_nazwa, 8, "italic"), bg=KARTA_BG, fg="#9AA5C0")\
            .grid(row=5, column=0, columnspan=4, sticky="w", pady=(0, 4))

        # ---- Wygląd ----
        k3 = karta("Wygląd")
        tk.Label(k3, text="Kolor tła:", font=(self.czcionka_nazwa, 10), bg=KARTA_BG, fg=self.kolor_czcionki)\
            .grid(row=1, column=0, sticky="e", padx=(0, 8), pady=5)
        ft = tk.Frame(k3, bg=KARTA_BG); ft.grid(row=1, column=1, sticky="w", pady=5)
        self.kolor_tla_var = tk.StringVar(value=self.kolor_tla)
        self.pill(ft, text="Wybierz kolor", command=self.wybierz_kolor_tla, bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg,
                  relief="flat", padx=10).pack(side=tk.LEFT)
        self.lbl_wybrany_kolor_tla = tk.Label(ft, text=self.kolor_tla, bg=self.kolor_tla, fg=self.kolor_czcionki, padx=8)
        self.lbl_wybrany_kolor_tla.pack(side=tk.LEFT, padx=8)

        tk.Label(k3, text="Kolor tekstu:", font=(self.czcionka_nazwa, 10), bg=KARTA_BG, fg=self.kolor_czcionki)\
            .grid(row=2, column=0, sticky="e", padx=(0, 8), pady=5)
        fcz = tk.Frame(k3, bg=KARTA_BG); fcz.grid(row=2, column=1, sticky="w", pady=5)
        self.kolor_czcionki_var = tk.StringVar(value=self.kolor_czcionki)
        self.pill(fcz, text="Wybierz kolor", command=self.wybierz_kolor_czcionki, bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg,
                  relief="flat", padx=10).pack(side=tk.LEFT)
        self.lbl_wybrany_kolor_czcionki = tk.Label(fcz, text=self.kolor_czcionki, bg=self.kolor_tla, fg=self.kolor_czcionki, padx=8)
        self.lbl_wybrany_kolor_czcionki.pack(side=tk.LEFT, padx=8)

        self.szerokosc_var = pole(k3, 1, "Szerokość okna (px):", str(self.szerokosc_okna), kolumna=2, szer=10)
        self.wysokosc_var = pole(k3, 2, "Wysokość okna (px):", str(self.wysokosc_okna), kolumna=2, szer=10)

        # ---- Zaawansowane: nazwy kolumn i przyciskow ----
        k4 = karta("Zaawansowane — nazwy kolumn i przycisków",
                   opis="Zwykle nie trzeba tego zmieniać. Dotyczy tylko podpisów widocznych na ekranie, nie wpływa na dane.")
        etykiety_kol = [
            ("VIP / Straty:", "kol_vip", self.kol_vip), ("Karta Mob.:", "kol_kmob", self.kol_kmob),
            ("Karta Kasa:", "kol_kkasa", self.kol_kkasa), ("Przelewy:", "kol_przelewy", self.kol_przelewy),
            ("Foodsi:", "kol_foodsi", self.kol_foodsi), ("Serwis / Tip:", "kol_serwis", self.kol_serwis),
            ("Suma POS:", "kol_razem", self.kol_razem), ("Gotówka oczek.:", "kol_gotowka", self.kol_gotowka),
            ("Zdano fizycznie:", "kol_sprawdz", self.kol_sprawdz), ("Różnica:", "kol_roznica", self.kol_roznica),
        ]
        for idx, (etykieta, klucz, wartosc) in enumerate(etykiety_kol):
            wiersz = 1 + idx // 2
            kolumna = 0 if idx % 2 == 0 else 2
            setattr(self, f"{klucz}_var", pole(k4, wiersz, etykieta, wartosc, kolumna=kolumna, szer=18))

        etykiety_txt = [
            ("Nagłówek daty:", "txt_wybor_daty", self.txt_wybor_daty), ("Nagłówek weryfikacji:", "txt_weryfikacja", self.txt_weryfikacja),
            ("Nagłówek kalkulatora:", "txt_kalkulator", self.txt_kalkulator), ("Nagłówek tabeli:", "txt_wprowadzanie", self.txt_wprowadzanie),
            ("Etykieta „Rozlicza”:", "txt_kto_rozlicza", self.txt_kto_rozlicza), ("Etykieta „Komentarz”:", "txt_komentarz", self.txt_komentarz),
            ("Przycisk „Wczytaj”:", "txt_wczytaj_btn", self.txt_wczytaj_btn), ("Przycisk „Wyczyść”:", "txt_wyczysc_btn", self.txt_wyczysc_btn),
            ("Przycisk „Zapisz”:", "txt_zapisz_btn", self.txt_zapisz_btn), ("Przycisk „E-mail”:", "txt_wyslij_btn", self.txt_wyslij_btn),
            ("Przycisk „Podziel i Przypisz”:", "txt_podziel_btn", self.txt_podziel_btn),
            ("Przycisk „+ Dodaj Kelnera”:", "txt_dodaj_kelnera_btn", self.txt_dodaj_kelnera_btn),
            ("Przycisk „- Usuń” (wiersz):", "txt_usun_wiersz_btn", self.txt_usun_wiersz_btn),
            ("Przycisk „Oblicz podział”:", "txt_oblicz_gotowka_btn", self.txt_oblicz_gotowka_btn),
            ("Przycisk „Wczytaj Dzień” (Obecności):", "txt_wczytaj_obc_btn", self.txt_wczytaj_obc_btn),
            ("Przycisk „Wyczyść” (Obecności):", "txt_wyczysc_obc_btn", self.txt_wyczysc_obc_btn),
            ("Przycisk „+ Dodaj Pracownika”:", "txt_dodaj_obc_btn", self.txt_dodaj_obc_btn),
            ("Przycisk „Zapisz Obecności Dnia”:", "txt_zapisz_obc_btn", self.txt_zapisz_obc_btn),
            ("Nagłówek kalkulatora gotówki:", "txt_kalkulator_gotowka_naglowek", self.txt_kalkulator_gotowka_naglowek),
            ("Nagłówek daty (Obecności):", "txt_wybor_daty_obc", self.txt_wybor_daty_obc),
            ("Nagłówek godzin pracy:", "txt_godziny_naglowek", self.txt_godziny_naglowek),
            ("Nagłówek podsumowania (Obecności):", "txt_podsumowanie_obc_naglowek", self.txt_podsumowanie_obc_naglowek),
        ]
        baza = 1 + (len(etykiety_kol) + 1) // 2
        for idx, (etykieta, klucz, wartosc) in enumerate(etykiety_txt):
            wiersz = baza + idx // 2
            kolumna = 0 if idx % 2 == 0 else 2
            setattr(self, f"{klucz}_var", pole(k4, wiersz, etykieta, wartosc, kolumna=kolumna, szer=18))

        btn_zapisz = self.pill(self.tab_ustawienia, text="Zapisz i Odśwież", command=self.zapisz_ustawienia,
                                bg=self.kolor_akcent, fg=self.kolor_akcent_fg, relief="flat", font=(self.czcionka_nazwa, 13, "bold"), height=2)
        btn_zapisz.pack(fill="x", padx=16, pady=(10, 24))

    def buduj_zakladke_rozliczenie(self):
        now = datetime.now()
        
        # 1. SEKJA DATY
        _zw_top, frame_top = self.zrob_karte(self.tab_rozliczenie, self.txt_wybor_daty)
        _zw_top.pack(fill="x", padx=20, pady=10)
        
        tk.Label(frame_top, text="Rok:", font=(self.czcionka_nazwa, 10, "bold"), bg=self.karta_bg, fg=self.kolor_czcionki).grid(row=0, column=0, padx=(0, 6), pady=6)
        self.rok_var = tk.StringVar(value=now.strftime("%Y")); tk.Entry(frame_top, textvariable=self.rok_var, width=5, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, relief="flat", highlightthickness=1, highlightbackground=self.karta_border, highlightcolor=self.kolor_akcent).grid(row=0, column=1, pady=6)

        tk.Label(frame_top, text="Msc (01-12):", font=(self.czcionka_nazwa, 10, "bold"), bg=self.karta_bg, fg=self.kolor_czcionki).grid(row=0, column=2, padx=(16, 6), pady=6)
        self.miesiac_var = tk.StringVar(value=now.strftime("%m")); tk.Entry(frame_top, textvariable=self.miesiac_var, width=4, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, relief="flat", highlightthickness=1, highlightbackground=self.karta_border, highlightcolor=self.kolor_akcent).grid(row=0, column=3, pady=6)
        self.miesiac_var.trace_add("write", lambda *a: self.odswiez_liste_pracownikow())

        tk.Label(frame_top, text="Dzień:", font=(self.czcionka_nazwa, 10, "bold"), bg=self.karta_bg, fg=self.kolor_czcionki).grid(row=0, column=4, padx=(16, 6), pady=6)
        self.dzien_var = tk.StringVar(value=now.strftime("%d")); tk.Entry(frame_top, textvariable=self.dzien_var, width=4, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, relief="flat", highlightthickness=1, highlightbackground=self.karta_border, highlightcolor=self.kolor_akcent).grid(row=0, column=5, pady=6)
        
        self.pill(frame_top, text=self.txt_wczytaj_btn, command=self.wczytaj_dzien, bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg, relief="flat", font=(self.czcionka_nazwa, 10, "bold")).grid(row=0, column=6, padx=(24, 8), pady=6)
        self.pill(frame_top, text=self.txt_wyczysc_btn, command=self.wyczysc_formularz, bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg, relief="flat").grid(row=0, column=7, padx=8, pady=6)
        self.pill(frame_top, text="Plik na kolejny miesiąc", command=self.wygeneruj_plik_kolejnego_miesiaca, bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg, relief="flat", font=(self.czcionka_nazwa, 10, "bold")).grid(row=0, column=8, padx=(8, 0), pady=6)

        # 2. WERYFIKACJA
        _zw_bot, frame_bot = self.zrob_karte(self.tab_rozliczenie, self.txt_weryfikacja)
        _zw_bot.pack(fill="x", padx=20, pady=10)

        _pole_kw = dict(relief="flat", highlightthickness=1)
        self.fiskal_var = tk.StringVar(value="0.0"); self.fiskal_var.trace_add("write", self.weryfikuj_dane)
        tk.Label(frame_bot, text="Fiskalny [Z]:", font=(self.czcionka_nazwa, 10, "bold"), bg=self.karta_bg, fg=self.kolor_czcionki).grid(row=0, column=0, padx=(0, 6), pady=(6, 10), sticky="e")
        ef = tk.Entry(frame_bot, textvariable=self.fiskal_var, width=12, bg=self.pole_uwaga_bg, fg=self.pole_uwaga_fg, insertbackground=self.pole_uwaga_fg, highlightbackground=self.karta_border, highlightcolor=self.kolor_akcent, **_pole_kw); ef.grid(row=0, column=1, padx=(0, 16), pady=(6, 10)); self.przypnij_czyszczenie_zer(ef)

        self.raport_karta_mob_var = tk.StringVar(value="0.0"); self.raport_karta_mob_var.trace_add("write", self.weryfikuj_dane)
        tk.Label(frame_bot, text="Karta Mobilna:", font=(self.czcionka_nazwa, 10, "bold"), bg=self.karta_bg, fg=self.kolor_czcionki).grid(row=0, column=2, padx=(0, 6), pady=(6, 10), sticky="e")
        ekm = tk.Entry(frame_bot, textvariable=self.raport_karta_mob_var, width=12, bg=self.pole_info_bg, fg=self.pole_info_fg, insertbackground=self.pole_info_fg, highlightbackground=self.karta_border, highlightcolor=self.kolor_akcent, **_pole_kw); ekm.grid(row=0, column=3, padx=(0, 16), pady=(6, 10)); self.przypnij_czyszczenie_zer(ekm)
        
        self.raport_karta_kasa_var = tk.StringVar(value="0.0"); self.raport_karta_kasa_var.trace_add("write", self.weryfikuj_dane)
        tk.Label(frame_bot, text="Karta Kasa:", font=(self.czcionka_nazwa, 10, "bold"), bg=self.karta_bg, fg=self.kolor_czcionki).grid(row=0, column=4, padx=(0, 6), pady=(6, 10), sticky="e")
        ekk = tk.Entry(frame_bot, textvariable=self.raport_karta_kasa_var, width=12, bg=self.pole_info_bg, fg=self.pole_info_fg, insertbackground=self.pole_info_fg, highlightbackground=self.karta_border, highlightcolor=self.kolor_akcent, **_pole_kw); ekk.grid(row=0, column=5, padx=(0, 16), pady=(6, 10)); self.przypnij_czyszczenie_zer(ekk)

        self.faktury_var = tk.StringVar(value="0.0"); self.faktury_var.trace_add("write", self.weryfikuj_dane)
        tk.Label(frame_bot, text="Faktury [AA]:", font=(self.czcionka_nazwa, 10, "bold"), bg=self.karta_bg, fg=self.kolor_czcionki).grid(row=0, column=6, padx=(10, 6), pady=(6, 10), sticky="e")
        efak = tk.Entry(frame_bot, textvariable=self.faktury_var, width=12, bg=self.pole_uwaga_bg, fg=self.pole_uwaga_fg, insertbackground=self.pole_uwaga_fg, highlightbackground=self.karta_border, highlightcolor=self.kolor_akcent, **_pole_kw); efak.grid(row=0, column=7, pady=(6, 10)); self.przypnij_czyszczenie_zer(efak)

        self.lbl_suma_kart = tk.Label(frame_bot, text="Suma kart: 0.00", font=(self.czcionka_nazwa, 9, "italic"), bg=self.karta_bg, fg=self.tekst_muted)
        self.lbl_suma_kart.grid(row=1, column=2, columnspan=4)

        self.lbl_weryfikacja = tk.Label(frame_bot, text="Wpisz dane...", font=(self.czcionka_nazwa, 10, "bold"), bg=self.karta_bg, fg=self.kolor_czcionki)
        self.lbl_weryfikacja.grid(row=1, column=0, columnspan=2, pady=5)

        self.lbl_gotowka_suma = tk.Label(frame_bot, text=f"OCZEKIWANA GOTÓWKA: 0.00", font=(self.czcionka_nazwa, 11, "bold"), fg="#7DD3FC", bg=self.karta_bg)
        self.lbl_gotowka_suma.grid(row=1, column=6, columnspan=2)

        self.lbl_sprawdz_suma = tk.Label(frame_bot, text=f"SPRAWDZONO W KASIE: 0.00", font=(self.czcionka_nazwa, 11, "bold"), fg="#4ADE80", bg=self.karta_bg)
        self.lbl_sprawdz_suma.grid(row=2, column=6, columnspan=2)

        self.lbl_roznica_suma = tk.Label(frame_bot, text=f"RÓŻNICA: ---", font=(self.czcionka_nazwa, 12, "bold"), bg=self.karta_bg, fg=self.kolor_czcionki)
        self.lbl_roznica_suma.grid(row=3, column=6, columnspan=2, pady=(0,5))

        # 3. KALKULATOR
        _zw_fc, fc = self.zrob_karte(self.tab_rozliczenie, self.txt_kalkulator)
        _zw_fc.pack(fill="x", padx=20, pady=10)
        
        _calc_kw = dict(relief="flat", highlightthickness=1, highlightbackground=self.karta_border, highlightcolor=self.kolor_akcent, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg)
        self.calc_1_var = tk.StringVar(value="0.0"); ec1 = tk.Entry(fc, textvariable=self.calc_1_var, width=7, **_calc_kw); tk.Label(fc, text="Raport 1:", font=(self.czcionka_nazwa, 10), bg=self.karta_bg, fg=self.kolor_czcionki).grid(row=0, column=0, pady=6); ec1.grid(row=0, column=1, padx=(6, 5), pady=6); self.przypnij_czyszczenie_zer(ec1)
        self.calc_2_var = tk.StringVar(value="0.0"); ec2 = tk.Entry(fc, textvariable=self.calc_2_var, width=7, **_calc_kw); tk.Label(fc, text="Raport 2:", font=(self.czcionka_nazwa, 10), bg=self.karta_bg, fg=self.kolor_czcionki).grid(row=0, column=2, pady=6); ec2.grid(row=0, column=3, padx=(6, 5), pady=6); self.przypnij_czyszczenie_zer(ec2)
        self.calc_3_var = tk.StringVar(value="0.0"); ec3 = tk.Entry(fc, textvariable=self.calc_3_var, width=7, **_calc_kw); tk.Label(fc, text="Raport 3:", font=(self.czcionka_nazwa, 10), bg=self.karta_bg, fg=self.kolor_czcionki).grid(row=0, column=4, pady=6); ec3.grid(row=0, column=5, padx=(6, 5), pady=6); self.przypnij_czyszczenie_zer(ec3)
        self.pill(fc, text=self.txt_podziel_btn, command=self.podziel_serwis, bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg, relief="flat", font=(self.czcionka_nazwa, 10, "bold")).grid(row=0, column=6, padx=15, pady=6)
        tk.Label(fc, text="(proporcjonalnie do godzin z zakładki Obecności, kuchnia zbiorczo + kelnerzy, bar bez zmian)", font=(self.czcionka_nazwa, 8, "italic"), bg=self.karta_bg, fg=self.tekst_muted).grid(row=1, column=0, columnspan=7, sticky="w", pady=(2,4))

        # 4. KELNERZY
        self._zw_mid, self.frame_mid = self.zrob_karte(self.tab_rozliczenie, self.txt_wprowadzanie)
        self._zw_mid.pack(fill="both", expand=True, padx=20, pady=10)

        nags = ["Pracownik", self.kol_vip, self.kol_kmob, self.kol_kkasa, self.kol_przelewy, self.kol_foodsi, self.kol_serwis, self.kol_razem, self.kol_gotowka, self.kol_sprawdz, self.kol_roznica]
        for i, n in enumerate(nags): tk.Label(self.frame_mid, text=n, font=(self.czcionka_nazwa, 9, "bold"), bg=self.karta_bg, fg=self.tekst_muted, wraplength=84, justify="center").grid(row=0, column=i, padx=4, pady=(0, 10), sticky="s")

        fb = tk.Frame(self.tab_rozliczenie, padx=10, bg=self.kolor_tla); fb.pack(fill="x")
        self.pill(fb, text=self.txt_dodaj_kelnera_btn, command=self.dodaj_wiersz, bg=self.kolor_akcent, fg=self.kolor_akcent_fg, relief="flat", font=(self.czcionka_nazwa, 10, "bold")).pack(side=tk.LEFT, padx=10)
        self.pill(fb, text=self.txt_usun_wiersz_btn, command=self.usun_wiersz, bg="#fee2e2", fg="#991b1b", relief="flat").pack(side=tk.LEFT, padx=5)

        # 5. ROZLICZA I MAIL
        fa = tk.Frame(self.tab_rozliczenie, bg=self.kolor_tla); fa.pack(fill="x", padx=10, pady=15)
        tk.Label(fa, text=self.txt_kto_rozlicza, bg=self.kolor_tla, fg=self.kolor_czcionki, font=(self.czcionka_nazwa, 10, "bold")).grid(row=0, column=0, sticky="ne", pady=5)
        self.osoba_var = tk.StringVar()
        self.osoba_cb = ttk.Combobox(fa, textvariable=self.osoba_var, values=self.lista_kelnerow, width=20); self.osoba_cb.grid(row=0, column=1, sticky="nw", pady=5)
        self.osoba_cb.bind('<KeyRelease>', lambda e, cb=self.osoba_cb: self.autocomplete_pracownik(e, cb))
        
        tk.Label(fa, text=self.txt_komentarz, bg=self.kolor_tla, fg=self.kolor_czcionki, justify="right").grid(row=0, column=2, sticky="ne", pady=5, padx=20)
        kf = tk.Frame(fa, bg=self.kolor_tla, highlightthickness=1, highlightbackground=self.karta_border); kf.grid(row=0, column=3, sticky="we", pady=5)
        self.komentarz_txt = tk.Text(kf, width=60, height=3, wrap="word", bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, relief="flat", bd=8); self.komentarz_txt.pack(side="left", fill="both", expand=True)
        self.lbl_komentarz_licznik = tk.Label(fa, text=f"0 / {MIN_DLUGOSC_KOMENTARZA} znaków", font=(self.czcionka_nazwa, 8, "italic"), bg=self.kolor_tla, fg="#8B96AD")
        self.lbl_komentarz_licznik.grid(row=1, column=3, sticky="w", pady=(2, 0))
        self.komentarz_txt.bind("<KeyRelease>", self._odswiez_licznik_komentarza)

        bf = tk.Frame(self.tab_rozliczenie, bg=self.kolor_tla); bf.pack(pady=5, fill="x", padx=10)
        self.btn_zapisz_dzien = self.pill(bf, text=self.txt_zapisz_btn, command=self.zapisz_excel, bg=self.kolor_akcent, fg=self.kolor_akcent_fg, relief="flat", font=(self.czcionka_nazwa, 13, "bold"), height=2)
        self.btn_zapisz_dzien.pack(side=tk.LEFT, expand=True, fill="x", padx=5)
        self.pill(bf, text=self.txt_wyslij_btn, command=self.wyslij_raport_email, bg=self.kolor_akcent, fg=self.kolor_akcent_fg, outline=True, relief="flat", font=(self.czcionka_nazwa, 13, "bold"), height=2).pack(side=tk.RIGHT, expand=True, fill="x", padx=5)

    def buduj_zakladke_kalkulator_gotowka(self):
        tk.Label(self.tab_kalkulator_gotowka, text="Kalkulator Napiwków Gotówki", font=(self.czcionka_nazwa, 18, "bold"), bg=self.kolor_tla, fg=self.kolor_czcionki).pack(pady=15)

        _zw_fg, fg = self.zrob_karte(self.tab_kalkulator_gotowka, self.txt_kalkulator_gotowka_naglowek)
        _zw_fg.pack(fill="x", padx=20, pady=12)

        tk.Label(fg, text="Gotówka do podziału:", bg=self.kolor_tla, fg=self.kolor_czcionki).grid(row=0, column=0, padx=5)
        self.gotowka_pula_var = tk.StringVar(value="0.0")
        egp = tk.Entry(fg, textvariable=self.gotowka_pula_var, width=10, bg="white", fg="black", insertbackground="black"); egp.grid(row=0, column=1, padx=5); self.przypnij_czyszczenie_zer(egp)
        self.pill(fg, text=self.txt_oblicz_gotowka_btn, command=self.oblicz_napiwki_gotowka, bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg, relief="flat", font=(self.czcionka_nazwa, 10, "bold")).grid(row=0, column=2, padx=15)
        tk.Label(fg, text="Kuchnia zawsze dostaje 25% (wg godzin), reszta 75% (wg godzin) — wynik tylko na ekranie", font=(self.czcionka_nazwa, 8, "italic"), bg=self.kolor_tla, fg="#8B96AD").grid(row=0, column=3, padx=10, sticky="w")

        self.txt_wynik_gotowka = tk.Text(fg, width=66, height=3, wrap="word", state="disabled",
                                          bg=self.kolor_tla, fg=self.tekst_muted, relief="flat", bd=0,
                                          font=(self.czcionka_nazwa, 10))
        self.txt_wynik_gotowka.grid(row=1, column=0, columnspan=4, pady=(10, 2), sticky="w")
        self.txt_wynik_gotowka.config(state="normal")
        self.txt_wynik_gotowka.insert("1.0", "Wpisz kwotę gotówki i kliknij \u201eOblicz podział\u201d, aby zobaczyć wynik.")
        self.txt_wynik_gotowka.config(state="disabled")

        tk.Label(self.tab_kalkulator_gotowka, text='Kalkulator korzysta z godzin wpisanych w zakładce "Obecności" oraz dnia wybranego powyżej.', font=(self.czcionka_nazwa, 9, "italic"), bg=self.kolor_tla, fg="#8B96AD").pack(pady=(0,10))

    # ===================== KASA: WYDATKI I WPŁATY =====================
    # Wydatki gotowkowe (paragony z NIP) i wplaty gotowki do banku - kazdy
    # wpis zyje wylacznie w KOS (freyonator_kasa_operacje, patrz
    # supabase/018_freyonator_kasa_operacje.sql i
    # app/api/freyonator/kasa-operacje), tak jak "Rozliczenie Dnia" od Fazy 1
    # migracji. "Aktualny stan gotowki" NIE jest wpisywany recznie - liczony
    # jest po stronie serwera jako suma gotowki sprawdzonej (Rozliczenia Dnia)
    # minus suma wszystkich operacji kasowych, wiec kazdy dodany/usuniety
    # wpis od razu zmienia pokazywany stan.

    def buduj_zakladke_kasa(self):
        tk.Label(self.tab_kasa, text="Kasa: Wydatki i Wpłaty", font=(self.czcionka_nazwa, 18, "bold"),
                 bg=self.kolor_tla, fg=self.kolor_czcionki).pack(pady=15)
        tk.Label(self.tab_kasa,
                 text="Wydatki gotówkowe (paragony z NIP) i wpłaty gotówki do banku - każdy wpis pomniejsza\n"
                      "aktualny stan gotówki w kasie (liczony z gotówki sprawdzonej w Rozliczeniach Dnia).",
                 font=(self.czcionka_nazwa, 9, "italic"), bg=self.kolor_tla, fg="#8B96AD", justify="center").pack(pady=(0, 10))

        # ---- Karta: aktualny stan gotówki ----
        _zw_stan, fstan = self.zrob_karte(self.tab_kasa, "Aktualny stan gotówki w kasie")
        _zw_stan.pack(fill="x", padx=20, pady=12)
        self.lbl_kasa_stan = tk.Label(fstan, text="— zł", font=(self.czcionka_nazwa, 26, "bold"),
                                       fg="#7DD3FC", bg=self.karta_bg)
        self.lbl_kasa_stan.pack(side="left", padx=(4, 20), pady=6)
        self.pill(fstan, text="Odśwież", command=self._kos_wczytaj_kasa, bg=self.kolor_wygaszony_bg,
                  fg=self.kolor_wygaszony_fg, relief="flat").pack(side="left", pady=6)
        tk.Label(fstan, text="Suma gotówki sprawdzonej (Rozliczenia Dnia) minus wydatki i wpłaty do banku.",
                 font=(self.czcionka_nazwa, 9), bg=self.karta_bg, fg=self.tekst_muted, wraplength=320,
                 justify="left").pack(side="left", padx=16)

        # ---- Karta: dodaj operację ----
        _zw_dod, fd = self.zrob_karte(self.tab_kasa, "Dodaj operację")
        _zw_dod.pack(fill="x", padx=20, pady=12)
        _pole_kw = dict(relief="flat", highlightthickness=1, highlightbackground=self.karta_border,
                         highlightcolor=self.kolor_akcent, bg=self.pole_bg, fg=self.pole_fg,
                         insertbackground=self.pole_fg, font=(self.czcionka_nazwa, 10))

        tk.Label(fd, text="Typ:", bg=self.karta_bg, fg=self.kolor_czcionki,
                 font=(self.czcionka_nazwa, 10, "bold")).grid(row=0, column=0, padx=8, pady=8, sticky="e")
        self.kasa_typ_var = tk.StringVar(value=self._kasa_typy_opis[0])
        typ_cb = ttk.Combobox(fd, textvariable=self.kasa_typ_var, values=self._kasa_typy_opis, state="readonly", width=28)
        typ_cb.grid(row=0, column=1, padx=8, pady=8, sticky="w")

        tk.Label(fd, text="Data (RRRR-MM-DD):", bg=self.karta_bg, fg=self.kolor_czcionki,
                 font=(self.czcionka_nazwa, 10, "bold")).grid(row=0, column=2, padx=8, pady=8, sticky="e")
        self.kasa_data_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        tk.Entry(fd, textvariable=self.kasa_data_var, width=12, **_pole_kw).grid(row=0, column=3, padx=8, pady=8, sticky="w")

        tk.Label(fd, text="Kwota (zł):", bg=self.karta_bg, fg=self.kolor_czcionki,
                 font=(self.czcionka_nazwa, 10, "bold")).grid(row=1, column=0, padx=8, pady=8, sticky="e")
        self.kasa_kwota_var = tk.StringVar(value="0.0")
        ekw = tk.Entry(fd, textvariable=self.kasa_kwota_var, width=12, **_pole_kw)
        ekw.grid(row=1, column=1, padx=8, pady=8, sticky="w")
        self.przypnij_czyszczenie_zer(ekw)

        tk.Label(fd, text="Opis:", bg=self.karta_bg, fg=self.kolor_czcionki,
                 font=(self.czcionka_nazwa, 10, "bold")).grid(row=1, column=2, padx=8, pady=8, sticky="e")
        self.kasa_opis_var = tk.StringVar(value="")
        tk.Entry(fd, textvariable=self.kasa_opis_var, width=32, **_pole_kw).grid(row=1, column=3, padx=8, pady=8, sticky="w")

        self.pill(fd, text="Dodaj operację", command=self._kos_dodaj_operacje_kasa, bg=self.kolor_akcent,
                  fg="#FFFFFF", relief="flat", font=(self.czcionka_nazwa, 10, "bold")).grid(row=2, column=0, columnspan=4, pady=(10, 4))

        # ---- Karta: historia operacji ----
        _zw_hist, self.frame_kasa_historia = self.zrob_karte(self.tab_kasa, "Historia operacji (ostatnie 100)")
        _zw_hist.pack(fill="x", padx=20, pady=(12, 24))
        self.lbl_kasa_historia_pusta = tk.Label(self.frame_kasa_historia, text="Ładowanie...",
                                                 font=(self.czcionka_nazwa, 10, "italic"),
                                                 bg=self.karta_bg, fg=self.tekst_muted)
        self.lbl_kasa_historia_pusta.pack(anchor="w", pady=6)
        self.wiersze_kasa.append(self.lbl_kasa_historia_pusta)

    def _kasa_wyczysc_historie(self):
        for w in getattr(self, "wiersze_kasa", []):
            w.destroy()
        self.wiersze_kasa = []

    def _kasa_dodaj_wiersz_historii(self, op):
        typ = op.get("typ")
        typ_opis = "Wydatek" if typ == "wydatek" else "Wpłata do banku"
        kolor_typ = "#F87171" if typ == "wydatek" else "#60A5FA"
        kwota = self.bezpieczny_float(op.get("kwota"))
        opis = (op.get("opis") or "").strip()
        nip = (op.get("nip") or "").strip()
        numer = (op.get("numer_dokumentu") or "").strip()

        w = tk.Frame(self.frame_kasa_historia, bg=self.karta_bg)
        w.pack(fill="x", pady=3)
        tk.Label(w, text=str(op.get("data") or ""), width=11, font=(self.czcionka_nazwa, 10),
                 bg=self.karta_bg, fg=self.tekst_muted).pack(side="left", padx=(0, 10))
        tk.Label(w, text=typ_opis, width=16, font=(self.czcionka_nazwa, 10, "bold"),
                 bg=self.karta_bg, fg=kolor_typ, anchor="w").pack(side="left", padx=(0, 10))
        tk.Label(w, text=f"-{kwota:.2f} zł", width=12, font=(self.czcionka_nazwa, 10, "bold"),
                 bg=self.karta_bg, fg=self.kolor_czcionki, anchor="w").pack(side="left", padx=(0, 10))
        opis_pelny = opis
        if nip:
            opis_pelny += (" • " if opis_pelny else "") + f"NIP {nip}"
        if numer:
            opis_pelny += (" • " if opis_pelny else "") + f"nr {numer}"
        tk.Label(w, text=opis_pelny or "—", font=(self.czcionka_nazwa, 10), bg=self.karta_bg,
                 fg=self.kolor_czcionki, anchor="w", justify="left").pack(side="left", fill="x", expand=True, padx=(0, 10))
        op_id = op.get("id")
        self.pill(w, text="Usuń", command=lambda i=op_id: self._kos_usun_operacje_kasa(i),
                  bg="#7F1D1D", fg="#FECACA", relief="flat", width=70).pack(side="right")
        self.wiersze_kasa.append(w)

    def _kos_wczytaj_kasa(self):
        """Pobiera z KOS aktualny stan gotowki i historie operacji kasowych
        (freyonator_kasa_operacje, patrz app/api/freyonator/kasa-operacje).
        Wywolywane przy starcie programu, po (ponownym) zalogowaniu do KOS
        i po kazdej dodanej/usunietej operacji - zawsze pokazuje aktualny
        stan. Ciche niepowodzenie (np. brak internetu/logowania) - ta
        zakladka nigdy nie blokuje reszty programu."""
        if not hasattr(self, "lbl_kasa_stan"):
            return  # GUI jeszcze niezbudowane (wywolanie z __init__ przed buduj_gui)
        access_token = self._kos_access_token_swiezy()
        if not access_token:
            self.lbl_kasa_stan.config(text="— zł (zaloguj się do KOS)")
            self._kasa_wyczysc_historie()
            self.lbl_kasa_historia_pusta = tk.Label(
                self.frame_kasa_historia, text="Zaloguj się do konta KOS w Ustawieniach, aby zobaczyć historię.",
                font=(self.czcionka_nazwa, 10, "italic"), bg=self.karta_bg, fg=self.tekst_muted)
            self.lbl_kasa_historia_pusta.pack(anchor="w", pady=6)
            self.wiersze_kasa.append(self.lbl_kasa_historia_pusta)
            return

        status, dane = self._kos_http(
            "GET", f"{KOS_API_URL}/api/freyonator/kasa-operacje?limit=100",
            headers={"Authorization": f"Bearer {access_token}"},
        )
        if status != 200 or not isinstance(dane, dict):
            self.lbl_kasa_stan.config(text="— zł (błąd pobierania)")
            return

        self.lbl_kasa_stan.config(text=f"{self.bezpieczny_float(dane.get('stan_gotowki')):.2f} zł")
        self._kasa_wyczysc_historie()
        operacje = dane.get("operacje") or []
        if not operacje:
            pusta = tk.Label(self.frame_kasa_historia, text="Brak zapisanych operacji.",
                              font=(self.czcionka_nazwa, 10, "italic"), bg=self.karta_bg, fg=self.tekst_muted)
            pusta.pack(anchor="w", pady=6)
            self.wiersze_kasa.append(pusta)
        else:
            for op in operacje:
                self._kasa_dodaj_wiersz_historii(op)

    def _kos_dodaj_operacje_kasa(self):
        typ = self._kasa_typ_backend.get(self.kasa_typ_var.get(), "wydatek")
        data_wpis = self.kasa_data_var.get().strip()
        try:
            datetime.strptime(data_wpis, "%Y-%m-%d")
        except ValueError:
            return messagebox.showwarning("Nieprawidłowa data", "Podaj datę w formacie RRRR-MM-DD.")

        kwota = self.pobierz_float(self.kasa_kwota_var.get())
        if kwota <= 0:
            return messagebox.showwarning("Nieprawidłowa kwota", "Kwota musi być większa od zera.")

        access_token = self._kos_access_token_swiezy()
        if not access_token:
            return messagebox.showerror(
                "Brak połączenia z KOS",
                "Nie udało się połączyć z kontem KOS.\n"
                "Sprawdź internet albo zaloguj się ponownie w Ustawieniach → Konto KOS.\n"
                "Operacja NIE została zapisana.",
            )

        body = {
            "p_data": data_wpis,
            "p_typ": typ,
            "p_kwota": kwota,
            "p_opis": self.kasa_opis_var.get().strip(),
            "p_nip": "",
            "p_numer_dokumentu": "",
            "p_dodal": self.kos_login,
        }
        status, dane = self._kos_http(
            "POST", f"{KOS_API_URL}/api/freyonator/kasa-operacje", body=body,
            headers={"Content-Type": "application/json", "Authorization": f"Bearer {access_token}"},
        )
        if status != 200:
            szczegol = (dane.get("error") if isinstance(dane, dict) else "") or ""
            tresc = "Nie udało się dodać operacji w KOS. Operacja NIE została zapisana."
            if status is None:
                tresc += "\n\nBrak połączenia z internetem."
            elif szczegol:
                tresc += f"\n\n{szczegol}"
            return messagebox.showerror("Błąd zapisu", tresc)

        self.kasa_kwota_var.set("0.0")
        self.kasa_opis_var.set("")
        self._toast("Dodano operację ✓", typ="sukces")
        self._kos_wczytaj_kasa()

    def _kos_usun_operacje_kasa(self, op_id):
        if not op_id:
            return
        if not messagebox.askyesno("Usuń operację", "Czy na pewno usunąć tę operację kasową?"):
            return
        access_token = self._kos_access_token_swiezy()
        if not access_token:
            return messagebox.showerror(
                "Brak połączenia z KOS",
                "Nie udało się połączyć z kontem KOS.\nSprawdź internet albo zaloguj się ponownie w Ustawieniach → Konto KOS.",
            )
        status, dane = self._kos_http(
            "DELETE", f"{KOS_API_URL}/api/freyonator/kasa-operacje?id={urllib.parse.quote(str(op_id))}",
            headers={"Authorization": f"Bearer {access_token}"},
        )
        if status != 200:
            szczegol = (dane.get("error") if isinstance(dane, dict) else "") or ""
            tresc = "Nie udało się usunąć operacji."
            if status is None:
                tresc += "\n\nBrak połączenia z internetem."
            elif szczegol:
                tresc += f"\n\n{szczegol}"
            return messagebox.showerror("Błąd usuwania", tresc)
        self._toast("Usunięto operację", typ="sukces")
        self._kos_wczytaj_kasa()

    # ===================== OBECNOŚCI (GODZINY PRACY) =====================
    # Układ pliku: kolumna = pracownik (od/do), tworzona dynamicznie przy pierwszym użyciu.
    # Wiersz = rzeczywisty dzień kalendarzowy (bez sztucznego dopełniania do 31), więc plik
    # sam dopasowuje się do dowolnej liczby dni w miesiącu i dowolnej liczby pracowników.

    NAGLOWEK_OBC = 3  # wiersze 1=sekcja, 2=imię, 3=od/do
    MIESIACE_OBC = ["STYCZEŃ", "LUTY", "MARZEC", "KWIECIEŃ", "MAJ", "CZERWIEC",
                    "LIPIEC", "SIERPIEŃ", "WRZESIEŃ", "PAŹDZIERNIK", "LISTOPAD", "GRUDZIEŃ"]

    def get_plik_obecnosci_rok(self, rok):
        nazwa = self.format_nazwy_obecnosci.replace("{rok}", str(rok))
        # Jeśli plik istnieje lokalnie, użyj go
        if os.path.exists(nazwa):
            return nazwa
        # Jeśli ustawiono osobny folder dla obecności, użyj go
        fol_obc = getattr(self, "folder_zapisu_obecnosci", "")
        if fol_obc:
            try: os.makedirs(fol_obc, exist_ok=True)
            except Exception: pass
            return os.path.join(fol_obc, nazwa)
        return self.sciezka(nazwa)

    def get_plik_obecnosci(self):
        rok = self.obc_rok_var.get() if hasattr(self, 'obc_rok_var') else datetime.now().strftime("%Y")
        return self.get_plik_obecnosci_rok(rok)

    def dni_w_miesiacu(self, rok, miesiac):
        import calendar
        try:
            return calendar.monthrange(int(rok), int(miesiac))[1]
        except Exception:
            return 31

    def wiersz_obecnosci(self, rok, miesiac, dzien):
        """Wiersz w pliku Obecności dla konkretnego dnia (etykieta miesiąca + realne dni + SUMA na miesiąc)."""
        r = self.NAGLOWEK_OBC + 1
        for m in range(1, int(miesiac)):
            r += 1 + self.dni_w_miesiacu(rok, m) + 1  # etykieta + dni + SUMA
        r += 1  # etykieta bieżącego miesiąca
        return r + (int(dzien) - 1)

    def wiersz_suma_miesiaca(self, rok, miesiac):
        return self.wiersz_obecnosci(rok, miesiac, self.dni_w_miesiacu(rok, miesiac)) + 1

    def godz_na_dziesietna(self, tekst):
        """Zamienia wpisany czas na godziny dziesietnie (do obliczen). System
        jest MINUTOWY, nie dziesietny: "8:45", "8,45" i "8.45" to wszystko
        8 godzin i 45 MINUT (czyli 8.75h), a nie 8.45 godziny dziesietnej -
        tak jak na zegarku, zeby nie trzeba bylo przeliczac minut na ulamek
        godziny. Dwukropek, przecinek i kropka dzialaja tu identycznie jako
        separator godzin/minut. Sam numer bez separatora (np. "8") to dalej
        pelne godziny."""
        tekst = (tekst or "").strip()
        if not tekst: return None
        try:
            for sep in (":", ",", "."):
                if sep in tekst:
                    h, m = tekst.split(sep, 1)
                    h = h.strip(); m = m.strip() or "0"
                    return int(h) + int(m) / 60.0
            return float(tekst)
        except (ValueError, ZeroDivisionError):
            return None

    def dziesietna_na_godz(self, wartosc):
        if wartosc is None or wartosc == "": return ""
        try: wartosc = float(wartosc)
        except (ValueError, TypeError): return ""
        h = int(wartosc); m = round((wartosc - h) * 60)
        if m == 60: h += 1; m = 0
        return f"{h:02d}:{m:02d}"

    def utworz_plik_obecnosci(self, rok):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "OBECNOSCI"
        ws.cell(row=self.NAGLOWEK_OBC, column=1, value="Dzień")
        for m in range(1, 13):
            etykieta_r = self.wiersz_obecnosci(rok, m, 1) - 1
            ws.cell(row=etykieta_r, column=1, value=self.MIESIACE_OBC[m - 1])
            for d in range(1, self.dni_w_miesiacu(rok, m) + 1):
                ws.cell(row=self.wiersz_obecnosci(rok, m, d), column=1, value=d)
            ws.cell(row=self.wiersz_suma_miesiaca(rok, m), column=1, value="SUMA")
        self._style_obecnosci(ws, rok)
        plik = self.get_plik_obecnosci()
        wb.save(plik)
        wb.close()
        return plik

    def _style_obecnosci(self, ws, rok, miesiac_do_rozwiniecia=None):
        """Nadaje plikowi Obecności kolorowy, czytelny wygląd: nagłówki sekcji/imion,
        naprzemienne tło dni, wyróżnione weekendy, pogrubione wiersze SUMA, oraz
        grupowanie (zwijanie/rozwijanie) miesięcy - domyślnie rozwinięty jest tylko
        bieżący miesiąc kalendarzowy (pozostałe są domyślnie zwinięte), ale eksport
        miesiąca może podać miesiac_do_rozwiniecia, żeby to on był widoczny."""
        import calendar
        if miesiac_do_rozwiniecia:
            miesiac_biezacy = int(miesiac_do_rozwiniecia)
        else:
            try:
                miesiac_biezacy = int(datetime.now().strftime("%m"))
            except Exception:
                miesiac_biezacy = 1

        ostatnia_kolumna = 2
        col = 2
        while ws.cell(row=2, column=col).value:
            ostatnia_kolumna = col + 1
            col += 2

        cienka = Side(style="thin", color="d1d5db")
        ramka = Border(left=cienka, right=cienka, top=cienka, bottom=cienka)
        wyr_centrum = Alignment(horizontal="center", vertical="center")
        KOLORY_SEKCJI = {"bar": "0d9488", "kuchnia": "d97706"}
        KOLOR_SEKCJI_JASNY = {"bar": "ccfbf1", "kuchnia": "fef3c7"}

        kom = ws.cell(row=self.NAGLOWEK_OBC, column=1, value=ws.cell(row=self.NAGLOWEK_OBC, column=1).value or "Dzień")
        kom.font = Font(bold=True, color="ffffff"); kom.fill = PatternFill("solid", fgColor="374151"); kom.alignment = wyr_centrum
        ws.column_dimensions["A"].width = 9

        col = 2
        while ws.cell(row=2, column=col).value:
            sekcja = str(ws.cell(row=1, column=col).value or "bar").strip().lower()
            kol_ciemny = KOLORY_SEKCJI.get(sekcja, "6b7280")
            kol_jasny = KOLOR_SEKCJI_JASNY.get(sekcja, "e5e7eb")
            for c in (col, col + 1):
                k1 = ws.cell(row=1, column=c); k1.fill = PatternFill("solid", fgColor=kol_ciemny)
                k1.font = Font(bold=True, color="ffffff"); k1.alignment = wyr_centrum
                k2 = ws.cell(row=2, column=c); k2.fill = PatternFill("solid", fgColor=kol_jasny)
                k2.font = Font(bold=True, color="1f2933"); k2.alignment = wyr_centrum
                k3 = ws.cell(row=3, column=c); k3.fill = PatternFill("solid", fgColor="f1f5f9")
                k3.font = Font(bold=True, size=8, color="475569"); k3.alignment = wyr_centrum
                ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 7
            col += 2

        for m in range(1, 13):
            etykieta_r = self.wiersz_obecnosci(rok, m, 1) - 1
            pierwszy = self.wiersz_obecnosci(rok, m, 1)
            dni = self.dni_w_miesiacu(rok, m)
            ostatni = self.wiersz_obecnosci(rok, m, dni)
            suma_r = self.wiersz_suma_miesiaca(rok, m)
            rozwiniety = (m == miesiac_biezacy)

            for c in range(1, ostatnia_kolumna + 1):
                kom = ws.cell(row=etykieta_r, column=c)
                kom.fill = PatternFill("solid", fgColor="1e3a8a")
                kom.font = Font(bold=True, color="ffffff", size=11)
            ws.cell(row=etykieta_r, column=1).alignment = Alignment(horizontal="left", vertical="center")

            for d in range(1, dni + 1):
                r = pierwszy + d - 1
                try:
                    jest_weekend = calendar.weekday(int(rok), m, d) >= 5
                except Exception:
                    jest_weekend = False
                tlo = "fff1e6" if jest_weekend else ("ffffff" if d % 2 == 0 else "f8fafc")
                for c in range(1, ostatnia_kolumna + 1):
                    kom = ws.cell(row=r, column=c)
                    kom.fill = PatternFill("solid", fgColor=tlo)
                    kom.border = ramka
                    if c == 1:
                        kom.font = Font(bold=jest_weekend, color="9a3412" if jest_weekend else "1f2933")
                        kom.alignment = wyr_centrum
                ws.row_dimensions[r].outlineLevel = 1
                ws.row_dimensions[r].hidden = not rozwiniety

            for c in range(1, ostatnia_kolumna + 1):
                kom = ws.cell(row=suma_r, column=c)
                kom.fill = PatternFill("solid", fgColor="dcfce7")
                kom.font = Font(bold=True, color="14532d")
                kom.border = Border(top=Side(style="medium", color="16a34a"))
            ws.cell(row=suma_r, column=1).alignment = Alignment(horizontal="left", vertical="center")

        try:
            ws.sheet_properties.outlinePr.summaryBelow = True
        except Exception:
            pass
        ws.freeze_panes = f"B{self.NAGLOWEK_OBC + 1}"

    def mapuj_kolumny_obecnosci(self, ws):
        """Zwraca dict: nazwa(lower) -> (od_col, do_col, sekcja, nazwa_oryginalna), skanując nagłówek."""
        mapa = {}
        col = 2
        while True:
            nazwa = ws.cell(row=2, column=col).value
            if not nazwa: break
            sekcja = (ws.cell(row=1, column=col).value or "bar")
            mapa[str(nazwa).strip().lower()] = (col, col + 1, str(sekcja).strip().lower(), str(nazwa).strip())
            col += 2
        return mapa

    def dodaj_kolumne_pracownika_obecnosci(self, ws, rok, nazwa, sekcja):
        col = 2
        while ws.cell(row=2, column=col).value:
            col += 2
        od_col, do_col = col, col + 1
        od_l, do_l = openpyxl.utils.get_column_letter(od_col), openpyxl.utils.get_column_letter(do_col)

        ws.merge_cells(start_row=1, start_column=od_col, end_row=1, end_column=do_col)
        ws.cell(row=1, column=od_col, value=sekcja)
        ws.merge_cells(start_row=2, start_column=od_col, end_row=2, end_column=do_col)
        ws.cell(row=2, column=od_col, value=nazwa)
        ws.cell(row=3, column=od_col, value="od")
        ws.cell(row=3, column=do_col, value="do")

        for m in range(1, 13):
            suma_r = self.wiersz_suma_miesiaca(rok, m)
            pierwszy = self.wiersz_obecnosci(rok, m, 1)
            ostatni = self.wiersz_obecnosci(rok, m, self.dni_w_miesiacu(rok, m))
            ws.merge_cells(start_row=suma_r, start_column=od_col, end_row=suma_r, end_column=do_col)
            ws.cell(row=suma_r, column=od_col,
                    value=f"=SUM({do_l}{pierwszy}:{do_l}{ostatni})-SUM({od_l}{pierwszy}:{od_l}{ostatni})")
        return od_col, do_col

    def buduj_zakladke_obecnosci(self):
        now = datetime.now()
        _zw_top, frame_top = self.zrob_karte(self.tab_obecnosci, self.txt_wybor_daty_obc)
        _zw_top.pack(fill="x", padx=20, pady=10)

        etyk_font = (self.czcionka_nazwa, 10)
        pole_font = (self.czcionka_nazwa, 11)
        tk.Label(frame_top, text="Rok:", font=etyk_font, bg=self.karta_bg, fg=self.tekst_muted).grid(row=0, column=0, padx=(0, 6), pady=8)
        self.obc_rok_var = tk.StringVar(value=now.strftime("%Y")); tk.Entry(frame_top, textvariable=self.obc_rok_var, width=5, font=pole_font, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, relief="flat", highlightthickness=1, highlightbackground=self.karta_border).grid(row=0, column=1, ipady=4, padx=(0, 16))

        tk.Label(frame_top, text="Msc (01-12):", font=etyk_font, bg=self.karta_bg, fg=self.tekst_muted).grid(row=0, column=2, padx=(0, 6))
        self.obc_miesiac_var = tk.StringVar(value=now.strftime("%m")); tk.Entry(frame_top, textvariable=self.obc_miesiac_var, width=4, font=pole_font, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, relief="flat", highlightthickness=1, highlightbackground=self.karta_border).grid(row=0, column=3, ipady=4, padx=(0, 16))

        tk.Label(frame_top, text="Dzień:", font=etyk_font, bg=self.karta_bg, fg=self.tekst_muted).grid(row=0, column=4, padx=(0, 6))
        self.obc_dzien_var = tk.StringVar(value=now.strftime("%d")); tk.Entry(frame_top, textvariable=self.obc_dzien_var, width=4, font=pole_font, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, relief="flat", highlightthickness=1, highlightbackground=self.karta_border).grid(row=0, column=5, ipady=4, padx=(0, 20))

        self.pill(frame_top, text=self.txt_wczytaj_obc_btn, command=self.wczytaj_dzien_obecnosci, bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg, relief="flat", cursor="hand2", font=(self.czcionka_nazwa, 10, "bold")).grid(row=0, column=6, padx=6, ipady=4)
        self.pill(frame_top, text=self.txt_wyczysc_obc_btn, command=self.wyczysc_obecnosci, bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg, relief="flat", cursor="hand2").grid(row=0, column=7, padx=6, ipady=4)
        self.pill(frame_top, text="Odśwież wygląd", command=self.odswiez_wyglad_obecnosci, bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg, relief="flat", cursor="hand2").grid(row=0, column=8, padx=6, ipady=4)

        self._zw_obc, self.frame_obecnosci = self.zrob_karte(self.tab_obecnosci, self.txt_godziny_naglowek)
        self._zw_obc.pack(fill="both", expand=True, padx=20, pady=10)

        nags = ["Pracownik", "Sekcja", "Od", "Do", "Godziny"]
        for i, n in enumerate(nags):
            tk.Label(self.frame_obecnosci, text=n, font=(self.czcionka_nazwa, 9, "bold"), bg=self.karta_bg, fg=self.tekst_muted).grid(row=0, column=i, padx=8, pady=(0, 10), sticky="w")

        self.wiersze_obecnosci = []

        fb = tk.Frame(self.tab_obecnosci, bg=self.kolor_tla); fb.pack(fill="x", padx=20, pady=(4, 0))
        self.pill(fb, text=self.txt_dodaj_obc_btn, command=self.dodaj_wiersz_obecnosci, bg=self.kolor_akcent, fg=self.kolor_akcent_fg, relief="flat", cursor="hand2", font=(self.czcionka_nazwa, 10, "bold")).pack(side=tk.LEFT, padx=(0, 8), ipady=4)
        self.pill(fb, text=self.txt_usun_wiersz_btn, command=self.usun_wiersz_obecnosci, bg="#fee2e2", fg="#991b1b", relief="flat", cursor="hand2").pack(side=tk.LEFT, ipady=4)

        bf = tk.Frame(self.tab_obecnosci, bg=self.kolor_tla); bf.pack(fill="x", padx=20, pady=20)
        self.btn_zapisz_obecnosci = self.pill(bf, text=self.txt_zapisz_obc_btn, command=self.zapisz_dzien_obecnosci, bg=self.kolor_akcent, fg=self.kolor_akcent_fg, relief="flat", cursor="hand2", font=(self.czcionka_nazwa, 13, "bold"), height=2)
        self.btn_zapisz_obecnosci.pack(fill="x")

        self.dodaj_wiersz_obecnosci()
        self.dodaj_wiersz_obecnosci()

        # ---- Podsumowanie miesiaca (zywe dane z bazy KOS) + eksport plikow ----
        _zw_pods, frame_pods = self.zrob_karte(self.tab_obecnosci, "Podsumowanie miesiąca (z bazy KOS)")
        _zw_pods.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        wiersz_pods = tk.Frame(frame_pods, bg=self.karta_bg)
        wiersz_pods.pack(fill="x", pady=(0, 10))

        tk.Label(wiersz_pods, text="Pracownik:", font=etyk_font, bg=self.karta_bg, fg=self.tekst_muted).grid(row=0, column=0, padx=(0, 6))
        self.pods_pracownik_var = tk.StringVar(value="")
        self.pods_pracownik_cb = ttk.Combobox(wiersz_pods, textvariable=self.pods_pracownik_var, values=getattr(self, "lista_kelnerow_raport", []), width=22)
        self.pods_pracownik_cb.grid(row=0, column=1, padx=(0, 16), ipady=2)

        tk.Label(wiersz_pods, text="Rok:", font=etyk_font, bg=self.karta_bg, fg=self.tekst_muted).grid(row=0, column=2, padx=(0, 6))
        self.pods_rok_var = tk.StringVar(value=now.strftime("%Y"))
        tk.Entry(wiersz_pods, textvariable=self.pods_rok_var, width=5, font=pole_font, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, relief="flat", highlightthickness=1, highlightbackground=self.karta_border).grid(row=0, column=3, ipady=4, padx=(0, 16))

        tk.Label(wiersz_pods, text="Msc (01-12):", font=etyk_font, bg=self.karta_bg, fg=self.tekst_muted).grid(row=0, column=4, padx=(0, 6))
        self.pods_miesiac_var = tk.StringVar(value=now.strftime("%m"))
        tk.Entry(wiersz_pods, textvariable=self.pods_miesiac_var, width=4, font=pole_font, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, relief="flat", highlightthickness=1, highlightbackground=self.karta_border).grid(row=0, column=5, ipady=4, padx=(0, 16))

        self.pill(wiersz_pods, text="Pokaż", command=self.pokaz_podsumowanie_miesiaca_kos, bg=self.kolor_akcent, fg=self.kolor_akcent_fg, relief="flat", cursor="hand2", font=(self.czcionka_nazwa, 10, "bold")).grid(row=0, column=6, padx=(0, 6), ipady=4)
        self.pill(wiersz_pods, text="Eksportuj pliki miesiąca", command=self.eksportuj_pliki_miesiaca_kos, bg=self.kolor_wygaszony_bg, fg=self.kolor_wygaszony_fg, relief="flat", cursor="hand2").grid(row=0, column=7, padx=6, ipady=4)

        self.txt_podsumowanie_godzin = tk.Text(frame_pods, height=12, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, relief="flat", highlightthickness=1, highlightbackground=self.karta_border, font=("Courier New", 10), wrap="none")
        self.txt_podsumowanie_godzin.pack(fill="both", expand=True)
        self.txt_podsumowanie_godzin.insert("1.0", "Wybierz pracownika, rok i miesiąc, a potem kliknij „Pokaż”.")
        self.txt_podsumowanie_godzin.config(state="disabled")

    def dodaj_wiersz_obecnosci(self, domyslny_pracownik=None, domyslna_sekcja="bar", domyslny_od="", domyslny_do=""):
        idx = len(self.wiersze_obecnosci) + 1

        cb = ttk.Combobox(self.frame_obecnosci, values=self.lista_kelnerow, width=16)
        if domyslny_pracownik: cb.set(domyslny_pracownik)
        cb.grid(row=idx, column=0, padx=8, pady=6, sticky="w")
        cb.bind('<KeyRelease>', lambda e, c=cb: self.autocomplete_pracownik(e, c))

        # Sekcja (bar/kuchnia/admin) nie jest juz wybierana recznie - wynika
        # automatycznie z przypisania pracownika w KOS (Grafik pracy na
        # stronie, patrz self.sekcja_kos_map / odswiez_liste_pracownikow),
        # z zapasowym spadkiem na lokalna tabelke "Funkcje pracownikow" w
        # Ustawieniach (jest_kuchnia/jest_bar/jest_admin), gdy kogos nie ma na
        # liscie z KOS. Pole zostaje jako wyszarzony podglad, bo dalej w
        # kodzie (zapis do KOS, podzial napiwkow) nadal korzysta z tej
        # wartosci. "admin" wyklucza dana osobe z podzialu serwisu/napiwkow i
        # z listy raportow zmianowych (Rozliczenie Dnia).
        sekcja_cb = ttk.Combobox(self.frame_obecnosci, values=["bar", "kuchnia", "admin"], width=8, state="disabled")
        sekcja_cb.set(domyslna_sekcja); sekcja_cb.grid(row=idx, column=1, padx=8, pady=6)

        def _auto_sekcja(event=None, c=cb, sc=sekcja_cb):
            _, nazwa = self._rozbij_nr_nazwe(c.get())
            nazwa = nazwa.strip()
            if not nazwa:
                return
            nowa = getattr(self, "sekcja_kos_map", {}).get(nazwa.lower())
            if nowa not in ("bar", "kuchnia", "admin"):
                if self.jest_kuchnia(nazwa):
                    nowa = "kuchnia"
                elif self.jest_bar(nazwa):
                    nowa = "bar"
                elif self.jest_admin(nazwa):
                    nowa = "admin"
                else:
                    return  # nieznana rola - zostaw domyslna/wczytana wartosc
            sc.config(state="normal"); sc.set(nowa); sc.config(state="disabled")

        cb.bind("<<ComboboxSelected>>", _auto_sekcja)
        cb.bind("<FocusOut>", _auto_sekcja)

        od_e = tk.Entry(self.frame_obecnosci, width=8, font=(self.czcionka_nazwa, 11), bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, relief="flat", highlightthickness=1, highlightbackground=self.karta_border); od_e.insert(0, domyslny_od); od_e.grid(row=idx, column=2, padx=8, pady=6, ipady=3)
        do_e = tk.Entry(self.frame_obecnosci, width=8, font=(self.czcionka_nazwa, 11), bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, relief="flat", highlightthickness=1, highlightbackground=self.karta_border); do_e.insert(0, domyslny_do); do_e.grid(row=idx, column=3, padx=8, pady=6, ipady=3)
        lbl_g = tk.Label(self.frame_obecnosci, text="---", font=(self.czcionka_nazwa, 10, "bold"), bg=self.karta_bg, fg=self.kolor_czcionki, width=8)
        lbl_g.grid(row=idx, column=4, padx=8, pady=6)

        def przelicz(event, oe=od_e, de=do_e, lg=lbl_g):
            o = self.godz_na_dziesietna(oe.get()); d = self.godz_na_dziesietna(de.get())
            lg.config(text=f"{(d - o):.2f}" if (o is not None and d is not None) else "---")

        od_e.bind("<FocusOut>", przelicz); do_e.bind("<FocusOut>", przelicz)
        self.wiersze_obecnosci.append((cb, sekcja_cb, od_e, do_e, lbl_g))
        if domyslny_od and domyslny_do:
            przelicz(None)

    def usun_wiersz_obecnosci(self):
        if self.wiersze_obecnosci:
            for w in self.wiersze_obecnosci.pop():
                w.destroy()

    def wyczysc_obecnosci(self):
        while self.wiersze_obecnosci: self.usun_wiersz_obecnosci()
        self.dodaj_wiersz_obecnosci(); self.dodaj_wiersz_obecnosci()

    def odswiez_wyglad_obecnosci(self):
        """Odświeża kolorowy wygląd pliku Obecności i grupowanie miesięcy (rozwija
        tylko bieżący miesiąc kalendarzowy) - przydatne do 'upiększenia' starszego
        pliku bez kolorów oraz do zwinięcia z powrotem miesięcy ręcznie rozwiniętych w Excelu."""
        rok = self.obc_rok_var.get() if hasattr(self, 'obc_rok_var') else datetime.now().strftime("%Y")
        plik = self.get_plik_obecnosci_rok(rok)
        if not os.path.exists(plik):
            self.utworz_plik_obecnosci(rok)
            return messagebox.showinfo("Wygląd", "Utworzono nowy, kolorowy plik obecności.")
        try:
            wb = openpyxl.load_workbook(plik)
            ws = wb.active
            self._style_obecnosci(ws, rok)
            wb.save(plik)
            wb.close()
            messagebox.showinfo("Wygląd", "Odświeżono wygląd pliku obecności. Rozwinięty jest bieżący miesiąc.")
        except Exception as e:
            messagebox.showerror("Błąd", str(e))

    def pokaz_podsumowanie_miesiaca_kos(self):
        """Dzien-po-dniu podsumowanie godzin pracy wybranego pracownika w
        wybranym miesiacu, czytane na zywo z bazy KOS (nie z lokalnego Excela
        - patrz komentarz w wczytaj_dzien_obecnosci) - zakres calego miesiaca
        zwraca endpoint GET /api/freyonator/obecnosci?rok=..&miesiac=.. (bez
        dzien)."""
        import calendar
        pracownik = (self.pods_pracownik_var.get() or "").strip()
        if not pracownik:
            return messagebox.showerror("Błąd", "Wybierz pracownika.")
        try:
            rok = int(self.pods_rok_var.get()); miesiac = int(self.pods_miesiac_var.get())
            if not (1 <= miesiac <= 12): raise ValueError
        except ValueError:
            return messagebox.showerror("Błąd", "Nieprawidłowy rok/miesiąc.")

        access_token = self._kos_access_token_swiezy()
        if not access_token:
            return messagebox.showerror(
                "Brak połączenia z KOS",
                "Nie udało się połączyć z kontem KOS.\n"
                "Sprawdź internet albo zaloguj się ponownie w Ustawieniach → Konto KOS.",
            )

        status, dane = self._kos_http(
            "GET", f"{KOS_API_URL}/api/freyonator/obecnosci?rok={rok}&miesiac={miesiac}",
            headers={"Authorization": f"Bearer {access_token}"},
        )
        if status != 200:
            szczegol = (dane.get("error") if isinstance(dane, dict) else "") or ""
            tresc = "Nie udało się wczytać podsumowania miesiąca z KOS."
            if status is None:
                tresc += "\n\nBrak połączenia z internetem."
            elif szczegol:
                tresc += f"\n\n{szczegol}"
            return messagebox.showerror("Błąd wczytywania", tresc)

        wszystkie = dane.get("pracownicy") if isinstance(dane, dict) else []
        wiersze = [w for w in (wszystkie or []) if str(w.get("pracownik", "")).strip().lower() == pracownik.lower()]

        dni_tyg = ["Pon", "Wt", "Śr", "Czw", "Pt", "Sob", "Niedz"]
        miesiac_nazwa = self.MIESIACE_OBC[miesiac - 1] if 1 <= miesiac <= 12 else str(miesiac)
        linie = [f"GODZINY PRACY — {pracownik} — {miesiac_nazwa} {rok}", ""]
        if not wiersze:
            linie.append("Brak zapisanych godzin w tym miesiącu dla tego pracownika.")
        else:
            suma = 0.0
            for w in sorted(wiersze, key=lambda x: x.get("data") or ""):
                data_str = w.get("data") or ""
                try:
                    rr, mm, dd = data_str.split("-")
                    etykieta_dnia = f"{dd}.{mm} ({dni_tyg[calendar.weekday(int(rr), int(mm), int(dd))]})"
                except Exception:
                    etykieta_dnia = data_str
                godz_od = self.dziesietna_na_godz(w.get("godz_od")) or "--:--"
                godz_do = self.dziesietna_na_godz(w.get("godz_do")) or "--:--"
                godziny = w.get("godziny") or 0.0
                suma += godziny
                linie.append(f"  {etykieta_dnia:<14} {godz_od}–{godz_do}  = {godziny:.2f}h   ({w.get('sekcja', '')})")
            linie.append("")
            linie.append(f"RAZEM: {suma:.2f}h")

        self.txt_podsumowanie_godzin.config(state="normal")
        self.txt_podsumowanie_godzin.delete("1.0", tk.END)
        self.txt_podsumowanie_godzin.insert("1.0", "\n".join(linie))
        self.txt_podsumowanie_godzin.config(state="disabled")

    def eksportuj_pliki_miesiaca_kos(self):
        """Eksportuje swiezy plik Obecnosci (w tym samym, kolorowym formacie
        co bazowy Obecnosci_<rok>.xlsx) oraz plik Rozliczenia miesiaca - oba
        budowane na zywo z danych w bazie KOS, bo od migracji na Faze 1/2
        lokalne pliki Obecnosci_<rok>.xlsx / Rozliczenie_<miesiac>.xlsx juz
        nie sa przez aplikacje uzywane (patrz komentarze w
        wczytaj_dzien_obecnosci / wczytaj_dzien) - eksport NIE kopiuje tych
        martwych plikow, tylko odtwarza ten sam wyglad ze zrodla danych,
        ktorym faktycznie teraz jest KOS."""
        try:
            rok = int(self.pods_rok_var.get()); miesiac = int(self.pods_miesiac_var.get())
            if not (1 <= miesiac <= 12): raise ValueError
        except ValueError:
            return messagebox.showerror("Błąd", "Nieprawidłowy rok/miesiąc.")

        access_token = self._kos_access_token_swiezy()
        if not access_token:
            return messagebox.showerror(
                "Brak połączenia z KOS",
                "Nie udało się połączyć z kontem KOS.\n"
                "Sprawdź internet albo zaloguj się ponownie w Ustawieniach → Konto KOS.",
            )

        status_o, dane_o = self._kos_http(
            "GET", f"{KOS_API_URL}/api/freyonator/obecnosci?rok={rok}&miesiac={miesiac}",
            headers={"Authorization": f"Bearer {access_token}"},
        )
        if status_o != 200:
            szczegol = (dane_o.get("error") if isinstance(dane_o, dict) else "") or ""
            return messagebox.showerror("Błąd", f"Nie udało się wczytać Obecności z KOS.\n\n{szczegol}")

        status_r, dane_r = self._kos_http(
            "GET", f"{KOS_API_URL}/api/freyonator/dzien-rozliczenia?rok={rok}&miesiac={miesiac}",
            headers={"Authorization": f"Bearer {access_token}"},
        )
        if status_r != 200:
            szczegol = (dane_r.get("error") if isinstance(dane_r, dict) else "") or ""
            return messagebox.showerror("Błąd", f"Nie udało się wczytać Rozliczeń z KOS.\n\n{szczegol}")

        folder = filedialog.askdirectory(title="Wybierz folder do zapisu eksportu")
        if not folder:
            return

        try:
            plik_obc = self._eksportuj_obecnosci_xlsx(folder, rok, miesiac, dane_o.get("pracownicy") or [])
            plik_rozl = self._eksportuj_rozliczenie_xlsx(folder, rok, miesiac, dane_r.get("pracownicy") or [], dane_r.get("weryfikacje") or [])
        except Exception as e:
            return messagebox.showerror("Błąd eksportu", str(e))

        messagebox.showinfo(
            "Eksport zakończony",
            f"Zapisano w folderze:\n{folder}\n\n• {os.path.basename(plik_obc)}\n• {os.path.basename(plik_rozl)}",
        )

    def _eksportuj_obecnosci_xlsx(self, folder, rok, miesiac, wiersze):
        """Buduje plik Obecnosci w tym samym stylu co utworz_plik_obecnosci /
        _style_obecnosci (caly szkielet roku, kolorowe sekcje bar/kuchnia,
        wyroznione weekendy, wiersze SUMA), wypelniony danymi z bazy KOS dla
        wybranego miesiaca - rozwiniety jest eksportowany miesiac."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "OBECNOSCI"
        ws.cell(row=self.NAGLOWEK_OBC, column=1, value="Dzień")
        for m in range(1, 13):
            etykieta_r = self.wiersz_obecnosci(rok, m, 1) - 1
            ws.cell(row=etykieta_r, column=1, value=self.MIESIACE_OBC[m - 1])
            for d in range(1, self.dni_w_miesiacu(rok, m) + 1):
                ws.cell(row=self.wiersz_obecnosci(rok, m, d), column=1, value=d)
            ws.cell(row=self.wiersz_suma_miesiaca(rok, m), column=1, value="SUMA")

        kolumny = {}
        for w in wiersze:
            nazwa = str(w.get("pracownik") or "").strip()
            if not nazwa:
                continue
            sekcja = str(w.get("sekcja") or "bar").strip().lower()
            if nazwa not in kolumny:
                kolumny[nazwa] = self.dodaj_kolumne_pracownika_obecnosci(ws, rok, nazwa, sekcja)
            od_col, do_col = kolumny[nazwa]
            data_str = w.get("data") or ""
            try:
                _, _, dd = data_str.split("-"); dzien = int(dd)
            except Exception:
                continue
            r = self.wiersz_obecnosci(rok, miesiac, dzien)
            if w.get("godz_od") is not None: ws.cell(row=r, column=od_col, value=w.get("godz_od"))
            if w.get("godz_do") is not None: ws.cell(row=r, column=do_col, value=w.get("godz_do"))

        self._style_obecnosci(ws, rok, miesiac_do_rozwiniecia=miesiac)
        miesiac_nazwa = self.MIESIACE_OBC[miesiac - 1].capitalize()
        plik = os.path.join(folder, f"Obecnosci_{rok}_{miesiac:02d}_{miesiac_nazwa}_eksport.xlsx")
        wb.save(plik)
        wb.close()
        return plik

    def _eksportuj_rozliczenie_xlsx(self, folder, rok, miesiac, wiersze, weryfikacje):
        """Buduje plik Rozliczenia miesiaca w czytelnym, kolorowym formacie
        (ten sam jezyk wizualny co eksport Obecnosci: naglowek na ciemnym
        tle, naprzemienne wiersze, zamrozony naglowek) z danych zapisanych w
        bazie KOS. Stary szablon.xlsx (eksport z Numbers, formuly CHOOSE(),
        stala lista pracownikow wpisana w wierszach 419-448) jest martwy od
        migracji na Faze 1 i nie jest tutaj odtwarzany jeden-do-jeden."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "ROZLICZENIA"
        naglowki = ["Data", "Pracownik", "Nr", "VIP", "Karta mob.", "Karta kasa", "Przelewy", "Foodsi", "Serwis", "Razem", "Gotówka", "Zdano"]
        for i, n in enumerate(naglowki, start=1):
            ws.cell(row=1, column=i, value=n)
        for w in sorted(wiersze, key=lambda x: (x.get("data") or "", x.get("nr_pracownika") or 0)):
            ws.append([
                w.get("data"), w.get("pracownik"), w.get("nr_pracownika"), w.get("vip"), w.get("karta_mob"),
                w.get("karta_kasa"), w.get("przelewy"), w.get("foodsi"), w.get("serwis"), w.get("razem"),
                w.get("gotowka"), w.get("zdano"),
            ])
        self._styl_tabela_eksportu(ws, len(naglowki))

        ws2 = wb.create_sheet("WERYFIKACJA DNIA")
        naglowki2 = ["Data", "Fiskal", "Faktury", "Karta mob. (raport)", "Karta kasa (raport)", "Gotówka sprawdzona", "Rozlicza", "Komentarz"]
        for i, n in enumerate(naglowki2, start=1):
            ws2.cell(row=1, column=i, value=n)
        for w in sorted(weryfikacje, key=lambda x: x.get("data") or ""):
            ws2.append([
                w.get("data"), w.get("fiskal"), w.get("faktury"), w.get("karta_mob_raport"), w.get("karta_kasa_raport"),
                w.get("gotowka_sprawdzona"), w.get("osoba"), w.get("komentarz"),
            ])
        self._styl_tabela_eksportu(ws2, len(naglowki2))

        miesiac_nazwa = self.MIESIACE_OBC[miesiac - 1].capitalize()
        plik = os.path.join(folder, f"Rozliczenie_{rok}_{miesiac:02d}_{miesiac_nazwa}_eksport.xlsx")
        wb.save(plik)
        wb.close()
        return plik

    def _styl_tabela_eksportu(self, ws, liczba_kolumn):
        """Prosty, czytelny styl tabeli eksportu (pogrubiony ciemny naglowek,
        obramowanie, naprzemienne tlo wierszy, zamrozony naglowek) - ten sam
        jezyk wizualny co plik Obecnosci."""
        cienka = Side(style="thin", color="d1d5db")
        ramka = Border(left=cienka, right=cienka, top=cienka, bottom=cienka)
        for c in range(1, liczba_kolumn + 1):
            kom = ws.cell(row=1, column=c)
            kom.font = Font(bold=True, color="ffffff")
            kom.fill = PatternFill("solid", fgColor="374151")
            kom.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 14
        for r in range(2, ws.max_row + 1):
            tlo = "ffffff" if r % 2 == 0 else "f8fafc"
            for c in range(1, liczba_kolumn + 1):
                kom = ws.cell(row=r, column=c)
                kom.fill = PatternFill("solid", fgColor=tlo)
                kom.border = ramka
        ws.freeze_panes = "A2"

    def wczytaj_dzien_obecnosci(self):
        """Wczytuje obecnosci dnia z bazy KOS (nie z lokalnego Excela - patrz
        app/api/freyonator/obecnosci). Od Fazy 2 migracji zakladka
        "Obecnosci" juz NIE czyta zadnego pliku .xlsx."""
        try:
            dz = int(self.obc_dzien_var.get())
        except ValueError:
            return messagebox.showerror("Błąd", "Nieprawidłowy dzień.")

        access_token = self._kos_access_token_swiezy()
        if not access_token:
            return messagebox.showerror(
                "Brak połączenia z KOS",
                "Nie udało się połączyć z kontem KOS.\n"
                "Sprawdź internet albo zaloguj się ponownie w Ustawieniach → Konto KOS.",
            )

        status, dane = self._kos_http(
            "GET",
            f"{KOS_API_URL}/api/freyonator/obecnosci?rok={int(self.obc_rok_var.get())}&miesiac={int(self.obc_miesiac_var.get())}&dzien={dz}",
            headers={"Authorization": f"Bearer {access_token}"},
        )
        if status != 200:
            szczegol = (dane.get("error") if isinstance(dane, dict) else "") or ""
            tresc = "Nie udało się wczytać dnia z KOS."
            if status is None:
                tresc += "\n\nBrak połączenia z internetem."
            elif szczegol:
                tresc += f"\n\n{szczegol}"
            return messagebox.showerror("Błąd wczytywania", tresc)

        while self.wiersze_obecnosci: self.usun_wiersz_obecnosci()
        pracownicy = dane.get("pracownicy") if isinstance(dane, dict) else []
        for p in (pracownicy or []):
            self.dodaj_wiersz_obecnosci(
                p.get("pracownik") or "",
                p.get("sekcja") or "bar",
                self.dziesietna_na_godz(p.get("godz_od")),
                self.dziesietna_na_godz(p.get("godz_do")),
            )
        if not pracownicy:
            self.dodaj_wiersz_obecnosci(); self.dodaj_wiersz_obecnosci()
        self._toast("Wczytano obecności dnia" if pracownicy else "Brak zapisanych obecności dla tego dnia", typ="sukces")

    def zapisz_dzien_obecnosci(self):
        """Zapisuje obecnosci dnia bezposrednio do bazy KOS (nie do lokalnego
        Excela - patrz app/api/freyonator/obecnosci). Od Fazy 2 migracji
        zakladka "Obecnosci" juz NIE tworzy/modyfikuje zadnego pliku .xlsx."""
        access_token = self._kos_access_token_swiezy()
        if not access_token:
            return messagebox.showerror(
                "Brak połączenia z KOS",
                "Nie udało się połączyć z kontem KOS.\n"
                "Sprawdź internet albo zaloguj się ponownie w Ustawieniach → Konto KOS.\n"
                "Dzień NIE został zapisany.",
            )
        try:
            dz = int(self.obc_dzien_var.get())
        except ValueError:
            return messagebox.showerror("Błąd", "Nieprawidłowy dzień.")

        pracownicy = []
        for cb, sekcja_cb, od_e, do_e, lbl_g in self.wiersze_obecnosci:
            nazwa = cb.get().strip()
            if not nazwa: continue
            godz_od = self.godz_na_dziesietna(od_e.get())
            godz_do = self.godz_na_dziesietna(do_e.get())
            godziny = 0.0
            if godz_od is not None and godz_do is not None:
                godziny = godz_do - godz_od
                if godziny < 0:
                    # Zmiana przechodzaca przez polnoc (np. 20:00-01:00) - bez
                    # tego godziny wyszlyby ujemne i pracownik znikalby po cichu
                    # z podzialu napiwkow/serwisu (filtr "godz > 0" nizej).
                    godziny += 24.0
            pracownicy.append({
                "pracownik": nazwa, "sekcja": sekcja_cb.get() or "bar",
                "godz_od": godz_od, "godz_do": godz_do, "godziny": godziny,
            })

        body = {"p_rok": int(self.obc_rok_var.get()), "p_miesiac": int(self.obc_miesiac_var.get()), "p_dzien": dz, "p_pracownicy": pracownicy}
        if getattr(self, "btn_zapisz_obecnosci", None):
            self.btn_zapisz_obecnosci.ustaw_aktywny(False)
        try:
            status, dane = self._kos_http(
                "POST",
                f"{KOS_API_URL}/api/freyonator/obecnosci",
                body=body,
                headers={"Content-Type": "application/json", "Authorization": f"Bearer {access_token}"},
            )
        finally:
            if getattr(self, "btn_zapisz_obecnosci", None):
                self.btn_zapisz_obecnosci.ustaw_aktywny(True)

        if status != 200:
            szczegol = (dane.get("error") if isinstance(dane, dict) else "") or ""
            tresc = "Nie udało się zapisać dnia w KOS. Dzień NIE został zapisany."
            if status is None:
                tresc += "\n\nBrak połączenia z internetem."
            elif szczegol:
                tresc += f"\n\n{szczegol}"
            return messagebox.showerror("Błąd zapisu", tresc)

        self._toast("Zapisano obecności dnia ✓", typ="sukces")

    def pokaz_podsumowanie_miesiaca_obecnosci(self):
        rok = self.obc_rok_var.get(); msc = self.obc_miesiac_var.get()
        plik = self.get_plik_obecnosci()
        self.txt_podsumowanie_godzin.config(state="normal")
        self.txt_podsumowanie_godzin.delete("1.0", tk.END)
        if not os.path.exists(plik):
            self.txt_podsumowanie_godzin.insert("1.0", f"Plik {plik} jeszcze nie istnieje.")
            self.txt_podsumowanie_godzin.config(state="disabled")
            return
        try:
            pierwszy = self.wiersz_obecnosci(rok, msc, 1)
            ostatni = self.wiersz_obecnosci(rok, msc, self.dni_w_miesiacu(rok, msc))
            wb = openpyxl.load_workbook(plik, data_only=True)
            ws = wb.active
            mapa = self.mapuj_kolumny_obecnosci(ws)
            wyniki = []
            for klucz, (od_col, do_col, sekcja, nazwa) in mapa.items():
                suma = 0.0
                for r in range(pierwszy, ostatni + 1):
                    ov = ws.cell(row=r, column=od_col).value
                    dv = ws.cell(row=r, column=do_col).value
                    if ov is not None and dv is not None:
                        try: suma += float(dv) - float(ov)
                        except (TypeError, ValueError): pass
                if suma > 0:
                    wyniki.append((nazwa, sekcja, suma))
            wb.close()

            miesiac_nazwa = self.MIESIACE_OBC[int(msc) - 1] if 1 <= int(msc) <= 12 else msc
            linie = [f"GODZINY PRACY — {miesiac_nazwa} {rok}", ""]
            if not wyniki:
                linie.append("Brak zapisanych godzin w tym miesiącu.")
            else:
                for nazwa, sekcja, suma in sorted(wyniki, key=lambda x: (-x[2])):
                    linie.append(f"  {nazwa} ({sekcja}): {suma:.2f}h")
                linie.append("")
                linie.append(f"RAZEM: {sum(s for _, _, s in wyniki):.2f}h")
            self.txt_podsumowanie_godzin.insert("1.0", "\n".join(linie))
        except Exception as e:
            self.txt_podsumowanie_godzin.insert("1.0", f"Błąd: {e}")
        self.txt_podsumowanie_godzin.config(state="disabled")

    def weryfikuj_dane(self, *args):
        calk_gotowka = 0.0; calk_sprawdz = 0.0; suma_razem = 0.0; suma_vip = 0.0; suma_kart_wpisy = 0.0
        
        for cb, vip_e, km_e, kk_e, prz_e, foo_e, serw_e, razem_e, lbl_g, sprawdz_e, lbl_r in self.wiersze_danych:
            nazwa = cb.get().lower()
            v, km, kk, prz, foo, r, spr = map(self.pobierz_float, [vip_e.get(), km_e.get(), kk_e.get(), prz_e.get(), foo_e.get(), razem_e.get(), sprawdz_e.get()])
            suma_kart_wpisy += (km + kk)
            if not self._wiersz_to_kuchnia(nazwa):
                suma_razem += r; suma_vip += v
                got = r - v - (km + kk + prz + foo)
                calk_gotowka += got; calk_sprawdz += spr
                lbl_g.config(text=f"{got:.2f}")
                roz = spr - got
                if spr == 0.0 and got == 0.0: lbl_r.config(text="0.00", fg=self.tekst_muted)
                elif roz == 0: lbl_r.config(text="OK", fg=self.kolor_sukces)
                elif roz > 0: lbl_r.config(text=f"+{roz:.2f}", fg="#7DD3FC")
                else: lbl_r.config(text="---", fg=self.tekst_muted) # Ukrycie manka u pracownika
            else:
                lbl_g.config(text="---"); lbl_r.config(text="---", fg=self.tekst_muted)

        croz = calk_sprawdz - calk_gotowka
        self.lbl_gotowka_suma.config(text=f"CAŁK. OCZEKIWANA: {calk_gotowka:.2f} zł")
        self.lbl_sprawdz_suma.config(text=f"SPRAWDZONO: {calk_sprawdz:.2f} zł")
        if croz == 0 and calk_gotowka > 0: self.lbl_roznica_suma.config(text=f"RÓŻNICA W KASIE: OK", fg=self.kolor_sukces)
        elif croz > 0: self.lbl_roznica_suma.config(text=f"RÓŻNICA W KASIE: +{croz:.2f} zł", fg="#7DD3FC")
        else: self.lbl_roznica_suma.config(text=f"RÓŻNICA W KASIE: ---", fg=self.tekst_muted) # Ukrycie manka głównego
            
        self.lbl_suma_kart.config(text=f"Suma z kart wpisów: {suma_kart_wpisy:.2f} zł")
        fisk, fakt, term_mob, term_kas = map(self.pobierz_float, [self.fiskal_var.get(), self.faktury_var.get(), self.raport_karta_mob_var.get(), self.raport_karta_kasa_var.get()])
        term_suma = term_mob + term_kas
        rfisk = round((suma_razem - suma_vip) - (fisk + fakt), 2)
        rkart = round(suma_kart_wpisy - term_suma, 2)
        msg = ""
        if rfisk == 0 and rkart == 0: self.lbl_weryfikacja.config(text="Fiskalizacja i Terminale OK!", fg=self.kolor_sukces)
        else:
            if rfisk != 0: msg += f"Fiskal BŁĄD: {rfisk:.2f} | "
            if rkart != 0: msg += f"Karty BŁĄD: {rkart:.2f}"
            self.lbl_weryfikacja.config(text=msg, fg=self.kolor_blad)

    def dodaj_wiersz(self, domyslny_pracownik=None):
        if len(self.wiersze_danych) >= self.max_wierszy: return messagebox.showwarning("Limit", "Limit wierszy osiągnięty.")
        wid = len(self.wiersze_danych) + 1
        
        _wiersz_tlo = self.karta_bg if wid % 2 else self._rozjasnij_kolor(self.karta_bg, 3)
        _wej_kw = dict(relief="flat", highlightthickness=1, highlightbackground=self.karta_border, highlightcolor=self.kolor_akcent)

        cb = ttk.Combobox(self.frame_mid, values=self._lista_kelnerow_bez_kuchni(), width=14); cb.grid(row=wid, column=0, padx=4, pady=4)
        if domyslny_pracownik: cb.set(domyslny_pracownik)
        cb.bind('<KeyRelease>', lambda e, c=cb: self.autocomplete_pracownik(e, c, self._lista_kelnerow_bez_kuchni()))
        
        vip = tk.Entry(self.frame_mid, width=8, bg=self.pole_ostrzezenie_bg, fg=self.pole_ostrzezenie_fg, insertbackground=self.pole_ostrzezenie_fg, **_wej_kw); vip.insert(0, "0.0"); vip.grid(row=wid, column=1, padx=4, pady=4)
        kmob = tk.Entry(self.frame_mid, width=8, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, **_wej_kw); kmob.insert(0, "0.0"); kmob.grid(row=wid, column=2, padx=4, pady=4)
        kkasa = tk.Entry(self.frame_mid, width=8, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, **_wej_kw); kkasa.insert(0, "0.0"); kkasa.grid(row=wid, column=3, padx=4, pady=4)
        prz = tk.Entry(self.frame_mid, width=8, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, **_wej_kw); prz.insert(0, "0.0"); prz.grid(row=wid, column=4, padx=4, pady=4)
        foo = tk.Entry(self.frame_mid, width=8, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, **_wej_kw); foo.insert(0, "0.0"); foo.grid(row=wid, column=5, padx=4, pady=4)
        serw = tk.Entry(self.frame_mid, width=8, bg=self.pole_bg, fg=self.pole_fg, insertbackground=self.pole_fg, **_wej_kw); serw.insert(0, "0.0"); serw.grid(row=wid, column=6, padx=4, pady=4)
        razem = tk.Entry(self.frame_mid, width=9, bg=self.pole_info_bg, fg=self.pole_info_fg, insertbackground=self.pole_info_fg, **_wej_kw); razem.insert(0, "0.0"); razem.grid(row=wid, column=7, padx=4, pady=4)
        lbl_g = tk.Label(self.frame_mid, text="0.00", font=(self.czcionka_nazwa, 9, "bold"), fg="#E2E8F0", bg=_wiersz_tlo, width=8); lbl_g.grid(row=wid, column=8, padx=4, pady=4, sticky="nsew")
        spr = tk.Entry(self.frame_mid, width=9, bg=self.pole_sukces_bg, fg=self.pole_sukces_fg, insertbackground=self.pole_sukces_fg, **_wej_kw); spr.insert(0, "0.0"); spr.grid(row=wid, column=9, padx=4, pady=4)
        lbl_r = tk.Label(self.frame_mid, text="0.00", font=(self.czcionka_nazwa, 9, "bold"), fg=self.tekst_muted, bg=_wiersz_tlo, width=8); lbl_r.grid(row=wid, column=10, padx=4, pady=4, sticky="nsew")
        
        for p in [vip, kmob, kkasa, prz, foo, serw, razem, spr]:
            p.bind("<KeyRelease>", self.weryfikuj_dane)
            self.przypnij_czyszczenie_zer(p)
            
        cb.bind("<<ComboboxSelected>>", lambda e: self.obsluga_typu_pracownika(e, cb, [vip, kmob, kkasa, prz, foo, razem, spr], serw))
        self.wiersze_danych.append((cb, vip, kmob, kkasa, prz, foo, serw, razem, lbl_g, spr, lbl_r))
        if domyslny_pracownik: self.obsluga_typu_pracownika(None, cb, [vip, kmob, kkasa, prz, foo, razem, spr], serw)

    def usun_wiersz(self):
        if len(self.wiersze_danych) > 0:
            for w in self.wiersze_danych.pop(): w.destroy()
            self.weryfikuj_dane()

    def wyczysc_formularz(self):
        while self.wiersze_danych: self.usun_wiersz()
        for v in [self.fiskal_var, self.raport_karta_mob_var, self.raport_karta_kasa_var, self.faktury_var]: v.set("0.0")
        self.osoba_var.set(""); self.komentarz_txt.delete("1.0", tk.END)
        self._odswiez_licznik_komentarza()
        # Kuchnia i Bar sa stalymi, zbiorczymi wierszami "Raportow Osobowych"
        # (patrz obsluga_typu_pracownika: Kuchnia ma odblokowane tylko pole
        # "Serwis", Bar - wszystko oprocz "Serwis") - maja byc zawsze obecne,
        # wiec odtwarzamy je od razu po kazdym czyszczeniu formularza.
        self.dodaj_wiersz(self.str_kuchnia)
        self.dodaj_wiersz(self.str_bar)

    def lista_wykluczonych_napiwki(self):
        """Wykluczenia z podzialu napiwkow - z dwoch zrodel naraz: lokalna
        lista tekstowa w Ustawieniach (self.wykluczeni_napiwki) ORAZ flaga
        "wykluczony_z_napiwkow" ustawiana na stronie (Rozliczenie Dnia ->
        Wykluczenia z napiwkow, zapisywana w shift_employees i pobierana w
        odswiez_liste_pracownikow). Kazde z tych dwoch zrodel samo wystarczy,
        zeby kogos wykluczyc."""
        lokalni = {n.strip().lower() for n in (self.wykluczeni_napiwki or "").split(",") if n.strip()}
        return list(lokalni | getattr(self, "pracownicy_wykluczeni_kos", set()))

    def zbierz_godziny_obecnosci_dnia(self, rok, miesiac, dzien):
        """Zwraca liste (nazwa, sekcja, godziny) pracownikow z bazy KOS dla
        wskazanego dnia - zasila kalkulator podzialu napiwkow (podziel_serwis,
        oblicz_napiwki_gotowka). Od Fazy 2 migracji Obecnosci zyja wylacznie w
        KOS (patrz wczytaj_dzien_obecnosci/zapisz_dzien_obecnosci) - to samo
        zrodlo trzeba czytac tutaj, zeby kalkulator nie dzielil napiwkow na
        podstawie pustego/nieaktualnego lokalnego pliku Excel."""
        try:
            r, m, d = int(rok), int(miesiac), int(dzien)
        except (TypeError, ValueError):
            return []
        access_token = self._kos_access_token_swiezy()
        if not access_token:
            return []
        status, dane = self._kos_http(
            "GET",
            f"{KOS_API_URL}/api/freyonator/obecnosci?rok={r}&miesiac={m}&dzien={d}",
            headers={"Authorization": f"Bearer {access_token}"},
        )
        if status != 200 or not isinstance(dane, dict):
            return []
        wynik = []
        for p in (dane.get("pracownicy") or []):
            nazwa = str(p.get("pracownik") or "").strip()
            if not nazwa:
                continue
            sekcja = p.get("sekcja") or "bar"
            try:
                godz = float(p.get("godziny"))
            except (TypeError, ValueError):
                godz = 0.0
            if godz > 0:
                wynik.append((nazwa, sekcja, godz))
        return wynik

    def podziel_serwis(self):
        suma = sum(map(self.pobierz_float, [self.calc_1_var.get(), self.calc_2_var.get(), self.calc_3_var.get()]))
        if suma <= 0:
            return messagebox.showwarning("Kalkulator", "Wpisz kwotę do podziału (Raport 1/2/3).")

        wykluczeni = self.lista_wykluczonych_napiwki()
        godziny = [(n, s, g) for n, s, g in self.zbierz_godziny_obecnosci_dnia(self.rok_var.get(), self.miesiac_var.get(), self.dzien_var.get())
                   if n.lower() not in wykluczeni and str(s).strip().lower() != "admin"]
        if not godziny:
            return messagebox.showerror("Kalkulator", "Brak wpisanych godzin w zakładce Obecności dla tego dnia.\nUzupełnij Obecności, żeby podzielić proporcjonalnie.")

        suma_godzin = sum(g for _, _, g in godziny)
        if suma_godzin <= 0:
            return messagebox.showerror("Kalkulator", "Suma godzin wynosi 0 — nie można podzielić proporcjonalnie.")

        kwoty = {n: round(suma * (g / suma_godzin), 2) for n, s, g in godziny}

        wiersze_kuchnia = [w for w in self.wiersze_danych if self._wiersz_to_kuchnia(w[0].get())]
        wiersze_bar = [w for w in self.wiersze_danych if self._wiersz_to_bar(w[0].get())]
        wiersze_kelnerzy = [w for w in self.wiersze_danych if w not in wiersze_kuchnia and w not in wiersze_bar]

        dopasowani = set()
        for w in wiersze_kelnerzy:
            nazwa_wiersza = w[0].get()
            _, _nazwa_czysta_surowa = self._rozbij_nr_nazwe(nazwa_wiersza)
            surowa = _nazwa_czysta_surowa.lower()
            for n, s, g in godziny:
                if n.lower() == surowa:
                    w[6].config(state="normal"); w[6].delete(0, tk.END); w[6].insert(0, str(kwoty[n]))
                    dopasowani.add(n.lower())
                    break

        # "Kuchnia" i "Bar" w Rozliczeniu Dnia to zbiorcze wiersze stanowiska
        # (kod kasowy), a nie konkretna osoba - w ciagu dnia na barze/w
        # kuchni moglo pracowac kilka roznych osob z Obecnosci, wiec ich
        # udzialy sumujemy do jednej kwoty na wspolny wiersz danego
        # stanowiska (bar byl tu wczesniej liczony, ale nigdy nic nie
        # zapisywano do jego wiersza - to byl blad).
        kuchnia_suma = round(sum(kwoty[n] for n, s, g in godziny if s == "kuchnia" and n.lower() not in dopasowani), 2)
        for w in wiersze_kuchnia:
            w[6].config(state="normal"); w[6].delete(0, tk.END); w[6].insert(0, str(kuchnia_suma))

        bar_suma = round(sum(kwoty[n] for n, s, g in godziny if s == "bar" and n.lower() not in dopasowani), 2)
        for w in wiersze_bar:
            w[6].config(state="normal"); w[6].delete(0, tk.END); w[6].insert(0, str(bar_suma))

        self.weryfikuj_dane()
        podsumowanie = "\n".join(f"{n} ({s}): {g:.2f}h -> {kwoty[n]:.2f} zł" for n, s, g in sorted(godziny, key=lambda x: -x[2]))
        niedopasowani = [n for n, s, g in godziny if s not in ("kuchnia", "bar") and n.lower() not in dopasowani]
        info_extra = f"\n\nNie znaleziono w tabeli kelnerów wiersza dla: {', '.join(niedopasowani)} (ich udział nie został nigdzie wpisany)" if niedopasowani else ""
        messagebox.showinfo("Kalkulator", f"Podzielono proporcjonalnie do godzin (suma {suma_godzin:.2f}h):\n\n{podsumowanie}{info_extra}")

    def oblicz_napiwki_gotowka(self):
        pula = self.pobierz_float(self.gotowka_pula_var.get())
        if pula <= 0:
            return messagebox.showwarning("Kalkulator gotówki", "Wpisz kwotę gotówki do podziału.")

        wykluczeni = self.lista_wykluczonych_napiwki()
        godziny = [(n, s, g) for n, s, g in self.zbierz_godziny_obecnosci_dnia(self.rok_var.get(), self.miesiac_var.get(), self.dzien_var.get())
                   if n.lower() not in wykluczeni and str(s).strip().lower() != "admin"]

        # Wspolna pula dla WSZYSTKICH (kuchnia + bar/kelnerzy razem) - dzielona
        # proporcjonalnie do godzin, bez osobnego "podzialu 25/75" dla kuchni.
        suma_godzin = sum(g for _, _, g in godziny)

        linie = [f"PULA GOTÓWKI: {pula:.2f} zł  (wspólna pula dla wszystkich, w tym kuchni)", ""]

        if godziny and suma_godzin > 0:
            linie.append(f"--- WSZYSCY ({suma_godzin:.2f}h łącznie) ---")
            for n, s, g in sorted(godziny, key=lambda x: -x[2]):
                etykieta = " [kuchnia]" if s == "kuchnia" else ""
                linie.append(f"  {n}{etykieta}: {g:.2f}h -> {(pula * g / suma_godzin):.2f} zł")
            razem = pula
        else:
            linie.append("  Brak godzin w Obecnościach dla tego dnia — nie można podzielić proporcjonalnie.")
            razem = 0.0

        linie.append("")
        linie.append(f"RAZEM WYPŁACONE: {razem:.2f} zł")
        linie.append("")
        linie.append("Wynik nie jest zapisywany do Excela ani nigdzie indziej — tylko podgląd.")

        self.txt_wynik_gotowka.config(state="normal", height=max(3, min(22, len(linie))))
        self.txt_wynik_gotowka.delete("1.0", tk.END)
        self.txt_wynik_gotowka.insert("1.0", "\n".join(linie))
        self.txt_wynik_gotowka.config(state="disabled")

    def uzupelnij_kelnerow_z_obecnosci(self, rok, miesiac, dzien):
        """Dodaje w Rozliczeniu Dnia wiersze dla pracowników wpisanych w zakładce Obecności
        danego dnia (bar/kelnerzy), którzy jeszcze nie mają swojego wiersza. Kuchnia liczona
        jest zbiorczo (wg godzin) i nie potrzebuje osobnego wiersza. Admin nie bierze udziału
        w rozliczeniach zmianowych w ogóle - nigdy nie dostaje tu własnego wiersza."""
        try:
            obecni = self.zbierz_godziny_obecnosci_dnia(rok, miesiac, dzien)
        except Exception:
            return
        if not obecni:
            return
        def _nazwa_wiersza(w):
            n = w[0].get()
            _, nazwa_czysta = self._rozbij_nr_nazwe(n)
            return nazwa_czysta.lower()
        istniejace = {_nazwa_wiersza(w) for w in self.wiersze_danych}
        for nazwa, sekcja, godz in obecni:
            if str(sekcja).strip().lower() in ("kuchnia", "admin"):
                continue
            if nazwa.strip().lower() in istniejace:
                continue
            wpis = next((k for k in self.lista_kelnerow if self._rozbij_nr_nazwe(k)[1].lower() == nazwa.strip().lower()), nazwa)
            self.dodaj_wiersz(wpis)
            istniejace.add(nazwa.strip().lower())

    def wygeneruj_plik_kolejnego_miesiaca(self):
        """Tworzy z szablonu pusty plik rozliczen na miesiac nastepujacy po tym,
        ktory jest aktualnie wybrany w polach Rok/Msc (przydatne np. pod koniec
        miesiaca, zeby przygotowac plik z wyprzedzeniem)."""
        try:
            msc_akt = int(self.miesiac_var.get())
            rok_akt = int(self.rok_var.get())
        except ValueError:
            return messagebox.showerror("Błąd", "Nieprawidłowy rok/miesiąc w polach powyżej.")
        msc_nast, rok_nast = (1, rok_akt + 1) if msc_akt >= 12 else (msc_akt + 1, rok_akt)
        nazwa_msc = self.slownik_miesiecy.get(f"{msc_nast:02d}", "MAJ")
        plik = self.sciezka(self.format_nazwy_pliku.replace("{miesiac}", nazwa_msc))
        if os.path.exists(plik):
            return messagebox.showinfo("Info", f"Plik na {nazwa_msc} {rok_nast} już istnieje:\n{plik}")
        szablon = self.sciezka(self.szablon_plik)
        if not os.path.exists(szablon):
            return messagebox.showerror("Błąd", "Brak szablonu.")
        try:
            shutil.copy(szablon, plik)
            wb = openpyxl.load_workbook(plik)
            ws = wb.active
            ws["G2"] = nazwa_msc
            wb.save(plik)
            wb.close()
        except Exception as e:
            return messagebox.showerror("Błąd", str(e))
        messagebox.showinfo("Gotowe", f"Utworzono plik na {nazwa_msc} {rok_nast}:\n{plik}")

    def wczytaj_dzien(self):
        """Wczytuje rozliczenie dnia z bazy KOS (nie z lokalnego Excela - patrz
        _kos_wczytaj_dzien, od Fazy 1 migracji "Rozliczenie Dnia" dane dnia
        zyja wylacznie w KOS)."""
        dz = int(self.dzien_var.get())
        access_token = self._kos_access_token_swiezy()
        if not access_token:
            return messagebox.showerror(
                "Brak połączenia z KOS",
                "Nie udało się połączyć z kontem KOS.\n"
                "Sprawdź internet albo zaloguj się ponownie w Ustawieniach → Konto KOS.",
            )
        url = (
            f"{KOS_API_URL}/api/freyonator/dzien-rozliczenia"
            f"?rok={int(self.rok_var.get())}&miesiac={int(self.miesiac_var.get())}&dzien={dz}"
        )
        status, dane = self._kos_http("GET", url, headers={"Authorization": f"Bearer {access_token}"})
        if status != 200 or not isinstance(dane, dict):
            szczegol = (dane.get("error") if isinstance(dane, dict) else "") or ""
            tresc = "Nie udało się wczytać dnia z KOS."
            if status is None:
                tresc += "\n\nBrak połączenia z internetem."
            elif szczegol:
                tresc += f"\n\n{szczegol}"
            return messagebox.showerror("Błąd", tresc)

        self.wyczysc_formularz()
        # wyczysc_formularz() juz odtworzyl stale wiersze Kuchnia+Bar (zawsze
        # pierwsze dwa, patrz obsluga_typu_pracownika) - bierzemy je wprost,
        # zamiast dodawac drugi raz.
        w_kuchnia = self.wiersze_danych[0]
        w_bar = self.wiersze_danych[1]

        for p in (dane.get("pracownicy") or []):
            imie = str(p.get("pracownik") or "").strip()
            if not imie:
                continue
            if self._wiersz_to_kuchnia(imie):
                w_kuchnia[6].delete(0, tk.END); w_kuchnia[6].insert(0, f"{self.bezpieczny_float(p.get('serwis')):.2f}")
                continue
            if self._wiersz_to_bar(imie):
                for pole, klucz in ((1, 'vip'), (2, 'karta_mob'), (3, 'karta_kasa'), (4, 'przelewy'), (5, 'foodsi'), (7, 'razem'), (9, 'zdano')):
                    w_bar[pole].delete(0, tk.END); w_bar[pole].insert(0, f"{self.bezpieczny_float(p.get(klucz)):.2f}")
                continue
            nr = p.get("nr_pracownika")
            wpis = f"{int(nr)} - {imie}" if isinstance(nr, (int, float)) else imie
            self.dodaj_wiersz(wpis); w = self.wiersze_danych[-1]

            w[7].delete(0, tk.END); w[7].insert(0, f"{self.bezpieczny_float(p.get('razem')):.2f}")
            w[1].delete(0, tk.END); w[1].insert(0, f"{self.bezpieczny_float(p.get('vip')):.2f}")
            w[2].delete(0, tk.END); w[2].insert(0, f"{self.bezpieczny_float(p.get('karta_mob')):.2f}")
            w[3].delete(0, tk.END); w[3].insert(0, f"{self.bezpieczny_float(p.get('karta_kasa')):.2f}")
            w[4].delete(0, tk.END); w[4].insert(0, f"{self.bezpieczny_float(p.get('przelewy')):.2f}")
            w[5].delete(0, tk.END); w[5].insert(0, f"{self.bezpieczny_float(p.get('foodsi')):.2f}")
            w[6].delete(0, tk.END); w[6].insert(0, f"{self.bezpieczny_float(p.get('serwis')):.2f}")
            w[9].delete(0, tk.END); w[9].insert(0, f"{self.bezpieczny_float(p.get('zdano')):.2f}")

        self.uzupelnij_kelnerow_z_obecnosci(self.rok_var.get(), self.miesiac_var.get(), dz)

        wer = dane.get("weryfikacja") or {}
        self.fiskal_var.set(str(wer.get("fiskal") if wer.get("fiskal") is not None else "0.0"))
        self.faktury_var.set(str(wer.get("faktury") if wer.get("faktury") is not None else "0.0"))
        self.raport_karta_mob_var.set(str(wer.get("karta_mob_raport") if wer.get("karta_mob_raport") is not None else "0.0"))
        self.raport_karta_kasa_var.set(str(wer.get("karta_kasa_raport") if wer.get("karta_kasa_raport") is not None else "0.0"))
        self.osoba_var.set(str(wer.get("osoba") or ""))
        self.komentarz_txt.delete("1.0", tk.END)
        self.komentarz_txt.insert("1.0", str(wer.get("komentarz") or ""))
        self._odswiez_licznik_komentarza()
        self.weryfikuj_dane()
        messagebox.showinfo("OK", "Załadowano z KOS.")

    def zapisz_excel(self):
        """Zapisuje rozliczenie dnia bezposrednio do bazy KOS (nie do lokalnego
        Excela - patrz app/api/freyonator/dzien-rozliczenia). Od Fazy 1
        migracji zakladka "Rozliczenie Dnia" juz NIE tworzy/modyfikuje zadnego
        pliku .xlsx - dane dnia zyja wylacznie w KOS, zeby nie dalo sie ich
        skopiowac z dysku. Obecnosci i Podsumowanie Miesiac na razie zostaja
        przy Excelu (kolejne fazy)."""
        osoba = self.osoba_var.get().strip()
        if not osoba or osoba == "Wybierz/Wpisz...": return messagebox.showwarning("Wymagane pole", "Wybierz rozliczającego!")
        dz = int(self.dzien_var.get())

        komentarz = self.komentarz_txt.get("1.0", tk.END).strip()
        if len(komentarz) < MIN_DLUGOSC_KOMENTARZA:
            return messagebox.showwarning(
                "Komentarz zbyt krótki",
                f"Komentarz / raport zmianowy musi mieć minimum {MIN_DLUGOSC_KOMENTARZA} znaków "
                f"(obecnie: {len(komentarz)}).\nOpisz dokładniej przebieg dnia - zapisuje się razem z dniem w KOS.",
            )

        access_token = self._kos_access_token_swiezy()
        if not access_token:
            return messagebox.showerror(
                "Brak połączenia z KOS",
                "Nie udało się połączyć z kontem KOS.\n"
                "Sprawdź internet albo zaloguj się ponownie w Ustawieniach → Konto KOS.\n"
                "Dzień NIE został zapisany.",
            )

        pracownicy = []
        suma_sprawdz = 0.0
        for w in self.wiersze_danych:
            n = w[0].get()
            nr, im = self._rozbij_nr_nazwe(n)
            vip, km, kk, prz, foo, serw, raz = map(
                self.pobierz_float,
                [w[1].get(), w[2].get(), w[3].get(), w[4].get(), w[5].get(), w[6].get(), w[7].get()],
            )
            zdano = self.pobierz_float(w[9].get())
            if self._wiersz_to_kuchnia(n):
                # Tak jak w starym Excelu: pracownicy kuchni nie maja zadnych
                # danych finansowych poza serwisem.
                vip = km = kk = prz = foo = raz = got = zdano = 0.0
            else:
                got = raz - vip - (km + kk) - prz - foo
                suma_sprawdz += zdano
            pracownicy.append({
                "pracownik": im,
                "nr_pracownika": int(nr) if str(nr).isdigit() else None,
                "vip": vip, "karta_mob": km, "karta_kasa": kk,
                "przelewy": prz, "foodsi": foo, "serwis": serw,
                "razem": raz, "gotowka": got, "zdano": zdano,
            })

        fisk, fakt, tmob, tkas = map(
            self.pobierz_float,
            [self.fiskal_var.get(), self.faktury_var.get(), self.raport_karta_mob_var.get(), self.raport_karta_kasa_var.get()],
        )
        weryfikacja = {
            "fiskal": fisk, "faktury": fakt,
            "karta_mob_raport": tmob, "karta_kasa_raport": tkas,
            "gotowka_sprawdzona": suma_sprawdz,
            "osoba": osoba,
            "komentarz": komentarz,
        }

        body = {
            "p_rok": int(self.rok_var.get()),
            "p_miesiac": int(self.miesiac_var.get()),
            "p_dzien": dz,
            "p_pracownicy": pracownicy,
            "p_weryfikacja": weryfikacja,
        }
        if getattr(self, "btn_zapisz_dzien", None):
            self.btn_zapisz_dzien.ustaw_aktywny(False)  # blokada podwojnego klikniecia w trakcie zapisu
        try:
            status, dane = self._kos_http(
                "POST",
                f"{KOS_API_URL}/api/freyonator/dzien-rozliczenia",
                body=body,
                headers={"Content-Type": "application/json", "Authorization": f"Bearer {access_token}"},
            )
        finally:
            if getattr(self, "btn_zapisz_dzien", None):
                self.btn_zapisz_dzien.ustaw_aktywny(True)
        if status != 200:
            szczegol = (dane.get("error") if isinstance(dane, dict) else "") or ""
            tresc = "Nie udało się zapisać dnia w KOS. Dzień NIE został zapisany."
            if status is None:
                tresc += "\n\nBrak połączenia z internetem."
            elif szczegol:
                tresc += f"\n\n{szczegol}"
            return messagebox.showerror("Błąd zapisu", tresc)

        self._toast("Zapisano w KOS ✓", typ="sukces")

    def wyslij_raport_email(self):
        osoba = self.osoba_var.get().strip()
        if not osoba or osoba == "Wybierz/Wpisz...": return messagebox.showwarning("Wymagane pole", "Wybierz rozliczającego!")
        dz, msc, rok = self.dzien_var.get(), self.miesiac_var.get(), self.rok_var.get()
        nazwa_msc = self.slownik_miesiecy.get(msc, msc).capitalize()
        temat = f"{self.nazwa_programu} — Raport dnia {dz}.{msc}.{rok}"
        komentarz = self.komentarz_txt.get("1.0", tk.END).strip()
        if len(komentarz) < MIN_DLUGOSC_KOMENTARZA:
            return messagebox.showwarning("Komentarz zbyt krótki", f"Komentarz / raport zmianowy musi mieć minimum {MIN_DLUGOSC_KOMENTARZA} znaków (obecnie: {len(komentarz)}).\nOpisz dokładniej przebieg dnia.")

        kreska = "━" * 32

        # Modyfikacja podsumowania e-mail - ukryte manko
        croz = self.pobierz_float(self.lbl_sprawdz_suma.cget('text').split()[-2]) - self.pobierz_float(self.lbl_gotowka_suma.cget('text').split()[-2])
        roz_text = "OK" if croz == 0 else (f"+{croz:.2f} zł" if croz > 0 else "---")

        tmob, tkas = self.pobierz_float(self.raport_karta_mob_var.get()), self.pobierz_float(self.raport_karta_kasa_var.get())
        razem_karty = tmob + tkas
        fiskal = self.pobierz_float(self.fiskal_var.get())
        faktury = self.pobierz_float(self.faktury_var.get())

        obecni = [w[0].get() for w in self.wiersze_danych if w[0].get() and w[0].get() != "Wybierz/Wpisz..."]

        body = (
            f"{self.nazwa_programu.upper()} — RAPORT DZIENNY\n"
            f"{dz} {nazwa_msc} {rok}\n"
            f"Rozlicza: {osoba}\n\n"

            f"{kreska}\n"
            f"FINANSE\n"
            f"{kreska}\n"
            f"{self.lbl_gotowka_suma.cget('text')}\n"
            f"{self.lbl_sprawdz_suma.cget('text')}\n"
            f"Różnica w kasie: {roz_text}\n"
            f"Weryfikacja POS: {self.lbl_weryfikacja.cget('text')}\n\n"

            f"{kreska}\n"
            f"TERMINALE I KASA\n"
            f"{kreska}\n"
            f"Fiskal Z:        {fiskal:.2f} zł\n"
            f"Faktury:         {faktury:.2f} zł\n"
            f"Karta Mob.:      {tmob:.2f} zł\n"
            f"Karta Kasa:      {tkas:.2f} zł\n"
            f"RAZEM (Karty):   {razem_karty:.2f} zł\n\n"

            f"{kreska}\n"
            f"ZESPÓŁ ({len(obecni)})\n"
            f"{kreska}\n"
            + ("\n".join(f"• {o}" for o in obecni) if obecni else "Brak.") + "\n\n"

            f"{kreska}\n"
            f"KOMENTARZ / OPIS DNIA\n"
            f"{kreska}\n"
            f"{komentarz}\n\n"

            f"— Wysłano z {self.nazwa_programu}"
        )
        try:
            webbrowser.open(f"https://mail.google.com/mail/?view=cm&fs=1&to={self.adresy_email.strip()}&su={urllib.parse.quote(temat)}&body={urllib.parse.quote(body)}")
        except Exception as e: messagebox.showerror("Błąd", str(e))

    def bezpieczny_float(self, val):
        if val is None or val == "": return 0.0
        try: return float(str(val).replace(",", "."))
        except ValueError: return 0.0

    def pobierz_float(self, val):
        try: return float(val.replace(",", "."))
        except ValueError: return 0.0

    def get_plik_docelowy(self):
        msc = self.slownik_miesiecy.get(self.miesiac_var.get(), "MAJ")
        fmt = getattr(self, 'format_nazwy_var', self).get() if hasattr(self, 'format_nazwy_var') else self.format_nazwy_pliku
        return self.sciezka(fmt.replace("{miesiac}", msc))

    def odswiez_liste_pracownikow(self):
        """Buduje self.lista_kelnerow - domyslnie z KOS (tabela
        shift_employees, ta sama lista co w zakladce "Grafik pracy" na
        stronie - patrz app/api/freyonator/pracownicy), zeby nie trzeba bylo
        prowadzic drugiej, osobnej listy pracownikow. Jesli KOS jest
        niedostepny (brak internetu/tokenu/konta), wraca do starego odczytu
        z lokalnego pliku Excel, zeby ta zakladka nigdy nie zostala bez listy
        pracownikow. Sekcja kuchnia/bar per pracownik (self.sekcja_kos_map)
        idzie wprost z KOS (pole "sekcja" w odpowiedzi pracownicy - to samo
        excluded_from_schedule, co dzieli Zespol w Grafiku pracy na stronie)
        - uzywane w zakladce Obecnosci do automatycznego przypisania sekcji.
        Mapa ma klucze pod PELNYM imieniem i nazwiskiem (tak jak jest w KOS),
        a dodatkowo (zapasowo) pod samym pierwszym imieniem - zeby wpisanie
        np. samego "Daniel" tez trafilo w sekcje, bez koniecznosci wybrania
        pelnej nazwy z podpowiedzi. Gdy dwie osoby maja to samo pierwsze imie
        w roznych sekcjach, ten zapasowy klucz jest pomijany (niejednoznaczny)
        - wtedy trzeba wybrac/wpisac pelne imie i nazwisko. UWAGA: jesli ktos
        jest wpisywany innym imieniem/pseudonimem niz w KOS (np. "Anastazja"
        w programie a "Anastasiia Verchyk" na stronie), dopasowanie sie nie
        uda - trzeba albo wybrac podpowiedz z listy, albo ujednolicic imie
        na stronie (Grafik pracy -> Zespol).
        jest_kuchnia/jest_bar (lokalna tabelka "Funkcje pracownikow" w
        Ustawieniach) sa tylko zapasowym zrodlem, gdy kogos nie ma na
        liscie z KOS (np. KOS niedostepny)."""
        self.lista_kelnerow = []; self.str_kuchnia = f"1 - {self.klucz_kuchnia.capitalize()}"; self.str_bar = f"17 - {self.klucz_bar.capitalize()}"
        self.sekcja_kos_map = {}

        pobrano_z_kos = False
        try:
            access_token = self._kos_access_token_swiezy()
            if access_token:
                status, dane = self._kos_http(
                    "GET", f"{KOS_API_URL}/api/freyonator/pracownicy",
                    headers={"Authorization": f"Bearer {access_token}"},
                )
                if status == 200 and isinstance(dane, dict):
                    wykluczeni_kos = set()
                    imiona_po_pierwszym = {}  # pierwsze imie (male) -> sekcja, albo False gdy niejednoznaczne (dwie osoby, rozne sekcje)
                    for p in (dane.get("pracownicy") or []):
                        imie = str(p.get("imie") or "").strip()
                        if not imie:
                            continue
                        self.lista_kelnerow.append(imie)
                        sekcja_p = p.get("sekcja")
                        if sekcja_p in ("bar", "kuchnia"):
                            self.sekcja_kos_map[imie.lower()] = sekcja_p
                            # Zapasowy klucz po samym pierwszym imieniu (bez nazwiska) -
                            # w Obecnosciach ludzie czesto wpisuja/wybieraja tylko imie
                            # (np. "Daniel" zamiast pelnego "Daniel Belavezha"). Gdy dwie
                            # osoby maja to samo pierwsze imie w roznych sekcjach (np. dwaj
                            # "Michal"), klucz zostaje wylaczony (False), zeby nie zgadywac.
                            pierwsze = imie.split()[0].lower() if imie.split() else ""
                            if pierwsze:
                                if pierwsze not in imiona_po_pierwszym:
                                    imiona_po_pierwszym[pierwsze] = sekcja_p
                                elif imiona_po_pierwszym[pierwsze] != sekcja_p:
                                    imiona_po_pierwszym[pierwsze] = False
                        if p.get("wykluczony_z_napiwkow"):
                            wykluczeni_kos.add(imie.lower())
                    for pierwsze, sekcja_p in imiona_po_pierwszym.items():
                        if sekcja_p and pierwsze not in self.sekcja_kos_map:
                            self.sekcja_kos_map[pierwsze] = sekcja_p
                    pobrano_z_kos = bool(self.lista_kelnerow)
                    if pobrano_z_kos:
                        self.pracownicy_wykluczeni_kos = wykluczeni_kos
        except Exception:
            pobrano_z_kos = False

        if not pobrano_z_kos:
            plik = self.get_plik_docelowy(); plik = plik if os.path.exists(plik) else self.sciezka(self.szablon_plik)
            if os.path.exists(plik):
                try:
                    wb = openpyxl.load_workbook(plik, data_only=True); ws = wb.active
                    for row in range(419, 450):
                        nr, imie = ws[f"D{row}"].value, ws[f"E{row}"].value
                        if imie and isinstance(imie, str) and len(imie.strip()) > 1:
                            wpis = f"{int(nr)} - {imie.strip()}" if nr is not None else imie.strip()
                            self.lista_kelnerow.append(wpis)
                            if self.klucz_kuchnia.lower() in imie.lower(): self.str_kuchnia = wpis
                            elif self.klucz_bar.lower() in imie.lower(): self.str_bar = wpis
                    wb.close()
                except Exception: pass
        if not self.lista_kelnerow: self.lista_kelnerow = [self.str_kuchnia, self.str_bar, "Wybierz/Wpisz..."]
        self.lista_kelnerow_raport = [w for w in self.lista_kelnerow if w != "Wybierz/Wpisz..."]
        self._rebuduj_tabele_funkcji_pracownikow()
        if hasattr(self, 'osoba_cb'): self.osoba_cb['values'] = self.lista_kelnerow_raport
        if hasattr(self, 'pods_pracownik_cb'): self.pods_pracownik_cb['values'] = self.lista_kelnerow_raport
        for cb, sekcja_cb, od_e, do_e, lbl_g in getattr(self, 'wiersze_obecnosci', []):
            cb['values'] = self.lista_kelnerow
        if hasattr(self, 'usun_pracownik_cb'): self.usun_pracownik_cb['values'] = self.lista_kelnerow
        self.wyczysc_formularz()

    def dodaj_pracownika_do_excela(self):
        imie = self.nowy_pracownik_var.get().strip()
        if not imie:
            return messagebox.showwarning("Wymagane pole", "Podaj imię i nazwisko pracownika!")
        funkcja = self.nowy_pracownik_funkcja_var.get()
        plik = self.get_plik_docelowy(); plik = plik if os.path.exists(plik) else self.sciezka(self.szablon_plik)
        if not os.path.exists(plik):
            return messagebox.showerror("Błąd", f"Nie znaleziono pliku:\n{plik}")
        try:
            wb = openpyxl.load_workbook(plik); ws = wb.active
            wolny_wiersz, maks_nr = None, 0
            for row in range(419, 450):
                imie_ist, nr_ist = ws[f"E{row}"].value, ws[f"D{row}"].value
                if isinstance(nr_ist, (int, float)): maks_nr = max(maks_nr, int(nr_ist))
                if wolny_wiersz is None and (not imie_ist or not str(imie_ist).strip()): wolny_wiersz = row
            if wolny_wiersz is None:
                wb.close()
                return messagebox.showerror("Brak miejsca", "Lista pracowników w Excelu jest pełna (wiersze 419-450).")
            ws[f"D{wolny_wiersz}"] = maks_nr + 1
            ws[f"E{wolny_wiersz}"] = imie
            wb.save(plik); wb.close()
        except Exception as e:
            return messagebox.showerror("Błąd zapisu", str(e))

        if funkcja == "Kuchnia":
            lista = [n.strip() for n in self.pracownicy_kuchnia.split(",") if n.strip()]
            if imie not in lista: lista.append(imie)
            self.pracownicy_kuchnia = ", ".join(lista)
            if hasattr(self, "pracownicy_kuchnia_var"): self.pracownicy_kuchnia_var.set(self.pracownicy_kuchnia)
        elif funkcja == "Bar":
            lista = [n.strip() for n in self.pracownicy_bar.split(",") if n.strip()]
            if imie not in lista: lista.append(imie)
            self.pracownicy_bar = ", ".join(lista)
            if hasattr(self, "pracownicy_bar_var"): self.pracownicy_bar_var.set(self.pracownicy_bar)

        self._zapisz_json_ustawien()
        self.nowy_pracownik_var.set("")
        self.odswiez_liste_pracownikow()
        messagebox.showinfo("Dodano", f"Dodano pracownika: {imie} ({funkcja})")

    def usun_pracownika_z_excela(self):
        wpis = self.usun_pracownik_var.get().strip()
        if not wpis or wpis in ("Wybierz/Wpisz...",):
            return messagebox.showwarning("Wymagane pole", "Wybierz pracownika do usunięcia!")
        imie = wpis.split(" - ", 1)[1].strip() if " - " in wpis else wpis.strip()
        if not messagebox.askyesno("Potwierdzenie", f"Czy na pewno usunąć pracownika:\n{imie}?"):
            return
        plik = self.get_plik_docelowy(); plik = plik if os.path.exists(plik) else self.sciezka(self.szablon_plik)
        if not os.path.exists(plik):
            return messagebox.showerror("Błąd", f"Nie znaleziono pliku:\n{plik}")
        try:
            wb = openpyxl.load_workbook(plik); ws = wb.active
            usunieto = False
            for row in range(419, 450):
                imie_ist = ws[f"E{row}"].value
                if imie_ist and str(imie_ist).strip().lower() == imie.lower():
                    ws[f"D{row}"] = None
                    ws[f"E{row}"] = None
                    usunieto = True
                    break
            if not usunieto:
                wb.close()
                return messagebox.showerror("Nie znaleziono", f"Nie znaleziono pracownika '{imie}' w arkuszu Excel.")
            wb.save(plik); wb.close()
        except Exception as e:
            return messagebox.showerror("Błąd zapisu", str(e))
        for attr, widget_attr in [("pracownicy_kuchnia", "pracownicy_kuchnia_var"), ("pracownicy_bar", "pracownicy_bar_var")]:
            lista = [n.strip() for n in getattr(self, attr, "").split(",") if n.strip()]
            if imie in lista:
                lista.remove(imie)
                setattr(self, attr, ", ".join(lista))
                if hasattr(self, widget_attr): getattr(self, widget_attr).set(getattr(self, attr))
        self._zapisz_json_ustawien()
        self.odswiez_liste_pracownikow()
        if hasattr(self, 'usun_pracownik_cb'):
            self.usun_pracownik_cb['values'] = self.lista_kelnerow
        self.usun_pracownik_var.set("")
        messagebox.showinfo("Usunięto", f"Usunięto pracownika: {imie}")

    def autocomplete_pracownik(self, event, cb, lista=None):
        if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Return', 'Tab', 'Escape'): return
        baza = self.lista_kelnerow if lista is None else lista
        wpis = cb.get()
        dopasowania = baza if wpis == '' else [i for i in baza if wpis.lower() in i.lower()]
        cb['values'] = dopasowania
        if wpis and dopasowania:
            cb.event_generate('<Down>')

    def obsluga_typu_pracownika(self, event, cb, pola, serwis_ent):
        wartosc = cb.get().lower()
        if self._wiersz_to_kuchnia(wartosc):
            for p in pola: p.config(state="disabled")
            serwis_ent.config(state="normal")
        elif self._wiersz_to_bar(wartosc):
            for p in pola: p.config(state="normal")
            serwis_ent.delete(0, tk.END); serwis_ent.insert(0, "0.0"); serwis_ent.config(state="disabled")
        else:
            for p in pola: p.config(state="normal")
            serwis_ent.config(state="normal")
        self.weryfikuj_dane()

def _pokaz_ekran_ladowania(root):
    """Krotki ekran powitalny na starcie programu - gra slow: nazwa
    "KawiarniaOS" przeczytana na glos brzmi jak ostrzezenie "Kawiarnia...
    OSY!" (klasyczny letni problem kazdej kawiarni z cukrem na blacie),
    wiec puszczamy kilka latajacych os, zanim pokaze sie glowne okno."""
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.configure(bg="#1E2433")
    szer, wys = 440, 230
    x = (splash.winfo_screenwidth() - szer) // 2
    y = (splash.winfo_screenheight() - wys) // 2
    splash.geometry(f"{szer}x{wys}+{x}+{y}")

    tk.Label(
        splash, text="☕ KawiarniaOS", font=("Helvetica", 24, "bold"),
        bg="#1E2433", fg="white",
    ).pack(pady=(32, 6))
    tk.Label(
        splash, text="Kawiarnia... OSY! 🐝  (lecą po cukier, nie po dane)",
        font=("Helvetica", 11, "italic"), bg="#1E2433", fg="#8B96AD",
    ).pack(pady=(0, 16))

    canvas = tk.Canvas(splash, width=szer, height=50, bg="#1E2433", highlightthickness=0)
    canvas.pack()
    osy = [
        canvas.create_text(-20 - i * 70, 25, text="🐝", font=("Helvetica", 20))
        for i in range(3)
    ]

    stan = {"zakonczono": False}

    def zamknij():
        if stan["zakonczono"]:
            return
        stan["zakonczono"] = True
        try:
            splash.destroy()
        except tk.TclError:
            pass

    def lec(klatka=0):
        if stan["zakonczono"]:
            return
        for osa in osy:
            canvas.move(osa, 7, 0)
        if klatka < 65:
            splash.after(16, lambda: lec(klatka + 1))
        else:
            zamknij()

    splash.after(100, lec)
    splash.after(1800, zamknij)  # bezpiecznik, gdyby animacja sie zacielа
    splash.update()
    splash.wait_window(splash)


def _globalny_handler_bledow(root, app):
    """Bezpiecznik na nieoczekiwane bledy w obsludze przyciskow/zdarzen Tk.
    Domyslnie Tkinter w takiej sytuacji NIC nie wyswietla uzytkownikowi - po
    cichu wypisuje traceback na stderr/stdout i konczy obsluge zdarzenia,
    co z perspektywy uzytkownika wyglada jak klikniecie, ktore nic nie robi
    (np. "Zapisz obecnosci" / "Podziel serwis" bez zadnego komunikatu). Ten
    handler zamiast tego: dociaga okno na wierzch, pokazuje czerwone okno z
    opisem bledu i dopisuje go do log_sieci.txt, zeby bylo wiadomo co
    konkretnie poszlo nie tak."""
    def obsluz(exc_type, exc_value, exc_tb):
        import traceback as _tb
        szczegoly = "".join(_tb.format_exception(exc_type, exc_value, exc_tb))
        try:
            app._log_siec(f"BLAD (nieobsluzony w interfejsie): {exc_type.__name__}: {exc_value}\n{szczegoly}")
        except Exception:
            pass
        try:
            root.lift()
            root.attributes("-topmost", True)
            root.after(200, lambda: root.attributes("-topmost", False))
            root.focus_force()
        except Exception:
            pass
        try:
            messagebox.showerror(
                "Nieoczekiwany błąd - akcja NIE została wykonana",
                "Coś poszło nie tak i ostatnia akcja (np. zapis/podział) najpewniej "
                "NIE została wykonana.\n\n"
                f"{exc_type.__name__}: {exc_value}\n\n"
                "Szczegóły zapisano do log_sieci.txt. Spróbuj ponownie; jeśli się "
                "powtarza, przekaż ten komunikat.",
            )
        except Exception:
            pass
    root.report_callback_exception = obsluz


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    _pokaz_ekran_ladowania(root)
    root.deiconify()
    app = KawiarniaOSApp(root)
    _globalny_handler_bledow(root, app)
    try:
        root.mainloop()
    except tk.TclError:
        pass
