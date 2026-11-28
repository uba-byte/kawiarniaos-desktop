"""
Prototyp "opcji (c)" z audytu aplikacji desktopowej KawiarniaOS: cienki
wrapper pywebview wokol strony KOS (https://kos-kawiarniaos.vercel.app/kasa-desktop)
zamiast wlasnego interfejsu w Tkinterze. Logowanie, stylowanie i cala logika
"Kasy" zyja w jednym miejscu - na stronie webowej KOS - wiec nie trzeba ich
duplikowac ani utrzymywac osobno dla desktopu.

To jest swiadomie MALY plik - cala robota jest po stronie strony webowej
(app/kasa-desktop/page.tsx w repo fv-strona-www). Ten skrypt tylko otwiera
natywne okno systemowe wskazujace na ten adres.

Wymaga pakietu "pywebview" (pip install pywebview). Na Windows pywebview
korzysta domyslnie z silnika Edge WebView2 (ten sam silnik co Edge/Chrome),
wiec wyglad i zachowanie strony sa identyczne jak w zwyklej przegladarce.

Uruchomienie: python kawiarniaOS_webview.py
"""

import sys

try:
    import webview
except ImportError:
    print(
        "Brak pakietu 'pywebview'. Zainstaluj go poleceniem:\n"
        "    pip install pywebview\n"
    )
    sys.exit(1)

# Adres strony KOS z prototypem zakladki Kasa. W docelowej wersji (po
# ewentualnym przeniesieniu kolejnych zakladek) ten adres bedzie wskazywal
# na pelny panel "desktop", a nie tylko Kase.
ADRES_KASA = "https://kos-kawiarniaos.vercel.app/kasa-desktop"

TYTUL_OKNA = "KawiarniaOS — Kasa (prototyp web)"


def main():
    webview.create_window(
        TYTUL_OKNA,
        ADRES_KASA,
        width=960,
        height=760,
        min_size=(720, 600),
    )
    webview.start()


if __name__ == "__main__":
    main()
