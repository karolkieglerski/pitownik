# pitownik

**Kalkulator PIT-38** dla inwestorów korzystających z **E\*TRADE** (RSU, ESPP),
**XTB**, **Interactive Brokers**, **Trading 212**, **Revolut** lub **eToro**.

Wrzuć raport `.xlsx` z brokera, a aplikacja policzy przychód, koszt uzyskania,
kwotę do opodatkowania i podatek (19%) — z przeliczeniem na PLN po kursach NBP.

**Wersja online (GitHub Pages):** <https://karolkieglerski.github.io/pitownik/>

Wszystko działa **wyłącznie w przeglądarce** — pliki nie opuszczają Twojego komputera.
Jedyne zapytania sieciowe to pobranie kursów walut z publicznego API NBP.

## Funkcje

- **E\*TRADE** — raport *Gains & Losses* (`.xlsx`):
  - sprzedaż **RSU**: przychód = Total Proceeds, koszt = 0
  - sprzedaż **ESPP**: koszt = Acquisition Cost, przeliczony po kursie NBP z dnia zakupu
    (przychód — po kursie z dnia sprzedaży)
  - podsumowanie transakcji, suma główna i podatek 19%, zysk ESPP po opodatkowaniu
  - opłata za przelew ($20) pokazana informacyjnie (nie pomniejsza podstawy opodatkowania)
- **XTB** — wyciąg konta (`.xlsx`):
  - pozycje zamknięte (long i short), operacje gotówkowe, pozycje otwarte
  - automatyczne wykrywanie waluty konta (PLN — bez przeliczania)
  - odsetki od wolnych środków (XTB pobrał już 19% podatku — pokazane informacyjnie)
  - **podsumowanie per rok podatkowy**: przychód / koszt / podatek 19%
- **Interactive Brokers** (a także Lynx, Mexem, CapTrader — ten sam format) —
  Activity Statement (CSV):
  - sekcja `Trades` (wiersze Order/Trade), parowanie FIFO, akcje i ETF
  - dywidendy i podatek u źródła — pokazane informacyjnie
- **Trading 212** — eksport historii (CSV):
  - zlecenia kupna/sprzedaży (market, limit, stop), ułamkowe akcje
  - obsługa GBX (pensy) — przeliczenie na GBP
  - dywidendy i odsetki — informacyjnie
- **Revolut** — wyciąg z rachunku Brokerage (CSV/XLSX):
  - operacje BUY/SELL, dywidendy z podatkiem u źródła
- **eToro** — Account Statement (XLSX/XLSM):
  - arkusz Closed Positions — pozycje długie i krótkie, CFD-y oznaczane wg dźwigni
  - dywidendy z arkusza Dividends — informacyjnie
- **Parowanie FIFO** — dla brokerów bez gotowych par kupno/sprzedaż
  (IBKR, Trading 212, Revolut): transakcje łączone są FIFO per instrument,
  z uwzględnieniem godzin (kupno i sprzedaż tego samego dnia)
- **Kursy NBP** — tabela A, automatycznie, bez klucza API:
  - kurs z dnia poprzedzającego transakcję (w weekendy/święta — ostatni opublikowany,
    z wyszukiwaniem do 7 dni wstecz)
  - zapytania zakresowe (chunki co 365 dni) — brak błędów API przy plikach z wieloma datami
  - wiele walut naraz (np. IBKR) — kursy pobierane osobno dla USD, EUR, GBP…
  - podgląd logu pobieranych kursów w trakcie przeliczania
- **Automatyczne wykrywanie typu raportu** (E\*TRADE, XTB, IBKR, Trading 212,
  Revolut, eToro) po wgraniu pliku `.xlsx` lub `.csv`

## Szybki start

Bez instalacji — wystarczy przeglądarka:

1. Otwórz **<https://karolkieglerski.github.io/pitownik/>** (wersja online na
   GitHub Pages) albo lokalnie `web/index.html`.
2. Przeciągnij raport `.xlsx` z brokera na stronę (lub wybierz plik).
3. Poczekaj na pobranie kursów NBP i gotowe.

Lokalny serwer deweloperski:

```bash
npm install        # zależność dev: SheetJS (używany przez testy)
npm run serve      # http://localhost:8080 → otwórz web/index.html
```

## Testy

```bash
npm test           # node --test tests/test_app.mjs (Node.js ≥ 20)
```

Testy obejmują parsowanie XLSX, logikę podatkową RSU/ESPP, przeliczenia NBP
(mockowane), reguły opłaty za przelew oraz parsowanie i podsumowania XTB.
Szczegóły w [tests/README.md](tests/README.md).

Dane testowe są anonimizowane (fikcyjne „ACME") — można je wygenerować ponownie:

```bash
python3 tests/input/generate_example.py
```

## Struktura projektu

```
web/            # cała aplikacja (statyczna strona, bez bundlera)
  index.html    # interfejs (przeciągnij i upuść, wyniki, log kursów)
  app.js        # logika: parsowanie XLSX/CSV, kursy NBP, obliczenia podatkowe
  style.css     # wygląd
  vendor/       # SheetJS (xlsx.full.min.js) — vendored
tests/          # testy (node --test) i dane testowe
.github/        # CI + wdrożenie na GitHub Pages
```

## Wdrożenie

Push na gałąź `main` uruchamia GitHub Actions: testy (`npm ci && npm test`),
a następnie publikację katalogu `web/` na **GitHub Pages** (wdrożenie następuje
tylko jeśli testy przejdą).

## Technologie

Vanilla JavaScript (ES2020+), HTML/CSS, [SheetJS](https://sheetjs.com/) do parsowania
`.xlsx`, publiczne [API NBP](https://api.nbp.pl/). Bez frameworków, bez bundlera,
bez kluczy API i bez zmiennych środowiskowych.

## Licencja i zastrzeżenie

[MIT](LICENSE) — szczegóły w [OPEN_SOURCE.md](OPEN_SOURCE.md).

⚠️ **Zastrzeżenie:** Program ma charakter wyłącznie informacyjny i pomocniczy.
Autorzy nie ponoszą odpowiedzialności za błędy w obliczeniach ani za konsekwencje
wynikające z użycia programu w deklaracjach podatkowych. Przed złożeniem PIT-38
zweryfikuj obliczenia z doradcą podatkowym.
