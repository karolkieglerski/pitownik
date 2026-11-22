# Testy

## Uruchomienie

```bash
npm install
npm test          # node --test tests/test_app.mjs (Node.js ≥ 20)
```

> Historyczna notka: starszy zestaw `tests/test_main.py` (pytest) odwoływał się
> do usuniętego katalogu `cli/` i nie jest już częścią CI. Aktywna suita testowa
> to `tests/test_app.mjs`.

## Dane testowe

Testy korzystają z zanonimizowanego pliku `tests/input/example_2025.xlsx` wygenerowanego przez `tests/input/generate_example.py`. Plik zawiera fikcyjne dane firmy **ACME**:

- **2 transakcje RSU** — sprzedane 10.04.2025 @ $52/szt (50 + 30 akcji)
- **1 transakcja ESPP** — kupione 05.06.2025, sprzedane 10.06.2025 @ $53/szt (120 akcji)
- **1 transakcja ESPP** — kupione 05.12.2025, sprzedane 10.12.2025 @ $48/szt (80 akcji)

Regeneracja pliku:
```bash
python tests/input/generate_example.py
```

## Pokrycie testami

| Grupa                | Testy | Co pokrywają                                                      |
|----------------------|-------|-------------------------------------------------------------------|
| **ReadTransactions** | 6     | Parsowanie xlsx, filtrowanie Sell vs Summary, wymagane pola       |
| **FormatDate**       | 3     | Konwersja dat na format DD.MM.YYYY                                |
| **ParseDate**        | 3     | Parsowanie datetime i stringów MM/DD/YYYY                         |
| **GroupBySaleDate**  | 4     | Grupowanie pakietów sprzedaży, kolejność, ilości                  |
| **RSUTaxLogic**      | 2     | Przychód RSU = Total Proceeds, koszt = 0                          |
| **ESPPTaxLogic**     | 4     | Przychód, koszt, kwota do opodatkowania, Purchase Date            |
| **NBPRates**         | 4     | Pobieranie kursu, retry na weekendy, deduplikacja dat (mockowane) |
| **PLNConversion**    | 2     | Różne kursy NBP dla przychodu vs kosztu ESPP                      |
| **WireFee**          | 1     | Opłata $20 nie pomniejsza kwoty do opodatkowania                  |

**Łącznie: 29 testów**

## Obsługa XTB (testy Node)

Wsparcie dla raportów XTB (`web/app.js`) jest testowane w `tests/test_app.mjs`
(nie wymaga pliku zewnętrznego — fixture XTB budowany jest w pamięci przez
SheetJS na podstawie wzorca prawdziwego raportu `Closed Positions` /
`Cash Operations` / `Open Positions`).

| Grupa                 | Testy | Co pokrywają                                                    |
|-----------------------|-------|-----------------------------------------------------------------|
| **ExcelSerialToIso**  | 3     | Konwersja dat Excela (serial) na ISO, część UTC                  |
| **XtbDateToIso**      | 4     | Serial, stringi z godziną, MM/DD/YYYY, wartości puste            |
| **IsXtbWorkbook**     | 2     | Auto-detekcja raportu XTB vs Gains & Losses z E*TRADE            |
| **ReadXtbReport**     | 6     | Parsowanie pozycji, pomijanie wiersza Profit/loss, waluta konta, |
| |               | | odsetki i pobrany podatek, mapowanie nóg BUY/SELL (krótka sprzedaż) |
| **XtbGroupBySaleDate**| 1     | Grupowanie wg daty sprzedaży, kolejność chronologiczna           |
| **XtbRate**           | 2     | Kurs NBP z mapy dla walut obcych, kurs 1 dla kont PLN            |
| **XtbYearSummaries**  | 6     | Przychód/koszt po własnych kursach NBP, podatek 19% (zaokrąglenie),|
| |               | | strata = brak podatku, konta PLN, podział na lata podatkowe       |

Zasady liczenia (jak dla E*TRADE): przychód przeliczany kursem NBP z dnia
poprzedzającego sprzedaż, koszt — kursem z dnia poprzedzającego zakup,
podatek 19% od nadwyżki przychodów nad kosztami (zaokrąglony do 2 miejsc).

## Obsługa wielu brokerów (IBKR, Trading 212, Revolut, eToro)

Parsery dodatkowych brokerów (`web/app.js`) są testowane w `tests/test_app.mjs`
na syntetycznych fixture'ach odwzorowujących realne eksporty (sekcjonowany CSV
z IBKR, kolumny zmienne w Trading 212, wiersze operacji Revolut, arkusze
wyciągu eToro budowane w pamięci przez SheetJS).

| Grupa                        | Testy | Co pokrywają                                                   |
|------------------------------|-------|----------------------------------------------------------------|
| **ParseAmount**              | 4     | Prefiksy walut, separatory tysięcy/dziesiętne, wartości ujemne |
| **ParseDateAny/TimeAny/Norm**| 4     | ISO z czasem, „YYYY-MM-DD, HH:MM:SS", MM/DD/YYYY, serial, GBX  |
| **MatchLegsFifo**            | 5     | FIFO z dopasowaniami cząstkowymi, krótka sprzedaż, opłaty,     |
| | | | pomijanie niedomkniętych nóg, kolejność tego samego dnia          |
| **IBKR Activity Statement**  | 4     | Detekcja, pomijanie Closed Lot/SubTotal/Total, dywidendy i      |
| | | | podatek u źródła, pomijanie opcji/futuresów                       |
| **Trading 212 export**       | 4     | Detekcja, parowanie buy/sell, GBX→GBP, dywidendy, pomijanie     |
| | | | wpłat                                                            |
| **Revolut statement**        | 2     | Detekcja, BUY/SELL z opłatami w koszcie, dywidendy, top-upy     |
| **eToro account statement**  | 4     | Detekcja workbooka, long/short, pomijanie otwartych pozycji,    |
| | | | waluta konta, dywidendy (brutto = netto + podatek)               |
| **RateFor/BrokerYearSummaries**| 3   | Kursy wielowalutowe („CUR\|ISO"), przeliczenie każdej strony    |
| | | | po własnym kursie, strata = brak podatku                          |

Zasady liczenia (jak dla XTB): przychód przeliczany kursem NBP z dnia
poprzedzającego sprzedaż, koszt — kursem z dnia poprzedzającego zakup.
Prowizje kupna i sprzedaży doliczane są do kosztów uzyskania. Dywidendy
i odsetki prezentowane są wyłącznie informacyjnie (osobne opodatkowanie).

