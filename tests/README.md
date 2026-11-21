# Testy

## Uruchomienie

```bash
source .venv/bin/activate
pip install -r requirements.txt
python -m pytest tests/ -v
```

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

