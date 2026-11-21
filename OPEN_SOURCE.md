# pitownik — Open Source

**pitownik** jest projektem open source, rozpowszechnianym na warunkach
licencji **MIT** (pełny tekst: [LICENSE](LICENSE)).

## Co to oznacza w praktyce

Na licencji MIT możesz swobodnie:

- **używać** programu — również do celów komercyjnych,
- **modyfikować** kod źródłowy według własnych potrzeb,
- **kopiować i dystrybuować** — także w produktach zamkniętych,
- **sprzedawać** kopje lub produkty oparte na tym kodzie.

Jedyny warunek:

- zachowaj informację o prawach autorskich i treść licencji MIT
  (plik [LICENSE](LICENSE)) w dystrybuowanych kopiach lub istotnych częściach
  oprogramowania.

Licencja MIT nie nakłada żadnych innych zobowiązań — nie musisz udostępniać
własnych zmian ani podawać, że korzystasz z tego projektu.

## Brak gwarancji

Oprogramowanie jest dostarczane „tak jak jest", bez jakiejkolwiek gwarancji
— również bez gwarancji poprawności obliczeń podatkowych. Szczegóły w
[sekcji z zastrzeżeniem w aplikacji](web/index.html) oraz w tekście licencji.
Przed złożeniem deklaracji PIT-38 zweryfikuj wyniki z doradcą podatkowym.

## Jak uruchomić projekt lokalnie

```bash
npm install     # zależność: SheetJS (xlsx)
npm test        # testy jednostkowe (Node)
npm run serve   # serwer lokalny na http://localhost:8080
```

## Jak kontrybuować

1. Zrób fork i utwórz gałąź ze zmianami.
2. Upewnij się, że `npm test` przechodzi i nowa funkcjonalność jest pokryta
   testami (`tests/test_app.mjs`).
3. Otwórz pull request z opisem zmian.

Pytania i błędy zgłaszaj przez issue.

## Licencje zależności

- [SheetJS (xlsx)](https://sheetjs.com/) — Apache-2.0 (dystrybuowane lokalnie
  w `web/vendor/xlsx.full.min.js`).
