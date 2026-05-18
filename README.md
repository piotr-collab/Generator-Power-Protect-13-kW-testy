# Generator Power Protect 13 kW - badania i testy

Repozytorium zawiera projekt planu badań i arkusz pomiarowy dla generatora
EcoGenerator Power Protect 13 kW 3F.

## Pliki

- `PP_EcoGenerator_13kW_plan_badan.xlsx` - główny skoroszyt Excel do
  prowadzenia testów, wpisywania pomiarów i podsumowania wyniku FAT.
- `build_pp_13kw_test_workbook.py` - skrypt odtwarzający skoroszyt.

## Zakres arkusza

Skoroszyt obejmuje:

- metryczkę urządzenia, wersje firmware, dane BESS, PV, agregatu i obciążnicy,
- checklistę bezpieczeństwa przed testem, w trakcie testu i po teście,
- dobowy profil obciążenia z kolumnami `Godzina doby` i
  `Obciążenie planowane [kW]`,
- rejestr pomiarów: U, I, P, Hz, cosφ, kVA, THD, SOC, PV, agregat,
  ładowanie BESS, paliwo, temperatury, hałas, alarmy i wynik,
- rejestr paliwa, alarmów, temperatur i interwencji,
- scenariusze: profil dobowy, FAT 25/50/75/100%, asymetria faz,
  skoki obciążenia, E-STOP, restart i symulacja końca paliwa,
- podsumowanie wyników i decyzję końcową.

## Założenia operacyjne

Test zakłada ręczną obsługę wyłącznie w oknie dziennym. Godziny nocne są
zaplanowane jako automatyczny zapis danych i monitoring bez ręcznej obsługi.
Maksymalna moc testowa EcoGeneratora wynosi 13 kW, nawet jeśli obciążnica ma
większy zapas mocy.
