### Kotolizator 3000
Program do analizy strony Schoniska na Paluchu. Umożliwia utworzenie zestawień kotów i psów w schonisku plus zrobienie podstawowych statystyk.
Program używa jedynie danych dostępnych publicznie na stronie https://napaluchu.waw.pl/.  Zadaniem programu jest ułatwienie życia wolontariuszom schroniska

#### Oszczędzanie połączeń
Dane są brane wyłącznie ze stron z listą zwierząt (`/zwierzeta/koty-do-adopcji/` lub `/zwierzeta/psy-do-adopcji/` oraz ich wersji `plec-samica/` i `plec-samiec/` – z nich pochodzi płeć). Strony poszczególnych zwierząt nie są pobierane, więc jedno uruchomienie to 3 zapytania plus zdjęcia nowych zwierząt.
Miniatury zdjęć są zapisywane w pamięci podręcznej (`zdjecia.kotolizer` w katalogu danych programu) i pobierane tylko raz.
Pobieranie korzysta z jednej sesji HTTP, robi przerwy z losowym odstępem między zapytaniami i ponawia je z rosnącym opóźnieniem przy błędach 429/5xx. Jeśli strona zacznie blokować zapytania (403/429), pozostałe zdjęcia są pomijane.
