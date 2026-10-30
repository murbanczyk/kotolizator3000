### Kotolizator 3000
Program do analizy strony Schoniska na Paluchu. Umożliwia utworzenie zestawień kotów i psów w schonisku plus zrobienie podstawowych statystyk.
Program używa jedynie danych dostępnych publicznie na stronie https://napaluchu.waw.pl/.  Zadaniem programu jest ułatwienie życia wolontariuszom schroniska

#### Oszczędzanie połączeń
Lista zwierząt (`/zwierzeta/koty-do-adopcji/` lub `/zwierzeta/psy-do-adopcji/`) to jedno zapytanie – są na niej imię, numer, wiek, zdjęcie i oznaczenie „do pilnej adopcji”.
Strona każdego zwierzęcia (płeć, data przyjęcia, lokalizacja, opis, liczba zdjęć) jest pobierana raz i trzymana w pamięci podręcznej przez 7 dni, a miniatury zdjęć są pobierane tylko raz (`cache.kotolizer` w katalogu danych programu).
Pierwsze uruchomienie pobiera wszystko (kilka minut), kolejne – tylko listę i nowe zwierzaki.
Pobieranie korzysta z jednej sesji HTTP, robi przerwy z losowym odstępem między zapytaniami i ponawia je z rosnącym opóźnieniem przy błędach 429/5xx. Jeśli strona zacznie blokować zapytania (403/429), program przestaje pobierać i korzysta z danych z pamięci podręcznej.
