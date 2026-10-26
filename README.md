### Kotolizator 3000
Program do analizy strony Schoniska na Paluchu. Umożliwia utworzenie zestawień kotów i psów w schonisku plus zrobienie podstawowych statystyk.
Program używa jedynie danych dostępnych publicznie na stronie https://napaluchu.waw.pl/.  Zadaniem programu jest ułatwienie życia wolontariuszom schroniska

#### Oszczędzanie połączeń
Pobieranie korzysta z jednej sesji HTTP (keep-alive), robi przerwy z losowym odstępem między zapytaniami i ponawia je z rosnącym opóźnieniem przy błędach 429/5xx.
Szczegóły i miniatury zwierząt są zapisywane w pamięci podręcznej (`cache.kotolizer` w katalogu danych programu) i odświeżane co 7 dni, więc przy kolejnych uruchomieniach pobierane są głównie listy oraz nowe zwierzaki.
Jeśli strona zacznie blokować zapytania (403/429), program przerywa pobieranie i korzysta z danych z pamięci podręcznej.
