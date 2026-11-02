### Kotolizator 3000
Program do analizy strony Schroniska na Paluchu. Umożliwia utworzenie zestawień kotów i psów w schronisku plus zrobienie podstawowych statystyk.
Program używa jedynie danych dostępnych publicznie na stronie https://napaluchu.waw.pl/.  Zadaniem programu jest ułatwienie życia wolontariuszom schroniska.

#### Co potrafi
- pobiera aktualną listę kotów lub psów do adopcji ze strony schroniska razem z ostatnio znalezionymi zwierzętami i zapisuje ją pod dzisiejszą datą; koty przebywające w punkcie Ochota na Kota są oznaczane,
- **Przeglądaj i szukaj** – tabela wszystkich zwierząt z wyszukiwarką (imię, numer, wiek, płeć, miejsce, status – np. „Nie do adopcji” albo „znaleziony”) i sortowaniem po kliknięciu nagłówka; widoczne wiersze można zapisać do Excela,
- **Najdłużej czekające** – ta sama tabela posortowana według liczby dni w schronisku (warto je promować),
- lista zwierząt w Wordzie (ze zdjęciami, pilne adopcje na czerwono, Ochota na Kota na fioletowo), krótka tabelka w Excelu, histogramy wieku i liczba zwierząt w poszczególnych miejscach,
- porównanie dwóch dni: które zwierzęta są nowe, a które zniknęły z listy (np. zostały adoptowane),
- wykres liczby zwierząt w czasie ze wszystkich zapisanych dni.

#### Instalacja i uruchomienie
Potrzebny jest Python 3.8 lub nowszy.
```
pip install -r requirements.txt
python Kotolizator3000.py
```
W Anacondzie wxPython najłatwiej zainstalować poleceniem `conda install -c conda-forge wxpython`.

Zapisane dni i pamięć podręczna są w katalogu `kotolizator` w:
- macOS: `~/Library/Application Support`,
- Windows: `C:\Users\<użytkownik>\AppData\Roaming`,
- Linux: `~/.local/share`.

#### Testy
```
pip install -r requirements-dev.txt
python -m pytest
```
Testy parsera korzystają z zapisanych fragmentów prawdziwych stron w `tests/fixtures`. Gdy strona schroniska zmieni wygląd, wystarczy zapisać nowe strony do tego katalogu i uruchomić testy, żeby zobaczyć, co trzeba poprawić.

#### Aplikacja dla osób bez Pythona
Gotową aplikację (macOS: `Kotolizator3000.app`, Windows: katalog z `Kotolizator3000.exe`) buduje się na danym systemie poleceniem:
```
pip install -r requirements-dev.txt
pyinstaller --noconfirm --windowed --name Kotolizator3000 --icon analizatorkotow/icon.ico --add-data "analizatorkotow/splash.png:analizatorkotow" --add-data "analizatorkotow/icon.ico:analizatorkotow" Kotolizator3000.py
```
Wynik jest w katalogu `dist`. Aplikację dla macOS trzeba zbudować na Macu, a dla Windows na Windows.
Na macOS przy pierwszym uruchomieniu niepodpisanej aplikacji trzeba kliknąć ją prawym przyciskiem i wybrać „Otwórz”.

#### Oszczędzanie połączeń
Lista zwierząt (`/zwierzeta/koty-do-adopcji/` lub `/zwierzeta/psy-do-adopcji/`) to jedno zapytanie – są na niej imię, numer, wiek, zdjęcie i oznaczenie „do pilnej adopcji”.
Do tego jedno zapytanie o ostatnio znalezione zwierzęta danego gatunku (`/zwierzeta/ostatnio-znalezione/gatunek-kot/` lub `gatunek-pies/`), a dla kotów jedno o stronę `/o-schronisku/ochota-na-kota/`, na której są koty przebywające w tym punkcie.
Strona każdego zwierzęcia (płeć, data przyjęcia, lokalizacja, status, opis, liczba zdjęć) jest pobierana raz i trzymana w pamięci podręcznej przez 7 dni, a miniatury zdjęć są pobierane tylko raz (`cache.kotolizer` w katalogu danych programu).
Pierwsze uruchomienie pobiera wszystko (kilka minut), kolejne – tylko listę i nowe zwierzaki.
Pobieranie korzysta z jednej sesji HTTP, robi przerwy z losowym odstępem między zapytaniami i ponawia je z rosnącym opóźnieniem przy błędach 429/5xx. Jeśli strona zacznie blokować zapytania (403/429), program przestaje pobierać i korzysta z danych z pamięci podręcznej.
