import bz2
import pickle
import random
import re
import time
from io import BytesIO
from urllib.parse import urljoin

import requests
from PIL import Image
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = 'https://napaluchu.waw.pl/'

# listy zwierzat do adopcji (cala lista miesci sie na jednej stronie)
LIST_URLS = {
    'koty': BASE_URL + 'zwierzeta/koty-do-adopcji/',
    'psy': BASE_URL + 'zwierzeta/psy-do-adopcji/',
}

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/124.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'pl-PL,pl;q=0.9,en-US;q=0.6,en;q=0.5',
    'Connection': 'keep-alive',
}


class BlockedError(Exception):
    """Strona odmawia dostepu (403/429) mimo ponawiania."""


class PaluchScraper:
    """Pobiera dane ze strony schroniska oszczedzajac polaczenia.

    Imie, numer, wiek, zdjecie i "pilne" sa na stronie z lista zwierzat (1 zapytanie).
    Plec, data przyjecia, lokalizacja, opis i liczba zdjec sa tylko na stronie
    zwierzecia - ta jest pobierana raz i trzymana w pamieci podrecznej przez
    details_days dni. Miniatury zdjec sa pobierane tylko raz.

    - jedna sesja HTTP (keep-alive), naglowki przegladarki i Referer,
    - odstepy z losowym jitterem i ponawianie z backoffem
      (z respektowaniem Retry-After) przy 429/5xx,
    - po 403/429 nic wiecej nie jest pobierane; uzywane sa dane z pamieci
      podrecznej (nawet starsze), a brakujace zdjecia zastepuje pusty obrazek.
    """

    DETAIL_LABELS = ('Typ rasy', 'Płeć', 'Data przyjęcia', 'Wiek', 'Identyfikator', 'Grupa Zwierząt',
                     'Waga', 'Lokalizacja (boks)', 'Kastracja')
    DESC_END = ('Zapraszamy do', 'Planujesz adopcję', 'Poradnik:', 'Zapraszamy do zapoznania się z ankietą')

    def __init__(self, cache_path, details_days=7, delay=(1.0, 2.5)):
        self.cache_path = str(cache_path)
        self.details_max_age = details_days * 24 * 3600
        self.delay = delay
        self._last_request = 0.0
        self.blocked = False
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        retry = Retry(total=4, backoff_factor=2, status_forcelist=(429, 500, 502, 503, 504),
                      allowed_methods=('GET',), respect_retry_after_header=True,
                      raise_on_status=False)
        adapter = HTTPAdapter(max_retries=retry, pool_connections=2, pool_maxsize=1)
        self.session.mount('https://', adapter)
        self.session.mount('http://', adapter)
        cache = self._load_cache()
        self.images = cache.get('images', {})    # url zdjecia -> miniatura w PNG
        self.details = cache.get('details', {})  # link do zwierzecia -> dane ze strony zwierzecia

    # --- pamiec podreczna ---

    def _load_cache(self):
        try:
            with bz2.open(self.cache_path, 'rb') as f:
                return pickle.load(f)
        except Exception:
            return {}

    def save_cache(self):
        try:
            with bz2.open(self.cache_path, 'wb') as f:
                pickle.dump({'images': self.images, 'details': self.details}, f)
        except Exception as e:
            print('Nie udalo sie zapisac cache: ' + str(e))

    # --- HTTP ---

    def _get(self, url, referer=None):
        wait = random.uniform(*self.delay) - (time.time() - self._last_request)
        if wait > 0:
            time.sleep(wait)
        headers = {'Referer': referer} if referer else None
        try:
            response = self.session.get(url, headers=headers, timeout=30)
        finally:
            self._last_request = time.time()
        if response.status_code in (403, 429):
            raise BlockedError('%d dla %s' % (response.status_code, url))
        response.raise_for_status()
        return response

    def getsoup(self, url, referer=None):
        return BeautifulSoup(self._get(url, referer).text, 'lxml')

    # --- listy zwierzat ---

    @staticmethod
    def parse_cards(soup):
        """Zwraca liste slownikow z danymi zwierzat z kart na stronie listy."""
        animals = []
        for card in soup.select('li.post-card'):
            title = card.find('h3')
            link = title.find('a', href=True) if title else None
            if link is None or '/animal/' not in link['href']:
                continue
            text = link.get_text(' ', strip=True)
            m = re.match(r'(?:(.*?)\s*[\u2013\u2014-]\s*)?(\d+)/(\d+)\s*$', text)
            if m:
                name, number, year = m.group(1) or '', m.group(2), m.group(3)
            else:
                name = text
                m = re.search(r'/animal/(\d+)-(\d+)', link['href'])
                number, year = (m.group(1), m.group(2)) if m else ('0', '')
            fields = {}
            for item in card.select('.fusion-li-item-content'):
                label = item.find('strong')
                if label is not None:
                    key = label.get_text(strip=True).rstrip(':')
                    label.extract()
                    fields[key] = item.get_text(' ', strip=True)
            img = card.find('img')
            imgurl = (img.get('data-src') or img.get('src')) if img is not None else None
            if imgurl and imgurl.startswith('data:'):
                imgurl = img.get('data-orig-src') or img.get('data-lazy-src')
            animals.append({
                'link': link['href'],
                'name': name,
                'number': number,
                'year': year,
                'breed': fields.get('Rasa', ''),
                'age': fields.get('Wiek', ''),
                'image': imgurl,
                'urgent': 'pilnej adopcji' in card.get_text(' ', strip=True).lower(),
                'sex': '',
            })
        return animals

    def get_list(self, url, referer=None, max_pages=20):
        """Pobiera liste zwierzat; idzie za linkiem do nastepnej strony, jesli taki jest."""
        animals = []
        for _ in range(max_pages):
            soup = self.getsoup(url, referer=referer)
            animals += self.parse_cards(soup)
            nxt = soup.select_one('.pagination a.pagination-next[href], .pagination a[rel=next][href]')
            if nxt is None:
                break
            referer, url = url, urljoin(url, nxt['href'])
        return animals

    def get_animals(self, species):
        """Zwraca liste zwierzat danego gatunku ('koty' lub 'psy')."""
        base = LIST_URLS[species]
        animals = self.get_list(base, referer=BASE_URL)
        if not animals:
            raise ValueError('Na stronie ' + base + ' nie znaleziono żadnych zwierząt. '
                             'Strona mogła zablokować pobieranie albo zmienić wygląd.')
        return list(dict((a['link'], a) for a in animals).values())

    # --- strona zwierzecia ---

    @classmethod
    def parse_details(cls, soup):
        texts = [d.get_text(' ', strip=True) for d in soup.select('div.fusion-text')]
        fields = {}
        for i, t in enumerate(texts[:-1]):
            if t in cls.DETAIL_LABELS and t not in fields:
                fields[t] = texts[i + 1]
        desc = ''
        for h in soup.find_all(['h2', 'h3', 'h4']):
            if h.get_text(strip=True) == 'Informacje':
                body = h.find_next('div', class_='fusion-text')
                if body is not None:
                    desc = body.get_text('\n', strip=True)
                break
        for marker in cls.DESC_END:
            desc = desc.split(marker)[0]
        return {
            'sex': fields.get('Płeć', '').lower(),
            'date': fields.get('Data przyjęcia', ''),
            'location': fields.get('Lokalizacja (boks)', ''),
            'desc': desc.strip(),
            'photos': len(soup.select('a.fusion-lightbox[data-rel^="iLightbox[gallery"]')),
            'fields': fields,
        }

    def get_details(self, link, referer=None):
        """Dane ze strony zwierzecia lub None. Zwraca (dane, czy_pobrano)."""
        entry = self.details.get(link)
        fresh = entry is not None and time.time() - entry['time'] < self.details_max_age
        if fresh or self.blocked:
            return entry, False
        try:
            data = self.parse_details(self.getsoup(link, referer=referer))
        except BlockedError as e:
            print('Strona blokuje zapytania (' + str(e) + '), dalej używam tylko pamięci podręcznej')
            self.blocked = True
            return entry, False
        except Exception as e:
            print('Problem z ' + link + ': ' + str(e))
            return entry, False
        data['time'] = time.time()
        self.details[link] = data
        return data, True

    # --- zdjecia ---

    def get_thumbnail(self, url, referer=None, fixed_height=150):
        """Miniatura zdjecia; pobierana tylko raz, potem brana z pamieci podrecznej."""
        data = self.images.get(url) if url else None
        if data is None and url and not self.blocked:
            try:
                image = Image.open(BytesIO(self._get(url, referer=referer).content))
                height_percent = (fixed_height / float(image.size[0]))
                width_size = int((float(image.size[1]) * float(height_percent)))
                image = image.convert('RGB').resize((fixed_height, width_size), Image.NEAREST)
                buf = BytesIO()
                image.save(buf, format='PNG')
                data = buf.getvalue()
                self.images[url] = data
            except BlockedError as e:
                print('Strona blokuje zapytania (' + str(e) + '), dalej używam tylko pamięci podręcznej')
                self.blocked = True
            except Exception as e:
                print('Problem ze zdjęciem ' + str(url) + ': ' + str(e))
        if data is None:
            return Image.new('RGB', (fixed_height, fixed_height * 3 // 4), 'white')
        image = Image.open(BytesIO(data))
        image.load()
        return image
