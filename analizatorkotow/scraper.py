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

    Wszystkie potrzebne dane (imie, numer, wiek, plec, zdjecie, "pilne") sa na
    stronach z lista zwierzat, wiec nie pobieramy stron poszczegolnych zwierzat:
    jedno uruchomienie to 3 strony (lista + lista samic + lista samcow) oraz
    miniatury tylko tych zwierzat, ktorych zdjec nie ma jeszcze w pamieci podrecznej.

    - jedna sesja HTTP (keep-alive), naglowki przegladarki i Referer,
    - odstepy z losowym jitterem i ponawianie z backoffem
      (z respektowaniem Retry-After) przy 429/5xx,
    - po 403/429 dalsze zdjecia nie sa pobierane (wstawiany jest pusty obrazek).
    """

    def __init__(self, cache_path, delay=(1.0, 2.5)):
        self.cache_path = str(cache_path)
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
        self.cache = self._load_cache()  # url zdjecia -> miniatura w PNG

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
                pickle.dump(self.cache, f)
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
        """Zwraca liste zwierzat danego gatunku ('koty' lub 'psy') z uzupelniona plcia."""
        base = LIST_URLS[species]
        animals = self.get_list(base, referer=BASE_URL)
        if not animals:
            raise ValueError('Na stronie ' + base + ' nie znaleziono żadnych zwierząt. '
                             'Strona mogła zablokować pobieranie albo zmienić wygląd.')
        bylink = dict((a['link'], a) for a in animals)
        animals = list(bylink.values())
        for sex in ('samica', 'samiec'):
            try:
                for a in self.get_list(base + 'plec-' + sex + '/', referer=base):
                    if a['link'] in bylink:
                        bylink[a['link']]['sex'] = sex
            except BlockedError:
                raise
            except Exception as e:
                print('Nie udało się pobrać listy ' + sex + ': ' + str(e))
        return animals

    # --- zdjecia ---

    def get_thumbnail(self, url, referer=None, fixed_height=150):
        """Miniatura zdjecia; pobierana tylko raz, potem brana z pamieci podrecznej."""
        data = self.cache.get(url) if url else None
        if data is None and url and not self.blocked:
            try:
                image = Image.open(BytesIO(self._get(url, referer=referer).content))
                height_percent = (fixed_height / float(image.size[0]))
                width_size = int((float(image.size[1]) * float(height_percent)))
                image = image.convert('RGB').resize((fixed_height, width_size), Image.NEAREST)
                buf = BytesIO()
                image.save(buf, format='PNG')
                data = buf.getvalue()
                self.cache[url] = data
            except BlockedError as e:
                print('Strona blokuje zapytania (' + str(e) + '), pomijam pozostałe zdjęcia')
                self.blocked = True
            except Exception as e:
                print('Problem ze zdjęciem ' + str(url) + ': ' + str(e))
        if data is None:
            return Image.new('RGB', (fixed_height, fixed_height * 3 // 4), 'white')
        image = Image.open(BytesIO(data))
        image.load()
        return image
