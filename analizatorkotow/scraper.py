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

    - jedna sesja HTTP (keep-alive) zamiast nowego polaczenia na kazde zapytanie,
    - naglowki przegladarki, odstepy z losowym jitterem i ponawianie z backoffem
      (z respektowaniem Retry-After) przy 429/5xx,
    - dyskowa pamiec podreczna szczegolow zwierzat (razem z miniaturka zdjecia),
      wiec przy kolejnych uruchomieniach pobierane sa tylko nowe zwierzaki
      lub te, ktorych dane sa starsze niz cache_days.
    """

    def __init__(self, cache_path, cache_days=7, delay=(1.0, 2.5)):
        self.cache_path = str(cache_path)
        self.cache_max_age = cache_days * 24 * 3600
        self.delay = delay
        self._last_request = 0.0
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        retry = Retry(total=4, backoff_factor=2, status_forcelist=(429, 500, 502, 503, 504),
                      allowed_methods=('GET',), respect_retry_after_header=True,
                      raise_on_status=False)
        adapter = HTTPAdapter(max_retries=retry, pool_connections=1, pool_maxsize=1)
        self.session.mount('https://', adapter)
        self.session.mount('http://', adapter)
        self.cache = self._load_cache()

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

    def getsoup(self, url, referer=None, parser='html.parser'):
        return BeautifulSoup(self._get(url, referer).text, parser)

    # --- listy zwierzat ---

    @staticmethod
    def get_links_to_cats(soup, links):
        # sciaga linki z danej podstrony
        for link in soup.findAll('a', attrs={'href': re.compile("/pet/")}):
            links.append(link.get('href'))
        return links

    @staticmethod
    def get_max_page(soup):
        # podaje liczbe podstron do sciagania
        nums = [1]
        for link in soup.findAll('a', attrs={'href': re.compile("pet_page=")}):
            m = re.search(r'pet_page=(\d+)', link.get('href'))
            if m:
                nums.append(int(m.group(1)))
        return max(nums)

    def get_all_cat_links(self, linkMain, linkEnd):
        first = linkMain + '1' + linkEnd
        soup = self.getsoup(first, referer=BASE_URL, parser='lxml')
        links = self.get_links_to_cats(soup, [])  # strona 1 nie jest pobierana drugi raz
        for k in range(2, self.get_max_page(soup) + 1):
            soup = self.getsoup(linkMain + str(k) + linkEnd, referer=first)
            links = self.get_links_to_cats(soup, links)
        return list(dict.fromkeys(links))

    # --- szczegoly zwierzecia ---

    def get_cat_details(self, link, referer=None):
        """Zwraca (CatID, CatName, CatAge, CatIn, CatDesc, CatStat, CatImage, from_cache)."""
        entry = self.cache.get(link)
        if entry is not None and time.time() - entry['time'] < self.cache_max_age:
            return self._unpack(entry) + (True,)
        try:
            data = self._fetch_cat_details(link, referer)
        except BlockedError:
            raise
        except Exception as e:
            print('Problem z ' + link + ': ' + str(e))
            if entry is not None:  # lepiej stare dane niz brak zwierzaka
                return self._unpack(entry) + (True,)
            return (0, 0, 0, 0, 0, '', 0, False)
        self.cache[link] = data
        return self._unpack(data) + (False,)

    @staticmethod
    def _unpack(entry):
        image = entry['image']
        if image is not None:
            image = Image.open(BytesIO(image))
            image.load()
        else:
            image = 0
        return (entry['id'], entry['name'], entry['age'], entry['in'], entry['desc'], entry['stat'], image)

    def _fetch_cat_details(self, link, referer):
        url = urljoin(BASE_URL, link)
        soup = self.getsoup(url, referer=referer)
        strongs = soup.body.find_all('strong')
        CatStat = str(strongs[5])
        if any(s in CatStat for s in ("tymczaso", "fundacja")):
            # zwierzak pominiety - zapamietujemy tylko status, bez pobierania zdjecia
            return {'time': time.time(), 'id': 0, 'name': 0, 'age': 0, 'in': 0, 'desc': 0,
                    'stat': CatStat, 'image': None}

        CatID = int(str(soup.body.find_all('h2')[0]).split(' ')[-1].split('>')[1].split('/')[0])
        CatName = str(soup.body.find_all('h2')[0]).split(' ')[0].split('>')[1] + '\n' + \
                  str(strongs[2]).split('>')[1].split('<')[0]
        CatAge = str(str(strongs[1]).split('>')[1].split('</')[0])
        CatIn = (str(strongs[6]).split('>')[1].split('<')[0]) + '\n' + \
                str(strongs[7]).split('>')[1].split('<')[0]
        CatIn += '\n\nZdjęć:' + str(len(soup.body.find_all('div', class_='pet-detail-gallery-column-photo')))
        try:
            if 'facebook' not in str(strongs[8]):
                CatIn += '\n\n' + str(strongs[8]).replace('<strong>', '').replace('</strong>', '')
        except IndexError:
            print('Problem with' + link)
        CatDesc = str(soup.body.find_all('div', class_='pet-description')).split('>')[1].split('<')[0].split(
            'Zapraszamy do zapoznania się z ankietą')[0]
        if "pilne" in CatStat:
            CatIn += '\n\nPILNE'
        if 'Ochota na Kota' in CatStat:
            CatIn += '\n\nOchota na Kota'

        imgurl = str(soup.body.find_all('img', class_='pet-detail-main-image')).split('data-src=')[1].split(' ')[0] \
            .replace('"', '')
        image = Image.open(BytesIO(self._get(urljoin(url, imgurl), referer=url).content))
        fixed_height = 150
        height_percent = (fixed_height / float(image.size[0]))
        width_size = int((float(image.size[1]) * float(height_percent)))
        image = image.convert('RGB').resize((fixed_height, width_size), Image.NEAREST)
        buf = BytesIO()
        image.save(buf, format='PNG')

        return {'time': time.time(), 'id': CatID, 'name': CatName, 'age': CatAge, 'in': CatIn,
                'desc': CatDesc, 'stat': CatStat, 'image': buf.getvalue()}
