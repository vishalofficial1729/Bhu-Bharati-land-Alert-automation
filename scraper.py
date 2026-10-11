"""Fetch one survey record from Bhu Bharati -> Land Details Search -> Survey No."""
import base64
import difflib
import json
import queue
import threading
import time
import urllib.parse
import urllib.request

import ddddocr
from playwright.sync_api import sync_playwright

URL = "https://bhubharati.telangana.gov.in/knowLandStatus"
CAPTCHA_LEN = 5
OCR_TRIES = 5


class FetchError(Exception):
    pass


def _twocaptcha(png: bytes, key: str) -> str:
    """Paid fallback, used only when OCR keeps failing and a 2Captcha key is configured."""
    body = urllib.parse.urlencode({"key": key, "method": "base64", "json": 1,
                                   "body": base64.b64encode(png).decode()}).encode()
    r = json.load(urllib.request.urlopen("https://2captcha.com/in.php", body, timeout=30))
    if r.get("status") != 1:
        raise FetchError(f"2Captcha upload failed: {r}")
    for _ in range(24):
        time.sleep(5)
        q = urllib.parse.urlencode({"key": key, "action": "get", "id": r["request"], "json": 1})
        res = json.load(urllib.request.urlopen(f"https://2captcha.com/res.php?{q}", timeout=30))
        if res.get("status") == 1:
            return res["request"].strip().upper()
        if res.get("request") != "CAPCHA_NOT_READY":
            raise FetchError(f"2Captcha failed: {res}")
    raise FetchError("2Captcha timed out")


class Scraper:
    """Usage: with Scraper() as s: s.fetch("Rangareddy", "Moinabad", "Chilkoor", "122")"""

    def __init__(self, twocaptcha_key: str = "", headless: bool = True):
        self.twocaptcha_key = twocaptcha_key
        self.headless = headless
        self.ocr = ddddocr.DdddOcr(show_ad=False)

    def __enter__(self):
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(headless=self.headless)
        return self

    def __exit__(self, *exc):
        self._browser.close()
        self._pw.stop()

    def fetch(self, district: str, mandal: str, village: str, survey: str) -> dict:
        (_, res), = self.fetch_village(district, mandal, village, [survey])
        if isinstance(res, Exception):
            raise res
        return res

    def fetch_village(self, district, mandal, village, surveys):
        """Yield (survey, data | FetchError). District/mandal/village are picked once and the
        page is reused for every survey in the village; after any error the page is reopened."""
        ctx = pg = None
        try:
            for survey in surveys:
                try:
                    if pg is None:
                        ctx, pg, alerts = self._open(district, mandal, village)
                    yield survey, self._fetch_survey(pg, alerts, survey)
                except Exception as ex:
                    yield survey, ex if isinstance(ex, FetchError) else FetchError(f"{type(ex).__name__}: {ex}")
                    if ctx:
                        ctx.close()
                    ctx = pg = None
        finally:
            if ctx:
                ctx.close()

    def _open(self, district, mandal, village):
        ctx = self._browser.new_context()
        # Site JS crashes in cookie() when document.cookie is empty (its only cookie is HttpOnly).
        ctx.add_cookies([{"name": "bb", "value": "1", "domain": "bhubharati.telangana.gov.in", "path": "/"}])
        pg = ctx.new_page()
        alerts = []
        pg.on("dialog", lambda d: (alerts.append(d.message), d.accept()))
        try:
            pg.goto(URL, wait_until="networkidle", timeout=60000)
            self._pick(pg, "districtID", "District", district)
            self._pick(pg, "mandalID", "Mandal", mandal)
            self._pick(pg, "villageId", "Village", village)
        except Exception:
            ctx.close()
            raise
        return ctx, pg, alerts

    def _fetch_survey(self, pg, alerts, survey) -> dict:
        # Clear the previous survey's khata list and results so we wait for fresh ones.
        pg.evaluate("document.getElementById('khataNoIdselect').length = 1;"
                    "document.getElementById('searchDataGrid').innerHTML = ''")
        self._pick(pg, "surveyIdselect", "Survey No", survey, by_value=True)
        khatas = self._options(pg, "khataNoIdselect", "Khata (none found for this survey no)")
        if len(khatas) > 1:
            raise FetchError(f"Multiple khatas for this survey no: {', '.join(t for _, t in khatas)}")
        pg.select_option("#khataNoIdselect", khatas[0][0])
        data = self._solve_and_read(pg, alerts)
        pg.wait_for_timeout(1200)  # site swaps the captcha 1s after results; let that happen first
        return data

    @staticmethod
    def _options(pg, sel, what):
        try:
            pg.wait_for_function(f"document.getElementById('{sel}').options.length > 1", timeout=30000)
        except Exception:
            raise FetchError(f"{what}: dropdown did not load")
        opts = pg.eval_on_selector(f"#{sel}", "s => [...s.options].map(o => [o.value, o.text.trim()])")
        return [(v, t) for v, t in opts if v not in ("0", "")]

    def _pick(self, pg, sel, what, wanted, by_value=False):
        opts = self._options(pg, sel, what)
        wanted = str(wanted).strip()
        # Options look like "Rangareddy|రంగా రెడ్డి"; survey options are exact values like "22/ఇ".
        keys = {(v if by_value else t.split("|")[0]).strip().lower().replace(" ", ""): v for v, t in opts}
        val = keys.get(wanted.lower().replace(" ", ""))
        if val is None:
            close = difflib.get_close_matches(wanted.lower(), list(keys), n=3)
            raise FetchError(f"{what} '{wanted}' not found" + (f" (did you mean: {', '.join(close)}?)" if close else ""))
        pg.select_option(f"#{sel}", val)

    def _captcha_png(self, pg) -> bytes:
        pg.evaluate("changecaptcha()")
        pg.wait_for_function("(i => i.complete && i.naturalWidth > 0)(document.getElementById('imgcapcha'))")
        # Read the image at native size; the element screenshot is downscaled and OCR misreads it.
        return base64.b64decode(pg.evaluate("""() => { const i = document.getElementById('imgcapcha');
            const c = document.createElement('canvas'); c.width = i.naturalWidth; c.height = i.naturalHeight;
            c.getContext('2d').drawImage(i, 0, 0); return c.toDataURL('image/png').split(',')[1]; }"""))

    def _solve_and_read(self, pg, alerts) -> dict:
        tries = OCR_TRIES + (1 if self.twocaptcha_key else 0)
        for attempt in range(tries):
            png = self._captcha_png(pg)
            if attempt < OCR_TRIES:
                text = "".join(c for c in self.ocr.classification(png).upper() if c.isalnum())
                if len(text) != CAPTCHA_LEN:
                    continue  # obvious misread, get a fresh image without submitting
            else:
                text = _twocaptcha(png, self.twocaptcha_key)
            n = len(alerts)
            pg.fill("#captchavalue", text)
            pg.click("input[value='Fetch']")
            for _ in range(120):  # up to 60s for captcha check + results
                pg.wait_for_timeout(500)
                if len(alerts) > n:
                    break
                if pg.locator("#searchDataGrid strong").count():
                    pg.wait_for_timeout(500)
                    return self._parse(pg)
            if len(alerts) > n and "captcha" in alerts[-1].lower():
                continue
            raise FetchError(f"Site said: {alerts[-1]}" if len(alerts) > n else "Timed out waiting for results")
        raise FetchError(f"Captcha not solved after {tries} tries")

    @staticmethod
    def _parse(pg) -> dict:
        pairs = pg.evaluate("""() => [...document.querySelectorAll('#searchDataGrid strong')].map(s =>
            [s.textContent.trim(), s.parentElement.textContent.replace(s.textContent, '').replace(/\\s+/g, ' ').trim()])""")
        data = {}
        for label, value in pairs:
            data[label] = f"{data[label]} | {value}" if label in data else value
        return data


def fetch_all(items, workers=3, twocaptcha_key="", headless=True, delay=1.0, chunk=25, stop=None):
    """items: [(district, mandal, village, survey)]. Yields ((d, m, v, s), data | FetchError)
    as results arrive, using `workers` browsers in parallel, one village chunk at a time.
    Setting the `stop` threading.Event makes every browser finish its current survey and quit."""
    stop = stop or threading.Event()
    groups = {}
    for d, m, v, s in items:
        groups.setdefault((d, m, v), []).append(s)
    work = queue.Queue()
    for loc, surveys in groups.items():
        size = min(chunk, max(5, -(-len(surveys) // workers)))  # big villages are shared across browsers
        for i in range(0, len(surveys), size):
            work.put((loc, surveys[i:i + size]))
    out = queue.Queue()

    def worker():
        try:
            with Scraper(twocaptcha_key, headless) as sc:
                while not stop.is_set():
                    try:
                        (d, m, v), surveys = work.get_nowait()
                    except queue.Empty:
                        return
                    for s, res in sc.fetch_village(d, m, v, surveys):
                        out.put(((d, m, v, s), res))
                        if stop.is_set():
                            break
                        time.sleep(delay)  # be polite to the govt server
        except Exception as ex:  # browser failed to start; other workers pick up the rest
            out.put(("crash", ex))
        finally:
            out.put(None)

    n = max(1, min(workers, work.qsize()))
    for _ in range(n):
        threading.Thread(target=worker, daemon=True).start()
    done = 0
    while done < n:
        r = out.get()
        if r is None:
            done += 1
        elif r[0] != "crash":
            yield r


if __name__ == "__main__":
    import sys
    args = sys.argv[1:] or ["Rangareddy", "Moinabad", "Chilkoor", "122"]
    with Scraper(headless="--show" not in args) as s:
        print(json.dumps(s.fetch(*[a for a in args if a != "--show"][:4]), ensure_ascii=False, indent=2))
