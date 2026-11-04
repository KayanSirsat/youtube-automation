"""Download stock assets for a query and log each file's source, author and license.

Usage:
    python fetch_assets.py commons "Gulfstream G650" -n 10    # real, named subjects (photos)
    python fetch_assets.py unsplash "private jet interior"     # generic, polished photos
    python fetch_assets.py pixabay "yacht ocean" -n 3          # b-roll video clips (1080p)

Files go to assets/<query>/ (or -o DIR) with one sources.csv row per file. Whenever the
`attribution` column is not empty, it must be credited in the video description.

Environment: UNSPLASH_ACCESS_KEY, PIXABAY_API_KEY, WIKIMEDIA_CONTACT (an email or URL;
Wikimedia blocks scripts that don't name a contact: https://w.wiki/4wJS).
Pixabay is used for video only: its image server answers HTTP 429 to scripts.
"""
import argparse
import csv
import os
import re
import sys
from itertools import islice
from pathlib import Path
from urllib.parse import urlsplit

import requests

PIXABAY_LICENSE = "Pixabay Content License"
MAX_BYTES = 300_000_000  # Commons hosts multi-GB files; nothing a 1080p video needs is that large
# Commons licenses safe for a monetized video: public domain, CC0, CC BY, CC BY-SA. Never NC, ND or GFDL-only.
FREE_LICENSE = re.compile(r"(public domain|pd\b|cc0|cc[- ]by(-sa)?[- ]\d).*", re.I)
assert all(FREE_LICENSE.fullmatch(s) for s in ["CC BY-SA 4.0", "CC BY 2.0", "Public domain", "CC0", "CC BY-SA 2.5,2.0,1.0"])
assert not any(FREE_LICENSE.fullmatch(s) for s in ["CC BY-NC-SA 4.0", "CC BY-ND 2.0", "GFDL", "Attribution", ""])


def env(name):
    return os.environ.get(name) or sys.exit(f"{name} is not set")


def get(url, **kw):
    r = requests.get(url, timeout=60, **kw)
    if not r.ok:  # no raise_for_status: its message prints the full URL, which can carry an API key
        raise RuntimeError(f"HTTP {r.status_code} from {url.split('?')[0]}: {r.text[:200]}")
    return r


def unsplash(query, n):
    h = {"Authorization": f"Client-ID {env('UNSPLASH_ACCESS_KEY')}", "Accept-Version": "v1"}
    results = get("https://api.unsplash.com/search/photos", headers=h, params={
        "query": query, "per_page": min(n, 30), "orientation": "landscape"}).json()["results"]
    for p in results:
        name = p["user"]["name"]
        yield dict(id=p["id"], url=p["urls"]["raw"] + "&w=2560&fm=jpg&q=85", ext=".jpg",
                   page=p["links"]["html"], author=name, license="Unsplash License",
                   attribution=f"Photo by {name} on Unsplash",
                   # Unsplash API terms: report each real download to this endpoint
                   ping=lambda u=p["links"]["download_location"]: get(u, headers=h))


def commons(query, n):
    ua = {"User-Agent": f"yt-asset-fetch/0.1 ({env('WIKIMEDIA_CONTACT')}) python-requests/{requests.__version__}"}
    pages = get("https://commons.wikimedia.org/w/api.php", headers=ua, params={
        "action": "query", "format": "json", "generator": "search", "gsrsearch": f"{query} filetype:bitmap",
        "gsrnamespace": 6, "gsrlimit": min(n * 3, 50),  # over-fetch: the license filter drops some
        "prop": "imageinfo", "iiprop": "url|extmetadata", "iiurlwidth": 2560,
    }).json().get("query", {}).get("pages", {})
    for p in sorted(pages.values(), key=lambda p: p["index"]):  # keep search relevance order
        ii = p["imageinfo"][0]
        meta = ii["extmetadata"]
        lic = meta.get("LicenseShortName", {}).get("value", "")
        if not FREE_LICENSE.fullmatch(lic):
            continue
        author = re.sub(r"<[^>]+>", "", meta.get("Artist", {}).get("value", "unknown")).strip()[:100]
        url = ii.get("thumburl", ii["url"])
        no_credit = re.match(r"(public domain|pd\b|cc0)", lic, re.I)
        yield dict(id=p["pageid"], url=url, ext=Path(urlsplit(url).path).suffix.lower(), headers=ua,
                   page=ii["descriptionurl"], author=author, license=lic,
                   attribution="" if no_credit else f"{p['title'][5:]} by {author}, {lic}, {ii['descriptionurl']}")


def pixabay(query, n):
    hits = get("https://pixabay.com/api/videos/", params={
        "key": env("PIXABAY_API_KEY"), "q": query[:100], "safesearch": "true",
        "per_page": max(3, min(n, 200)),  # API accepts 3..200
    }).json()["hits"]
    for h in hits:
        yield dict(id=h["id"], url=h["videos"]["medium"]["url"], ext=".mp4",  # medium = 1920x1080
                   page=h["pageURL"], author=h["user"], license=PIXABAY_LICENSE, attribution="")


SOURCES = {"commons": commons, "unsplash": unsplash, "pixabay": pixabay}


def download(item, path):
    with requests.get(item["url"], headers=item.get("headers"), stream=True, timeout=60) as r:
        if not r.ok:
            raise RuntimeError(f"HTTP {r.status_code}")
        if int(r.headers.get("content-length", 0)) > MAX_BYTES:
            raise RuntimeError("file too large")
        part = path.with_name(path.name + ".part")  # a crash mid-download must not leave a "done" file
        with part.open("wb") as f:
            for chunk in r.iter_content(1 << 20):
                f.write(chunk)
        part.rename(path)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("source", choices=SOURCES)
    ap.add_argument("query")
    ap.add_argument("-n", type=int, default=5, help="number of files")
    ap.add_argument("-o", "--out", type=Path)
    args = ap.parse_args()

    out = args.out or Path("assets") / re.sub(r"[^a-z0-9]+", "-", args.query.lower()).strip("-")
    out.mkdir(parents=True, exist_ok=True)
    log = out / "sources.csv"
    new_log = not log.exists()
    got = 0
    with log.open("a", newline="") as f:
        w = csv.writer(f)
        if new_log:
            w.writerow(["file", "source", "query", "source_url", "author", "license", "attribution"])
        try:
            for item in islice(SOURCES[args.source](args.query, args.n), args.n):
                path = out / f"{args.source}_{item['id']}{item['ext']}"
                if path.exists():
                    continue
                try:
                    download(item, path)
                    if "ping" in item:
                        item["ping"]()
                except (RuntimeError, requests.RequestException) as e:
                    print(f"skip {item['page']}: {e}", file=sys.stderr)
                    continue
                w.writerow([path.name, args.source, args.query, item["page"], item["author"],
                            item["license"], item["attribution"]])
                print(path)
                got += 1
        except RuntimeError as e:  # search request failed
            sys.exit(f"{args.source} search failed: {e}")
    if not got:
        print(f"No new files for {args.query!r}", file=sys.stderr)


if __name__ == "__main__":
    main()
