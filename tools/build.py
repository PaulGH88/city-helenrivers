"""Build the City of Delafield homepage with the HELEN banner (Helen Rivers).

python tools/build.py   -> fetch the live homepage fresh, mirror its styles, scripts and images,
                           put the banner on the hero photo, write docs/ (GitHub Pages: main /docs)

Cloned from the Town build (town-helenmaple-demo). Every navigation link and form points at the
real www.cityofdelafieldwi.gov. Styles, scripts and images the page needs are copied under docs/
so the copy looks right and does not hot-link the City's server. CivicPlus serves its images as
/ImageRepository/Document?documentID=N; those become ImageRepository/doc-N.<ext>. Nothing visible
is added except the banner. The City's paid widgets (DocAccess, Google Translate) are left off.
"""
import os, re, sys, shutil, urllib.request, urllib.parse, urllib.error, hashlib, time

CITY = 'https://www.cityofdelafieldwi.gov/'
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
DIST = os.path.join(REPO, 'docs')
SNAP = os.path.join(REPO, 'snapshot')
UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/130 Safari/537.36'}

def get(url, tries=3):
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers=UA)
            with urllib.request.urlopen(req, timeout=60) as r:
                return r.read(), r.headers.get('Content-Type', '')
        except urllib.error.HTTPError as e:
            if e.code in (403, 404) or i == tries - 1:
                raise
        except Exception:
            if i == tries - 1:
                raise
        time.sleep(1.5)

def is_city(u):
    return urllib.parse.urlparse(u).netloc in ('www.cityofdelafieldwi.gov', 'cityofdelafieldwi.gov')

EXT = {'image/jpeg': '.jpg', 'image/png': '.png', 'image/gif': '.gif', 'image/svg+xml': '.svg',
       'image/webp': '.webp', 'application/pdf': '.pdf', 'image/x-icon': '.ico'}

def sniff(data, ctype):
    c = (ctype or '').split(';')[0].strip().lower()
    if c in EXT:
        return EXT[c]
    if data[:3] == b'\xff\xd8\xff': return '.jpg'
    if data[:8] == b'\x89PNG\r\n\x1a\n': return '.png'
    if data[:4] == b'GIF8': return '.gif'
    if b'<svg' in data[:400]: return '.svg'
    return '.bin'

mirrored = {}
def local_name(abs_url):
    """Where an asset lives under docs/ (no leading slash), or None for a page."""
    p = urllib.parse.urlparse(abs_url)
    path = urllib.parse.unquote(p.path).lstrip('/')
    m = re.match(r'(?i)ImageRepository/Document$', path)
    if m:
        q = urllib.parse.parse_qs(p.query.lower())
        doc = (q.get('documentid') or [''])[0]
        return ('ImageRepository/doc-' + doc) if doc.isdigit() else None
    if not path or path.endswith('/'):
        return None
    return path

def city_search(text):
    """The City's search builds a site-root address in script; send it to the real City search."""
    for q in ('"', "'"):
        text = text.replace(q + '/Search?', q + CITY + 'Search?')
    return text

def mirror(abs_url, depth=0):
    """Copy a City asset into docs/; return its site-root-relative path (URL-quoted) or None."""
    key = local_name(abs_url)
    if not key:
        return None
    if key in mirrored:
        return mirrored[key]
    cache = os.path.join(SNAP, 'assets-cache', key.replace('/', os.sep))
    data = ctype = None
    for c in [cache] + [cache + e for e in set(EXT.values()) | {'.bin'}]:
        if os.path.isfile(c) and not os.environ.get('FRESH'):
            data = open(c, 'rb').read(); ctype = ''
            break
    if data is None:
        try:
            data, ctype = get(abs_url.split('#')[0])
        except Exception as e:
            print('  MISS', abs_url, e)
            mirrored[key] = None
            return None
        os.makedirs(os.path.dirname(cache), exist_ok=True)
        with open(cache, 'wb') as f:
            f.write(data)
    rel = key
    if key.startswith('ImageRepository/doc-'):
        rel = key + sniff(data, ctype)
    if rel.lower().endswith('.css') and depth < 3:
        css = data.decode('utf-8', 'replace')
        def fix(m):
            q, ref = m.group(1), m.group(2).strip()
            if ref.startswith(('data:', '#')):
                return m.group(0)
            full = urllib.parse.urljoin(abs_url, ref)
            if not is_city(full):
                return m.group(0)
            loc = mirror(full, depth + 1)
            if not loc:
                return 'url(%s%s%s)' % (q, full, q)
            back = os.path.relpath(urllib.parse.unquote(loc), os.path.dirname(rel) or '.').replace(os.sep, '/')
            return 'url(%s%s%s)' % (q, urllib.parse.quote(back), q)
        css = re.sub(r'url\(\s*([\'"]?)([^\'")]+)\1\s*\)', fix, css)
        data = css.encode('utf-8')
    if rel.lower().endswith('.js'):
        data = city_search(data.decode('utf-8', 'replace')).encode('utf-8')
    dest = os.path.join(DIST, rel.replace('/', os.sep))
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    with open(dest, 'wb') as f:
        f.write(data)
    out = urllib.parse.quote(rel)
    mirrored[key] = out
    return out

SKIP = ('#', 'mailto:', 'tel:', 'javascript:', 'data:', 'sms:')
ASSET_EXT = re.compile(r'\.(css|js|png|jpe?g|gif|svg|webp|ico|woff2?|ttf|eot)$', re.I)

def rewrite(html):
    """Steps 3-5: images mirrored, links and forms to the real City site."""
    # 3. CivicPlus image documents, wherever they appear (attributes, inline CSS, inline JS).
    def fix_doc(m):
        full = urllib.parse.urljoin(CITY, m.group(0).replace('&amp;', '&'))
        loc = mirror(full)
        return loc if loc else full
    html = re.sub(r'(?i)(?:https?://www\.cityofdelafieldwi\.gov)?/ImageRepository/Document\?documentI[Dd]=\d+', fix_doc, html)

    # 4. Attributes: assets mirrored, links and forms to the real City site.
    def fix_attr(m):
        tag, attr, q, val = m.group(1).lower(), m.group(2).lower(), m.group(3), m.group(4)
        v = val.strip()
        if not v or v.startswith(SKIP) or v.startswith('ImageRepository/'):
            return m.group(0)
        full = urllib.parse.urljoin(CITY, v.replace('&amp;', '&'))
        is_asset = attr == 'src' or (tag == 'link' and attr == 'href')
        if is_asset and is_city(full):
            loc = mirror(full)
            if loc:
                return m.group(0).replace(q + val + q, q + loc + q)
        if not v.startswith(('http:', 'https:', '//')):
            return m.group(0).replace(q + val + q, q + full.replace('&', '&amp;') + q)
        return m.group(0)
    html = re.sub(r'<(\w+)\b[^>]*?\s(href|src|action)=(["\'])(.*?)\3', fix_attr, html, flags=re.S | re.I)

    # 5. Inline style url()s for theme images.
    def fix_style_url(m):
        q, ref = m.group(1), m.group(2)
        if ref.startswith(('data:', '#', 'ImageRepository/')):
            return m.group(0)
        full = urllib.parse.urljoin(CITY, ref)
        if is_city(full):
            loc = mirror(full)
            if loc:
                return 'url(%s%s%s)' % (q, loc, q)
        return m.group(0)
    html = re.sub(r'url\(\s*([\'"]?)([^\'")]+)\1\s*\)', fix_style_url, html)

    return html

def main():
    os.makedirs(DIST, exist_ok=True)
    for name in os.listdir(DIST):
        p = os.path.join(DIST, name)
        shutil.rmtree(p) if os.path.isdir(p) else os.remove(p)
    raw, _ = get(CITY)
    html = raw.decode('utf-8', 'replace')
    os.makedirs(SNAP, exist_ok=True)
    with open(os.path.join(SNAP, 'city-homepage-raw.html'), 'w', encoding='utf-8') as f:
        f.write(html)
    print('fetched', len(html), 'chars, sha', hashlib.sha256(html.encode()).hexdigest()[:12])

    # 1. No <base>; the City's paid widgets and analytics stay off.
    html = re.sub(r'<base [^>]*>\s*', '', html, flags=re.I)
    def drop_script(m):
        s = m.group(0)
        for bad in ('docaccess.com', 'translate.google.com', 'googleTranslateElementInit',
                    'googletagmanager', 'google-analytics', 'gtag(', 'SplashModal',
                    'instrumentationKey', 'appInsights', 'monitor.azure.com'):
            if bad in s:
                return ''
        return s
    html = re.sub(r'<script\b[^>]*>.*?</script>', drop_script, html, flags=re.S | re.I)
    html = re.sub(r'<div id="google_translate_element"[^>]*>\s*</div>', '', html, flags=re.I)
    # The City's own account keys are not ours to reuse (Google Maps key, unused on the homepage).
    html, n = re.subn(r'<input type="hidden" id="GoogleMapsKey"[^>]*>', '', html)
    if n != 1:
        sys.exit('FAIL: GoogleMapsKey input found %d times' % n)
    for leak in ('instrumentationKey', 'AIzaSy', 'monitor.azure.com'):
        if leak in html:
            sys.exit('FAIL: City key or tracker still on the page: ' + leak)

    # The copy does not call itself the City's official website (Paul YES, Gloria item 8, 2026-11-06).
    html, n = re.subn(r'<title>[^<]*Official Website[^<]*</title>', '<title>City of Delafield - Wisconsin</title>', html, count=1)
    if n != 1:
        sys.exit('FAIL: page title not found exactly once (%d)' % n)

    # 2. Search engines stay out.
    if re.search(r'<meta name="robots"', html, re.I):
        html = re.sub(r'<meta name="robots"[^>]*>', '<meta name="robots" content="noindex,nofollow">', html, flags=re.I)
    else:
        html = html.replace('<head>', '<head>\n<meta name="robots" content="noindex,nofollow">', 1)

    html = city_search(rewrite(html))

    # 6. Theme folders named inside inline scripts (banner arrows): make them page-relative.
    html = html.replace("'/App_Themes/", "'App_Themes/")
    # Scripts the City's code loads by name as it runs.
    for extra in ['Assets/Mystique/Shared/Scripts/webfontloader/webfont.js',
                  'Assets/Scripts/RWD/MediaFramework.js', 'Assets/Scripts/attrchange.js',
                  'Assets/Scripts/attrchange_ext.js', 'Assets/Scripts/RWD/quo.debug.js']:
        if not mirror(CITY + extra):
            sys.exit('FAIL: runtime script missing: ' + extra)
    # What the City's scripts fetch at run time (tools/capture-endpoints.mjs), with the same
    # link and image rewriting as the page, served by banner/cp-shim.js.
    cpd = os.path.join(DIST, 'cp')
    os.makedirs(cpd, exist_ok=True)
    for name in ['toggle.json', 'antiforgery.json', 'submenus.html', 'content-load.html']:
        src = os.path.join(SNAP, 'cp', name)
        if not os.path.exists(src):
            sys.exit('FAIL: snapshot/cp/%s missing - run node tools/capture-endpoints.mjs first' % name)
        txt = open(src, encoding='utf-8').read()
        if name.endswith('.html'):
            txt = rewrite(txt)
        if name == 'antiforgery.json':
            txt = '{"token":"static-copy"}'  # never ship a token the City issued
        with open(os.path.join(cpd, name), 'w', encoding='utf-8', newline='\n') as f:
            f.write(txt)
    html = re.sub(r'(<head[^>]*>)', r'\1\n<script src="helen/cp-shim.js"></script>', html, count=1, flags=re.I)

    # 7. The HELEN banner sits on the hero photo, under the menu, above the round quick links.
    banner = open(os.path.join(REPO, 'banner', 'banner.html'), encoding='utf-8').read()
    anchor = '<div data-cpRole="structuralContainer" id="bannerSearchTS">'
    if html.count(anchor) != 1:
        sys.exit('FAIL: hero anchor found %d times' % html.count(anchor))
    # Just after the Pay Bills / social row closes, so on phones the banner follows it.
    i = html.index(anchor) + len(anchor); depth = 1
    for m in re.finditer(r'<(/?)div\b', html[i:], flags=re.I):
        depth += -1 if m.group(1) else 1
        if depth == 0:
            end = i + html[i:].index('>', m.start()) + 1
            break
    else:
        sys.exit('FAIL: end of the Pay Bills row not found')
    html = html[:end] + '\n' + banner + '\n' + html[end:]
    html = html.replace('</head>', '<link rel="stylesheet" href="helen/banner.css">\n</head>', 1)
    html = html.replace('</body>', '<script src="helen/banner.js"></script>\n</body>', 1)

    hd = os.path.join(DIST, 'helen')
    os.makedirs(hd, exist_ok=True)
    for f in os.listdir(os.path.join(REPO, 'banner')):
        if f != 'banner.html':
            shutil.copy2(os.path.join(REPO, 'banner', f), os.path.join(hd, f))

    with open(os.path.join(DIST, 'index.html'), 'w', encoding='utf-8', newline='\n') as f:
        f.write(html)
    with open(os.path.join(DIST, 'robots.txt'), 'w', newline='\n') as f:
        f.write('User-agent: *\nDisallow: /\n')
    open(os.path.join(DIST, '.nojekyll'), 'w').close()
    with open(os.path.join(DIST, 'CNAME'), 'w', newline='\n') as f:
        f.write('city.helenrivers.online\n')
    print('mirrored', sum(1 for v in mirrored.values() if v), 'files; missed', sum(1 for v in mirrored.values() if not v))

if __name__ == '__main__':
    main()
