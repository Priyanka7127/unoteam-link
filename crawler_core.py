"""
crawler_core.py
Core crawling, link-graph, and report-generation logic for the
Internal Linking Audit tool. Kept separate from the Streamlit UI
so it can be unit-tested without a browser or network access.
"""
import re
import time
from collections import defaultdict, deque
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup

USER_AGENT = "InternalLinkAuditBot/1.0 (+local audit tool; run by site owner)"

# ---------------------------------------------------------------------------
# URL helpers
# ---------------------------------------------------------------------------

SKIP_SCHEMES = ("mailto:", "tel:", "javascript:", "whatsapp:", "sms:")
NON_HTML_EXT = (
    ".jpg", ".jpeg", ".png", ".gif", ".svg", ".webp", ".pdf", ".zip",
    ".mp4", ".mp3", ".css", ".js", ".ico", ".xml", ".json", ".woff", ".woff2",
)

def normalize_url(url: str) -> str:
    """Strip fragments and trailing-slash-inconsistency so the same page
    isn't counted twice under two slightly different URLs."""
    url, _frag = urldefrag(url)
    parsed = urlparse(url)
    path = parsed.path
    if path == "":
        path = "/"
    # drop a single trailing slash duplicate (but keep root "/")
    if path != "/" and path.endswith("/"):
        pass  # keep as-is; WordPress-style sites are consistently slash-terminated
    normalized = parsed._replace(path=path, query=parsed.query).geturl()
    return normalized


def is_same_domain(url: str, root_netloc: str) -> bool:
    try:
        netloc = urlparse(url).netloc.lower()
    except Exception:
        return False
    netloc = netloc[4:] if netloc.startswith("www.") else netloc
    root = root_netloc[4:] if root_netloc.startswith("www.") else root_netloc
    return netloc == root


def is_crawlable_url(url: str) -> bool:
    low = url.lower()
    if any(low.startswith(s) for s in SKIP_SCHEMES):
        return False
    if any(low.split("?")[0].endswith(ext) for ext in NON_HTML_EXT):
        return False
    return True


# ---------------------------------------------------------------------------
# Page classification
# ---------------------------------------------------------------------------

BLOG_PATTERNS = ("/blog/", "/blogs/", "/news/", "/articles/", "/insights/")
CASE_STUDY_PATTERNS = ("/case-stud", "/case_stud", "/success-stor", "/patient-stor", "/testimonial")
UTILITY_KEYWORDS = (
    "contact", "privacy", "terms", "cookie", "sitemap", "career",
    "gallery", "faq", "author/", "category/", "tag/", "page/",
    "login", "cart", "checkout", "search", "404", "thank-you",
)
# Pages that are clearly "about the business" rather than a specific
# service/treatment, and shouldn't be forced into the Treatments bucket.
ABOUT_KEYWORDS = ("about-us", "about", "our-team", "meet-", "dentist-in", "clinic-in")


def classify_url(url: str, home_netloc: str) -> str:
    """Return one of: 'home', 'blog', 'case_study', 'utility', 'about', 'content'."""
    path = urlparse(url).path.lower()
    if path in ("", "/"):
        return "home"
    if any(p in path for p in BLOG_PATTERNS):
        return "blog"
    if any(p in path for p in CASE_STUDY_PATTERNS):
        return "case_study"
    if any(k in path for k in UTILITY_KEYWORDS):
        return "utility"
    if any(k in path for k in ABOUT_KEYWORDS):
        return "about"
    return "content"  # candidate treatment/service page


# ---------------------------------------------------------------------------
# HTML parsing
# ---------------------------------------------------------------------------

def get_title(soup: BeautifulSoup) -> str:
    if soup.title and soup.title.string:
        return soup.title.string.strip()
    h1 = soup.find("h1")
    if h1:
        return h1.get_text(strip=True)
    return "(untitled)"


def get_text_sample(soup: BeautifulSoup, max_words: int = 150) -> str:
    for tag in soup(["script", "style", "nav", "header", "footer"]):
        tag.decompose()
    text = soup.get_text(" ", strip=True)
    words = text.split()
    return " ".join(words[:max_words])


def extract_links(soup: BeautifulSoup, base_url: str):
    """Return list of dicts: {target, anchor_text, in_boilerplate}.
    in_boilerplate=True for links inside <nav>/<header>/<footer>, which are
    excluded from 'contextual inbound link' counts, mirroring the manual
    audit methodology (nav/footer links repeat on every page and don't
    reflect real topical relevance)."""
    results = []
    boilerplate_tags = soup.find_all(["nav", "header", "footer"])
    boilerplate_ids = {id(t) for t in boilerplate_tags}

    def is_inside_boilerplate(tag):
        parent = tag
        while parent is not None:
            if id(parent) in boilerplate_ids:
                return True
            parent = parent.parent
        return False

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href or not is_crawlable_url(href):
            continue
        target = urljoin(base_url, href)
        target = normalize_url(target)
        anchor_text = a.get_text(" ", strip=True) or "(image or empty anchor)"
        results.append({
            "target": target,
            "anchor_text": anchor_text,
            "in_boilerplate": is_inside_boilerplate(a),
        })
    return results


# ---------------------------------------------------------------------------
# Crawl
# ---------------------------------------------------------------------------

def crawl_site(start_url: str, max_pages: int = 150, delay: float = 0.4,
                progress_cb=None, log_cb=None):
    """Breadth-first crawl of a single domain.

    Returns:
        pages: dict[url] -> {
            'title': str, 'category': str, 'text_sample': str,
            'outbound': list[{'target','anchor_text','in_boilerplate'}],
            'status': int or None, 'error': str or None,
        }
    """
    start_url = normalize_url(start_url)
    root_netloc = urlparse(start_url).netloc

    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})

    pages = {}
    queue = deque([start_url])
    seen = {start_url}

    while queue and len(pages) < max_pages:
        url = queue.popleft()
        if log_cb:
            log_cb(f"Fetching {url}")
        try:
            resp = session.get(url, timeout=15, allow_redirects=True)
            final_url = normalize_url(resp.url)
        except requests.RequestException as e:
            pages[url] = {
                "title": "(failed to load)", "category": classify_url(url, root_netloc),
                "text_sample": "", "outbound": [], "status": None, "error": str(e),
            }
            if progress_cb:
                progress_cb(len(pages), max_pages)
            continue

        content_type = resp.headers.get("Content-Type", "")
        if resp.status_code != 200 or "text/html" not in content_type:
            pages[url] = {
                "title": "(non-HTML or error)", "category": classify_url(url, root_netloc),
                "text_sample": "", "outbound": [], "status": resp.status_code, "error": None,
            }
            if progress_cb:
                progress_cb(len(pages), max_pages)
            continue

        soup = BeautifulSoup(resp.text, "html.parser")
        title = get_title(soup)
        # IMPORTANT: extract_links must run on the untouched soup, because
        # get_text_sample() below decomposes (deletes) nav/header/footer
        # tags in place to build a clean text sample. Extracting links
        # first, text second, avoids losing nav/footer links entirely.
        links = extract_links(soup, final_url)
        text_sample = get_text_sample(soup)

        pages[final_url] = {
            "title": title,
            "category": classify_url(final_url, root_netloc),
            "text_sample": text_sample,
            "outbound": links,
            "status": resp.status_code,
            "error": None,
        }

        for link in links:
            tgt = link["target"]
            if tgt in seen:
                continue
            if not is_same_domain(tgt, root_netloc):
                continue
            seen.add(tgt)
            if len(seen) <= max_pages * 3:  # cap queue growth
                queue.append(tgt)

        if progress_cb:
            progress_cb(len(pages), max_pages)
        time.sleep(delay)

    return pages


# ---------------------------------------------------------------------------
# Link graph
# ---------------------------------------------------------------------------

def build_inbound_index(pages: dict):
    """Return dict[target_url] -> list of {'source','anchor_text'} for
    CONTEXTUAL (non-boilerplate) links only."""
    inbound = defaultdict(list)
    for source, data in pages.items():
        for link in data["outbound"]:
            if link["in_boilerplate"]:
                continue
            if link["target"] == source:
                continue  # self-link
            if link["target"] in pages:  # only count links landing on a crawled page
                inbound[link["target"]].append({
                    "source": source,
                    "anchor_text": link["anchor_text"],
                })
    return inbound


# ---------------------------------------------------------------------------
# Similarity-based recommendations (no external NLP deps)
# ---------------------------------------------------------------------------

_STOPWORDS = set("""
a an the of for and or to in on with is are your our we you at from by
best top treatment treatments in kandivali hyderabad india clinic dental
dentist smile connect jaydev
""".split())

def _tokenize(text: str):
    words = re.findall(r"[a-z]+", text.lower())
    return set(w for w in words if w not in _STOPWORDS and len(w) > 2)


def similarity(a_tokens: set, b_tokens: set) -> float:
    if not a_tokens or not b_tokens:
        return 0.0
    inter = len(a_tokens & b_tokens)
    union = len(a_tokens | b_tokens)
    return inter / union if union else 0.0


MIN_RELEVANCE_SCORE = 0.10  # floor below which a candidate is not "relevant
                             # enough" to recommend — never padded past this
                             # just to hit a target link count.
SAME_CATEGORY_PENALTY = 0.5  # a transactional/service page doesn't need to
                              # be linked from another service page; a
                              # topically related blog post or case study is
                              # usually the more natural, honest link. This
                              # only nudges ranking order for content→content
                              # candidates — it never overrides genuine
                              # relevance (the min-score floor still applies
                              # to the real similarity score, not this key).


def recommend_link_sources(target_url, pages, top_n=5, exclude_urls=None,
                            min_score=MIN_RELEVANCE_SCORE):
    """Find up to top_n topically-relevant OTHER pages to recommend as new
    inbound-link sources for target_url.

    Only candidates whose real similarity score clears min_score are
    returned — if fewer than top_n pages are genuinely relevant, fewer are
    returned rather than padding the list with unrelated pages.

    Cross-type intelligence: when target_url is a transactional/service
    ('content') page, another 'content' page is still eligible but is
    ranked behind an equally-or-less-relevant blog/case-study candidate,
    since service pages should preferentially link out to relevant content
    (blogs, case studies) rather than to each other.
    """
    target = pages[target_url]
    target_category = target["category"]
    target_tokens = _tokenize(target["title"] + " " + target["text_sample"])
    exclude_urls = exclude_urls or set()

    scored = []
    for url, data in pages.items():
        if url == target_url or url in exclude_urls:
            continue
        if data["category"] in ("utility", "home", "about"):
            continue
        if data.get("status") != 200:
            continue
        tokens = _tokenize(data["title"] + " " + data["text_sample"])
        score = similarity(target_tokens, tokens)
        if score < min_score:
            continue
        same_category = target_category == "content" and data["category"] == "content"
        rank_key = score * SAME_CATEGORY_PENALTY if same_category else score
        scored.append((rank_key, score, url, data["title"], data["category"]))

    scored.sort(key=lambda t: t[0], reverse=True)
    return [(score, url, title, category) for _, score, url, title, category in scored[:top_n]]
