"""
Discovery Engine v0.2 - Prospecting Suite / Engine 1.

Discovers companies matching our ICP by searching public agency directories
through DuckDuckGo with a visible Chromium browser driven by Playwright.

Responsibility ends after discovery. This engine does not qualify, score,
enrich contacts, or use AI. It feeds the Company Intelligence Engine.

Outputs (written to the directory containing this file):
    master_company_list.xlsx  - full discovery metadata
    prospector_input.csv      - exactly two columns: Company Name, Website
    run_report.xlsx           - run statistics
    discovery.log             - detailed runtime log
    discovery_autosave.json   - checkpoint for resume support

Usage:
    python discovery_engine.py
    python discovery_engine.py --fresh
    python discovery_engine.py --max-queries 20

Requires: playwright (with chromium installed), beautifulsoup4, requests,
tldextract, openpyxl.
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import signal
import time
from collections import Counter
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from types import FrameType
from urllib.parse import parse_qs, quote, unquote, urljoin, urlparse

import requests
import tldextract
from bs4 import BeautifulSoup
from openpyxl import Workbook
from playwright.sync_api import Browser, BrowserContext, Page, Playwright
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

# ----------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent

LOG_FILE = BASE_DIR / "discovery.log"
AUTOSAVE_FILE = BASE_DIR / "discovery_autosave.json"
OUTPUT_MASTER = BASE_DIR / "master_company_list.xlsx"
OUTPUT_PROSPECTOR = BASE_DIR / "prospector_input.csv"
OUTPUT_RUN_REPORT = BASE_DIR / "run_report.xlsx"

AUTOSAVE_VERSION = 3

SEARCH_KEYWORDS = [
    "Performance Marketing Agency",
    "Digital Marketing Agency",
    "Paid Media Agency",
    "PPC Agency",
    "Growth Marketing Agency",
    "Media Buying Agency",
    "Meta Ads Agency",
    "Google Ads Agency",
    "Media Buying Agency",
    "Paid Social Agency",
]

INDIAN_CITIES = [
    "Pune", "Mumbai", "Delhi", "Bangalore", "Hyderabad",
    "Ahmedabad", "Chennai", "Kolkata", "Surat", "Jaipur", 
    "Indore", "Nagpur", "Kochi", "Coimbatore", "Chandigarh",
    "Noida", "Gurgaon", "Bhubaneswar", "Mysore", "Vadodara",
    "Lucknow", "Vizag", "Navi Mumbai", "Goa", "Thane",
]

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/138.0 Safari/537.36"
)

REQUEST_HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept-Language": "en-US,en;q=0.9",
}

HEADLESS = False
SLOW_MO_MS = 250
PAGE_TIMEOUT_MS = 25_000
VERIFY_TIMEOUT_SECONDS = 12

MAX_RESULTS_PER_QUERY = 30
MAX_INTERNAL_PAGES = 5
MAX_META_DESCRIPTION_CHARS = 300
MAX_RESPONSE_SNIFF_CHARS = 20_000

RETRY_ATTEMPTS = 3
RETRY_BACKOFF_SECONDS = 2.0
CHECKPOINT_EVERY_QUERIES = 5

# A directory that serves this many challenge pages in a row is left alone for
# the rest of the run rather than being hammered with more requests.
MAX_CONSECUTIVE_BLOCKS = 3

# Hosts that are never a company's official website.
SOCIAL_DOMAINS = {
    "facebook.com", "fb.com", "fb.me", "instagram.com", "linkedin.com",
    "twitter.com", "x.com", "t.co", "youtube.com", "youtu.be",
    "pinterest.com", "tiktok.com", "snapchat.com", "threads.net",
    "whatsapp.com", "wa.me", "t.me", "telegram.me", "vimeo.com",
    "behance.net", "dribbble.com", "reddit.com", "quora.com", "tumblr.com",
}

DIRECTORY_DOMAINS = {
    "clutch.co", "goodfirms.co", "designrush.com", "sortlist.com",
    "agencyspotter.com", "upcity.com", "expertise.com", "trustpilot.com",
    "yelp.com", "bbb.org", "crunchbase.com", "glassdoor.com",
    "glassdoor.co.in", "indeed.com", "ambitionbox.com", "justdial.com",
    "sulekha.com", "manta.com", "yellowpages.com", "topdevelopers.co",
    "themanifest.com", "g2.com", "capterra.com", "producthunt.com",
}

AI_DOMAINS = {
    "openai.com", "chatgpt.com", "anthropic.com", "claude.ai",
    "perplexity.ai", "gemini.google.com", "bard.google.com",
    "copilot.microsoft.com", "midjourney.com", "huggingface.co",
    "character.ai", "jasper.ai", "copy.ai", "writesonic.com",
}

DOCUMENTATION_DOMAINS = {
    "wikipedia.org", "wikimedia.org", "medium.com", "substack.com",
    "notion.so", "notion.site", "gitbook.io", "readthedocs.io",
    "readthedocs.org", "docs.google.com", "drive.google.com",
    "sites.google.com", "github.com", "gitlab.com", "stackoverflow.com",
    "w3.org", "mozilla.org", "schema.org", "creativecommons.org",
}

INFRASTRUCTURE_DOMAINS = {
    "google.com", "gstatic.com", "googleapis.com", "googletagmanager.com",
    "google-analytics.com", "doubleclick.net", "apple.com", "microsoft.com",
    "adobe.com", "cloudflare.com", "cloudfront.net", "gravatar.com",
    "jquery.com", "bootstrapcdn.com", "jsdelivr.net", "unpkg.com",
    "duckduckgo.com", "bing.com", "yahoo.com", "amazonaws.com",
    "wordpress.org", "wix.com", "squarespace.com", "godaddy.com",
    "sedo.com", "namecheap.com", "hostinger.com", "bluehost.com",
}

BLOCKED_DOMAINS = (
    SOCIAL_DOMAINS
    | DIRECTORY_DOMAINS
    | AI_DOMAINS
    | DOCUMENTATION_DOMAINS
    | INFRASTRUCTURE_DOMAINS
)

# Anchor text that reliably marks the official website link on a profile page.
WEBSITE_LINK_HINTS = (
    "visit website",
    "visit site",
    "view website",
    "go to website",
    "company website",
    "official website",
    "visit profile website",
    "website",
)

# Path fragments used by directories to wrap outbound links in a redirect.
REDIRECT_MARKERS = ("visit-website", "/go/", "/redirect", "/out/", "/link/", "/away")

# Query parameters that carry the real destination inside a redirect URL.
REDIRECT_PARAMS = ("u", "url", "redirect", "target", "link", "to", "r", "dest")

# File extensions that are assets rather than websites.
ASSET_SUFFIXES = (".jpg", ".jpeg", ".png", ".gif", ".svg", ".webp", ".pdf", ".zip")

# Text that marks a bot check or challenge page rather than real content.
BLOCK_SIGNATURES = (
    "verify you are human",
    "verifying you are human",
    "checking your browser",
    "just a moment",
    "attention required",
    "enable javascript and cookies",
    "unusual traffic",
    "access denied",
    "security check",
    "captcha",
    "cf-challenge",
    "cf_chl_",
    "ray id",
)

# Titles that mean the page exists but holds no company.
SOFT_404_SIGNATURES = (
    "404", "not found", "page missing", "page unavailable",
    "domain for sale", "buy this domain", "this domain is parked",
    "under construction", "coming soon", "default web page",
    "account suspended", "site not configured",
)

# Company names that are placeholders rather than real companies.
PLACEHOLDER_NAMES = {
    "", "home", "homepage", "untitled", "index", "welcome", "loading",
    "error", "404", "not found", "page not found", "just a moment",
    "access denied", "attention required", "security check", "cloudflare",
    "verify", "verification", "profile", "company", "agency", "companies",
    "agencies", "reviews", "n/a", "na", "none", "null", "unknown",
}

# Titles from directory listing pages rather than a single company.
LISTING_NAME_PATTERN = re.compile(
    r"\b(top|best|leading|list of|\d{1,3}\+?\s+(best|top))\b.*\b(agenc|compan|firm|service)",
    re.IGNORECASE,
)

INTERNAL_PAGE_TOKENS = (
    "about", "services", "contact", "case", "client", "work", "portfolio",
    "who-we-are", "our-story", "team", "expertise", "solutions",
)

# Anchor labels that identify the same pages when the URL gives nothing away.
INTERNAL_PAGE_LABELS = (
    "about", "about us", "who we are", "our story", "services", "what we do",
    "contact", "contact us", "get in touch", "work", "our work", "portfolio",
    "case studies", "clients", "team", "expertise", "solutions",
)

# Words dropped before matching a company name against its own domain.
COMPANY_NAME_STOPWORDS = {
    "pvt", "private", "ltd", "limited", "llp", "inc", "incorporated", "co",
    "company", "corp", "the", "and", "solutions", "services", "group",
    "technologies", "technology", "agency", "agencies", "consulting",
}

# Boilerplate removed from directory page titles to recover the company name.
NAME_NOISE_PATTERN = re.compile(
    r"\b(reviews?|profile|pricing|clutch\.co|clutch|goodfirms|designrush|"
    r"top\s+\d+|agency\s+profile)\b",
    re.IGNORECASE,
)

TITLE_PATTERN = re.compile(r"<title[^>]*>(.*?)</title>", re.IGNORECASE | re.DOTALL)

CONTROL_CHARS_PATTERN = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

PRIVATE_HOST_PATTERN = re.compile(
    r"^(localhost|127\.|0\.0\.0\.0|10\.|192\.168\.|172\.(1[6-9]|2\d|3[01])\.)",
    re.IGNORECASE,
)

IP_HOST_PATTERN = re.compile(r"^\d{1,3}(\.\d{1,3}){3}$")

# Use the bundled public suffix snapshot so startup never depends on the network.
DOMAIN_EXTRACTOR = tldextract.TLDExtract(suffix_list_urls=())

# Confidence is scored from evidence. Each item is worth the points shown.
EVIDENCE_POINTS = {
    "explicit-website-link": 2,
    "website-reachable": 2,
    "name-matches-site": 2,
    "meta-description": 1,
    "internal-pages": 1,
}
CONFIDENCE_HIGH_SCORE = 6
CONFIDENCE_MEDIUM_SCORE = 4

MASTER_COLUMNS = [
    "Company Name",
    "Website",
    "Source",
    "Source URL",
    "Search Query",
    "Discovery Confidence",
    "Meta Description",
    "Internal Pages",
    "Keyword",
    "City",
    "Root Domain",
    "Verified",
    "Confidence Score",
    "Evidence",
    "Website Source",
    "HTTP Status",
    "Site Title",
    "Discovered At",
]


# ----------------------------------------------------------------------
# Directory adapters
# ----------------------------------------------------------------------

@dataclass(frozen=True)
class DirectoryAdapter:
    """
    Everything that differs between the three directories.

    DesignRush was the reliable adapter in the v0.1 smoke test, so its layered
    approach - specific selectors first, labelled anchors second, scored
    external links last - is the shape all three now use. The selector lists
    are per-directory because the markup differs; the fallbacks are shared.
    """

    name: str
    search_operator: str
    domain: str
    profile_markers: tuple[str, ...]
    website_selectors: tuple[str, ...]
    name_selectors: tuple[str, ...]


DIRECTORIES: dict[str, DirectoryAdapter] = {
    "Clutch": DirectoryAdapter(
        name="Clutch",
        search_operator="site:clutch.co",
        domain="clutch.co",
        profile_markers=("/profile/",),
        website_selectors=(
            "a[data-link_text='Visit website']",
            "a[data-testid='visit-website']",
            "a.website-link__item",
            ".website-link a[href]",
            "a.sg-button--website",
            "a[href*='/go/']",
            "li.website-link a[href]",
        ),
        name_selectors=("h1.profile-header__title", "h1[data-testid='provider-name']", "h1"),
    ),
    "GoodFirms": DirectoryAdapter(
        name="GoodFirms",
        search_operator="site:goodfirms.co",
        domain="goodfirms.co",
        profile_markers=("/company/",),
        website_selectors=(
            "a[href*='/visit-website/']",
            "a.visit-website",
            "a.visit-website-link",
            ".profile-visit-website a[href]",
            ".company-visit-website a[href]",
            "a[title*='Visit Website']",
            "a[aria-label*='Visit Website']",
            ".detail-visit-website a[href]",
            "a.entity-website",
        ),
        name_selectors=(
            "h1.profile-name",
            "h1.entity-name",
            ".profile-header h1",
            "h1[itemprop='name']",
            "h1",
        ),
    ),
    "DesignRush": DirectoryAdapter(
        name="DesignRush",
        search_operator="site:designrush.com",
        domain="designrush.com",
        profile_markers=("/agency/profile/",),
        website_selectors=(
            "a[data-testid='agency-website']",
            "a.agency-website",
            "a.visit-website",
            ".agency-profile__website a[href]",
            "a[href][rel*='nofollow'][target='_blank']",
        ),
        name_selectors=("h1.agency-name", ".agency-profile__title h1", "h1"),
    ),
}


# ----------------------------------------------------------------------
# Failure reasons (used in logs and the run report)
# ----------------------------------------------------------------------

REASON_SECURITY_VERIFICATION = "security verification"
REASON_SELECTOR_MISSING = "selector missing"
REASON_TIMEOUT = "timeout"
REASON_REDIRECTED = "redirected"
REASON_DIRECTORY_BLOCKED = "directory blocked earlier in run"
REASON_UNREADABLE = "page content unreadable"
REASON_NO_CANDIDATE_LINKS = "no external links on page"


# ----------------------------------------------------------------------
# Logging
# ----------------------------------------------------------------------

def setup_logger() -> logging.Logger:
    """Configure the file + console logger used across the whole run."""
    logger = logging.getLogger("discovery")
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        formatter = logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(funcName)s | %(message)s"
        )

        file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.WARNING)
        console_handler.setFormatter(logging.Formatter("%(levelname)s | %(message)s"))
        logger.addHandler(console_handler)

    return logger


LOGGER = setup_logger()


# ----------------------------------------------------------------------
# Data model
# ----------------------------------------------------------------------

@dataclass
class SearchQuery:
    """One generated DuckDuckGo query targeting a single directory."""

    directory: str
    keyword: str
    city: str
    query: str


@dataclass
class SearchResult:
    """A single DuckDuckGo result link."""

    title: str
    url: str


@dataclass
class WebsiteMatch:
    """The outcome of looking for the official website on a profile page."""

    url: str = ""
    method: str = ""
    reason: str = ""


@dataclass
class SiteCheck:
    """The outcome of fetching a candidate company website."""

    reachable: bool = False
    status_code: int = 0
    final_url: str = ""
    title: str = ""
    reason: str = ""


@dataclass
class CompanyRecord:
    """A discovered company. Every field written to output lives here."""

    company_name: str = ""
    website: str = ""
    source: str = ""
    source_url: str = ""
    search_query: str = ""
    discovery_confidence: str = ""
    meta_description: str = ""
    internal_pages: list[str] = field(default_factory=list)
    keyword: str = ""
    city: str = ""
    root_domain: str = ""
    verified: bool = False
    website_source: str = ""
    http_status: int = 0
    site_title: str = ""
    evidence: list[str] = field(default_factory=list)
    confidence_score: int = 0
    discovered_at: str = ""


@dataclass
class Runtime:
    """Single source of truth for the run: work list, results and counters."""

    queries: list[SearchQuery]
    records: list[CompanyRecord] = field(default_factory=list)
    seen_domains: set[str] = field(default_factory=set)
    visited_urls: set[str] = field(default_factory=set)
    error_messages: list[str] = field(default_factory=list)
    skip_reasons: Counter[str] = field(default_factory=Counter)
    blocked_directories: set[str] = field(default_factory=set)
    consecutive_blocks: Counter[str] = field(default_factory=Counter)
    query_index: int = 0
    started_at: float = field(default_factory=time.time)
    stop_requested: bool = False
    queries_completed: int = 0
    results_seen: int = 0
    results_kept: int = 0
    profiles_visited: int = 0
    profiles_blocked: int = 0
    websites_checked: int = 0
    records_rejected: int = 0
    duplicates_removed: int = 0
    errors: int = 0

    @property
    def companies_found(self) -> int:
        """Companies kept after validation and deduplication."""
        return len(self.records)

    @property
    def websites_verified(self) -> int:
        """Exported companies whose website responded. Never exceeds records."""
        return sum(1 for record in self.records if record.verified)


def initialize_runtime() -> Runtime:
    """Create the shared runtime state used by the main loop."""
    return Runtime(queries=generate_queries())


def record_error(runtime: Runtime, message: str) -> None:
    """Log an error once and count it in the single statistics system."""
    runtime.errors += 1
    runtime.error_messages.append(f"{datetime.now().isoformat(timespec='seconds')} | {message}")
    LOGGER.error(message)


def record_skip(runtime: Runtime, url: str, reason: str) -> None:
    """Log why a candidate was dropped, with the reason kept for the report."""
    runtime.skip_reasons[reason] += 1
    LOGGER.info("Skipped %s\nReason: %s", url, reason)


# ----------------------------------------------------------------------
# Query generation
# ----------------------------------------------------------------------

def generate_queries() -> list[SearchQuery]:
    """Build every directory x keyword x city query combination."""
    queries: list[SearchQuery] = []

    for adapter in DIRECTORIES.values():
        for keyword in SEARCH_KEYWORDS:
            for city in INDIAN_CITIES:
                queries.append(
                    SearchQuery(
                        directory=adapter.name,
                        keyword=keyword,
                        city=city,
                        query=f"{adapter.search_operator} {keyword} {city}",
                    )
                )

    return queries


# ----------------------------------------------------------------------
# URL helpers
# ----------------------------------------------------------------------

def root_domain(url: str) -> str:
    """Return the registrable domain (example.co.in) used for deduplication."""
    extracted = DOMAIN_EXTRACTOR(url)
    return ".".join(part for part in (extracted.domain, extracted.suffix) if part).lower()


def is_blocked_domain(url: str) -> bool:
    """True for social networks, directories, AI tools and other non-companies."""
    domain = root_domain(url)
    if not domain:
        return True
    return domain in BLOCKED_DOMAINS


def is_private_host(url: str) -> bool:
    """True for localhost, bare IPs and private network addresses."""
    host = urlparse(url).netloc.split(":")[0].lower()
    if not host:
        return True
    if PRIVATE_HOST_PATTERN.match(host) or IP_HOST_PATTERN.match(host):
        return True
    return host.endswith((".local", ".internal", ".test", ".invalid", ".localhost"))


def is_directory_result(adapter: DirectoryAdapter, url: str) -> bool:
    """Ensure a search result belongs to the directory currently being searched."""
    return root_domain(url) == adapter.domain


def is_profile_url(adapter: DirectoryAdapter, url: str) -> bool:
    """True when the URL points at a single company page rather than a listing."""
    path = urlparse(url).path.lower()
    return any(marker in path for marker in adapter.profile_markers)


def looks_like_redirect(url: str) -> bool:
    """True when a directory link wraps the destination instead of linking it."""
    lowered = url.lower()
    return any(marker in lowered for marker in REDIRECT_MARKERS)


def unwrap_redirect(url: str) -> str:
    """Pull the real destination out of a directory redirect URL when present."""
    query = parse_qs(urlparse(url).query)

    for parameter in REDIRECT_PARAMS:
        for raw_value in query.get(parameter, []):
            candidate = unquote(raw_value).strip()
            if candidate.startswith("http"):
                return candidate

    return url


def is_plausible_website(url: str) -> bool:
    """Reject assets, blocked hosts and anything that is not a real http(s) URL."""
    parsed = urlparse(url)

    if parsed.scheme not in ("http", "https"):
        return False
    if "." not in parsed.netloc:
        return False
    if parsed.path.lower().endswith(ASSET_SUFFIXES):
        return False
    if is_private_host(url):
        return False

    return not is_blocked_domain(url)


def clean_company_name(name: str) -> str:
    """Normalize a directory page title down to the company name."""
    name = re.split(r"\s*[|–—]\s*", name)[0]
    name = NAME_NOISE_PATTERN.sub("", name)
    name = re.sub(r"\s*-\s*(top|best|leading)\b.*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"[\s,;:\-]+$", "", name)
    return re.sub(r"\s+", " ", name).strip()


def name_tokens(name: str) -> list[str]:
    """Split a company name into comparable tokens, dropping legal boilerplate."""
    words = re.findall(r"[a-z0-9]+", name.lower())
    return [word for word in words if len(word) > 2 and word not in COMPANY_NAME_STOPWORDS]


def name_matches_site(record: CompanyRecord) -> bool:
    """True when the company name is visible in its own domain or page title."""
    tokens = name_tokens(record.company_name)
    if not tokens:
        return False

    domain_slug = re.sub(r"[^a-z0-9]", "", root_domain(record.website).split(".")[0])
    title = record.site_title.lower()

    for token in tokens:
        if token in domain_slug or token in title:
            return True

    joined = "".join(tokens)
    return bool(joined) and joined in domain_slug


# ----------------------------------------------------------------------
# Browser lifecycle
# ----------------------------------------------------------------------

def launch_browser(headless: bool = HEADLESS) -> tuple[Playwright, Browser, BrowserContext, Page]:
    """Start Playwright with a visible Chromium and return the browsing page."""
    playwright = sync_playwright().start()

    try:
        browser = playwright.chromium.launch(headless=headless, slow_mo=SLOW_MO_MS)
        context = browser.new_context(
            user_agent=USER_AGENT,
            viewport={"width": 1400, "height": 900},
        )
        page = context.new_page()
        page.set_default_timeout(PAGE_TIMEOUT_MS)
    except Exception:
        playwright.stop()
        raise

    LOGGER.info("Chromium launched (headless=%s).", headless)
    return playwright, browser, context, page


def close_browser(playwright: Playwright, browser: Browser, context: BrowserContext) -> None:
    """Shut the browser down step by step so one failure cannot leak a process."""
    LOGGER.info("Closing browser.")

    for label, close in (("context", context.close), ("browser", browser.close)):
        try:
            close()
        except PlaywrightError as error:
            LOGGER.warning("Failed to close %s cleanly: %s", label, error)

    try:
        playwright.stop()
    except Exception as error:  # noqa: BLE001 - shutdown must never raise
        LOGGER.warning("Failed to stop Playwright cleanly: %s", error)


def navigate(page: Page, url: str, runtime: Runtime) -> str:
    """
    Navigate with retries.

    Returns an empty string on success, otherwise the reason it failed, so the
    caller can log something more useful than a bare failure.
    """
    reason = REASON_TIMEOUT

    for attempt in range(1, RETRY_ATTEMPTS + 1):
        if runtime.stop_requested:
            return "shutdown requested"

        try:
            page.goto(url, wait_until="domcontentloaded")
            return ""
        except PlaywrightTimeoutError:
            reason = REASON_TIMEOUT
            LOGGER.warning("Navigation attempt %d/%d timed out for %s", attempt, RETRY_ATTEMPTS, url)
        except PlaywrightError as error:
            reason = str(error).splitlines()[0]
            LOGGER.warning(
                "Navigation attempt %d/%d failed for %s: %s",
                attempt, RETRY_ATTEMPTS, url, reason,
            )

        if attempt < RETRY_ATTEMPTS:
            time.sleep(RETRY_BACKOFF_SECONDS * attempt)

    return reason


def read_page_html(page: Page, runtime: Runtime) -> str:
    """Return the rendered HTML, or an empty string if it cannot be read."""
    try:
        return page.content()
    except PlaywrightError as error:
        record_error(runtime, f"Could not read page content for {page.url}: {error}")
        return ""


def looks_blocked(html: str) -> bool:
    """True when a page is a bot check or challenge rather than real content."""
    sample = html[:MAX_RESPONSE_SNIFF_CHARS].lower()
    return any(signature in sample for signature in BLOCK_SIGNATURES)


# ----------------------------------------------------------------------
# DuckDuckGo search
# ----------------------------------------------------------------------

RESULT_LINK_SELECTORS = (
    "a[data-testid='result-title-a']",
    "article[data-testid='result'] h2 a[href]",
    "h2 a[href]",
    "a.result__a",
)


def run_search(page: Page, query: str, runtime: Runtime) -> list[SearchResult]:
    """Execute a DuckDuckGo search and return the result links it rendered."""
    LOGGER.info("Searching: %s", query)

    search_url = f"https://duckduckgo.com/?q={quote(query)}&ia=web"
    failure = navigate(page, search_url, runtime)
    if failure:
        record_error(runtime, f"Search navigation failed\nQuery: {query}\nReason: {failure}")
        return []

    links = _wait_for_result_links(page)
    if links is None:
        record_error(runtime, f"Search returned no results\nQuery: {query}\nReason: {REASON_SELECTOR_MISSING}")
        return []

    results: list[SearchResult] = []

    for index in range(min(links.count(), MAX_RESULTS_PER_QUERY)):
        link = links.nth(index)

        try:
            href = link.get_attribute("href")
            title = (link.inner_text() or "").strip()
        except PlaywrightError as error:
            LOGGER.warning("Could not read result %d: %s", index, error)
            continue

        if href:
            results.append(SearchResult(title=title, url=href))

    LOGGER.info("Returned %d search results.", len(results))
    return results


def _wait_for_result_links(page: Page):
    """Try each known result layout and return the first locator that matches."""
    for selector in RESULT_LINK_SELECTORS:
        locator = page.locator(selector)
        try:
            locator.first.wait_for(
                state="attached",
                timeout=PAGE_TIMEOUT_MS // len(RESULT_LINK_SELECTORS),
            )
        except PlaywrightTimeoutError:
            continue

        if locator.count():
            LOGGER.debug("Matched result layout with selector: %s", selector)
            return locator

    return None


def filter_results(
    search: SearchQuery,
    results: list[SearchResult],
    runtime: Runtime,
) -> list[SearchResult]:
    """Keep only unvisited company profile pages on the searched directory."""
    adapter = DIRECTORIES[search.directory]

    if adapter.name in runtime.blocked_directories:
        runtime.skip_reasons[REASON_DIRECTORY_BLOCKED] += len(results)
        LOGGER.info("Skipping %d results\nReason: %s", len(results), REASON_DIRECTORY_BLOCKED)
        return []

    filtered: list[SearchResult] = []
    seen_in_batch: set[str] = set()

    for result in results:
        url = result.url.strip()

        if not url or url in seen_in_batch or url in runtime.visited_urls:
            continue
        if not is_directory_result(adapter, url):
            continue
        if not is_profile_url(adapter, url):
            LOGGER.debug("Skipping non-profile URL: %s", url)
            continue

        seen_in_batch.add(url)
        filtered.append(result)

    LOGGER.info(
        "Filtered %d -> %d results for %s.",
        len(results), len(filtered), adapter.name,
    )
    return filtered


# ----------------------------------------------------------------------
# Directory parsing
# ----------------------------------------------------------------------

def resolve_candidate(href: str, source_url: str) -> str:
    """Turn a raw href into an absolute URL, unwrapping directory redirects."""
    absolute = urljoin(source_url, href.strip())
    return unwrap_redirect(absolute) if looks_like_redirect(absolute) else absolute


def extract_company_name(soup: BeautifulSoup, adapter: DirectoryAdapter, fallback: str) -> str:
    """Read the company name using the directory's own heading selectors."""
    for selector in adapter.name_selectors:
        element = soup.select_one(selector)
        if element:
            name = clean_company_name(element.get_text(" ", strip=True))
            if name:
                return name

    if soup.title:
        name = clean_company_name(soup.title.get_text(strip=True))
        if name:
            return name

    return clean_company_name(fallback)


def extract_official_website(
    soup: BeautifulSoup,
    adapter: DirectoryAdapter,
    source_url: str,
) -> WebsiteMatch:
    """
    Find the company's own website on a directory profile page.

    Three layers, most reliable first: the directory's own website selectors,
    then anchors explicitly labelled "Visit Website", then external links that
    survive the blocklist. Social networks, directories, AI tools, docs hosts
    and infrastructure domains are never treated as a company website.
    """
    match = _match_by_selector(soup, adapter, source_url)
    if match.url:
        return match

    match = _match_by_label(soup, source_url)
    if match.url:
        return match

    return _match_by_external_link(soup, source_url)


def _match_by_selector(
    soup: BeautifulSoup,
    adapter: DirectoryAdapter,
    source_url: str,
) -> WebsiteMatch:
    """Layer 1: the selectors known to hold the website link on this directory."""
    for selector in adapter.website_selectors:
        for anchor in soup.select(selector):
            href = anchor.get("href", "")
            if not href:
                continue

            candidate = resolve_candidate(href, source_url)
            if is_plausible_website(candidate):
                return WebsiteMatch(url=candidate, method="selector")

    return WebsiteMatch(reason=REASON_SELECTOR_MISSING)


def _match_by_label(soup: BeautifulSoup, source_url: str) -> WebsiteMatch:
    """Layer 2: any anchor whose visible text says it leads to the website."""
    for anchor in soup.find_all("a", href=True):
        label = " ".join(
            part.lower()
            for part in (
                anchor.get_text(" ", strip=True),
                anchor.get("aria-label", ""),
                anchor.get("title", ""),
            )
            if part
        )

        if not any(hint in label for hint in WEBSITE_LINK_HINTS):
            continue

        candidate = resolve_candidate(anchor["href"], source_url)
        if is_plausible_website(candidate):
            return WebsiteMatch(url=candidate, method="label")

    return WebsiteMatch(reason=REASON_SELECTOR_MISSING)


def _match_by_external_link(soup: BeautifulSoup, source_url: str) -> WebsiteMatch:
    """
    Layer 3: the best remaining external link.

    Links are scored rather than taken in document order, because the first
    external link on a profile page is usually a social icon or a partner
    badge. The most repeated external domain on the page wins.
    """
    counts: Counter[str] = Counter()
    first_seen: dict[str, str] = {}

    for anchor in soup.find_all("a", href=True):
        href = anchor["href"].strip()
        if not href or href.startswith(("#", "mailto:", "tel:", "javascript:")):
            continue

        candidate = resolve_candidate(href, source_url)
        if not is_plausible_website(candidate):
            continue

        domain = root_domain(candidate)
        counts[domain] += 1
        first_seen.setdefault(domain, candidate)

    if not counts:
        return WebsiteMatch(reason=REASON_NO_CANDIDATE_LINKS)

    best_domain, _count = counts.most_common(1)[0]
    return WebsiteMatch(url=first_seen[best_domain], method="heuristic")


def open_profile(
    page: Page,
    adapter: DirectoryAdapter,
    result: SearchResult,
    runtime: Runtime,
) -> BeautifulSoup | None:
    """
    Open a directory profile page and return its parsed markup.

    Challenge pages are detected and counted. A directory that serves several
    in a row is dropped for the rest of the run instead of being retried.
    """
    runtime.visited_urls.add(result.url)

    failure = navigate(page, result.url, runtime)
    if failure:
        record_skip(runtime, result.url, failure)
        return None

    html = read_page_html(page, runtime)
    if not html:
        record_skip(runtime, result.url, REASON_UNREADABLE)
        return None

    if looks_blocked(html):
        runtime.profiles_blocked += 1
        note_directory_block(runtime, adapter)
        record_skip(runtime, result.url, REASON_SECURITY_VERIFICATION)
        return None

    if root_domain(page.url) != adapter.domain:
        record_skip(runtime, result.url, f"{REASON_REDIRECTED} to {page.url}")
        return None

    runtime.consecutive_blocks[adapter.name] = 0
    runtime.profiles_visited += 1
    return BeautifulSoup(html, "html.parser")


def note_directory_block(runtime: Runtime, adapter: DirectoryAdapter) -> None:
    """Count a challenge page and retire the directory once they keep coming."""
    runtime.consecutive_blocks[adapter.name] += 1

    if runtime.consecutive_blocks[adapter.name] < MAX_CONSECUTIVE_BLOCKS:
        return

    if adapter.name not in runtime.blocked_directories:
        runtime.blocked_directories.add(adapter.name)
        LOGGER.warning(
            "Dropping %s for the rest of this run\nReason: %s on %d consecutive pages",
            adapter.name, REASON_SECURITY_VERIFICATION, MAX_CONSECUTIVE_BLOCKS,
        )


def build_company_record(
    soup: BeautifulSoup,
    search: SearchQuery,
    result: SearchResult,
    runtime: Runtime,
) -> CompanyRecord | None:
    """Turn a parsed profile page into a record, or explain why it cannot."""
    adapter = DIRECTORIES[search.directory]

    record = CompanyRecord(
        source=adapter.name,
        source_url=result.url,
        search_query=search.query,
        keyword=search.keyword,
        city=search.city,
        discovered_at=datetime.now().isoformat(timespec="seconds"),
    )

    record.company_name = extract_company_name(soup, adapter, result.title)

    match = extract_official_website(soup, adapter, result.url)
    if not match.url:
        LOGGER.info("Official website not found on %s\nReason: %s", result.url, match.reason)
        runtime.skip_reasons[f"website not found: {match.reason}"] += 1
        return None

    record.website = match.url
    record.website_source = match.method
    record.root_domain = root_domain(match.url)
    return record


# ----------------------------------------------------------------------
# Website verification and validation
# ----------------------------------------------------------------------

def check_website(url: str) -> SiteCheck:
    """Fetch a candidate website and report status, final URL and title."""
    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            response = requests.get(
                url,
                headers=REQUEST_HEADERS,
                timeout=VERIFY_TIMEOUT_SECONDS,
                allow_redirects=True,
            )
        except requests.RequestException as error:
            LOGGER.debug(
                "Verification attempt %d/%d failed for %s: %s",
                attempt, RETRY_ATTEMPTS, url, error,
            )
            if attempt < RETRY_ATTEMPTS:
                time.sleep(RETRY_BACKOFF_SECONDS * attempt)
            continue

        return SiteCheck(
            reachable=response.status_code < 400,
            status_code=response.status_code,
            final_url=response.url,
            title=extract_title(response.text),
        )

    return SiteCheck(reason="unreachable")


def extract_title(html: str) -> str:
    """Pull the page title out of raw HTML without a full parse."""
    found = TITLE_PATTERN.search(html[:MAX_RESPONSE_SNIFF_CHARS])
    if not found:
        return ""
    return re.sub(r"\s+", " ", found.group(1)).strip()[:200]


def validate_record(record: CompanyRecord, check: SiteCheck) -> str:
    """
    Decide whether a candidate is a genuine company.

    Returns an empty string when the record may be exported, otherwise the
    reason it was rejected.
    """
    name = record.company_name.strip()

    if name.lower() in PLACEHOLDER_NAMES or len(name) < 2:
        return "placeholder company name"
    if not re.search(r"[A-Za-z]", name):
        return "company name has no letters"
    if LISTING_NAME_PATTERN.search(name):
        return "directory listing page, not a company"

    if not record.website or not record.root_domain:
        return "invalid domain"
    if is_private_host(record.website):
        return "localhost or private address"
    if root_domain(record.website) in DIRECTORY_DOMAINS:
        return "directory page, not a company"
    if root_domain(record.website) in AI_DOMAINS:
        return "AI or chat website"
    if is_blocked_domain(record.website):
        return "blocked domain"

    if check.status_code >= 400:
        return f"website returned HTTP {check.status_code}"

    if check.final_url:
        if is_private_host(check.final_url):
            return "redirected to a private address"
        if is_blocked_domain(check.final_url):
            return f"{REASON_REDIRECTED} to a blocked domain"

    title = check.title.lower()
    if title and any(signature in title for signature in SOFT_404_SIGNATURES):
        return f"placeholder or missing page: {check.title[:60]}"

    return ""


def apply_site_check(record: CompanyRecord, check: SiteCheck) -> None:
    """Copy verification results onto the record before it is validated."""
    record.verified = check.reachable
    record.http_status = check.status_code
    record.site_title = check.title

    # Follow the redirect chain so the exported website is the one that answers.
    if check.reachable and check.final_url and is_plausible_website(check.final_url):
        record.website = check.final_url
        record.root_domain = root_domain(check.final_url)


# ----------------------------------------------------------------------
# Enrichment
# ----------------------------------------------------------------------

def extract_meta_description(soup: BeautifulSoup) -> str:
    """Read the page description, falling back through the usual meta tags."""
    candidates = (
        {"name": re.compile(r"^description$", re.IGNORECASE)},
        {"property": re.compile(r"^og:description$", re.IGNORECASE)},
        {"name": re.compile(r"^twitter:description$", re.IGNORECASE)},
        {"itemprop": re.compile(r"^description$", re.IGNORECASE)},
    )

    for attributes in candidates:
        meta = soup.find("meta", attrs=attributes)
        if meta and meta.get("content", "").strip():
            text = re.sub(r"\s+", " ", meta["content"]).strip()
            return text[:MAX_META_DESCRIPTION_CHARS]

    return ""


def discover_internal_pages(soup: BeautifulSoup, base_url: str) -> list[str]:
    """
    Collect a small set of useful same-site pages from the company website.

    Both the URL and the link text are matched, because plenty of sites label
    a link "About Us" while pointing it at /who-we-are or /company. Depth stays
    at one page: this reads the homepage's links and never follows them.
    """
    site_domain = root_domain(base_url)
    found: dict[str, str] = {}

    for anchor in soup.find_all("a", href=True):
        href = anchor["href"].strip()
        if not href or href.startswith(("#", "mailto:", "tel:", "javascript:")):
            continue

        absolute = normalize_internal_url(urljoin(base_url, href))
        if not absolute or root_domain(absolute) != site_domain:
            continue

        path = urlparse(absolute).path.lower().strip("/")
        if not path:
            continue

        label = anchor.get_text(" ", strip=True).lower()
        url_match = any(token in path for token in INTERNAL_PAGE_TOKENS)
        label_match = any(label == text or label.startswith(text) for text in INTERNAL_PAGE_LABELS)

        if url_match or label_match:
            found.setdefault(path, absolute)

        if len(found) >= MAX_INTERNAL_PAGES:
            break

    return list(found.values())


def normalize_internal_url(url: str) -> str:
    """Drop query strings and fragments so the same page is not listed twice."""
    parsed = urlparse(url)
    if parsed.scheme not in ("http", "https"):
        return ""
    return f"{parsed.scheme}://{parsed.netloc}{parsed.path}".rstrip("/")


def enrich_company_record(page: Page, record: CompanyRecord, runtime: Runtime) -> None:
    """Add the meta description and a few internal page links to a record."""
    failure = navigate(page, record.website, runtime)
    if failure:
        LOGGER.info("Enrichment skipped for %s\nReason: %s", record.website, failure)
        runtime.skip_reasons[f"enrichment failed: {failure}"] += 1
        return

    html = read_page_html(page, runtime)
    if not html:
        runtime.skip_reasons[f"enrichment failed: {REASON_UNREADABLE}"] += 1
        return

    soup = BeautifulSoup(html, "html.parser")

    record.meta_description = extract_meta_description(soup)
    record.internal_pages = discover_internal_pages(soup, page.url or record.website)

    if not record.site_title and soup.title:
        record.site_title = soup.title.get_text(strip=True)[:200]


# ----------------------------------------------------------------------
# Confidence
# ----------------------------------------------------------------------

def collect_evidence(record: CompanyRecord) -> list[str]:
    """List what the run actually proved about this company."""
    evidence: list[str] = []

    if record.website_source in ("selector", "label"):
        evidence.append("explicit-website-link")
    if record.verified:
        evidence.append("website-reachable")
    if name_matches_site(record):
        evidence.append("name-matches-site")
    if record.meta_description:
        evidence.append("meta-description")
    if record.internal_pages:
        evidence.append("internal-pages")

    return evidence


def grade_confidence(record: CompanyRecord) -> tuple[str, int, list[str]]:
    """
    Score the collected evidence and turn it into a confidence band.

    A website that never answered is capped at Low no matter what else was
    found: without a reachable site the rest of the evidence is unconfirmed.
    """
    evidence = collect_evidence(record)
    score = sum(EVIDENCE_POINTS[item] for item in evidence)

    if not record.verified:
        band = "Low"
    elif score >= CONFIDENCE_HIGH_SCORE:
        band = "High"
    elif score >= CONFIDENCE_MEDIUM_SCORE:
        band = "Medium"
    else:
        band = "Low"

    return band, score, evidence


# ----------------------------------------------------------------------
# Collection, progress and checkpointing
# ----------------------------------------------------------------------

def safe_append(runtime: Runtime, record: CompanyRecord) -> bool:
    """
    Add a record to the run, deduplicating by root domain.

    This is the only place records enter the run, so it is the only place
    deduplication happens.
    """
    if not record.website or not record.company_name:
        return False

    domain = record.root_domain or root_domain(record.website)
    if not domain:
        return False

    if domain in runtime.seen_domains:
        runtime.duplicates_removed += 1
        LOGGER.info("Duplicate skipped: %s (%s)", record.company_name, domain)
        return False

    record.root_domain = domain
    record.company_name = clean_company_name(record.company_name)
    runtime.seen_domains.add(domain)
    runtime.records.append(record)

    LOGGER.info("Company added: %s | %s", record.company_name, record.website)
    return True


def format_eta(started_at: float, completed: int, total: int) -> str:
    """Estimate remaining runtime as HH:MM:SS from the average query duration."""
    if completed <= 0:
        return "--:--:--"

    average = (time.time() - started_at) / completed
    remaining = max(total - completed, 0)
    minutes, seconds = divmod(int(average * remaining), 60)
    hours, minutes = divmod(minutes, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def progress_line(runtime: Runtime, search: SearchQuery) -> None:
    """Print one console progress line for the query about to run."""
    total = len(runtime.queries)
    current = runtime.query_index + 1
    percent = (current / total * 100) if total else 0.0

    print(
        f"[{current:>4}/{total}] {percent:5.1f}% | "
        f"companies {runtime.companies_found:>4} | "
        f"rejected {runtime.records_rejected:>3} | "
        f"errors {runtime.errors:>3} | "
        f"ETA {format_eta(runtime.started_at, runtime.queries_completed, total)} | "
        f"{search.query}"
    )


def checkpoint(runtime: Runtime) -> None:
    """Write the autosave file so an interrupted run can resume where it stopped."""
    state = {
        "version": AUTOSAVE_VERSION,
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "query_index": runtime.query_index,
        "queries_completed": runtime.queries_completed,
        "results_seen": runtime.results_seen,
        "results_kept": runtime.results_kept,
        "profiles_visited": runtime.profiles_visited,
        "profiles_blocked": runtime.profiles_blocked,
        "websites_checked": runtime.websites_checked,
        "records_rejected": runtime.records_rejected,
        "duplicates_removed": runtime.duplicates_removed,
        "errors": runtime.errors,
        "error_messages": runtime.error_messages,
        "skip_reasons": dict(runtime.skip_reasons),
        "blocked_directories": sorted(runtime.blocked_directories),
        "visited_urls": sorted(runtime.visited_urls),
        "records": [asdict(record) for record in runtime.records],
    }

    save_autosave(state)
    LOGGER.info(
        "Checkpoint saved at query %d (%d companies).",
        runtime.query_index, runtime.companies_found,
    )


def save_autosave(state: dict) -> None:
    """Write the autosave payload to disk."""
    AUTOSAVE_FILE.write_text(json.dumps(state, indent=2), encoding="utf-8")


def load_autosave() -> dict | None:
    """Read the autosave payload, returning None when it is missing or unusable."""
    if not AUTOSAVE_FILE.exists():
        return None

    try:
        state = json.loads(AUTOSAVE_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as error:
        LOGGER.warning("Ignoring unreadable autosave file: %s", error)
        return None

    if state.get("version") != AUTOSAVE_VERSION:
        LOGGER.warning("Ignoring autosave from an incompatible version.")
        return None

    return state


def restore_runtime(runtime: Runtime, state: dict) -> None:
    """Rebuild the runtime from an autosave so counters and records survive."""
    known_fields = set(CompanyRecord.__dataclass_fields__)

    for payload in state.get("records", []):
        record = CompanyRecord(**{k: v for k, v in payload.items() if k in known_fields})
        runtime.records.append(record)
        if record.root_domain:
            runtime.seen_domains.add(record.root_domain)

    runtime.visited_urls.update(state.get("visited_urls", []))
    runtime.error_messages.extend(state.get("error_messages", []))
    runtime.skip_reasons.update(state.get("skip_reasons", {}))
    runtime.blocked_directories.update(state.get("blocked_directories", []))
    runtime.query_index = state.get("query_index", 0)
    runtime.queries_completed = state.get("queries_completed", 0)
    runtime.results_seen = state.get("results_seen", 0)
    runtime.results_kept = state.get("results_kept", 0)
    runtime.profiles_visited = state.get("profiles_visited", 0)
    runtime.profiles_blocked = state.get("profiles_blocked", 0)
    runtime.websites_checked = state.get("websites_checked", 0)
    runtime.records_rejected = state.get("records_rejected", 0)
    runtime.duplicates_removed = state.get("duplicates_removed", 0)
    runtime.errors = state.get("errors", 0)

    LOGGER.info(
        "Resumed from query %d with %d companies already discovered.",
        runtime.query_index, runtime.companies_found,
    )
    print(
        f"Resuming from query {runtime.query_index + 1} "
        f"with {runtime.companies_found} companies already discovered."
    )


def remove_autosave() -> None:
    """Delete the autosave file once the run has finished cleanly."""
    if AUTOSAVE_FILE.exists():
        AUTOSAVE_FILE.unlink()
        LOGGER.info("Autosave removed.")


# ----------------------------------------------------------------------
# Statistics and exports
# ----------------------------------------------------------------------

def build_run_statistics(runtime: Runtime) -> dict:
    """
    Derive every reported number from the runtime, the single state object.

    Anything describing exported companies is counted from runtime.records, so
    those totals can never exceed the number of companies found.
    """
    return {
        "companies_found": runtime.companies_found,
        "verified_websites": runtime.websites_verified,
        "queries_planned": len(runtime.queries),
        "queries_completed": runtime.queries_completed,
        "results_seen": runtime.results_seen,
        "results_kept": runtime.results_kept,
        "profiles_visited": runtime.profiles_visited,
        "profiles_blocked": runtime.profiles_blocked,
        "websites_checked": runtime.websites_checked,
        "records_rejected": runtime.records_rejected,
        "duplicates_removed": runtime.duplicates_removed,
        "errors": runtime.errors,
        "runtime": format_duration(time.time() - runtime.started_at),
        "blocked_directories": sorted(runtime.blocked_directories),
        "sources": dict(Counter(record.source for record in runtime.records)),
        "confidence": dict(Counter(record.discovery_confidence for record in runtime.records)),
        "website_sources": dict(Counter(record.website_source for record in runtime.records)),
        "skip_reasons": dict(runtime.skip_reasons.most_common()),
    }


def format_duration(seconds: float) -> str:
    """Format an elapsed duration as HH:MM:SS."""
    minutes, secs = divmod(int(seconds), 60)
    hours, minutes = divmod(minutes, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def print_statistics(stats: dict) -> None:
    """Print the end-of-run summary to the console."""
    print("\n========== Discovery Summary ==========\n")
    print(f"Companies Found   : {stats['companies_found']}")
    print(f"Verified Websites : {stats['verified_websites']}")
    print(f"Queries Completed : {stats['queries_completed']} / {stats['queries_planned']}")
    print(f"Profiles Visited  : {stats['profiles_visited']}")
    print(f"Profiles Blocked  : {stats['profiles_blocked']}")
    print(f"Rejected          : {stats['records_rejected']}")
    print(f"Duplicates Removed: {stats['duplicates_removed']}")
    print(f"Errors            : {stats['errors']}")
    print(f"Runtime           : {stats['runtime']}")

    if stats["blocked_directories"]:
        print(f"Dropped Sources   : {', '.join(stats['blocked_directories'])}")

    print("\nSource Breakdown")
    for source, count in stats["sources"].items():
        print(f"  {source:12} {count}")

    print("\nConfidence Breakdown")
    for level, count in stats["confidence"].items():
        print(f"  {level:8} {count}")

    if stats["skip_reasons"]:
        print("\nTop Skip Reasons")
        for reason, count in list(stats["skip_reasons"].items())[:8]:
            print(f"  {count:>4}  {reason}")

    print("\n=======================================\n")


def clean_cell(value: object) -> object:
    """Strip control characters that openpyxl refuses to write."""
    if isinstance(value, str):
        return CONTROL_CHARS_PATTERN.sub("", value)
    return value


def export_master(runtime: Runtime) -> None:
    """Write master_company_list.xlsx with the full discovery metadata."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Companies"
    sheet.append(MASTER_COLUMNS)

    for record in runtime.records:
        sheet.append([
            clean_cell(record.company_name),
            clean_cell(record.website),
            record.source,
            clean_cell(record.source_url),
            record.search_query,
            record.discovery_confidence,
            clean_cell(record.meta_description),
            clean_cell(" | ".join(record.internal_pages)),
            record.keyword,
            record.city,
            record.root_domain,
            "Yes" if record.verified else "No",
            record.confidence_score,
            ", ".join(record.evidence),
            record.website_source,
            record.http_status,
            clean_cell(record.site_title),
            record.discovered_at,
        ])

    workbook.save(OUTPUT_MASTER)
    LOGGER.info("Master list written to %s (%d rows).", OUTPUT_MASTER.name, runtime.companies_found)


def export_prospector(runtime: Runtime) -> None:
    """Write prospector_input.csv with exactly the two columns Engine 2 expects."""
    with OUTPUT_PROSPECTOR.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Company Name", "Website"])

        for record in runtime.records:
            writer.writerow([record.company_name, record.website])

    LOGGER.info("Prospector input written to %s.", OUTPUT_PROSPECTOR.name)


def create_run_report(runtime: Runtime, stats: dict) -> None:
    """Write run_report.xlsx with the run summary, breakdowns and errors."""
    workbook = Workbook()

    summary = workbook.active
    summary.title = "Run Report"
    summary.append(["Metric", "Value"])
    summary.append(["Companies Found", stats["companies_found"]])
    summary.append(["Verified Websites", stats["verified_websites"]])
    summary.append(["Queries Planned", stats["queries_planned"]])
    summary.append(["Queries Completed", stats["queries_completed"]])
    summary.append(["Search Results Seen", stats["results_seen"]])
    summary.append(["Search Results Kept", stats["results_kept"]])
    summary.append(["Profiles Visited", stats["profiles_visited"]])
    summary.append(["Profiles Blocked", stats["profiles_blocked"]])
    summary.append(["Websites Checked", stats["websites_checked"]])
    summary.append(["Records Rejected", stats["records_rejected"]])
    summary.append(["Duplicates Removed", stats["duplicates_removed"]])
    summary.append(["Errors", stats["errors"]])
    summary.append(["Runtime", stats["runtime"]])
    summary.append(["Dropped Sources", ", ".join(stats["blocked_directories"])])

    _append_breakdown(summary, "Source Breakdown", stats["sources"])
    _append_breakdown(summary, "Confidence Breakdown", stats["confidence"])
    _append_breakdown(summary, "Website Source Breakdown", stats["website_sources"])

    skips = workbook.create_sheet("Skip Reasons")
    skips.append(["Reason", "Count"])
    for reason, count in stats["skip_reasons"].items():
        skips.append([clean_cell(reason), count])

    errors = workbook.create_sheet("Errors")
    errors.append(["Error"])
    for message in runtime.error_messages:
        errors.append([clean_cell(message)])

    workbook.save(OUTPUT_RUN_REPORT)
    LOGGER.info("Run report written to %s.", OUTPUT_RUN_REPORT.name)


def _append_breakdown(sheet, heading: str, values: dict) -> None:
    """Append one titled key/value block to the summary sheet."""
    sheet.append([])
    sheet.append([heading, ""])
    for key, count in values.items():
        sheet.append([key or "unknown", count])


def finalize_run(runtime: Runtime) -> dict:
    """Write every output file and print the summary. Records are already unique."""
    stats = build_run_statistics(runtime)

    export_master(runtime)
    export_prospector(runtime)
    create_run_report(runtime, stats)
    print_statistics(stats)

    LOGGER.info("Discovery run completed with %d companies.", runtime.companies_found)
    return stats


# ----------------------------------------------------------------------
# Orchestration
# ----------------------------------------------------------------------

def process_result(
    page: Page,
    runtime: Runtime,
    search: SearchQuery,
    result: SearchResult,
) -> None:
    """Take one search result from profile page to stored, graded company."""
    adapter = DIRECTORIES[search.directory]

    soup = open_profile(page, adapter, result, runtime)
    if soup is None:
        return

    record = build_company_record(soup, search, result, runtime)
    if record is None:
        return

    runtime.websites_checked += 1
    check = check_website(record.website)
    apply_site_check(record, check)

    rejection = validate_record(record, check)
    if rejection:
        runtime.records_rejected += 1
        record_skip(runtime, record.website, rejection)
        return

    if not safe_append(runtime, record):
        return

    if record.verified:
        enrich_company_record(page, record, runtime)

    record.discovery_confidence, record.confidence_score, record.evidence = grade_confidence(record)
    LOGGER.info(
        "Graded %s as %s (%d points: %s)",
        record.company_name, record.discovery_confidence,
        record.confidence_score, ", ".join(record.evidence) or "none",
    )


def process_query(page: Page, runtime: Runtime, search: SearchQuery) -> None:
    """Run one query end to end: search, filter, scrape, verify, enrich, store."""
    results = run_search(page, search.query, runtime)
    runtime.results_seen += len(results)

    kept = filter_results(search, results, runtime)
    runtime.results_kept += len(kept)

    for result in kept:
        if runtime.stop_requested:
            return
        if search.directory in runtime.blocked_directories:
            return

        process_result(page, runtime, search, result)


def run_discovery(runtime: Runtime, page: Page) -> None:
    """Drive the query loop, checkpointing as it goes."""
    total = len(runtime.queries)

    while runtime.query_index < total and not runtime.stop_requested:
        search = runtime.queries[runtime.query_index]
        progress_line(runtime, search)

        try:
            process_query(page, runtime, search)
        except PlaywrightError as error:
            record_error(runtime, f"Query failed ({search.query})\nReason: {error}")

        runtime.query_index += 1
        runtime.queries_completed += 1

        if runtime.query_index % CHECKPOINT_EVERY_QUERIES == 0:
            checkpoint(runtime)

    if runtime.stop_requested:
        LOGGER.warning("Run stopped early at query %d.", runtime.query_index)


def install_shutdown_handler(runtime: Runtime) -> None:
    """Ask the run to stop at the next safe point when Ctrl+C is pressed."""

    def handle_interrupt(_signum: int, _frame: FrameType | None) -> None:
        if runtime.stop_requested:
            raise KeyboardInterrupt

        runtime.stop_requested = True
        LOGGER.warning("Shutdown requested. Finishing current step.")
        print("\nShutdown requested - finishing current step, press Ctrl+C again to abort.")

    signal.signal(signal.SIGINT, handle_interrupt)


def parse_arguments() -> argparse.Namespace:
    """Parse the optional command line flags. Running with no flags is normal."""
    parser = argparse.ArgumentParser(description="Discover ICP companies from public directories.")
    parser.add_argument("--fresh", action="store_true", help="ignore any existing autosave")
    parser.add_argument("--max-queries", type=int, default=None, help="limit how many queries run")
    parser.add_argument("--headless", action="store_true", help="run Chromium without a window")
    return parser.parse_args()


def main() -> int:
    """Single entry point: set up, discover, export, shut down."""
    args = parse_arguments()

    LOGGER.info("Discovery Engine v0.2 started.")
    runtime = initialize_runtime()

    if args.max_queries is not None:
        runtime.queries = runtime.queries[: max(args.max_queries, 0)]

    if args.fresh:
        remove_autosave()
    else:
        state = load_autosave()
        if state:
            restore_runtime(runtime, state)

    runtime.query_index = min(runtime.query_index, len(runtime.queries))
    runtime.started_at = time.time()

    install_shutdown_handler(runtime)

    LOGGER.info("Planned %d search queries.", len(runtime.queries))
    print(f"Planned {len(runtime.queries)} search queries.")

    try:
        playwright, browser, context, page = launch_browser(headless=args.headless or HEADLESS)
    except PlaywrightError as error:
        record_error(runtime, f"Could not start Chromium\nReason: {error}")
        print("Could not start Chromium. Run 'playwright install chromium' and try again.")
        return 2

    try:
        run_discovery(runtime, page)
    except KeyboardInterrupt:
        runtime.stop_requested = True
        LOGGER.warning("Run aborted by user.")
    finally:
        checkpoint(runtime)
        close_browser(playwright, browser, context)

    finalize_run(runtime)

    if not runtime.stop_requested:
        remove_autosave()

    return 1 if runtime.stop_requested else 0


if __name__ == "__main__":
    raise SystemExit(main())
