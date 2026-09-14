"""
Discovery Engine v0.3 - Prospecting Suite / Engine 1.

Discovers companies matching our ICP by searching public agency directories
through DuckDuckGo with a visible Chromium browser driven by Playwright.

Responsibility ends after discovery. This engine does not qualify, score,
enrich contacts, or use AI. It feeds the Company Intelligence Engine.

v0.3 is a data-quality revision of v0.2. The architecture is unchanged; the
precision of the company -> official website association is not:

    1. Website extraction now requires positive evidence that a link is the
       company's own site. A generic external link is no longer a candidate.
    2. Every candidate passes a dedicated identity check that compares the
       directory's company name against the site's domain, title, meta
       description and homepage text before the record can be exported.
    3. Records that cannot be confidently matched go to a Review Queue
       instead of being silently dropped or silently exported.
    4. A persistent registry stops the same companies being re-exported by
       every run over the same query universe.

Outputs (written to the directory containing this file):
    master_company_list.xlsx  - accepted companies, review queue, rejects
    prospector_input.csv      - exactly two columns: Company Name, Website
    run_report.xlsx           - run statistics and per-source diagnostics
    discovery.log             - detailed runtime log
    discovery_autosave.json   - checkpoint for resume support
    discovery_registry.json   - cross-run registry of discovered companies

Usage:
    python discovery_engine.py
    python discovery_engine.py --fresh                 # ignore the autosave
    python discovery_engine.py --fresh-universe        # ignore the registry
    python discovery_engine.py --max-queries 20

Requires: playwright (with chromium installed), beautifulsoup4, requests,
tldextract, openpyxl.
"""

from __future__ import annotations

import argparse
import csv
import difflib
import json
import logging
import os
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
REGISTRY_FILE = BASE_DIR / "discovery_registry.json"
OUTPUT_MASTER = BASE_DIR / "master_company_list.xlsx"
OUTPUT_PROSPECTOR = BASE_DIR / "prospector_input.csv"
OUTPUT_RUN_REPORT = BASE_DIR / "run_report.xlsx"

# Bumped for v0.3: the checkpoint now carries audit records, per-source
# diagnostics and the consecutive-block counters. Older autosaves are ignored
# by the existing version check rather than being misread.
AUTOSAVE_VERSION = 4
REGISTRY_VERSION = 1

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

# How much homepage body text is kept for the identity check. The identity
# check only needs to know whether the company name appears near the top of
# the page, so this stays small on purpose.
MAX_HOMEPAGE_TEXT_CHARS = 4_000

RETRY_ATTEMPTS = 3
RETRY_BACKOFF_SECONDS = 2.0
CHECKPOINT_EVERY_QUERIES = 5

# A directory that serves this many challenge pages in a row is left alone for
# the rest of the run rather than being hammered with more requests. Kept
# configurable from the command line with --max-consecutive-blocks.
MAX_CONSECUTIVE_BLOCKS = 3

# Audit sheets are capped so a long run cannot produce an unopenable workbook
# or an enormous checkpoint. The counters stay exact; only the stored rows
# are trimmed.
MAX_AUDIT_ROWS = 2_000

# Cross-run behaviour. "accumulate" keeps previously discovered companies out
# of prospector_input.csv; "fresh" deliberately rediscovers the universe.
REGISTRY_MODE_ACCUMULATE = "accumulate"
REGISTRY_MODE_FRESH = "fresh"
DEFAULT_REGISTRY_MODE = REGISTRY_MODE_ACCUMULATE

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

# Anchor text that names the official website outright. These are the only
# labels trusted on their own; "website" by itself is handled separately
# because it also appears in "Website Design", "Website Audit" and so on.
STRONG_WEBSITE_LINK_HINTS = (
    "visit website",
    "visit site",
    "visit the website",
    "view website",
    "go to website",
    "open website",
    "company website",
    "official website",
    "agency website",
    "visit profile website",
    "visit company website",
    "website:",
)

# Accepted only when it is essentially the whole label, not a word inside a
# sentence or a service name.
WEAK_WEBSITE_LINK_HINTS = ("website", "site", "web")
MAX_WEAK_LABEL_CHARS = 24

# Labels that contain a website hint but never point at the company's own
# site. Checked before the hints so "website design services" cannot qualify.
WEBSITE_LABEL_EXCLUSIONS = (
    "website design", "website development", "website redesign",
    "website audit", "website maintenance", "website builder",
    "website cost", "website pricing", "how much", "best website",
    "top website", "website company", "website companies",
    "website agencies", "website services", "website marketing",
)

# Path fragments used by directories to wrap outbound links in a redirect.
# A link that goes through the directory's own outbound redirect is strong
# evidence: directories only wrap the profile's own website that way.
REDIRECT_MARKERS = ("visit-website", "/go/", "/redirect", "/out/", "/link/", "/away")

# Query parameters that carry the real destination inside a redirect URL.
REDIRECT_PARAMS = ("u", "url", "redirect", "target", "link", "to", "r", "dest")

# File extensions that are assets rather than websites.
ASSET_SUFFIXES = (".jpg", ".jpeg", ".png", ".gif", ".svg", ".webp", ".pdf", ".zip")

# Page regions that hold OTHER companies: related agencies, sponsored cards,
# comparison strips and carousels. Every extraction layer below the
# directory's own selectors refuses to look inside these, because a
# neighbouring agency's "Visit Website" button is exactly how v0.2 attached
# the wrong website to the wrong company.
OTHER_COMPANY_REGION_HINTS = (
    "related", "similar", "recommend", "also-view", "also-viewed",
    "you-may", "sponsor", "advert", "promo", "carousel", "comparison",
    "competitor", "alternative", "other-agencies", "other-companies",
    "top-agencies", "more-agencies", "listing", "cross-sell", "upsell",
)

# Page furniture: navigation, footers, cookie bars and share widgets. These
# hold no company website either, but a profile's own website button can sit
# in a sidebar or a header block, so only the weakest layers - the ones that
# infer rather than read a label - refuse to look here.
#
# "header" is deliberately absent: the profile containers below are named
# .profile-header and .agency-header, and excluding the word would disable
# the very region the website link lives in.
CHROME_REGION_HINTS = (
    "sidebar", "side-bar", "widget", "breadcrumb", "cookie", "newsletter",
    "subscribe", "social", "share", "menu", "navigation", "footer",
)

EXCLUDED_ANCESTOR_TAGS = ("nav", "footer", "aside")

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

# Industry vocabulary that carries no identity. "Digital Media" matching
# "digitalmediaworks.com" proves nothing, so these words are ignored when
# deciding whether a name and a website belong together. They are still used
# for names built entirely from generic words, where an exact whole-name
# match is then required.
GENERIC_NAME_TOKENS = {
    "digital", "media", "marketing", "advertising", "ads", "adv",
    "creative", "creatives", "studio", "studios", "lab", "labs",
    "consultants", "consultancy", "communications", "communication",
    "brand", "brands", "branding", "performance", "growth", "global",
    "international", "india", "indian", "online", "web", "website",
    "internet", "interactive", "network", "networks", "partners",
    "partner", "ventures", "enterprise", "enterprises", "systems",
    "system", "works", "house", "hub", "world", "point",
    "infotech", "softech", "soft", "tech", "seo", "ppc", "smm", "social",
    "search", "content", "design", "designs", "development", "developers",
    "software", "best", "top", "leading", "experts", "expert",
    "professional", "professionals", "service", "management", "manage",
    "business", "corporate", "commerce", "ecommerce",
}

# Boilerplate removed from directory page titles to recover the company name.
NAME_NOISE_PATTERN = re.compile(
    r"\b(reviews?|profile|pricing|clutch\.co|clutch|goodfirms|designrush|"
    r"top\s+\d+|agency\s+profile)\b",
    re.IGNORECASE,
)

TITLE_PATTERN = re.compile(r"<title[^>]*>(.*?)</title>", re.IGNORECASE | re.DOTALL)

META_TAG_PATTERN = re.compile(r"<meta\s+([^>]+?)/?>", re.IGNORECASE)
META_ATTR_PATTERN = re.compile(
    r"""(\w[\w:-]*)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+))""",
    re.IGNORECASE,
)

SCRIPT_STYLE_PATTERN = re.compile(
    r"<(script|style|noscript|svg)\b[^>]*>.*?</\1>",
    re.IGNORECASE | re.DOTALL,
)
HTML_TAG_PATTERN = re.compile(r"<[^>]+>")

# Anchor text that is itself a domain or URL. When a profile shows the
# website as its own link text and that text matches the href, the link is
# almost certainly the company's site.
DOMAIN_TEXT_PATTERN = re.compile(
    r"^\s*(?:https?://)?(?:www\.)?([a-z0-9][a-z0-9.-]*\.[a-z]{2,})(?:/[^\s]*)?\s*$",
    re.IGNORECASE,
)

CONTROL_CHARS_PATTERN = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

PRIVATE_HOST_PATTERN = re.compile(
    r"^(localhost|127\.|0\.0\.0\.0|10\.|192\.168\.|172\.(1[6-9]|2\d|3[01])\.)",
    re.IGNORECASE,
)

IP_HOST_PATTERN = re.compile(r"^\d{1,3}(\.\d{1,3}){3}$")

# Use the bundled public suffix snapshot so startup never depends on the network.
DOMAIN_EXTRACTOR = tldextract.TLDExtract(suffix_list_urls=())


# ----------------------------------------------------------------------
# Identity scoring
# ----------------------------------------------------------------------
#
# The identity score answers one question: does this website belong to the
# company the directory profile named? It is a plain additive score so any
# row in the audit output can be explained from its listed signals.
#
# Points are deliberately arranged so that a single strong signal - the full
# company name in the site title, or the name in the domain - is enough to
# accept, while a reachable site with metadata but no name evidence anywhere
# cannot be. That is the "Softex Media -> salesflint.com" case from QA.

IDENTITY_POINTS = {
    # The name, or the whole significant part of it, is in the domain.
    "domain-exact": 45,
    # A distinctive word from the name is in the domain.
    "domain-token": 30,
    # The name and the domain are near-identical allowing for typos/spacing.
    "domain-fuzzy": 22,
    # The domain is the name's initials (e.g. "Sharma Kapoor Rao" -> skr.com).
    "domain-acronym": 25,
    # The full name appears in the site title.
    "title-strong": 45,
    # A distinctive word from the name appears in the site title.
    "title-token": 22,
    # The title's brand segment is a near match for the name.
    "title-fuzzy": 20,
    # The name appears in the meta description.
    "meta-name": 12,
    # The name appears in the homepage body text.
    "homepage-name": 10,
    # The directory published this link as the profile's own website.
    "explicit-website-link": 18,
    # The profile showed the website address as the link text.
    "link-text-domain": 12,
    # The link came from a weaker but still profile-scoped extraction layer.
    "profile-scoped-link": 8,
}

# Applied only when nothing at all connects the name to the site.
IDENTITY_PENALTIES = {
    # The site clearly belongs to a differently named company.
    "identity-mismatch": -35,
    # The site answered and had a title, but the name is nowhere on it.
    "no-name-evidence": -25,
}

IDENTITY_ACCEPT_SCORE = 55
IDENTITY_REVIEW_SCORE = 30

# When the site never answered, an absent name is missing evidence rather
# than contradicting evidence, so the review floor is lower. Such records
# still cannot reach High confidence.
IDENTITY_REVIEW_SCORE_UNVERIFIED = 15

# The identity score a record needs before confidence may reach High.
IDENTITY_STRONG_SCORE = 70

# Minimum similarity ratios for the two fuzzy signals.
DOMAIN_FUZZY_RATIO = 0.72
TITLE_FUZZY_RATIO = 0.75

# A token must be at least this long before a bare substring test against a
# domain is meaningful. Shorter tokens must sit at one end of the domain.
MIN_SUBSTRING_TOKEN_CHARS = 4

VERDICT_ACCEPT = "ACCEPT"
VERDICT_REVIEW = "REVIEW"
VERDICT_REJECT = "REJECT"

STATUS_ACCEPTED = "ACCEPTED"
STATUS_REVIEW = "REVIEW"
STATUS_REJECTED = "REJECTED"

# Extraction methods that count as the directory explicitly publishing the
# company's website, rather than the engine inferring it.
STRONG_WEBSITE_METHODS = ("selector", "redirect", "label", "link-text-domain")

# Extraction methods that are scoped to the profile but inferred.
SCOPED_WEBSITE_METHODS = ("profile-container", "name-affinity")

# Confidence is scored from evidence. Each item is worth the points shown.
EVIDENCE_POINTS = {
    "explicit-website-link": 2,
    "website-reachable": 2,
    "identity-strong": 3,
    "identity-moderate": 1,
    "meta-description": 1,
    "internal-pages": 1,
}
CONFIDENCE_HIGH_SCORE = 7
CONFIDENCE_MEDIUM_SCORE = 5


# ----------------------------------------------------------------------
# Output schemas
# ----------------------------------------------------------------------

MASTER_COLUMNS = [
    "Company Name",
    "Website",
    "Source",
    "Source URL",
    "Search Query",
    "Status",
    "Discovery Confidence",
    "Confidence Score",
    "Evidence",
    "Identity Score",
    "Identity Verdict",
    "Identity Signals",
    "Meta Description",
    "Internal Pages",
    "Keyword",
    "City",
    "Root Domain",
    "Verified",
    "Website Source",
    "HTTP Status",
    "Site Title",
    "Previously Discovered",
    "Exported",
    "Discovered At",
]

AUDIT_COLUMNS = [
    "Company Name",
    "Candidate Website",
    "Source",
    "Source URL",
    "Identity Score",
    "Identity Verdict",
    "Confidence",
    "Website Title",
    "Meta Description",
    "Reason Flagged",
    "Identity Signals",
    "Website Source",
    "HTTP Status",
    "Verified",
    "Search Query",
    "Discovered At",
]

KNOWN_COLUMNS = [
    "Company Name",
    "Website",
    "Root Domain",
    "Source",
    "Source URL",
    "First Discovered",
    "Times Seen",
    "Search Query",
    "Seen At",
]

# The frozen handoff contract with the Company Intelligence Engine.
PROSPECTOR_COLUMNS = ["Company Name", "Website"]


# ----------------------------------------------------------------------
# Directory adapters
# ----------------------------------------------------------------------

@dataclass(frozen=True)
class DirectoryAdapter:
    """
    Everything that differs between the three directories.

    The extraction layers are shared; only the selectors differ. Every
    selector list is ordered most-specific first, because the first layer to
    produce a plausible URL wins.

    website_selectors     markup known to hold the profile's website link
    website_containers    regions of the page that belong to this profile,
                          used to scope the weaker fallback layers
    name_selectors        markup known to hold the company name
    """

    name: str
    search_operator: str
    domain: str
    profile_markers: tuple[str, ...]
    website_selectors: tuple[str, ...]
    name_selectors: tuple[str, ...]
    website_containers: tuple[str, ...] = ()


# Selector note: these lists are written defensively and are layered, so a
# directory that changes its markup degrades to the next layer instead of
# failing outright or, worse, attaching an unrelated link. The markup of
# Clutch and GoodFirms could not be inspected live from the build
# environment, so their selectors are informed guesses covering several
# plausible shapes - they are NOT live-verified. When every layer misses, the
# adapter records the exact layer that failed in the run report rather than
# falling back to a generic external link.

DIRECTORIES: dict[str, DirectoryAdapter] = {
    "Clutch": DirectoryAdapter(
        name="Clutch",
        search_operator="site:clutch.co",
        domain="clutch.co",
        profile_markers=("/profile/",),
        website_selectors=(
            "a[data-link_text='Visit website']",
            "a[data-testid='visit-website']",
            "a[data-testid='provider-website']",
            "a.website-link__item",
            ".website-link a[href]",
            "li.website-link a[href]",
            "a.sg-button--website",
            "a.provider__cta-link--website",
            ".profile-summary__website a[href]",
        ),
        name_selectors=(
            "h1.profile-header__title",
            "h1[data-testid='provider-name']",
            "h1.provider__title",
            "h1[itemprop='name']",
            "h1",
        ),
        website_containers=(
            ".profile-header",
            ".profile-summary",
            ".provider__header",
            "#profile-summary",
            "[data-testid='profile-header']",
        ),
    ),
    "GoodFirms": DirectoryAdapter(
        name="GoodFirms",
        search_operator="site:goodfirms.co",
        domain="goodfirms.co",
        profile_markers=("/company/",),
        website_selectors=(
            "a[href*='/visit-website/']",
            "a[href*='/visit-website']",
            "a.visit-website",
            "a.visit-website-link",
            "a.visit_website",
            ".profile-visit-website a[href]",
            ".company-visit-website a[href]",
            ".detail-visit-website a[href]",
            ".profile-detail-website a[href]",
            "a[title*='Visit Website' i]",
            "a[aria-label*='Visit Website' i]",
            "a[data-original-title*='Visit Website' i]",
            "a.entity-website",
            "a[itemprop='url'][href^='http']",
        ),
        name_selectors=(
            "h1.profile-name",
            "h1.entity-name",
            "h1.company-name",
            ".profile-header h1",
            "h1[itemprop='name']",
            "h1",
        ),
        website_containers=(
            ".profile-header",
            ".company-profile-header",
            ".entity-header",
            ".profile-detail",
            "[itemtype*='Organization']",
        ),
    ),
    "DesignRush": DirectoryAdapter(
        name="DesignRush",
        search_operator="site:designrush.com",
        domain="designrush.com",
        profile_markers=("/agency/profile/",),
        # The v0.2 list ended with a[href][rel*='nofollow'][target='_blank'],
        # which matched every outbound link on the page - related agencies,
        # sponsored placements and partner badges included. That single
        # selector produced the wrong-company pairings QA reported, so it is
        # gone. The replacements below are all website-specific.
        website_selectors=(
            "a[data-testid='agency-website']",
            "a[data-testid='visit-website']",
            "a.agency-website",
            "a.visit-website",
            "a[href*='/visit-website']",
            ".agency-profile__website a[href]",
            ".agency-header__website a[href]",
            ".profile-website a[href]",
            "a[data-gtm*='website' i]",
            "a[class*='website' i][href^='http']",
        ),
        name_selectors=(
            "h1.agency-name",
            ".agency-profile__title h1",
            "h1[itemprop='name']",
            "h1",
        ),
        website_containers=(
            ".agency-profile__header",
            ".agency-header",
            ".profile-header",
            "[data-testid='agency-header']",
            "[itemtype*='Organization']",
        ),
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
REASON_NO_TRUSTED_LINK = "no link on the page identifies the company website"
REASON_KNOWN_COMPANY = "already in the discovery registry"


# ----------------------------------------------------------------------
# Per-source diagnostics
# ----------------------------------------------------------------------
#
# Attempt-level counters, kept per directory. These describe what happened
# while scraping and are deliberately separate from the final-record
# statistics, which are always derived from runtime.records.

EVENT_PROFILES_ATTEMPTED = "profiles attempted"
EVENT_PROFILES_LOADED = "profiles loaded"
EVENT_SECURITY_CHALLENGES = "security challenge pages"
EVENT_TIMEOUTS = "timeouts"
EVENT_NAVIGATION_ERRORS = "navigation errors"
EVENT_UNREADABLE = "unreadable pages"
EVENT_REDIRECTED = "redirected away"
EVENT_CANDIDATES_FOUND = "candidate websites found"
EVENT_SELECTOR_FAILURES = "selector failures"
EVENT_WEBSITE_UNREACHABLE = "candidate websites unreachable"
EVENT_VALIDATION_REJECTED = "rejected by validation"
EVENT_IDENTITY_REJECTED = "identity mismatches"
EVENT_IDENTITY_REVIEW = "sent to review queue"
EVENT_RUN_DUPLICATES = "duplicates within run"
EVENT_REGISTRY_DUPLICATES = "known from earlier runs"
EVENT_ACCEPTED = "accepted companies"
EVENT_EXPORTED = "exported companies"

SOURCE_EVENTS = (
    EVENT_PROFILES_ATTEMPTED,
    EVENT_PROFILES_LOADED,
    EVENT_SECURITY_CHALLENGES,
    EVENT_TIMEOUTS,
    EVENT_NAVIGATION_ERRORS,
    EVENT_UNREADABLE,
    EVENT_REDIRECTED,
    EVENT_SELECTOR_FAILURES,
    EVENT_CANDIDATES_FOUND,
    EVENT_WEBSITE_UNREACHABLE,
    EVENT_VALIDATION_REJECTED,
    EVENT_IDENTITY_REJECTED,
    EVENT_IDENTITY_REVIEW,
    EVENT_RUN_DUPLICATES,
    EVENT_REGISTRY_DUPLICATES,
    EVENT_ACCEPTED,
    EVENT_EXPORTED,
)


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
    link_text: str = ""


@dataclass
class SiteCheck:
    """
    The outcome of fetching a candidate company website.

    v0.3 keeps the meta description and a slice of the homepage text from the
    same response the reachability check already downloaded, so the identity
    check has something to compare the company name against without a second
    request.
    """

    reachable: bool = False
    status_code: int = 0
    final_url: str = ""
    title: str = ""
    meta_description: str = ""
    text_sample: str = ""
    blocked: bool = False
    reason: str = ""


@dataclass
class IdentityAssessment:
    """Whether a candidate website belongs to the company that was named."""

    score: int = 0
    verdict: str = VERDICT_REJECT
    signals: list[str] = field(default_factory=list)
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
    website_link_text: str = ""
    http_status: int = 0
    site_title: str = ""
    evidence: list[str] = field(default_factory=list)
    confidence_score: int = 0
    identity_score: int = 0
    identity_verdict: str = ""
    identity_signals: list[str] = field(default_factory=list)
    identity_reason: str = ""
    status: str = ""
    rejection_reason: str = ""
    previously_discovered: bool = False
    exported: bool = False
    discovered_at: str = ""


@dataclass
class KnownCompany:
    """A company the registry had already seen, met again during this run."""

    company_name: str = ""
    website: str = ""
    root_domain: str = ""
    source: str = ""
    source_url: str = ""
    first_discovered_at: str = ""
    times_seen: int = 0
    search_query: str = ""
    seen_at: str = ""


# ----------------------------------------------------------------------
# Persistent cross-run registry
# ----------------------------------------------------------------------

class DiscoveryRegistry:
    """
    A JSON file remembering every company earlier runs exported.

    Deduplication is keyed on the normalized root domain. The normalized
    company name is kept as a secondary key but is never used on its own:
    two unrelated agencies can share a name, so a name match only counts as
    a duplicate when it comes from the same directory as the earlier one.

    Nothing here needs a database. The file is small - one entry per exported
    company - and is rewritten atomically so an interrupted run cannot leave
    a half-written registry behind.
    """

    def __init__(self, path: Path = REGISTRY_FILE, mode: str = DEFAULT_REGISTRY_MODE) -> None:
        self.path = path
        self.mode = mode
        self.entries: dict[str, dict] = {}
        self.loaded_count = 0
        self.added_count = 0
        self.updated_count = 0

    @property
    def enforcing(self) -> bool:
        """True when known companies should be held back from the export."""
        return self.mode == REGISTRY_MODE_ACCUMULATE

    def load(self) -> None:
        """Read the registry, tolerating a missing or damaged file."""
        if not self.path.exists():
            LOGGER.info("No discovery registry at %s. Starting a new one.", self.path.name)
            return

        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError) as error:
            LOGGER.warning("Ignoring unreadable discovery registry: %s", error)
            return

        if not isinstance(payload, dict) or payload.get("version") != REGISTRY_VERSION:
            LOGGER.warning("Ignoring discovery registry from an incompatible version.")
            return

        companies = payload.get("companies", {})
        if not isinstance(companies, dict):
            LOGGER.warning("Ignoring discovery registry with an unexpected shape.")
            return

        for domain, entry in companies.items():
            if isinstance(domain, str) and domain and isinstance(entry, dict):
                self.entries[domain] = entry

        self.loaded_count = len(self.entries)
        LOGGER.info(
            "Discovery registry loaded: %d companies from earlier runs (mode=%s).",
            self.loaded_count, self.mode,
        )

    def lookup(self, domain: str) -> dict | None:
        """Return the stored entry for a root domain, if there is one."""
        return self.entries.get(domain) if domain else None

    def find_by_name(self, normalized: str, source: str) -> dict | None:
        """
        Secondary lookup: the same normalized name from the same directory.

        Name alone is never enough - this pairs it with the source so two
        different agencies that happen to share a name are not merged.
        """
        if not normalized:
            return None

        for entry in self.entries.values():
            if entry.get("normalized_name") == normalized and entry.get("source") == source:
                return entry

        return None

    def knows(self, domain: str, normalized_name: str, source: str) -> dict | None:
        """Return the matching entry when this company has been seen before."""
        return self.lookup(domain) or self.find_by_name(normalized_name, source)

    def remember(self, record: CompanyRecord, timestamp: str) -> None:
        """Record a newly exported company, or refresh one already stored."""
        domain = record.root_domain
        if not domain:
            return

        existing = self.entries.get(domain)

        if existing is None:
            self.entries[domain] = {
                "domain": domain,
                "company_name": record.company_name,
                "normalized_name": normalized_company_name(record.company_name),
                "source": record.source,
                "source_url": record.source_url,
                "website": record.website,
                "first_discovered_at": record.discovered_at or timestamp,
                "last_seen_at": timestamp,
                "times_seen": 1,
            }
            self.added_count += 1
            return

        existing["last_seen_at"] = timestamp
        existing["times_seen"] = int(existing.get("times_seen", 1)) + 1
        existing.setdefault("company_name", record.company_name)
        existing.setdefault("normalized_name", normalized_company_name(record.company_name))
        existing.setdefault("source", record.source)
        existing.setdefault("source_url", record.source_url)
        existing.setdefault("website", record.website)
        existing.setdefault("first_discovered_at", record.discovered_at or timestamp)
        self.updated_count += 1

    def touch(self, entry: dict, timestamp: str) -> None:
        """Note that a already-known company was met again this run."""
        entry["last_seen_at"] = timestamp
        entry["times_seen"] = int(entry.get("times_seen", 1)) + 1
        self.updated_count += 1

    def save(self) -> None:
        """Write the registry atomically so a crash cannot corrupt it."""
        payload = {
            "version": REGISTRY_VERSION,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "count": len(self.entries),
            "companies": self.entries,
        }

        temporary = self.path.with_suffix(self.path.suffix + ".tmp")

        try:
            temporary.write_text(json.dumps(payload, indent=2), encoding="utf-8")
            os.replace(temporary, self.path)
        except OSError as error:
            LOGGER.error("Could not write the discovery registry: %s", error)
            return

        LOGGER.info(
            "Discovery registry saved: %d companies (%d new, %d refreshed).",
            len(self.entries), self.added_count, self.updated_count,
        )


# ----------------------------------------------------------------------
# Runtime
# ----------------------------------------------------------------------

@dataclass
class Runtime:
    """Single source of truth for the run: work list, results and counters."""

    queries: list[SearchQuery]
    records: list[CompanyRecord] = field(default_factory=list)
    review_records: list[CompanyRecord] = field(default_factory=list)
    rejected_records: list[CompanyRecord] = field(default_factory=list)
    known_companies: list[KnownCompany] = field(default_factory=list)
    seen_domains: set[str] = field(default_factory=set)
    visited_urls: set[str] = field(default_factory=set)
    error_messages: list[str] = field(default_factory=list)
    skip_reasons: Counter[str] = field(default_factory=Counter)
    blocked_directories: set[str] = field(default_factory=set)
    consecutive_blocks: Counter[str] = field(default_factory=Counter)
    source_stats: dict[str, Counter[str]] = field(default_factory=dict)
    registry: DiscoveryRegistry = field(default_factory=DiscoveryRegistry)
    max_consecutive_blocks: int = MAX_CONSECUTIVE_BLOCKS
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
    records_review: int = 0
    duplicates_removed: int = 0
    registry_duplicates: int = 0
    errors: int = 0

    @property
    def companies_found(self) -> int:
        """Companies kept after validation, identity checking and dedup."""
        return len(self.records)

    @property
    def companies_exported(self) -> int:
        """Accepted companies actually written to prospector_input.csv."""
        return sum(1 for record in self.records if record.exported)

    @property
    def websites_verified(self) -> int:
        """Exported companies whose website responded. Never exceeds records."""
        return sum(1 for record in self.records if record.verified)


def initialize_runtime(
    registry_path: Path = REGISTRY_FILE,
    registry_mode: str = DEFAULT_REGISTRY_MODE,
    max_consecutive_blocks: int = MAX_CONSECUTIVE_BLOCKS,
) -> Runtime:
    """Create the shared runtime state used by the main loop."""
    return Runtime(
        queries=generate_queries(),
        registry=DiscoveryRegistry(path=registry_path, mode=registry_mode),
        max_consecutive_blocks=max_consecutive_blocks,
    )


def record_error(runtime: Runtime, message: str) -> None:
    """Log an error once and count it in the single statistics system."""
    runtime.errors += 1
    runtime.error_messages.append(f"{datetime.now().isoformat(timespec='seconds')} | {message}")
    LOGGER.error(message)


def record_skip(runtime: Runtime, url: str, reason: str) -> None:
    """Log why a candidate was dropped, with the reason kept for the report."""
    runtime.skip_reasons[reason] += 1
    LOGGER.info("Skipped %s\nReason: %s", url, reason)


def note_source_event(runtime: Runtime, source: str, event: str, amount: int = 1) -> None:
    """Count one attempt-level event against the directory that caused it."""
    runtime.source_stats.setdefault(source, Counter())[event] += amount


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


def domain_slug(url: str) -> str:
    """Return the brandable part of a domain, stripped to letters and digits."""
    return re.sub(r"[^a-z0-9]", "", root_domain(url).split(".")[0])


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


def resolve_candidate(href: str, source_url: str) -> str:
    """Turn a raw href into an absolute URL, unwrapping directory redirects."""
    absolute = urljoin(source_url, href.strip())
    return unwrap_redirect(absolute) if looks_like_redirect(absolute) else absolute


# ----------------------------------------------------------------------
# Name helpers
# ----------------------------------------------------------------------

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


def distinctive_name_tokens(name: str) -> list[str]:
    """
    The tokens that actually identify a company.

    "Digital Media" tells us nothing about which agency we are looking at, so
    industry vocabulary is removed before any name-to-website comparison. A
    name built entirely from such words keeps no distinctive tokens, and is
    handled separately by requiring a whole-name match instead.
    """
    return [token for token in name_tokens(name) if token not in GENERIC_NAME_TOKENS]


def normalized_company_name(name: str) -> str:
    """A stable key for a company name, used as the registry's secondary key."""
    return "".join(name_tokens(name))


def normalize_text(text: str) -> str:
    """Collapse text to lowercase letters and digits for substring matching."""
    return re.sub(r"[^a-z0-9]", "", text.lower())


def text_words(text: str) -> set[str]:
    """The set of lowercase words in a piece of text."""
    return set(re.findall(r"[a-z0-9]+", text.lower()))


def similarity(left: str, right: str) -> float:
    """Deterministic 0-1 similarity between two normalized strings."""
    if not left or not right:
        return 0.0
    return difflib.SequenceMatcher(None, left, right).ratio()


def token_in_domain(token: str, slug: str) -> bool:
    """
    True when a name token is present in a domain slug.

    Short tokens are only accepted at either end of the slug, because a
    three-letter sequence appears inside plenty of unrelated words.
    """
    if not token or not slug:
        return False
    if len(token) >= MIN_SUBSTRING_TOKEN_CHARS:
        return token in slug
    return slug.startswith(token) or slug.endswith(token)


def title_brand_segment(title: str) -> str:
    """
    The part of a page title that usually carries the company's own name.

    Titles are normally "Brand | Tagline" or "Tagline - Brand", so both ends
    are considered and the shorter one is treated as the brand.
    """
    parts = [part.strip() for part in re.split(r"\s*[|\-–—:·•]\s*", title) if part.strip()]
    if not parts:
        return ""
    if len(parts) == 1:
        return parts[0]
    return min(parts[0], parts[-1], key=len)


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


def _region_identifiers(element) -> str:
    """The class, id and section names of one element, lowercased."""
    if not hasattr(element, "get"):
        return ""

    classes = element.get("class", []) or []
    if isinstance(classes, str):
        classes = [classes]

    return " ".join(
        [
            " ".join(str(item) for item in classes),
            str(element.get("id", "") or ""),
            str(element.get("data-section", "") or ""),
        ]
    ).lower()


def _in_excluded_region(anchor, strict: bool = False) -> bool:
    """
    True when a link sits in a part of the page that belongs to someone else.

    Directory profiles surround the company with "related agencies", sponsored
    cards and comparison carousels. Links from those regions were the direct
    cause of the wrong company -> wrong website pairings QA found, so every
    layer below the directory's own selectors refuses to look inside them.

    strict also rules out page furniture - sidebars, footers, share widgets.
    Only the inferring layers set it: a link labelled "Visit Website" says
    what it is wherever it sits, but an unlabelled external link in a footer
    is just a footer link.
    """
    hints = OTHER_COMPANY_REGION_HINTS + (CHROME_REGION_HINTS if strict else ())

    for parent in anchor.parents:
        if strict and (getattr(parent, "name", "") or "") in EXCLUDED_ANCESTOR_TAGS:
            return True

        identifiers = _region_identifiers(parent)
        if identifiers and any(hint in identifiers for hint in hints):
            return True

    return False


def _anchor_label(anchor) -> str:
    """The visible text plus accessible labels of a link, lowercased."""
    return " ".join(
        part.strip().lower()
        for part in (
            anchor.get_text(" ", strip=True),
            anchor.get("aria-label", "") or "",
            anchor.get("title", "") or "",
            anchor.get("data-original-title", "") or "",
        )
        if part and part.strip()
    )


def _label_names_website(label: str) -> bool:
    """
    True when a link's label says it leads to the company's own website.

    "Website Design Services" and "Best Website Companies" contain the word
    but point somewhere else entirely, so exclusions are checked first and a
    bare "Website" only counts when it is essentially the whole label.
    """
    if not label:
        return False

    if any(exclusion in label for exclusion in WEBSITE_LABEL_EXCLUSIONS):
        return False

    if any(hint in label for hint in STRONG_WEBSITE_LINK_HINTS):
        return True

    if len(label) <= MAX_WEAK_LABEL_CHARS:
        stripped = label.strip(" :-–—>»|")
        return stripped in WEAK_WEBSITE_LINK_HINTS

    return False


def _anchor_text_domain(anchor) -> str:
    """The domain a link displays as its own text, if that is what it shows."""
    text = anchor.get_text(" ", strip=True)
    if not text or len(text) > 120:
        return ""

    found = DOMAIN_TEXT_PATTERN.match(text)
    return found.group(1).lower() if found else ""


def extract_official_website(
    soup: BeautifulSoup,
    adapter: DirectoryAdapter,
    source_url: str,
    company_name: str = "",
) -> WebsiteMatch:
    """
    Find the company's own website on a directory profile page.

    v0.2 ended with "best remaining external link", which is how profiles
    handed back the website of a neighbouring agency. v0.3 requires positive
    evidence at every layer and returns nothing when it has none:

        1. the directory's own website selectors
        2. the directory's outbound redirect wrapper
        3. an anchor labelled "Visit Website"
        4. an anchor that displays the website address as its text
        5. a single external link inside the profile's own header block
        6. an external link whose domain carries the company's name

    A link that is merely external, nofollow or target=_blank is not
    evidence and is never used on its own.
    """
    layers = (
        lambda: _match_by_selector(soup, adapter, source_url),
        lambda: _match_by_redirect_wrapper(soup, adapter, source_url),
        lambda: _match_by_label(soup, source_url),
        lambda: _match_by_domain_anchor(soup, source_url),
        lambda: _match_by_profile_container(soup, adapter, source_url, company_name),
        lambda: _match_by_name_affinity(soup, source_url, company_name),
    )

    for layer in layers:
        match = layer()
        if match.url:
            return match

    return WebsiteMatch(reason=REASON_NO_TRUSTED_LINK)


def _match_by_selector(
    soup: BeautifulSoup,
    adapter: DirectoryAdapter,
    source_url: str,
) -> WebsiteMatch:
    """
    Layer 1: the selectors known to hold the website link on this directory.

    These are website-specific markup hooks, so they are trusted wherever
    they appear on the page and are not filtered by page region.
    """
    for selector in adapter.website_selectors:
        try:
            anchors = soup.select(selector)
        except Exception as error:  # noqa: BLE001 - a bad selector must not stop the run
            LOGGER.debug("Selector %r rejected on %s: %s", selector, adapter.name, error)
            continue

        for anchor in anchors:
            href = anchor.get("href", "")
            if not href:
                continue

            candidate = resolve_candidate(href, source_url)
            if is_plausible_website(candidate):
                return WebsiteMatch(
                    url=candidate,
                    method="selector",
                    link_text=anchor.get_text(" ", strip=True)[:120],
                )

    return WebsiteMatch(reason=REASON_SELECTOR_MISSING)


def _match_by_redirect_wrapper(
    soup: BeautifulSoup,
    adapter: DirectoryAdapter,
    source_url: str,
) -> WebsiteMatch:
    """
    Layer 2: links routed through the directory's own outbound redirect.

    Directories only wrap a profile's own website that way, so the wrapper
    itself is the evidence. Related-agency cards are still excluded, because
    some directories wrap those links too.
    """
    for anchor in soup.find_all("a", href=True):
        href = anchor["href"].strip()
        if not href or not looks_like_redirect(href):
            continue
        if _in_excluded_region(anchor):
            continue

        candidate = resolve_candidate(href, source_url)

        # An unwrapped redirect that still points at the directory tells us
        # nothing; only a real outbound destination counts.
        if root_domain(candidate) == adapter.domain:
            continue

        if is_plausible_website(candidate):
            return WebsiteMatch(
                url=candidate,
                method="redirect",
                link_text=anchor.get_text(" ", strip=True)[:120],
            )

    return WebsiteMatch(reason=REASON_SELECTOR_MISSING)


def _match_by_label(soup: BeautifulSoup, source_url: str) -> WebsiteMatch:
    """Layer 3: an anchor whose visible label says it leads to the website."""
    for anchor in soup.find_all("a", href=True):
        if not _label_names_website(_anchor_label(anchor)):
            continue
        if _in_excluded_region(anchor):
            continue

        candidate = resolve_candidate(anchor["href"], source_url)
        if is_plausible_website(candidate):
            return WebsiteMatch(
                url=candidate,
                method="label",
                link_text=anchor.get_text(" ", strip=True)[:120],
            )

    return WebsiteMatch(reason=REASON_SELECTOR_MISSING)


def _match_by_domain_anchor(soup: BeautifulSoup, source_url: str) -> WebsiteMatch:
    """
    Layer 4: an anchor that displays the website address as its own text.

    Profiles that print "www.example.com" and link it to example.com are
    stating the website outright. The displayed domain must agree with the
    href, so a link showing one company and pointing at another is ignored.
    """
    for anchor in soup.find_all("a", href=True):
        shown = _anchor_text_domain(anchor)
        if not shown:
            continue
        if _in_excluded_region(anchor, strict=True):
            continue

        candidate = resolve_candidate(anchor["href"], source_url)
        if not is_plausible_website(candidate):
            continue

        if root_domain(candidate) != root_domain(f"http://{shown}"):
            LOGGER.debug("Link text %s disagrees with href %s. Ignored.", shown, candidate)
            continue

        return WebsiteMatch(
            url=candidate,
            method="link-text-domain",
            link_text=anchor.get_text(" ", strip=True)[:120],
        )

    return WebsiteMatch(reason=REASON_SELECTOR_MISSING)


def _external_links_in(element, source_url: str, strict: bool = True) -> dict[str, str]:
    """Distinct external company domains linked from one region of a page."""
    found: dict[str, str] = {}

    for anchor in element.find_all("a", href=True):
        href = anchor["href"].strip()
        if not href or href.startswith(("#", "mailto:", "tel:", "javascript:")):
            continue
        if _in_excluded_region(anchor, strict=strict):
            continue

        candidate = resolve_candidate(href, source_url)
        if not is_plausible_website(candidate):
            continue

        found.setdefault(root_domain(candidate), candidate)

    return found


def _match_by_profile_container(
    soup: BeautifulSoup,
    adapter: DirectoryAdapter,
    source_url: str,
    company_name: str,
) -> WebsiteMatch:
    """
    Layer 5: an external link inside the profile's own header block.

    The container selectors scope this to markup that belongs to the company
    being viewed. When a container offers several different external domains
    the page is ambiguous, so the layer takes the one that carries the
    company's name and otherwise declines to guess.
    """
    for selector in adapter.website_containers:
        try:
            container = soup.select_one(selector)
        except Exception as error:  # noqa: BLE001 - a bad selector must not stop the run
            LOGGER.debug("Container %r rejected on %s: %s", selector, adapter.name, error)
            continue

        if container is None:
            continue

        candidates = _external_links_in(container, source_url, strict=False)
        if not candidates:
            continue

        if len(candidates) == 1:
            url = next(iter(candidates.values()))
            return WebsiteMatch(url=url, method="profile-container")

        named = _domain_matching_name(candidates, company_name)
        if named:
            return WebsiteMatch(url=named, method="profile-container")

        LOGGER.debug(
            "Profile container %r on %s offered %d external domains. Declined to guess.",
            selector, source_url, len(candidates),
        )

    return WebsiteMatch(reason=REASON_SELECTOR_MISSING)


def _match_by_name_affinity(
    soup: BeautifulSoup,
    source_url: str,
    company_name: str,
) -> WebsiteMatch:
    """
    Layer 6: an external link whose domain carries the company's own name.

    This is the last layer and the only one that looks at the whole page. It
    is safe there because the evidence is the name itself: a link is only
    taken when its domain already agrees with the company being profiled.
    """
    if not distinctive_name_tokens(company_name):
        return WebsiteMatch(reason=REASON_NO_TRUSTED_LINK)

    candidates = _external_links_in(soup, source_url)
    if not candidates:
        return WebsiteMatch(reason=REASON_NO_CANDIDATE_LINKS)

    named = _domain_matching_name(candidates, company_name)
    if named:
        return WebsiteMatch(url=named, method="name-affinity")

    return WebsiteMatch(reason=REASON_NO_TRUSTED_LINK)


def _domain_matching_name(candidates: dict[str, str], company_name: str) -> str:
    """Pick the candidate domain that carries the company's distinctive name."""
    tokens = distinctive_name_tokens(company_name)
    if not tokens:
        return ""

    joined = "".join(tokens)

    for domain, url in candidates.items():
        slug = re.sub(r"[^a-z0-9]", "", domain.split(".")[0])
        if not slug:
            continue
        if joined and (joined == slug or joined in slug):
            return url
        if any(token_in_domain(token, slug) for token in tokens):
            return url

    return ""


def open_profile(
    page: Page,
    adapter: DirectoryAdapter,
    result: SearchResult,
    runtime: Runtime,
) -> BeautifulSoup | None:
    """
    Open a directory profile page and return its parsed markup.

    Every way this can fail is counted against the directory that caused it,
    so the run report can tell a blocked source from a slow one from a source
    whose markup has moved.
    """
    runtime.visited_urls.add(result.url)

    failure = navigate(page, result.url, runtime)
    if failure:
        event = EVENT_TIMEOUTS if failure == REASON_TIMEOUT else EVENT_NAVIGATION_ERRORS
        note_source_event(runtime, adapter.name, event)
        record_skip(runtime, result.url, failure)
        return None

    html = read_page_html(page, runtime)
    if not html:
        note_source_event(runtime, adapter.name, EVENT_UNREADABLE)
        record_skip(runtime, result.url, REASON_UNREADABLE)
        return None

    if looks_blocked(html):
        runtime.profiles_blocked += 1
        note_source_event(runtime, adapter.name, EVENT_SECURITY_CHALLENGES)
        note_directory_block(runtime, adapter)
        record_skip(runtime, result.url, REASON_SECURITY_VERIFICATION)
        return None

    if root_domain(page.url) != adapter.domain:
        note_source_event(runtime, adapter.name, EVENT_REDIRECTED)
        record_skip(runtime, result.url, f"{REASON_REDIRECTED} to {page.url}")
        return None

    runtime.consecutive_blocks[adapter.name] = 0
    runtime.profiles_visited += 1
    note_source_event(runtime, adapter.name, EVENT_PROFILES_LOADED)
    return BeautifulSoup(html, "html.parser")


def note_directory_block(runtime: Runtime, adapter: DirectoryAdapter) -> None:
    """
    Count a challenge page and retire the directory once they keep coming.

    One transient challenge is not enough; the directory is only dropped
    after max_consecutive_blocks in a row, and the decision is carried in the
    checkpoint so a resumed run does not start hammering it again.
    """
    runtime.consecutive_blocks[adapter.name] += 1
    seen = runtime.consecutive_blocks[adapter.name]

    if seen < runtime.max_consecutive_blocks:
        LOGGER.info(
            "%s served a challenge page (%d/%d before this source is dropped).",
            adapter.name, seen, runtime.max_consecutive_blocks,
        )
        return

    if adapter.name not in runtime.blocked_directories:
        runtime.blocked_directories.add(adapter.name)
        LOGGER.warning(
            "Dropping %s for the rest of this run\nReason: %s on %d consecutive pages",
            adapter.name, REASON_SECURITY_VERIFICATION, runtime.max_consecutive_blocks,
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

    match = extract_official_website(soup, adapter, result.url, record.company_name)
    if not match.url:
        note_source_event(runtime, adapter.name, EVENT_SELECTOR_FAILURES)
        LOGGER.info(
            "Official website not found for %s on %s\nReason: %s",
            record.company_name or "(unnamed)", result.url, match.reason,
        )
        runtime.skip_reasons[f"website not found: {match.reason}"] += 1
        return None

    note_source_event(runtime, adapter.name, EVENT_CANDIDATES_FOUND)

    record.website = match.url
    record.website_source = match.method
    record.website_link_text = match.link_text
    record.root_domain = root_domain(match.url)

    LOGGER.info(
        "Candidate website for %s: %s (via %s)",
        record.company_name or "(unnamed)", record.website, record.website_source,
    )
    return record


# ----------------------------------------------------------------------
# Website verification and validation
# ----------------------------------------------------------------------

def check_website(url: str) -> SiteCheck:
    """
    Fetch a candidate website and report what it says about itself.

    The identity check needs the title, meta description and a slice of the
    homepage text, and this response already contains all three, so they are
    taken here rather than with a second request.
    """
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

        content_type = response.headers.get("Content-Type", "").lower()
        html = "" if ("html" not in content_type and content_type) else response.text

        return SiteCheck(
            reachable=response.status_code < 400,
            status_code=response.status_code,
            final_url=response.url,
            title=extract_title(html),
            meta_description=extract_meta_description_from_html(html),
            text_sample=extract_visible_text(html),
            blocked=bool(html) and looks_blocked(html),
        )

    return SiteCheck(reason="unreachable")


def extract_title(html: str) -> str:
    """Pull the page title out of raw HTML without a full parse."""
    found = TITLE_PATTERN.search(html[:MAX_RESPONSE_SNIFF_CHARS])
    if not found:
        return ""
    return re.sub(r"\s+", " ", found.group(1)).strip()[:200]


def extract_meta_description_from_html(html: str) -> str:
    """Read a description meta tag out of raw HTML without a full parse."""
    wanted = ("description", "og:description", "twitter:description")

    for found in META_TAG_PATTERN.finditer(html[:MAX_RESPONSE_SNIFF_CHARS]):
        attributes: dict[str, str] = {}
        for attribute in META_ATTR_PATTERN.finditer(found.group(1)):
            key = attribute.group(1).lower()
            value = attribute.group(2) or attribute.group(3) or attribute.group(4) or ""
            attributes[key] = value

        label = (attributes.get("name") or attributes.get("property") or "").lower()
        content = attributes.get("content", "").strip()

        if label in wanted and content:
            return re.sub(r"\s+", " ", content)[:MAX_META_DESCRIPTION_CHARS]

    return ""


def extract_visible_text(html: str) -> str:
    """
    A small slice of readable page text, used only for the identity check.

    Scripts and styles are removed first so a company name is not "found"
    inside an analytics snippet.
    """
    if not html:
        return ""

    sample = html[:MAX_RESPONSE_SNIFF_CHARS]
    sample = SCRIPT_STYLE_PATTERN.sub(" ", sample)
    sample = HTML_TAG_PATTERN.sub(" ", sample)
    return re.sub(r"\s+", " ", sample).strip()[:MAX_HOMEPAGE_TEXT_CHARS]


def validate_record(record: CompanyRecord, check: SiteCheck) -> str:
    """
    Decide whether a candidate is a genuine company.

    Returns an empty string when the record may continue to the identity
    check, otherwise the reason it was rejected.
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

    if check.meta_description and not record.meta_description:
        record.meta_description = check.meta_description

    # Follow the redirect chain so the exported website is the one that answers.
    if check.reachable and check.final_url and is_plausible_website(check.final_url):
        record.website = check.final_url
        record.root_domain = root_domain(check.final_url)


# ----------------------------------------------------------------------
# Identity consistency validation
# ----------------------------------------------------------------------
#
# Directory Profile -> Company Name -> Candidate Website -> Website Fetch
# -> Identity Validation -> Accept / Review / Reject
#
# The question is only ever "does this site belong to this company". The
# deeper research is the Intelligence Engine's job, so nothing here crawls
# further than the homepage the reachability check already downloaded.

def _score_domain(
    slug: str,
    tokens_key: list[str],
    joined_all: str,
    joined_key: str,
    generic_only: bool,
) -> str:
    """Compare the company name with the domain. Returns a signal name."""
    if not slug:
        return ""

    # A whole-name match is conclusive. A substring match only counts for
    # names long enough that the overlap cannot be a coincidence.
    if joined_key and (joined_key == slug or (len(joined_key) >= 5 and joined_key in slug)):
        return "domain-exact"

    if generic_only and joined_all:
        if joined_all == slug or (len(joined_all) >= 5 and joined_all in slug):
            return "domain-exact"
        if similarity(joined_all, slug) >= DOMAIN_FUZZY_RATIO:
            return "domain-fuzzy"
        return ""

    if not tokens_key:
        return ""

    if any(token_in_domain(token, slug) for token in tokens_key):
        return "domain-token"

    acronym = "".join(token[0] for token in tokens_key)
    if len(acronym) >= 3 and acronym == slug:
        return "domain-acronym"

    if joined_key and similarity(joined_key, slug) >= DOMAIN_FUZZY_RATIO:
        return "domain-fuzzy"

    return ""


def _score_title(
    title: str,
    tokens_key: list[str],
    joined_all: str,
    joined_key: str,
    generic_only: bool,
) -> str:
    """Compare the company name with the website's own title."""
    if not title:
        return ""

    title_norm = normalize_text(title)
    words = text_words(title)

    # One short distinctive word is not a strong match on its own: plenty of
    # unrelated titles contain a three letter word.
    strong_enough = len(tokens_key) >= 2 or len(joined_key) >= 5

    if joined_key and strong_enough and joined_key in title_norm:
        return "title-strong"

    if generic_only and joined_all and len(joined_all) >= 5 and joined_all in title_norm:
        return "title-strong"

    if tokens_key:
        if strong_enough and all(token in words for token in tokens_key):
            return "title-strong"
        if any(token in words for token in tokens_key):
            return "title-token"

    brand = normalize_text(title_brand_segment(title))
    if joined_key and brand and similarity(joined_key, brand) >= TITLE_FUZZY_RATIO:
        return "title-fuzzy"

    return ""


def _name_in_text(
    text: str,
    tokens_key: list[str],
    joined_all: str,
    generic_only: bool,
) -> bool:
    """True when the company name appears in a block of page text."""
    if not text:
        return False

    if tokens_key:
        return any(token in text_words(text) for token in tokens_key)

    if generic_only and joined_all and len(joined_all) >= 5:
        return joined_all in normalize_text(text)

    return False


def _extraction_signals(website_source: str) -> list[str]:
    """How strongly the directory itself vouched for this link."""
    signals: list[str] = []

    if website_source in STRONG_WEBSITE_METHODS:
        signals.append("explicit-website-link")
    elif website_source in SCOPED_WEBSITE_METHODS:
        signals.append("profile-scoped-link")

    if website_source == "link-text-domain":
        signals.append("link-text-domain")

    return signals


def _names_a_different_company(title: str, tokens_all: list[str], slug: str) -> bool:
    """
    True when the site's own title brands it as some other company.

    Only consulted when nothing else connects the name to the site, and a
    generic title such as "Home" is never treated as a competing identity.

    Titles are not reliably "Brand | Tagline", so when the brand segment
    yields nothing the whole title is considered - but then the domain has
    to corroborate it. A title token that is also the domain name is the
    site stating its own identity, and if that identity is not the company
    we were looking for, the pairing is wrong.
    """
    owned = set(tokens_all)

    brand = title_brand_segment(title)
    if brand and brand.strip().lower() not in PLACEHOLDER_NAMES:
        brand_tokens = distinctive_name_tokens(brand)
        if brand_tokens:
            return not set(brand_tokens) & owned

    title_tokens = distinctive_name_tokens(title)
    if not title_tokens or set(title_tokens) & owned:
        return False

    return any(token_in_domain(token, slug) for token in title_tokens)


def assess_identity(record: CompanyRecord, check: SiteCheck) -> IdentityAssessment:
    """
    Decide whether the candidate website belongs to the discovered company.

    The score is additive and every point is attributable to a named signal,
    so any row in the audit workbook can be explained from its Identity
    Signals column. The name is never required to appear in the domain:
    "Example Growth" at agencyexample.com is matched on its title just as
    readily as on its domain.
    """
    tokens_all = name_tokens(record.company_name)
    tokens_key = distinctive_name_tokens(record.company_name)
    generic_only = bool(tokens_all) and not tokens_key
    joined_all = "".join(tokens_all)
    joined_key = "".join(tokens_key)

    signals: list[str] = []

    domain_signal = _score_domain(
        domain_slug(record.website), tokens_key, joined_all, joined_key, generic_only,
    )
    if domain_signal:
        signals.append(domain_signal)

    title_signal = _score_title(check.title, tokens_key, joined_all, joined_key, generic_only)
    if title_signal:
        signals.append(title_signal)

    if _name_in_text(check.meta_description, tokens_key, joined_all, generic_only):
        signals.append("meta-name")

    if _name_in_text(check.text_sample, tokens_key, joined_all, generic_only):
        signals.append("homepage-name")

    has_name_signal = bool(signals)
    signals.extend(_extraction_signals(record.website_source))

    # The site answered and told us who it is. If none of that mentions the
    # company the directory named, the pairing is not merely unproven - it is
    # contradicted.
    evidence_available = check.reachable and bool(check.title) and not check.blocked

    if not has_name_signal and evidence_available:
        signals.append("no-name-evidence")
        if _names_a_different_company(check.title, tokens_all, domain_slug(record.website)):
            signals.append("identity-mismatch")

    score = 0
    for name in signals:
        score += IDENTITY_POINTS.get(name, 0) + IDENTITY_PENALTIES.get(name, 0)
    score = max(0, min(100, score))

    mismatch = "identity-mismatch" in signals
    floor = IDENTITY_REVIEW_SCORE if evidence_available else IDENTITY_REVIEW_SCORE_UNVERIFIED

    if mismatch:
        verdict = VERDICT_REJECT
    elif score >= IDENTITY_ACCEPT_SCORE:
        verdict = VERDICT_ACCEPT
    else:
        verdict = VERDICT_REVIEW if score >= floor else VERDICT_REJECT

    # A name built only from industry words cannot identify a company on its
    # own. "Digital Media Solutions" appears inside other companies' titles
    # as a description of what they sell, so only a whole-name match on the
    # domain itself is accepted; everything else goes to review for a human.
    if verdict == VERDICT_ACCEPT and generic_only and "domain-exact" not in signals:
        verdict = VERDICT_REVIEW
        signals.append("generic-name-only")

    return IdentityAssessment(
        score=score,
        verdict=verdict,
        signals=signals,
        reason=_identity_reason(verdict, score, signals, mismatch, evidence_available, check),
    )


def _identity_reason(
    verdict: str,
    score: int,
    signals: list[str],
    mismatch: bool,
    evidence_available: bool,
    check: SiteCheck,
) -> str:
    """One readable sentence explaining the verdict, for the audit output."""
    if mismatch:
        return (
            f"Website is branded as a different company "
            f"(site title: {check.title[:80] or 'none'})"
        )

    if verdict == VERDICT_ACCEPT:
        return f"Identity confirmed by: {', '.join(signals) or 'no signals'}"

    if "generic-name-only" in signals:
        return "Company name is made only of industry words; needs a whole-name match"

    if not evidence_available and not check.reachable:
        return f"Website did not respond, so identity could not be confirmed (score {score})"

    if "no-name-evidence" in signals:
        return "Company name appears nowhere on the website"

    return f"Identity evidence too weak to export (score {score} of {IDENTITY_ACCEPT_SCORE} needed)"


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

    description = extract_meta_description(soup)
    if description:
        record.meta_description = description

    record.internal_pages = discover_internal_pages(soup, page.url or record.website)

    if not record.site_title and soup.title:
        record.site_title = soup.title.get_text(strip=True)[:200]


# ----------------------------------------------------------------------
# Confidence
# ----------------------------------------------------------------------

def collect_evidence(record: CompanyRecord) -> list[str]:
    """List what the run actually proved about this company."""
    evidence: list[str] = []

    if record.website_source in STRONG_WEBSITE_METHODS:
        evidence.append("explicit-website-link")
    if record.verified:
        evidence.append("website-reachable")

    if record.identity_score >= IDENTITY_STRONG_SCORE:
        evidence.append("identity-strong")
    elif record.identity_score >= IDENTITY_ACCEPT_SCORE:
        evidence.append("identity-moderate")

    if record.meta_description:
        evidence.append("meta-description")
    if record.internal_pages:
        evidence.append("internal-pages")

    return evidence


def grade_confidence(record: CompanyRecord) -> tuple[str, int, list[str]]:
    """
    Score the collected evidence and turn it into a confidence band.

    Three caps keep the band honest, and all of them are absolute:

        - a website that never answered stays Low, because without a
          reachable site the rest of the evidence is unconfirmed;
        - a record whose identity never reached the accept threshold stays
          Low, so a review-queue or rejected row can never read as a
          confident find;
        - a record whose identity is only moderately established cannot
          reach High however much surrounding metadata it collected. A
          reachable site with a meta description is not evidence that the
          site belongs to this company.
    """
    evidence = collect_evidence(record)
    score = sum(EVIDENCE_POINTS[item] for item in evidence)

    if not record.verified or record.identity_score < IDENTITY_ACCEPT_SCORE:
        band = "Low"
    elif score >= CONFIDENCE_HIGH_SCORE and record.identity_score >= IDENTITY_STRONG_SCORE:
        band = "High"
    elif score >= CONFIDENCE_MEDIUM_SCORE:
        band = "Medium"
    else:
        band = "Low"

    return band, score, evidence


def apply_confidence(record: CompanyRecord) -> None:
    """Grade a record and log the reasoning behind the band it received."""
    record.discovery_confidence, record.confidence_score, record.evidence = grade_confidence(record)

    LOGGER.info(
        "Graded %s as %s (%d points: %s | identity %d: %s)",
        record.company_name, record.discovery_confidence, record.confidence_score,
        ", ".join(record.evidence) or "none",
        record.identity_score, ", ".join(record.identity_signals) or "none",
    )


# ----------------------------------------------------------------------
# Collection, progress and checkpointing
# ----------------------------------------------------------------------

def safe_append(runtime: Runtime, record: CompanyRecord) -> bool:
    """
    Add a record to the run, deduplicating by root domain.

    This is the only place records enter the run, so it is the only place
    within-run deduplication happens. Cross-run deduplication is separate and
    lives in the registry.
    """
    if not record.website or not record.company_name:
        return False

    domain = record.root_domain or root_domain(record.website)
    if not domain:
        return False

    if domain in runtime.seen_domains:
        runtime.duplicates_removed += 1
        note_source_event(runtime, record.source, EVENT_RUN_DUPLICATES)
        LOGGER.info("Duplicate skipped: %s (%s)", record.company_name, domain)
        return False

    record.root_domain = domain
    record.company_name = clean_company_name(record.company_name)
    record.status = STATUS_ACCEPTED
    runtime.seen_domains.add(domain)
    runtime.records.append(record)

    LOGGER.info("Company added: %s | %s", record.company_name, record.website)
    return True


def store_audit_record(bucket: list[CompanyRecord], record: CompanyRecord) -> None:
    """Keep an audit row, bounded so a long run cannot bloat the workbook."""
    if len(bucket) < MAX_AUDIT_ROWS:
        bucket.append(record)


def reject_record(runtime: Runtime, record: CompanyRecord, reason: str, event: str) -> None:
    """Record a rejection in the audit trail rather than dropping it silently."""
    record.status = STATUS_REJECTED
    record.rejection_reason = reason
    runtime.records_rejected += 1
    note_source_event(runtime, record.source, event)
    store_audit_record(runtime.rejected_records, record)
    record_skip(runtime, record.website, reason)


def review_record(runtime: Runtime, record: CompanyRecord, reason: str) -> None:
    """Send a record that cannot be confidently matched to the review queue."""
    record.status = STATUS_REVIEW
    record.rejection_reason = reason
    runtime.records_review += 1
    note_source_event(runtime, record.source, EVENT_IDENTITY_REVIEW)
    store_audit_record(runtime.review_records, record)

    LOGGER.info(
        "Review queue: %s | %s\nReason: %s",
        record.company_name, record.website, reason,
    )


def registry_gate(runtime: Runtime, record: CompanyRecord) -> bool:
    """
    Stop a company earlier runs already discovered from being processed again.

    Returns True when the candidate was recognised and should go no further.
    The company is still written to the audit workbook and its registry entry
    is refreshed, so nothing is lost - it simply does not reach
    prospector_input.csv a second time.

    In fresh mode the registry is consulted for logging only, so the whole
    universe can be deliberately rediscovered.
    """
    registry = runtime.registry
    entry = registry.knows(
        record.root_domain,
        normalized_company_name(record.company_name),
        record.source,
    )

    if entry is None:
        return False

    record.previously_discovered = True

    if not registry.enforcing:
        LOGGER.info(
            "Known company %s kept anyway (registry mode=%s).",
            record.company_name, registry.mode,
        )
        return False

    timestamp = datetime.now().isoformat(timespec="seconds")
    registry.touch(entry, timestamp)

    runtime.registry_duplicates += 1
    note_source_event(runtime, record.source, EVENT_REGISTRY_DUPLICATES)
    runtime.skip_reasons[REASON_KNOWN_COMPANY] += 1

    store_known_company(runtime, record, entry, timestamp)

    LOGGER.info(
        "Known from an earlier run: %s (%s, first discovered %s)",
        record.company_name, record.root_domain,
        entry.get("first_discovered_at", "unknown"),
    )
    return True


def store_known_company(
    runtime: Runtime,
    record: CompanyRecord,
    entry: dict,
    timestamp: str,
) -> None:
    """Keep a previously discovered company in the audit data for this run."""
    if len(runtime.known_companies) >= MAX_AUDIT_ROWS:
        return

    runtime.known_companies.append(
        KnownCompany(
            company_name=record.company_name or entry.get("company_name", ""),
            website=record.website or entry.get("website", ""),
            root_domain=record.root_domain or entry.get("domain", ""),
            source=record.source,
            source_url=record.source_url,
            first_discovered_at=entry.get("first_discovered_at", ""),
            times_seen=int(entry.get("times_seen", 1)),
            search_query=record.search_query,
            seen_at=timestamp,
        )
    )


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
        f"review {runtime.records_review:>3} | "
        f"rejected {runtime.records_rejected:>3} | "
        f"known {runtime.registry_duplicates:>3} | "
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
        "records_review": runtime.records_review,
        "duplicates_removed": runtime.duplicates_removed,
        "registry_duplicates": runtime.registry_duplicates,
        "errors": runtime.errors,
        "error_messages": runtime.error_messages,
        "skip_reasons": dict(runtime.skip_reasons),
        "blocked_directories": sorted(runtime.blocked_directories),
        "consecutive_blocks": dict(runtime.consecutive_blocks),
        "source_stats": {
            source: dict(counter) for source, counter in runtime.source_stats.items()
        },
        "visited_urls": sorted(runtime.visited_urls),
        "records": [asdict(record) for record in runtime.records],
        "review_records": [asdict(record) for record in runtime.review_records],
        "rejected_records": [asdict(record) for record in runtime.rejected_records],
        "known_companies": [asdict(company) for company in runtime.known_companies],
    }

    save_autosave(state)
    runtime.registry.save()

    LOGGER.info(
        "Checkpoint saved at query %d (%d companies, %d in review).",
        runtime.query_index, runtime.companies_found, runtime.records_review,
    )


def save_autosave(state: dict) -> None:
    """Write the autosave payload to disk atomically."""
    temporary = AUTOSAVE_FILE.with_suffix(AUTOSAVE_FILE.suffix + ".tmp")

    try:
        temporary.write_text(json.dumps(state, indent=2), encoding="utf-8")
        os.replace(temporary, AUTOSAVE_FILE)
    except OSError as error:
        LOGGER.error("Could not write the autosave file: %s", error)


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


def _records_from_state(payloads: list) -> list[CompanyRecord]:
    """Rebuild company records from an autosave, ignoring unknown fields."""
    known_fields = set(CompanyRecord.__dataclass_fields__)
    rebuilt: list[CompanyRecord] = []

    for payload in payloads or []:
        if isinstance(payload, dict):
            rebuilt.append(
                CompanyRecord(**{k: v for k, v in payload.items() if k in known_fields})
            )

    return rebuilt


def restore_runtime(runtime: Runtime, state: dict) -> None:
    """Rebuild the runtime from an autosave so counters and records survive."""
    for record in _records_from_state(state.get("records", [])):
        runtime.records.append(record)
        if record.root_domain:
            runtime.seen_domains.add(record.root_domain)

    runtime.review_records.extend(_records_from_state(state.get("review_records", [])))
    runtime.rejected_records.extend(_records_from_state(state.get("rejected_records", [])))

    known_fields = set(KnownCompany.__dataclass_fields__)
    for payload in state.get("known_companies", []) or []:
        if isinstance(payload, dict):
            runtime.known_companies.append(
                KnownCompany(**{k: v for k, v in payload.items() if k in known_fields})
            )

    runtime.visited_urls.update(state.get("visited_urls", []))
    runtime.error_messages.extend(state.get("error_messages", []))
    runtime.skip_reasons.update(state.get("skip_reasons", {}))
    runtime.blocked_directories.update(state.get("blocked_directories", []))
    runtime.consecutive_blocks.update(state.get("consecutive_blocks", {}))

    for source, counts in (state.get("source_stats", {}) or {}).items():
        runtime.source_stats.setdefault(source, Counter()).update(counts)

    runtime.query_index = state.get("query_index", 0)
    runtime.queries_completed = state.get("queries_completed", 0)
    runtime.results_seen = state.get("results_seen", 0)
    runtime.results_kept = state.get("results_kept", 0)
    runtime.profiles_visited = state.get("profiles_visited", 0)
    runtime.profiles_blocked = state.get("profiles_blocked", 0)
    runtime.websites_checked = state.get("websites_checked", 0)
    runtime.records_rejected = state.get("records_rejected", 0)
    runtime.records_review = state.get("records_review", 0)
    runtime.duplicates_removed = state.get("duplicates_removed", 0)
    runtime.registry_duplicates = state.get("registry_duplicates", 0)
    runtime.errors = state.get("errors", 0)

    LOGGER.info(
        "Resumed from query %d with %d companies already discovered "
        "(%d dropped sources carried over).",
        runtime.query_index, runtime.companies_found, len(runtime.blocked_directories),
    )
    print(
        f"Resuming from query {runtime.query_index + 1} "
        f"with {runtime.companies_found} companies already discovered."
    )

    if runtime.blocked_directories:
        print(f"Still skipping blocked sources: {', '.join(sorted(runtime.blocked_directories))}")


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

    Two families of numbers are kept apart on purpose:

        - attempt-level counters describe what the scrape did and can exceed
          the number of companies (many profiles produce no company);
        - final-record counters describe the exported dataset and are always
          counted from runtime.records, so verified companies can never
          exceed discovered companies.
    """
    return {
        "companies_found": runtime.companies_found,
        "companies_exported": runtime.companies_exported,
        "verified_websites": runtime.websites_verified,
        "review_queue": runtime.records_review,
        "queries_planned": len(runtime.queries),
        "queries_completed": runtime.queries_completed,
        "results_seen": runtime.results_seen,
        "results_kept": runtime.results_kept,
        "profiles_visited": runtime.profiles_visited,
        "profiles_blocked": runtime.profiles_blocked,
        "websites_checked": runtime.websites_checked,
        "records_rejected": runtime.records_rejected,
        "duplicates_removed": runtime.duplicates_removed,
        "registry_duplicates": runtime.registry_duplicates,
        "registry_size": len(runtime.registry.entries),
        "registry_mode": runtime.registry.mode,
        "errors": runtime.errors,
        "runtime": format_duration(time.time() - runtime.started_at),
        "blocked_directories": sorted(runtime.blocked_directories),
        "sources": dict(Counter(record.source for record in runtime.records)),
        "confidence": dict(Counter(record.discovery_confidence for record in runtime.records)),
        "website_sources": dict(Counter(record.website_source for record in runtime.records)),
        "identity_bands": dict(
            Counter(identity_band(record.identity_score) for record in runtime.records)
        ),
        "skip_reasons": dict(runtime.skip_reasons.most_common()),
    }


def identity_band(score: int) -> str:
    """Group an identity score into a band for the summary breakdown."""
    if score >= IDENTITY_STRONG_SCORE:
        return f"Strong ({IDENTITY_STRONG_SCORE}+)"
    if score >= IDENTITY_ACCEPT_SCORE:
        return f"Accepted ({IDENTITY_ACCEPT_SCORE}-{IDENTITY_STRONG_SCORE - 1})"
    return f"Below accept (<{IDENTITY_ACCEPT_SCORE})"


def format_duration(seconds: float) -> str:
    """Format an elapsed duration as HH:MM:SS."""
    minutes, secs = divmod(int(seconds), 60)
    hours, minutes = divmod(minutes, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def print_statistics(stats: dict) -> None:
    """Print the end-of-run summary to the console."""
    print("\n========== Discovery Summary ==========\n")
    print(f"Companies Found   : {stats['companies_found']}")
    print(f"Exported to CSV   : {stats['companies_exported']}")
    print(f"Verified Websites : {stats['verified_websites']}")
    print(f"Review Queue      : {stats['review_queue']}")
    print(f"Queries Completed : {stats['queries_completed']} / {stats['queries_planned']}")
    print(f"Profiles Visited  : {stats['profiles_visited']}")
    print(f"Profiles Blocked  : {stats['profiles_blocked']}")
    print(f"Rejected          : {stats['records_rejected']}")
    print(f"Duplicates Removed: {stats['duplicates_removed']}")
    print(f"Known From Before : {stats['registry_duplicates']} (registry mode: {stats['registry_mode']})")
    print(f"Registry Size     : {stats['registry_size']}")
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


def master_row(record: CompanyRecord) -> list:
    """One row of the accepted-companies sheet."""
    return [
        clean_cell(record.company_name),
        clean_cell(record.website),
        record.source,
        clean_cell(record.source_url),
        record.search_query,
        record.status,
        record.discovery_confidence,
        record.confidence_score,
        ", ".join(record.evidence),
        record.identity_score,
        record.identity_verdict,
        ", ".join(record.identity_signals),
        clean_cell(record.meta_description),
        clean_cell(" | ".join(record.internal_pages)),
        record.keyword,
        record.city,
        record.root_domain,
        "Yes" if record.verified else "No",
        record.website_source,
        record.http_status,
        clean_cell(record.site_title),
        "Yes" if record.previously_discovered else "No",
        "Yes" if record.exported else "No",
        record.discovered_at,
    ]


def audit_row(record: CompanyRecord) -> list:
    """One row of the review queue or rejected sheet."""
    return [
        clean_cell(record.company_name),
        clean_cell(record.website),
        record.source,
        clean_cell(record.source_url),
        record.identity_score,
        record.identity_verdict,
        record.discovery_confidence,
        clean_cell(record.site_title),
        clean_cell(record.meta_description),
        clean_cell(record.rejection_reason or record.identity_reason),
        ", ".join(record.identity_signals),
        record.website_source,
        record.http_status,
        "Yes" if record.verified else "No",
        record.search_query,
        record.discovered_at,
    ]


def export_master(runtime: Runtime) -> None:
    """
    Write master_company_list.xlsx.

    Accepted, review and rejected records are kept on separate sheets so the
    three outcomes are never confused with one another, and companies known
    from earlier runs stay in the audit data on a sheet of their own.
    """
    workbook = Workbook()

    accepted = workbook.active
    accepted.title = "Companies"
    accepted.append(MASTER_COLUMNS)
    for record in runtime.records:
        accepted.append(master_row(record))

    review = workbook.create_sheet("Review Queue")
    review.append(AUDIT_COLUMNS)
    for record in runtime.review_records:
        review.append(audit_row(record))

    rejected = workbook.create_sheet("Rejected")
    rejected.append(AUDIT_COLUMNS)
    for record in runtime.rejected_records:
        rejected.append(audit_row(record))

    known = workbook.create_sheet("Previously Discovered")
    known.append(KNOWN_COLUMNS)
    for company in runtime.known_companies:
        known.append([
            clean_cell(company.company_name),
            clean_cell(company.website),
            company.root_domain,
            company.source,
            clean_cell(company.source_url),
            company.first_discovered_at,
            company.times_seen,
            company.search_query,
            company.seen_at,
        ])

    workbook.save(OUTPUT_MASTER)
    LOGGER.info(
        "Master list written to %s (%d accepted, %d review, %d rejected, %d known).",
        OUTPUT_MASTER.name, runtime.companies_found,
        len(runtime.review_records), len(runtime.rejected_records),
        len(runtime.known_companies),
    )


def export_prospector(runtime: Runtime) -> None:
    """
    Write prospector_input.csv with exactly the two columns Engine 2 expects.

    The schema is frozen. Only accepted companies that are new to the
    registry are written: a record in the review queue, a rejected record and
    a company an earlier run already handed over all stay out.
    """
    exported = 0

    with OUTPUT_PROSPECTOR.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(PROSPECTOR_COLUMNS)

        for record in runtime.records:
            if not record.exported:
                continue
            writer.writerow([record.company_name, record.website])
            exported += 1

    LOGGER.info(
        "Prospector input written to %s (%d companies).",
        OUTPUT_PROSPECTOR.name, exported,
    )


def create_run_report(runtime: Runtime, stats: dict) -> None:
    """Write run_report.xlsx with the run summary, breakdowns and diagnostics."""
    workbook = Workbook()

    summary = workbook.active
    summary.title = "Run Report"
    summary.append(["Metric", "Value"])

    summary.append(["Companies Found", stats["companies_found"]])
    summary.append(["Companies Exported", stats["companies_exported"]])
    summary.append(["Verified Websites", stats["verified_websites"]])
    summary.append(["Review Queue", stats["review_queue"]])
    summary.append(["Records Rejected", stats["records_rejected"]])
    summary.append([])
    summary.append(["Queries Planned", stats["queries_planned"]])
    summary.append(["Queries Completed", stats["queries_completed"]])
    summary.append(["Search Results Seen", stats["results_seen"]])
    summary.append(["Search Results Kept", stats["results_kept"]])
    summary.append(["Profiles Visited", stats["profiles_visited"]])
    summary.append(["Profiles Blocked", stats["profiles_blocked"]])
    summary.append(["Websites Checked", stats["websites_checked"]])
    summary.append(["Duplicates Removed (this run)", stats["duplicates_removed"]])
    summary.append(["Known From Earlier Runs", stats["registry_duplicates"]])
    summary.append(["Registry Mode", stats["registry_mode"]])
    summary.append(["Registry Size", stats["registry_size"]])
    summary.append(["Errors", stats["errors"]])
    summary.append(["Runtime", stats["runtime"]])
    summary.append(["Dropped Sources", ", ".join(stats["blocked_directories"])])

    _append_breakdown(summary, "Source Breakdown", stats["sources"])
    _append_breakdown(summary, "Confidence Breakdown", stats["confidence"])
    _append_breakdown(summary, "Identity Score Breakdown", stats["identity_bands"])
    _append_breakdown(summary, "Website Source Breakdown", stats["website_sources"])

    _write_source_diagnostics(workbook, runtime)

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


def _write_source_diagnostics(workbook: Workbook, runtime: Runtime) -> None:
    """
    One row per directory showing where its candidates were lost.

    Reading across a row explains a source's yield: how many profiles it was
    asked for, how many loaded, how many were challenge pages, how many
    produced a candidate website, and how many of those survived validation
    and the identity check.
    """
    sheet = workbook.create_sheet("Source Diagnostics")
    sheet.append(["Source"] + list(SOURCE_EVENTS) + ["Dropped For This Run"])

    for name in DIRECTORIES:
        counter = runtime.source_stats.get(name, Counter())
        sheet.append(
            [name]
            + [counter.get(event, 0) for event in SOURCE_EVENTS]
            + ["Yes" if name in runtime.blocked_directories else "No"]
        )


def _append_breakdown(sheet, heading: str, values: dict) -> None:
    """Append one titled key/value block to the summary sheet."""
    sheet.append([])
    sheet.append([heading, ""])
    for key, count in values.items():
        sheet.append([key or "unknown", count])


def print_source_diagnostics(runtime: Runtime) -> None:
    """Print the per-source funnel to the console at the end of the run."""
    print("Source Diagnostics")

    for name in DIRECTORIES:
        counter = runtime.source_stats.get(name, Counter())
        dropped = " [dropped]" if name in runtime.blocked_directories else ""
        print(
            f"  {name:11} attempted {counter.get(EVENT_PROFILES_ATTEMPTED, 0):>4} | "
            f"loaded {counter.get(EVENT_PROFILES_LOADED, 0):>4} | "
            f"challenges {counter.get(EVENT_SECURITY_CHALLENGES, 0):>3} | "
            f"timeouts {counter.get(EVENT_TIMEOUTS, 0):>3} | "
            f"no website link {counter.get(EVENT_SELECTOR_FAILURES, 0):>4} | "
            f"candidates {counter.get(EVENT_CANDIDATES_FOUND, 0):>4} | "
            f"mismatches {counter.get(EVENT_IDENTITY_REJECTED, 0):>3} | "
            f"review {counter.get(EVENT_IDENTITY_REVIEW, 0):>3} | "
            f"accepted {counter.get(EVENT_ACCEPTED, 0):>4}{dropped}"
        )

    print()


def verify_statistics(runtime: Runtime, stats: dict) -> None:
    """
    Guard the invariants QA asked for, and say so in the log either way.

    These are cheap checks on numbers that have been wrong before: anything
    describing exported companies must be counted from the final records.
    """
    problems = []

    if stats["verified_websites"] > stats["companies_found"]:
        problems.append("verified websites exceed companies found")
    if stats["companies_exported"] > stats["companies_found"]:
        problems.append("exported companies exceed companies found")
    if sum(stats["sources"].values()) != stats["companies_found"]:
        problems.append("source breakdown does not add up to companies found")
    if len(runtime.seen_domains) != stats["companies_found"]:
        problems.append("deduplication index disagrees with the record count")

    for problem in problems:
        LOGGER.error("Statistics inconsistency: %s", problem)

    if not problems:
        LOGGER.info("Statistics consistency checks passed.")


def finalize_run(runtime: Runtime) -> dict:
    """Write every output file and print the summary. Records are already unique."""
    stats = build_run_statistics(runtime)
    verify_statistics(runtime, stats)

    export_master(runtime)
    export_prospector(runtime)
    create_run_report(runtime, stats)
    print_statistics(stats)
    print_source_diagnostics(runtime)

    runtime.registry.save()

    LOGGER.info(
        "Discovery run completed with %d companies (%d exported, %d in review).",
        runtime.companies_found, runtime.companies_exported, runtime.records_review,
    )
    return stats


# ----------------------------------------------------------------------
# Orchestration
# ----------------------------------------------------------------------

def register_accepted(runtime: Runtime, record: CompanyRecord) -> None:
    """
    Decide whether an accepted company is new, and record it in the registry.

    The check is repeated here because following a redirect can change the
    root domain after the early registry gate has already run.
    """
    registry = runtime.registry
    timestamp = datetime.now().isoformat(timespec="seconds")

    entry = registry.knows(
        record.root_domain,
        normalized_company_name(record.company_name),
        record.source,
    )

    record.previously_discovered = entry is not None
    record.exported = not (registry.enforcing and record.previously_discovered)

    if record.exported:
        registry.remember(record, timestamp)
    elif entry is not None:
        registry.touch(entry, timestamp)
        runtime.registry_duplicates += 1
        note_source_event(runtime, record.source, EVENT_REGISTRY_DUPLICATES)
        LOGGER.info(
            "Accepted but held back from the export: %s is already in the registry.",
            record.company_name,
        )


def process_result(
    page: Page,
    runtime: Runtime,
    search: SearchQuery,
    result: SearchResult,
) -> None:
    """
    Take one search result from profile page to stored, graded company.

    The path is: load the profile, read the name, find a website the page
    actually identifies as the company's own, fetch it, validate it, check
    that its identity agrees with the name, then accept, review or reject.
    """
    adapter = DIRECTORIES[search.directory]
    note_source_event(runtime, adapter.name, EVENT_PROFILES_ATTEMPTED)

    soup = open_profile(page, adapter, result, runtime)
    if soup is None:
        return

    record = build_company_record(soup, search, result, runtime)
    if record is None:
        return

    if registry_gate(runtime, record):
        return

    runtime.websites_checked += 1
    check = check_website(record.website)
    apply_site_check(record, check)

    if not check.reachable:
        note_source_event(runtime, adapter.name, EVENT_WEBSITE_UNREACHABLE)

    rejection = validate_record(record, check)
    if rejection:
        record.identity_verdict = VERDICT_REJECT
        record.identity_reason = rejection
        apply_confidence(record)
        reject_record(runtime, record, rejection, EVENT_VALIDATION_REJECTED)
        return

    identity = assess_identity(record, check)
    record.identity_score = identity.score
    record.identity_verdict = identity.verdict
    record.identity_signals = identity.signals
    record.identity_reason = identity.reason

    if identity.verdict != VERDICT_ACCEPT:
        apply_confidence(record)
        if identity.verdict == VERDICT_REJECT:
            reject_record(runtime, record, identity.reason, EVENT_IDENTITY_REJECTED)
        else:
            review_record(runtime, record, identity.reason)
        return

    if not safe_append(runtime, record):
        return

    register_accepted(runtime, record)

    if record.verified:
        enrich_company_record(page, record, runtime)

    apply_confidence(record)

    note_source_event(runtime, adapter.name, EVENT_ACCEPTED)
    if record.exported:
        note_source_event(runtime, adapter.name, EVENT_EXPORTED)


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


def parse_arguments(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse the optional command line flags. Running with no flags is normal."""
    parser = argparse.ArgumentParser(
        description="Discover ICP companies from public directories.",
    )
    parser.add_argument(
        "--fresh",
        action="store_true",
        help="ignore any existing autosave and start the query list from the top",
    )
    parser.add_argument(
        "--fresh-universe",
        action="store_true",
        help="ignore the cross-run registry and rediscover companies earlier runs found",
    )
    parser.add_argument(
        "--registry-mode",
        choices=(REGISTRY_MODE_ACCUMULATE, REGISTRY_MODE_FRESH),
        default=DEFAULT_REGISTRY_MODE,
        help=(
            "accumulate (default) keeps previously discovered companies out of "
            "prospector_input.csv; fresh rediscovers them"
        ),
    )
    parser.add_argument(
        "--registry-file",
        type=Path,
        default=REGISTRY_FILE,
        help="path to the cross-run discovery registry",
    )
    parser.add_argument(
        "--max-consecutive-blocks",
        type=int,
        default=MAX_CONSECUTIVE_BLOCKS,
        help="challenge pages in a row before a directory is dropped for the run",
    )
    parser.add_argument("--max-queries", type=int, default=None, help="limit how many queries run")
    parser.add_argument("--headless", action="store_true", help="run Chromium without a window")

    arguments = parser.parse_args(argv)

    # --fresh-universe is the friendly spelling of --registry-mode fresh.
    if arguments.fresh_universe:
        arguments.registry_mode = REGISTRY_MODE_FRESH

    if arguments.max_consecutive_blocks < 1:
        parser.error("--max-consecutive-blocks must be at least 1")

    return arguments


def main() -> int:
    """Single entry point: set up, discover, export, shut down."""
    args = parse_arguments()

    LOGGER.info(
        "Discovery Engine v0.3 started (registry mode=%s, max consecutive blocks=%d).",
        args.registry_mode, args.max_consecutive_blocks,
    )

    runtime = initialize_runtime(
        registry_path=args.registry_file,
        registry_mode=args.registry_mode,
        max_consecutive_blocks=args.max_consecutive_blocks,
    )
    runtime.registry.load()

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
    print(
        f"Registry: {len(runtime.registry.entries)} companies known from earlier runs "
        f"(mode: {runtime.registry.mode})."
    )

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
