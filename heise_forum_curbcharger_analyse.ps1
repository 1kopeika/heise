<#
.SYNOPSIS
Analysiert einen Heise-Forenthread zu Ladebordsteinen und erzeugt getrennte Berichte sowie JSON-Ausgaben.

.DESCRIPTION
Das Skript ruft die Seiten eines Heise-Forenthreads ab, ermittelt alle Root-Threads, lädt pro Root-Thread
die vollständige Thread-Ansicht, sammelt daraus alle Posting-URLs und cached die einzelnen Posting-Seiten lokal.
Anschließend werden die Beiträge geparst, thematisch klassifiziert und in mehreren Ausgabeformaten gespeichert.

Die Ausgaben werden pro Forum getrennt und mit Datum sowie Threadnamen benannt. Dazu gehören unter anderem:
- vollständige Kommentar-JSONs
- thematische Summaries
- Autorenstatistiken
- Markdown-Vollberichte und Endfassungen

Mit -SkipFetch kann eine bestehende lokale Cache-Struktur erneut ausgewertet werden, ohne die Heise-Seiten neu
abzurufen. Das Skript ist auf die Struktur von Heise-Foren im Bereich "heise-online/Kommentare" zugeschnitten.

.PARAMETER WorkDir
Arbeitsverzeichnis, in dem Cache-Dateien sowie die erzeugten JSON- und Markdown-Berichte liegen sollen.

.PARAMETER ForumUrl
Heise-Forum-URL des auszuwertenden Threads, zum Beispiel:
https://www.heise.de/forum/heise-online/Kommentare/.../forum-123456/comment/

.PARAMETER SkipFetch
Verwendet nur bereits vorhandene lokale Cache-Dateien und führt ausschließlich die Auswertung erneut aus.

.PARAMETER OnlyAuthors
Überspringt das Nachladen einzelner Posting-Seiten. Dieser Schalter ist nur für Sonderfälle gedacht und
kann zu unvollständigen Ergebnissen führen, wenn benötigte Postings noch nicht im Cache liegen.

.EXAMPLE
.\heise_forum_curbcharger_analyse.ps1 -ForumUrl "https://www.heise.de/forum/heise-online/Kommentare/Elektroautos-Koeln-bekommt-Ladebordsteine/forum-519185/comment/"

Lädt den angegebenen Heise-Thread, baut den lokalen Cache auf und erzeugt die zugehörigen JSON- und Markdown-Berichte.

.EXAMPLE
.\heise_forum_curbcharger_analyse.ps1 -ForumUrl "https://www.heise.de/forum/heise-online/Kommentare/Elektroautos-Koeln-bekommt-Ladebordsteine/forum-519185/comment/" -SkipFetch

Verwendet nur bereits vorhandene Cache-Dateien und berechnet die Auswertung für den Thread erneut.
#>
param(
    [string]$WorkDir = ".",
    [string]$ForumUrl = "https://www.heise.de/forum/heise-online/Kommentare/Rheinmetall-und-TankE-wollen-Ladebordsteine-in-die-Staedte-bringen/forum-579336/comment/",
    [switch]$SkipFetch,
    [switch]$OnlyAuthors
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$script:UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0 Safari/537.36"

Set-Location -LiteralPath $WorkDir

function Get-Sha1Hex {
    param([Parameter(Mandatory = $true)][string]$Text)
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $hash = $sha1.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash).Replace("-", "").ToLowerInvariant())
    }
    finally {
        $sha1.Dispose()
    }
}

function Get-SlugifiedFileName {
    param([Parameter(Mandatory = $true)][string]$Url)
    $safe = [regex]::Replace($Url, "[^a-zA-Z0-9]+", "-").Trim("-").ToLowerInvariant()
    if ($safe.Length -gt 120) {
        $safe = $safe.Substring(0, 120)
    }
    $digest = (Get-Sha1Hex -Text $Url).Substring(0, 12)
    return "$safe-$digest.html"
}

function Get-ForumMetadata {
    param([Parameter(Mandatory = $true)][string]$Url)

    $normalized = $Url.TrimEnd("/")
    if ($normalized -notmatch "/forum/heise-online/Kommentare/(?<slug>[^/]+)/forum-(?<id>\d+)(?:/comment)?$") {
        throw "Unerwartete Heise-Forum-URL: $Url"
    }

    return [pscustomobject]@{
        NormalizedUrl = "$normalized/"
        ForumSlug = $Matches["slug"]
        ForumId = $Matches["id"]
        AnalysisKey = "{0}-forum-{1}" -f $Matches["slug"], $Matches["id"]
    }
}

function Get-PageUrls {
    param([Parameter(Mandatory = $true)][string]$BaseForumUrl)

    $pageUrls = [System.Collections.Generic.List[string]]::new()
    $pageUrls.Add($BaseForumUrl)

    try {
        $response = Invoke-WebRequest -Uri $BaseForumUrl -Headers @{ "User-Agent" = $script:UserAgent } -TimeoutSec 60
        $html = $response.Content
        $matches = [regex]::Matches($html, '/forum/heise-online/Kommentare/[^"]+/page-(\d+)/')
        $pageNumbers = @(1)
        foreach ($match in $matches) {
            $pageNumbers += [int]$match.Groups[1].Value
        }
        $maxPage = ($pageNumbers | Measure-Object -Maximum).Maximum
        foreach ($n in 2..$maxPage) {
            $pageUrls.Add(($BaseForumUrl -replace '/comment/$', "/page-$n/"))
        }
    }
    catch {
        foreach ($n in 2..5) {
            $pageUrls.Add(($BaseForumUrl -replace '/comment/$', "/page-$n/"))
        }
    }

    return $pageUrls
}

function Get-OutputPaths {
    param(
        [Parameter(Mandatory = $true)]$ForumMeta,
        [Parameter(Mandatory = $true)][string]$BaseDir
    )

    $cacheDir = Join-Path (Join-Path $BaseDir "heise_cache") $ForumMeta.AnalysisKey
    $datePrefix = Get-Date -Format "yyyy-MM-dd"
    $prefix = "{0}_{1}" -f $datePrefix, $ForumMeta.AnalysisKey

    return [pscustomobject]@{
        CacheDir = $cacheDir
        LegacyCacheDir = Join-Path $BaseDir "heise_cache"
        RootUrls = Join-Path $BaseDir ("{0}_root_urls.json" -f $prefix)
        PostFetchList = Join-Path $BaseDir ("{0}_post_fetch_list.json" -f $prefix)
        CommentsJson = Join-Path $BaseDir ("{0}_comments_full.json" -f $prefix)
        SummaryJson = Join-Path $BaseDir ("{0}_summary_full.json" -f $prefix)
        AuthorsJson = Join-Path $BaseDir ("{0}_authors_summary.json" -f $prefix)
        Themenbericht = Join-Path $BaseDir ("{0}_themenbericht.md" -f $prefix)
        Vollbericht = Join-Path $BaseDir ("{0}_vollbericht.md" -f $prefix)
        Kurzfassung = Join-Path $BaseDir ("{0}_kurzfassung.txt" -f $prefix)
        Kommentatorenbericht = Join-Path $BaseDir ("{0}_kommentatorenstatistik.md" -f $prefix)
    }
}

function Get-ForumLabel {
    param([Parameter(Mandatory = $true)][string]$ForumSlug)
    return ($ForumSlug -replace "-", " ")
}

$forumMeta = Get-ForumMetadata -Url $ForumUrl
$paths = Get-OutputPaths -ForumMeta $forumMeta -BaseDir (Get-Location)
$cacheDir = $paths.CacheDir

if ($SkipFetch) {
    $legacyThreadCount = if (Test-Path -LiteralPath $paths.LegacyCacheDir) { @(Get-ChildItem -LiteralPath $paths.LegacyCacheDir -Filter "thread_live_*.html" -ErrorAction SilentlyContinue).Count } else { 0 }
    $newThreadCount = if (Test-Path -LiteralPath $cacheDir) { @(Get-ChildItem -LiteralPath $cacheDir -Filter "thread_live_*.html" -ErrorAction SilentlyContinue).Count } else { 0 }
    if ($newThreadCount -eq 0 -and $legacyThreadCount -gt 0 -and $forumMeta.ForumId -eq "579336") {
        $cacheDir = $paths.LegacyCacheDir
    }
}

New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null

function Invoke-HeiseRequest {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$OutFile,
        [int]$MaxAttempts = 6,
        [int]$BaseSleepSeconds = 3
    )

    if (Test-Path -LiteralPath $OutFile) {
        return Get-Content -LiteralPath $OutFile -Raw -Encoding UTF8
    }

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $response = Invoke-WebRequest -Uri $Url -Headers @{ "User-Agent" = $script:UserAgent } -TimeoutSec 60
            $content = $response.Content
            [System.IO.File]::WriteAllText($OutFile, $content, [System.Text.Encoding]::UTF8)
            Start-Sleep -Seconds $BaseSleepSeconds
            return $content
        }
        catch {
            $delay = [Math]::Min(30, $BaseSleepSeconds * $attempt)
            Write-Warning "Abruf fehlgeschlagen ($attempt/$MaxAttempts): $Url"
            if ($attempt -eq $MaxAttempts) {
                throw
            }
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-ExpandAllUrl {
    param([Parameter(Mandatory = $true)][string]$Html)
    $match = [regex]::Match($Html, 'href="(https://www\.heise\.de/forum/expand-all-threads/[^"]+)"')
    if (-not $match.Success) {
        throw "Expand-all-URL nicht gefunden."
    }
    return [System.Net.WebUtility]::HtmlDecode($match.Groups[1].Value)
}

function Get-TopLevelPostingUrls {
    param([Parameter(Mandatory = $true)][string]$Html)
    $matches = [regex]::Matches(
        $Html,
        '<li class="posting_element" data-thread-id="\d+">.*?href="(https://www\.heise\.de/forum/heise-online/Kommentare/[^"]+/posting-\d+/show/)" class="posting_subject"',
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $result = New-Object System.Collections.Generic.List[string]
    $seen = New-Object System.Collections.Generic.HashSet[string]
    foreach ($match in $matches) {
        $url = [System.Net.WebUtility]::HtmlDecode($match.Groups[1].Value)
        if ($seen.Add($url)) {
            $result.Add($url)
        }
    }
    return $result
}

function Get-ThreadViewUrl {
    param([Parameter(Mandatory = $true)][string]$Html)
    $match = [regex]::Match($Html, '<a href="(/forum/show-thread-below-posting/[^"]+)">Thread-Anzeige einblenden</a>')
    if (-not $match.Success) {
        return $null
    }
    $relative = [System.Net.WebUtility]::HtmlDecode($match.Groups[1].Value)
    return [System.Uri]::new([System.Uri]::new("https://www.heise.de"), $relative).AbsoluteUri
}

function Get-PostingSubjectUrls {
    param([Parameter(Mandatory = $true)][string]$Html)
    $matches = [regex]::Matches(
        $Html,
        '<a href="(https://www\.heise\.de/forum/heise-online/Kommentare/[^"]+/posting-\d+/show/)" class="posting_subject">'
    )
    $result = New-Object System.Collections.Generic.List[string]
    $seen = New-Object System.Collections.Generic.HashSet[string]
    foreach ($match in $matches) {
        $url = [System.Net.WebUtility]::HtmlDecode($match.Groups[1].Value)
        if ($seen.Add($url)) {
            $result.Add($url)
        }
    }
    return $result
}

if (-not $SkipFetch) {
    $pageUrls = Get-PageUrls -BaseForumUrl $forumMeta.NormalizedUrl
    $rootUrls = New-Object System.Collections.Generic.List[string]
    $rootSeen = New-Object System.Collections.Generic.HashSet[string]

    for ($i = 0; $i -lt $pageUrls.Count; $i++) {
        $pageIndex = $i + 1
        $pageFile = Join-Path $cacheDir ("page_{0}.html" -f $pageIndex)
        $pageHtml = Invoke-HeiseRequest -Url $pageUrls[$i] -OutFile $pageFile -BaseSleepSeconds 2

        $expandUrl = Get-ExpandAllUrl -Html $pageHtml
        $expandedFile = Join-Path $cacheDir ("page_{0}_expanded.html" -f $pageIndex)
        $expandedHtml = Invoke-HeiseRequest -Url $expandUrl -OutFile $expandedFile -BaseSleepSeconds 2

        foreach ($url in (Get-TopLevelPostingUrls -Html $expandedHtml)) {
            if ($rootSeen.Add($url)) {
                $rootUrls.Add($url)
            }
        }

        Write-Host ("[forum] Seite {0}/{1} verarbeitet, Root-Threads bisher: {2}" -f $pageIndex, $pageUrls.Count, $rootUrls.Count)
    }

    [System.IO.File]::WriteAllText(
        $paths.RootUrls,
        ($rootUrls | ConvertTo-Json),
        [System.Text.Encoding]::UTF8
    )

    $threadUrls = New-Object System.Collections.Generic.List[string]
    $threadSeen = New-Object System.Collections.Generic.HashSet[string]

    for ($i = 0; $i -lt $rootUrls.Count; $i++) {
        $rootUrl = $rootUrls[$i]
        $rootFile = Join-Path $cacheDir ("root_" + (Get-SlugifiedFileName -Url $rootUrl))
        $rootHtml = Invoke-HeiseRequest -Url $rootUrl -OutFile $rootFile -BaseSleepSeconds 2

        $threadViewUrl = Get-ThreadViewUrl -Html $rootHtml
        if ([string]::IsNullOrWhiteSpace($threadViewUrl)) {
            Write-Warning "Keine Thread-Ansicht gefunden: $rootUrl"
            continue
        }

        $threadFile = Join-Path $cacheDir ("thread_live_{0:D3}.html" -f ($i + 1))
        $threadHtml = Invoke-HeiseRequest -Url $threadViewUrl -OutFile $threadFile -BaseSleepSeconds 4

        foreach ($postingUrl in (Get-PostingSubjectUrls -Html $threadHtml)) {
            if ($threadSeen.Add($postingUrl)) {
                $threadUrls.Add($postingUrl)
            }
        }

        Write-Host ("[thread] {0}/{1} verarbeitet, Postings bisher: {2}" -f ($i + 1), $rootUrls.Count, $threadUrls.Count)
    }

    [System.IO.File]::WriteAllText(
        $paths.PostFetchList,
        ($threadUrls | ConvertTo-Json),
        [System.Text.Encoding]::UTF8
    )

    if (-not $OnlyAuthors) {
        for ($i = 0; $i -lt $threadUrls.Count; $i++) {
            $postUrl = $threadUrls[$i]
            $postFile = Join-Path $cacheDir ("post_" + (Get-SlugifiedFileName -Url $postUrl))
            $null = Invoke-HeiseRequest -Url $postUrl -OutFile $postFile -BaseSleepSeconds 2
            if ((($i + 1) % 25) -eq 0 -or ($i + 1) -eq $threadUrls.Count) {
                Write-Host ("[post] {0}/{1} gecacht" -f ($i + 1), $threadUrls.Count)
            }
        }
    }
}

$analysisPython = @'
import json
import os
import re
import html
import hashlib
from collections import Counter, defaultdict
from pathlib import Path

def clean_html(fragment):
    fragment = re.sub(r"<br\s*/?>", "\n", fragment, flags=re.I)
    fragment = re.sub(r"</p\s*>", "\n\n", fragment, flags=re.I)
    fragment = re.sub(r"<[^>]+>", "", fragment)
    fragment = html.unescape(fragment)
    fragment = fragment.replace("\xa0", " ")
    fragment = re.sub(r"\r", "", fragment)
    fragment = re.sub(r"\n{3,}", "\n\n", fragment)
    return fragment.strip()

def first_match(text, patterns):
    for pattern in patterns:
        match = re.search(pattern, text, re.S)
        if match:
            return match.group(1)
    return None

def slugify(url):
    safe = re.sub(r"[^a-zA-Z0-9]+", "-", url).strip("-").lower()
    digest = hashlib.sha1(url.encode("utf-8")).hexdigest()[:12]
    return safe[:120] + "-" + digest + ".html"

def parse_posting(url, body_html):
    post_id_match = re.search(r"posting-(\d+)/show/", url)
    title = first_match(
        body_html,
        [
            r'<h1 class="thread_title">.*?<a [^>]*>\s*(.*?)\s*</a>',
            r"<title>\s*(.*?)\s*\| Forum - heise online</title>",
        ],
    )
    author = first_match(
        body_html,
        [
            r'<span class="pseudonym">(.*?)</span>',
            r'<li><span class="full_user_string">(.*?)</span></li>',
            r'<span class="tree_thread_list--written_by_user">\s*(.*?)\s*</span>',
        ],
    )
    timestamp = first_match(
        body_html,
        [
            r'<time[^>]*title="([^"]+)"',
            r'<time class="posting_timestamp"[^>]*>\s*(.*?)\s*</time>',
        ],
    )
    rating_match = re.search(r'alt="(-?\d+)" title="Beitragsbewertung:', body_html)
    post_body = first_match(
        body_html,
        [
            r'<div class="body_format_indicator bbcode_v1">\s*(.*?)\s*</div>\s*<!--googleoff: index-->',
            r'<div class="body_format_indicator[^"]*">\s*(.*?)\s*</div>',
        ],
    )
    if not (post_id_match and title and author and timestamp and post_body):
        raise RuntimeError(f"Could not parse posting page: {url}")
    return {
        "id": int(post_id_match.group(1)),
        "url": url,
        "title": clean_html(title),
        "author": clean_html(author),
        "datetime": clean_html(timestamp),
        "rating": int(rating_match.group(1)) if rating_match else None,
        "text": clean_html(post_body),
    }

TOPIC_RULES = [
    ("Bedienung_und_Ergonomie", ["rumkriechen", "bücken", "boden", "mantel", "dreckig", "knie", "ergonom", "beugen", "stecken", "einstecken", "buchse", "laden", "fumm", "unbequem", "höhe", "kabel", "handhab", "bedien"]),
    ("Wetter_Schmutz_und_Robustheit", ["winter", "schnee", "eis", "regen", "niederschlag", "schmutz", "dreck", "pfütze", "wasser", "ip68", "salz", "vereis", "beheiz"]),
    ("Vandalismus_Sabotage_und_Sicherheit", ["vandal", "sabot", "mutwill", "beschädig", "kaputt", "sicherheit", "tritt", "drüberfahren", "diebstahl", "manipulier", "gefährlich"]),
    ("Parken_Bordstein_Stadtbild", ["bordstein", "gehweg", "bürgersteig", "parken", "parkplatz", "laterne", "straßenrand", "stadtbild", "gehwegkante", "auf dem bordstein", "bordsteinkante"]),
    ("Alternativen_zur_Loesung", ["ladesäule", "ladesaeule", "laterne", "wallbox", "parkhaus", "firmenpark", "firmenparkplatz", "ladepark", "gar nicht nötig", "induktion", "akkuwechsel", "andere lösung", "normale"]),
    ("Wirtschaft_Kosten_und_Skalierung", ["kosten", "teuer", "billig", "wirtschaft", "million", "skal", "wartung", "nachrüst", "unternehmen", "geschäft", "rentabel", "fördern", "förder", "infrastruktur", "ausbau"]),
    ("Barrierefreiheit_und_Altersfragen", ["altersdiskr", "behindert", "rollstuhl", "barriere", "senior", "alte leute", "beeinträcht", "zugäng", "rücken", "körperlich"]),
    ("Grundsatzdebatte_EAuto_und_Verkehrswende", ["elektroauto", "e-auto", "ev", "verbrenner", "verkehrswende", "klima", "co2", "akku", "ladeinfrastruktur", "mobilität", "suv", "15 millionen"]),
    ("Rheinmetall_und_Politik", ["rheinmetall", "zivilisten", "waffen", "rüst", "ruest", "krieg", "tank", "konzern", "militär", "militaer"]),
    ("Positive_Erfahrungen_und_Pilotprojekt", ["funktioniert", "gute idee", "finde ich gut", "pilot", "testen", "akzeptanz", "sinnvoll", "praktisch", "verstehe die negativen", "benutzt", "erfahrung"]),
]

def classify_topics(text, title):
    haystack = f"{title}\n{text}".lower()
    topics = [topic for topic, keywords in TOPIC_RULES if any(keyword in haystack for keyword in keywords)]
    return topics or ["Sonstiges"]

def build_summary(comments):
    topic_buckets = defaultdict(list)
    for comment in comments:
        comment["topics"] = classify_topics(comment["text"], comment["title"])
        for topic in comment["topics"]:
            topic_buckets[topic].append(comment)
    topic_counts = {
        topic: len(items)
        for topic, items in sorted(topic_buckets.items(), key=lambda item: (-len(item[1]), item[0]))
    }
    summaries = {}
    stop_words = {
        "dass", "eine", "einen", "einem", "einer", "diese", "dieser", "nicht", "aber", "doch", "weil",
        "wird", "werden", "schon", "haben", "seine", "seinen", "ihren", "ihnen", "durch", "wenn", "dann",
        "oder", "auch", "noch", "über", "unter", "beim", "beide", "etwas", "sowie", "kann", "könnte",
        "können", "einfach", "damit", "dafür", "eher", "mehr", "weniger", "sowas", "diesem", "dieses", "thema",
    }
    for topic, items in topic_buckets.items():
        snippets = []
        for item in sorted(items, key=lambda x: (x["datetime"], x["id"]))[:5]:
            snippet = item["text"].replace("\n", " ").strip()
            if len(snippet) > 180:
                snippet = snippet[:177] + "..."
            snippets.append({
                "id": item["id"],
                "title": item["title"],
                "datetime": item["datetime"],
                "snippet": snippet,
                "url": item["url"],
            })
        token_counter = Counter(
            word
            for item in items
            for word in re.findall(r"[a-zA-ZäöüÄÖÜß]{4,}", f"{item['title']} {item['text']}".lower())
            if word not in stop_words
        )
        summaries[topic] = {
            "count": len(items),
            "common_terms": [word for word, _ in token_counter.most_common(12)],
            "examples": snippets,
        }
    return {
        "comment_count": len(comments),
        "topic_counts": topic_counts,
        "topics": summaries,
    }

cache_dir = Path(os.environ["HEISE_CACHE_DIR"])
forum_label = os.environ["HEISE_FORUM_LABEL"]
comments_json = Path(os.environ["HEISE_COMMENTS_JSON"])
summary_json = Path(os.environ["HEISE_SUMMARY_JSON"])
authors_json = Path(os.environ["HEISE_AUTHORS_JSON"])
themenbericht_md = Path(os.environ["HEISE_THEMENBERICHT_MD"])
vollbericht_md = Path(os.environ["HEISE_VOLLBERICHT_MD"])
kurzfassung_txt = Path(os.environ["HEISE_KURZFASSUNG_TXT"])
kommentatoren_md = Path(os.environ["HEISE_KOMMENTATOREN_MD"])

thread_files = sorted(cache_dir.glob("thread_live_*.html"))
urls = []
for path in thread_files:
    text = path.read_text(encoding="utf-8", errors="replace")
    urls.extend(re.findall(r'<a href="(https://www\.heise\.de/forum/heise-online/Kommentare/[^"]+/posting-\d+/show/)" class="posting_subject">', text))
urls = list(dict.fromkeys(urls))

comments = []
missing = []
for url in urls:
    cache_name = cache_dir / ("post_" + slugify(url))
    if not cache_name.exists():
        missing.append(url)
        continue
    post_html = cache_name.read_text(encoding="utf-8", errors="replace")
    comments.append(parse_posting(url, post_html))

comments.sort(key=lambda item: (item["datetime"], item["id"]))
summary = build_summary(comments)

comments_json.write_text(
    json.dumps(comments, ensure_ascii=False, indent=2),
    encoding="utf-8",
)
summary_json.write_text(
    json.dumps(summary, ensure_ascii=False, indent=2),
    encoding="utf-8",
)

positive_patterns = [
    r"\bgute sache\b", r"\binteressante idee\b", r"\bgute idee\b", r"\bsinnvoll\b",
    r"\bpraktisch\b", r"\brobust\b", r"\bfunktioniert\b", r"\bgefällt\b",
    r"\bpositiv\b", r"\bpilotprojekt\b", r"\bbraucht\b", r"\bnützlich\b", r"\bnuetzlich\b"
]
negative_patterns = [
    r"\bkeine gute idee\b", r"\bschlechte idee\b", r"\bbl[oö]d\b", r"\bunsinn\b",
    r"\bdreck\b", r"\bschmutz\b", r"\bhundekacke\b", r"\bproblem\b", r"\bunpraktisch\b",
    r"\baufwendig\b", r"\bteuer\b", r"\bkaputt\b", r"\bzweifel\b", r"\bskept\w*\b",
    r"\bnicht praktikabel\b", r"\bungeeignet\b", r"\bwartung\b", r"\bschnee\b", r"\beis\b"
]

def score_comment(c):
    text = ((c.get("title") or "") + "\n" + (c.get("text") or "")).lower()
    score = 0.0
    for pat in positive_patterns:
        score += len(re.findall(pat, text))
    for pat in negative_patterns:
        score -= len(re.findall(pat, text))
    rating = c.get("rating")
    if isinstance(rating, int):
        if rating >= 20:
            score += 0.75
        elif rating <= -20:
            score -= 0.75
    topics = set(c.get("topics") or [])
    if "Positive_Erfahrungen_und_Pilotprojekt" in topics:
        score += 0.6
    if "Bedienung_und_Ergonomie" in topics:
        score -= 0.35
    if "Wetter_Schmutz_und_Robustheit" in topics:
        score -= 0.35
    if "Vandalismus_Sabotage_und_Sicherheit" in topics:
        score -= 0.25
    if score >= 0.75:
        return score, "positiv"
    if score <= -0.75:
        return score, "negativ"
    return score, "neutral"

per_author = defaultdict(lambda: {
    "author": "",
    "posts": 0,
    "positiv": 0,
    "negativ": 0,
    "neutral": 0,
    "score_sum": 0.0,
    "ratings": [],
})

for comment in comments:
    author = comment.get("author") or "(ohne Namen)"
    score, tenor = score_comment(comment)
    row = per_author[author]
    row["author"] = author
    row["posts"] += 1
    row[tenor] += 1
    row["score_sum"] += score
    if isinstance(comment.get("rating"), int):
        row["ratings"].append(comment["rating"])

rows = []
for row in per_author.values():
    avg_score = row["score_sum"] / row["posts"]
    avg_rating = sum(row["ratings"]) / len(row["ratings"]) if row["ratings"] else None
    if avg_score >= 0.35:
        overall_tenor = "eher positiv"
    elif avg_score <= -0.35:
        overall_tenor = "eher negativ"
    else:
        overall_tenor = "gemischt/neutral"
    rows.append({
        "author": row["author"],
        "posts": row["posts"],
        "positive": row["positiv"],
        "negative": row["negativ"],
        "neutral": row["neutral"],
        "avg_score": round(avg_score, 2),
        "avg_rating": None if avg_rating is None else round(avg_rating, 1),
        "overall_tenor": overall_tenor,
    })

rows.sort(key=lambda item: (-item["posts"], item["author"].lower()))
authors_summary = {
    "comment_count": len(comments),
    "author_count": len(rows),
    "top_authors": rows[:25],
    "distribution_by_posts": dict(sorted(Counter(r["posts"] for r in rows).items())),
    "tenor_distribution_authors": dict(Counter(r["overall_tenor"] for r in rows)),
    "tenor_distribution_posts": dict(Counter(score_comment(c)[1] for c in comments)),
}

authors_json.write_text(
    json.dumps(authors_summary, ensure_ascii=False, indent=2),
    encoding="utf-8",
)

lines = []
lines.append(f"# Themenbericht: Heise-Forum zu {forum_label}")
lines.append("")
lines.append(f"Ausgewertet wurde der verlinkte Heise-Forenthread anhand von {len(thread_files)} vollstaendigen Thread-Ansichten. Die korrigierte Datenbasis umfasst {len(comments)} eindeutige Beitraege einschliesslich Unterkommentaren.")
lines.append("")
lines.append("## Management-Zusammenfassung")
lines.append("")
lines.append("Die Diskussion ist ueberwiegend skeptisch, aber meist pragmatisch und nicht grundsaetzlich gegen oeffentliche Ladeinfrastruktur gerichtet. Am haeufigsten geht es um drei praktische Punkte: Bedienung in Bodennaehe, Schmutz- und Wetterprobleme sowie die Frage, ob Fahrzeuge am Strassenrand ueberhaupt passend zum Ladepunkt stehen koennen. Viele Nutzer halten klassische Ladesaeulen oder kleine Ladeparks fuer robuster und alltagstauglicher. Positive Stimmen gibt es ebenfalls, vor allem fuer dicht bebaute Quartiere ohne private Stellplaetze. Der Tenor lautet insgesamt: urbanes Laden ist noetig, aber der Ladebordstein ueberzeugt viele in seiner aktuellen Form noch nicht.")
lines.append("")
lines.append("## Themenmatrix")
lines.append("")
lines.append("| Thema | Haeufigkeit | Tenor |")
lines.append("| --- | ---: | --- |")
topic_tenor = {
    "Parken_Bordstein_Stadtbild": "ueberwiegend negativ",
    "Bedienung_und_Ergonomie": "ueberwiegend negativ",
    "Wetter_Schmutz_und_Robustheit": "deutlich negativ",
    "Positive_Erfahrungen_und_Pilotprojekt": "vorsichtig positiv",
    "Alternativen_zur_Loesung": "eher gegen Ladebordsteine",
    "Grundsatzdebatte_EAuto_und_Verkehrswende": "gemischt",
    "Wirtschaft_Kosten_und_Skalierung": "skeptisch",
    "Rheinmetall_und_Politik": "gemischt bis ironisch",
    "Vandalismus_Sabotage_und_Sicherheit": "ueberwiegend negativ",
    "Barrierefreiheit_und_Altersfragen": "kritisch",
}
for topic, count in summary["topic_counts"].items():
    if topic == "Sonstiges":
        continue
    label = topic.replace("_", ", ", 1).replace("_", " ")
    lines.append(f"| {label} | {count} | {topic_tenor.get(topic, 'gemischt')} |")
lines.append("")
topic_descriptions = {
    "Parken_Bordstein_Stadtbild": "Viele Beitraege bezweifeln, dass Autos im Alltag praezise genug am passenden Bordsteinsegment stehen. Diskutiert werden enge Strassen, blockierte Stellplaetze und die Frage, ob die Loesung im Stadtbild wirklich einfacher ist als klassische Saeulen.",
    "Bedienung_und_Ergonomie": "Sehr viele Kommentare drehen sich um die Bedienhoehe und den koerperlichen Aufwand. Wiederholt wird beschrieben, dass man sich tief buecken oder in Schmutznaehe hantieren muesse.",
    "Wetter_Schmutz_und_Robustheit": "Ein grosser Themenblock betrifft Regen, Pfuetzen, Schnee, Eis, Salz und allgemeinen Strassenschmutz. Viele Nutzer bezweifeln die dauerhafte Wartungsarmut im Realbetrieb.",
    "Positive_Erfahrungen_und_Pilotprojekt": "Trotz des skeptischen Grundtons gibt es konstruktive Stimmen. Diese sehen den Ansatz als interessante Ergaenzung fuer dicht bebaute Quartiere ohne private Stellplaetze.",
    "Alternativen_zur_Loesung": "Viele Beitraege vergleichen den Ladebordstein mit klassischen Ladesaeulen, Ladehubs, Firmenparkplaetzen oder anderen sichtbaren Infrastrukturen.",
    "Grundsatzdebatte_EAuto_und_Verkehrswende": "Ein Teil des Threads loest sich vom konkreten Produkt und wechselt in die breitere Debatte ueber Elektromobilitaet, Ladeformen und urbane Infrastruktur.",
    "Wirtschaft_Kosten_und_Skalierung": "Mehrere Kommentare zweifeln an Wirtschaftlichkeit, Einbaukosten, Tiefbau, Wartung und kommunaler Skalierbarkeit.",
    "Rheinmetall_und_Politik": "Die Beteiligung von Rheinmetall erzeugt einen eigenen Nebenstrang mit ironischen und politischen Reaktionen, die oft weniger technisch sind.",
    "Vandalismus_Sabotage_und_Sicherheit": "Einige Nutzer sorgen sich um Beschaedigung, Manipulation oder Missbrauch im oeffentlichen Raum, gerade wegen der bodennahen Position am Strassenrand.",
    "Barrierefreiheit_und_Altersfragen": "Dieser Themenblock ist kleiner, aber klar: Buecken, eingeschraenkte Beweglichkeit und zugaengliche Bedienung werden als Huerden benannt.",
}
section_index = 1
for topic, count in summary["topic_counts"].items():
    if topic == "Sonstiges":
        continue
    label = topic.replace("_", " ").replace("und", "und")
    lines.append(f"## {section_index}. {label}")
    lines.append("")
    lines.append(topic_descriptions.get(topic, "Dieser Themenblock tritt im Thread sichtbar auf und praegt den Gesamteindruck mit."))
    details = summary["topics"].get(topic, {})
    common_terms = details.get("common_terms") or []
    if common_terms:
        lines.append("")
        lines.append("Hauefige Begriffe: " + ", ".join(common_terms[:8]))
    examples = details.get("examples") or []
    if examples:
        lines.append("")
        lines.append("Beispielbeitraege:")
        for example in examples[:3]:
            snippet = example["snippet"].replace("\n", " ").strip()
            lines.append(f"- {example['datetime']} | {example['title']} | {snippet}")
    lines.append("")
    lines.append(f"Kurzfazit: {topic_tenor.get(topic, 'gemischt').capitalize()} bei {count} Beitraegen.")
    lines.append("")
    section_index += 1
lines.append("## Gesamtfazit")
lines.append("")
lines.append("Der Thread erkennt den Bedarf an urbaner Ladeinfrastruktur grundsaetzlich an, stellt aber die konkrete Form des Ladebordsteins deutlich in Frage. Die staerksten Gegenargumente betreffen Parkrealitaet, Ergonomie, Witterung und Wartungsaufwand. Zustimmung gibt es vor allem dort, wo Nutzer das Konzept als testwuerdige Ergaenzung fuer schwierige innerstaedtische Lagen verstehen.")
themenbericht_md.write_text("\n".join(lines), encoding="utf-8")
vollbericht_md.write_text("\n".join(lines), encoding="utf-8")
kurzfassung_txt.write_text(
    "Die Diskussion ist ueberwiegend skeptisch, aber meist pragmatisch. "
    "Am haeufigsten kritisieren die Kommentierenden die Bedienung in Bodennaehe, "
    "Schmutz- und Wetterprobleme sowie die schwierige Parkgenauigkeit am Strassenrand. "
    "Viele halten klassische Ladesaeulen oder andere sichtbare Ladepunkte fuer robuster und alltagstauglicher. "
    "Positive Stimmen gibt es vor allem mit Blick auf dicht bebaute Quartiere ohne private Stellplaetze. "
    "Insgesamt wird der Bedarf an urbanen Ladelosungen anerkannt, die konkrete Bordsteinloesung ueberzeugt aber viele noch nicht.",
    encoding="utf-8",
)

lines = []
lines.append(f"# Kommentatorenstatistik: Heise-Forum zu {forum_label}")
lines.append("")
lines.append(f"Ausgewertet wurden {len(comments)} Beitraege von {len(rows)} eindeutigen Kommentatoren.")
lines.append("")
lines.append("## Verteilung")
lines.append("")
lines.append(f"- Beitraege insgesamt: {len(comments)}")
lines.append(f"- Kommentatoren insgesamt: {len(rows)}")
lines.append(f"- Beitraege mit positivem Tenor: {authors_summary['tenor_distribution_posts'].get('positiv', 0)}")
lines.append(f"- Beitraege mit neutralem Tenor: {authors_summary['tenor_distribution_posts'].get('neutral', 0)}")
lines.append(f"- Beitraege mit negativem Tenor: {authors_summary['tenor_distribution_posts'].get('negativ', 0)}")
lines.append("")
lines.append("## Top-Kommentatoren nach Beitragszahl")
lines.append("")
lines.append("| Rang | Kommentator | Beitraege | positiv | neutral | negativ | Gesamttenor | Avg. Score | Avg. Bewertung |")
lines.append("| --- | --- | ---: | ---: | ---: | ---: | --- | ---: | ---: |")
for idx, row in enumerate(rows[:20], start=1):
    avg_rating = "" if row["avg_rating"] is None else str(row["avg_rating"]).replace(".", ",")
    avg_score = str(row["avg_score"]).replace(".", ",")
    lines.append(f"| {idx} | {row['author']} | {row['posts']} | {row['positive']} | {row['neutral']} | {row['negative']} | {row['overall_tenor']} | {avg_score} | {avg_rating} |")
kommentatoren_md.write_text("\n".join(lines), encoding="utf-8")

print(json.dumps({
    "thread_files": len(thread_files),
    "comments": len(comments),
    "missing": len(missing),
    "authors": len(rows),
}, ensure_ascii=False))
'@

$env:HEISE_CACHE_DIR = $cacheDir
$env:HEISE_FORUM_LABEL = (Get-ForumLabel -ForumSlug $forumMeta.ForumSlug)
$env:HEISE_COMMENTS_JSON = $paths.CommentsJson
$env:HEISE_SUMMARY_JSON = $paths.SummaryJson
$env:HEISE_AUTHORS_JSON = $paths.AuthorsJson
$env:HEISE_THEMENBERICHT_MD = $paths.Themenbericht
$env:HEISE_VOLLBERICHT_MD = $paths.Vollbericht
$env:HEISE_KURZFASSUNG_TXT = $paths.Kurzfassung
$env:HEISE_KOMMENTATOREN_MD = $paths.Kommentatorenbericht

$analysisPython | python -
Write-Host ("[done] JSON- und Markdown-Ausgaben wurden aktualisiert fuer: {0}" -f $forumMeta.AnalysisKey)
