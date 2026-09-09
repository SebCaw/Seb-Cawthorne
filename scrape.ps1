# Cloud scraper for the Degree Apprenticeship Tracker.
# Runs on GitHub Actions (PowerShell Core, Linux) twice a day.
# Reads & rewrites the three sector data files (sales / finance / consulting).
# Source: GOV.UK "Find an Apprenticeship" only. Non-technical (sales/finance/consulting) roles.

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

# Test path: set TEST_NOTIFICATION=true (via the workflow_dispatch input) to send one
# dummy message and exit before scraping. Without this there is no way to distinguish
# "notifications work but nothing is new" from "notifications are broken" - which
# matters most right before the September season opens.
if ($env:TEST_NOTIFICATION -eq 'true') {
  # ALERT_TEXT (workflow_dispatch input) sends that text instead of the dummy
  # message - use it to push a real alert by hand when something opens that the
  # scraper cannot see, e.g. a Workday-hosted vacancy behind JavaScript.
  if ($env:ALERT_TEXT) {
    $testMsg = $env:ALERT_TEXT
  } else {
    $testMsg = "Test notification from the apprenticeship tracker." + [char]10 + [char]10 + "If you can read this, notifications are wired up correctly." + [char]10 + "Sent: " + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + " UTC"
  }
  $sentOk = $false
  if ($env:TELEGRAM_BOT_TOKEN -and $env:TELEGRAM_CHAT_ID) {
    try {
      $tBody = @{ chat_id = $env:TELEGRAM_CHAT_ID; text = $testMsg } | ConvertTo-Json
      Invoke-RestMethod -Uri "https://api.telegram.org/bot$($env:TELEGRAM_BOT_TOKEN)/sendMessage" -Method Post -ContentType 'application/json; charset=utf-8' -Body $tBody | Out-Null
      Write-Host "Test notification sent via Telegram."
      $sentOk = $true
    } catch { Write-Host "Telegram test FAILED: $_" }
  } else {
    Write-Host "Telegram not configured - TELEGRAM_BOT_TOKEN and/or TELEGRAM_CHAT_ID is missing."
  }
  if ($env:NTFY_TOPIC) {
    try {
      Invoke-RestMethod -Uri "https://ntfy.sh/$($env:NTFY_TOPIC)" -Method Post -Body $testMsg -Headers @{ Title = "Apprenticeship tracker test" } | Out-Null
      Write-Host "Test notification sent via ntfy."
      $sentOk = $true
    } catch { Write-Host "ntfy test FAILED: $_" }
  }
  if (-not $sentOk) { Write-Host "No test notification could be sent - check the repo secrets."; exit 1 }
  exit 0
}


$TECH = 'software|web develop|developer|programmer|data scien|data engineer|data analy|\bdata\b|cyber|\bcloud|devops|infrastructure|\bnetwork|enginee|structural|mechanical|electrical|chemical|aerospace|laborator|quantity survey|architect\b'
# Always exclude these, however the rest of the title reads (core technical / data / cyber roles).
$HARD_TECH = 'software|web develop|\bdeveloper\b|programmer|data scien|data engineer|data analy|\bdata\b|cyber|devops|network engineer|infrastructure engineer|cloud engineer'
# A clear business / commercial / finance / consulting signal — keep the role however it is worded.
$BIZ_SIGNAL = 'sales|account|business|revenue|go-to-market|relationship|commercial|client|partnership|customer|\bfinanc|\bbank|invest|wealth|capital market|trading|broker|portfolio|asset manage|underwrit|actuar|insurance|pension|treasury|lending|mortgage|\bcredit|\brisk|complian|consult|advis|strateg|transformation|audit|assur|\btax\b|market|procure|supply chain|buyer|merchandis|negotiat|channel|corporate|human resource|\bhr\b|\bmanagement\b'
$BIG  = 'microsoft|ibm|salesforce|amazon|oracle|\bsap\b|hsbc|barclays|goldman|jpmorgan|j\.p\. morgan|jp morgan|pwc|pricewaterhouse|deloitte|kpmg|ernst|\bey\b|accenture|capgemini|cognizant|infosys|tata|\btcs\b|bt group|\bbt\b|vodafone|\bsky\b|virgin media|\bo2\b|lloyds|natwest|santander|nationwide|standard chartered|aviva|legal & general|prudential|\baxa\b|allianz|zurich|morgan stanley|\bciti\b|bank of america|merrill|\bubs\b|deutsche bank|nomura|bnp paribas|schroders|fidelity|unilever|procter|nestle|coca-cola|pepsico|diageo|\bmars\b|johnson|\bgsk\b|glaxo|astrazeneca|pfizer|siemens|bosch|cisco|\bdell\b|\bhp\b|hewlett|intel|google|meta|apple|adobe|tesco|sainsbury|\basda\b|marks & spencer|m&s|john lewis|boots|\bbp\b|shell|centrica|national grid|rolls-royce|\bbae\b|airbus|jaguar|land rover|nissan|toyota|\bford\b|\bbmw\b|mercedes|volkswagen|network rail|royal mail|\bdhl\b|fedex|\bups\b|american express|\bamex\b|visa|mastercard|paypal|\bsage\b|softcat|computacenter'

function Parse-CloseDate([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $s2 = $s -replace ' at .*$','' -replace '^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s+',''
  try { ([datetime]::Parse($s2.Trim(), [System.Globalization.CultureInfo]::GetCultureInfo('en-GB'))).ToString('yyyy-MM-dd') } catch { $null }
}
function Status-From([string]$iso) {
  if ([string]::IsNullOrWhiteSpace($iso)) { return 'open' }
  try { $d = ([datetime]$iso - (Get-Date)).TotalDays; if ($d -lt 0) { 'closed' } elseif ($d -le 14) { 'closing_soon' } else { 'open' } } catch { 'open' }
}
function Is-Technical([string]$title, [string]$course) {
  $t = ("$title $course").ToLower()
  if ($t -match $BIZ_SIGNAL) { return $false }   # any sales/finance/consulting signal wins, even with a tech word (e.g. "Software Sales")
  if ($t -match $HARD_TECH) { return $true }     # otherwise pure tech / data / cyber -> drop
  return $t -match $TECH                          # broad technical screen for anything left
}
function Category-From([string]$title) {
  $t = $title.ToLower()
  if ($t -match 'sales|business develop|account|customer|relationship|commercial|partnership') { 'Sales' }
  elseif ($t -match 'audit|assurance|\btax\b') { 'Audit' }
  elseif ($t -match 'finance|financ|bank|invest|wealth') { 'Finance' }
  elseif ($t -match 'consult|advis|strateg') { 'Consulting' }
  elseif ($t -match 'market') { 'Marketing' }
  else { 'Business' }
}
# Route a broad-search "interest" listing to the sector page it belongs on.
function Sector-ForCategory([string]$cat) {
  switch ($cat) {
    'Finance'    { 'finance' }
    'Audit'      { 'consulting' }
    'Consulting' { 'consulting' }
    default      { 'sales' }   # Sales, Marketing, Business
  }
}
function Get-Vacancies([string]$html) {
  $out = @()
  $titles = [regex]::Matches($html, '<span id="(VAC\d+)-vacancy-title">([^<]*)</span>')
  for ($k = 0; $k -lt $titles.Count; $k++) {
    $id = $titles[$k].Groups[1].Value; $title = ($titles[$k].Groups[2].Value).Trim()
    $bStart = $titles[$k].Index
    $bEnd = if ($k + 1 -lt $titles.Count) { $titles[$k+1].Index } else { $html.Length }
    $block = $html.Substring($bStart, $bEnd - $bStart)
    $employer=''; $location=''; $course=''; $closeRaw=''
    $em = [regex]::Match($block, '<p class="govuk-body govuk-!-margin-bottom-0">([^<]+)</p>'); if ($em.Success) { $employer = $em.Groups[1].Value.Trim() }
    $lm = [regex]::Match($block, 'das-!-color-dark-grey">\s*([^<]+?)\s*</p>'); if ($lm.Success) { $location = ($lm.Groups[1].Value.Trim() -replace '\s+',' ') }
    $cm = [regex]::Match($block, '<b>Training course</b>\s*([^<]+)</p>'); if ($cm.Success) { $course = $cm.Groups[1].Value.Trim() }
    $clm = [regex]::Match($block, 'Closes[^(<]*\(([^)]+)\)'); if ($clm.Success) { $closeRaw = $clm.Groups[1].Value.Trim() }
    $out += [pscustomobject]@{ id=$id; title=$title; employer=$employer; location=$location; course=$course; closeISO=(Parse-CloseDate $closeRaw) }
  }
  return $out
}
function Fetch-Html([string]$url) {
  try { (Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }).Content } catch { $null }
}

# --- vacancy detail: pull a short overview + skill keywords from a role's own GOV.UK page ---
function Html-Decode([string]$s) {
  if (-not $s) { return '' }
  $s = $s -replace '&nbsp;',' ' -replace '&amp;','&' -replace '&#39;',"'" -replace '&rsquo;',"'" -replace '&lsquo;',"'" -replace '&quot;','"' -replace '&ldquo;','"' -replace '&rdquo;','"' -replace '&bull;','' -replace '&pound;',([char]0xA3) -replace '&ndash;','-' -replace '&mdash;','-' -replace '&hellip;','...'
  return $s
}
function Strip-Tags([string]$s) {
  if (-not $s) { return '' }
  $s = [regex]::Replace($s, '(?s)<[^>]+>', ' ')
  $s = Html-Decode $s
  $s = ($s -replace '\s+',' ').Trim()
  return $s
}
function Get-VacancyDetail([string]$id) {
  $url = "https://www.findapprenticeship.service.gov.uk/apprenticeship/$id"
  $overview = ''; $keywords = @()
  $html = Fetch-Html $url
  if ($html) {
    $m = [regex]::Match($html, '(?s)<h2[^>]*>Summary</h2>(.*?)<dl')
    if ($m.Success) { $overview = Strip-Tags $m.Groups[1].Value }
    if (-not $overview) {
      $m2 = [regex]::Match($html, "(?s)What you'll do at work</h3>(.*?)</h3>")
      if ($m2.Success) { $overview = Strip-Tags $m2.Groups[1].Value }
    }
    $sm = [regex]::Match($html, '(?s)<h3[^>]*>\s*Skills\s*</h3>(.*?)</ul>')
    if ($sm.Success) {
      foreach ($li in [regex]::Matches($sm.Groups[1].Value, '(?s)<li[^>]*>(.*?)</li>')) {
        $k = ((Strip-Tags $li.Groups[1].Value) -replace '\s*[Ss]kills?$','').Trim()
        if ($k) { $keywords += $k }
      }
    }
  }
  if ($overview.Length -gt 320) { $overview = ($overview.Substring(0,320) -replace '\s+\S*$','') + [char]0x2026 }
  return [pscustomobject]@{ url=$url; overview=$overview; keywords=@($keywords | Select-Object -First 8) }
}



# ---- customer-facing classification -------------------------------------------------
# Mirrors the customerFacing tags in data.json so scraped roles are classified the
# same way as seeded ones. core = owns a customer or a number, client = advisory but
# client-facing, internal = back office.
function CustomerFacing-From([string]$title, [string]$course) {
  $t = ("$title $course").ToLower()
  if ($t -match 'sales|account manage|business develop|customer success|customer solution|client engagement|relationship manage|commercial bank|corporate bank|customer service|channel|partnership') { return 'core' }
  if ($t -match 'consult|advis|wealth|client|tax|strateg') { return 'client' }
  return 'internal'
}

# ---- employer careers-site check ----------------------------------------------------
# GOV.UK carries almost none of these employers - the register showed 46 Level 6
# apprenticeships nationally and zero matching "sales". This checks each employer's
# own careers page for a change in apprenticeship signal, so an opening is caught even
# when GOV.UK never lists it. Cheap, no browser, no Claude credits.
# Workday-hosted vacancies.
#
# Barclays posted 14 apprenticeships on 9 Sept 2026 and this scraper saw none of
# them: the page it reads is a marketing page, while the vacancies live in
# Workday behind JavaScript. The tracker said "0 open" while the role Seb went on
# to apply for was live. That is the worst possible failure for a tracker.
#
# Workday exposes a plain JSON endpoint that needs no browser:
#   POST /wday/cxs/<tenant>/<site>/jobs  {"limit":20,"offset":0,"searchText":"..."}
# Configure per company in data.json as:
#   "ats": { "type":"workday", "host":"...", "tenant":"...", "site":"...", "search":"..." }
function Get-WorkdayRoles($co) {
  if (-not ($co.PSObject.Properties.Name -contains 'ats') -or -not $co.ats) { return $null }
  if ($co.ats.type -ne 'workday') { return $null }
  $api = "https://" + $co.ats.host + "/wday/cxs/" + $co.ats.tenant + "/" + $co.ats.site + "/jobs"
  $body = @{ appliedFacets = @{}; limit = 20; offset = 0; searchText = $co.ats.search } | ConvertTo-Json -Depth 4
  # Record the outcome on the company so a failure is visible in data.json.
  # GitHub Actions runs from a datacentre IP and Workday sits behind Akamai, so
  # a call that works from a home connection can still be blocked here - and the
  # Actions log is awkward to read after the fact.
  $diag = [pscustomobject]@{
    checkedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ssZ')
    ok = $false; total = 0; kept = 0; error = ''
  }
  try {
    $r = Invoke-RestMethod -Uri $api -Method Post -ContentType 'application/json' -Body $body -TimeoutSec 25 -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36'; 'Accept' = 'application/json' }
    $diag.ok = $true
    $diag.total = [int]$r.total
  } catch {
    $diag.error = ($_.Exception.Message)
    $co | Add-Member -NotePropertyName atsCheck -NotePropertyValue $diag -Force
    Write-Host ("Workday check failed for " + $co.name + ": " + $_)
    return $null
  }
  $out = @()
  foreach ($p in $r.jobPostings) {
    # Degree-level only. "Higher" and plain apprenticeships are lower level, and
    # a title without "Degree" is a different, lower-level scheme - at Barclays
    # the two Manchester Business Banking roles differ by exactly that word.
    if ($p.title -notmatch '(?i)degree') { continue }
    # Optional per-company patterns so the tracker shows only the roles Seb would
    # actually apply for. Without this Barclays returns Technology Developer,
    # Technology Analyst and an Operations Analyst that is really a Data Scientist
    # qualification - all degree-level, none of them what he wants.
    if ($co.ats.PSObject.Properties.Name -contains 'exclude' -and $co.ats.exclude) {
      if ($p.title -match $co.ats.exclude) { continue }
    }
    if ($co.ats.PSObject.Properties.Name -contains 'include' -and $co.ats.include) {
      if ($p.title -notmatch $co.ats.include) { continue }
    }
    $out += [pscustomobject]@{
      title    = $p.title
      location = $p.locationsText
      postedOn = $p.postedOn
      url      = "https://" + $co.ats.host + "/" + $co.ats.site + $p.externalPath
    }
  }
  $diag.kept = $out.Count
  $co | Add-Member -NotePropertyName atsCheck -NotePropertyValue $diag -Force
  return ,$out
}

function Check-EmployerSite($co) {
  $result = [pscustomobject]@{
    checkedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ssZ')
    reachable = $false
    httpStatus = $null
    mentions  = 0
    openSignal = $false
    note = ''
  }
  if (-not $co.careerUrl) { $result.note = 'no careerUrl'; return $result }
  try {
    $resp = Invoke-WebRequest -Uri $co.careerUrl -UseBasicParsing -TimeoutSec 25 -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36' }
    $result.reachable = $true
    $result.httpStatus = [int]$resp.StatusCode
    $html = $resp.Content
    $result.mentions = ([regex]::Matches($html, 'apprentice', 'IgnoreCase')).Count
    # An explicit "applications are open" style phrase is a stronger signal than a
    # raw mention count, which barely moves on a static marketing page.
    if ([regex]::IsMatch($html, 'applications?s+(ares+)?(nows+)?open|applys+now|nows+recruiting|applications?s+open(ing)?s+', 'IgnoreCase')) {
      $result.openSignal = $true
    }
    if ($result.mentions -eq 0) { $result.note = 'page loaded but no apprenticeship text - likely JavaScript-rendered' }
  } catch {
    $code = $null
    try { $code = [int]$_.Exception.Response.StatusCode } catch {}
    $result.httpStatus = $code
    if ($code -eq 403) { $result.note = 'blocked - check manually in the browser' }
    elseif ($code -eq 404) { $result.note = 'careerUrl is dead - needs updating' }
    else { $result.note = 'unreachable' }
  }
  return $result
}

# ---- the three sector datasets ----
$sectorFiles = [ordered]@{ finance = 'data.json' }
$today = (Get-Date).ToString('yyyy-MM-dd')

$data = @{}
foreach ($s in $sectorFiles.Keys) {
  $p = Join-Path $root $sectorFiles[$s]
  $data[$s] = [System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
}

$errors = 0
$liveItems = @()       # {id, company, title, sector} for everything found this run (for notifications)
$coreIdsAll = @{}      # vacancy ids matched to a core company (excluded from interest)
$coreLive = 0
$siteSignals = @()   # employers whose own careers page started reading as open

# ---- core companies: scrape each sector's tracked employers ----
foreach ($s in $sectorFiles.Keys) {
  foreach ($co in $data[$s].companies) {

    # Check the employer's own careers page alongside GOV.UK.
    $prevSignal = $false
    if ($co.PSObject.Properties.Name -contains 'siteCheck' -and $co.siteCheck) { $prevSignal = [bool]$co.siteCheck.openSignal }
    $sc = Check-EmployerSite $co
    $co | Add-Member -NotePropertyName siteCheck -NotePropertyValue $sc -Force
    if ($sc.openSignal -and -not $prevSignal) {
      $siteSignals += [pscustomobject]@{ company = $co.name; url = $co.careerUrl; sector = $s }
      Write-Host ("Careers-site signal: " + $co.name + " now reads as open - " + $co.careerUrl)
    }
    # Workday roles are real, dated vacancies, so they take priority over the
    # seeded programme list. When they are present the card shows what is
    # actually open rather than a curated guess.
    $wd = Get-WorkdayRoles $co
    if ($wd -and $wd.Count -gt 0) {
      $wdProgs = @()
      foreach ($w in $wd) {
        $wdProgs += [pscustomobject]@{
          id             = ($co.id + '_wd_' + ([regex]::Replace($w.title.ToLower(), '[^a-z0-9]+', '_')).Trim('_'))
          name           = $w.title
          standard       = 'Level 6'
          location       = $w.location
          salary         = $null
          duration       = $null
          customerFacing = 'core'
          nameConfidence = 'official'
          sourceUrl      = $w.url
          status         = 'open'
          closingDate    = $null
          applyUrl       = $w.url
          govVacancyId   = $null
          firstSeen      = (Get-Date).ToString('yyyy-MM-dd')
          lastSeen       = (Get-Date).ToString('yyyy-MM-dd')
        }
      }
      $co.programs = $wdProgs
      $coreLive += $wdProgs.Count
      Write-Host ("Workday: " + $co.name + " has " + $wdProgs.Count + " live degree apprenticeships")
      foreach ($w in $wd) {
        $liveItems += [pscustomobject]@{ id = $w.url; company = $co.name; title = $w.title; sector = $s; close = $null }
      }
    }

    $html = Fetch-Html $co.govSearchUrl
    if ($null -eq $html) { $errors++; continue }
    $matched = @()
    foreach ($v in (Get-Vacancies $html)) {
      $empOk = $false
      foreach ($variant in $co.employerNameVariants) {
        if ($v.employer -and [regex]::IsMatch($v.employer, '\b' + [regex]::Escape($variant) + '\b', 'IgnoreCase')) { $empOk = $true; break }
      }
      if (-not $empOk) { continue }
      if (Is-Technical $v.title $v.course) { continue }
      $matched += $v
    }
    if ($matched.Count -gt 0) {
      $progs = @()
      foreach ($v in ($matched | Sort-Object { if ($_.closeISO) { [datetime]$_.closeISO } else { [datetime]::MaxValue } })) {
        $coreIdsAll[$v.id] = $true; $coreLive++
        $liveItems += [pscustomobject]@{ id=$v.id; company=$co.name; title=$v.title; sector=$s; location=$v.location; close=$v.closeISO }
        $progs += [pscustomobject]@{
          id=$v.id; name=$v.title; standard="Level 6"; location=$v.location; salary=$null; duration=""
          status=(Status-From $v.closeISO); closingDate=$v.closeISO
          applyUrl=("https://www.findapprenticeship.service.gov.uk/apprenticeship/" + $v.id)
          govVacancyId=$v.id; firstSeen=$today; lastSeen=$today
          customerFacing=(CustomerFacing-From $v.title $v.course)
          nameConfidence='official'   # straight from a live GOV.UK posting
          sourceUrl=("https://www.findapprenticeship.service.gov.uk/apprenticeship/" + $v.id)
        }
      }
      $co.programs = $progs
    }
    # else: leave the company's existing (seeded / pre-season) programs untouched.
    # Curated programmes carry official names researched from the employer's own site;
    # GOV.UK listings only replace them when GOV.UK actually has something.
  }
}

# ---- interest: broad GOV.UK searches, routed to the matching sector page ----
$terms = @('sales','business+development','account+management','finance','banking','consulting','audit','business','marketing','commercial')
$interestBySector = @{ finance = @() }
$seenInterest = @{}
foreach ($term in $terms) {
  $html = Fetch-Html "https://www.findapprenticeship.service.gov.uk/apprenticeships?searchTerm=$term&levelIds=6&distanceType=England&sort=AgeDesc"
  if ($null -eq $html) { $errors++; continue }
  foreach ($v in (Get-Vacancies $html)) {
    if ($seenInterest.ContainsKey($v.id) -or $coreIdsAll.ContainsKey($v.id)) { continue }
    if (Is-Technical $v.title $v.course) { continue }
    if (-not $v.employer) { continue }
    if ($v.employer.ToLower() -notmatch $BIG) { continue }
    if ($v.closeISO -and ([datetime]$v.closeISO -lt (Get-Date))) { continue }
    $seenInterest[$v.id] = $true
    $cat = Category-From $v.title
    $sec = Sector-ForCategory $cat; if ($sec -ne 'finance') { continue }
    $liveItems += [pscustomobject]@{ id=$v.id; company=$v.employer; title=$v.title; sector=$sec; location=$v.location; close=$v.closeISO }
    $interestBySector[$sec] += [pscustomobject]@{
      id=$v.id; company=$v.employer; title=$v.title; standard="Level 6"; location=$v.location; salary=$null
      category=$cat; closingDate=$v.closeISO
      applyUrl=("https://www.findapprenticeship.service.gov.uk/apprenticeship/" + $v.id); firstSeen=$today
    }
  }
}

# ---- per-sector: update meta, compute what's new, write the file ----
$allNew = @()   # {id, sector}
foreach ($s in $sectorFiles.Keys) {
  $d = $data[$s]
  $d.interestListings = @($interestBySector[$s])

  $sectorIds = @()
  foreach ($co in $d.companies) { foreach ($p in $co.programs) { if ($p.govVacancyId) { $sectorIds += $p.govVacancyId } } }
  foreach ($it in $interestBySector[$s]) { $sectorIds += $it.id }
  $sectorIds = @($sectorIds | Select-Object -Unique)

  $seen = @(); if ($d.meta.seenVacancyIds) { $seen = @($d.meta.seenVacancyIds) }
  $new = @($sectorIds | Where-Object { $seen -notcontains $_ })
  foreach ($n in $new) { $allNew += [pscustomobject]@{ id=$n; sector=$s } }

  $d.meta.lastUpdated = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
  $d.meta.seenVacancyIds = @($seen + $sectorIds | Select-Object -Unique)
  $d.meta.newSinceLastScrape = @($new)

  $json = $d | ConvertTo-Json -Depth 12
  [System.IO.File]::WriteAllText((Join-Path $root $sectorFiles[$s]), $json, (New-Object System.Text.UTF8Encoding($false)))
}

Write-Host "Done. coreLive=$coreLive new=$($allNew.Count) errors=$errors"

# --- Notifications - one message per genuinely NEW opening this run ---
# Each new role gets its own message: title, employer, location + closing date, a short
# overview and skill keywords pulled from the role's own GOV.UK page, and an apply link.
$tgToken = $env:TELEGRAM_BOT_TOKEN; $tgChat = $env:TELEGRAM_CHAT_ID
$ntfyTopic = $env:NTFY_TOPIC
if ($allNew.Count -gt 0) {
  $anySent = $false
  foreach ($n in $allNew) {
    $it = $liveItems | Where-Object { $_.id -eq $n.id -and $_.sector -eq $n.sector } | Select-Object -First 1
    if (-not $it) { continue }

    $detail = Get-VacancyDetail $it.id

    $closeTxt = ''
    if ($it.close) { try { $closeTxt = ([datetime]$it.close).ToString('d MMM yyyy') } catch { $closeTxt = [string]$it.close } }

    $parts = @()
    $parts += ("New Level 6 apprenticeship - " + $n.sector.ToUpper())
    $parts += ""
    $parts += $it.title
    $parts += $it.company
    $meta = @()
    if ($it.location) { $meta += $it.location }
    if ($closeTxt)    { $meta += ("Closes " + $closeTxt) }
    if ($meta.Count -gt 0) { $parts += ($meta -join " | ") }
    if ($detail.overview) { $parts += ""; $parts += ("Overview: " + $detail.overview) }
    if ($detail.keywords -and $detail.keywords.Count -gt 0) { $parts += ""; $parts += ("Keywords: " + ($detail.keywords -join ", ")) }
    $parts += ""
    $parts += ("Apply: " + $detail.url)
    $msg = ($parts -join "`n")

    $sent = $false
    if ($tgToken -and $tgChat) {
      try {
        $body = @{ chat_id = $tgChat; text = $msg; disable_web_page_preview = $false } | ConvertTo-Json
        Invoke-RestMethod -Uri "https://api.telegram.org/bot$tgToken/sendMessage" -Method Post -ContentType 'application/json; charset=utf-8' -Body $body | Out-Null
        $sent = $true; $anySent = $true
      } catch { Write-Host "Telegram notification failed for $($it.id): $_" }
    }
    if ($ntfyTopic) {
      try {
        Invoke-RestMethod -Uri "https://ntfy.sh/$ntfyTopic" -Method Post -Body $msg -Headers @{ Title = ($it.company + " - " + $it.title); Click = $detail.url } | Out-Null
        $sent = $true; $anySent = $true
      } catch { Write-Host "ntfy notification failed for $($it.id): $_" }
    }
    if ($sent) { Write-Host ("Notified: {0} - {1}" -f $it.company, $it.title) }
    Start-Sleep -Milliseconds 400
  }
  if (-not $anySent) { Write-Host "New roles found but no notifier configured (set TELEGRAM_BOT_TOKEN+TELEGRAM_CHAT_ID, or NTFY_TOPIC)." }
} else {
  Write-Host "No new roles this run; no notification sent."
}
