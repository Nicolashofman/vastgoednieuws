# Vastgoednieuws: haalt recente artikels op uit RSS-feeds en maakt er één overzichtspagina van.
# Gebruik: rechtsklik > "Met PowerShell uitvoeren", of dubbelklik op Vastgoednieuws.bat
param([int]$Dagen = 7, [switch]$GeenBrowser)

$ErrorActionPreference = 'Continue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$map = Split-Path -Parent $MyInvocation.MyCommand.Path

# ---- Bronnen: pas deze lijst gerust aan ----
$bronnen = @(
    @{ naam = 'Vastgoedmarkt';        url = 'https://www.vastgoedmarkt.nl/rss.xml' },
    @{ naam = 'Vastgoedactueel';      url = 'https://www.vastgoedactueel.nl/feed/' },
    @{ naam = 'Google News (BE)';     url = 'https://news.google.com/rss/search?q=vastgoed+when:7d&hl=nl&gl=BE&ceid=BE:nl'; google = $true },
    @{ naam = 'Google News (NL)';     url = 'https://news.google.com/rss/search?q=vastgoed+when:7d&hl=nl&gl=NL&ceid=NL:nl'; google = $true },
    @{ naam = 'Google News (immo BE)'; url = 'https://news.google.com/rss/search?q=immosector+OR+vastgoedmarkt+OR+woningmarkt+when:7d&hl=nl&gl=BE&ceid=BE:nl'; google = $true },
    # --- Regio Kortrijk: zoekfeeds ---
    @{ naam = 'Google News (Kortrijk)'; url = 'https://news.google.com/rss/search?q=Kortrijk+(vastgoed+OR+woningen+OR+appartementen+OR+bouwproject+OR+verkaveling+OR+immo+OR+stadsontwikkeling)+when:14d&hl=nl&gl=BE&ceid=BE:nl'; google = $true; filter = $true; forceregio = $true },
    @{ naam = 'Google News (omgeving)'; url = 'https://news.google.com/rss/search?q=(Harelbeke+OR+Kuurne+OR+Wevelgem+OR+Menen+OR+Izegem+OR+Waregem+OR+Zwevegem+OR+Avelgem+OR+Anzegem)+(vastgoed+OR+woningen+OR+appartementen+OR+bouwproject+OR+verkaveling)+when:14d&hl=nl&gl=BE&ceid=BE:nl'; google = $true; filter = $true },
    # --- Regio Kortrijk: regionale kranten en stad (algemeen nieuws, daarom gefilterd op vastgoedwoorden) ---
    @{ naam = 'Krant van West-Vlaanderen'; url = 'https://www.kw.be/rss'; filter = $true },
    @{ naam = 'HLN Kortrijk'; url = 'https://www.hln.be/kortrijk/rss.xml'; filter = $true; forceregio = $true },
    @{ naam = 'Stad Kortrijk'; url = 'https://www.kortrijk.be/rss.xml'; filter = $true; forceregio = $true }
)

# Woorden die een artikel vastgoed-gerelateerd maken (enkel gebruikt voor de algemene regionale feeds)
$vastgoedWoorden = 'vastgoed|woning|woningen|appartement|verkaveling|bouwgrond|bouwproject|bouwvergunning|omgevingsvergunning|renovatie|projectontwikkel|immo|huurder|verhuur|koopwoning|sociale woningen|leegstand|stadsontwikkeling|herbestemming|woonproject|woonwijk|woonzorg|kantoorgebouw|bedrijfsgebouw|winkelpand|nieuwbouw|ruimtelijk|masterplan|wooncomplex|sloop|site '
# Plaatsen die als "jouw regio" tellen
$regioWoorden = '\b(Kortrijk|Courtrai|Kuurne|Harelbeke|Wevelgem|Menen|Izegem|Waregem|Zwevegem|Anzegem|Deerlijk|Avelgem|Lendelede|Heule|Bissegem|Rollegem|Aalbeke|Wielsbeke|Ingelmunster|Lauwe|Bavikhove|Hulste|Kooigem|Bellegem|Vichte|Leiestreek|Zuid-West-Vlaanderen|Kortrijk Xpo)\b'

function Get-Tekst($s) {
    if (-not $s) { return '' }
    $t = [System.Net.WebUtility]::HtmlDecode(($s -replace '<[^>]+>', ' '))
    return (($t -replace '\s+', ' ').Trim())
}

function Get-Node($n) {
    # Een XML-element zonder attributen komt als string binnen, anders als element
    if ($null -eq $n) { return '' }
    if ($n -is [string]) { return $n }
    return $n.InnerText
}

$grens =[DateTimeOffset]::UtcNow.AddDays(-$Dagen)
$artikels = New-Object System.Collections.ArrayList
$status = @()

foreach ($b in $bronnen) {
    try {
        $r = Invoke-WebRequest -Uri $b.url -UseBasicParsing -TimeoutSec 20 -UserAgent 'Mozilla/5.0 (Vastgoednieuws-lezer)'
        $inhoud = [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
        $xml = [xml]$inhoud
        $items = @()
        if ($xml.rss) { $items = @($xml.rss.channel.item) } elseif ($xml.feed) { $items = @($xml.feed.entry) }
        $n = 0
        foreach ($i in $items) {
            if ($xml.rss) {
                $titel = Get-Tekst $i.title
                $link = [string]$i.link
                $datum = [string]$i.pubDate
                $samenvatting = Get-Tekst $i.description
                $bron = $b.naam
            } else {
                $titel = Get-Tekst (Get-Node $i.title)
                $l = @($i.link) | Where-Object { $_.href } | Select-Object -First 1
                $link = [string]$l.href
                $datum = if ($i.published) { [string]$i.published } else { [string]$i.updated }
                $samenvatting = Get-Tekst (Get-Node $(if ($i.summary) { $i.summary } else { $i.content }))
                $bron = $b.naam
            }
            if ($b.google) {
                # Google News: "Titel - Krant"; de beschrijving is enkel een herhaling van de titel
                $samenvatting = ''
                if ($i.source -and $i.source.'#text') { $bron = [string]$i.source.'#text' }
                if ($titel -match '^(.*) - [^-]+$') { $titel = $Matches[1] }
            }
            try { $d = [DateTimeOffset]::Parse($datum, [Globalization.CultureInfo]::InvariantCulture) } catch { continue }
            if ($d -lt $grens -or -not $link -or -not $titel) { continue }
            if ($b.filter -and ("$titel $samenvatting" -notmatch $vastgoedWoorden)) { continue }
            $isRegio = [bool]("$titel $samenvatting" -match $regioWoorden)
            if ($b.forceregio) { $isRegio = $true }   # bron gaat uitsluitend over Kortrijk
            if ($samenvatting.Length -gt 280) { $samenvatting = $samenvatting.Substring(0, 280).TrimEnd() + '…' }
            [void]$artikels.Add([pscustomobject]@{ titel = $titel; link = $link; datum = $d.UtcDateTime.ToString('o'); bron = $bron; samenvatting = $samenvatting; regio = $isRegio })
            $n++
        }
        $status += "OK    $($b.naam): $n artikels"
    } catch {
        $status += "FOUT  $($b.naam): $($_.Exception.Message)"
    }
}

# Dubbels weg (zelfde titel), nieuwste eerst
$gezien = @{}
# Bij dubbels blijft de regio-versie behouden
$uniek = $artikels | Sort-Object @{ Expression = { $_.regio }; Descending = $true }, @{ Expression = { [datetime]$_.datum }; Descending = $true } | Where-Object {
    $k = ($_.titel.ToLower() -replace '[^a-z0-9]', '')
    if ($gezien.ContainsKey($k)) { $false } else { $gezien[$k] = 1; $true }
} | Sort-Object { [datetime]$_.datum } -Descending

$json = (@($uniek) | ConvertTo-Json -Depth 4 -Compress)
if (-not $json -or $json -eq 'null') { $json = '[]' }
if (@($uniek).Count -eq 1) { $json = "[$json]" }
$json = $json -replace '</', '<\/'
$bijgewerkt = (Get-Date).ToString('dd/MM/yyyy HH:mm')

$sjabloon = Get-Content (Join-Path $map 'sjabloon.html') -Raw -Encoding UTF8
$html = $sjabloon.Replace('__DATA__', $json).Replace('__BIJGEWERKT__', $bijgewerkt)
$uit = Join-Path $map 'index.html'
[IO.File]::WriteAllText($uit, $html, (New-Object Text.UTF8Encoding($true)))

$status | ForEach-Object { Write-Host $_ }
Write-Host "Totaal: $(@($uniek).Count) unieke artikels -> $uit"
if (-not $GeenBrowser) { Start-Process $uit }
