$baseUri = "http://puroland.co.jp/chara_gre/"
$listUriTemplate = $baseUri + "chara_sentaku.asp?tchk={0}"
$detailUriTemplate = $baseUri + "chara_sche.asp?tchk={0}&C_KEY={1}"
# User-Agent には携帯っぽい文字列を含んでおく必要あり
$userAgent = "Mozilla/5.0 (PowerShell; https://github.com/ohtake/spl-greeting) (Android)"
$encoding = [Text.Encoding]::GetEncoding("Shift_JIS")

function wget([String]$uri) {
    $wc = New-Object Net.WebClient
    $wc.Encoding = $encoding
    $wc.Headers.Add([Net.HttpRequestHeader]::UserAgent, $userAgent)
    $wc.DownloadString($uri)
}
function get-tchk() {
    $body = wget($baseUri)
    if($body -match 'name="TCHK" value="(\d+)"'){
        [int]$Matches[1]
    }else{
        Write-Verbose $body -Verbose
        throw "Cannot find tchk"
    }
}
function get-ids() {
    $body = wget($listUriTemplate -f $tchk)
    $ids = @($body -split "<BR>" |% {if ($_ -match "C_KEY=(\d+)") {[int]$Matches[1]}})
    if($ids.Count -eq 0) {
        Write-Verbose $body -Verbose
        throw "Cannot find any CIDs"
    }
    $ids
}
function get-items($id) {
    $body = wget($detailUriTemplate -f $tchk,$id)
    if($body -match '(\d+)年(\d+)月(\d+)日') {
        $date = New-Object DateTime @([int]$Matches[1],[int]$Matches[2],[int]$Matches[3])
    }else{
        Write-Verbose $body -Verbose
        throw "Cannot find date"
    }
    if($body -match "<P align=center><FONT size=-1>(.+)</FONT></P>"){
        $name = $Matches[1]
    }else{
        Write-Verbose $body -Verbose
        throw "Cannot find name"
    }
    $body -split "</P>" |
        % {if($_ -match "<FONT Size=-1>([\d:]+)-([\d:]+)<BR>(.+?)</FONT>"){$Matches}} |
        % {
            New-Object PSObject -Property @{
                CID = $id;
                Name = $name;
                Start = $date.Add([TimeSpan]::Parse($_[1]));
                End = $date.Add([TimeSpan]::Parse($_[2]));
                Location = $_[3]}
        }
}

Write-Progress "Fetching TCHK" "Fetching" -PercentComplete 0
$tchk = get-tchk
Write-Progress "Fetching TCHK" ("TCHK = {0}" -f $tchk) -PercentComplete 100
Write-Progress "Fetching IDs" "Fetching" -PercentComplete 0
$ids = get-ids
Write-Progress "Fetching IDs" ("# of IDs: {0}" -f $ids.Count) -PercentComplete 100
$items = $ids |
    % -Begin {
        $i = 0
        $count = $ids.Length
    } -Process {
        Write-Progress "Fetching schedules" ("Character ID: {0}" -f $_) -PercentComplete (100 * $i++ / $count)
        get-items($_)
    } -End {
        Write-Progress "Fetching schedules" ("# of items: {0}" -f $_.Count) -PercentComplete 100
    }

# $items | group Start,End,Location | sort Name | select Name,{@($_.Group|%{$_.Name}) -join ", "} | ft -Wrap
# $items | Export-Csv ("{0:yyyyMMdd}.csv" -f (Get-Date)) -Encoding UTF8
