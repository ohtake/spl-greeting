$baseUri = "http://puroland.jp/chara_gre/"
$listUriTemplate = $baseUri + "chara_sentaku.asp?tchk={0}"
$detailUriTemplate = $baseUri + "chara_sche.asp?tchk={0}&C_KEY={1}"
# User-Agent には携帯っぽい文字列を含んでおく必要あり
$userAgent = "Mozilla/5.0 (PowerShell; https://github.com/ohtake/spl-greeting) (Android)"

function wget([String]$uri) {
    $wc = New-Object Net.WebClient
    $wc.Headers.Add([Net.HttpRequestHeader]::UserAgent, $userAgent)
    $wc.DownloadString($uri)
}
function get-tchk() {
    if((wget($baseUri)) -match 'name="TCHK" value="(\d+)"'){
        $Matches[1]
    }else{
        throw "Cannot find tchk"
    }
}
function get-ids() {
    (wget($listUriTemplate -f $tchk)) -split "<BR>" |% {if ($_ -match "C_KEY=(\d+)") {$Matches[1]}}
}
function get-items($id) {
    $body = wget($detailUriTemplate -f $tchk,$id)
    if($body -match '(\d+)年(\d+)月(\d+)日') {
        $date = New-Object DateTime @([int]$Matches[1],[int]$Matches[2],[int]$Matches[3])
    }else{
        throw "Cannot find date"
    }
    if($body -match "<P align=center><FONT size=-1>(.+)</FONT></P>"){
        $name = $Matches[1]
    }else{
        throw "Cannot find name"
    }
    $body -split "</P>" |
        % {if($_ -match "<FONT Size=-1>([\d:]+)-([\d:]+)<BR>(.+?)</FONT>"){$Matches}} |
        % {
            New-Object PSObject -Property @{
                Name = $name;
                Start = $date.Add([TimeSpan]::Parse($_[1]));
                End = $date.Add([TimeSpan]::Parse($_[2]));
                Location = $_[3]}
        }
}

$tchk = get-tchk
$ids = get-ids
$items = $ids |% {get-items($_)}

# $items | group start,end,location | sort name | select name,{$_.Group|%{$_.Name}} | ft -Wrap
