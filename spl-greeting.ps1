﻿[Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null

function Get-SplLocalTime() {
	[TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::UtcNow, "Tokyo Standard Time")
}

function Get-SplGreeting() {
	$proxy = $null
	$baseUri = "http://puroland.co.jp/chara_gre/"
	$listUriTemplate = $baseUri + "chara_sentaku.asp?tchk={0}"
	$detailUriTemplate = $baseUri + "chara_sche.asp?tchk={0}&C_KEY={1}"
	$tomorrowUriTemplate = $baseUri + "chara_sentaku_nextday.asp?tchk={0}"
	$userAgent = "Mozilla/5.0 (PowerShell; https://github.com/ohtake/spl-greeting)"
	$encoding = [Text.Encoding]::GetEncoding("Shift_JIS")
	$maxTry = 5

	function wget-splgreeting([String]$uri) {
		$wc = New-Object Net.WebClient
		if ($proxy) {
			$wc.Proxy = $proxy
		}

		for($try=1; $try -le $maxTry; $try++) {
			try {
				$wc.Encoding = $encoding
				$wc.Headers.Add([Net.HttpRequestHeader]::UserAgent, $userAgent)
				return $wc.DownloadString($uri)
			} catch [System.Net.WebException] {
				$ex = $_.Exception
				switch ($ex.Response.StatusCode) {
					ProxyAuthenticationRequired {
						$proxy = New-Object System.Net.WebProxy (Read-Host -Prompt "Proxy address")
						$proxy.Credentials = Get-Credential
						$wc.Proxy = $proxy
						break
					}
					default {
						Write-Warning $ex
						break
					}
				}
				continue
			} catch {
				Write-Warning $_.Exception
				continue
			}
		}
		throw ("Failed to retrieve {0}" -f $uri)
	}
	function get-tchk() {
		$body = wget-splgreeting($baseUri)
		if($body -match '公開されておりません。P'){
			Write-Verbose "Retrying with 'para' paramter" -Verbose
			$body = wget-splgreeting($baseUri + ("?para={0:yyyyMMdd}" -f (Get-SplLocalTime)))
		}
		if($body -match 'name="TCHK" value="(\d+)"'){
			[int]$Matches[1]
		}else{
			Write-Verbose $body -Verbose
			throw "Cannot find tchk"
		}
	}
	function get-ids() {
		$body = wget-splgreeting($listUriTemplate -f $tchk)
		$ids = @($body -split "<form " |% {if ($_ -match "name=C_KEY value=(\d+)") {[int]$Matches[1]}})
		if($ids.Count -eq 0) {
			Write-Verbose $body -Verbose
			throw "Cannot find any CIDs"
		}
		$ids
	}
	function get-tomorrow() {
		$body = wget-splgreeting($tomorrowUriTemplate -f $tchk)
		$names = $body -split "\n" |
			? {$_ -match "<tr align=center>" } |
			% { $_ -split "</?td>"} |
			? { $_ -ne "" } |
			? { $_ -NotMatch "<" }
		if ($body -match '<div class="newsTop3">(\d+)年(\d+)月(\d+)日は') {
			$date = New-Object DateTime @([int]$Matches[1],[int]$Matches[2],[int]$Matches[3])
		} else {
			Write-Verbose $body -Verbose
			throw "Cannot find date of next day"
		}
		$names |% { New-Object PSObject -Property @{Name=$_; Date = $date} }
	}
	function get-items($id) {
		$body = wget-splgreeting($detailUriTemplate -f $tchk,$id)
		if($body -match '(\d+)年(\d+)月(\d+)日') {
			$date = New-Object DateTime @([int]$Matches[1],[int]$Matches[2],[int]$Matches[3])
		}else{
			Write-Verbose $body -Verbose
			throw "Cannot find date"
		}
		if($body -match '<div id="date3">(.+)</div>'){
			$name = $Matches[1]
		}else{
			Write-Verbose $body -Verbose
			throw "Cannot find name"
		}
		$body -split "<div id=date>" |
			% {if($_ -match "^([\d:０-９：]+)[-－]([\d:０-９：]+)<BR>(.+?)</FONT>"){$Matches}} |
			% {
				New-Object PSObject -Property @{
					CID = $id;
					Name = $name;
					Start = $date.Add([TimeSpan]::Parse((to-hankaku($_[1]))));
					End = $date.Add([TimeSpan]::Parse((to-hankaku($_[2]))));
					Location = $_[3]}
			}
	}
	function to-hankaku($str) {
		[Microsoft.VisualBasic.Strings]::StrConv($str, [Microsoft.VisualBasic.VbStrConv]::Narrow)
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
	$tomorrow = get-tomorrow
	return @{today = $items; tomorrow = $tomorrow }
}

function Invoke-SplGreetingMain() {
	$result = Get-SplGreeting
	$result["today"] | Export-Csv ("{0:yyyyMMdd}.csv" -f (Get-SplLocalTime)) -Encoding UTF8 -NoTypeInformation
	$result["tomorrow"] | Export-Csv ("{0:yyyyMMdd}_next.csv" -f $result["tomorrow"][0].Date) -Encoding UTF8 -NoTypeInformation
	$result["today"] |% {
		$readable = ([DateTime]$_.Start).ToString("HH:mm")
		$readable += "-"
		$readable += ([DateTime]$_.End).ToString("HH:mm")
		$readable += " "
		$readable += $_.Location -replace "[(（].+[)）]",""
		$_ | Add-Member -MemberType NoteProperty "FriendlyTimeAndLocation" $readable -PassThru -Force
	} | group FriendlyTimeAndLocation | sort Name | select Name,{@($_.Group|%{$_.Name}) -join ", "} | ft -AutoSize -Wrap
	diff ($result["today"] | select Name -Unique |% {$_.Name}) ($result["tomorrow"] |% {$_.Name}) -IncludeEqual | ft -AutoSize -Wrap
}

function Merge-SplGreetingCsv() {
	function merge($suffix="") {
		ls -Filter *.csv |? { $_.Name -match "\d{8}$suffix\.csv" } |% {
			New-Object PSObject -Property @{
				Name = $_.Name
				YYYYMM = $_.Name.Substring(0,6)
			}
		} | group YYYYMM |% {
			$ym = $_.Name
			$items = @()
			$_.Group |% {
				$items += Import-Csv $_.Name
			}
			$items | Export-Csv "$ym$suffix.csv" -Encoding UTF8 -NoTypeInformation
		}
	}
	merge
	merge "_next"
}

function Import-SplCsv() {
	param(
		[parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[String[]]
		$filenames
	)
	function to-zenkaku($str) {
		[Microsoft.VisualBasic.Strings]::StrConv($str, [Microsoft.VisualBasic.VbStrConv]::Wide)
	}
	$raw = Import-Csv $filenames
	$raw |% {
		$_.CID = [int]$_.CID
		$_.Name = to-zenkaku($_.Name)
		$_.Location = to-zenkaku($_.Location)
		$_.Start = [DateTime]$_.Start
		$_.End = [DateTime]$_.End
		$_
	}
}
