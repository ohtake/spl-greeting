[Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null
$LCID_JAJP = (New-Object -TypeName System.Globalization.CultureInfo -ArgumentList @("ja-jp", $true)).LCID
$outdir = "schedule"

function Get-SplLocalTime() {
	[TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::UtcNow, "Tokyo Standard Time")
}

function Get-SplGreeting() {
	$proxy = $null
	$baseUri = "http://puroland.co.jp/chara_gre/"
	$homeUri = $baseUri + "mobile/"
	$listUriTemplate = $baseUri + "mobile/chara_sentaku.asp?tchk={0}"
	$detailUriTemplate = $baseUri + "mobile/chara_sche.asp?tchk={0}&C_KEY={1}"
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
		$body = wget-splgreeting($homeUri)
		if($body -match '公開されておりません。P'){
			Write-Verbose "Retrying with 'para' paramter" -Verbose
			$body = wget-splgreeting($homeUri + ("?para={0:yyyyMMdd}" -f (Get-SplLocalTime)))
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
		$ids = @($body -split "<BR>" |% {if ($_ -match "C_KEY=(\d+)") {[int]$Matches[1]}})
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
		if($body -match "<P align=center><FONT size=-1>(.+)</FONT></P>"){
			$name = $Matches[1]
		}else{
			Write-Verbose $body -Verbose
			throw "Cannot find name"
		}
		$body -split "</P>" |
			% {if($_ -match "<FONT Size=-1>([\d:０-９：]+)[-－]([\d:０-９：]+)<BR>(.+?)</FONT>"){$Matches}} |
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
		[Microsoft.VisualBasic.Strings]::StrConv($str, [Microsoft.VisualBasic.VbStrConv]::Narrow, $LCID_JAJP)
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
	mkdir $outdir -ErrorAction SilentlyContinue
	$result["today"] | Export-Csv (Join-Path $outdir ("{0:yyyyMMdd}.csv" -f (Get-SplLocalTime))) -Encoding UTF8 -NoTypeInformation
	$result["tomorrow"] | Export-Csv (Join-Path $outdir ("{0:yyyyMMdd}_next.csv" -f $result["tomorrow"][0].Date)) -Encoding UTF8 -NoTypeInformation
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
		ls -Filter (Join-Path $outdir "*.csv") |? { $_.Name -match "\d{8}$suffix\.csv" } |% {
			New-Object PSObject -Property @{
				Name = $_.Name
				YYYYMM = $_.Name.Substring(0,6)
			}
		} | group YYYYMM |% {
			$ym = $_.Name
			$items = @()
			$_.Group |% {
				$items += Import-Csv (Join-Path $outdir $_.Name)
			}
			$items | Export-Csv (Join-Path $outdir "$ym$suffix.csv") -Encoding UTF8 -NoTypeInformation
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
		[Microsoft.VisualBasic.Strings]::StrConv($str, [Microsoft.VisualBasic.VbStrConv]::Wide, $LCID_JAJP)
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

function Format-SplCsvDateTime() {
	$temp_dir = "temp_format"
	mkdir $temp_dir -ErrorAction SilentlyContinue
	ls -Filter (Join-Path $outdir "*.csv") |% {
		$changed = $false
		$name = $_
		$items = Import-Csv (Join-Path $outdir $name)
		$items |% {
			# Code duplitaes? It is required for faster execution.
			if ($_.Start) {
				$parsed = [DateTime]$_.Start
				if ($_.Start -ne [string]$parsed) {
					$_.Start = $parsed
					$changed = $true
				}
			}
			if ($_.End) {
				$parsed = [DateTime]$_.End
				if ($_.End -ne [string]$parsed) {
					$_.End = $parsed
					$changed = $true
				}
			}
			if ($_.Date) {
				$parsed = [DateTime]$_.Date
				if ($_.Date -ne [string]$parsed) {
					$_.Date = $parsed
					$changed = $true
				}
			}
		}
		if ($changed) {
			$items | Export-Csv (Join-Path $temp_dir $name) -Encoding UTF8 -NoTypeInformation
		}
	}
	mv -Force $temp_dir/*.csv $outdir/
	rm -Recurse $temp_dir
}
