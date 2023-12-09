function Get-SplLocalTime() {
	[TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::UtcNow, "Tokyo Standard Time")
}

function Get-ScheduleMonth() {
	$fn = "{0:yyyyMM}.json" -f (Get-SplLocalTime)
	$uri = "https://www.puroland.jp/common_v2/json/{0}" -f $fn
	iwr $uri -OutFile ("schedule2m/{0}" -f $fn)
}

function Get-ScheduleDay() {
	$fn = "{0:yyyyMMdd}.json" -f (Get-SplLocalTime)
	$uri = "https://www.puroland.jp/system/json/schedule/day/{0}" -f $fn
	iwr $uri -OutFile ("schedule2d/{0}" -f $fn)
}

function Invoke-SplSchelduleMain() {
	mkdir -f schedule2m
	mkdir -f schedule2d
	Get-ScheduleMonth
	Get-ScheduleDay
}
