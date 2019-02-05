param (
  [string]$datecapture = 'thimonth',
  [string]$ReportFile  = 'C:\Windows\Temp\collect_log_update.csv'
)

$os_name = (Get-WmiObject -Class Win32_OperatingSystem).caption
$UpdateActivityTable = New-Object system.Data.DataTable "Update Activity"

$os = New-Object system.Data.DataColumn "os",([string])
$date = New-Object system.Data.DataColumn "date",([string])
$id = New-Object system.Data.DataColumn "id",([string])
$status = New-Object system.Data.DataColumn "status",([string])
$reference = New-Object system.Data.DataColumn "reference",([string])
$describe = New-Object system.Data.DataColumn "describe",([string])
$verifylog = New-Object system.Data.DataColumn "verifylog",([string])

$UpdateActivityTable.columns.add($os)
$UpdateActivityTable.columns.add($date)
$UpdateActivityTable.columns.add($id)
$UpdateActivityTable.columns.add($status)
$UpdateActivityTable.columns.add($reference)
$UpdateActivityTable.columns.add($describe)
$UpdateActivityTable.columns.add($verifylog)

if ( $datecapture -eq 'lastmonth' ) {
  $lastDayOfMonth = (Get-Date -day 1 -hour 23 -minute 59 -second 59).AddDays(-1)
  $firstDayOfMonth = (Get-Date $lastDayOfMonth -day 1 -hour 0 -minute 0 -second 0)
} else {
  $firstDayOfMonth = Get-Date -day 1 -hour 0 -minute 0 -second 0
  $lastDayOfMonth = (Get-Date $firstDayOfMonth).AddMonths(1).addSeconds(-1)
}

$firstDayOfMonthConverted=($firstDayOfMonth.ToString("yyyy-MM-dd"))+"T00:00:00.000Z"
$lastDayOfMonthConverted=($lastDayOfMonth.ToString("yyyy-MM-dd"))+"T23:59:59.999Z"

$filterXml = @"
<QueryList>
  <Query Id="0" Path="System">
    <Select Path="System">
      *[System
        [Provider
          [@Name='Microsoft-Windows-WindowsUpdateClient']
          and Task = 1
          and (band(Keywords,8200))
          and (EventID=19 or EventID=20)
          and
          TimeCreated[
            @SystemTime&gt;='$firstDayOfMonthConverted'
            and
            @SystemTime&lt;='$lastDayOfMonthConverted'
          ]
        ]
      ]
</Select>
</Query>
</QueryList>
"@

$systemEvents = Get-WinEvent -ErrorAction SilentlyContinue -FilterXml $filterXml | select TimeCreated, RecordId, LevelDisplayName, KeywordsDisplayNames, Properties, Message

if ($systemEvents) {
  foreach ($i in $systemEvents){
    $row = $UpdateActivityTable.NewRow()
    $row.os = $os_name
    $row.date=$i.TimeCreated.ToString("yyyy-MM-dd HH:MM:ss")
    $id_now = $i.RecordId
    $row.id = $id_now
    $row.status = $i.KeywordsDisplayNames[1]
    $row.reference =  $i.Message.split('(|)')[-2]
    $row.describe = $i.Message.Split(':')[-1]
    $row.verifylog = "Get-WinEvent -LogName System -FilterXPath `"*[System[EventRecordID=$id_now]]`""
    $UpdateActivityTable.Rows.Add($row)
  }
}else {
  $row = $UpdateActivityTable.NewRow()
  $not_status = 'Nada Encontrado'
  $row.os = $os_name
  $row.date = $not_status
  $row.id = $not_status
  $row.status = $not_status
  $row.reference = $not_status
  $row.describe = $not_status
  $row.verifylog = $not_status

  $UpdateActivityTable.Rows.Add($row)
}

$UpdateActivityTable | Export-Csv -path "$ReportFile" -NoTypeInformation
