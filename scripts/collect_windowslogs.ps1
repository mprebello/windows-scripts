param (
  [string]$datecapture = 'lastmonth',
  [string]$ReportFile  = 'C:\Windows\Temp\collect.csv'
)

# Table Creation
$LogonActivityTable = New-Object system.Data.DataTable "Logon/Logoff Activity"

# Create Columns
$date = New-Object system.Data.DataColumn "Date",([string])
$type = New-Object system.Data.DataColumn "Type",([string])
$status = New-Object system.Data.DataColumn "Status",([string])
$user = New-Object system.Data.DataColumn "User",([string])
$ipaddress = New-Object system.Data.DataColumn "IPAddress",([string])

# Add Columns to Table
$LogonActivityTable.columns.add($date)
$LogonActivityTable.columns.add($type)
$LogonActivityTable.columns.add($status)
$LogonActivityTable.columns.add($user)
$LogonActivityTable.columns.add($ipaddress)

if ( $datecapture -eq 'lastmonth' ) {
  $lastDayOfMonth = (Get-Date -day 1 -hour 23 -minute 59 -second 59).AddDays(-1)
  $firstDayOfMonth = (Get-Date $lastDayOfMonth -day 1 -hour 0 -minute 0 -second 0)
} else {
  $firstDayOfMonth = Get-Date -day 1 -hour 0 -minute 0 -second 0
  $lastDayOfMonth = (Get-Date $firstDayOfMonth).AddMonths(1).addSeconds(-1)
}

#$firstDayOfMonthConverted=$firstDayOfMonth.ToString("yyyy-MM-ddTHH:MM:ss")
#$lastDayOfMonthConverted=$lastDayOfMonth.ToString("yyyy-MM-ddTHH:MM:ss")
$firstDayOfMonthConverted=($firstDayOfMonth.ToString("yyyy-MM-dd"))+"T00:00:00.000Z"
$lastDayOfMonthConverted=($lastDayOfMonth.ToString("yyyy-MM-dd"))+"T23:59:59.999Z"

$filterXml="<QueryList>
  <Query Id='0' Path='Security'>
    <Select Path='Security'>
     *[
        (System[
          (EventID=4624 or EventID=4634 or EventID=4625)
          and
          TimeCreated[@SystemTime&gt;='$firstDayOfMonthConverted'
          and
          @SystemTime&lt;='$lastDayOfMonthConverted']
        ]
        and
        EventData[
          Data[@Name='LogonType']
          and
          (Data='10' or Data='2' or Data='11')
        ])
        or
        (System[(EventID=4647)
        and
        TimeCreated[@SystemTime&gt;='$firstDayOfMonthConverted'
        and
        @SystemTime&lt;='$lastDayOfMonthConverted']
        ])
     ]
    </Select>
  </Query>
</QueryList>"

$log = Get-WinEvent -FilterXml $filterXml | select Id,TimeCreated,Properties

foreach ($i in $log){
  if ((($i.Id -eq 4624) -or ($i.Id -eq 4634)) -and ($i.Properties[8].value -eq 2)) {
    $row = $LogonActivityTable.NewRow()
    $row.date =  $i.TimeCreated.ToString("yyyy-MM-dd HH:MM:ss")
    $row.type =  "Logon - Local"
    $row.status =  "Success"
    $row.user =  $i.Properties[5].value
    $row.ipaddress = ""
    $LogonActivityTable.Rows.Add($row)
  }

  if ((($i.Id -eq 4624) -or ($i.Id -eq 4634)) -and ($i.Properties[8].value -eq 10)) {
    $row = $LogonActivityTable.NewRow()
    $row.date =  $i.TimeCreated.ToString("yyyy-MM-dd HH:MM:ss")
    $row.type =  "Logon - Remote"
    $row.status =  "Success"
    $row.user =  $i.Properties[5].value
    $row.ipaddress = $i.Properties[18].value
    $LogonActivityTable.Rows.Add($row)
  }

  if ((($i.Id -eq 4624) -or ($i.Id -eq 4634)) -and ($i.Properties[8].value -eq 11)) {
    $row = $LogonActivityTable.NewRow()
    $row.date =  $i.TimeCreated.ToString("yyyy-MM-dd HH:MM:ss")
    $row.type =  "Logon - Cache"
    $row.status =  "Success"
    $row.user =  $i.Properties[5].value
    $row.ipaddress = $i.Properties[18].value
    $LogonActivityTable.Rows.Add($row)
  }

  if ($i.Id -eq 4647){
    $row = $LogonActivityTable.NewRow()
    $row.date =  $i.TimeCreated.ToString("yyyy-MM-dd HH:MM:ss")
    $row.type =  "Logoff"
    $row.status =  "Success"
    $row.user =  $i.Properties[1].value
    $row.ipaddress = ""
    $LogonActivityTable.Rows.Add($row)
  }

  if (($i.Id -eq 4625) -and ($i.Properties[10].value -eq 2)) {
    $row = $LogonActivityTable.NewRow()
    $row.date =  $i.TimeCreated.ToString("yyyy-MM-dd HH:MM:ss")
    $row.type =  "Logon - Local"
    $row.status =  "Failure"
    $row.user =  $i.Properties[5].value
    $row.ipaddress = ""
    $LogonActivityTable.Rows.Add($row)
  }

  if (($i.Id -eq 4625) -and ($i.Properties[10].value -eq 10)) {
    $row = $LogonActivityTable.NewRow()
    $row.date =  $i.TimeCreated.ToString("yyyy-MM-dd HH:MM:ss")
    $row.type =  "Logon - Remote"
    $row.status =  "Failure"
    $row.user =  $i.Properties[5].value
    $row.ipaddress = $i.Properties[19].value
    $LogonActivityTable.Rows.Add($row)
  }

}

$LogonActivityTable | Export-Csv -path "$ReportFile" -NoTypeInformation
