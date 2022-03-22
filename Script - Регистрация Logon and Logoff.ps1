$time = $null
$comp = $env:COMPUTERNAME
[System.IO.Directory]::CreateDirectory("\\orion\k$\Event Logs\$comp\")
$date = Get-Date -Format "dd-MM-yyyy"
$head = "Дата;Event;Имя пользователя" | out-file "\\orion\k$\Event Logs\$comp\$comp $date.csv" -Encoding utf8
$period = 7
$logs = Get-EventLog System -Source microsoft-windows-winlogon -After (get-date).AddDays(-$period) -ComputerName $comp

If ($logs)
{
    foreach ($log in $logs){
            if ($log.InstanceId -eq 7001)
            {
            $eventname = "Logon"
            $time = $log.TimeWritten
            $user = (New-Object System.Security.Principal.SecurityIdentifier $log.ReplacementStrings[1]).Translate([System.Security.Principal.NTAccount])
            $user1 = $user.Replace("PREMIERE-FEO\", "")
            $time + ";" + $eventname + ";" + $user1 | out-file "\\orion\k$\Event Logs\$comp\$comp $date.csv" -Encoding utf8 -Append
            }
        elseif ($log.InstanceId -eq 7002)
        {
        $eventname = "Logoff"
        $time = $log.TimeWritten
        $user = (New-Object System.Security.Principal.SecurityIdentifier $log.ReplacementStrings[1]).Translate([System.Security.Principal.NTAccount])
        $user1 = $user.Replace("PREMIERE-FEO\", "")
        $time + ";" + $eventname + ";" + $user1 | out-file "\\orion\k$\Event Logs\$comp\$comp $date.csv" -Encoding utf8 -Append
        }
        
Else
{
continue
}
}
}

