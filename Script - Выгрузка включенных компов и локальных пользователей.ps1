$comps = (Get-ADComputer -Filter *).name
$head = "Имя компа;Имя пользователя;SamAccountName" |out-file C:\Users\radmin\Desktop\test3.csv -Encoding utf8
Foreach($c in $comps) {
IF (Test-Connection -BufferSize 32 -Count 1 -ComputerName $c -Quiet)
{
$loginname = ((Get-WmiObject -ComputerName $c -Class win32_computersystem).username -split '\\')
$loginname1 = $loginname[1]
$name = (Get-ADUser -Filter "samaccountname -eq '$loginname1'" -Properties name).name
$c + ";" + $name + ";" + $loginname1 | out-file C:\Users\radmin\Desktop\test3.csv -Encoding utf8 -Append
}
else
{
$name = "Нет данных"
$loginname1 = "Нет данных"
$c + ";" + $name + ";" + $loginname1 | out-file C:\Users\radmin\Desktop\test3.csv -Encoding utf8 -Append
}
}
