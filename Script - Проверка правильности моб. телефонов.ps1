$head = "Пользователь ;Моб. тел.;Верно?;" |out-file C:\Users\radmin\Desktop\Mobile.csv -Encoding utf8
$users = Get-ADUser -Filter * -SearchBase 'OU=TestOU, DC=premiere-feo, DC=com' -Properties name,mobile | select name,mobile

$i = 0
$k = 0
$j = 0

foreach ($user in $users)
{
    if ($user.mobile -match '\+\d{1}\(\d{3}\)\d{3}-\d{2}-\d{2}')
    {
    $user.name + ";" + $user.mobile + ";" + "Да" |out-file C:\Users\radmin\Desktop\Mobile.csv -Encoding utf8 -Append
    $i++
    }
elseif ($user.mobile -eq $null)
{
$user.name + ";" + $user.mobile + ";" + "Нет" |out-file C:\Users\radmin\Desktop\Mobile.csv -Encoding utf8 -Append
$k++
}
else
{
$user.name + ";" + $user.mobile + ";" + "Нет" |out-file C:\Users\radmin\Desktop\Mobile.csv -Encoding utf8 -Append
$j++
}
}
"Правильных:" + ";" + $i + ";" + [string][math]::Round(($i/($i+$j+$k)*100), 2) + "%" |out-file C:\Users\radmin\Desktop\Mobile.csv -Encoding utf8 -Append
"Незаполненных:" + ";" + $k + ";" + [string][math]::Round(($k/($i+$j+$k)*100), 2) + "%" |out-file C:\Users\radmin\Desktop\Mobile.csv -Encoding utf8 -Append
"Неправильных:" + ";" + $j + ";" + [string][math]::Round(($j/($i+$j+$k)*100), 2) + "%" |out-file C:\Users\radmin\Desktop\Mobile.csv -Encoding utf8 -Append
"Незаполненных в неправильных:" + ";" + $k + ";" + [string][math]::Round(($k/($j+$k)*100), 2) + "%" |out-file C:\Users\radmin\Desktop\Mobile.csv -Encoding utf8 -Append