#Серийный номер
#Фирма
#ЦПУ
#Кол-во ОЗУ (ГБ)
#Кол-во планок
#Объём ХДД
#Занято/своб на диске C:\
#Current User
$head = "Компьютер;Пользователь;Серийный номер;Фирма производитель;ЦПУ;Объём ОЗУ;Кол-во планок ОЗУ;Объём дисков; Кол-во дисков;Занято на C:\;Свободно на C:\;" |out-file C:\Users\radmin\Desktop\Comp-info.csv -Encoding utf8
$comps = (Get-ADComputer -Filter *).name
foreach ($comp in $comps)
{
IF (Test-Connection -BufferSize 32 -Count 1 -ComputerName $c -Quiet)
{
$serial = (Get-WmiObject -class win32_Bios -ComputerName "$comp").SerialNumber #Серийник

$firma =(Get-WmiObject -ClassName Win32_ComputerSystem -ComputerName "$comp").Manufacturer #Фирма

$cpu = (Get-WmiObject -ClassName Win32_Processor -ComputerName "$comp").name #Процессор

$rams = (Get-WmiObject -ComputerName "$comp" win32_physicalmemory).capacity #Массив с объёмами планок RAM
$ram_size = [string](($rams | Measure-Object -sum).sum/1GB) + 'GB'
$ram_count = $rams.count

$hdds = (Get-WmiObject -ComputerName "$comp" win32_diskdrive).size #Массив с объёмами HDD/SSD
$hdd_size = [string]([math]::Round(($hdds | Measure-Object -sum).sum/1GB, 2)) + 'GB'
$hdd_count = $hdds.count

# Колво планок $rams.count
# Объём планок суммарно [string](($rams | Measure-Object -sum).sum/1GB) + 'GB'
# Колво HDD/SSD $hdds.count
# Объём HDD/SSD суммарно [string](($hdds | Measure-Object -sum).sum/1GB) + 'GB'

    $volumes = Get-WmiObject win32_logicaldisk -ComputerName "$comp" #Выбираем диск C:\
    foreach ($volume in $volumes)
    {
    if ($volume -like "*C:*")
    {
    $vol_C = $volume
    }
    }

$c_free = [string]([math]::Round(($vol_c).FreeSpace/1gb, 2)) + 'GB'
$c_occ = [string]([math]::Round(($vol_c).size/1gb, 2)-[math]::Round(($vol_c).FreeSpace/1gb, 2)) + 'GB'


# Занято места на диске с 2 зн. после запятой C:\ [string]([math]::Round(($vol_c).size/1gb, 2)-[math]::Round(($vol_c).FreeSpace/1gb, 2))
# Своб. место на диске с 2 зн. после запятой [string]([math]::Round(($vol_c).FreeSpace/1gb, 2))

$loginname = ((Get-WmiObject -ComputerName $comp -Class win32_computersystem).username -split '\\') #CurrentUser
$loginname1 = $loginname[1]

"$comp" + ";" + $loginname[1] + ";" + $serial + ";" + $firma + ";" + $cpu + ";" + $ram_size + ";" + $ram_count + ";" + $hdd_size + ";" + $hdd_count + ";" + $c_occ + ";" + $c_free |out-file C:\Users\radmin\Desktop\Comp-info.csv -Encoding utf8 -Append

}
else
{
$serial = "Нет данных"
$firma = "Нет данных"
$cpu = "Нет данных"
$ram_size = "Нет данных"
$ram_count = "Нет данных"
$hdd_size = "Нет данных"
$hdd_count = "Нет данных"
$c_free = "Нет данных"
$c_occ = "Нет данных"
$loginname1 = "Нет данных"
"$comp" + ";" + $loginname[1] + ";" + $serial + ";" + $firma + ";" + $cpu + ";" + $ram_size + ";" + $ram_count + ";" + $hdd_size + ";" + $hdd_count + ";" + $c_occ + ";" + $c_free |out-file C:\Users\radmin\Desktop\Comp-info.csv -Encoding utf8 -Append

}

}