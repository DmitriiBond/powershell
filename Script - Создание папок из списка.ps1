# Основной файл разбили на три отдельных с объединением ячеек для удобства в данном кокретном случае.
$A = Get-Content "Path to file A.csv here" | Select-Object -Skip 1 # Получили номера группы из файла "A"
$B_C = Get-Content "Path to file B_C.csv here" | Select-Object -Skip 1 # Получили номера п/п и даты из файла "B_C"
$D_E = Get-Content "Path to file D_E.csv here" | Select-Object -Skip 1 # Получили имена клиентов и даты из файла "D_E"
$i = 0
foreach ($elem in $A)
{
$A1 = $A[$i].replace(";","")
$B_C1 = $B_C[$i].replace(";","")
$D_E1 = $D_E[$i].replace(";","")

[System.IO.Directory]::CreateDirectory("C:\Users\radmin\Desktop\csv\papki\$A1\") #Создаём папку с группой
[System.IO.Directory]::CreateDirectory("C:\Users\radmin\Desktop\csv\papki\$A1\$B_C1\") #Создаём папку с номером п/п и датой
[System.IO.Directory]::CreateDirectory("C:\Users\radmin\Desktop\csv\papki\$A1\$B_C1\$D_E1\") #Создаём папку с именем и датой
[System.IO.Directory]::CreateDirectory("C:\Users\radmin\Desktop\csv\papki\$A1\$B_C1\$D_E1\Выписки, платежные документы") #Создаём стандартные папки у каждого клиента
[System.IO.Directory]::CreateDirectory("C:\Users\radmin\Desktop\csv\papki\$A1\$B_C1\$D_E1\Документы по кредиту")
[System.IO.Directory]::CreateDirectory("C:\Users\radmin\Desktop\csv\papki\$A1\$B_C1\$D_E1\Информация по заемщику")
[System.IO.Directory]::CreateDirectory("C:\Users\radmin\Desktop\csv\papki\$A1\$B_C1\$D_E1\Иные существенные факторы")
[System.IO.Directory]::CreateDirectory("C:\Users\radmin\Desktop\csv\papki\$A1\$B_C1\$D_E1\Направление ссуды")
[System.IO.Directory]::CreateDirectory("C:\Users\radmin\Desktop\csv\papki\$A1\$B_C1\$D_E1\Обеспечение")
[System.IO.Directory]::CreateDirectory("C:\Users\radmin\Desktop\csv\papki\$A1\$B_C1\$D_E1\Оценка")
[System.IO.Directory]::CreateDirectory("C:\Users\radmin\Desktop\csv\papki\$A1\$B_C1\$D_E1\Прочие документы")
[System.IO.Directory]::CreateDirectory("C:\Users\radmin\Desktop\csv\papki\$A1\$B_C1\$D_E1\Решения")
[System.IO.Directory]::CreateDirectory("C:\Users\radmin\Desktop\csv\papki\$A1\$B_C1\$D_E1\Финансовое состояние")
$i++
}



