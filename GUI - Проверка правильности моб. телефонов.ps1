Add-Type -assembly System.Windows.Forms
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Проверка правильности введённых телефонов в AD'
$main_form.Width = 460
$main_form.Height = 400
$main_form.AutoSize = $true
$main_form.FormBorderStyle = 'FixedDialog'

# Заголовок 1
$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text = "Выберите подразделение:"
$Label2.Location  = New-Object System.Drawing.Point(10,10)
$Label2.AutoSize = $true
$main_form.Controls.Add($Label2)

# Выбор OU
$ComboBox = New-Object System.Windows.Forms.ComboBox
$ComboBox.Width = 300
$OU_list = Get-ADOrganizationalUnit -filter * -Properties name, distinguishedname
    Foreach ($OU in $OU_List)
    {
    $ComboBox.Items.Add($OU.name);
    }
$ComboBox.Location  = New-Object System.Drawing.Point(10,30)
$main_form.Controls.Add($ComboBox)

# Заголовок 2
$Label3 = New-Object System.Windows.Forms.Label
$Label3.Text = "Выберите путь для сохранения отчета:"
$Label3.Location  = New-Object System.Drawing.Point(10,60)
$Label3.AutoSize = $true
$main_form.Controls.Add($Label3)

# Кнопка выбора папки
$Button_browse = New-Object System.Windows.Forms.Button
$Button_browse.Location = New-Object System.Drawing.Size(350,80)
$Button_browse.Size = New-Object System.Drawing.Size(110,20)
$Button_browse.Text = "Выбрать"
$main_form.Controls.Add($Button_browse)

# Строка выбора папки
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,80)
$textBox.Size = New-Object System.Drawing.Size(300,20)
$main_form.Controls.Add($textBox)

# Нажатие кнопки выбора папки
$Button_browse.Add_Click(
{
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
[void]$FolderBrowser.ShowDialog()
$FolderBrowser.SelectedPath
$textBox.Text = $FolderBrowser.SelectedPath
}
)

# Кнопка генерации отчета
$Button_export = New-Object System.Windows.Forms.Button
$Button_export.Location = New-Object System.Drawing.Size(160,150)
$Button_export.Size = New-Object System.Drawing.Size(150,20)
$Button_export.Text = "Сформировать"
$main_form.Controls.Add($Button_export)

# Строка загрузки
$ProgressBar              = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location     = New-Object System.Drawing.Point(10,250)
$ProgressBar.Size         = New-Object System.Drawing.Size(460,20)
$ProgressBar.Style        = "Blocks"
$ProgressBar.MarqueeAnimationSpeed = 20

$main_form.Controls.Add($ProgressBar);

# Статус выполнения в текст-боксе
$StatusBox = New-Object system.Windows.Forms.TextBox
$StatusBox.multiline = $true
$StatusBox.width = 460
$StatusBox.height = 100
$StatusBox.location = New-Object System.Drawing.Point(10,300)
$StatusBox.Font = 'Microsoft Sans Serif,10'
$StatusBox.Scrollbars = "Vertical" 
$main_form.Controls.Add($StatusBox);

# Нажатие кнопки генерации отчета
$Button_export.Add_Click(
{
$date = Get-Date -Format "MM-dd-yyyy-HH_mm_ss"
$path = $textbox.Text + "\" + $date + ".csv"
$head = "Пользователь ;Моб. тел.;Верно?;" |out-file $path -Encoding utf8
$users = Get-ADUser -Filter * -SearchBase $OU_list[$ComboBox.SelectedIndex] -Properties name,mobile | select name,mobile

$progressbar.Maximum = $users.Count
$progressbar.Step = 1
$progressbar.Value = 0

$i = 0
$k = 0
$j = 0
$Statusbox.AppendText("Script by Bondarev D.A. ver. 0.1`r`n")
$Statusbox.AppendText("Starting...`r`n")
Start-Sleep -Seconds 1
foreach ($user in $users)
{
    if ($user.mobile -match '\+\d{1}\(\d{3}\)\d{3}-\d{2}-\d{2}')
    {
    $user.name + ";" + $user.mobile + ";" + "Да" |out-file $path -Encoding utf8 -Append
    $i++
    $progressbar.PerformStep()
    $nameOutput = $user.name + " - Пользователь обработан"
    $Statusbox.AppendText("$nameOutput`r`n")
    }
elseif ($user.mobile -eq $null)
{
$user.name + ";" + $user.mobile + ";" + "Нет" |out-file $path -Encoding utf8 -Append
$k++
$progressbar.PerformStep()
$nameOutput = $user.name + " - Пользователь обработан"
    $Statusbox.AppendText("$nameOutput`r`n")
}
else
{
$user.name + ";" + $user.mobile + ";" + "Нет" |out-file $path -Encoding utf8 -Append
$j++
$progressbar.PerformStep()
$nameOutput = $user.name + " - Пользователь обработан"
    $Statusbox.AppendText("$nameOutput`r`n")
}
}
"Правильных:" + ";" + $i + ";" + [string][math]::Round(($i/($i+$j+$k)*100), 2) + "%" |out-file $path -Encoding utf8 -Append
"Незаполненных:" + ";" + $k + ";" + [string][math]::Round(($k/($i+$j+$k)*100), 2) + "%" |out-file $path -Encoding utf8 -Append
"Неправильных:" + ";" + $j + ";" + [string][math]::Round(($j/($i+$j+$k)*100), 2) + "%" |out-file $path -Encoding utf8 -Append
"Незаполненных в неправильных:" + ";" + $k + ";" + [string][math]::Round(($k/($j+$k)*100), 2) + "%" |out-file $path -Encoding utf8 -Append
$Statusbox.AppendText("Готово`r`n")
$ChildForm = New-Object system.Windows.Forms.Form
$ChildForm.Width = 150
$ChildForm.Height = 100
$ChildForm.AutoSize = $true
$ChildForm.FormBorderStyle = 'FixedDialog'

$Child_text = New-Object System.Windows.Forms.Label
$Child_text.Text = "Готово."
$Child_text.Location  = New-Object System.Drawing.Point(50,35)
$Child_text.Font = 'Microsoft Sans Serif,10'
$ChildForm.Controls.Add($Child_text)

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(40,70)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$ChildForm.AcceptButton = $OKButton
$ChildForm.Controls.Add($OKButton)

$ChildForm.ShowDialog()
}

)


$main_form.ShowDialog()
