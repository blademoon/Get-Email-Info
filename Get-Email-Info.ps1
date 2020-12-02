Import-Module ActiveDirectory
#add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010

# Тестовая версия! Получение общей статистики по почтовому ящику
$email = "username@example.com"
$path = "C:\PATH_TO_RESULT_FILE\" + $email + ".txt"

$email_name = (Get-MailboxStatistics $email).DisplayName
$email_total_size = (Get-MailboxStatistics $email).TotalItemSize.Value.ToMB()
$email_total_items = (Get-MailboxStatistics $email).ItemCount

"Почтовый ящик: " + $email + "." | Out-File -FilePath $path -Encoding utf8 -Append
"Имя почтового ящика: " + $email_name | Out-File -FilePath $path -Encoding utf8 -Append
"Всего занято: " + $email_total_size + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
"Всего элементов (писем и т. д.): " + $email_total_items | Out-File -FilePath $path -Encoding utf8 -Append

$email_db = ""
$email_IssueWarningQuota = ""
$email_ProhibitSendQuota = ""
$email_ProhibitSendReceiveQuota = ""

# Проверяем на каком уровне заданы текущие квоты почтового ящика.
$temp = ((get-mailbox $email -ResultSize Unlimited).UseDatabaseQuotaDefaults)

if (($temp) -eq 1) {
	$email_db = (Get-Mailbox $email | select Database).Database.Name
    "Mailbox database: " + $email_db
    $email_IssueWarningQuota = (Get-MailboxDatabase -Identity $email_db).IssueWarningQuota.Value.ToMB()
    $email_ProhibitSendQuota = (Get-MailboxDatabase -Identity $email_db).ProhibitSendQuota.Value.ToMB()
    $email_ProhibitSendReceiveQuota = (Get-MailboxDatabase -Identity "PRB DB01").ProhibitSendReceiveQuota.Value.ToMB()

    "Квоты заданы на уровне БД " + $email_db + ". " + "Текущие квоты: " | Out-File -FilePath $path -Encoding utf8 -Append
    "Запретить отправку и получение при достижении: " + $email_ProhibitSendReceiveQuota + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
    "Запретить отправку при достижении: " + $email_ProhibitSendQuota + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
    "Выдавать предупреждение при достижении: " + $email_IssueWarningQuota + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
    
}
if (($temp) -eq 0) {
    $email_IssueWarningQuota = (Get-Mailbox -Identity $email).IssueWarningQuota.Value.ToMB()
    $email_ProhibitSendQuota = (Get-Mailbox -Identity $email).ProhibitSendQuota.Value.ToMB()
    $email_ProhibitSendReceiveQuota = (Get-Mailbox -Identity $email).ProhibitSendReceiveQuota.Value.ToMB()
    "Квоты заданы на уровне почтового ящика. " +  "Текущие квоты: " | Out-File -FilePath $path -Encoding utf8 -Append
    "Запретить отправку и получение при достижении: " + $email_ProhibitSendReceiveQuota + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
    "Запретить отправку при достижении: " + $email_ProhibitSendQuota + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
    "Выдавать предупреждение при достижении: " + $email_IssueWarningQuota + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
}

$old_email_data = ""
$old_email_date = ""
$old_email_fold = ""


$items = Get-MailboxFolderStatistics $email -IncludeOldestAndNewestItems  | Sort name | Select Name, Folderpath, ItemsInFolder, FolderSize, OldestItemReceivedDate | Sort OldestItemReceivedDate

$items | Out-File -FilePath 'C:\PATH_TO_DEBUG_FILE\DEBUG.txt' -Encoding utf8 -Append


# Фильтруем папки в почтовом ящике.
if ($items.count -gt 0) {
    foreach ($item in $items) {

        #IF DEBUG
        $item.Name

        $unnecessary_folders_found = $false

        if ($item.OldestItemReceivedDate -eq $null) {
            "NULL DETECTED!!!"
            continue
        }

        switch ($item.Name) {
        
            {$_ -contains 'Deletions'} { 
                '"Deletions" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Recoverable Items'} {
                '"Recoverable Items" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Purges'} {
                '"Purges" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Versions'} {
                '"Versions" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Conversation Action Settings'} {
                '"Conversation Action Settings" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Versions'} {
                
                '"Versions" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Отправленные'} {
                '"Отправленные" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Удаленные'} {
                '"Удаленные" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Исходящие'} {
                '"Исходящие" detected.'
                $unnecessary_folders_found = $true
                break
            }
            
            {$_ -contains 'Календарь'} {
                '"Календарь" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Контакты'} {
                '"Контакты" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Нежелательная почта'} {
                '"Нежелательная почта" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Черновики'} {
                '"Черновики" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Конфликты'} {
                
                '"Конфликты" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Сбои сервера'} {
                
                '"Сбои сервера" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'RSS-подписки'} {
                
                '"Сбои сервера" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Пакет новостей'} {
                
                '"Пакет новостей" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Дневник'} {
                
                '"Дневник" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Предлагаемые контакты'} {
                
                '"Предлагаемые контакты" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Настройка быстрых действий'} {
                
                '"Настройка быстрых действий" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Локальные ошибки'} {
                
                '"Локальные ошибки" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Ошибки синхронизации'} {
                
                '"Ошибки синхронизации" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Задачи'} {
                
                '"Задачи" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Заметки'} {
                
                '"Заметки" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Журнал разговоров'} {
                
                '"Журнал разговоров" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'Контакты Skype для бизнеса'} {
                
                '"Журнал разговоров" detected.'
                $unnecessary_folders_found = $true
                break
            }

                    
       }

       if ($unnecessary_folders_found) {'This is unnecessary folder!!!'; continue}
    
       $old_email_data = $item
       $old_email_date = $item.OldestItemReceivedDate
       $old_email_fold = $item.FolderPath
       break
       '---- Ending loop -----'
       
    }
}

'Самое старое письмо в почтовом ящике датировано: ' +  $old_email_date + '. ' + 'Письмо находится в папке "' + $old_email_fold + '"' | Out-File -FilePath $path -Encoding utf8 -Append

$ad_SamAccountName = (Get-Mailbox $email).SamAccountName
$groups = (Get-ADPrincipalGroupMembership $ad_SamAccountName | Select-Object name)
$mab_groups_detected = $false
$mab_groups_list = ""


# Проверяем членство пользователя в необходимых группа (проверка наличия мобильного доступа).
foreach ($item in $groups) {

    if (($item.name -ne $null) -and (($item.name -contains "GROUP_1") -or ($item.name -contains "GROUP_2") -or ($item.name -contains "GROUP_3") -or ($item.name -contains "GROUP_4"))) {
        "GROUP DETECTED!!!"
        $mab_groups_detected = $true
        $mab_groups_list = $mab_groups_list + $item.name + "; "
        continue
    }
}

if ($mab_groups_detected) {
    'Подключен мобильный доступ. Почтовый ящик является членом следующих групп: "' +  $mab_groups_list + '"' | Out-File -FilePath $path -Encoding utf8 -Append
}
else {
    'Мобильный доступ не подключен.' | Out-File -FilePath $path -Encoding utf8 -Append
}

#https://www.itprotoday.com/powershell/powershell-contains
