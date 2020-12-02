Import-Module ActiveDirectory
#add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010

# �������� ������! ��������� ����� ���������� �� ��������� �����
$email = "username@example.com"
$path = "C:\PATH_TO_RESULT_FILE\" + $email + ".txt"

$email_name = (Get-MailboxStatistics $email).DisplayName
$email_total_size = (Get-MailboxStatistics $email).TotalItemSize.Value.ToMB()
$email_total_items = (Get-MailboxStatistics $email).ItemCount

"�������� ����: " + $email + "." | Out-File -FilePath $path -Encoding utf8 -Append
"��� ��������� �����: " + $email_name | Out-File -FilePath $path -Encoding utf8 -Append
"����� ������: " + $email_total_size + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
"����� ��������� (����� � �. �.): " + $email_total_items | Out-File -FilePath $path -Encoding utf8 -Append

$email_db = ""
$email_IssueWarningQuota = ""
$email_ProhibitSendQuota = ""
$email_ProhibitSendReceiveQuota = ""

# ��������� �� ����� ������ ������ ������� ����� ��������� �����.
$temp = ((get-mailbox $email -ResultSize Unlimited).UseDatabaseQuotaDefaults)

if (($temp) -eq 1) {
	$email_db = (Get-Mailbox $email | select Database).Database.Name
    "Mailbox database: " + $email_db
    $email_IssueWarningQuota = (Get-MailboxDatabase -Identity $email_db).IssueWarningQuota.Value.ToMB()
    $email_ProhibitSendQuota = (Get-MailboxDatabase -Identity $email_db).ProhibitSendQuota.Value.ToMB()
    $email_ProhibitSendReceiveQuota = (Get-MailboxDatabase -Identity "PRB DB01").ProhibitSendReceiveQuota.Value.ToMB()

    "����� ������ �� ������ �� " + $email_db + ". " + "������� �����: " | Out-File -FilePath $path -Encoding utf8 -Append
    "��������� �������� � ��������� ��� ����������: " + $email_ProhibitSendReceiveQuota + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
    "��������� �������� ��� ����������: " + $email_ProhibitSendQuota + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
    "�������� �������������� ��� ����������: " + $email_IssueWarningQuota + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
    
}
if (($temp) -eq 0) {
    $email_IssueWarningQuota = (Get-Mailbox -Identity $email).IssueWarningQuota.Value.ToMB()
    $email_ProhibitSendQuota = (Get-Mailbox -Identity $email).ProhibitSendQuota.Value.ToMB()
    $email_ProhibitSendReceiveQuota = (Get-Mailbox -Identity $email).ProhibitSendReceiveQuota.Value.ToMB()
    "����� ������ �� ������ ��������� �����. " +  "������� �����: " | Out-File -FilePath $path -Encoding utf8 -Append
    "��������� �������� � ��������� ��� ����������: " + $email_ProhibitSendReceiveQuota + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
    "��������� �������� ��� ����������: " + $email_ProhibitSendQuota + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
    "�������� �������������� ��� ����������: " + $email_IssueWarningQuota + " MB" | Out-File -FilePath $path -Encoding utf8 -Append
}

$old_email_data = ""
$old_email_date = ""
$old_email_fold = ""


$items = Get-MailboxFolderStatistics $email -IncludeOldestAndNewestItems  | Sort name | Select Name, Folderpath, ItemsInFolder, FolderSize, OldestItemReceivedDate | Sort OldestItemReceivedDate

$items | Out-File -FilePath 'C:\PATH_TO_DEBUG_FILE\DEBUG.txt' -Encoding utf8 -Append


# ��������� ����� � �������� �����.
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

            {$_ -contains '������������'} {
                '"������������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '���������'} {
                '"���������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '���������'} {
                '"���������" detected.'
                $unnecessary_folders_found = $true
                break
            }
            
            {$_ -contains '���������'} {
                '"���������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '��������'} {
                '"��������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '������������� �����'} {
                '"������������� �����" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '���������'} {
                '"���������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '���������'} {
                
                '"���������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '���� �������'} {
                
                '"���� �������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains 'RSS-��������'} {
                
                '"���� �������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '����� ��������'} {
                
                '"����� ��������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '�������'} {
                
                '"�������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '������������ ��������'} {
                
                '"������������ ��������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '��������� ������� ��������'} {
                
                '"��������� ������� ��������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '��������� ������'} {
                
                '"��������� ������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '������ �������������'} {
                
                '"������ �������������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '������'} {
                
                '"������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '�������'} {
                
                '"�������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '������ ����������'} {
                
                '"������ ����������" detected.'
                $unnecessary_folders_found = $true
                break
            }

            {$_ -contains '�������� Skype ��� �������'} {
                
                '"������ ����������" detected.'
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

'����� ������ ������ � �������� ����� ����������: ' +  $old_email_date + '. ' + '������ ��������� � ����� "' + $old_email_fold + '"' | Out-File -FilePath $path -Encoding utf8 -Append

$ad_SamAccountName = (Get-Mailbox $email).SamAccountName
$groups = (Get-ADPrincipalGroupMembership $ad_SamAccountName | Select-Object name)
$mab_groups_detected = $false
$mab_groups_list = ""


# ��������� �������� ������������ � ����������� ������ (�������� ������� ���������� �������).
foreach ($item in $groups) {

    if (($item.name -ne $null) -and (($item.name -contains "GROUP_1") -or ($item.name -contains "GROUP_2") -or ($item.name -contains "GROUP_3") -or ($item.name -contains "GROUP_4"))) {
        "GROUP DETECTED!!!"
        $mab_groups_detected = $true
        $mab_groups_list = $mab_groups_list + $item.name + "; "
        continue
    }
}

if ($mab_groups_detected) {
    '��������� ��������� ������. �������� ���� �������� ������ ��������� �����: "' +  $mab_groups_list + '"' | Out-File -FilePath $path -Encoding utf8 -Append
}
else {
    '��������� ������ �� ���������.' | Out-File -FilePath $path -Encoding utf8 -Append
}

#https://www.itprotoday.com/powershell/powershell-contains
