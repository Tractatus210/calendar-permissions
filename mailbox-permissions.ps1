$all_relevant_mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox,RoomMailbox,EquipmentMailbox,TeamMailbox,SharedMailbox
$user_mailboxes = @{}
$relevant_mailboxes = New-Object System.Collections.Generic.List[Object]

foreach($mailbox in $all_relevant_mailboxes){
    if($mailbox.RecipientTypeDetails -eq "UserMailbox"){
        $user_mailboxes[$mailbox.PrimarySmtpAddress] = $true
    }
    $relevant_mailboxes.Add($mailbox.PrimarySmtpAddress)
}

[string]$overwrite_input
[bool]$overwrite_existing

$get_overwrite_input = {
    $overwrite_input = Read-Host "Do you want to overwrite existing permissions? y/n"
    return $overwrite_input
}

& $get_overwrite_input
if($overwrite_input -eq ("y" -or "Y")){
    $overwrite_existing = $true
}elseif($overwrite_input -eq ("n" -or "N")){
    $overwrite_existing = $false
}else{
    & $get_overwrite_input
}

$get_baseline_permission = {
    $baseline_permission = Read-Host "Do you want to set the baseline permissions - the access that all user mailboxes have for all other mailboxes - to Reviewer (R) or Editor (E)"
    return $baseline_permission
}

& $get_baseline_permission

$acceptable_baseline_permissions = @("R","r","E","e")
if($baseline_permission -notin $acceptable_baseline_permissions){
    & $get_baseline_permission
}

[string]$permission_to_set
if($baseline_permission -eq ("E" -or "e")){
    $permission_to_set = "Editor"
}else{
    $permission_to_set = "Reviewer"
}

$relevant_mailboxes | ForEach-Object -Parallel{
    $relevant_mailbox = $_
    
    foreach($user_mailbox in $using:user_mailboxes.Keys){
        if($user_mailbox -ne $relevant_mailbox){
            if($overwrite_existing -eq $true){
                Set-MailboxFolderPermission -Identity "${relevant_mailbox}:\Calendar" -User $user_mailbox -AccessRights $permission_to_set -ErrorAction SilentlyContinue | Out-Null
            }else{
                Add-MailboxFolderPermission -Identity "${relevant_mailbox}:\Calendar" -User $user_mailbox -AccessRights $permission_to_set -ErrorAction SilentlyContinue | Out-Null
            }
        }      
    }   -ThrottleLimit 5

