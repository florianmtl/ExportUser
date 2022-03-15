#Connexion à l'exchange local
$userCred = Get-Credential
$session = New-PSSession -Configurationname Microsoft.Exchange –ConnectionUri http://vsrvmail/powershell -Authentication Kerberos -Credential $userCred
Import-PSSession $session

#Recuperation de la liste des mail, stocké dans une variable, sur l'exchange local. Pour chaque utilisateur, extrait de son nom et prenom, de son email, de sa taille de boite aux lettres et indiquation si il est migré ou non
$List = get-mailbox -ResultSize unlimited | foreach { 

    $MailUser = $_.UserPrincipalName 

    $stats= Get-MailboxStatistics $MailUser

    New-Object -TypeName PSObject -Property @{
    'Prénom Nom' = $stats.DisplayName
    Email = $MailUser
    "Taille boîte aux lettres" = $stats.TotalItemSize
    Migré = "Non"
}

}

#Fermeture de la session courante pour pouvoir passer à la connection à Office 365
Remove-PSSession -session $session

#Connexion à l'office 365
$Cred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Cred -Authentication Basic –AllowRedirection
Import-PSSession $Session
Import-Module msonline
Get-Command –Module msonline
Connect-msolservice -Credential $Cred

#Ajout à la liste précedente les mêmes informations mais côté Office 365
$List += get-mailbox -ResultSize unlimited | foreach { 

    $MailUser = $_.UserPrincipalName 

    $stats= Get-MailboxStatistics $MailUser

    New-Object -TypeName PSObject -Property @{
    'Prénom Nom' = $stats.DisplayName
    Email = $MailUser
    "Taille boîte aux lettres" = $stats.TotalItemSize
    Migré = "Oui"
}

}
#Export de la liste en format CSV
$List | Export-Csv "C:\Script-FM\exportMailUser.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8

Remove-PSSession -session $Session