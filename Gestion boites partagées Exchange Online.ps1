# Script permettant de modifier la destination des mails envoyes depuis les boites partagees
# et de définir si oui ou non l'adresse de la boite partagée doit être visible 
#dans le carnet d'adresse global 
# il prend en charge toutes les boites partagées du tenant
# la configuration est pour chaque boite
 
$Module=Get-InstalledModule -Name ExchangeOnlineManagement -MinimumVersion 2.0.3
    if($Module.count -eq 0)
    {
     Write-Host Required Exchange Online'(EXO V2)' module is not available  -ForegroundColor yellow 
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -match "[yY]")
     {
      Install-Module ExchangeOnlineManagement
      Import-Module ExchangeOnlineManagement
     }
     else
     {
      Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet.
     }
   }
Connect-ExchangeOnline

# Install-Module -Name ExchangeOnlineManagement -MinimumVersion 2.0.3
# Import-Module ExchangeOnlineManagement
# Connect-ExchangeOnline

Clear-Host
Get-EXOMailbox -Filter '(RecipientTypeDetails -eq "SharedMailbox")' | Select Alias,HiddenFromAddressListsEnabled,MessageCopyForSentAsEnabled | Foreach-Object {

Write-Host "`r`nXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" -ForegroundColor Green
Write-Host ETAT DE LA BOITE: $_.Alias  
If ($_.HiddenFromAddressListsEnabled -eq $True)
{write-Host  Visible dans le GAL: OUI}
Else 
{write-Host  Visible dans le GAL: NON}
If ($_.MessageCopyForSentAsEnabled -eq $True)
{write-Host  Copie éléments envoyés "en tant que" dans la BAL partagée: OUI}
Else 
{write-Host  Copie éléments envoyés "en tant que" dans la BAL partagée: NON}

$modif= Read-Host Voulez-vous apporter des modifications sur la boite $_Alias  "(O/N)" ?
while("o","n","O","N" -notcontains $modif )
{Read-Host Voulez-vous apporter des modifications sur la BAL partagée $_Alias  "(O/N)" ?}

If ($modif -eq 'o' -Or $modif -eq 'O')
    {
        $modifGAL= Read-Host Affichage de la BAL $_Alias dans le GAL  "(O/N)" ?
        while("o","n","O","N" -notcontains $modif )
        {Read-Host Présence Affichage de la BAL $_Alias dans le GAL  "(O/N)" ?}
        If ($modifGAL -eq 'o' -Or $modifGAL -eq 'O')
        {Set-Mailbox $_.Alias -HiddenFromAddressListsEnabled $true}
        Else
        {Set-Mailbox $_.Alias -HiddenFromAddressListsEnabled $false}
        
        $modifCOPY= Read-Host Copie des messages envoyés dans la BAL partagée $_Alias "(O/N)" ?
        while("o","n","O","N" -notcontains $modifCOPY )
        {Read-Host Copie des messages envoyés dans la BAL partagée $_Alias "(O/N)" ?}
        If ($modifCOPY -eq 'o' -Or $modifCOPY -eq 'O')
        {Set-Mailbox $_.Alias -MessageCopyForSentAsEnabled $true}
        Else
        {Set-Mailbox $_.Alias -MessageCopyForSentAsEnabled $false}
    }
    Else
    {Write-Host AUCUN CHANGEMENT POUR LA BOITE $_Alias}
}
Write-Host "`r`nXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" -ForegroundColor Green
Write-Host "OPERATION TERMINEE"
Write-Host "DECONNEXION DE LA SESSION"
Write-Host "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" -ForegroundColor Green


Disconnect-ExchangeOnline -Confirm:$false
pause
exit
