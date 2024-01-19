<#
  .SYNOPSIS
  Schickt eine Mail an alle Nutzer, deren Passwörter in den kommenden 14 Tagen ablaufen und eine Übersicht über diese Nutzer an die IT-Abteilung.

  .DESCRIPTION
  Das Get-ADPWExpiryDates-Skript liest alle Nutzer aus dem lokalen Active Directory aus,
  deren Passwörter innerhalb der nächsten 14 Tage ablaufen.
  Im Anschluss bekommt jeder Mitarbeiter eine E-Mail, dass sein Passwort abläuft
  inkl. einer Anleitung zum Ändern des Passworts und den betroffenen Services,
  die bei abgelaufenem Passwort nicht mehr zur Verfügung stehen werden.
  Und es wird eine Gesamtübersicht der ablaufenden Passwörter per Mail an die IT-Abteilung gesendet.

  .INPUTS
  Keine.

  .OUTPUTS
  Keine.

  .NOTES
  Bei den AD-Benutzern muss die E-Mail-Adresse im AD-Profil hinterlegt sein (erste Seite).
  Bitte die 5 Parameter unten in dem ###-Block abändern: ADSearchBase, SmtpServer, ggf. Mailport und die Absenderadresse.
  Hinweis zu der Empfangsadresse: Diese bezieht sich auf dem Empfänger der Gesamtübersicht - nicht die der User,
  daher kann sie identisch mit der Absenderadresse sein.

  .EXAMPLE
  PS> .\Get-ADPWExpiryDates.ps1

  .LINK
  https://github.com/synergy1990/Get-ADPWExpiryDates
#>

### Change the parameters in this block to match your environment
$ADSearchBase         = "OU=MeineUserOU, DC=firma, DC=local"
$SmtpServer           = "meinmailserver.mail.protection.outlook.com"
$Mailport             = "25"
$FromAddress          = "Eure IT <it@firma.de>"
$GeneralMailToAddress = "it@firma.de"
###

$Users                = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties msDS-UserPasswordExpiryTimeComputed, mail, givenName -SearchBase $ADSearchBase
$Date                 = Get-Date
$PWExpiryUsers        = @()
$GeneralMailContent   = "Folgende Passwörter werden in den kommenden 14 Tagen ablaufen:`n`n"

foreach ($u in $Users) {
    $User = New-Object -TypeName PSObject -Property @{
        'Name' = $u.Name
        'PasswordExpiryDate' = [datetime]::FromFileTime($u.'msDS-UserPasswordExpiryTimeComputed')
        'MailAddress' = $u.Mail
        'GivenName' = $u.GivenName
    }

    $DateDiff = $User.PasswordExpiryDate - $Date
            
    if ($DateDiff.Days -le 14) {
        $PWExpiryUsers += $User
    }
}

$PWExpiryUsers = $PWExpiryUsers | Sort-Object -Property PasswordExpiryDate

foreach ($p in $PWExpiryUsers) {
    $GeneralMailContent  += "$($p.Name)`n"
    $GeneralMailContent  += "$($p.PasswordExpiryDate.ToString())`n`n"
    $IndividualMailBody   = "Hallo $($p.GivenName)!`n`nDein Passwort läuft am $($p.PasswordExpiryDate.ToString()) ab.`n"
    $IndividualMailBody  += "Bitte beachte, dass nach Ablauf des Passworts das Office-Portal, Mitarbeiter-WLAN, ERP, VPN uns das Wiki nicht mehr funktionieren werden.`n`n"
    $IndividualMailBody  += "Bitte ändere also vorher dein Passwort!`n"
    $IndividualMailBody  += "Drücke dazu Strg+Alt+Entf und klicke dann auf Kennwort ändern.`n`n"
    $IndividualMailBody  += "Hinweis: Wir werden euch NIEMALS einen Link schicken, den ihr anklicken sollt, UM euer Passwort zu ändern. "
    $IndividualMailBody  += "Derartige Aufforderungen sind allesamt Fake! Im Zweifel immer fragen! :)`n`n"
    $IndividualMailBody  += "Gruß von deiner IT :)"
    Send-MailMessage -SmtpServer $SmtpServer -Port $Mailport -to $p.MailAddress -from $FromAddress -Subject "Passwort läuft bald ab" -Body $IndividualMailBody -Encoding ([System.Text.Encoding]::UTF8)
}

Send-MailMessage -SmtpServer $SmtpServer -Port $Mailport -to $GeneralMailToAddress -from $FromAddress -Subject "Ablaufende Passwörter" -Body $GeneralMailContent -Encoding ([System.Text.Encoding]::UTF8)
