# M365 Bulk User Import

Een moderne PowerShell GUI applicatie voor het bulk importeren van gebruikers in Microsoft 365 via Microsoft Graph API.

## 🚀 Features

- **🎨 Modern Dark UI** — Professionele WPF interface
- **📊 Excel Import** — Importeer gebruikers vanuit Excel bestanden
- **🔑 License Assignment** — Automatisch Microsoft 365 licenties toewijzen
- **👥 Group Membership** — Gebruikers toevoegen aan Azure AD groepen
- **✅ Dry Run Mode** — Test imports zonder wijzigingen door te voeren
- **📝 Logging** — Real-time resultaten en voortgang

## 📋 Requirements

### PowerShell Modules

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
```

### Permissions

- `User.ReadWrite.All`
- `Directory.ReadWrite.All`
- `Organization.Read.All`
- `Group.Read.All`
- `GroupMember.ReadWrite.All`

## 🚀 Getting Started

1. **Start de applicatie:**
   ```powershell
   .\M365-BulkImport.ps1
   ```

2. **Verbind met M365:**
   - Klik op "Connect to M365"
   - Log in met je admin credentials

3. **Selecteer Excel bestand:**
   - Klik op "Browse..." om je Excel bestand te selecteren
   - Klik op "Preview Data" om de inhoud te bekijken

4. **Configureer opties:**
   - Force password change: Gebruikers moeten wachtwoord wijzigen bij eerste login
   - Assign license: Selecteer een Microsoft 365 licentie
   - Dry Run: Alleen preview, geen daadwerkelijke wijzigingen

5. **Importeer gebruikers:**
   - Klik op "Import Users"

## 📊 Excel Template Format

| Kolom | Beschrijving | Vereist |
|-------|--------------|---------|
| DisplayName | Volledige naam | ✅ Ja |
| UserPrincipalName | Email adres | ✅ Ja |
| MailNickname | Korte alias | ✅ Ja |
| GivenName | Voornaam | ❌ Nee |
| Surname | Achternaam | ❌ Nee |
| JobTitle | Functie | ❌ Nee |
| Department | Afdeling | ❌ Nee |
| OfficeLocation | Kantoor locatie | ❌ Nee |
| MobilePhone | Mobiel nummer | ❌ Nee |
| BusinessPhones | Telefoon | ❌ Nee |
| StreetAddress | Adres | ❌ Nee |
| City | Stad | ❌ Nee |
| State | Provincie | ❌ Nee |
| PostalCode | Postcode | ❌ Nee |
| Country | Land | ❌ Nee |
| UsageLocation | 2-letter land code (bijv. NL, US) | ❌ Nee |
| Password | Initieel wachtwoord | ✅ Ja |
| Groups | Groepen, gescheiden door puntkomma (;) | ❌ Nee |

## Voorbeeld Excel

Zie `template.csv` voor een voorbeeld. Sla op als `.xlsx` formaat.

## 🛡️ Security Tips

- ⚠️ Excel bestanden bevatten wachtwoorden in platte tekst
- 🔐 Gebruik altijd "Force password change" voor nieuwe accounts
- 👤 Gebruik een admin account met minimaal benodigde rechten
- 📝 Controleer altijd de Dry Run output vooraf
- 🔄 Laat gebruikers hun wachtwoord wijzigen na eerste login

## 📝 License

MIT License
