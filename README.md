# M365 Tenant Manager

Een moderne PowerShell GUI applicatie voor Microsoft 365 tenant beheer via Microsoft Graph API.

## 🚀 Features

- **🎨 Modern Dark UI** — Professionele WPF interface met dark theme
- **📊 Dashboard** — Real-time statistieken van je tenant
- **👤 User Management** — Gebruikers zoeken, filteren en exporteren
- **📈 Bulk Import** — Massaal gebruikers aanmaken vanuit Excel
- **🔑 License Management** — Overzicht van alle licenties
- **👥 Group Management** — Azure AD groepen beheren
- **✅ Dry Run Mode** — Test imports zonder wijzigingen door te voeren
- **📤 Export** — Exporteer data naar Excel

## 📋 Requirements

### PowerShell Modules

```powershell
# Microsoft Graph modules
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser

# Excel module
Install-Module ImportExcel -Scope CurrentUser
```

### Permissions

De applicatie vereist de volgende Microsoft Graph permissions:
- `User.ReadWrite.All`
- `Directory.ReadWrite.All`
- `Organization.Read.All`
- `Group.ReadWrite.All`
- `GroupMember.ReadWrite.All`

## 🚀 Getting Started

1. **Start de applicatie:**
   ```powershell
   .\M365-TenantManager.ps1
   ```

2. **Verbind met M365:**
   - Klik op "Connect to M365"
   - Log in met je admin credentials
   - Accepteer de permissions

3. **Gebruik het dashboard:**
   - Bekijk statistieken
   - Navigeer naar verschillende secties via het menu

## 📊 Bulk Import Users

### Excel Template Format

| Kolom | Beschrijving | Vereist |
|-------|--------------|---------|
| DisplayName | Volledige naam | ✅ |
| UserPrincipalName | Email adres | ✅ |
| MailNickname | Korte alias | ✅ |
| GivenName | Voornaam | ❌ |
| Surname | Achternaam | ❌ |
| JobTitle | Functie | ❌ |
| Department | Afdeling | ❌ |
| OfficeLocation | Kantoor locatie | ❌ |
| MobilePhone | Mobiel nummer | ❌ |
| BusinessPhones | Telefoon | ❌ |
| StreetAddress | Adres | ❌ |
| City | Stad | ❌ |
| State | Provincie | ❌ |
| PostalCode | Postcode | ❌ |
| Country | Land | ❌ |
| UsageLocation | 2-letter land code | ❌ |
| Password | Initieel wachtwoord | ✅ |
| Groups | Groepen (gescheiden door ;) | ❌ |

### Voorbeeld Excel Bestand

Zie `template.xlsx` voor een voorbeeld.

## 🛡️ Security

- ⚠️ Excel bestanden bevatten wachtwoorden — beveilig deze goed
- 🔐 Gebruik altijd "Force password change" voor nieuwe accounts
- 👤 Gebruik least-privilege admin accounts
- 📝 Controleer de Dry Run output voordat je imports uitvoert

## 📝 License

MIT License

## 🙋‍♂️ Support

Voor vragen of issues, maak een GitHub issue aan.
