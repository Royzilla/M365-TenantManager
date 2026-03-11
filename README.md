# M365 Bulk User Import

PowerShell GUI tool voor bulk importeren van gebruikers naar Microsoft 365 met uitgebreide validatie en rapportage.

## 🚀 Features

### Data Validatie
- ✅ **Verplichte velden check** — Controleert of alle vereiste velden aanwezig zijn
- ✅ **Email formaat validatie** — Controleert geldigheid van email adressen
- ✅ **Wachtwoord sterkte check** — Waarschuwt voor zwakke wachtwoorden
- ✅ **Duplicaten detectie** — Controleert op dubbele gebruikers in Excel en Azure AD
- ✅ **Bestaande gebruikers check** — Waarschuwt als gebruiker al bestaat

### Import Features
- 📊 **Excel import** — Importeer vanuit .xlsx bestanden
- 🔑 **Automatische licentie toewijzing** — Microsoft 365 licenties toewijzen
- 👥 **Groep membership** — Automatisch toevoegen aan Azure AD groepen
- ✅ **Dry Run mode** — Test imports zonder wijzigingen
- 📝 **Gedetailleerde logging** — Exporteer resultaten naar CSV

### UI Features
- 🎨 **Modern dark theme** — Professionele interface
- 📥 **Template download** — Genereer voorbeeld Excel bestand
- 👁️ **Data preview** — Bekijk data voor import
- 🔍 **Validatie rapport** — Zie alle errors en waarschuwingen
- 💾 **Log export** — Sla resultaten op

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
   - Klik op "Browse..." of "Download Template"
   - Data wordt automatisch geladen

4. **Valideer data:**
   - Klik op "Validate Data"
   - Controleer errors en waarschuwingen

5. **Importeer:**
   - Configureer opties
   - Klik op "Import Users"

## 📊 Excel Template Format

| Kolom | Beschrijving | Vereist | Validatie |
|-------|--------------|---------|-----------|
| DisplayName | Volledige naam | ✅ Ja | Mag niet leeg zijn |
| UserPrincipalName | Email adres | ✅ Ja | Valid email formaat |
| MailNickname | Korte alias | ✅ Ja | Mag niet leeg zijn |
| GivenName | Voornaam | ❌ Nee | - |
| Surname | Achternaam | ❌ Nee | - |
| JobTitle | Functie | ❌ Nee | - |
| Department | Afdeling | ❌ Nee | - |
| OfficeLocation | Kantoor locatie | ❌ Nee | - |
| MobilePhone | Mobiel nummer | ❌ Nee | - |
| BusinessPhones | Telefoon | ❌ Nee | - |
| StreetAddress | Adres | ❌ Nee | - |
| City | Stad | ❌ Nee | - |
| State | Provincie | ❌ Nee | - |
| PostalCode | Postcode | ❌ Nee | - |
| Country | Land | ❌ Nee | - |
| UsageLocation | 2-letter code (NL, US, etc.) | ❌ Nee | - |
| Password | Initieel wachtwoord | ✅ Ja | Min. 8 tekens, mixed case, cijfers |
| Groups | Groepen (; gescheiden) | ❌ Nee | - |

## 🔍 Validatie Checks

De tool controleert:

1. **Verplichte velden** — DisplayName, UserPrincipalName, MailNickname, Password
2. **Email formaat** — Moet geldig email formaat hebben
3. **Wachtwoord sterkte** — Minimaal 3 van 4:
   - 8+ tekens
   - Hoofdletters (A-Z)
   - Kleine letters (a-z)
   - Cijfers (0-9)
   - Speciale tekens (!@#$% etc.)
4. **Duplicaten in Excel** — Geen dubbele UserPrincipalName
5. **Bestaande gebruikers** — Checkt Azure AD voor bestaande accounts

## 🛡️ Security Tips

- ⚠️ **Bescherm je Excel bestanden** — Bevatten wachtwoorden in platte tekst
- 🔐 **Gebruik sterke wachtwoorden** — Minimaal 8 tekens met mixed case + cijfers + symbolen
- 🔄 **Force password change** — Laat gebruikers wachtwoord wijzigen bij eerste login
- 👤 **Gebruik least-privilege admin** — Admin account met minimale rechten
- 📝 **Review validatie resultaten** — Controleer alle errors voor import
- 💾 **Sla logs op** — Bewaar import logs voor audit

## 📝 License

MIT License
