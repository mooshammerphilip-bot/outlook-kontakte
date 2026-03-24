# Outlook Kontakte Exporter

Web-App zum Extrahieren von Kontakten aus Outlook-Mails.

## Deploy auf Vercel

### Schritt 1: GitHub Repository erstellen
1. github.com → "New repository" → Name: `outlook-kontakte`
2. Alle Dateien hochladen (oder via git push)

### Schritt 2: Vercel verbinden
1. vercel.com → "New Project" → GitHub Repository auswählen
2. "Deploy" klicken

### Schritt 3: Environment Variables in Vercel setzen
In Vercel → Settings → Environment Variables:

| Variable | Wert |
|---|---|
| AZURE_AD_CLIENT_ID | 25a81fcc-2eca-482d-99a3-75dac0472407 |
| AZURE_AD_CLIENT_SECRET | (dein Client Secret aus Azure) |
| AZURE_AD_TENANT_ID | thisisapex.com |
| NEXTAUTH_SECRET | (zufälliger String von generate-secret.vercel.app/32) |
| NEXTAUTH_URL | https://deine-app.vercel.app |

### Schritt 4: Redirect URI in Azure hinzufügen
Azure Portal → App "Kontakt Export" → Authentifizierung → Redirect URI hinzufügen:
```
https://deine-app.vercel.app/api/auth/callback/azure-ad
```

### Schritt 5: Fertig!
Die App ist jetzt unter deiner Vercel-URL erreichbar.
