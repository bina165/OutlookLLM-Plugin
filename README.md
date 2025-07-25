# OutlookLLM-Plugin

Ein Outlook Web Add-in, das ein Large Language Model (LLM) direkt in Outlook integriert, um kontextbewusste E-Mail-Analyse, -Zusammenfassung, -Übersetzung und -Antworten zu ermöglichen.

## Überblick

Das OutlookLLM-Plugin verbindet Outlook direkt mit einem NVIDIA Triton Inference Server, der Ihr bevorzugtes LLM hostet. Das Plugin ermöglicht es Ihnen, das LLM für verschiedene E-Mail-bezogene Aufgaben zu nutzen, einschließlich:

- Analyse von E-Mail-Inhalten
- Erstellung von Zusammenfassungen
- Generierung von kontextbewussten Antworten
- Übersetzung von E-Mails
- Extraktion von Kalenderdaten
- Benutzerdefinierte LLM-Anfragen

## Funktionen

- **Kontextbewusste Antworten**: Berücksichtigt den gesamten E-Mail-Thread für relevante Antworten
- **Direkte Integration**: Arbeitet nahtlos innerhalb der Outlook-Benutzeroberfläche
- **Verschiedene Antwortstile**: Formell, freundlich, kurz oder detailliert
- **Kalenderintegration**: Extrahiert Termindetails und erstellt Kalendereinträge
- **Konfigurierbare Prompts**: Anpassbare Vorlagen für verschiedene LLM-Aktionen
- **Sichere Verbindung**: Unterstützt API-Schlüssel und TLS für die Kommunikation mit dem Triton-Server

## Installation

### Voraussetzungen

- Microsoft Outlook (Web, Windows oder Mac)
- Zugang zu einem NVIDIA Triton Inference Server mit einem konfigurierten LLM
- Node.js und npm (für die lokale Entwicklung)

### Einrichtung für Entwickler

1. Klonen Sie dieses Repository:
   ```
   git clone https://github.com/bina165/OutlookLLM-Plugin.git
   cd OutlookLLM-Plugin
   ```

2. Installieren Sie einen lokalen Webserver (z.B. mit Node.js):
   ```
   npm install -g http-server
   ```

3. Starten Sie den Webserver im HTTPS-Modus:
   ```
   http-server -S -C cert.pem -K key.pem -p 3000
   ```
   
   Hinweis: Für die Entwicklung müssen Sie selbstsignierte Zertifikate erstellen. Tools wie [mkcert](https://github.com/FiloSottile/mkcert) können hierbei helfen.

4. Konfigurieren Sie die Verbindung zum Triton-Server:
   - Kopieren Sie `triton-config.example.json` zu `triton-config.json`
   - Passen Sie die Server-URL, den API-Schlüssel und andere Parameter nach Bedarf an

### Sideloading in Outlook

1. Öffnen Sie Outlook im Web oder Desktop
2. Gehen Sie zu Einstellungen > Erweiterungen verwalten
3. Wählen Sie "Benutzerdefinierte Add-ins" > "Meine Add-ins" > "Add-in aus Datei hinzufügen"
4. Wählen Sie die `manifest.xml` Datei aus diesem Repository
5. Folgen Sie den Anweisungen, um die Installation abzuschließen

## Verwendung

Nach der Installation erscheint ein neuer Button "LLM Assistant" in der Outlook-Oberfläche. Klicken Sie darauf, um das Plugin zu öffnen.

### E-Mail analysieren
1. Öffnen Sie eine E-Mail
2. Klicken Sie auf "E-Mail analysieren"
3. Das LLM extrahiert die wichtigsten Punkte und erforderlichen Aktionen

### E-Mail zusammenfassen
1. Öffnen Sie eine E-Mail
2. Klicken Sie auf "Zusammenfassen"
3. Erhalten Sie eine prägnante Zusammenfassung des Inhalts

### Antwort generieren
1. Öffnen Sie eine E-Mail
2. Klicken Sie auf "Antwort generieren"
3. Das LLM erstellt eine kontextbewusste Antwort, die Sie bearbeiten oder direkt verwenden können

### Direkt antworten
1. Öffnen Sie eine E-Mail
2. Klicken Sie auf "Direkt antworten"
3. Das LLM öffnet ein Antwortformular und fügt automatisch eine generierte Antwort ein

### Antwort-Stil wählen
1. Öffnen Sie eine E-Mail
2. Klicken Sie auf "Antwort-Stil"
3. Wählen Sie zwischen formell, freundlich, kurz oder detailliert
4. Das LLM generiert eine Antwort im gewählten Stil

### E-Mail übersetzen
1. Öffnen Sie eine E-Mail
2. Klicken Sie auf "Übersetzen"
3. Geben Sie die Zielsprache ein
4. Erhalten Sie eine Übersetzung des E-Mail-Inhalts

### Termin erstellen
1. Öffnen Sie eine E-Mail mit Termininformationen
2. Klicken Sie auf "Termin erstellen"
3. Das LLM extrahiert Datum, Uhrzeit, Teilnehmer und andere Details
4. Überprüfen Sie die Informationen und klicken Sie auf "Kalendereintrag erstellen"

### Benutzerdefinierte Anfrage
1. Öffnen Sie eine E-Mail
2. Klicken Sie auf "Benutzerdefiniert"
3. Geben Sie Ihre eigene Anweisung für das LLM ein
4. Erhalten Sie eine maßgeschneiderte Antwort basierend auf Ihrer Anfrage

## Konfiguration

Die Datei `triton-config.json` enthält alle Einstellungen für die Verbindung zum Triton-Server und die LLM-Parameter:

```json
{
  "serverConfig": {
    "url": "http://localhost:8000",
    "apiVersion": "v2",
    "timeout": 30000,
    "maxRetries": 3,
    "retryDelay": 1000,
    "debug": false,
    "apiKey": "",
    "useTLS": false
  },
  "modelConfig": {
    "name": "llm_model",
    "defaultParameters": {
      "max_tokens": 1024,
      "temperature": 0.7,
      "top_p": 0.9,
      "stop_sequences": ["\n###", "###", "</answer>"],
      "return_full_text": false
    }
  },
  "promptTemplates": {
    "analyze": "Analysiere diese E-Mail und gib mir die wichtigsten Punkte und erforderlichen Aktionen:\n\n{email_context}",
    "summarize": "Fasse diese E-Mail kurz und prägnant zusammen:\n\n{email_context}",
    "reply": "Generiere eine professionelle Antwort auf diese E-Mail:\n\n{email_context}",
    "translate": "Übersetze diese E-Mail ins {target_language}:\n\n{email_context}",
    "calendar": "Extrahiere Informationen für einen Kalendereintrag aus dieser E-Mail (Datum, Uhrzeit, Teilnehmer, Ort, Thema) und formatiere sie als JSON:\n\n{email_context}",
    "custom": "{custom_prompt}\n\n{email_context}"
  }
}
```

## Sicherheitshinweise

- Verwenden Sie immer TLS für die Produktionsumgebung
- Schützen Sie Ihren API-Schlüssel
- Überprüfen Sie die Berechtigungen des Add-ins in Outlook
- Stellen Sie sicher, dass Ihre Datenverarbeitungsrichtlinien die Übermittlung von E-Mail-Inhalten an externe LLM-Server erlauben

## Fehlerbehebung

### Das Plugin wird nicht geladen
- Überprüfen Sie, ob der Webserver läuft und über HTTPS erreichbar ist
- Stellen Sie sicher, dass die URLs im Manifest korrekt sind
- Prüfen Sie die Browser-Konsole auf JavaScript-Fehler

### Keine Verbindung zum Triton-Server
- Überprüfen Sie die Server-URL und den API-Schlüssel
- Stellen Sie sicher, dass der Triton-Server läuft und erreichbar ist
- Prüfen Sie die Netzwerkeinstellungen und Firewalls

### LLM-Antworten sind nicht wie erwartet
- Passen Sie die Prompt-Vorlagen in der Konfigurationsdatei an
- Experimentieren Sie mit verschiedenen Temperatur- und Top-P-Werten
- Stellen Sie sicher, dass das richtige Modell auf dem Triton-Server konfiguriert ist

## Lizenz

Dieses Projekt steht unter der MIT-Lizenz - siehe die [LICENSE](LICENSE) Datei für Details.

## Beitragen

Beiträge sind willkommen! Bitte lesen Sie [CONTRIBUTING.md](CONTRIBUTING.md) für Details zum Prozess für Pull-Requests.