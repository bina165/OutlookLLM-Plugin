# Outlook LLM Assistant

![Outlook LLM Assistant Logo](assets/logo.svg)

Ein professionelles Outlook Add-in, das LLM-Funktionalitäten (Large Language Model) direkt in Microsoft Outlook integriert. Dieses Add-in verbindet sich mit einem NVIDIA Triton Inference Server, um E-Mail-Kontext zu analysieren und intelligente Aktionen wie Zusammenfassung, Antwortgenerierung und Kalenderdatenextraktion zu ermöglichen.

## Funktionen

- **E-Mail-Analyse**: Extrahiert wichtige Punkte und erforderliche Aktionen aus E-Mails
- **Zusammenfassung**: Fasst lange E-Mails kurz und prägnant zusammen
- **Intelligente Antworten**: Generiert kontextbezogene Antworten basierend auf dem E-Mail-Verlauf
- **Direkte Antwortfunktion**: Antwortet auf E-Mails direkt in Outlook mit LLM-generierten Texten
- **Übersetzung**: Übersetzt E-Mail-Inhalte in verschiedene Sprachen
- **Kalenderdatenextraktion**: Erkennt Terminvorschläge in E-Mails und erstellt Kalendereinträge
- **Benutzerdefinierte Anfragen**: Ermöglicht individuelle Anfragen an das LLM

## Anforderungen

- Microsoft Outlook (Desktop oder Web)
- NVIDIA Triton Inference Server mit konfiguriertem LLM-Modell
- Netzwerkverbindung zwischen Outlook und dem Triton-Server

## Installation

### Für Endbenutzer

1. Laden Sie die neueste Version des Add-ins aus dem [Releases](https://github.com/bina165/OutlookLLM-Plugin/releases)-Bereich herunter
2. Öffnen Sie Outlook und navigieren Sie zu "Datei" > "Optionen" > "Add-ins"
3. Klicken Sie auf "Verwalten" (COM-Add-ins) und dann auf "Durchsuchen"
4. Wählen Sie die heruntergeladene Manifest-Datei aus
5. Starten Sie Outlook neu

### Für Entwickler

1. Klonen Sie das Repository:
   ```
   git clone https://github.com/bina165/OutlookLLM-Plugin.git
   ```

2. Konfigurieren Sie die Triton-Server-Verbindung in `triton-config.json`

3. Testen Sie das Add-in lokal:
   ```
   npm install -g office-addin-dev-certs
   office-addin-dev-certs install
   npm start
   ```

4. Sideloading in Outlook:
   - Öffnen Sie Outlook
   - Navigieren Sie zu "Datei" > "Optionen" > "Add-ins" > "COM-Add-ins" > "Gehe zu..."
   - Klicken Sie auf "Meine Add-ins" > "Benutzerdefiniertes Add-in hinzufügen" > "Manifest aus Datei..."
   - Wählen Sie die Manifest-Datei aus dem Projektverzeichnis aus

## Konfiguration

Die Konfiguration des Add-ins erfolgt über die Datei `triton-config.json`:

```json
{
  "serverConfig": {
    "url": "http://localhost:8000",
    "apiVersion": "v2",
    "timeout": 30000,
    "maxRetries": 3,
    "retryDelay": 1000,
    "debug": false
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

## Architektur

Das Add-in besteht aus folgenden Hauptkomponenten:

1. **UI-Komponenten** (`taskpane.html`, `taskpane.css`): Benutzeroberfläche des Add-ins
2. **Triton-Connector** (`triton-connector.js`): Kommunikation mit dem NVIDIA Triton Inference Server
3. **E-Mail-Kontext-Extraktor** (`email-context-extractor.js`): Extraktion und Aufbereitung von E-Mail-Daten
4. **LLM-Aktionen** (`llm-actions.js`): Implementierung der verschiedenen LLM-gesteuerten Aktionen
5. **Direct-Reply-Handler** (`direct-reply-handler.js`): Direkte Antwortfunktion mit E-Mail-Thread-Kontext
6. **Initialisierung** (`initialize.js`): Verbindung aller Komponenten und Event-Handling

## Sicherheit

- Das Add-in läuft in einer Sandbox-Umgebung
- Die Kommunikation mit dem Triton-Server kann über HTTPS abgesichert werden
- API-Schlüssel können für die Authentifizierung konfiguriert werden
- Keine Speicherung von E-Mail-Daten außerhalb des Add-ins

## Beitragen

Beiträge sind willkommen! Bitte lesen Sie [CONTRIBUTING.md](CONTRIBUTING.md) für Details zum Prozess für Pull Requests.

## Lizenz

Dieses Projekt ist unter der MIT-Lizenz lizenziert - siehe [LICENSE](LICENSE) für Details.

## Kontakt

Bei Fragen oder Problemen erstellen Sie bitte ein [Issue](https://github.com/bina165/OutlookLLM-Plugin/issues) oder kontaktieren Sie den Projektbetreuer.

---

Entwickelt mit ❤️ für effizientere E-Mail-Kommunikation