/**
 * LLM Actions für Outlook
 * 
 * Diese Klasse implementiert die verschiedenen LLM-gesteuerten Aktionen,
 * die in Outlook ausgeführt werden können, wie z.B. E-Mails analysieren,
 * Antworten generieren, Termine erstellen, etc.
 */

class LLMActions {
    /**
     * Initialisiert die LLM Actions
     * 
     * @param {TritonConnector} tritonConnector - Der Triton-Connector für die Kommunikation mit dem LLM
     * @param {EmailContextExtractor} contextExtractor - Der E-Mail-Kontext-Extraktor
     * @param {Object} config - Die Konfiguration für die LLM-Aktionen
     */
    constructor(tritonConnector, contextExtractor, config = {}) {
        this.tritonConnector = tritonConnector;
        this.contextExtractor = contextExtractor;
        this.config = {
            promptTemplates: config.promptTemplates || {},
            defaultParameters: config.defaultParameters || {
                max_tokens: 1024,
                temperature: 0.7,
                top_p: 0.9
            },
            ...config
        };
        
        // Event-Handler
        this.onBeforeAction = null;
        this.onAfterAction = null;
        this.onError = null;
    }

    /**
     * Führt eine LLM-Aktion aus
     * 
     * @param {string} actionType - Der Typ der Aktion (analyze, summarize, reply, etc.)
     * @param {Office.Item} item - Das Outlook-Element (E-Mail, Termin, etc.)
     * @param {Object} options - Zusätzliche Optionen für die Aktion
     * @returns {Promise<Object>} - Das Ergebnis der Aktion
     */
    async executeAction(actionType, item, options = {}) {
        try {
            // Event vor der Aktion auslösen
            if (this.onBeforeAction) {
                this.onBeforeAction(actionType, item, options);
            }
            
            // Kontext extrahieren
            const context = await this.contextExtractor.extractContext(item);
            
            // Prompt erstellen
            const prompt = this.createPrompt(actionType, context, options);
            
            // Parameter für die LLM-Anfrage erstellen
            const parameters = {
                ...this.config.defaultParameters,
                ...options.parameters
            };
            
            // LLM-Anfrage senden
            const response = await this.tritonConnector.generateText(prompt, parameters);
            
            // Ergebnis verarbeiten
            const result = this.processResponse(actionType, response, context, options);
            
            // Event nach der Aktion auslösen
            if (this.onAfterAction) {
                this.onAfterAction(actionType, item, result);
            }
            
            return result;
        } catch (error) {
            console.error(`Fehler bei der Ausführung der Aktion ${actionType}:`, error);
            
            // Fehler-Event auslösen
            if (this.onError) {
                this.onError(actionType, error);
            }
            
            throw error;
        }
    }

    /**
     * Erstellt einen Prompt für eine LLM-Aktion
     * 
     * @param {string} actionType - Der Typ der Aktion
     * @param {Object} context - Der extrahierte Kontext
     * @param {Object} options - Zusätzliche Optionen
     * @returns {string} - Der erstellte Prompt
     */
    createPrompt(actionType, context, options = {}) {
        // Formatiere den Kontext für das LLM
        const formattedContext = this.contextExtractor.formatContextForLLM(context);
        
        // Hole die Prompt-Vorlage für den Aktionstyp
        let promptTemplate = this.config.promptTemplates[actionType];
        
        // Fallback auf benutzerdefinierten Prompt, wenn keine Vorlage gefunden wurde
        if (!promptTemplate && options.customPrompt) {
            promptTemplate = options.customPrompt;
        }
        
        // Fallback auf Standard-Prompt, wenn keine Vorlage und kein benutzerdefinierter Prompt gefunden wurde
        if (!promptTemplate) {
            promptTemplate = "Analysiere den folgenden Inhalt:\n\n{email_context}";
        }
        
        // Ersetze Platzhalter im Prompt
        let prompt = promptTemplate.replace("{email_context}", formattedContext);
        
        // Ersetze weitere Platzhalter
        if (options.targetLanguage) {
            prompt = prompt.replace("{target_language}", options.targetLanguage);
        }
        
        if (options.customPrompt) {
            prompt = prompt.replace("{custom_prompt}", options.customPrompt);
        }
        
        return prompt;
    }

    /**
     * Verarbeitet die Antwort des LLM
     * 
     * @param {string} actionType - Der Typ der Aktion
     * @param {Object} response - Die Antwort des LLM
     * @param {Object} context - Der extrahierte Kontext
     * @param {Object} options - Zusätzliche Optionen
     * @returns {Object} - Das verarbeitete Ergebnis
     */
    processResponse(actionType, response, context, options = {}) {
        // Extrahiere den Text aus der Antwort
        let text = "";
        if (response && response.responses && response.responses.length > 0) {
            text = response.responses[0].text;
        } else if (typeof response === "string") {
            text = response;
        }
        
        // Verarbeite die Antwort je nach Aktionstyp
        switch (actionType) {
            case "calendar":
                return this.processCalendarResponse(text, context);
            case "translate":
                return this.processTranslationResponse(text, context, options.targetLanguage);
            default:
                return {
                    type: actionType,
                    text: text,
                    context: context
                };
        }
    }

    /**
     * Verarbeitet die Antwort des LLM für eine Kalenderaktion
     * 
     * @param {string} text - Der Text der Antwort
     * @param {Object} context - Der extrahierte Kontext
     * @returns {Object} - Das verarbeitete Ergebnis
     */
    processCalendarResponse(text, context) {
        try {
            // Versuche, JSON aus der Antwort zu extrahieren
            const jsonMatch = text.match(/```json\s*([\s\S]*?)\s*```/) || 
                             text.match(/\{[\s\S]*\}/);
            
            let eventData = {};
            if (jsonMatch) {
                eventData = JSON.parse(jsonMatch[1] || jsonMatch[0]);
            }
            
            return {
                type: "calendar",
                text: text,
                context: context,
                eventData: eventData
            };
        } catch (error) {
            console.error("Fehler beim Verarbeiten der Kalenderantwort:", error);
            return {
                type: "calendar",
                text: text,
                context: context,
                error: "Konnte keine gültigen Kalenderdaten extrahieren"
            };
        }
    }

    /**
     * Verarbeitet die Antwort des LLM für eine Übersetzungsaktion
     * 
     * @param {string} text - Der Text der Antwort
     * @param {Object} context - Der extrahierte Kontext
     * @param {string} targetLanguage - Die Zielsprache
     * @returns {Object} - Das verarbeitete Ergebnis
     */
    processTranslationResponse(text, context, targetLanguage) {
        return {
            type: "translate",
            text: text,
            context: context,
            sourceLanguage: "Deutsch", // Annahme: Quellsprache ist Deutsch
            targetLanguage: targetLanguage
        };
    }

    /**
     * Analysiert eine E-Mail mit dem LLM
     * 
     * @param {Office.Item} item - Das Outlook-Element
     * @returns {Promise<Object>} - Das Ergebnis der Analyse
     */
    async analyzeEmail(item) {
        return this.executeAction("analyze", item);
    }

    /**
     * Fasst eine E-Mail mit dem LLM zusammen
     * 
     * @param {Office.Item} item - Das Outlook-Element
     * @returns {Promise<Object>} - Das Ergebnis der Zusammenfassung
     */
    async summarizeEmail(item) {
        return this.executeAction("summarize", item);
    }

    /**
     * Generiert eine Antwort auf eine E-Mail mit dem LLM
     * 
     * @param {Office.Item} item - Das Outlook-Element
     * @returns {Promise<Object>} - Das Ergebnis der Antwortgenerierung
     */
    async generateReply(item) {
        return this.executeAction("reply", item);
    }

    /**
     * Übersetzt eine E-Mail mit dem LLM
     * 
     * @param {Office.Item} item - Das Outlook-Element
     * @param {string} targetLanguage - Die Zielsprache
     * @returns {Promise<Object>} - Das Ergebnis der Übersetzung
     */
    async translateEmail(item, targetLanguage) {
        return this.executeAction("translate", item, { targetLanguage });
    }

    /**
     * Extrahiert Kalenderdaten aus einer E-Mail mit dem LLM
     * 
     * @param {Office.Item} item - Das Outlook-Element
     * @returns {Promise<Object>} - Das Ergebnis der Kalenderdatenextraktion
     */
    async extractCalendarData(item) {
        return this.executeAction("calendar", item);
    }

    /**
     * Führt eine benutzerdefinierte Aktion mit dem LLM aus
     * 
     * @param {Office.Item} item - Das Outlook-Element
     * @param {string} customPrompt - Der benutzerdefinierte Prompt
     * @returns {Promise<Object>} - Das Ergebnis der benutzerdefinierten Aktion
     */
    async executeCustomAction(item, customPrompt) {
        return this.executeAction("custom", item, { customPrompt });
    }

    /**
     * Erstellt einen Kalendereintrag in Outlook basierend auf extrahierten Daten
     * 
     * @param {Object} eventData - Die extrahierten Kalenderdaten
     * @returns {Promise<boolean>} - true, wenn der Kalendereintrag erfolgreich erstellt wurde
     */
    async createOutlookAppointment(eventData) {
        return new Promise((resolve, reject) => {
            try {
                // Erstelle ein neues Appointment-Item
                Office.context.mailbox.displayNewAppointmentForm({
                    subject: eventData.subject || "Neuer Termin",
                    start: new Date(eventData.start) || new Date(),
                    end: new Date(eventData.end) || new Date(Date.now() + 3600000), // +1 Stunde
                    location: eventData.location || "",
                    body: eventData.description || "",
                    attendees: eventData.attendees || []
                });
                
                resolve(true);
            } catch (error) {
                console.error("Fehler beim Erstellen des Kalendereintrags:", error);
                reject(error);
            }
        });
    }

    /**
     * Fügt eine Antwort in eine E-Mail ein
     * 
     * @param {string} text - Der einzufügende Text
     * @returns {Promise<boolean>} - true, wenn der Text erfolgreich eingefügt wurde
     */
    async insertTextIntoEmail(text) {
        return new Promise((resolve, reject) => {
            try {
                // Prüfen, ob wir im Compose-Modus sind
                if (Office.context.mailbox.item.body && Office.context.mailbox.item.body.setSelectedDataAsync) {
                    // Text einfügen
                    Office.context.mailbox.item.body.setSelectedDataAsync(
                        text,
                        { coercionType: Office.CoercionType.Text },
                        (result) => {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                resolve(true);
                            } else {
                                reject(new Error(result.error.message));
                            }
                        }
                    );
                } else {
                    // Lese-Modus: Antwortformular öffnen
                    if (Office.context.mailbox.item.displayReplyForm) {
                        Office.context.mailbox.item.displayReplyForm(text);
                        resolve(true);
                    } else {
                        reject(new Error("Konnte Text nicht einfügen: Nicht im Compose-Modus und keine Reply-Funktion verfügbar"));
                    }
                }
            } catch (error) {
                console.error("Fehler beim Einfügen des Textes:", error);
                reject(error);
            }
        });
    }
}

// Exportiere die Klasse für die Verwendung in anderen Modulen
if (typeof module !== 'undefined' && module.exports) {
    module.exports = LLMActions;
}