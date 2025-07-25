/**
 * Direct Reply Handler
 * 
 * Diese Klasse implementiert die Funktionalität, um direkt in Outlook eine E-Mail mit Hilfe des LLMs zu beantworten,
 * wobei der gesamte E-Mail-Verlauf als Kontext berücksichtigt wird.
 */

class DirectReplyHandler {
    /**
     * Initialisiert den Direct Reply Handler
     * 
     * @param {TritonConnector} tritonConnector - Der Triton-Connector für die Kommunikation mit dem LLM
     * @param {EmailContextExtractor} contextExtractor - Der E-Mail-Kontext-Extraktor
     * @param {Object} config - Die Konfiguration für den Direct Reply Handler
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
            includeThread: config.includeThread !== undefined ? config.includeThread : true,
            ...config
        };
        
        // Event-Handler
        this.onBeforeReply = null;
        this.onAfterReply = null;
        this.onError = null;
    }

    /**
     * Generiert eine Antwort auf eine E-Mail und fügt sie direkt in das Antwortformular ein
     * 
     * @param {Office.MessageItem} item - Die E-Mail, auf die geantwortet werden soll
     * @param {Object} options - Zusätzliche Optionen für die Antwort
     * @returns {Promise<boolean>} - true, wenn die Antwort erfolgreich generiert und eingefügt wurde
     */
    async replyToEmail(item, options = {}) {
        try {
            // Event vor der Antwort auslösen
            if (this.onBeforeReply) {
                this.onBeforeReply(item, options);
            }
            
            // Kontext extrahieren (mit Thread)
            const contextOptions = { ...this.contextExtractor.options };
            this.contextExtractor.options.includeThread = this.config.includeThread;
            const context = await this.contextExtractor.extractContext(item);
            this.contextExtractor.options = contextOptions; // Ursprüngliche Optionen wiederherstellen
            
            // Prompt erstellen
            const prompt = this.createReplyPrompt(context, options);
            
            // Parameter für die LLM-Anfrage erstellen
            const parameters = {
                ...this.config.defaultParameters,
                ...options.parameters
            };
            
            // LLM-Anfrage senden
            const response = await this.tritonConnector.generateText(prompt, parameters);
            
            // Antwort extrahieren
            let replyText = "";
            if (response && response.responses && response.responses.length > 0) {
                replyText = response.responses[0].text;
            } else if (typeof response === "string") {
                replyText = response;
            }
            
            // Antwort in das Antwortformular einfügen
            const success = await this.insertReplyIntoEmail(item, replyText, options);
            
            // Event nach der Antwort auslösen
            if (this.onAfterReply) {
                this.onAfterReply(item, replyText, success);
            }
            
            return success;
        } catch (error) {
            console.error("Fehler bei der Antwortgenerierung:", error);
            
            // Fehler-Event auslösen
            if (this.onError) {
                this.onError(item, error);
            }
            
            throw error;
        }
    }

    /**
     * Erstellt einen Prompt für die Antwortgenerierung
     * 
     * @param {Object} context - Der extrahierte Kontext
     * @param {Object} options - Zusätzliche Optionen
     * @returns {string} - Der erstellte Prompt
     */
    createReplyPrompt(context, options = {}) {
        // Formatiere den Kontext für das LLM
        const formattedContext = this.contextExtractor.formatContextForLLM(context);
        
        // Hole die Prompt-Vorlage für die Antwort
        let promptTemplate = this.config.promptTemplates.reply || options.promptTemplate;
        
        // Fallback auf Standard-Prompt, wenn keine Vorlage gefunden wurde
        if (!promptTemplate) {
            promptTemplate = `Generiere eine professionelle und hilfreiche Antwort auf die folgende E-Mail. 
Die Antwort sollte höflich, präzise und auf den Inhalt der E-Mail bezogen sein.
Berücksichtige dabei den gesamten E-Mail-Verlauf und beziehe dich auf relevante Informationen aus früheren Nachrichten.
Schreibe die Antwort direkt, ohne Einleitungen wie "Hier ist meine Antwort:" oder ähnliches.
Verwende einen professionellen, aber freundlichen Ton und achte auf eine korrekte Anrede und Grußformel.

E-Mail-Kontext:
{email_context}`;
        }
        
        // Ersetze Platzhalter im Prompt
        let prompt = promptTemplate.replace("{email_context}", formattedContext);
        
        // Ersetze weitere Platzhalter
        if (options.customInstructions) {
            prompt = prompt.replace("{custom_instructions}", options.customInstructions);
        }
        
        if (options.responseStyle) {
            prompt = prompt.replace("{response_style}", options.responseStyle);
        }
        
        return prompt;
    }

    /**
     * Fügt eine generierte Antwort in das Antwortformular ein
     * 
     * @param {Office.MessageItem} item - Die E-Mail, auf die geantwortet werden soll
     * @param {string} replyText - Der generierte Antworttext
     * @param {Object} options - Zusätzliche Optionen
     * @returns {Promise<boolean>} - true, wenn die Antwort erfolgreich eingefügt wurde
     */
    async insertReplyIntoEmail(item, replyText, options = {}) {
        return new Promise((resolve, reject) => {
            try {
                // Prüfen, ob wir im Lese- oder Compose-Modus sind
                if (Office.context.mailbox.item.displayReplyForm) {
                    // Lese-Modus: Antwortformular öffnen
                    Office.context.mailbox.item.displayReplyForm(replyText);
                    resolve(true);
                } else if (Office.context.mailbox.item.body && Office.context.mailbox.item.body.setSelectedDataAsync) {
                    // Compose-Modus: Text in die E-Mail einfügen
                    Office.context.mailbox.item.body.setSelectedDataAsync(
                        replyText,
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
                    reject(new Error("Konnte Antwort nicht einfügen: Weder im Lese- noch im Compose-Modus"));
                }
            } catch (error) {
                console.error("Fehler beim Einfügen der Antwort:", error);
                reject(error);
            }
        });
    }

    /**
     * Öffnet das Antwortformular und generiert eine Antwort
     * 
     * @param {Office.MessageItem} item - Die E-Mail, auf die geantwortet werden soll
     * @param {Object} options - Zusätzliche Optionen
     * @returns {Promise<boolean>} - true, wenn die Antwort erfolgreich generiert wurde
     */
    async openReplyFormAndGenerate(item, options = {}) {
        return new Promise((resolve, reject) => {
            try {
                // Prüfen, ob wir im Lese-Modus sind
                if (Office.context.mailbox.item.displayReplyForm) {
                    // Antwortformular öffnen
                    Office.context.mailbox.item.displayReplyForm("Generiere Antwort mit LLM...");
                    
                    // Warte einen Moment, bis das Antwortformular geöffnet ist
                    setTimeout(async () => {
                        try {
                            // Generiere die Antwort und füge sie ein
                            const success = await this.replyToEmail(item, options);
                            resolve(success);
                        } catch (error) {
                            reject(error);
                        }
                    }, 1000);
                } else {
                    // Wir sind bereits im Compose-Modus, generiere die Antwort direkt
                    this.replyToEmail(item, options)
                        .then(resolve)
                        .catch(reject);
                }
            } catch (error) {
                console.error("Fehler beim Öffnen des Antwortformulars:", error);
                reject(error);
            }
        });
    }

    /**
     * Generiert eine Antwort auf eine E-Mail mit einem bestimmten Stil
     * 
     * @param {Office.MessageItem} item - Die E-Mail, auf die geantwortet werden soll
     * @param {string} style - Der Stil der Antwort (z.B. "formal", "freundlich", "kurz")
     * @returns {Promise<boolean>} - true, wenn die Antwort erfolgreich generiert wurde
     */
    async replyWithStyle(item, style) {
        const styleInstructions = {
            formal: "Schreibe eine formelle und professionelle Antwort.",
            freundlich: "Schreibe eine freundliche und persönliche Antwort.",
            kurz: "Schreibe eine kurze und prägnante Antwort.",
            detailliert: "Schreibe eine detaillierte und ausführliche Antwort."
        };
        
        return this.replyToEmail(item, {
            responseStyle: styleInstructions[style] || styleInstructions.freundlich
        });
    }
}

// Exportiere die Klasse für die Verwendung in anderen Modulen
if (typeof module !== 'undefined' && module.exports) {
    module.exports = DirectReplyHandler;
}