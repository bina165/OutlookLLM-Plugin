/**
 * Email Context Extractor
 * 
 * Diese Klasse ist verantwortlich für die Extraktion und Aufbereitung des E-Mail-Kontexts
 * aus Outlook-Elementen (E-Mails, Termine, etc.) für die Verwendung mit dem LLM.
 */

class EmailContextExtractor {
    /**
     * Initialisiert den Email Context Extractor
     * 
     * @param {Object} options - Konfigurationsoptionen
     */
    constructor(options = {}) {
        this.options = {
            includeAttachments: options.includeAttachments || false,
            maxBodyLength: options.maxBodyLength || 10000,
            includeHeaders: options.includeHeaders || true,
            includeRecipients: options.includeRecipients || true,
            includeCc: options.includeCc || true,
            includeBcc: options.includeBcc || false,
            includeThread: options.includeThread || false,
            maxThreadDepth: options.maxThreadDepth || 3,
            ...options
        };
    }

    /**
     * Extrahiert den Kontext aus einem Outlook-Element
     * 
     * @param {Office.Item} item - Das Outlook-Element (E-Mail, Termin, etc.)
     * @returns {Promise<Object>} - Der extrahierte Kontext
     */
    async extractContext(item) {
        if (!item) {
            throw new Error("Kein Outlook-Element angegeben");
        }

        // Bestimme den Elementtyp und wende die entsprechende Extraktionsmethode an
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            return this.extractEmailContext(item);
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            return this.extractAppointmentContext(item);
        } else {
            throw new Error(`Nicht unterstützter Elementtyp: ${item.itemType}`);
        }
    }

    /**
     * Extrahiert den Kontext aus einer E-Mail
     * 
     * @param {Office.MessageItem} item - Die E-Mail
     * @returns {Promise<Object>} - Der extrahierte E-Mail-Kontext
     */
    async extractEmailContext(item) {
        const context = {
            type: "email",
            subject: item.subject || "",
            sender: {
                name: item.sender ? item.sender.displayName : "",
                email: item.sender ? item.sender.emailAddress : ""
            },
            recipients: [],
            cc: [],
            bcc: [],
            receivedTime: item.dateTimeCreated ? item.dateTimeCreated.toISOString() : "",
            importance: this.getImportanceLevel(item.importance),
            hasAttachments: false,
            attachments: [],
            body: "",
            thread: [],
            conversationId: item.conversationId || ""
        };

        // Empfänger extrahieren
        if (this.options.includeRecipients && item.to) {
            context.recipients = await this.extractRecipients(item.to);
        }

        // CC-Empfänger extrahieren
        if (this.options.includeCc && item.cc) {
            context.cc = await this.extractRecipients(item.cc);
        }

        // BCC-Empfänger extrahieren (nur im Compose-Modus verfügbar)
        if (this.options.includeBcc && item.bcc) {
            context.bcc = await this.extractRecipients(item.bcc);
        }

        // E-Mail-Text extrahieren
        try {
            context.body = await this.getItemBody(item);
        } catch (error) {
            console.error("Fehler beim Extrahieren des E-Mail-Textes:", error);
            context.body = "Fehler beim Laden des E-Mail-Textes";
        }

        // Anhänge extrahieren
        if (this.options.includeAttachments && item.attachments) {
            context.hasAttachments = item.attachments.length > 0;
            if (context.hasAttachments) {
                context.attachments = await this.extractAttachments(item.attachments);
            }
        }

        // E-Mail-Thread extrahieren (falls verfügbar und aktiviert)
        if (this.options.includeThread && context.conversationId && Office.context.mailbox.getCallbackTokenAsync) {
            try {
                context.thread = await this.extractThread(context.conversationId);
            } catch (error) {
                console.error("Fehler beim Extrahieren des E-Mail-Threads:", error);
            }
        }

        return context;
    }

    /**
     * Extrahiert den Kontext aus einem Termin
     * 
     * @param {Office.AppointmentItem} item - Der Termin
     * @returns {Promise<Object>} - Der extrahierte Termin-Kontext
     */
    async extractAppointmentContext(item) {
        const context = {
            type: "appointment",
            subject: item.subject || "",
            organizer: {
                name: item.organizer ? item.organizer.displayName : "",
                email: item.organizer ? item.organizer.emailAddress : ""
            },
            location: item.location || "",
            start: item.start ? item.start.toISOString() : "",
            end: item.end ? item.end.toISOString() : "",
            attendees: {
                required: [],
                optional: []
            },
            hasAttachments: false,
            attachments: [],
            body: ""
        };

        // Pflicht-Teilnehmer extrahieren
        if (item.requiredAttendees) {
            context.attendees.required = await this.extractRecipients(item.requiredAttendees);
        }

        // Optionale Teilnehmer extrahieren
        if (item.optionalAttendees) {
            context.attendees.optional = await this.extractRecipients(item.optionalAttendees);
        }

        // Termin-Text extrahieren
        try {
            context.body = await this.getItemBody(item);
        } catch (error) {
            console.error("Fehler beim Extrahieren des Termin-Textes:", error);
            context.body = "Fehler beim Laden des Termin-Textes";
        }

        // Anhänge extrahieren
        if (this.options.includeAttachments && item.attachments) {
            context.hasAttachments = item.attachments.length > 0;
            if (context.hasAttachments) {
                context.attachments = await this.extractAttachments(item.attachments);
            }
        }

        return context;
    }

    /**
     * Extrahiert Empfänger aus einem Outlook-Empfängerobjekt
     * 
     * @param {Office.Recipients} recipients - Die Empfänger
     * @returns {Promise<Array<Object>>} - Die extrahierten Empfänger
     */
    async extractRecipients(recipients) {
        return new Promise((resolve) => {
            try {
                const extractedRecipients = [];
                
                if (!recipients || !recipients.getAsync) {
                    resolve(extractedRecipients);
                    return;
                }
                
                recipients.getAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const recipientsArray = result.value || [];
                        
                        recipientsArray.forEach((recipient) => {
                            extractedRecipients.push({
                                name: recipient.displayName || "",
                                email: recipient.emailAddress || ""
                            });
                        });
                    }
                    
                    resolve(extractedRecipients);
                });
            } catch (error) {
                console.error("Fehler beim Extrahieren der Empfänger:", error);
                resolve([]);
            }
        });
    }

    /**
     * Extrahiert Anhänge aus einem Outlook-Anhangsobjekt
     * 
     * @param {Office.AttachmentDetails[]} attachments - Die Anhänge
     * @returns {Promise<Array<Object>>} - Die extrahierten Anhangsinformationen
     */
    async extractAttachments(attachments) {
        const extractedAttachments = [];
        
        if (!attachments || !Array.isArray(attachments)) {
            return extractedAttachments;
        }
        
        attachments.forEach((attachment) => {
            extractedAttachments.push({
                id: attachment.id || "",
                name: attachment.name || "",
                contentType: attachment.contentType || "",
                size: attachment.size || 0,
                isInline: attachment.isInline || false
            });
        });
        
        return extractedAttachments;
    }

    /**
     * Extrahiert den E-Mail-Thread für eine bestimmte Konversations-ID
     * 
     * @param {string} conversationId - Die Konversations-ID
     * @returns {Promise<Array<Object>>} - Der extrahierte E-Mail-Thread
     */
    async extractThread(conversationId) {
        return new Promise((resolve, reject) => {
            try {
                // Diese Funktion würde in einer realen Implementierung die EWS-API verwenden,
                // um den E-Mail-Thread zu extrahieren. Da dies jedoch komplex ist und
                // zusätzliche Berechtigungen erfordert, liefern wir hier eine Dummy-Implementierung.
                
                // In einer vollständigen Implementierung würde hier die Office.js API verwendet werden,
                // um auf die EWS-API zuzugreifen und den Thread zu extrahieren.
                
                // Beispiel für eine Dummy-Implementierung:
                setTimeout(() => {
                    resolve([]);
                }, 100);
                
                // Echte Implementierung würde etwa so aussehen:
                /*
                Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const token = result.value;
                        
                        // Verwende den Token, um die REST-API aufzurufen und den Thread zu extrahieren
                        // ...
                        
                        resolve(threadMessages);
                    } else {
                        reject(new Error("Fehler beim Abrufen des Callback-Tokens"));
                    }
                });
                */
            } catch (error) {
                console.error("Fehler beim Extrahieren des E-Mail-Threads:", error);
                reject(error);
            }
        });
    }

    /**
     * Holt den Text eines Outlook-Elements
     * 
     * @param {Office.Item} item - Das Outlook-Element
     * @returns {Promise<string>} - Der Text des Elements
     */
    async getItemBody(item) {
        return new Promise((resolve, reject) => {
            try {
                if (!item.body) {
                    resolve("");
                    return;
                }
                
                item.body.getAsync(Office.CoercionType.Text, { maxBytes: this.options.maxBodyLength }, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(result.value || "");
                    } else {
                        reject(new Error("Fehler beim Abrufen des Textes"));
                    }
                });
            } catch (error) {
                console.error("Fehler beim Abrufen des Textes:", error);
                reject(error);
            }
        });
    }

    /**
     * Konvertiert den Wichtigkeitswert in einen lesbaren String
     * 
     * @param {Office.MailboxEnums.ImportanceType} importance - Der Wichtigkeitswert
     * @returns {string} - Die lesbare Wichtigkeit
     */
    getImportanceLevel(importance) {
        if (!importance) {
            return "normal";
        }
        
        switch (importance) {
            case Office.MailboxEnums.ImportanceType.Low:
                return "niedrig";
            case Office.MailboxEnums.ImportanceType.High:
                return "hoch";
            case Office.MailboxEnums.ImportanceType.Normal:
            default:
                return "normal";
        }
    }

    /**
     * Formatiert den extrahierten Kontext für die Verwendung mit dem LLM
     * 
     * @param {Object} context - Der extrahierte Kontext
     * @returns {string} - Der formatierte Kontext
     */
    formatContextForLLM(context) {
        if (!context) {
            return "";
        }
        
        let formattedContext = "";
        
        if (context.type === "email") {
            formattedContext += `Betreff: ${context.subject}\n`;
            formattedContext += `Von: ${context.sender.name} <${context.sender.email}>\n`;
            
            if (context.recipients && context.recipients.length > 0) {
                formattedContext += "An: ";
                formattedContext += context.recipients.map(r => `${r.name} <${r.email}>`).join(", ");
                formattedContext += "\n";
            }
            
            if (context.cc && context.cc.length > 0) {
                formattedContext += "CC: ";
                formattedContext += context.cc.map(r => `${r.name} <${r.email}>`).join(", ");
                formattedContext += "\n";
            }
            
            formattedContext += `Datum: ${new Date(context.receivedTime).toLocaleString()}\n`;
            formattedContext += `Wichtigkeit: ${context.importance}\n`;
            
            if (context.hasAttachments) {
                formattedContext += "Anhänge: ";
                formattedContext += context.attachments.map(a => a.name).join(", ");
                formattedContext += "\n";
            }
            
            formattedContext += "\n";
            formattedContext += context.body;
            
            // Thread-Nachrichten hinzufügen, falls vorhanden
            if (context.thread && context.thread.length > 0) {
                formattedContext += "\n\n--- Vorherige Nachrichten ---\n\n";
                
                context.thread.forEach((message, index) => {
                    if (index < this.options.maxThreadDepth) {
                        formattedContext += `Von: ${message.sender.name} <${message.sender.email}>\n`;
                        formattedContext += `Datum: ${new Date(message.receivedTime).toLocaleString()}\n`;
                        formattedContext += `Betreff: ${message.subject}\n\n`;
                        formattedContext += message.body;
                        formattedContext += "\n\n---\n\n";
                    }
                });
            }
        } else if (context.type === "appointment") {
            formattedContext += `Betreff: ${context.subject}\n`;
            formattedContext += `Organisator: ${context.organizer.name} <${context.organizer.email}>\n`;
            formattedContext += `Ort: ${context.location}\n`;
            formattedContext += `Start: ${new Date(context.start).toLocaleString()}\n`;
            formattedContext += `Ende: ${new Date(context.end).toLocaleString()}\n`;
            
            if (context.attendees.required.length > 0) {
                formattedContext += "Pflicht-Teilnehmer: ";
                formattedContext += context.attendees.required.map(a => `${a.name} <${a.email}>`).join(", ");
                formattedContext += "\n";
            }
            
            if (context.attendees.optional.length > 0) {
                formattedContext += "Optionale Teilnehmer: ";
                formattedContext += context.attendees.optional.map(a => `${a.name} <${a.email}>`).join(", ");
                formattedContext += "\n";
            }
            
            if (context.hasAttachments) {
                formattedContext += "Anhänge: ";
                formattedContext += context.attachments.map(a => a.name).join(", ");
                formattedContext += "\n";
            }
            
            formattedContext += "\n";
            formattedContext += context.body;
        }
        
        return formattedContext;
    }
}

// Exportiere die Klasse für die Verwendung in anderen Modulen
if (typeof module !== 'undefined' && module.exports) {
    module.exports = EmailContextExtractor;
}