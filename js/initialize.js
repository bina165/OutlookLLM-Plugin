/**
 * Initialisierungsdatei für das LLM Assistant Outlook-Plugin
 * 
 * Diese Datei initialisiert alle Komponenten des Plugins und stellt die Verbindung
 * zwischen der Benutzeroberfläche, dem Triton-Server und Outlook her.
 */

// Globale Variablen für die Komponenten
let tritonConnector = null;
let contextExtractor = null;
let llmActions = null;
let directReplyHandler = null;
let currentItem = null;
let config = null;

// Initialisierung beim Laden des Office-Add-ins
Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        console.log("Outlook LLM Assistant wird initialisiert...");
        
        // Konfiguration laden
        loadConfig()
            .then(initializeComponents)
            .then(setupEventListeners)
            .then(() => {
                console.log("Outlook LLM Assistant erfolgreich initialisiert");
                updateUIStatus("Bereit");
                loadCurrentItem();
            })
            .catch(error => {
                console.error("Fehler bei der Initialisierung:", error);
                updateUIStatus("Fehler bei der Initialisierung: " + error.message);
            });
    }
});

/**
 * Lädt die Konfiguration aus der Konfigurationsdatei oder dem localStorage
 */
async function loadConfig() {
    try {
        // Versuche zuerst, die Konfiguration aus dem localStorage zu laden
        const savedConfig = localStorage.getItem("llmAssistantConfig");
        if (savedConfig) {
            config = JSON.parse(savedConfig);
            console.log("Konfiguration aus localStorage geladen");
            return config;
        }
        
        // Wenn keine Konfiguration im localStorage gefunden wurde, lade die Standardkonfiguration
        const response = await fetch("triton-config.json");
        if (!response.ok) {
            throw new Error(`HTTP-Fehler ${response.status}`);
        }
        
        config = await response.json();
        console.log("Standardkonfiguration geladen");
        
        // Speichere die Konfiguration im localStorage
        localStorage.setItem("llmAssistantConfig", JSON.stringify(config));
        
        return config;
    } catch (error) {
        console.error("Fehler beim Laden der Konfiguration:", error);
        
        // Fallback auf Standardkonfiguration
        config = {
            serverConfig: {
                url: "http://localhost:8000",
                apiVersion: "v2",
                timeout: 30000,
                maxRetries: 3,
                retryDelay: 1000,
                debug: false
            },
            modelConfig: {
                name: "llm_model",
                defaultParameters: {
                    max_tokens: 1024,
                    temperature: 0.7,
                    top_p: 0.9,
                    stop_sequences: ["\n###", "###", "</answer>"],
                    return_full_text: false
                }
            },
            promptTemplates: {
                analyze: "Analysiere diese E-Mail und gib mir die wichtigsten Punkte und erforderlichen Aktionen:\n\n{email_context}",
                summarize: "Fasse diese E-Mail kurz und prägnant zusammen:\n\n{email_context}",
                reply: "Generiere eine professionelle Antwort auf diese E-Mail:\n\n{email_context}",
                translate: "Übersetze diese E-Mail ins {target_language}:\n\n{email_context}",
                calendar: "Extrahiere Informationen für einen Kalendereintrag aus dieser E-Mail (Datum, Uhrzeit, Teilnehmer, Ort, Thema) und formatiere sie als JSON:\n\n{email_context}",
                custom: "{custom_prompt}\n\n{email_context}"
            }
        };
        
        return config;
    }
}

/**
 * Initialisiert alle Komponenten des Plugins
 */
function initializeComponents() {
    // Initialisiere den Triton-Connector
    tritonConnector = new TritonConnector(
        config.serverConfig.url,
        config.modelConfig.name,
        {
            apiKey: config.securityConfig?.apiKey,
            timeout: config.serverConfig.timeout,
            maxRetries: config.serverConfig.maxRetries,
            retryDelay: config.serverConfig.retryDelay,
            debug: config.serverConfig.debug
        }
    );
    
    // Initialisiere den E-Mail-Kontext-Extraktor
    contextExtractor = new EmailContextExtractor({
        includeAttachments: true,
        maxBodyLength: 10000,
        includeHeaders: true,
        includeRecipients: true,
        includeCc: true,
        includeBcc: false,
        includeThread: false,
        maxThreadDepth: 3
    });
    
    // Initialisiere die LLM-Aktionen
    llmActions = new LLMActions(tritonConnector, contextExtractor, {
        promptTemplates: config.promptTemplates,
        defaultParameters: config.modelConfig.defaultParameters
    });
    
    // Initialisiere den Direct Reply Handler
    directReplyHandler = new DirectReplyHandler(tritonConnector, contextExtractor, {
        promptTemplates: config.promptTemplates,
        defaultParameters: config.modelConfig.defaultParameters,
        includeThread: true
    });
    
    // Event-Handler für LLM-Aktionen
    llmActions.onBeforeAction = (actionType) => {
        updateUIStatus(`Führe ${actionType}-Aktion aus...`);
        showLoadingIndicator(true);
    };
    
    llmActions.onAfterAction = (actionType) => {
        updateUIStatus(`${actionType}-Aktion abgeschlossen`);
        showLoadingIndicator(false);
    };
    
    llmActions.onError = (actionType, error) => {
        updateUIStatus(`Fehler bei ${actionType}-Aktion: ${error.message}`);
        showLoadingIndicator(false);
    };
    
    return Promise.resolve();
}

/**
 * Richtet Event-Listener für UI-Elemente ein
 */
function setupEventListeners() {
    // Analyse-Button
    document.getElementById("analyze-btn").addEventListener("click", () => {
        if (!currentItem) {
            showResponse("Bitte warten Sie, bis die E-Mail geladen wurde.");
            return;
        }
        
        llmActions.analyzeEmail(currentItem)
            .then(result => {
                showResponse(result.text);
                showResponseActions(true);
            })
            .catch(error => {
                showResponse("Fehler bei der Analyse: " + error.message);
            });
    });
    
    // Zusammenfassungs-Button
    document.getElementById("summarize-btn").addEventListener("click", () => {
        if (!currentItem) {
            showResponse("Bitte warten Sie, bis die E-Mail geladen wurde.");
            return;
        }
        
        llmActions.summarizeEmail(currentItem)
            .then(result => {
                showResponse(result.text);
                showResponseActions(true);
            })
            .catch(error => {
                showResponse("Fehler bei der Zusammenfassung: " + error.message);
            });
    });
    
    // Antwort-Button
    document.getElementById("reply-btn").addEventListener("click", () => {
        if (!currentItem) {
            showResponse("Bitte warten Sie, bis die E-Mail geladen wurde.");
            return;
        }
        
        llmActions.generateReply(currentItem)
            .then(result => {
                showResponse(result.text);
                showResponseActions(true);
            })
            .catch(error => {
                showResponse("Fehler bei der Antwortgenerierung: " + error.message);
            });
    });
    
    // Direkt-Antwort-Button
    document.getElementById("direct-reply-btn").addEventListener("click", () => {
        if (!currentItem) {
            showResponse("Bitte warten Sie, bis die E-Mail geladen wurde.");
            return;
        }
        
        updateUIStatus("Generiere Antwort und öffne Antwortformular...");
        showLoadingIndicator(true);
        
        directReplyHandler.openReplyFormAndGenerate(currentItem)
            .then(success => {
                if (success) {
                    updateUIStatus("Antwort wurde im Antwortformular eingefügt.");
                } else {
                    updateUIStatus("Antwort konnte nicht eingefügt werden.");
                }
                showLoadingIndicator(false);
            })
            .catch(error => {
                console.error("Fehler bei der direkten Antwortgenerierung:", error);
                updateUIStatus("Fehler bei der direkten Antwortgenerierung: " + error.message);
                showLoadingIndicator(false);
                showResponse("Fehler bei der direkten Antwortgenerierung: " + error.message);
            });
    });
    
    // Antwort-Stil-Button
    document.getElementById("reply-style-btn").addEventListener("click", () => {
        if (!currentItem) {
            showResponse("Bitte warten Sie, bis die E-Mail geladen wurde.");
            return;
        }
        
        // Dialog für Antwortstil anzeigen
        const styles = [
            { id: "formal", name: "Formell" },
            { id: "freundlich", name: "Freundlich" },
            { id: "kurz", name: "Kurz und prägnant" },
            { id: "detailliert", name: "Detailliert" }
        ];
        
        // Erstelle einen einfachen Dialog
        const styleDialog = document.createElement("div");
        styleDialog.className = "style-dialog";
        styleDialog.style.position = "fixed";
        styleDialog.style.top = "50%";
        styleDialog.style.left = "50%";
        styleDialog.style.transform = "translate(-50%, -50%)";
        styleDialog.style.backgroundColor = "white";
        styleDialog.style.padding = "20px";
        styleDialog.style.boxShadow = "0 0 10px rgba(0,0,0,0.5)";
        styleDialog.style.zIndex = "1000";
        styleDialog.style.borderRadius = "4px";
        
        styleDialog.innerHTML = `
            <h3>Wählen Sie einen Antwortstil</h3>
            <div class="style-options">
                ${styles.map(style => `
                    <button class="ms-Button ms-Button--primary style-option" data-style="${style.id}">
                        <span class="ms-Button-label">${style.name}</span>
                    </button>
                `).join('')}
            </div>
            <button class="ms-Button ms-Button--default cancel-btn">
                <span class="ms-Button-label">Abbrechen</span>
            </button>
        `;
        
        document.body.appendChild(styleDialog);
        
        // Event-Listener für die Stil-Buttons
        const styleButtons = styleDialog.querySelectorAll(".style-option");
        styleButtons.forEach(button => {
            button.addEventListener("click", () => {
                const style = button.getAttribute("data-style");
                document.body.removeChild(styleDialog);
                
                updateUIStatus(`Generiere ${button.textContent.trim()}-Antwort...`);
                showLoadingIndicator(true);
                
                directReplyHandler.replyWithStyle(currentItem, style)
                    .then(success => {
                        if (success) {
                            updateUIStatus("Antwort wurde im Antwortformular eingefügt.");
                        } else {
                            updateUIStatus("Antwort konnte nicht eingefügt werden.");
                        }
                        showLoadingIndicator(false);
                    })
                    .catch(error => {
                        console.error("Fehler bei der Antwortgenerierung mit Stil:", error);
                        updateUIStatus("Fehler bei der Antwortgenerierung: " + error.message);
                        showLoadingIndicator(false);
                        showResponse("Fehler bei der Antwortgenerierung: " + error.message);
                    });
            });
        });
        
        // Event-Listener für den Abbrechen-Button
        const cancelButton = styleDialog.querySelector(".cancel-btn");
        cancelButton.addEventListener("click", () => {
            document.body.removeChild(styleDialog);
        });
    });
    
    // Übersetzungs-Button
    document.getElementById("translate-btn").addEventListener("click", () => {
        if (!currentItem) {
            showResponse("Bitte warten Sie, bis die E-Mail geladen wurde.");
            return;
        }
        
        // Dialog für Zielsprache anzeigen
        const targetLanguage = prompt("In welche Sprache soll die E-Mail übersetzt werden?", "Englisch");
        if (!targetLanguage) return;
        
        llmActions.translateEmail(currentItem, targetLanguage)
            .then(result => {
                showResponse(result.text);
                showResponseActions(true);
            })
            .catch(error => {
                showResponse("Fehler bei der Übersetzung: " + error.message);
            });
    });
    
    // Kalender-Button
    document.getElementById("calendar-btn").addEventListener("click", () => {
        if (!currentItem) {
            showResponse("Bitte warten Sie, bis die E-Mail geladen wurde.");
            return;
        }
        
        llmActions.extractCalendarData(currentItem)
            .then(result => {
                if (result.eventData && Object.keys(result.eventData).length > 0) {
                    // Zeige die extrahierten Kalenderdaten an
                    let eventHtml = "<h3>Extrahierte Kalenderdaten:</h3>";
                    eventHtml += "<div class='calendar-data'>";
                    
                    if (result.eventData.subject) {
                        eventHtml += `<p><strong>Betreff:</strong> ${escapeHtml(result.eventData.subject)}</p>`;
                    }
                    
                    if (result.eventData.start) {
                        const startDate = new Date(result.eventData.start);
                        eventHtml += `<p><strong>Start:</strong> ${formatDate(startDate)}</p>`;
                    }
                    
                    if (result.eventData.end) {
                        const endDate = new Date(result.eventData.end);
                        eventHtml += `<p><strong>Ende:</strong> ${formatDate(endDate)}</p>`;
                    }
                    
                    if (result.eventData.location) {
                        eventHtml += `<p><strong>Ort:</strong> ${escapeHtml(result.eventData.location)}</p>`;
                    }
                    
                    if (result.eventData.attendees && result.eventData.attendees.length > 0) {
                        eventHtml += "<p><strong>Teilnehmer:</strong></p>";
                        eventHtml += "<ul>";
                        result.eventData.attendees.forEach(attendee => {
                            eventHtml += `<li>${escapeHtml(attendee)}</li>`;
                        });
                        eventHtml += "</ul>";
                    }
                    
                    if (result.eventData.description) {
                        eventHtml += `<p><strong>Beschreibung:</strong></p>`;
                        eventHtml += `<div class="description">${escapeHtml(result.eventData.description)}</div>`;
                    }
                    
                    eventHtml += "</div>";
                    
                    // Füge einen Button zum Erstellen des Kalendereintrags hinzu
                    eventHtml += `
                        <button id="create-appointment-btn" class="ms-Button ms-Button--primary">
                            <span class="ms-Button-label">Kalendereintrag erstellen</span>
                        </button>
                    `;
                    
                    showResponse(eventHtml);
                    
                    // Event-Listener für den Kalendereintrag-Button
                    document.getElementById("create-appointment-btn").addEventListener("click", () => {
                        llmActions.createOutlookAppointment(result.eventData)
                            .then(() => {
                                showTemporaryMessage("Kalendereintrag wurde erstellt.");
                            })
                            .catch(error => {
                                showResponse("Fehler beim Erstellen des Kalendereintrags: " + error.message);
                            });
                    });
                } else {
                    showResponse(result.text);
                }
                
                showResponseActions(true);
            })
            .catch(error => {
                showResponse("Fehler bei der Kalenderdatenextraktion: " + error.message);
            });
    });
    
    // Benutzerdefinierte Aktion-Button
    document.getElementById("custom-btn").addEventListener("click", () => {
        if (!currentItem) {
            showResponse("Bitte warten Sie, bis die E-Mail geladen wurde.");
            return;
        }
        
        // Dialog für benutzerdefinierten Prompt anzeigen
        const customPrompt = prompt("Geben Sie einen benutzerdefinierten Prompt ein:", "");
        if (!customPrompt) return;
        
        llmActions.executeCustomAction(currentItem, customPrompt)
            .then(result => {
                showResponse(result.text);
                showResponseActions(true);
            })
            .catch(error => {
                showResponse("Fehler bei der benutzerdefinierten Aktion: " + error.message);
            });
    });
    
    // Einfügen-Button
    document.getElementById("insert-btn").addEventListener("click", () => {
        const responseText = document.getElementById("response-text").textContent;
        if (!responseText) {
            showTemporaryMessage("Nichts zum Einfügen vorhanden.");
            return;
        }
        
        llmActions.insertTextIntoEmail(responseText)
            .then(success => {
                if (success) {
                    showTemporaryMessage("Text wurde in die E-Mail eingefügt.");
                } else {
                    showTemporaryMessage("Text konnte nicht eingefügt werden.");
                }
            })
            .catch(error => {
                showTemporaryMessage("Fehler beim Einfügen: " + error.message);
            });
    });
    
    // Kopieren-Button
    document.getElementById("copy-btn").addEventListener("click", () => {
        const responseText = document.getElementById("response-text").textContent;
        if (!responseText) {
            showTemporaryMessage("Nichts zum Kopieren vorhanden.");
            return;
        }
        
        navigator.clipboard.writeText(responseText)
            .then(() => {
                showTemporaryMessage("Text wurde in die Zwischenablage kopiert.");
            })
            .catch(error => {
                showTemporaryMessage("Fehler beim Kopieren: " + error.message);
            });
    });
    
    // Konfigurationsbutton
    document.getElementById("config-btn").addEventListener("click", () => {
        // Hier könnte ein Konfigurationsdialog angezeigt werden
        alert("Konfiguration wird in einer zukünftigen Version verfügbar sein.");
    });
    
    // Event-Listener für Änderungen am aktuellen Element
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadCurrentItem);
}

/**
 * Lädt das aktuelle Outlook-Element
 */
function loadCurrentItem() {
    currentItem = Office.context.mailbox.item;
    
    if (currentItem) {
        updateEmailInfo();
        updateUIStatus("E-Mail geladen");
    } else {
        updateUIStatus("Keine E-Mail ausgewählt");
    }
}

/**
 * Aktualisiert die Anzeige der E-Mail-Informationen
 */
function updateEmailInfo() {
    if (!currentItem) return;
    
    const emailInfoElement = document.getElementById("email-info");
    if (!emailInfoElement) return;
    
    // Betreff anzeigen
    const subject = currentItem.subject || "Kein Betreff";
    
    // Absender anzeigen (falls verfügbar)
    let sender = "";
    if (currentItem.from) {
        sender = currentItem.from.displayName || currentItem.from.emailAddress || "";
    }
    
    // Empfänger anzeigen (falls verfügbar)
    let recipients = "";
    if (currentItem.to && currentItem.to.length > 0) {
        recipients = currentItem.to.map(recipient => recipient.displayName || recipient.emailAddress).join(", ");
    }
    
    // Datum anzeigen (falls verfügbar)
    let date = "";
    if (currentItem.dateTimeCreated) {
        date = formatDate(currentItem.dateTimeCreated);
    }
    
    // Informationen in HTML formatieren
    let infoHtml = `
        <div class="email-header">
            <div class="email-subject">${escapeHtml(subject)}</div>
            ${sender ? `<div class="email-sender">Von: ${escapeHtml(sender)}</div>` : ""}
            ${recipients ? `<div class="email-recipients">An: ${escapeHtml(recipients)}</div>` : ""}
            ${date ? `<div class="email-date">Datum: ${date}</div>` : ""}
        </div>
    `;
    
    emailInfoElement.innerHTML = infoHtml;
}

/**
 * Aktualisiert den Status in der UI
 */
function updateUIStatus(status) {
    const statusElement = document.getElementById("status");
    if (statusElement) {
        statusElement.textContent = status;
    }
    
    // Aktualisiere auch den Status im Titel
    const titleElement = document.getElementById("app-title");
    if (titleElement) {
        const baseTitle = "LLM Assistant";
        if (status && status !== "Bereit") {
            titleElement.textContent = `${baseTitle} - ${status}`;
        } else {
            titleElement.textContent = baseTitle;
        }
    }
    
    console.log("Status:", status);
}

/**
 * Zeigt oder versteckt den Ladeindikator
 */
function showLoadingIndicator(show) {
    const loadingElement = document.getElementById("loading");
    if (loadingElement) {
        loadingElement.style.display = show ? "block" : "none";
    }
}

/**
 * Zeigt die LLM-Antwort im Antwortbereich an
 */
function showResponse(response) {
    const responseElement = document.getElementById("response-text");
    if (responseElement) {
        responseElement.innerHTML = response;
    }
}

/**
 * Zeigt oder versteckt die Aktionsbuttons für die Antwort
 */
function showResponseActions(show) {
    const actionsElement = document.getElementById("response-actions");
    if (actionsElement) {
        actionsElement.style.display = show ? "flex" : "none";
    }
}

/**
 * Zeigt eine temporäre Nachricht an
 */
function showTemporaryMessage(message) {
    const messageElement = document.createElement("div");
    messageElement.className = "temporary-message";
    messageElement.textContent = message;
    
    // Stil für die temporäre Nachricht
    messageElement.style.position = "fixed";
    messageElement.style.bottom = "20px";
    messageElement.style.left = "50%";
    messageElement.style.transform = "translateX(-50%)";
    messageElement.style.backgroundColor = "rgba(0, 0, 0, 0.7)";
    messageElement.style.color = "white";
    messageElement.style.padding = "10px 20px";
    messageElement.style.borderRadius = "4px";
    messageElement.style.zIndex = "1000";
    
    document.body.appendChild(messageElement);
    
    // Nachricht nach 3 Sekunden ausblenden
    setTimeout(() => {
        messageElement.style.opacity = "0";
        messageElement.style.transition = "opacity 0.5s";
        
        // Nachricht nach dem Ausblenden entfernen
        setTimeout(() => {
            document.body.removeChild(messageElement);
        }, 500);
    }, 3000);
}

/**
 * Formatiert ein Datum für die Anzeige
 */
function formatDate(dateObj) {
    if (!dateObj) return "";
    
    try {
        const date = new Date(dateObj);
        return date.toLocaleString("de-DE", {
            day: "2-digit",
            month: "2-digit",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit"
        });
    } catch (error) {
        return "";
    }
}

/**
 * Escaped HTML-Sonderzeichen
 */
function escapeHtml(text) {
    if (!text) return "";
    
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}