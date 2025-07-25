/**
 * Triton LLM Connector
 * 
 * Dieses Modul stellt die Verbindung zwischen dem Outlook-Plugin und dem NVIDIA Triton Inference Server her,
 * der das LLM hostet. Es bietet Funktionen für die sichere Kommunikation, Authentifizierung und Fehlerbehandlung.
 */

class TritonConnector {
    /**
     * Initialisiert den Triton-Connector
     * 
     * @param {string} serverUrl - URL des Triton-Servers (z.B. http://localhost:8000)
     * @param {string} modelName - Name des LLM-Modells auf dem Triton-Server
     * @param {Object} options - Zusätzliche Optionen (apiKey, timeout, etc.)
     */
    constructor(serverUrl, modelName, options = {}) {
        this.serverUrl = serverUrl;
        this.modelName = modelName;
        this.apiKey = options.apiKey || null;
        this.timeout = options.timeout || 30000; // 30 Sekunden Timeout
        this.maxRetries = options.maxRetries || 3;
        this.retryDelay = options.retryDelay || 1000; // 1 Sekunde zwischen Wiederholungsversuchen
        this.debug = options.debug || false;
    }

    /**
     * Sendet eine Anfrage an das LLM
     * 
     * @param {string} prompt - Der Prompt für das LLM
     * @param {Object} parameters - Zusätzliche Parameter für die Inferenz (temperature, max_tokens, etc.)
     * @returns {Promise<Object>} - Die Antwort des LLM
     */
    async generateText(prompt, parameters = {}) {
        const requestData = {
            prompt: prompt,
            max_tokens: parameters.max_tokens || 1024,
            temperature: parameters.temperature || 0.7,
            top_p: parameters.top_p || 1.0,
            stop_sequences: parameters.stop_sequences || [],
            return_full_text: parameters.return_full_text || false
        };

        return this._sendRequest('/v2/models/' + this.modelName + '/generate', requestData);
    }

    /**
     * Sendet eine Batch-Anfrage an das LLM für mehrere Prompts
     * 
     * @param {Array<string>} prompts - Liste von Prompts
     * @param {Object} parameters - Zusätzliche Parameter für die Inferenz
     * @returns {Promise<Array<Object>>} - Die Antworten des LLM
     */
    async generateBatch(prompts, parameters = {}) {
        const requestData = {
            prompts: prompts,
            max_tokens: parameters.max_tokens || 1024,
            temperature: parameters.temperature || 0.7,
            top_p: parameters.top_p || 1.0,
            stop_sequences: parameters.stop_sequences || [],
            return_full_text: parameters.return_full_text || false
        };

        return this._sendRequest('/v2/models/' + this.modelName + '/generate_batch', requestData);
    }

    /**
     * Ruft Informationen über das Modell ab
     * 
     * @returns {Promise<Object>} - Informationen über das Modell
     */
    async getModelInfo() {
        return this._sendRequest('/v2/models/' + this.modelName, {}, 'GET');
    }

    /**
     * Ruft eine Liste aller verfügbaren Modelle ab
     * 
     * @returns {Promise<Array<Object>>} - Liste der verfügbaren Modelle
     */
    async listModels() {
        return this._sendRequest('/v2/models', {}, 'GET');
    }

    /**
     * Prüft die Verbindung zum Triton-Server
     * 
     * @returns {Promise<boolean>} - true, wenn die Verbindung erfolgreich ist
     */
    async testConnection() {
        try {
            const response = await this._sendRequest('/v2/health/ready', {}, 'GET');
            return response && response.status === 'READY';
        } catch (error) {
            this._logDebug('Verbindungstest fehlgeschlagen:', error);
            return false;
        }
    }

    /**
     * Sendet eine HTTP-Anfrage an den Triton-Server
     * 
     * @param {string} endpoint - Der API-Endpunkt
     * @param {Object} data - Die Daten für die Anfrage
     * @param {string} method - Die HTTP-Methode (GET, POST, etc.)
     * @returns {Promise<Object>} - Die Antwort des Servers
     * @private
     */
    async _sendRequest(endpoint, data = {}, method = 'POST') {
        const url = this.serverUrl + endpoint;
        
        const headers = {
            'Content-Type': 'application/json'
        };
        
        // Füge API-Key hinzu, falls vorhanden
        if (this.apiKey) {
            headers['Authorization'] = `Bearer ${this.apiKey}`;
        }
        
        const options = {
            method: method,
            headers: headers,
            timeout: this.timeout
        };
        
        // Füge Body hinzu, falls es sich um eine POST-Anfrage handelt
        if (method === 'POST') {
            options.body = JSON.stringify(data);
        }
        
        this._logDebug(`Sende ${method}-Anfrage an ${url}`, data);
        
        // Implementiere Wiederholungslogik
        let lastError = null;
        for (let attempt = 1; attempt <= this.maxRetries; attempt++) {
            try {
                const response = await this._fetchWithTimeout(url, options);
                
                if (!response.ok) {
                    const errorText = await response.text();
                    throw new Error(`HTTP-Fehler ${response.status}: ${errorText}`);
                }
                
                const jsonResponse = await response.json();
                this._logDebug('Antwort erhalten:', jsonResponse);
                return jsonResponse;
            } catch (error) {
                lastError = error;
                this._logDebug(`Versuch ${attempt}/${this.maxRetries} fehlgeschlagen:`, error);
                
                // Wenn dies nicht der letzte Versuch ist, warte vor dem nächsten Versuch
                if (attempt < this.maxRetries) {
                    await new Promise(resolve => setTimeout(resolve, this.retryDelay * attempt));
                }
            }
        }
        
        // Wenn wir hier ankommen, sind alle Versuche fehlgeschlagen
        throw new Error(`Alle ${this.maxRetries} Anfrageversuche fehlgeschlagen. Letzter Fehler: ${lastError.message}`);
    }

    /**
     * Führt einen Fetch mit Timeout durch
     * 
     * @param {string} url - Die URL
     * @param {Object} options - Die Fetch-Optionen
     * @returns {Promise<Response>} - Die Fetch-Antwort
     * @private
     */
    async _fetchWithTimeout(url, options) {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), this.timeout);
        
        try {
            const response = await fetch(url, {
                ...options,
                signal: controller.signal
            });
            
            clearTimeout(timeoutId);
            return response;
        } catch (error) {
            clearTimeout(timeoutId);
            if (error.name === 'AbortError') {
                throw new Error(`Anfrage-Timeout nach ${this.timeout}ms`);
            }
            throw error;
        }
    }

    /**
     * Loggt Debug-Informationen, wenn der Debug-Modus aktiviert ist
     * 
     * @param {string} message - Die Nachricht
     * @param {*} data - Zusätzliche Daten
     * @private
     */
    _logDebug(message, data) {
        if (this.debug) {
            console.log(`[TritonConnector] ${message}`, data);
        }
    }
}

// Exportiere die Klasse für die Verwendung in anderen Modulen
if (typeof module !== 'undefined' && module.exports) {
    module.exports = TritonConnector;
}