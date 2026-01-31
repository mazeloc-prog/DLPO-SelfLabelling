/**
 * MazeShield DRM Client
 * Chiffrement/Déchiffrement côté client pour Office Add-in
 */

const DRM_API_BASE = "https://fn-mazeshield-drm.azurewebsites.net/api";
const TENANT_ID = "mazeloc";

// ============================================================
// CRYPTO FUNCTIONS (AES-256-GCM)
// ============================================================

/**
 * Génère une clé AES-256 (DEK) aléatoire
 * @returns {Promise<CryptoKey>}
 */
async function generateDEK() {
    return await window.crypto.subtle.generateKey(
        { name: "AES-GCM", length: 256 },
        true, // extractable
        ["encrypt", "decrypt"]
    );
}

/**
 * Exporte une CryptoKey en bytes (pour envoyer à l'API)
 * @param {CryptoKey} key 
 * @returns {Promise<Uint8Array>}
 */
async function exportKey(key) {
    const raw = await window.crypto.subtle.exportKey("raw", key);
    return new Uint8Array(raw);
}

/**
 * Importe des bytes en CryptoKey
 * @param {Uint8Array} keyBytes 
 * @returns {Promise<CryptoKey>}
 */
async function importKey(keyBytes) {
    return await window.crypto.subtle.importKey(
        "raw",
        keyBytes,
        { name: "AES-GCM", length: 256 },
        false,
        ["encrypt", "decrypt"]
    );
}

/**
 * Chiffre du contenu avec AES-256-GCM
 * @param {string} plaintext 
 * @param {CryptoKey} key 
 * @returns {Promise<{ciphertext: Uint8Array, iv: Uint8Array}>}
 */
async function encryptContent(plaintext, key) {
    const iv = window.crypto.getRandomValues(new Uint8Array(12)); // 96 bits pour GCM
    const encoded = new TextEncoder().encode(plaintext);
    
    const ciphertext = await window.crypto.subtle.encrypt(
        { name: "AES-GCM", iv: iv },
        key,
        encoded
    );
    
    return {
        ciphertext: new Uint8Array(ciphertext),
        iv: iv
    };
}

/**
 * Déchiffre du contenu avec AES-256-GCM
 * @param {Uint8Array} ciphertext 
 * @param {CryptoKey} key 
 * @param {Uint8Array} iv 
 * @returns {Promise<string>}
 */
async function decryptContent(ciphertext, key, iv) {
    const decrypted = await window.crypto.subtle.decrypt(
        { name: "AES-GCM", iv: iv },
        key,
        ciphertext
    );
    
    return new TextDecoder().decode(decrypted);
}

// ============================================================
// HELPERS
// ============================================================

function arrayBufferToBase64(buffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}

function base64ToArrayBuffer(base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
}

// ============================================================
// API CALLS
// ============================================================

/**
 * Protège un document (envoie la DEK à l'API)
 * @param {Object} params
 * @returns {Promise<{success: boolean, docId: string}>}
 */
async function protectDocument({ dek, iv, documentName, classification, policy, userEmail }) {
    const response = await fetch(`${DRM_API_BASE}/keys/protect`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            dek: arrayBufferToBase64(dek),
            iv: arrayBufferToBase64(iv),
            documentName,
            classification,
            policy,
            tenantId: TENANT_ID,
            userEmail
        })
    });
    
    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error || "Failed to protect document");
    }
    
    return await response.json();
}

/**
 * Demande l'accès à un document protégé
 * @param {Object} params
 * @returns {Promise<{success: boolean, dek: string, rights: string[]}>}
 */
async function requestAccess({ docId, userEmail, authMethod }) {
    const response = await fetch(`${DRM_API_BASE}/keys/access`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            docId,
            tenantId: TENANT_ID,
            userEmail,
            authMethod
        })
    });
    
    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error || "Access denied");
    }
    
    return await response.json();
}

/**
 * Révoque l'accès à un document
 * @param {Object} params
 * @returns {Promise<{success: boolean}>}
 */
async function revokeDocument({ docId, userEmail }) {
    const response = await fetch(`${DRM_API_BASE}/keys/revoke`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            docId,
            tenantId: TENANT_ID,
            userEmail
        })
    });
    
    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error || "Failed to revoke");
    }
    
    return await response.json();
}

// ============================================================
// MAIN DRM FUNCTIONS (pour l'add-in)
// ============================================================

/**
 * Protège le document actuel
 * @param {string} content - Contenu du document
 * @param {string} documentName - Nom du fichier
 * @param {string} classification - PUBLIC, INTERNAL, CONFIDENTIAL, RESTRICTED
 * @param {Object} policy - { allowedDomains, allowedEmails, rights, expiry }
 * @param {string} userEmail - Email de l'utilisateur
 * @returns {Promise<{docId: string, encryptedContent: string, iv: string}>}
 */
async function protectCurrentDocument(content, documentName, classification, policy, userEmail) {
    console.log("🔐 Protecting document:", documentName);
    
    // 1. Générer une DEK
    const dek = await generateDEK();
    const dekBytes = await exportKey(dek);
    console.log("✅ DEK generated");
    
    // 2. Chiffrer le contenu
    const { ciphertext, iv } = await encryptContent(content, dek);
    console.log("✅ Content encrypted");
    
    // 3. Envoyer la DEK à l'API
    const result = await protectDocument({
        dek: dekBytes,
        iv: iv,
        documentName,
        classification,
        policy,
        userEmail
    });
    console.log("✅ DEK stored, docId:", result.docId);
    
    // 4. Retourner le contenu chiffré + docId
    return {
        docId: result.docId,
        encryptedContent: arrayBufferToBase64(ciphertext),
        iv: arrayBufferToBase64(iv),
        policy: result.policy
    };
}

/**
 * Ouvre un document protégé
 * @param {string} encryptedContent - Contenu chiffré (base64)
 * @param {string} iv - IV (base64)
 * @param {string} docId - ID du document
 * @param {string} userEmail - Email de l'utilisateur
 * @param {string} authMethod - azure_ad, google, email_otp
 * @returns {Promise<{content: string, rights: string[], classification: string}>}
 */
async function openProtectedDocument(encryptedContent, iv, docId, userEmail, authMethod) {
    console.log("🔓 Opening protected document:", docId);
    
    // 1. Demander l'accès (récupérer la DEK)
    const access = await requestAccess({ docId, userEmail, authMethod });
    console.log("✅ Access granted, rights:", access.rights);
    
    // 2. Importer la DEK
    const dekBytes = base64ToArrayBuffer(access.dek);
    const dek = await importKey(dekBytes);
    console.log("✅ DEK imported");
    
    // 3. Déchiffrer le contenu
    const ciphertext = base64ToArrayBuffer(encryptedContent);
    const ivBytes = base64ToArrayBuffer(iv);
    const content = await decryptContent(ciphertext, dek, ivBytes);
    console.log("✅ Content decrypted");
    
    return {
        content,
        rights: access.rights,
        classification: access.classification,
        documentName: access.documentName,
        watermark: access.watermark
    };
}

// ============================================================
// EXPORTS (pour utiliser dans le taskpane)
// ============================================================

window.MazeShieldDRM = {
    protectCurrentDocument,
    openProtectedDocument,
    revokeDocument,
    // Helpers exposés pour debug
    generateDEK,
    encryptContent,
    decryptContent
};

console.log("🛡️ MazeShield DRM Client loaded");
