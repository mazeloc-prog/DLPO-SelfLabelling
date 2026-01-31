/*
 * MazeShield - Smart Alerts Handler v2.1
 * Avec popup justification style MIP (iFrame = pas de Allow)
 */

const SENSITIVE_LABELS = ['Confidential', 'Restricted', 'CONFIDENTIAL', 'RESTRICTED'];
const AUDIT_URL = "https://dlpo-audit-api.azurewebsites.net/api/log-event";
const DIALOG_URL = "https://mazeloc-prog.github.io/DLPO-SelfLabelling/justification-dialog.html";

let internalDomain = null;
let currentUser = null;

function onMessageSendHandler(event) {
    try {
        const item = Office.context.mailbox.item;
        
        // 1. Récupérer le domaine de l'utilisateur (SSO)
        currentUser = Office.context.mailbox.userProfile.emailAddress;
        const atIndex = currentUser.indexOf('@');
        if (atIndex > -1) {
            internalDomain = currentUser.substring(atIndex + 1).toLowerCase();
        }
        
        // 2. Récupérer le sujet pour le label
        item.subject.getAsync(function(subjectResult) {
            if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
                event.completed({ allowEvent: true });
                return;
            }
            
            const subject = subjectResult.value || '';
            let label = extractLabelFromSubject(subject);
            
            // 3. Si pas de label dans le sujet, vérifier custom properties
            if (!label) {
                item.loadCustomPropertiesAsync(function(propsResult) {
                    if (propsResult.status === Office.AsyncResultStatus.Succeeded) {
                        label = propsResult.value.get("MazeShield_Classification");
                    }
                    checkRecipientsAndProcess(item, label, subject, event);
                });
            } else {
                checkRecipientsAndProcess(item, label, subject, event);
            }
        });
    } catch (error) {
        console.error("MazeShield Error:", error);
        event.completed({ allowEvent: true });
    }
}

function checkRecipientsAndProcess(item, label, subject, event) {
    getAllRecipients(item, function(recipients) {
        const externalRecipients = recipients.filter(function(r) {
            return isExternal(r.emailAddress);
        });
        
        const hasExternal = externalRecipients.length > 0;
        const isSensitive = label && SENSITIVE_LABELS.some(function(l) {
            return l.toLowerCase() === label.toLowerCase();
        });
        
        console.log('MazeShield Check:', { label, isSensitive, hasExternal });
        
        // Si sensible + externe → Ouvrir le dialog de justification
        if (isSensitive && hasExternal) {
            openJustificationDialog(label, externalRecipients, subject, event);
        } else {
            event.completed({ allowEvent: true });
        }
    });
}

function openJustificationDialog(label, externalRecipients, subject, event) {
    const recipientEmails = externalRecipients.map(function(r) { return r.emailAddress; }).join(', ');
    
    // Construire l'URL avec les paramètres
    const dialogUrl = DIALOG_URL + 
        '?label=' + encodeURIComponent(label) +
        '&recipients=' + encodeURIComponent(recipientEmails) +
        '&subject=' + encodeURIComponent(subject);
    
    // displayInIframe: true = PAS de popup "Allow/Ignore" !
    Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 70, width: 40, displayInIframe: true },
        function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Dialog failed:", asyncResult.error.message);
                // Fallback : utiliser la popup native
                showNativeAlert(label, recipientEmails, event);
                return;
            }
            
            var dialog = asyncResult.value;
            
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                dialog.close();
                
                try {
                    var response = JSON.parse(arg.message);
                    
                    if (response.action === 'SEND') {
                        // Log la justification dans Azure
                        logAudit({
                            action: "EMAIL_SENT_WITH_JUSTIFICATION",
                            label: label,
                            justificationCode: response.justification,
                            justificationText: response.justificationText,
                            comment: response.comment,
                            externalRecipients: externalRecipients.map(function(r) { return r.emailAddress; }),
                            subject: subject
                        });
                        
                        // Autoriser l'envoi
                        event.completed({ allowEvent: true });
                    } else {
                        // Annulé → Bloquer l'envoi
                        logAudit({
                            action: "EMAIL_SEND_CANCELLED",
                            label: label,
                            externalRecipients: externalRecipients.map(function(r) { return r.emailAddress; })
                        });
                        
                        event.completed({ allowEvent: false });
                    }
                } catch (e) {
                    console.error("Parse error:", e);
                    event.completed({ allowEvent: false });
                }
            });
            
            dialog.addEventHandler(Office.EventType.DialogEventReceived, function(arg) {
                // Dialog fermé sans réponse
                console.log("Dialog closed:", arg.error);
                event.completed({ allowEvent: false });
            });
        }
    );
}

// Fallback si le dialog ne marche pas
function showNativeAlert(label, recipientEmails, event) {
    event.completed({
        allowEvent: false,
        errorMessage: "⚠️ MAZESHIELD ALERT\n\n" +
            "This email is classified \"" + label.toUpperCase() + "\" and contains EXTERNAL recipients:\n\n" +
            recipientEmails + "\n\n" +
            "Please use the MazeShield taskpane to provide justification before sending."
    });
}

function extractLabelFromSubject(subject) {
    var match = subject.match(/\[(PUBLIC|INTERNAL|CONFIDENTIAL|RESTRICTED)\]/i);
    return match ? match[1] : null;
}

function getAllRecipients(item, callback) {
    var recipients = [];
    var pending = 3;
    
    function done() {
        pending--;
        if (pending === 0) {
            callback(recipients);
        }
    }
    
    item.to.getAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            recipients = recipients.concat(result.value || []);
        }
        done();
    });
    
    item.cc.getAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            recipients = recipients.concat(result.value || []);
        }
        done();
    });
    
    item.bcc.getAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            recipients = recipients.concat(result.value || []);
        }
        done();
    });
}

function isExternal(email) {
    if (!internalDomain) return false;
    return !email.toLowerCase().endsWith('@' + internalDomain);
}

function logAudit(data) {
    try {
        var xhr = new XMLHttpRequest();
        xhr.open("POST", AUDIT_URL, true);
        xhr.setRequestHeader("Content-Type", "application/json");
        xhr.send(JSON.stringify({
            tenantId: internalDomain,
            user: currentUser,
            timestamp: new Date().toISOString(),
            source: "outlook-smart-alert",
            ...data
        }));
    } catch (e) {
        console.error("Audit log failed:", e);
    }
}

// Mapper le handler
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
