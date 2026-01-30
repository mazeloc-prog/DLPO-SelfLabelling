/*
 * MazeShield - Smart Alerts Handler
 * Ce fichier est utilisé par Outlook Desktop (Windows classique)
 */

// Labels sensibles
const SENSITIVE_LABELS = ['Confidential', 'Restricted', 'CONFIDENTIAL', 'RESTRICTED'];

// Azure Function URL pour audit
const AUDIT_URL = "https://dlpo-audit-api.azurewebsites.net/api/log-event";

// Variables globales
let internalDomain = null;
let currentUser = null;

// Handler principal
function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;
    
    // Récupérer le domaine de l'utilisateur (SSO dynamique)
    currentUser = Office.context.mailbox.userProfile.emailAddress;
    const atIndex = currentUser.indexOf('@');
    if (atIndex > -1) {
      internalDomain = currentUser.substring(atIndex + 1).toLowerCase();
    }
    
    // Récupérer le sujet
    item.subject.getAsync(function(subjectResult) {
      if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
        event.completed({ allowEvent: true });
        return;
      }
      
      const subject = subjectResult.value || '';
      const label = extractLabelFromSubject(subject);
      
      // Récupérer les destinataires
      getAllRecipients(item, function(recipients) {
        const externalRecipients = recipients.filter(function(r) {
          return isExternal(r.emailAddress);
        });
        
        const hasExternal = externalRecipients.length > 0;
        const isSensitive = label && SENSITIVE_LABELS.some(function(l) {
          return l.toLowerCase() === label.toLowerCase();
        });
        
        console.log('MazeShield Check:', { label: label, isSensitive: isSensitive, hasExternal: hasExternal });
        
        if (isSensitive && hasExternal) {
          // Log l'alerte
          logAudit({
            action: "EMAIL_SEND_BLOCKED",
            label: label,
            hasExternal: true,
            externalCount: externalRecipients.length,
            externalRecipients: externalRecipients.map(function(r) { return r.emailAddress; }),
            user: currentUser
          });
          
          var externalEmails = externalRecipients.map(function(r) { return r.emailAddress; }).join('\n• ');
          
          event.completed({ 
            allowEvent: false,
            errorMessage: "⚠️ ALERTE MAZESHIELD\n\n" +
              "Cet email est classifié \"" + label.toUpperCase() + "\" et contient des destinataires EXTERNES :\n\n" +
              "• " + externalEmails + "\n\n" +
              "Voulez-vous vraiment envoyer cet email sensible à l'extérieur de l'organisation ?"
          });
        } else {
          event.completed({ allowEvent: true });
        }
      });
    });
  } catch (error) {
    console.error("MazeShield Error:", error);
    event.completed({ allowEvent: true });
  }
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

// IMPORTANT: Mapper le handler
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
