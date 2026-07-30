/* ============================================================
   LETTERE — script di consegna
   ------------------------------------------------------------
   Riceve un indirizzo mail dal sito, lo cerca nel foglio Google
   e spedisce a quell'indirizzo il PDF corrispondente.

   Le lettere restano in una cartella privata su Drive: sul sito
   pubblico non c'è nessun file da scaricare.
   ============================================================ */


/* ===== 1. CONFIGURAZIONE — le uniche righe da modificare ===== */

// L'ID del foglio Google (lo trovi nel suo indirizzo, tra /d/ e /edit)
const ID_FOGLIO = "INCOLLA_QUI_L_ID_DEL_FOGLIO";

// Il nome della scheda in basso nel foglio
const NOME_SCHEDA = "Lettere";

// L'ID della cartella Drive con i PDF (nell'indirizzo, dopo /folders/)
const ID_CARTELLA = "INCOLLA_QUI_L_ID_DELLA_CARTELLA";

// Il messaggio che accompagna la lettera. {nome} viene sostituito
// con il nome scritto nel foglio.
const OGGETTO = "La tua lettera";
const TESTO =
  "Ciao {nome},\n\n" +
  "ho scritto un pensiero per te dopo quello che abbiamo vissuto insieme.\n" +
  "Lo trovi qui in allegato.\n\n" +
  "Un abbraccio";

// Quante volte la stessa persona può richiedere la lettera in un'ora
const MAX_INVII_ORA = 3;


/* ===== 2. Da qui in giù non serve toccare niente ===== */


/**
 * Riceve la richiesta dal sito.
 * Risponde sempre "ok", anche se l'indirizzo non è in lista:
 * così nessuno può usare la pagina per scoprire chi c'è nella lista.
 */
function doPost(e) {
  try {
    const dati = JSON.parse(e.postData.contents);
    const email = String(dati.email || "").trim().toLowerCase();
    if (email) {
      inviaLettera(email);
    }
  } catch (errore) {
    console.error("Richiesta non valida: " + errore);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Se apri l'indirizzo dello script nel browser, vedi questo.
 * Serve solo a controllare che la pubblicazione sia andata a buon fine.
 */
function doGet() {
  return ContentService.createTextOutput("Il servizio è attivo.");
}


/**
 * Cerca l'indirizzo nel foglio e spedisce il PDF associato.
 */
function inviaLettera(email) {
  const cache = CacheService.getScriptCache();
  const chiave = "invii_" + email;
  const giaInviate = Number(cache.get(chiave) || 0);

  // Freno anti-abuso: evita che qualcuno tempesti di mail un amico
  if (giaInviate >= MAX_INVII_ORA) {
    console.warn("Troppe richieste per " + email);
    return;
  }

  const scheda = SpreadsheetApp.openById(ID_FOGLIO).getSheetByName(NOME_SCHEDA);
  if (!scheda) {
    console.error("Scheda non trovata: " + NOME_SCHEDA);
    return;
  }

  const righe = scheda.getDataRange().getValues();

  // Parte da 1 per saltare la riga delle intestazioni
  for (let i = 1; i < righe.length; i++) {
    const nome = String(righe[i][0]).trim();
    const indirizzo = String(righe[i][1]).trim().toLowerCase();
    const nomeFile = String(righe[i][2]).trim();

    if (!indirizzo || indirizzo !== email) continue;

    const file = DriveApp.getFolderById(ID_CARTELLA).getFilesByName(nomeFile);
    if (!file.hasNext()) {
      console.error("PDF non trovato nella cartella: " + nomeFile);
      return;
    }

    MailApp.sendEmail({
      to: indirizzo,
      subject: OGGETTO,
      body: TESTO.replace("{nome}", nome),
      attachments: [file.next().getBlob()]
    });

    cache.put(chiave, String(giaInviate + 1), 3600);
    registraConsegna(nome, indirizzo);
    return;
  }

  // Indirizzo non presente in lista: non si fa niente, in silenzio.
  console.info("Indirizzo non in lista: " + email);
}


/**
 * Segna data e ora di ogni consegna in una scheda "Consegne",
 * così vedi chi ha già ritirato la sua lettera. La crea al bisogno.
 */
function registraConsegna(nome, email) {
  try {
    const foglio = SpreadsheetApp.openById(ID_FOGLIO);
    let log = foglio.getSheetByName("Consegne");

    if (!log) {
      log = foglio.insertSheet("Consegne");
      log.appendRow(["Data e ora", "Nome", "Email"]);
    }

    log.appendRow([new Date(), nome, email]);
  } catch (errore) {
    console.error("Non sono riuscito a registrare la consegna: " + errore);
  }
}


/**
 * PROVA — scrivi qui sotto il tuo indirizzo e premi "Esegui" nell'editor.
 * Serve a due cose: dare i permessi allo script la prima volta
 * e verificare che la lettera arrivi davvero.
 */
function provaInvio() {
  inviaLettera("c.schiavella@gmail.com");
}
