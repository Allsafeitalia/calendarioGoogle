# Script per Creare Eventi di Calendario da Google Sheets

Questo script utilizza Google Apps Script per creare eventi in un calendario Google a partire dai dati presenti in un foglio Google Sheets. È utile per automatizzare la creazione degli eventi in calendario direttamente da un foglio, evitando l'inserimento manuale.

## Funzionalità

- Legge le informazioni sugli eventi da un foglio Google Sheets.
- Crea eventi nel calendario Google con titolo, descrizione, data e orari specificati.
- Permette la personalizzazione per l’utilizzo di calendari specifici.

## Requisiti

- Un account Google con accesso a Google Sheets e Google Calendar.
- Permessi per eseguire script su Google Apps.

## Configurazione

1. **Crea un nuovo foglio Google Sheets** e aggiungi le seguenti intestazioni nella prima riga:

   - `Titolo`
   - `Descrizione`
   - `Data Inizio`
   - `Ora Inizio`
   - `Data Fine`
   - `Ora Fine`

2. **Aggiungi i dettagli degli eventi** sotto le intestazioni create:
   - Ogni riga rappresenta un evento.
   - `Data Inizio` e `Ora Inizio` indicano quando inizia l'evento.
   - `Data Fine` e `Ora Fine` indicano quando termina l'evento.

3. **Aggiungi lo script** al foglio Google Sheets:

   - Vai su **Estensioni > Apps Script**.
   - Incolla il seguente codice e salva il progetto.

   ```javascript
   function creaEventiCalendario() {
     // Ottieni il foglio di calcolo attivo
     var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
     var sheet = spreadsheet.getActiveSheet();

     // Ottieni il calendario su cui creare gli eventi (principale dell'utente)
     var calendar = CalendarApp.getDefaultCalendar();
     
     // Definisci l'intervallo di righe con i dati degli eventi
     var dataRange = sheet.getDataRange();
     var data = dataRange.getValues();

     // Loop sugli eventi (salta la prima riga con le intestazioni)
     for (var i = 1; i < data.length; i++) {
       var row = data[i];
       
       // Estrai i dettagli dell'evento
       var titolo = row[0];
       var descrizione = row[1];
       var dataInizio = new Date(row[2] + " " + row[3]);
       var dataFine = new Date(row[4] + " " + row[5]);
       
       // Crea l'evento nel calendario
       calendar.createEvent(titolo, dataInizio, dataFine, {
         description: descrizione
       });
     }
     
     // Messaggio di conferma
     SpreadsheetApp.getUi().alert("Eventi creati con successo nel calendario!");
   }
