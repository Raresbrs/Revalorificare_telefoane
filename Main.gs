/**
@OnlyCurrentDoc
Acest fișier conține scripturi pentru un sistem de notificare și procesare a răspunsurilor.
Funcționalități:
Trimiterea automată de emailuri de notificare inițiale (cu memorie pentru ultima linie
procesată).
Trimiterea manuală a unui email de reiterare la un interval definit (ex: 6 luni).
Verificarea răspunsurilor dintr-un formular cu logică de validare (ID Angajat + Email).
Actualizarea condiționată a foii de calcul principale doar pentru răspunsurile pozitive.
Protecție la suprascriere pentru a accepta doar primul răspuns valid.
Notificare automată prin email către admin în caz de erori. */

// --- CONSTANTE GLOBALE PENTRU FOAIA PRINCIPALĂ ---
const COL_EMPLOYEE_ID = 0; // Coloana A
const COL_EMAIL_EMPLOYEE = 2; // Coloana C
const COL_LAST_NAME = 3; // Coloana D
const COL_FIRST_NAME = 4; // Coloana E
const COL_EMAIL_SUPERVISOR = 5; // Coloana F
const COL_DEVICE_MODEL = 7; // Coloana H
const COL_DEVICE_SERIAL = 8; // Coloana I (ex: IMEI)
const COL_ALLOCATION_DATE = 9; // Coloana J
const COL_PRICE = 10; // Coloana K
const COL_EMAIL_STATUS = 11; // Coloana L
const COL_FIRST_EMAIL_DATE = 12; // Coloana M
const COL_LAST_EMAIL_DATE = 13; // Coloana N (Folosită pentru reiterare)
const COL_CONFIRMATION_TIMESTAMP = 14; // Coloana O
// --- SCRIPT PENTRU TRIMITEREA EMAILURILOR INIȚIALE CU MEMORIE ---
function sendInitialEmails() {
const sheetName = &quot;MainSheet&quot;;
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName(sheetName);
const ui = SpreadsheetApp.getUi();
if (!sheet) {
const errorMsg = Foaia de calcul &quot;${sheetName}&quot; nu a fost găsită!;
Logger.log(errorMsg);
sendErrorNotification(&#39;sendInitialEmails&#39;, [errorMsg]);
ui.alert(errorMsg);
return;
}
const scriptProperties = PropertiesService.getScriptProperties();
const lastProcessedRow = parseInt(scriptProperties.getProperty(&#39;lastProcessedRow&#39;) || &#39;1&#39;);
const startRow = lastProcessedRow + 1;

if (startRow &gt; sheet.getLastRow()) {
ui.alert(&#39;Procesare finalizată! Toate liniile au fost deja verificate.&#39;);
return;
}
const promptResponse = ui.prompt(
&#39;Confirmare Rulare&#39;,
Scriptul va începe procesarea de la rândul ${startRow}.\nIntroduceți numărul de linii pe
care doriți să le procesați:,
ui.ButtonSet.OK_CANCEL
);
if (promptResponse.getSelectedButton() !== ui.Button.OK) {
Logger.log(&quot;Utilizatorul a anulat operația.&quot;);
return;
}
const numLinesToProcess = parseInt(promptResponse.getResponseText(), 10);
if (isNaN(numLinesToProcess) || numLinesToProcess &lt;= 0) {
ui.alert(&#39;Număr invalid. Vă rugăm să introduceți un număr pozitiv.&#39;);
return;
}
const maxRowsToFetch = Math.min(numLinesToProcess, sheet.getLastRow() - startRow +
1);
if (maxRowsToFetch &lt;= 0) {
ui.alert(Nu mai sunt linii de procesat începând cu rândul ${startRow}.);
return;
}
const dataRange = sheet.getRange(startRow, 1, maxRowsToFetch,
sheet.getLastColumn());
const dataValues = dataRange.getValues();
const currentDate = new Date();
currentDate.setHours(0, 0, 0, 0);
let emailsSent = 0;
let errorsEncountered = [];
let emailsSkipped = 0;
for (let i = 0; i &lt; dataValues.length; i++) {
const currentRowInSheet = startRow + i;
try {
const rowData = dataValues[i];
const confirmationTimestamp = rowData[COL_CONFIRMATION_TIMESTAMP];
if (confirmationTimestamp) {
Logger.log(`Rând ${currentRowInSheet}: Răspuns deja confirmat. Se omite.`);

emailsSkipped++;
continue;
}
const employeeEmail = rowData[COL_EMAIL_EMPLOYEE];
const supervisorEmail = rowData[COL_EMAIL_SUPERVISOR];
let finalRecipient = &quot;&quot;;
if (employeeEmail &amp;&amp; String(employeeEmail).trim() !== &quot;&quot;) {
finalRecipient = employeeEmail;
} else if (supervisorEmail &amp;&amp; String(supervisorEmail).trim() !== &quot;&quot;) {
finalRecipient = supervisorEmail;
}
if (!finalRecipient) {
Logger.log(`Rând ${currentRowInSheet}: Adresa de email lipsește. Se omite.`);
continue;
}
const allocationDateRaw = rowData[COL_ALLOCATION_DATE];
const emailStatus = rowData[COL_EMAIL_STATUS];
if (!(allocationDateRaw instanceof Date) || isNaN(allocationDateRaw.getTime())) {
Logger.log(`Rând ${currentRowInSheet}: Data alocării este invalidă. Se omite.`);
if (emailStatus !== &quot;TRIMIS&quot; &amp;&amp; !String(emailStatus).startsWith(&quot;EROARE&quot;)) {
sheet.getRange(currentRowInSheet, COL_EMAIL_STATUS +
1).setValue(&quot;DATA_INVALIDA&quot;);
}
continue;
}
let shouldSendEmail = false;
if (!emailStatus &amp;&amp; !String(emailStatus).startsWith(&quot;EROARE&quot;)) {
const notificationDate = new Date(allocationDateRaw);
notificationDate.setFullYear(notificationDate.getFullYear() + 2);
notificationDate.setMonth(notificationDate.getMonth() - 1);
if (currentDate &gt;= notificationDate) {
shouldSendEmail = true;
}
}
if (shouldSendEmail) {
const lastName = rowData[COL_LAST_NAME];
const firstName = rowData[COL_FIRST_NAME];
const deviceModel = rowData[COL_DEVICE_MODEL];
const deviceSerial = rowData[COL_DEVICE_SERIAL];
const priceRaw = rowData[COL_PRICE];

const subject = `Oportunitate achiziție dispozitiv: ${deviceModel}`;
const priceText = typeof priceRaw === &#39;number&#39; ? priceRaw.toFixed(2) :
String(priceRaw).replace(&#39;,&#39;, &#39;.&#39;);
const formLink = &quot;https://docs.google.com/forms/d/e/YOUR_FORM_ID_HERE/viewform&quot;;
const htmlBody = `&lt; p &gt;Bună ziua ${firstName} ${lastName},&lt;/p&gt;&lt; p &gt;Dispozitivul dvs. &lt;
strong &gt;${deviceModel}&lt;/strong&gt; (Serie: &lt; strong &gt;${deviceSerial}&lt;/strong&gt;) poate fi
achiziționat la prețul de &lt; strong &gt;${priceText} RON&lt;/strong&gt;.&lt;/p&gt;&lt; p &gt;Dacă doriți să îl
achiziționați, vă rugăm să completați formularul de mai jos:&lt;/p&gt;&lt;p style=&quot;margin-top: 15px;
margin-bottom: 15px;&quot;&gt;&lt;a href=&quot;${formLink}&quot; style=&quot;background-color: #4CAF50; color:
white; padding: 10px 20px; text-align: center; text-decoration: none; display: inline-block;
border-radius: 5px;&quot;&gt;Accesează Formularul&lt;/a&gt;&lt;/p&gt;&lt; p &gt;Vă mulțumim,&lt; br &gt;&lt; b &gt;Echipa
IT&lt;/b&gt;&lt;/p&gt;`;
GmailApp.sendEmail(finalRecipient, subject, &quot;&quot;, { htmlBody: htmlBody });
const sentDate = new Date();
sheet.getRange(currentRowInSheet, COL_EMAIL_STATUS + 1).setValue(&quot;TRIMIS&quot;);
sheet.getRange(currentRowInSheet, COL_FIRST_EMAIL_DATE +
1).setValue(sentDate);
emailsSent++;
Logger.log(`Rând ${currentRowInSheet}: Email trimis cu succes către ${finalRecipient}.`);
} else {
emailsSkipped++;
}
} catch (e) {
const errorMessage = `Eroare la procesarea rândului ${currentRowInSheet}:
${e.message}`;
Logger.log(errorMessage);
errorsEncountered.push(errorMessage);
try {
sheet.getRange(currentRowInSheet, COL_EMAIL_STATUS + 1).setValue(`EROARE:
${e.message.substring(0, 150)}`);
} catch (sheetError) {
Logger.log(`Nu s-a putut scrie eroarea în foaia de calcul la rândul ${currentRowInSheet}:
${sheetError.message}`);
}
}
Copy
}
const newLastProcessedRow = startRow + dataValues.length - 1;
scriptProperties.setProperty(&#39;lastProcessedRow&#39;, newLastProcessedRow.toString());
if (errorsEncountered.length &gt; 0) {

sendErrorNotification(&#39;sendInitialEmails&#39;, errorsEncountered);
}
SpreadsheetApp.flush();
ui.alert(Procesare finalizată pentru rândurile ${startRow} -
${newLastProcessedRow}.\n\nEmailuri trimise: ${emailsSent}\nEmailuri sărite:
${emailsSkipped}\nErori: ${errorsEncountered.length}\n\nUrmătoarea rulare va începe de la
rândul ${newLastProcessedRow + 1}.);
}
// --- FUNCȚIE PENTRU TRIMITEREA REITERĂRILOR ---
function sendReminderEmails() {
const sheetName = &quot;MainSheet&quot;;
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName(sheetName);
const ui = SpreadsheetApp.getUi();
if (!sheet) {
const errorMsg = Foaia de calcul &quot;${sheetName}&quot; nu a fost găsită!;
Logger.log(errorMsg);
sendErrorNotification(&#39;sendReminderEmails&#39;, [errorMsg]);
ui.alert(errorMsg);
return;
}
const promptResponse = ui.alert(
&#39;Confirmare Trimitere Reitărari&#39;,
&#39;Acest script va verifica TOATE rândurile pentru a trimite emailuri de
reiterare.\n\nCondiții:\n- Au trecut cel puțin 23 de luni de la alocare.\n- Au trecut cel puțin 6
luni de la ultimul email.\n- Nu s-a primit încă un răspuns.\n\nDoriți să continuați?&#39;,
ui.ButtonSet.YES_NO
);
if (promptResponse !== ui.Button.YES) {
Logger.log(&quot;Utilizatorul a anulat operația de reiterare.&quot;);
return;
}
const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
const dataValues = dataRange.getValues();
const currentDate = new Date();
currentDate.setHours(0, 0, 0, 0);
let emailsSent = 0;
let errorsEncountered = [];
let emailsSkipped = 0;

for (let i = 0; i &lt; dataValues.length; i++) {
const currentRowInSheet = i + 2;
try {
const rowData = dataValues[i];
const confirmationTimestamp = rowData[COL_CONFIRMATION_TIMESTAMP];
if (confirmationTimestamp) {
emailsSkipped++;
continue;
}
const emailStatus = rowData[COL_EMAIL_STATUS];
if (emailStatus !== &quot;TRIMIS&quot; &amp;&amp; emailStatus !== &quot;TRIMIS+REITERARE&quot;) {
emailsSkipped++;
continue;
}
const allocationDateRaw = rowData[COL_ALLOCATION_DATE];
if (!(allocationDateRaw instanceof Date) || isNaN(allocationDateRaw.getTime())) {
emailsSkipped++;
continue;
}
const eligibilityDate = new Date(allocationDateRaw);
eligibilityDate.setMonth(eligibilityDate.getMonth() + 23);
if (currentDate &lt; eligibilityDate) {
emailsSkipped++;
continue;
}
const lastEmailDateRaw = rowData[COL_LAST_EMAIL_DATE];
const firstEmailDateRaw = rowData[COL_FIRST_EMAIL_DATE];
let lastSentDate = null;
if (lastEmailDateRaw instanceof Date &amp;&amp; !isNaN(lastEmailDateRaw.getTime())) {
lastSentDate = lastEmailDateRaw;
} else if (firstEmailDateRaw instanceof Date &amp;&amp; !isNaN(firstEmailDateRaw.getTime())) {
lastSentDate = firstEmailDateRaw;
}
if (!lastSentDate) {
emailsSkipped++;
continue;
}
const reminderDate = new Date(lastSentDate);
reminderDate.setMonth(reminderDate.getMonth() + 6);

if (currentDate &gt;= reminderDate) {
const employeeEmail = rowData[COL_EMAIL_EMPLOYEE];
const supervisorEmail = rowData[COL_EMAIL_SUPERVISOR];
let finalRecipient = &quot;&quot;;
if (employeeEmail &amp;&amp; String(employeeEmail).trim() !== &quot;&quot;) {
finalRecipient = employeeEmail;
} else if (supervisorEmail &amp;&amp; String(supervisorEmail).trim() !== &quot;&quot;) {
finalRecipient = supervisorEmail;
}
if (!finalRecipient) {
continue;
}
const lastName = rowData[COL_LAST_NAME];
const firstName = rowData[COL_FIRST_NAME];
const deviceModel = rowData[COL_DEVICE_MODEL];
const deviceSerial = rowData[COL_DEVICE_SERIAL];
const priceRaw = rowData[COL_PRICE];
const subject = `[REITERARE] Oportunitate achiziție dispozitiv: ${deviceModel}`;
const priceText = typeof priceRaw === &#39;number&#39; ? priceRaw.toFixed(2) :
String(priceRaw).replace(&#39;,&#39;, &#39;.&#39;);
const formLink = &quot;https://docs.google.com/forms/d/e/YOUR_FORM_ID_HERE/viewform&quot;;
const htmlBody = `&lt; p &gt;Bună ziua ${firstName} ${lastName},&lt;/p&gt;&lt; p &gt;Acesta este un
memento referitor la posibilitatea de a achiziționa dispozitivul &lt; strong
&gt;${deviceModel}&lt;/strong&gt; (Serie: &lt; strong &gt;${deviceSerial}&lt;/strong&gt;), la prețul de &lt; strong
&gt;${priceText} RON&lt;/strong&gt;.&lt;/p&gt;&lt; p &gt;Dacă doriți să îl achiziționați, vă rugăm să completați
formularul de mai jos:&lt;/p&gt;&lt;p style=&quot;margin-top: 15px; margin-bottom: 15px;&quot;&gt;&lt;a
href=&quot;${formLink}&quot; style=&quot;background-color: #4CAF50; color: white; padding: 10px 20px;
text-align: center; text-decoration: none; display: inline-block; border-radius:
5px;&quot;&gt;Accesează Formularul&lt;/a&gt;&lt;/p&gt;&lt; p &gt;Vă mulțumim,&lt; br &gt;Echipa IT&lt;/p&gt;`;
GmailApp.sendEmail(finalRecipient, subject, &quot;&quot;, { htmlBody: htmlBody });
const reminderSentDate = new Date();
sheet.getRange(currentRowInSheet, COL_EMAIL_STATUS +
1).setValue(&quot;TRIMIS+REITERARE&quot;);
sheet.getRange(currentRowInSheet, COL_LAST_EMAIL_DATE +
1).setValue(reminderSentDate);
emailsSent++;
Logger.log(`Rând ${currentRowInSheet}: Email de reiterare trimis cu succes către
${finalRecipient}.`);

} else {
emailsSkipped++;
}
} catch (e) {
const errorMessage = `Eroare la procesarea rândului ${currentRowInSheet} pentru
reiterare: ${e.message}`;
Logger.log(errorMessage);
errorsEncountered.push(errorMessage);
}
Copy
}
if (errorsEncountered.length &gt; 0) {
sendErrorNotification(&#39;sendReminderEmails&#39;, errorsEncountered);
}
SpreadsheetApp.flush();
ui.alert(Procesare reitărări finalizată.\n\nEmailuri de reiterare trimise:
${emailsSent}\nRânduri verificate/sărite: ${emailsSkipped + emailsSent +
errorsEncountered.length}\nErori: ${errorsEncountered.length});
}
// --- FUNCȚIE PENTRU NOTIFICARE EROARE CĂTRE ADMIN ---
function sendErrorNotification(functionName, errorsArray) {
const ADMIN_EMAIL = &quot;admin@example.com&quot;;
if (!ADMIN_EMAIL) {
Logger.log(&quot;Adresa de email a administratorului nu este configurată.&quot;);
return;
}
const timestamp = new Date().toLocaleString(&quot;ro-RO&quot;, { timeZone:
Session.getScriptTimeZone() });
const subject = [EROARE SCRIPT] Notificare de Eroare - ${functionName};
let errorDetails = &quot;&quot;;
errorsArray.forEach((error, index) =&gt; {
errorDetails += &lt; p &gt;&lt; b &gt;Eroarea ${index + 1}:&lt;/b&gt;&lt; br &gt;&lt;pre style=&quot;background-
color:#f5f5f5; padding: 10px; border-radius: 4px; white-space: pre-wrap; word-wrap: break-
word;&quot;&gt;${error}&lt;/pre&gt;&lt;/p&gt;;
});
const htmlBody = &lt; html &gt;&lt; body &gt;&lt; h2 &gt;Raport de Erori Automat&lt;/h2&gt;&lt; p &gt;Au fost
detectate erori în timpul rulării scriptului.&lt;/p&gt;&lt; hr &gt;&lt; p &gt;&lt; b &gt;Data și Ora:&lt;/b&gt;
${timestamp}&lt;/p&gt;&lt; p &gt;&lt; b &gt;Funcția:&lt;/b&gt; ${functionName}&lt;/p&gt;&lt; p &gt;&lt; b &gt;Total erori:&lt;/b&gt;
${errorsArray.length}&lt;/p&gt;&lt; hr &gt;&lt; h3 &gt;Detalii Erori:&lt;/h3&gt;${errorDetails}&lt;/body&gt;&lt;/html&gt;;

try {
GmailApp.sendEmail(ADMIN_EMAIL, subject, &quot;&quot;, { htmlBody: htmlBody });
Logger.log(Notificare de eroare trimisă către ${ADMIN_EMAIL}.);
} catch (e) {
Logger.log(A eșuat trimiterea notificării de eroare: ${e.message});
}
}
// --- FUNCȚIE PENTRU VERIFICAREA RĂSPUNSURILOR DIN FORMULAR ---
function processFormResponses() {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const mainSheet = ss.getSheetByName(&#39;MainSheet&#39;);
const formResponsesSheet = ss.getSheetByName(&#39;FormResponsesSheet&#39;);
if (!mainSheet || !formResponsesSheet) {
const errorMsg = &quot;Una dintre foile de calcul &#39;MainSheet&#39; sau &#39;FormResponsesSheet&#39;
lipsește.&quot;;
Logger.log(errorMsg);
sendErrorNotification(&#39;processFormResponses&#39;, [errorMsg]);
return;
}
const colStatusForm = 7; // Coloana G în foaia de răspunsuri
const colEmployeeIdForm = 8; // Coloana H în foaia de răspunsuri
const mainSheetData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1,
mainSheet.getLastColumn()).getValues();
const formResponsesData = formResponsesSheet.getRange(2, 1,
formResponsesSheet.getLastRow() - 1,
formResponsesSheet.getLastColumn()).getValues();
const mainSheetMap = new Map();
mainSheetData.forEach((row, index) =&gt; {
const employeeId = row[COL_EMPLOYEE_ID];
if (employeeId) {
mainSheetMap.set(String(employeeId), {
rowIndex: index + 2,
email: String(row[COL_EMAIL_EMPLOYEE]).trim().toLowerCase(),
timestamp: row[COL_CONFIRMATION_TIMESTAMP]
});
}
});
const statusUpdates = [];
for (let i = 0; i &lt; formResponsesData.length; i++) {
const responseRow = formResponsesData[i];
const currentStatus = responseRow[colStatusForm - 1];

const finalStatuses = [&#39;Confirmat&#39;, &#39;Eroare: Emailul nu corespunde&#39;, &#39;Eroare: ID negăsit&#39;,
&#39;Procesat Anterior&#39;];
if (finalStatuses.includes(currentStatus)) {
statusUpdates.push([currentStatus]);
continue;
}
const emailFromForm = responseRow[1] ? String(responseRow[1]).trim().toLowerCase() :
null;
const employeeIdFromForm = responseRow[colEmployeeIdForm - 1] ?
String(responseRow[colEmployeeIdForm - 1]) : null;
let finalStatus = &#39;&#39;;
if (employeeIdFromForm &amp;&amp; mainSheetMap.has(employeeIdFromForm)) {
const mainSheetEntry = mainSheetMap.get(employeeIdFromForm);
if (emailFromForm === mainSheetEntry.email) {
if (mainSheetEntry.timestamp) {
finalStatus = &#39;Procesat Anterior&#39;;
} else {
const wantsToPurchase = responseRow[5] ?
String(responseRow[5]).trim().toLowerCase() : &#39;&#39;; // Coloana F în foaia de răspunsuri
if (wantsToPurchase === &#39;da&#39;) {
finalStatus = &#39;Confirmat&#39;;
const mainSheetRowIndex = mainSheetEntry.rowIndex;
mainSheet.getRange(mainSheetRowIndex, 3).setValue(finalStatus); // Actualizează
statusul în foaia principală
const formTimestamp = responseRow[0]; // Coloana A în foaia de răspunsuri
const timestampCell = mainSheet.getRange(mainSheetRowIndex,
COL_CONFIRMATION_TIMESTAMP + 1);
timestampCell.setValue(formTimestamp);
timestampCell.setNumberFormat(&quot;dd.MM.yyyy HH:mm:ss&quot;);
mainSheetEntry.timestamp = formTimestamp;
} else if (wantsToPurchase === &#39;nu&#39;) {
finalStatus = &#39;Refuzat&#39;;
} else {
finalStatus = &#39;Raspuns Necunoscut&#39;;
}
}
} else {
finalStatus = &#39;Eroare: Emailul nu corespunde&#39;;

}
} else {
finalStatus = &#39;Eroare: ID negăsit&#39;;
}
statusUpdates.push([finalStatus]);
Copy
}
if (statusUpdates.length &gt; 0) {
formResponsesSheet.getRange(2, colStatusForm, statusUpdates.length,
1).setValues(statusUpdates);
}
Logger.log(&quot;Verificarea răspunsurilor a fost finalizată.&quot;);
}
// --- FUNCȚII UTILITARE PENTRU MENIU ȘI TRIGGERE ---
function onOpen() {
SpreadsheetApp.getUi()
.createMenu(&#39;Scripturi Automatizare&#39;)
.addItem(&#39;[Email] Rulează Trimitere Inițială&#39;, &#39;sendInitialEmails&#39;)
.addItem(&#39;[Email] Trimite Reitare&#39;, &#39;sendReminderEmails&#39;)
.addItem(&#39;[Formular] Verifică Răspunsuri Noi&#39;, &#39;processFormResponses&#39;)
.addSeparator()
.addItem(&#39;[CONFIG] Configurează Trigger Zilnic (Emailuri)&#39;, &#39;setupDailyEmailTrigger&#39;)
.addItem(&#39;[CONFIG] Configurează Trigger la Submit (Formular)&#39;,
&#39;setupFormSubmitTrigger&#39;)
.addSeparator()
.addItem(&#39;[ADMIN] Resetează Contor Linii (Email)&#39;, &#39;resetProcessedRowCount&#39;)
.addItem(&#39;[ADMIN] Șterge Toate Trigger-ele&#39;, &#39;deleteAllProjectTriggers&#39;)
.addToUi();
}
function resetProcessedRowCount() {
const ui = SpreadsheetApp.getUi();
const result = ui.alert(
&#39;Confirmare Resetare Contor&#39;,
&#39;Sunteți sigur că doriți să resetați contorul de linii procesate? Următoarea rulare va începe
de la primul rând.&#39;,
ui.ButtonSet.YES_NO);
if (result == ui.Button.YES) {
PropertiesService.getScriptProperties().deleteProperty(&#39;lastProcessedRow&#39;);
ui.alert(&#39;Contorul a fost resetat.&#39;);
}
}

function setupDailyEmailTrigger() {
const functionToRun = &quot;sendInitialEmails&quot;;
const triggers = ScriptApp.getProjectTriggers();
if (triggers.some(t =&gt; t.getHandlerFunction() === functionToRun &amp;&amp; t.getEventType() ===
ScriptApp.EventType.CLOCK)) {
SpreadsheetApp.getUi().alert(Un trigger zilnic pentru funcția &#39;${functionToRun}&#39; există
deja.);
} else {
ScriptApp.newTrigger(functionToRun)
.timeBased()
.everyDays(1)
.atHour(8)
.inTimezone(Session.getScriptTimeZone())
.create();
SpreadsheetApp.getUi().alert(Trigger zilnic configurat pentru &#39;${functionToRun}&#39;.);
}
}
function setupFormSubmitTrigger() {
const functionToRun = &quot;processFormResponses&quot;;
const triggers = ScriptApp.getProjectTriggers();
if (triggers.some(t =&gt; t.getHandlerFunction() === functionToRun &amp;&amp; t.getEventType() ===
ScriptApp.EventType.ON_FORM_SUBMIT)) {
SpreadsheetApp.getUi().alert(Un trigger la submit pentru funcția &#39;${functionToRun}&#39; există
deja.);
} else {
ScriptApp.newTrigger(functionToRun)
.forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
.onFormSubmit()
.create();
SpreadsheetApp.getUi().alert(Trigger la submit configurat pentru &#39;${functionToRun}&#39;.);
}
}
function deleteAllProjectTriggers() {
const ui = SpreadsheetApp.getUi();
const result = ui.alert(
&#39;Confirmare Ștergere&#39;,
&#39;Sunteți sigur că doriți să ștergeți TOATE trigger-ele automate asociate acestui proiect?&#39;,
ui.ButtonSet.YES_NO);
if (result == ui.Button.YES) {
const triggers = ScriptApp.getProjectTriggers();
if (triggers.length &gt; 0) {
triggers.forEach(trigger =&gt; ScriptApp.deleteTrigger(trigger));
ui.alert(&#39;Toate trigger-ele au fost șterse.&#39;);

} else {
ui.alert(&#39;Nu există trigger-e de șters.&#39;);
}
}
}
