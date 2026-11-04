/**
 * Yasmina & Mario — RSVP backend
 * Bound to Google Sheet owned by yasminaenmario@gmail.com
 *
 * Deployment: Extensions > Apps Script (inside the RSVP spreadsheet);
 * paste this complete file into Code.gs; Deploy > New deployment > Web app
 * Execute as: Me; Who has access: Anyone.
 *
 * Do NOT publish the spreadsheet or share private guest responses.
 */
const SPREADSHEET_ID = '1JbD5CiKWCqY1Cug1NbKdt9ozEotfx--ooj-NRNhkqQw';
const INVITES_TAB = 'Uitnodigingen';
const RESPONSES_TAB = 'Reacties';

function doGet(e) {
  // Invitation type only affects explanatory text, never guest authorization.
  // Both URLs save into the same private spreadsheet ID.
  const ceremonyVariant = !!(e && e.parameter && e.parameter.ceremony === 'yes');
  return HtmlService.createHtmlOutput(RSVP_HTML_(ceremonyVariant))
    .setTitle('Yasmina & Mario — RSVP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function submitRSVP(p) {
  if (!p || typeof p !== 'object') throw new Error('Ongeldige aanmelding.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Het is momenteel druk. Probeer opnieuw.');
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const invites = ss.getSheetByName(INVITES_TAB);
    const responses = ss.getSheetByName(RESPONSES_TAB);
    if (!invites || !responses) throw new Error('De gastenlijst is nog niet ingericht.');
    const code = String(p.code || '').trim().toUpperCase();
    const name = clean_(p.name, 150);
    const email = clean_(p.email, 200);
    if (!/^[A-Z0-9][A-Z0-9-]{3,29}$/.test(code)) throw new Error('Vul een geldige uitnodigingscode in.');
    if (name.length < 2 || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Controleer naam en e-mailadres.');
    if (p.attendance !== 'Ja' && p.attendance !== 'Nee') throw new Error('Kies of je aanwezig bent.');
    const values = invites.getDataRange().getDisplayValues();
    const invitation = values.slice(1).find(r => String(r[0]).trim().toUpperCase() === code);
    if (!invitation) throw new Error('Deze uitnodigingscode is niet gevonden. Neem contact met ons op.');
    const atChurch = String(invitation[5]).trim().toLowerCase() === 'ja';
    const atParty = String(invitation[6]).trim().toLowerCase() === 'ja';
    const maxAdults = Math.max(1, Number(invitation[7]) || 1);
    const plusAllowed = String(invitation[8]).trim().toLowerCase() === 'ja';
    const attending = p.attendance === 'Ja';
    const adults = attending ? Number(p.adults) : 0;
    const children = attending ? Number(p.children) : 0;
    if (!Number.isInteger(adults) || !Number.isInteger(children) || adults < (attending ? 1 : 0) || adults > maxAdults || children < 0 || children > 20) {
      throw new Error('Controleer het aantal gasten. Het maximum volwassenen staat op jullie uitnodiging.');
    }
    const plusOneUsed = attending && p.plusOne === true;
    if (plusOneUsed && !plusAllowed) throw new Error('Een extra gast is niet toegestaan voor deze uitnodiging.');
    const plusOneName = plusOneUsed ? clean_(p.plusOneName, 150) : '';
    if (plusOneUsed && !plusOneName) throw new Error('Vul de naam van je extra gast in.');
    const names = clean_(p.adultNames, 600);
    const childNames = clean_(p.childNames, 600);
    if (attending && !names) throw new Error('Vul de namen van de volwassen gasten in.');
    if (attending && children > 0 && !childNames) throw new Error('Vul de namen van de kinderen in.');
    const previous = responses.getDataRange().getDisplayValues();
    const index = previous.findIndex((r,i)=>i>0 && String(r[2]).trim().toUpperCase()===code);
    const existingId = index > 0 ? previous[index][0] : Utilities.getUuid();
    const now = new Date();
    const row = [
      existingId, index > 0 ? previous[index][1] : now, code, name, email,
      p.attendance, attending && atChurch ? 'Ja' : 'Nee', attending && atParty ? 'Ja' : 'Nee',
      adults, children, names, childNames,
      plusOneUsed ? 'Ja' : 'Nee', plusOneName,
      clean_(p.diet, 1000), clean_(p.accessibility, 1000),
      clean_(p.message, 1500), now,
      index > 0 ? 'Gewijzigd' : 'Nieuw'
    ];
    if (index > 0) responses.getRange(index + 1, 1, 1, row.length).setValues([row]);
    else responses.appendRow(row);
    // Confirmation mail follows only successful Sheet write.
    try {
      MailApp.sendEmail({
        to: email,
        subject: 'RSVP bevestigd — Yasmina & Mario',
        body: 'Beste ' + name + ',\n\nWe hebben jullie RSVP ontvangen (' + p.attendance + ').\n\nBewaar deze bevestiging. Wil je iets wijzigen? Gebruik dezelfde uitnodigingscode.\n\nLiefs,\nYasmina & Mario'
      });
    } catch(err) { console.error('Mail confirmation could not be sent: ' + err); }
    return {ok:true, updated:index>0, message:'Bedankt! Je RSVP is opgeslagen.'};
  } finally {lock.releaseLock();}
}
function clean_(x,max) {
  let s = String(x == null ? '' : x).trim().slice(0,max);
  // Prevent spreadsheet formula injection when user content is stored.
  return /^[=+@\-]/.test(s) ? "'" + s : s;
}
function RSVP_HTML_(ceremonyVariant) {
  return `<!doctype html><html lang="nl"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;500;600&family=DM+Sans:wght@400;500;600&display=swap" rel="stylesheet">
  <style>*,*:before,*:after{box-sizing:border-box}body{margin:0;color:#493d33;background:#ede7dd;font:14px/1.65 "DM Sans",sans-serif}main{max-width:640px;padding:14px 18px 40px;margin:auto}h1{font:normal 65px/1 "Cormorant Garamond",serif;text-align:center;margin:10px 0}p.intro{text-align:center;font:23px "Cormorant Garamond",serif}label{display:block;margin:16px 0 6px;font-weight:600;letter-spacing:.06em;font-size:11px;text-transform:uppercase}input,select,textarea{width:100%;background:#fffdf7;border:1px solid #b9a185;color:#493d33;border-radius:0;padding:12px;font:inherit}textarea{min-height:70px;resize:vertical}button{width:100%;padding:15px;background:#493d33;color:white;border:0;letter-spacing:.1em;text-transform:uppercase;margin-top:20px;cursor:pointer}button:disabled{opacity:.5;cursor:wait}.tip{font-size:12px;color:#796555}.kicker{text-align:center;letter-spacing:.25em;font-size:10px;text-transform:uppercase}#out{margin:18px 0;font-weight:600}#out.error{color:#943d32}#out.success{color:#3b6846}[hidden]{display:none!important}</style></head><body><main><div class="kicker">YASMINA & MARIO · 30.10.2027</div><h1>Kindly Reply</h1><p class="intro">We hopen jullie te zien op onze grote dag.</p><p class="tip" style="text-align:center">${ceremonyVariant ? 'Deze uitnodiging bevat ook de kerkelijke ceremonie. Het systeem controleert jullie uitnodigingscode; de kerk is uitsluitend voor persoonlijk uitgenodigde gasten.' : 'Uitnodiging voor het feest. Alleen gasten die ook voor de kerk zijn uitgenodigd ontvangen daarvoor de aanvullende informatie.'}</p>
  <form id="f"><label for="code">Persoonlijke uitnodigingscode *</label><input name="code" id="code" required maxlength="30" autocomplete="off" placeholder="YM-0001"><p class="tip">Gebruik de code van jullie uitnodiging. Met dezelfde code kun je later je antwoord wijzigen.</p>
  <label for="name">Naam contactpersoon *</label><input id="name" name="name" required maxlength="150" autocomplete="name">
  <label for="email">E-mailadres *</label><input id="email" name="email" type="email" required maxlength="200" autocomplete="email">
  <label for="attendance">Zijn jullie erbij? *</label><select id="attendance" name="attendance" required><option value="">Kies een antwoord</option><option value="Ja">Met liefde aanwezig</option><option value="Nee">Helaas verhinderd</option></select>
  <div id="more" hidden><label for="adults">Aantal volwassenen inclusief jezelf *</label><input id="adults" name="adults" type="number" value="1" min="1" max="20">
  <label for="adultNames">Namen van de volwassen gasten *</label><textarea id="adultNames" name="adultNames" placeholder="Voor- en achternamen"></textarea>
  <label for="children">Aantal kinderen (alle kinderen welkom)</label><input id="children" name="children" type="number" value="0" min="0" max="20">
  <label for="childNames">Namen van de kinderen</label><textarea id="childNames" name="childNames"></textarea>
  <label for="plusOne">Staat op jullie uitnodiging dat een extra gast is toegestaan?</label><select id="plusOne" name="plusOne"><option value="Nee">Nee</option><option value="Ja">Ja</option></select>
  <div id="plusFields" hidden><label for="plusOneName">Naam van extra gast *</label><input id="plusOneName" name="plusOneName"></div>
  <p class="tip">Extra volwassen gasten zijn uitsluitend toegestaan wanneer dit op jullie persoonlijke uitnodiging staat; het systeem controleert deze toestemming.</p>
  <label for="diet">Dieetwensen of allergieën</label><textarea id="diet" name="diet"></textarea><label for="accessibility">Toegankelijkheidswensen</label><textarea id="accessibility" name="accessibility"></textarea></div>
  <label for="message">Een berichtje voor ons</label><textarea name="message" id="message"></textarea>
  <button id="send" type="submit">Verstuur RSVP</button><div id="out" role="status" aria-live="polite"></div>
  <p class="tip">Jullie antwoorden worden privé bewaard in onze gastenadministratie. Deel de persoonlijke code niet met anderen.</p></form></main>
  <script>
  const f=document.getElementById('f'),more=document.getElementById('more'),out=document.getElementById('out');
  document.getElementById('attendance').onchange=e=>{more.hidden=e.target.value!=='Ja'};
  document.getElementById('plusOne').onchange=e=>{document.getElementById('plusFields').hidden=e.target.value!=='Ja'};
  f.onsubmit=e=>{e.preventDefault();const btn=document.getElementById('send');if(!f.reportValidity())return;btn.disabled=true;out.textContent='We verwerken jullie antwoord…';out.className='';
  const d=Object.fromEntries(new FormData(f));d.plusOne=d.plusOne==='Ja';
  google.script.run.withSuccessHandler(result=>{btn.disabled=false;out.className='success';out.textContent=result.message;f.reset();more.hidden=true;document.getElementById('plusFields').hidden=true})
  .withFailureHandler(err=>{btn.disabled=false;out.className='error';out.textContent=err && err.message ? err.message : 'Er is iets misgegaan. Probeer opnieuw.'}).submitRSVP(d)};
  <\/script></body></html>`;
}
