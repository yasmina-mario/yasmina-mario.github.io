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
  if (e && e.parameter && e.parameter.music === 'yes') return HtmlService.createHtmlOutput(MUSIC_HTML_()).setTitle('Yasmina & Mario — Muziekaanvragen').addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
    const babiesWithoutSeat = attending ? Number(p.babiesWithoutSeat) : 0;
    if (!Number.isInteger(adults) || !Number.isInteger(children) || adults < (attending ? 1 : 0) || adults > maxAdults || children < 0 || children > 20) {
      throw new Error('Controleer het aantal gasten. Het maximum volwassenen staat op jullie uitnodiging.');
    }
    if (!Number.isInteger(babiesWithoutSeat) || babiesWithoutSeat < 0 || babiesWithoutSeat > children) {
      throw new Error('Het aantal baby’s zonder eigen stoel moet tussen 0 en het totale aantal kinderen liggen.');
    }
    const seatsNeeded = adults + children - babiesWithoutSeat;
    const plusOneUsed = attending && p.plusOne === true;
    if (plusOneUsed && !plusAllowed) throw new Error('Een extra gast is niet toegestaan voor deze uitnodiging.');
    const plusOneName = plusOneUsed ? clean_(p.plusOneName, 150) : '';
    if (plusOneUsed && !plusOneName) throw new Error('Vul de naam van je extra gast in.');
    if (plusOneUsed && adults < 2) throw new Error('Een plus-one telt mee als volwassene. Kies minimaal 2 volwassenen.');
    const names = clean_(p.adultNames, 600);
    const childNames = clean_(p.childNames, 600);
    if (attending && !names) throw new Error('Vul de namen van de volwassen gasten in.');
    // Each adult is entered on a separate line. A plus-one is entered separately,
    // but MUST be included in the total number of adults (maxAdults includes plus-one).
    const namedAdults = attending ? names.split(/\r?\n/).map(n => n.trim()).filter(Boolean) : [];
    if (attending && namedAdults.length !== adults - (plusOneUsed ? 1 : 0)) {
      throw new Error('Het aantal namen klopt niet. Vul één volwassen naam per regel in; tel de plus-one mee bij het aantal volwassenen, maar vul diens naam alleen in bij plus-one.');
    }
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
      index > 0 ? 'Gewijzigd' : 'Nieuw', babiesWithoutSeat, seatsNeeded
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
  <div id="more" hidden><label for="adults">Totaal aantal volwassenen (inclusief eventuele plus-one) *</label><input id="adults" name="adults" type="number" value="1" min="1" max="20">
  <label for="adultNames">Namen volwassenen (zonder eventuele plus-one) *</label><textarea id="adultNames" name="adultNames" placeholder="Eén naam per regel; plus-one apart invullen"></textarea>
  <label for="children">Aantal kinderen (alle kinderen welkom)</label><input id="children" name="children" type="number" value="0" min="0" max="20">
  <label for="babiesWithoutSeat">Waarvan baby's die op schoot zitten (geen eigen stoel)</label><input id="babiesWithoutSeat" name="babiesWithoutSeat" type="number" value="0" min="0" max="20"><p class="tip">Geef het aantal kinderen inclusief baby's op. Trek alleen de baby's zonder eigen stoel af van het aantal benodigde stoelen. Een kinderstoel telt wel als een stoel.</p>
  <label for="childNames">Namen van de kinderen (inclusief baby's)</label><textarea id="childNames" name="childNames"></textarea>
  <label for="plusOne">Staat op jullie uitnodiging dat een extra gast is toegestaan?</label><select id="plusOne" name="plusOne"><option value="Nee">Nee</option><option value="Ja">Ja</option></select>
  <div id="plusFields" hidden><label for="plusOneName">Naam van extra gast *</label><input id="plusOneName" name="plusOneName"></div>
  <p class="tip">Een plus-one telt mee in het totaal aantal volwassenen en in het maximale aantal op jullie uitnodiging. Geef diens naam alleen op in het plus-one-veld, niet dubbel in de namenlijst. Het systeem controleert de toestemming.</p>
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
  if(d.attendance==='Ja' && Number(d.babiesWithoutSeat)>Number(d.children)){btn.disabled=false;out.className='error';out.textContent='Het aantal baby’s zonder stoel mag niet hoger zijn dan het aantal kinderen.';return;}
  google.script.run.withSuccessHandler(result=>{btn.disabled=false;out.className='success';out.textContent=result.message;f.reset();more.hidden=true;document.getElementById('plusFields').hidden=true})
  .withFailureHandler(err=>{btn.disabled=false;out.className='error';out.textContent=err && err.message ? err.message : 'Er is iets misgegaan. Probeer opnieuw.'}).submitRSVP(d)};
  <\/script></body></html>`;
}

/** Music requests are private, invitation-code validated and never alter RSVP rows. */
function submitMusicRequest(p) {
  if (!p || typeof p !== 'object') throw new Error('Ongeldig verzoek.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Probeer het over een moment opnieuw.');
  try {
    const code = String(p.code || '').trim().toUpperCase();
    const guest = clean_(p.guest, 120), artist = clean_(p.artist, 120), song = clean_(p.song, 160), note = clean_(p.note, 300);
    if (!/^[A-Z0-9][A-Z0-9-]{3,29}$/.test(code)) throw new Error('Controleer de uitnodigingscode.');
    if (guest.length < 2 || artist.length < 2 || song.length < 2) throw new Error('Vul je naam, de artiest en het nummer in.');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const invites = ss.getSheetByName(INVITES_TAB);
    const music = ss.getSheetByName('Muziekaanvragen');
    if (!invites || !music) throw new Error('Muziekaanvragen zijn nog niet ingericht.');
    const invitation = invites.getDataRange().getDisplayValues().slice(1).find(r => String(r[0]).trim().toUpperCase() === code);
    if (!invitation || String(invitation[6]).trim().toLowerCase() !== 'ja') throw new Error('Deze code is niet geldig voor het feest.');
    // Limit accidental flooding: at most 5 requests per invitation code.
    const existing = music.getLastRow() > 1 ? music.getRange(2,2,music.getLastRow()-1,1).getDisplayValues().filter(r=>r[0]===code).length : 0;
    if (existing >= 5) throw new Error('Er zijn al vijf muziekaanvragen met deze uitnodigingscode ontvangen.');
    music.appendRow([new Date(),code,guest,artist,song,note,'Nieuw']);
    return {message:'Dankjewel! Jullie muziekaanvraag is ontvangen. We bekijken alle verzoeken voor de playlist.'};
  } finally {lock.releaseLock();}
}
function MUSIC_HTML_() {
  return `<!doctype html><html lang="nl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;500;600&family=DM+Sans:wght@400;500;600&display=swap" rel="stylesheet">
  <style>*{box-sizing:border-box}body{margin:0;background:#f8f3e9;color:#4d4135;font:14px/1.65 "DM Sans",sans-serif}main{max-width:540px;margin:0 auto;padding:28px 25px 45px;border:1px solid #c7ae7b;background:#fffcf6}h1{text-align:center;font:italic 48px/1.1 "Cormorant Garamond",serif;margin:10px 0;color:#92794f}.k{text-align:center;font-size:10px;letter-spacing:.28em;color:#9e8150}p{text-align:center}label{display:block;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.08em;margin:15px 0 5px}input,textarea{width:100%;border:1px solid #c6b28e;background:white;padding:12px;font:inherit;color:#493d33}textarea{min-height:70px}button{width:100%;margin-top:20px;padding:15px;background:#4b463b;color:white;border:1px solid #b79d67;text-transform:uppercase;letter-spacing:.12em;cursor:pointer}button:disabled{opacity:.6}#out{margin-top:15px;text-align:center;font-weight:600}.tip{font-size:12px;color:#7a6c5c}</style></head><body><main><div class="k">YASMINA & MARIO · 30.10.2027</div><h1>Een liedje voor ons</h1><p>Welk nummer mag volgens jou niet ontbreken op ons feest?</p><form id="music"><label for="code">Uitnodigingscode *</label><input id="code" name="code" required maxlength="30" placeholder="YM-0001" autocomplete="off"><label for="guest">Jouw naam *</label><input id="guest" name="guest" required maxlength="120"><label for="artist">Artiest *</label><input id="artist" name="artist" required maxlength="120"><label for="song">Titel van het nummer *</label><input id="song" name="song" required maxlength="160"><label for="note">Opmerking (optioneel)</label><textarea id="note" name="note" maxlength="300"></textarea><button id="send" type="submit">Verstuur muziekaanvraag</button><div id="out" role="status" aria-live="polite"></div><p class="tip">Maximaal vijf verzoeken per uitnodiging. We bekijken alle suggesties; aanvragen zijn geen garantie dat een nummer wordt gedraaid.</p></form></main><script>document.getElementById('music').onsubmit=function(e){e.preventDefault();const b=document.getElementById('send'),out=document.getElementById('out');b.disabled=true;out.textContent='Bezig met verzenden…';google.script.run.withSuccessHandler(function(r){b.disabled=false;out.textContent=r.message;document.getElementById('music').reset()}).withFailureHandler(function(e){b.disabled=false;out.textContent=e.message||'Verzenden mislukt.'}).submitMusicRequest(Object.fromEntries(new FormData(this)))};<\/script></body></html>`;
}
