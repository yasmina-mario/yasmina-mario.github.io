# RSVP activeren — Yasmina & Mario

De website is voorbereid op een veilige RSVP met persoonlijke uitnodigingscodes. De oude, kapotte FormSubmit-route is uitgeschakeld. **Er wordt nog niets opgeslagen totdat de web-app is geactiveerd en getest.**

## Benodigd
- Google-account **yasminaenmario@gmail.com**
- [Privé-gastenlijst in Google Sheets](https://docs.google.com/spreadsheets/d/1JbD5CiKWCqY1Cug1NbKdt9ozEotfx--ooj-NRNhkqQw/edit)
- [Code.gs](./Code.gs) uit deze map

## Eenmalige activatie
1. Open de Google Sheet als **yasminaenmario@gmail.com**.
2. Open **Extensies > Apps Script**. Verwijder de standaardinhoud van Code.gs en plak de volledige inhoud van dit [Code.gs-bestand](./Code.gs). Sla op.
3. Open **Implementeren > Nieuwe implementatie > Web-app**.
4. Kies **Uitvoeren als: Ik** en **Toegang: Iedereen**. Geef Google toestemming voor toegang tot de gastenlijst en om bevestigingsmails te verzenden. Controleer daarbij dat je op het juiste account bent ingelogd.
5. Kopieer de web-app-URL met `https://script.google.com/macros/s/...` en **/exec** aan het einde.
6. Open die URL in een privévenster, zonder je Google-account. Zo test je of gasten de RSVP kunnen gebruiken.
7. Vul eerst in het tabblad **Uitnodigingen** een proefcode in. Bijvoorbeeld `YM-TEST01`, een fictieve contactnaam, bij 'Kerk toegestaan?' `Ja`, bij 'Feest toegestaan?' `Ja`, bij 'Max. volwassenen' `2` en bij 'Plus-one toegestaan?' `Ja`. Test ook een ongeldige code.
8. Controleer of de RSVP in **Reacties** verschijnt, of de bevestigingsmail aankomt en of herinzenden met dezelfde code de bestaande reactie bijwerkt. Verwijder testgegevens en de testuitnodiging daarna.
9. Stel in het rootbestand **index.html** de regel `const RSVP_WEBAPP_URL = "";` in op de werkende **/exec**-URL. Vanaf dat moment toont de trouwwebsite het RSVP-formulier in plaats van de tijdelijke melding.

## Beleid
- **Alle kinderen zijn welkom.** Namen en aantal kinderen worden apart vastgelegd.
- **Plus-one alleen indien toegestaan in de uitnodigingslijst.** De backend controleert dit, niet alleen het formulier.
- **Per uitnodigingscode maximaal één actuele reactie.** Nieuwe inzendingen met dezelfde code overschrijven de vorige reactie. Zorg daarom dat alleen de uitgenodigde contactpersoon de code kent.
- **Kerk/feest** wordt bepaald door de uitnodigingslijst, niet door de gast.
- Deel de Google Sheet niet openbaar.
- Bij deployment onder 'Iedereen' kan de publieke webapp worden geopend; de uitnodigingscode beschermt tegen ongeautoriseerd aanmelden, maar is geen sterke authenticatie. Voor zeer gevoelige gastgegevens is een strengere oplossing wenselijk.
- Google Apps Script heeft eigen quota. Controleer verzendlimieten voor bevestigingsmails en test met een klein aantal aanmeldingen.

## Nog in te vullen vóór uitnodigingen
- RSVP-deadline
- Uitnodigingscodes, max. volwassenen, toegestane plus-one per huishouden
- Eventuele definitieve kerk- en feestgegevens
- Namen/foto's en definitieve teksten op de trouwwebsite

**Niets op deze pagina garandeert dat de koppeling al actief is.** Pas na bovenstaande succesvolle test wordt het formulier aangezet.