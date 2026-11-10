# Tilitysanalyysi

Paikallinen, suomenkielinen työkalu Suomisportin Excel-tilitysraporttien analyysiin.

## Käynnistys

Tarvitset Python 3.13:n tai uudemman sekä [uv:n](https://docs.astral.sh/uv/).

```sh
uv sync
uv run tiliraportti-analyysi
```

Avaa terminaalissa näkyvä paikallinen osoite (yleensä http://localhost:8501).
Raahaa `.xlsx`-tiedosto latausalueelle tai valitse se tiedostonvalitsimella.
Voit säilyttää omia raportteja paikallisessa `esimerkki-data`-kansiossa, jota ei lisätä Gitiin.

Työkalu näyttää joukkueittain maksut, palvelumaksut, komissiot, kulujen prosenttiosuudet
ja nettotilityksen. Voit suodattaa joukkueita, tarkastaa kohdistetut lähderivit ja ladata
yhteenvedon sekä lähderivit Excelissä avattavana CSV:nä. CSV sisältää valitun rajauksen.
Tiedosto käsitellään sovelluksen palvelimen muistissa eikä sovellus tallenna sitä levylle.
Paikallisesti ajettaessa palvelin on oma tietokoneesi. Julkisessa palvelussa tiedosto
lähetetään palveluntarjoajan palvelimelle.

## Julkinen Streamlit-palvelu

Käynnistyspolku Streamlit Community Cloudissa on `src/tiliraportti_analyysi/app.py`.
Asenna riippuvuudet lukitustiedoston mukaan (`uv sync --locked`). Pidä Python ja
riippuvuudet ajan tasalla ja tarkista ne säännöllisesti esimerkiksi `pip-audit`-työkalulla.

- Enimmäiskoko on 5 MiB, ZIP-sisällön purettu koko 20 MiB, tiedostojen määrä 128,
  rivimäärä 20 000 ja sarakemäärä 64. Soluja sallitaan yhteensä 200 000 ja
  XML-elementtejä 1 000 000. XML:n DTD:t ja entiteetit estetään.
- Maksuosuuden itseisarvo saa olla enintään 10¹² ja desimaaliesityksen eksponentti
  välillä −6…12. CSV-viennin vaaralliset tekstialut suojataan heittomerkillä.
- `.streamlit/config.toml` pitää CORS- ja XSRF-suojaukset käytössä, estää staattisten
  tiedostojen tarjoilun, rajaa latauksia ja piilottaa virheiden tekniset yksityiskohdat.
  Käynnistä sovellus projektin juuresta, jotta asetukset latautuvat.
- Sovellus ei käytä yhteistä datavälimuistia eikä kirjoita raportteja levylle.
  Palveluntarjoajan lokitus, muistivedokset ja tietojen säilytys on arvioitava erikseen.
- Sovelluksessa ei ole kirjautumista eikä pyyntötiheyden rajoitusta. Julkisessa
  palvelussa käytä HTTPS:ää sekä alustan pääsynhallintaa tarpeen mukaan. Rajoita
  liikennettä ja palvelimen resursseja alustalla tai välityspalvelimessa: tiedostorajat
  eivät estä monista yhtäaikaisista pyynnöistä syntyvää palvelunestohyökkäystä.
- Pidä oikeat raportit poissa Gitistä. `esimerkki-data/` on ohitettu ja testit luovat
  synteettisen raportin muistissa. Pidä salaisuudet alustan salaisuusasetuksissa,
  poissa Gitistä. `.gitignore` ei poista aiemmin tallennettuja tietoja Git-historiasta.

Tarkistus ja suojaukset vähentävät riskejä, mutta eivät takaa haavoittuvuudettomuutta.
Suojaukset perustuvat myös [openpyxlin tietoturvaohjeeseen](https://openpyxl.readthedocs.io/en/stable/index.html#security)
ja [Streamlitin asetusten dokumentaatioon](https://docs.streamlit.io/develop/api-reference/configuration/config.toml).

## Laskenta

- Lähde on `Tilitysrivit`-välilehti, jonka ensimmäinen rivi sisältää otsikot.
  Pakolliset sarakkeet: `Maksun tunniste`, `Joukkueet`, `Maksuosuuden tyyppi`, `Maksuosuus`.
- Summat luetaan vain `Maksuosuus`-sarakkeesta. Tuettuja tyyppejä ovat `Maksu`,
  `Palvelumaksu` ja `Komissio`. Tuntematon tyyppi tai virheellinen summa ilmoitetaan virheenä.
- Rivin oma joukkue on ensisijainen. Puuttuva joukkue haetaan saman maksun `Maksu`-riveiltä.
  Avaimessa käytetään organisaatiota, tilityksen viitenumeroa ja maksun tunnistetta.
  Kohdistus ei riipu rivijärjestyksestä. Epäselvät rivit näytetään `Kohdistamatta`-ryhmässä;
  usean joukkueen yhteisen maksun kuluja ei jaeta automaattisesti.
- Kulut näytetään vastakkaismerkkisinä lähderivin summaan nähden: veloitus on positiivinen
  kulu ja kuluhyvitys negatiivinen. Nettotilitys = maksut − palvelumaksut − komissiot.
- Kukin kuluprosentti = kyseinen kulu / joukkueen maksut × 100. Nollasummalla prosentti
  on tyhjä. Maksujen lukumäärä on joukkueen erillisten maksuavainten määrä.
- Laskenta käyttää desimaalilukuja. Käyttöliittymä näyttää summat kahden desimaalin tarkkuudella.
- `Yhteenveto`-välilehteä ei lasketa mukaan, jotta summat eivät kahdennu.
  Vain `.xlsx` on tuettu; vanha `.xls` pitää tallentaa ensin `.xlsx`-muotoon.

## Tarkistukset

```sh
uv run python -m unittest discover -s tests -v
```

Testit kattavat esimerkkiraportin summat, kulujen kohdistuksen, epäselvät kohdistukset,
hyvitykset, virheelliset tiedostot sekä käyttöliittymän yhteenvedon ja joukkuerajauksen.
