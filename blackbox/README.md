# PCD Black Box

Ovaj folder je jednostavan most između korporativnog računara na kome se PCD Excel testira i ChatGPT sesije na drugom računaru.

## Kako se koristi

1. Na korporativnom računaru otvori folder `blackbox/inbox` na GitHub-u.
2. Klikni **Add file → Upload files**.
3. Prevuci screenshot greške (PNG/JPG). Poželjno ime: `DEV-29-syntax-error.png`, ali nije obavezno.
4. Klikni **Commit changes** direktno na `main`.
5. U chatu napiši samo: **"Pogledaj crnu kutiju."**

To je sve. Ne moraš slati screenshot mejlom niti prepisivati tekst greške.

## Pravila

- `inbox/` služi za nove screenshotove i male tekstualne logove.
- Ne stavljaj XLSM buildove u crnu kutiju; oni ostaju u `dist/`.
- Ne menjaj postojeće screenshotove. Za novi test dodaj novi fajl.
- Ne briši stare fajlove dok dijagnostika traje.
- Ne stavljaj poverljive podatke, lozinke, tokene ili podatke korisnika/klijenata na GitHub.
- Crna kutija nije deo build/update procesa i workflow je ne koristi.

## Imenovanje

Preporuka:

`DEV-<broj>-<kratak-opis>.<ext>`

Primeri:

`DEV-29-syntax-error.png`

`DEV-30-updater-log.txt`

Ako ima više slika istog testa:

`DEV-29-syntax-error-1.png`
`DEV-29-syntax-error-2.png`

## Važno

Cilj je da Black Box ostane namerno jednostavan: GitHub je samo mesto za prenos dokaza sa test računara. Nema dodatnog programa, skripte, VBA koda ni GitHub Action-a.
