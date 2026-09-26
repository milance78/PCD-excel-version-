# PCD Excel Version

Excel/VBA replika aplikacije `interventions-pcd`.

## Uloga ovog repozitorijuma

- `interventions-pcd` je referentna implementacija i izvor poslovnih pravila.
- Ovaj repozitorijum je nova Excel/VBA implementacija.
- Korporacijski računar treba da može da preuzme **najnoviju verziju aplikacije** iz ovog repozitorijuma dok je razvoj aktivan.
- Preuzeta Excel aplikacija mora nakon toga da radi lokalno, bez Firebase-a, Node-a, React-a ili drugih serverskih komponenti.

## Plan

1. Napraviti Excel `.xlsm` aplikaciju.
2. Replikovati ključne funkcije PCD Tool-a, počev od Magic Import-a.
3. Napraviti kontrolu verzije/update mehanizam preko GitHub-a koji poštuje korporacijska ograničenja.
4. Svaka objavljena verzija dobija eksplicitnu verziju i checksum.
5. Nakon završetka razvoja, aplikacija može biti isporučena kao potpuno lokalna/offline verzija.

## Referentni projekat

`milance78/interventions-pcd`

Ovaj projekat se **ne menja** kao deo razvoja Excel verzije.
