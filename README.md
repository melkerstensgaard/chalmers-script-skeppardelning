
# Sjöfartsbevisprojekt
Syftet med detta program är att underlätta digitiseringen av utbildningsbevis från sjöfartsutbildningar i Chalmers arkiv. På grund av många dokumentformat, handskriven text som leder till dålig OCR kvalité och närvaron av oönskade bilagor är den serien inte lämplig att helt automatisera. Därmed utvecklades detta program för att utföra automatisk extraktion av personnummer i den mån det går. Sedan får en användare gå igenom resterande bevis för hand. 

## Funktioner
- Underliggande skript som detekterar personnummer utifrån OCR text i dokumenten.
- Verktyg för att hangripligen tillföra metadata och klippa dokument.
- Databasfunktion som tillåter att avsluta och återuppta arbetet.

## Skärdump
<img width="1645" height="1028" alt="image" src="https://github.com/user-attachments/assets/bef1cadc-bb71-4bc8-bf97-3439f8cd7aed" />

## Arbetsflöde
1. Skapa eller öppna en arbetsdatabas.
2. Läs in en seriemapp som innehåller en eller flera volymer med PDF-filer.
3. Programmet analyserar PDF-filerna och försöker automatiskt identifiera personnummer.
4. Användaren granskar varje dokument och klassificerar det.
5. Eventuella felaktiga personnummer korrigeras manuellt.
6. Resultatet exporteras till PDF-filer och en indexfil.

## Databasen

Alla granskningsresultat sparas i en SQLite-databas som innehåller:

- Klassificeringar
- Identifierade personnummer
- OCR-resultat
- Granskningsstatus
- Senaste arbetsposition

## Export

Vid export:

- Serie- och volymstrukturen bevaras.
- PDF-filer skapas utifrån användarens klassificeringar.
- En indexfil (`index.xlsx`) skapas.

Indexfilen innehåller:

- Filnamn
- Volym
- Serie
- Personnummer
- Dokumenttyp
- Personnummer bestämt av (människa eller skript)
