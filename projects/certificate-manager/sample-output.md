# Sample Output — Certificate Manager

> Illustrative excerpt with **fictional data**.

`dividi` — bulk print PDF → one file per person:

```
>> Input: stampa_massiva.pdf (26 pages = 13 people, front+back)
>> attestati_output/
     01_Maria_Rossi.pdf
     02_Luca_Bianchi.pdf
     03_Anna_Ferrari.pdf
     ... (13 files, name read from each certificate's text)
```

`organizza` — one edition sorted into company folders via the Excel directory:

```
>> CORSO10/
     Example Logistics S.r.l./
         01_Maria_Rossi.pdf
     Demo Services S.p.A./
         02_Luca_Bianchi.pdf
     Sample Consulting S.r.l./
         03_Anna_Ferrari.pdf
     _DA_VERIFICARE/
         07_persona_7.pdf     (name not found in directory)

  Matches: 12 exact, 1 fuzzy (~, worth a visual check), 1 to verify
```

`riorganizza` — cross-edition company view:

```
>> ATTESTATI_PER_AZIENDA/
     Example Logistics S.r.l./
         CORSO10/01_Maria_Rossi.pdf
         CORSO11/05_Maria_Rossi.pdf
     Demo Services S.p.A./
         CORSO10/02_Luca_Bianchi.pdf

  Summary: 3 companies across 2 editions — nothing lost
```

Note: fuzzy matches are flagged with `~` in the log so a human can double-check
only the uncertain ones.
