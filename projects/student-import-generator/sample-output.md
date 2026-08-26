# Sample Output — Student Import Generator

> Illustrative excerpt with **fictional data**. The tool produces
> `modello-importazione-corsisti.xlsx`, the import file for the course platform.

Each row of the source "student data request" becomes an enriched import row:

| Cognome | Nome | Codice Fiscale | Sesso | Comune Nascita | Prov. Nascita | ID Corso | PIVA Azienda |
|---|---|---|---|---|---|---|---|
| ROSSI | MARIA | RSSMRA80A01H501Z | F | ROMA | RM | 12345 | 01234567890 |
| BIANCHI | LUCA | BNCLCU85C22D612X | M | MILANO | MI | 12345 | 01234567890 |
| FERRARI | ANNA | FRRNNA92B41A944W | F | FIRENZE | FI | 12345 | 09876543210 |
| CONTI | PAOLO | CNTPLA78L03F839E | M | TORINO | TO | 12345 | 09876543210 |

**Sesso, Comune and Prov. are derived automatically** by parsing the Codice
Fiscale (position 9–11 for gender, Belfiore code for the municipality).

Console:

```
>> Read 4 students from Richiesta_dati_allievi.xlsx
>> Parsed tax IDs: 4/4 OK (gender + birth place resolved)
>> Written: modello-importazione-corsisti.xlsx (4 rows, ready for import)
```
