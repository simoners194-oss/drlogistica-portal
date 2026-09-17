# Costo orario hub iMile

Quanto costa un'ora di lavoro all'hub iMile: le **ore riconosciute dal cliente**
sul prospetto presenze contro gli **stipendi effettivamente pagati**, mese per mese.

`estrai_costo_orario.py` legge i file sorgente e produce un workbook di analisi.

## Perché non è dentro il portale

I prospetti presenze arrivano dal cliente come allegato e gli stipendi da paghe:
nessuno dei due passa da SharePoint, quindi non c'è una lista da cui il portale
possa leggerli. È uno script da lanciare a mano quando serve, come
`mezzi-seed/genera_seed.py`.

## Uso

```bash
pip install openpyxl          # solo la prima volta

python estrai_costo_orario.py \
  --presenze "Ore IMILE" \
  --stipendi "STIPENDI GENNAIO-AGOSTO.xlsx" \
  --out "costo-orario-imile.xlsx"
```

- `--presenze` è una **cartella**: ogni `.xlsx` al suo interno con il nome del mese
  in italiano nel filename viene letto (`PRESENZE IMILE MARZO.xlsx`, …).
  Aggiungere un mese = copiare il file nella cartella e rilanciare.
- `--stipendi` è un xlsx con il nome del mese in colonna A e l'importo in colonna B.

Lo script stampa ore e importo di ogni mese letto e il costo orario complessivo:
è il primo controllo che sia andato tutto bene.

## Cosa produce

Quattro fogli, con le formule vive (nessun numero incollato):

| Foglio | Contenuto |
|---|---|
| **Costo orario** | Un mese per riga: ore riconosciute, importo riconosciuto, stipendi, costo €/h, margine. L'unica cella da decidere è **B4**, gialla: il carico aggiuntivo sugli stipendi (0% se la colonna Stipendi è già il costo azienda pieno, altrimenti contributi + TFR + ratei). |
| **Dettaglio mensile** | Giorni aperti, ore ordinarie feriali, domenicali, straordinari, presenze in giorni-uomo, punta operatori/giorno. |
| **Giornaliero** | Tutte le righe dei prospetti presenze, filtro attivo. È la base di verifica: i totali degli altri fogli sono la somma di queste righe. |
| **Note e fonti** | Metodo, tariffe, perimetro, punti aperti. |

Testo blu = dato letto dai sorgenti, nero = formula. Il blu non si corregge a mano:
si rigenera rilanciando lo script.

## Tracciato del prospetto presenze

Dal file del cliente si usa **solo il foglio `MESE 2026`**, dalla riga 3 in giù,
una riga per giorno di calendario:

| Col | Contenuto |
|---|---|
| A | giorno del mese |
| B | giorno della settimana (`Sun` = domenica) |
| C | data |
| D | numero operatori |
| E | ore totali |
| F | ore ordinarie |
| G | ore straordinarie |
| H · I · J | costo ordinario · straordinario · totale |

Le giornate a zero ore vengono saltate. La riga `Total` già presente nel file del
cliente **non viene usata**: i totali sono risommati dai singoli giorni, così il
numero è verificato e non copiato. Se un domani iMile cambia il tracciato, le
costanti da toccare sono in cima allo script (`FOGLIO_MESE` e gli indici `C_*`).

Tariffe riconosciute, dal prospetto stesso: **17,50 €/h** ordinarie (07:00–21:59),
**20,50 €/h** notturne (22:00–06:00), **27,50 €/h** la domenica.

## Limiti da tenere presenti

- Il file stipendi è un importo unico per mese, **senza indicazione della sede**:
  l'analisi assume che sia esattamente e soltanto il personale dell'hub iMile.
- Se quell'importo è la sola retribuzione e non il costo azienda, mancano
  contributi, TFR e ratei: serve la cella B4.
- Il margine calcolato copre **il solo costo del personale**. Attrezzature,
  materiali, struttura e oneri di sede non ci sono.

## Dati

I file sorgente e il workbook prodotto contengono le buste paga: **restano in
locale**, non vanno committati. In questa cartella c'è solo lo script.
