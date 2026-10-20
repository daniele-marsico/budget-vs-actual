# Come replicarlo da un cliente

Guida pratica per usare il report su dati veri. Non serve saper programmare: si lavora con file Excel e tre comandi.
La prima volta richiede circa mezza giornata, dal mese successivo pochi minuti.

> **Riservatezza.** I dati del cliente restano sul tuo PC o sul suo. La cartella `data/` è esclusa da GitHub: non caricare mai file del cliente nel repository.

## 1. Preparare il PC (una volta sola)

1. Installa **Python** da [python.org/downloads](https://www.python.org/downloads/). Su Windows spunta **"Add python.exe to PATH"** nella prima schermata. L'installazione per il solo utente non richiede diritti di amministratore.
2. Scarica il progetto: in alto su questa pagina, pulsante verde **Code → Download ZIP**, poi estrai la cartella (per esempio in `Documenti\budget-vs-actual`).
3. Apri la cartella in Esplora file, scrivi `cmd` nella barra dell'indirizzo e premi Invio: si apre il Prompt dei comandi già posizionato nella cartella.
4. Installa le librerie (una volta sola):

   ```
   pip install -r requirements.txt
   ```

## 2. Prova con i dati di esempio

```
python generate_sample_data.py
python variance_report.py --month 2026-04
```

Apri `output\variance_report_2026-04.xlsx`: deve corrispondere all'anteprima nel README. Se funziona, il PC è pronto.

## 3. Dati da chiedere al cliente

Svuota la cartella `data\` e mettici questi file (Excel o CSV, va bene anche il formato italiano con `;` e la virgola decimale):

| File | Cosa contiene | Da dove si prende in Business Central |
|---|---|---|
| `actuals.xlsx` | Movimenti contabili dell'esercizio | **Movimenti C/G**, filtro sulle date dell'esercizio → *Apri in Excel* |
| `budget.xlsx` | Budget mensile per conto | **Movimenti budget C/G**, filtro sul nome del budget dell'anno → *Apri in Excel* |
| `accounts.xlsx` *(facoltativo)* | Nomi dei conti | **Piano dei conti** → *Apri in Excel* |

Colonne che lo script cerca (riconosce i nomi in italiano e in inglese):

- **actuals**: Data di registrazione, Nr. conto C/G, Importo, Codice dimensione globale 1 (centro di costo, facoltativo)
- **budget**: Data, Nr. conto C/G, Importo, Codice dimensione globale 1 (facoltativo)
- **accounts**: Nr., Nome

Se una colonna ha un nome diverso, lo script si ferma e scrive quale manca e quali ha trovato: basta rinominarla nel file.
Con un gestionale diverso da Business Central vale lo stesso: serve un export con queste colonne.

**Segni.** Come in Business Central, ricavi negativi e costi positivi. Lo script capisce da solo quali conti sono ricavi (saldo negativo) e nel report li mostra in positivo.

## 4. Lanciare il report

```
python variance_report.py --month 2026-09
```

- `--month` è il mese di chiusura; il progressivo parte dal primo mese presente nel budget.
- `--threshold 0.05` evidenzia gli scostamenti oltre il 5% (di default il 10%).

## 5. Controlli prima di consegnare

Questi controlli sono il tuo lavoro da controller: lo script fa i conti, tu garantisci che siano giusti.

1. **Quadratura con la contabilità.** Il consuntivo progressivo di ogni conto deve coincidere con il **Bilancio di verifica** di Business Central per lo stesso periodo.
2. **Ricavi e costi.** Controlla che ogni conto sia classificato bene (colonna *Account Type*). Un conto di ricavo con un saldo anomalo può finire tra i costi.
3. **Conti senza budget.** Compaiono con budget 0 e scostamento pari al consuntivo: chiedi al cliente se è voluto.
4. **Centro di costo "-".** Sono movimenti senza dimensione: di solito vanno segnalati all'amministrazione.

## 6. Cosa consegnare

- Il file Excel, con una call di 15–20 minuti per leggere insieme i 3–5 scostamenti principali (foglio *Top deviations*).
- Ogni mese: nuovo export da BC e rilancio del comando, oppure automatico con la GitHub Action su un repository **privato** del cliente.

## 7. Limiti da dire al cliente

- Lavora su file esportati, non con un collegamento diretto al gestionale (estensione possibile con l'API di Business Central).
- Il budget dev'essere mensile e sugli stessi conti della contabilità.
