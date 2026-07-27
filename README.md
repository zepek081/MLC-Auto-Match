# MLC Auto Match

Automazione Playwright per The MLC Matching Tool: legge un catalogo (Excel/CSV)
e prova il match automatico per ISRC, poi per Track Title + Publisher ("LOO"),
con fallback su Track Title + Writer Name.

## Setup

```
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
npx playwright install chromium
```

Credenziali via variabili d'ambiente (mai hardcoded nello script o nel file):

```
export MLC_EMAIL="tua_email@dominio.com"
export MLC_PASSWORD="tua_password"
```

Su Windows PowerShell:

```
$env:MLC_EMAIL="tua_email@dominio.com"
$env:MLC_PASSWORD="tua_password"
```

Il codice 2FA (OTP) arriva via email/SMS: lo script te lo chiede a terminale
solo se il sito lo richiede davvero (non succede a ogni login - a volte il
dispositivo/sessione e' gia' riconosciuto e si salta dritti alla Summary).
Non va salvato da nessuna parte.

## Come funziona il matching (due stage separati)

Il tool MLC non fa un match in un solo passaggio: sono due ricerche in
sequenza sulla stessa riga di catalogo.

**Stage 1 - trova la registrazione (recording)**: cerca per ISRC. Se trova
piu' gruppi per lo stesso ISRC (es. varianti "Original Mix" arrivate da DSP
diversi), li seleziona TUTTI e li conferma insieme - non e' un caso ambiguo,
e' normale. Se i risultati sono tanti, puo' comparire un pulsante "Load
More" anche piu' volte di seguito: lo script lo clicca finche' non
scompare, prima di selezionare tutti i gruppi (altrimenti ne selezionerebbe
solo una parte). Se l'ISRC non trova nulla, la riga si ferma qui
(`no_match_recording`). Se il gruppo risulta gia' "Submitted"/"Accepted"/
"Rejected" da una sessione precedente, la riga e' `already_submitted` e si
passa all'ISRC successivo senza toccare nulla.

**Stage 2 - trova la tua opera (work) gia' registrata** da abbinare alla
registrazione appena confermata: cerca per Titolo + un secondo criterio. Di
default il secondo criterio e' **Publisher Name con il valore fisso "LOO"**
(funziona per "LOOSE CLUB EDITION" a prescindere dal publisher esatto in
riga). Se questa ricerca non trova nulla o e' ambigua (piu' di un risultato)
e il writer e' disponibile in input, si ritenta con Titolo + **Writer Name**
(cognome dell'autore, colonna `Surname` nel master sheet): utile sui titoli
generici tipo "System" o "Contacto" dove il publisher da solo non
discrimina - abbiamo verificato un caso con 28 risultati tutti diversi sotto
lo stesso titolo+publisher.

Nota: la ricerca titolo di MLC non e' per frase esatta di default - un
titolo di due parole come "Real Life" puo' restituire decine di risultati
per singola parola ("Real Faces", "Still Life", ...). Se un'opera non
compare ne' con Publisher+LOO ne' con Writer, molto probabilmente non e'
ancora registrata su MLC (`no_match_work`), non e' un problema di ricerca.

## Uso

```
python mlc_auto_match.py catalogo.xlsx --sheet "FEB 25" --output risultati.xlsx
```

Parametri principali:

- `--sheet` nome dello sheet Excel da usare (default: il primo)
- `--isrc-col`, `--title-col` colonne obbligatorie (default: `ISRC Code`,
  `Track Title`, gia' i nomi usati nel master sheet)
- `--writer-col` colonna cognome autore, usata come fallback allo Stage 2 se
  Publisher Name ("LOO") non trova un risultato univoco (default: `Surname`)
- `--publisher-col` non usato per la ricerca (si usa il valore fisso "LOO"),
  tenuto solo per riferimento nel report (default: `Publisher Name`)
- `--output` file Excel con il report finale
- `--headless` esegue senza finestra visibile (sconsigliato al primo utilizzo,
  utile a regime una volta verificato che funziona)
- `--skip-ambigui` non si ferma sulle righe con piu' opere diverse: le segna
  `ambiguous_work` e prosegue. Serve per lotti lunghi da lasciare non
  presidiati, poi si rivedono tutte insieme dal report

Il report viene riscritto **dopo ogni riga**, non solo a fine run: su lotti
lunghi un Ctrl+C o un crash non fa perdere il lavoro gia' fatto. Se
interrompi, il file contiene le righe processate fino a quel momento.

## Stati possibili nel report

Il file Excel di output colora automaticamente ogni riga come nel flusso
manuale: **verde** per `matched` e `already_submitted`, **giallo** per
`no_match_recording` e `no_match_work`, **arancione** per `submit_failed`,
`manual_incomplete` e `error` (tentativi falliti da ritentare - tenerli
distinti dal giallo, che significa "opera non presente a catalogo"). Gli
stati `ambiguous_*` restano senza colore e richiedono la lettura della
colonna `note`.

- `matched` - Stage 1 e Stage 2 completati, match confermato
- `no_match_recording` - ISRC non trovato in MLC, nessuna registrazione
- `already_submitted` - il gruppo recording per questo ISRC risulta gia'
  "Submitted"/"Accepted"/"Rejected" da una sessione precedente: nessuna
  azione necessaria, riga gia' processata in passato
- `no_match_work` - la registrazione e' stata trovata (Stage 1 ok) ma nessuna
  opera nel tuo catalogo corrisponde ne' a Titolo+Publisher("LOO") ne' a
  Titolo+Writer: probabilmente l'opera non e' ancora registrata su MLC, serve
  verifica manuale (es. tramite Works Registration), non e' un errore dello
  script
- `submit_failed` - il match e' stato compilato correttamente ma **MLC ha
  rifiutato l'invio lato server**: il loro backend risponde HTTP 400
  ("Failed to contact recordings API") su
  `POST /current/matching/suggestions` e la loro UI non mostra alcun errore,
  lascia solo il dialog di conferma bloccato. Non e' un problema dello
  script e non si aggira cambiando opera: verificato su ISRC
  `GBLV62419182` (PLAYBOOYZ) sia con la registrazione `PN8C4S` sia con la
  doppia `PN8DH1`, stesso identico errore. Da ritentare piu' tardi o da
  segnalare al supporto MLC
- `ambiguous_recording` - layout inatteso: nessun controllo di selezione
  trovato per l'ISRC (caso raro, i gruppi multipli vengono gia' gestiti
  automaticamente selezionandoli tutti). Se capita su ISRC che rifatti a
  mano funzionano, e' un problema di attesa: vedi punto 18
- `ambiguous_work` - piu' di un'opera trovata con lo stesso titolo/criterio
  anche dopo il fallback su Writer, richiede scelta manuale (lo script si
  ferma e aspetta la tua selezione a schermo prima di continuare, oppure la
  segna e prosegue se hai passato `--skip-ambigui`)
- `manual_incomplete` - dopo una pausa manuale non e' stata trovata ne' la
  finestra di conferma ne' la schermata di successo: quasi sempre significa
  che l'invio e' stato dato senza aver selezionato un'opera. La riga va
  rifatta
- `error` - eccezione imprevista, dettaglio nella colonna `note`

## Comportamento sulle righe ambigue

Con piu' risultati nello Stage 2 lo script distingue due casi:

**Registrazioni doppie della stessa opera** - stesso titolo, stessi autori,
stesso publisher, ma MLC Song Code diversi (caso reale: PLAYBOOYZ, `PN8C4S`
con 7 recordings e `PN8DH1` con 6). Non c'e' una scelta vera da fare: prende
la prima e lo annota nel report. Il confronto ignora di proposito Song Code,
Member's Song ID e numero di recordings/artisti collegati, che sono dati
della singola registrazione e non dell'opera.

**Opere davvero diverse** - cambia il titolo, l'autore o il publisher. Qui lo
script NON sceglie da solo: si ferma, ti chiede di selezionare a mano nel
browser visibile, poi continua alla pressione di invio nel terminale.

## Assunzioni da verificare al primo run

Il codice e' stato costruito incrociando una sessione Playwright Codegen, un
video della sessione di lavoro e una sessione live di test guidata (MCP
Playwright) sull'account reale. Punti da controllare comunque con
`--headless` disattivato, dato che restano assunzioni su un sito che puo'
cambiare:

1. Stage 1 assume che il form recording abbia gia' di default riga 0 =
   Recording Title, riga 1 = Recording ISRC (confermato live).
2. Stage 2: la riga 1 (secondo criterio) esiste di norma gia' di suo, col
   dropdown che mantiene il valore scelto nella ricerca precedente - non va
   aggiunta con "Add Criteria" (confermato live). Il codice clicca il
   dropdown esistente e ricade su "Add Criteria" solo se non lo trova.
3. La selezione dell'opera trovata in Stage 2 usa `#select-link` (un id
   ripetuto su ogni riga risultato, non un pulsante "Select" - confermato
   live), quindi `_count_work_results` conta quell'elemento.
4. I testi di "nessun risultato" sono due formulazioni diverse tra i due
   stage (confermate entrambe live): "couldn't locate any results" (Stage 1)
   e "No results found." (Stage 2). Se cambiano, aggiorna
   `NO_RESULTS_PATTERNS`.
5. Dopo "Match N Group" compare un dialog intermedio "Let's continue
   matching recordings" con un secondo pulsante "Continue" - il click sul
   ruolo "Continue" lo gestisce comunque, ma se il dialog cambia testo puo'
   rompersi.
6. Quando lo Stage 2 non trova nulla (`no_match_work`), tornare a Stage 1
   richiede "Back" + click sull'icona "deselect" (`abandon_stage2`): il
   semplice click su "Matching Tool" lascia il gruppo recording ancora
   selezionato per la ricerca successiva (confermato live, bug corretto).
7. Il campo ISRC/Titolo non si svuota da solo tra una ricerca e l'altra
   (SPA, nessun remount di route): non e' un problema perche' `.fill()`
   sovrascrive comunque il valore precedente.
8. Il pulsante di conferma selezione gruppi si chiama "Match 1 Group" al
   singolare ma "Match N Groups" al plurale da 2 in su - il regex accetta
   entrambi (bug corretto, prima matchava solo il singolare).
9. Con molti risultati in Stage 1 compare un pulsante "Load More" che va
   cliccato ripetutamente (puo' ricomparire piu' volte) prima di selezionare
   tutti i gruppi - non ancora osservato dal vivo con un caso reale, gestito
   per analogia a quanto descritto.
10. Il 2FA non compare a ogni login (bug corretto: lo script restava
    bloccato 30s in attesa del campo "Code" quando il sito saltava dritto
    alla Summary perche' la sessione era gia' riconosciuta) - ora l'attesa
    del campo OTP e' con timeout breve e opzionale.
11. Il banner cookie (widget Cookiebot, elenca una decina di vendor) puo'
    metterci diversi secondi a diventare cliccabile e nel frattempo blocca
    ogni click sottostante (bug corretto: un tentativo di 3s falliva e
    faceva scadere il click su "Login" dopo 30s) - ora si aspetta fino a
    10s sull'id stabile del pulsante "Allow all cookies". Rimosso anche il
    `press("Enter")` dopo la password (ridondante col click esplicito su
    "Login", rischiava di correre in parallelo col banner).
12. **Bug serio corretto**: `search.1.searchTerm` non e' un campo dedicato,
    e' lo stesso slot riusato sia per l'ISRC (Stage 1) sia per il valore
    Publisher/Writer (Stage 2) - stesso test-id in entrambi gli stage. Se
    un errore interrompe un cambio di criterio a meta' in Stage 2, questo
    slot resta "sporco" (es. bloccato su "MLC Song Code") e si trascina
    sulle righe successive invece di resettarsi, senza sollevare
    eccezioni: risultati falsi (`no_match_recording`) silenziosi (osservato
    live su un'intera run dopo un singolo errore in mezzo). Corretto con due
    protezioni: (a) verifica che la riga 1 di Stage 1 mostri davvero
    "Recording ISRC" prima di ogni ricerca, reload completo se non lo
    mostra; (b) il recovery da errore fa un reload completo della pagina,
    non un semplice click sulla sidebar (che e' un no-op SPA). (Un primo
    tentativo di escludere i dropdown disabilitati in
    `_set_second_criteria` e' stato rimosso: rompeva il caso normale,
    lasciando il criterio Stage 2 bloccato su "MLC Song Code" senza valore -
    osservato live, vedi punto 13.)
13. **Gli indici dei gruppi slittano** (verificato live sul DOM): quando un
    gruppo viene selezionato il suo `select-rg-button` SPARISCE (diventa
    l'icona "deselect"), quindi la lista si accorcia a ogni click - da 3
    elementi a 2, poi a 1. Iterare con `nth(i)` fisso seleziona il 1o e il
    3o gruppo saltando il 2o ("Match 2 Groups" invece di 3), poi va in
    timeout su `nth(2)` ormai inesistente. Corretto cliccando sempre il
    PRIMO rimasto, tante volte quanti sono i gruppi.
14. **Il contenitore del criterio non e' cliccabile al centro** (verificato
    live con `elementFromPoint`): il div `search-with-criteria-select-div-*`
    copre tutta la riga (etichetta + casella + toggle) e il suo centro e'
    la CASELLA DI TESTO (`search.1.searchTerm`), non la tendina. Cliccarlo
    metteva il fuoco nel campo senza aprire il menu, e l'attesa
    dell'opzione andava in timeout dopo 30s. Corretto cliccando
    l'ETICHETTA del criterio corrente (`_open_criteria_menu`), come faceva
    la sessione Codegen originale.
15. **Deselezione multipla** (verificato live): dopo "Back" restano
    selezionati TUTTI i gruppi e ci sono altrettante icone "deselect" - un
    click singolo solleva "strict mode violation" di Playwright. Corretto
    ciclando sulla prima icona rimasta finche' non spariscono.
16. Il banner cookie puo' ricomparire **a meta' sessione** dopo un reload,
    bloccando i click con lo stesso timeout di 30s visto al login
    (osservato live) - `_dismiss_cookies` viene ora richiamato dopo ogni
    reload, non solo durante il login.
17. **Il portale puo' fallire in silenzio dopo il Confirm** (diagnosticato
    live leggendo console e traffico di rete): il backend MLC risponde
    HTTP 400 su `POST /current/matching/suggestions` e la loro UI non
    mostra nulla, lascia il dialog "Are you sure...?" aperto coi pulsanti
    spariti. Il match NON viene registrato (verificato: assente dal Match
    History). Ora l'attesa di "Done" e' limitata a 15s e l'esito diventa
    `submit_failed` con nota esplicita, invece di 30s di timeout e un
    generico `error`.
18. **Mai attese a tempo fisso dopo "Search"** (bug osservato live): con
    1200ms fissi, sulla prima ricerca dopo il login i risultati non erano
    ancora renderizzati, si contavano zero pulsanti di selezione e la riga
    finiva in `ambiguous_recording` pur essendo normalissima (ISRC
    `CARH11900303`, rifatto a mano: 2 gruppi regolari). `run_search` ora
    attende la RISPOSTA DI RETE della ricerca (`/search/unmatched-recordings`
    per lo Stage 1, `/search/works/catalog` per lo Stage 2) e poi il render:
    aspettare la risposta nuova evita anche di leggere per sbaglio i
    risultati della ricerca precedente rimasti a schermo.

## Sicurezza

La password non deve mai finire in chat, screenshot o log condivisi: se e'
gia' successo, cambiala prima di usare questo script in produzione.
