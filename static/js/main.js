document.addEventListener('DOMContentLoaded', () => {
    const zonaCaricamento = document.getElementById('upload-zone');
    const inputFile = document.getElementById('file-input');
    const contenitoreProgresso = document.getElementById('progress-container');
    const barraProgresso = document.getElementById('progress-bar-fill');
    const testoStato = document.getElementById('progress-status-text');
    const schedaRisultati = document.getElementById('results-card');
    const tabellaCorpo = document.getElementById('results-table-body');
    const statTrovati = document.getElementById('stat-trovati');
    const statNonTrovati = document.getElementById('stat-non-trovati');
    const bottoneScarica = document.getElementById('btn-download');
    const chkAnchored = document.getElementById('chk-anchored');

    let intervalloProgresso = null;
    let fileIdCorrente = null;

    // Gestione del click sulla zona di caricamento
    zonaCaricamento.addEventListener('click', () => {
        inputFile.click();
    });

    // Cambiamento dell'input file (selezione manuale)
    inputFile.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            gestisciFile(e.target.files[0]);
        }
    });

    // Gestione Drag & Drop
    ['dragenter', 'dragover'].forEach(nomeEvento => {
        zonaCaricamento.addEventListener(nomeEvento, (e) => {
            e.preventDefault();
            zonaCaricamento.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(nomeEvento => {
        zonaCaricamento.addEventListener(nomeEvento, (e) => {
            e.preventDefault();
            zonaCaricamento.classList.remove('dragover');
        }, false);
    });

    zonaCaricamento.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const file = dt.files[0];
        if (file) {
            gestisciFile(file);
        }
    });

    // Funzione principale di gestione ed elaborazione file
    function gestisciFile(file) {
        // Verifica estensione
        if (!file.name.endsWith('.xlsx')) {
            mostraErrore("Formato file non valido. Carica solo file Excel (.xlsx).");
            return;
        }

        // Mostra progresso ed resetta UI precedente
        schedaRisultati.style.display = 'none';
        contenitoreProgresso.style.display = 'flex';
        barraProgresso.style.width = '0%';
        testoStato.textContent = "Connessione al server...";

        // Avvia animazione progresso fittizia per migliorare la UX (arriva al 90% in attesa del backend)
        avviaProgressoAnimato();

        const datiForm = new FormData();
        datiForm.append('file', file);

        fetch('/carica', {
            method: 'POST',
            body: datiForm
        })
        .then(async risposta => {
            const dati = await risposta.json();
            if (!risposta.ok) {
                throw new Error(dati.errore || "Errore sconosciuto durante l'elaborazione.");
            }
            return dati;
        })
        .then(dati => {
            completaProgresso(() => {
                visualizzaRisultati(dati);
            });
        })
        .catch(errore => {
            fermaProgresso();
            contenitoreProgresso.style.display = 'none';
            mostraErrore(errore.message);
        });
    }

    // Gestione barra progresso fluida
    function avviaProgressoAnimato() {
        let percentuale = 0;
        testoStato.textContent = "Connessione al catalogo metadati RSDI...";
        
        intervalloProgresso = setInterval(() => {
            if (percentuale < 30) {
                percentuale += 2;
            } else if (percentuale < 60) {
                testoStato.textContent = "Analisi del file Excel e confronto con il catalogo...";
                percentuale += 1;
            } else if (percentuale < 90) {
                testoStato.textContent = "Scrittura delle colonne metadati e date...";
                percentuale += 0.5;
            }
            barraProgresso.style.width = `${percentuale}%`;
        }, 150);
    }

    function completaProgresso(callback) {
        clearInterval(intervalloProgresso);
        barraProgresso.style.width = '100%';
        testoStato.textContent = "Completato!";
        
        setTimeout(() => {
            contenitoreProgresso.style.display = 'none';
            if (callback) callback();
        }, 600);
    }

    function fermaProgresso() {
        clearInterval(intervalloProgresso);
    }

    // Visualizzazione dei dati elaborati
    function visualizzaRisultati(dati) {
        tabellaCorpo.innerHTML = '';
        let contatoreTrovati = 0;
        let contatoreNonTrovati = 0;

        dati.risultati.forEach(r => {
            const tr = document.createElement('tr');
            
            // Colonna Riga
            const tdRiga = document.createElement('td');
            tdRiga.className = 'col-riga';
            tdRiga.textContent = r.riga;
            tr.appendChild(tdRiga);

            // Colonna Nome Completo
            const tdNomeCompleto = document.createElement('td');
            tdNomeCompleto.className = 'col-layer';
            tdNomeCompleto.textContent = r.nome_completo || "-";
            tr.appendChild(tdNomeCompleto);

            // Colonna Nome (layer)
            const tdNome = document.createElement('td');
            tdNome.className = 'col-db';
            tdNome.textContent = r.nome || "-";
            tr.appendChild(tdNome);

            // Colonna Esito Catalogo (Status Badge)
            const tdEsito = document.createElement('td');
            const spanBadge = document.createElement('span');
            spanBadge.className = `status-badge ${r.esito.toLowerCase()}`;
            spanBadge.innerHTML = r.esito === 'YES' ? 'Trovato' : 'No';
            tdEsito.appendChild(spanBadge);
            tr.appendChild(tdEsito);

            // Colonna Catalogo Link
            const tdLink = document.createElement('td');
            tdLink.className = 'col-link';
            if (r.esito === 'YES' && r.link !== 'no') {
                const a = document.createElement('a');
                a.href = r.link;
                a.target = '_blank';
                a.innerHTML = 'Apri Scheda <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path><polyline points="15 3 21 3 21 9"></polyline><line x1="10" y1="14" x2="21" y2="3"></line></svg>';
                tdLink.appendChild(a);
                contatoreTrovati++;
            } else {
                const spanNo = document.createElement('span');
                spanNo.className = 'no-link';
                spanNo.textContent = 'no';
                tdLink.appendChild(spanNo);
                contatoreNonTrovati++;
            }
            tr.appendChild(tdLink);

            // Colonna Visualizzatore
            const tdVis = document.createElement('td');
            if (r.visualizzatore) {
                const spanVis = document.createElement('span');
                spanVis.className = `status-badge ${r.visualizzatore.toLowerCase()}`;
                spanVis.textContent = r.visualizzatore;
                tdVis.appendChild(spanVis);
            } else {
                tdVis.textContent = "-";
            }
            tr.appendChild(tdVis);

            // Colonna Data Pubblicazione
            const tdPub = document.createElement('td');
            tdPub.className = 'col-data';
            tdPub.textContent = r.pubblicazione || "-";
            tr.appendChild(tdPub);

            // Colonna Data Revisione
            const tdRev = document.createElement('td');
            tdRev.className = 'col-data';
            tdRev.textContent = r.revisione || "-";
            tr.appendChild(tdRev);

            tabellaCorpo.appendChild(tr);
        });

        // Aggiorna le statistiche in alto
        statTrovati.textContent = `${contatoreTrovati} Trovati`;
        statNonTrovati.textContent = `${contatoreNonTrovati} Non trovati`;

        // Imposta il link del file da scaricare
        fileIdCorrente = dati.fileId;
        aggiornaLinkDownload();

        // Rendi visibile la scheda dei risultati
        schedaRisultati.style.display = 'flex';
        schedaRisultati.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    function aggiornaLinkDownload() {
        if (!fileIdCorrente) return;
        const ancorato = chkAnchored.checked;
        bottoneScarica.href = `/scarica/${fileIdCorrente}?ancorato=${ancorato}`;
    }

    chkAnchored.addEventListener('change', aggiornaLinkDownload);

    function mostraErrore(messaggio) {
        alert("Errore: " + messaggio);
    }
});
