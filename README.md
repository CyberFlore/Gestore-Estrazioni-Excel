
# ExcelTemplateManager 📊

<div align="center">
  <img src="generated-icon.png" alt="ExcelTemplateManager Logo" width="120"/>
  <h3>Gestione avanzata di template Excel</h3>
  <p>Riorganizza, rinomina e applica validazioni alle colonne dei tuoi fogli Excel</p>
</div>

## 🌟 Panoramica

ExcelTemplateManager è un'applicazione Windows che semplifica la gestione e la trasformazione di fogli Excel. Consente di:

- **Riorganizzare** l'ordine delle colonne con semplici operazioni di drag & drop
- **Rinominare** le intestazioni delle colonne per una migliore leggibilità
- **Aggiungere nuove colonne** ai fogli esistenti
- **Applicare validazioni** per garantire l'inserimento di dati corretti
- **Salvare e riutilizzare template** per applicarli a fogli diversi

## 🖥️ Screenshot

<div align="center">
  <p><i>Interfaccia principale di ExcelTemplateManager</i></p>
  <p>(Aggiungi uno screenshot qui una volta disponibile)</p>
</div>

## ✨ Funzionalità principali

### 📋 Gestione delle colonne
- Visualizzazione delle colonne originali del file Excel
- Riordinamento tramite drag & drop intuitivo
- Aggiunta di nuove colonne (non presenti nel file originale)
- Eliminazione di colonne non necessarie

### 🔄 Template personalizzati
- Salvataggio di configurazioni colonnari come template riutilizzabili
- Gestione di template multipli in memoria
- Esportazione e importazione di template su file
- Applicazione automatica di template salvati a nuovi file

### ✅ Validazione dei dati
- Aggiunta di liste di valori per la validazione delle celle
- Configurazione rapida delle opzioni di validazione tramite interfaccia dedicata
- Mantenimento delle validazioni nei file esportati

## 📦 Installazione

### Download diretto
1. Scarica l'ultima versione dell'applicazione dalla [sezione Releases](https://github.com/username/ExcelTemplateManager/releases)
2. Esegui il file `ExcelTemplateManager.exe` (non richiede installazione)

### Compilazione da sorgente
1. Clona il repository
2. Assicurati di avere .NET SDK 7.0+ installato
3. Esegui il comando:
```
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:InvariantGlobalization=true -o ./publish
```
4. L'eseguibile si troverà nella cartella `publish`

## 📖 Utilizzo

### Aprire e modificare un file Excel
1. Avvia l'applicazione
2. Fai clic su "Apri file" e seleziona un file Excel (.xlsx)
3. Visualizza le colonne nella lista centrale
4. Modifica l'ordine trascinando le colonne nella posizione desiderata
5. Rinomina le colonne tramite il menu contestuale (clic destro)

### Gestire i template
1. Modifica l'ordine e i nomi delle colonne come desiderato
2. Fai clic su "Salva Template in Memoria" per salvare la configurazione attuale
3. Utilizza il menu a tendina per caricare i template salvati
4. Utilizza "Rimuovi Template" per eliminare i template non più necessari

### Aggiungere validazioni
1. Seleziona una colonna
2. Fai clic destro e seleziona "Modifica opzioni di validazione"
3. Aggiungi le opzioni desiderate nella finestra di dialogo
4. Conferma per salvare le opzioni di validazione

### Esportare il file
1. Dopo aver configurato le colonne, fai clic su "Esporta"
2. Seleziona la posizione in cui salvare il nuovo file Excel
3. Il file esportato conterrà tutte le modifiche applicate

## 🔧 Tecnologie utilizzate

- **C# / .NET 7.0**: Linguaggio e framework di sviluppo
- **Windows Forms**: Framework UI per l'interfaccia grafica
- **EPPlus**: Libreria per la manipolazione di file Excel
- **Serializzazione binaria**: Per il salvataggio e caricamento di template

## 🤝 Contributi

Contributi, suggerimenti e segnalazioni di bug sono benvenuti! Sentiti libero di:

- Aprire una issue per segnalare bug o richiedere nuove funzionalità
- Proporre miglioramenti attraverso pull request
- Condividere l'applicazione con chi potrebbe trovarla utile

## 📄 Licenza

Questo progetto è rilasciato sotto licenza MIT. Vedi il file `LICENSE` per i dettagli.

## 👤 Autore

Creato da Mirko Subri

---

<div align="center">
  <p>Made with ❤️ for Excel users</p>
</div>
