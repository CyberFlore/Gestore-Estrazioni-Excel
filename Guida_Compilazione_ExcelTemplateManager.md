
# Guida alla compilazione di ExcelTemplateManager in Windows

Questo documento fornisce una guida dettagliata su come compilare e generare l'eseguibile dell'applicazione ExcelTemplateManager su un sistema Windows.

## Prerequisiti

Prima di iniziare, assicurati di avere installato:

1. **.NET SDK 6.0 o superiore** - [Download .NET SDK](https://dotnet.microsoft.com/download)
2. **Visual Studio 2022** (opzionale ma consigliato) - [Download Visual Studio](https://visualstudio.microsoft.com/it/downloads/)

## Passaggi per la compilazione con riga di comando

### 1. Preparazione dell'ambiente

1. Dopo aver installato .NET SDK, apri il **Prompt dei comandi** o **PowerShell** come amministratore
2. Verifica che .NET sia installato correttamente eseguendo:
   ```
   dotnet --version
   ```

### 2. Download e posizionamento dei file sorgente

1. Crea una cartella per il progetto, ad esempio: `C:\Projects\ExcelTemplateManager`
2. Posiziona tutti i file sorgente del progetto nella cartella appena creata

### 3. Ripristino dei pacchetti NuGet

1. Naviga nella cartella del progetto:
   ```
   cd C:\Projects\ExcelTemplateManager
   ```
2. Esegui il comando per ripristinare i pacchetti:
   ```
   dotnet restore
   ```

### 4. Verifica che il file di progetto contenga EPPlus

Assicurati che il file `ExcelTemplateManager.csproj` contenga il riferimento al pacchetto EPPlus. Se manca, apri il file e aggiungi:

```xml
<ItemGroup>
  <PackageReference Include="EPPlus" Version="6.2.10" />
</ItemGroup>
```

### 5. Compilazione dell'eseguibile standalone

1. Esegui il seguente comando per generare un eseguibile Windows autonomo:
   ```
   dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:InvariantGlobalization=true -o .\publish
   ```

2. Al termine della compilazione, troverai l'eseguibile nella cartella `publish`:
   ```
   C:\Projects\ExcelTemplateManager\publish\ExcelTemplateManager.exe
   ```

## Passaggi per la compilazione con Visual Studio 2022

Se preferisci utilizzare Visual Studio, ecco come procedere:

### 1. Aprire il progetto

1. Avvia Visual Studio 2022
2. Seleziona **Apri un progetto o una soluzione**
3. Naviga nella cartella del progetto e seleziona il file `ExcelTemplateManager.csproj`

### 2. Configura le impostazioni di pubblicazione

1. Fai clic con il pulsante destro sul progetto nel **Solution Explorer**
2. Seleziona **Pubblica...**
3. Scegli **Cartella** come destinazione di pubblicazione
4. Nella schermata successiva, scegli **Più configurazioni...**
5. Configura le seguenti impostazioni:
   - **Configurazione**: Release
   - **Framework di destinazione**: net6.0-windows
   - **Sistema operativo di destinazione**: Windows
   - **Modalità di distribuzione**: Autonomo
   - **Versione di runtime di destinazione**: win-x64
   - **Produci file singolo**: Selezionato
   - **Pubblicazione trim**: Non selezionato
   - **Globalizzazione invariante**: Selezionato
   - **Cartella di destinazione**: .\publish

6. Fai clic su **Salva** e poi su **Pubblica**

### 3. Accesso all'eseguibile

Al termine della pubblicazione, troverai l'eseguibile nella cartella specificata:
```
[Percorso del progetto]\publish\ExcelTemplateManager.exe
```

## Utilizzo dell'applicazione

1. Fai doppio clic su `ExcelTemplateManager.exe` per avviare l'applicazione
2. Utilizza i pulsanti nell'interfaccia per:
   - Aprire un file Excel
   - Salvare un template
   - Caricare un template
   - Esportare un file Excel

## Risoluzione dei problemi

Se riscontri problemi durante la compilazione:

1. **Errore di dipendenze mancanti**:
   Verifica che il .NET SDK sia installato correttamente e che tutti i pacchetti NuGet siano stati ripristinati.

2. **Errore di compilazione**:
   Verifica che il codice sorgente sia completo e che non ci siano errori di sintassi.

3. **Errori relativi a EPPlus**:
   Assicurati che il pacchetto EPPlus sia installato correttamente:
   ```
   dotnet add package EPPlus --version 6.2.10
   ```

## Note finali

- L'eseguibile generato è completamente autonomo e non richiede .NET installato sulla macchina di destinazione
- La dimensione del file sarà maggiore (circa 60-80 MB) rispetto a un'applicazione non autonoma
- L'applicazione è ottimizzata per sistemi Windows a 64 bit
