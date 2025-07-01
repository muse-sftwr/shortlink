# Shortener Excel Tool

Un semplice tool con interfaccia grafica (Tkinter) per accorciare URL presenti in un file Excel. I link vengono accorciati tramite API (Bit.ly o altri) e salvati direttamente nella colonna "ShortLink" del file.

## 🔧 Funzionalità
- Interfaccia utente semplice con Tkinter
- Accorciamento multiplo di URL separati da virgole
- Gestione colonne flessibile ("OriginalLink", "original link", ecc.)
- Scrive gli URL accorciati nella colonna accanto, nello stesso file Excel
- Gestione errori API e connettività
- Protezione da blocchi IP (delay 0.2s tra le richieste)

## ▶️ Requisiti
- Python 3.9+
- `requests`, `pandas`, `openpyxl`

## 🛠️ Installazione

Clona il repository:

```bash
git clone https://github.com/TUO_USERNAME/NOME_REPO.git
cd NOME_REPO
pip install -r requirements.txt
