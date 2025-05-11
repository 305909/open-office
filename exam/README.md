# Analisi Vendite Prodotti Tipici Online

Il file **[test.xlsx](test.xlsx)** contiene una tabella con dati simulati relativi agli ordini di un e-commerce di prodotti tipici italiani.

## Struttura della tabella (anteprima)

| ID Ordine | Prodotto                      | Categoria Prodotto | Regione di Origine | Quantità Venduta | Prezzo Unitario | Spedizione Nazionale | Peso Lordo Pacco (kg) |
|-----------|-------------------------------|--------------------|--------------------|-----------------|-----------------|----------------------|-----------------------|
| 1001      | Olio EVO Taggiasco            | Olio               | Liguria            | 2               | 18,50           | SI                   | 1,2                   |
| 1002      | Parmigiano Reggiano 24 Mesi   | Formaggio          | Emilia-Romagna     | 1               | 22,00           | SI                   | 1,1                   |
| 1003      | Confettura di Fichi Bio       | Conserva           | Puglia             | 3               | 7,50            | NO                   | 0,9                   |
| ...       | ...                           | ...                | ...                | ...             | ...             | ...                  | ...                   |

## Colonne

- `ID Ordine` (colonna `A`)
- `Prodotto` (colonna `B`)
- `Categoria Prodotto` (colonna `C`)
- `Regione di Origine` (colonna `D`)
- `Quantità Venduta` (colonna `E`)
- `Prezzo Unitario` (colonna `F`)
- `Spedizione Nazionale` (colonna `G`)
- `Peso Lordo Pacco (kg)` (colonna `H`)

---

## Esercizi

### 1. **Totale quantità venduta di tutti i prodotti**

- **Dove**: cella `K2`

### 2. **Prezzo unitario medio di tutti i prodotti**

- **Dove**: cella `K3`

### 3. **Numero di prodotti della categoria "Vino”**

- **Dove**: cella `K4`

### 4. **Quantità totale venduta per i prodotti della regione "Toscana”**

- **Dove**: cella `K5`

### 5. **Prezzo unitario medio dei prodotti categoria "Olio”**

- **Dove**: cella `K6`

### 6. **Classificazione del Prezzo Unitario**

- **Dove**: aggiungere una nuova colonna `Fascia Prezzo` in colonna `I`, a partire da cella `I2` fino a `I61`
- **Descrizione**:

- Se il `Prezzo Unitario` (colonna `F`) è maggiore di `20€`, la fascia è `"Alta"`.
  
- Se il `Prezzo Unitario` è compreso tra `10€` (incluso) e `20€` (incluso), la fascia è `"Media"`.
  
- Altrimenti (minore di `10€`), la fascia è `"Bassa"`.

- **Estendere fino a `I61`**

### 7. **Valore Totale Ordine per riga**

- **Dove**: aggiungere una nuova colonna `Valore Ordine` in colonna `J`, da `J2` a `J61`
- **Descrizione**: Calcolare il valore totale per ciascuna riga d'ordine (`Quantità Venduta * Prezzo Unitario`).

- **Estendere fino a `J61`**

---

## Licenza

Questo materiale è distribuito con licenza **Creative Commons BY-NC-SA 4.0**

---

## Autore

**Prof. Francesco Giuseppe Gillio** – Corso di Informatica
Istituto di Istruzione Superiore Giovanni Cena di Ivrea
