# Analisi Stagione Turistica

Il file **test.xlsx** contiene una tabella con dati simulati relativi alle prenotazioni settimanali di diverse strutture ricettive.

## Struttura della tabella (anteprima)

| Settimana | Struttura          | Tipo Struttura | Città       | Posti Prenotati | Prezzo Medio a Notte | Clienti Italiani | Clienti Esteri | 
|-----------|--------------------|----------------|-------------|------------------|------------------------|-------------------|----------------|
| 1         | Hotel Sole   | Hotel          | Firenze     | 45               | 90                    | 30                | 15             | 
| 1         | B&B Luna | B&B            | Venezia     | 18               | 80                     | 12                | 6              | 
| 1         | Agriturismo Verde   | Agriturismo    | Siena       | 25               | 95                     | 20                | 5              | 
| ...       | ...                | ...            | ...         | ...              | ...                    | ...               | ...            | 

## Colonne

- `Settimana` (colonna `A`)
- `Struttura` (colonna `B`)
- `Tipo Struttura` (colonna `C`)
- `Città` (colonna `D`)
- `Posti Prenotati` (colonna `E`)
- `Prezzo Medio a Notte` (colonna `F`)
- `Clienti Italiani` (colonna `G`)
- `Clienti Esteri` (colonna `H`)

---

## Esercizi

### 1. **Totale posti prenotati in tutta la stagione**

- **Dove**: cella `K2`
- **Formula attesa**:

```excel
=SOMMA(E2:E61)
````

### 2. **Prezzo medio per notte su tutte le strutture**

* **Dove**: cella `K3`
* **Formula attesa**:

```excel
=MEDIA(F2:F61)
```

### 3. **Numero di Hotel**

* **Dove**: cella `K4`
* **Formula attesa**:

```excel
=CONTA.SE(C2:C61; "Hotel")
```

### 4. **Posti prenotati nei B\&B**

* **Dove**: cella `K5`
* **Formula attesa**:

```excel
=SOMMA.SE(C2:C61; "B&B"; E2:E61)
```

### 5. **Prezzo medio per notte solo per gli agriturismi**

* **Dove**: cella `K6`
* **Formula attesa**:

```excel
=MEDIA.SE(C2:C61; "Agriturismo"; F2:F61)
```

### 6. **Classificazione dell’affluenza**

* **Dove**: aggiungere una nuova colonna `Affluenza` in colonna `I`, a partire da cella `I2` fino a `I61`
* **Formula da inserire in `I2`**:

```excel
=SE(E2>40; "Alta"; SE(E2>=20; "Media"; "Bassa"))
```

* **Estendere fino a `I61`**

### 7. **Ricavo stimato per riga**

* **Dove**: aggiungere una nuova colonna `Ricavo Stimato` in colonna `J`, da `J2` a `J61`
* **Formula da inserire in `J2`**:

```excel
=E2 * F2
```

* **Estendere fino a `J61`**

---

## Licenza

Questo materiale è distribuito con licenza **Creative Commons BY-NC-SA 4.0**

---

## Autore

**Prof. Francesco Giuseppe Gillio** – Corso di Informatica
Istituto di Istruzione Superiore Giovanni Cena di Ivrea
