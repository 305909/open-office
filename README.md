<img src="apache.png" width="300"/>

# Guida OpenOffice Calc

##### di Francesco Giuseppe Gillio  
###### I.I.S. G. Cena di Ivrea

---

## Indice

1. [Funzioni di Base: SOMMA, MEDIA, CONTA](#1-funzioni-di-base-somma-media-conta)  
   - [1.1. SOMMA](#11-somma)  
   - [1.2. MEDIA](#12-media)  
   - [1.3. CONTA](#13-conta)  

2. [Funzioni di Ricerca e Condizionali: MAX, MIN, SE](#2-funzioni-di-ricerca-e-condizionali-max-min-se)  
   - [2.1. MAX](#21-max)  
   - [2.2. MIN](#22-min)  
   - [2.3. SE](#23-se)  

3. [Funzioni Condizionali: SOMMA.SE, MEDIA.SE, CONTA.SE](#3-funzioni-condizionali-sommase-mediase-contase)  
   - [3.1. SOMMA.SE](#31-sommase)  
   - [3.2. MEDIA.SE](#32-mediase)  
   - [3.3. CONTA.SE](#33-contase)  

4. [Funzioni Condizionali Avanzate: SOMMA.PIÙ.SE, MEDIA.PIÙ.SE, CONTA.PIÙ.SE](#4-funzioni-condizionali-avanzate-sommapiùse-mediapiùse-contapiùse)  
   - [4.1. SOMMA.PIÙ.SE](#41-sommapiùse)  
   - [4.2. MEDIA.PIÙ.SE](#42-mediapiùse)  
   - [4.3. CONTA.PIÙ.SE](#43-contapiùse)

5. [Manuale](#5-manuale)

---

## Introduzione a OpenOffice Calc

**OpenOffice Calc** è un software di calcolo elettronico della suite *OpenOffice* per l'analisi dati. Calc integra un set di funzioni matematiche e statistiche per eseguire operazioni aritmetiche come `SOMMA`, `MEDIA`, `CONTA`, nonché le varianti condizionali `SOMMA.SE`, `MEDIA.SE`, `CONTA.SE`, `SOMMA.PIÙ.SE`, `MEDIA.PIÙ.SE` e `CONTA.PIÙ.SE`.

Il foglio di calcolo si presenta nel formato:

|     | `A`  | `B`  | `C`  | `D`  |  
|-----|-----|-----|-----|-----|  
| `1` | `A1` | `B1` | `C1` | `D1` |  
| `2` | `A2` | `B2` | `C2` | `D2` |  
| `3` | `A3` | `B3` | `C3` | `D3` |  
| `4` | `A4` | `B4` | `C4` | `D4` |  

dove:  
- ciascuna **colonna** è identificata da una lettera (`A`, `B`, `C`, …);  
- ciascuna **riga** è identificata da un numero (`1`, `2`, `3`, …);
- ciascuna **cella** è identificate da una combinazione alfa-numerica che indica la colonna e la riga di appartenenza (ad esempio, `A1` si riferisce alla cella di intersezione della colonna `A` e della riga `1`).

---

## 1. Funzioni di Base: SOMMA, MEDIA, CONTA  

---

### 1.1. SOMMA

La funzione `SOMMA` calcola la **somma** dei valori numerici presenti in un intervallo di celle:



  $$\text{SOMMA}(\text{intervallo}) = \sum_{i=1}^{n} x_i$$


  
dove: 
- $\text{intervallo}$ è l’insieme di celle oggetto della funzione,
- $\sum_{i=1}^{n} x_i = x_1 + x_2 + \dots + x_n$,
- $n$ è il numero di celle con valori numerici nell’intervallo,
- $x_i$ è il valore numerico della $i$-esima cella.

La formula per la funzione `SOMMA` in Calc è:
  
  `=SOMMA(intervallo di celle)`

**Esempio:** 

Si consideri l'intervallo numerico `A1:A5 = {5; 10; 15; 20; 25}`. La funzione `=SOMMA(A1:A5)` restituisce il valore:



  $$5 + 10 + 15 + 20 + 25 = 75$$



---

### 1.2. MEDIA  

La funzione `MEDIA` calcola la **media aritmetica** dei valori numerici presenti in un intervallo di celle:



  $$\text{MEDIA}(\text{intervallo}) = \frac{\sum_{i=1}^{n} x_i}{n}$$


  
dove: 
- $\text{intervallo}$ è l’insieme di celle oggetto della funzione,
- $\sum_{i=1}^{n} x_i = x_1 + x_2 + \dots + x_n$,
- $n$ è il numero di celle con valori numerici nell’intervallo,
- $x_i$ è il valore numerico della $i$-esima cella.

La formula per la funzione `MEDIA` in Calc è:
  
  `=MEDIA(intervallo di celle)`
  
**Esempio:**  

Si consideri l'intervallo numerico `B1:B4 = {4; 6; 8; 10}`. La funzione `=MEDIA(B1:B4)` restituisce il valore:



  $$\frac{4 + 6 + 8 + 10}{4} = 7$$


 
---

### 1.3. CONTA  

La funzione `CONTA` calcola il **numero** di valori numerici presenti in un intervallo di celle:



  $$\text{CONTA}(\text{intervallo}) = \sum_{i=1}^{n} I_{c}(x_i)$$


  
dove: 
- $\text{intervallo}$ è l’insieme di celle oggetto della funzione,
- $\sum_{i=1}^{n} I_{c}(x_i) = I_{c}(x_1) + I_{c}(x_2) + \dots + I_{c}(x_n)$,
- $n$ è il numero di celle con valori numerici nell’intervallo,
- $x_i$ è il valore numerico della $i$-esima cella,
- $I_{c}(x_i)$ è la funzione indicatrice, che restituisce $1$ se $x_i$ è un valore numerico, ossia se rispetta la condizione $c = \text{è un numero}$, altrimenti $0$.

La formula per la funzione `CONTA` in Calc è:
  
  `=CONTA(intervallo di celle)`
  
**Esempio:**

Si consideri l'intervallo numerico `C1:C6 = {7; "testo"; 3; vuoto; 9; 5}`. La funzione `=CONTA(C1:C6)` restituisce il valore:



  $$1 + 0 + 1 + 0 + 1 + 1 = 4$$



---

## 2. Funzioni di Ricerca e Condizionali: MAX, MIN, SE  

---

### 2.1. MAX  

La funzione `MAX` cerca il **valore massimo** tra i valori numerici presenti in un intervallo di celle:



  $$\text{MAX}(\text{intervallo}) = \max(x_1, x_2, \dots, x_n)$$


  
dove: 
- $\text{intervallo}$ è l’insieme di celle oggetto della funzione,
- $x_i$ è il valore numerico della $i$-esima cella,
- $\max(x_1, x_2, \dots, x_n)$ è il valore *massimo* presente nell'intervallo.

La formula per la funzione `MAX` in Calc è:
  
  `=MAX(intervallo di celle)`

**Esempio:**  

Si consideri l'intervallo numerico `D1:D5 = {12; 34; 27; 8; 19}`. La funzione `=MAX(D1:D5)` restituisce il valore:



  $$\max(12, 34, 27, 8, 19) = 34$$



---

### 2.2. MIN  

La funzione `MIN` cerca il **valore minimo** tra i valori numerici presenti in un intervallo di celle:



  $$\text{MIN}(\text{intervallo}) = \min(x_1, x_2, \dots, x_n)$$


  
dove: 
- $\text{intervallo}$ è l’insieme di celle oggetto della funzione,
- $x_i$ è il valore numerico della $i$-esima cella,
- $\min(x_1, x_2, \dots, x_n)$ è il valore *minimo* presente nell'intervallo.

La formula per la funzione `MIN` in Calc è:
  
  `=MIN(intervallo di celle)`

**Esempio:**  

Si consideri l'intervallo numerico `E1:E4 = {22; 45; 18; 30}`. La funzione `=MIN(E1:E4)` restituisce il valore:



  $$\min(22, 45, 18, 30) = 18$$



---

### 2.3. SE  

La funzione `SE` restituisce un valore in base a una *condizione logica*:



  $$\text{SE}(\text{test}; \text{ vero}; \text{ falso})$$


  
dove: 
- `test` è la *condizione logica*,
- `vero` è il valore che la funzione restituisce se il `test` è vero,
- `falso` è il valore che la funzione restituisce se il `test` è falso.

La formula per la funzione `SE` in Calc è:
  
  `=SE(condizione; valore se vero; valore se falso)`

**Esempio:**  

Si consideri il valore numerico `A1 = 10`. La funzione `=SE(A1>=5; "Si"; "No")` restituisce il valore:



  $$\text{"Si"}$$



---

## 3. Funzioni Condizionali: SOMMA.SE, MEDIA.SE, CONTA.SE  

---

### 3.1. SOMMA.SE  

La funzione `SOMMA.SE` calcola la **somma** dei valori numerici presenti in un intervallo di celle che rispettano una *condizione*:



  $$\text{SOMMA.SE}(\text{intervallo condizionale}; \text{ condizione}; \text{ intervallo di somma}) = \sum_{i=1}^{n} I_{c}(y_i) \cdot x_i$$


  
dove: 
- $\text{intervallo condizionale}$ è l’insieme di celle oggetto della funzione `SE`,
- $\text{intervallo di somma}$ è l’insieme di celle oggetto della funzione `SOMMA`,
- $\sum_{i=1}^{n} I_{c}(y_i) \cdot x_i = I_{c}(y_1)\cdot x_1 + I_{c}(y_2) \cdot x_2 + \dots + I_{c}(y_n) \cdot x_n$,
- $n$ è il numero di celle con valori numerici nell’intervallo di somma,
- $x_i$ è il valore numerico della $i$-esima cella nell'intervallo di somma,
- $I_{c}(y_i)$ è la funzione indicatrice, che restituisce $1$ se $y_i$ - il valore alfa-numerico della $i$-esima cella nell'intervallo condizionale - rispetta la *condizione* $c$, altrimenti $0$.

La formula per la funzione `SOMMA.SE` in Calc è:
  
  `=SOMMA.SE(intervallo condizionale; condizione; intervallo di somma)`

**Esempio:**

Si consideri l'insieme di dati:

| Prodotto | Prezzo |
|----------|--------|
| A        | 15     |
| B        | 20     |
| A        | 25     |

La funzione `=SOMMA.SE(A1:A3;"A";B1:B3)` restituisce il valore:




  $$1 \cdot 15 + 0 \cdot 20 + 1 \cdot 25 = 40$$


  

---

### 3.2. MEDIA.SE  

La funzione `MEDIA.SE` calcola la **media aritmetica** dei valori numerici presenti in un intervallo di celle che rispettano una *condizione*:



  $$\text{MEDIA.SE}(\text{intervallo condizionale}; \text{ condizione}; \text{ intervallo di media}) = \frac{\sum_{i=1}^{n} I_{c}(y_i) \cdot x_i}{\sum_{i=1}^{n} I_{c}(y_i)}$$


  
dove:
- $\text{intervallo condizionale}$ è l’insieme di celle oggetto della funzione `SE`,
- $\text{intervallo di media}$ è l’insieme di celle oggetto della funzione `MEDIA`,
- $\sum_{i=1}^{n} I_{c}(y_i) \cdot x_i = I_{c}(y_1)\cdot x_1 + I_{c}(y_2) \cdot x_2 + \dots + I_{c}(y_n) \cdot x_n$,
- $\sum_{i=1}^{n} I_{c}(y_i) = I_{c}(y_1) + I_{c}(y_2) + \dots + I_{c}(y_n)$,
- $n$ è il numero di celle con valori numerici nell’intervallo di media,
- $x_i$ è il valore numerico della $i$-esima cella nell'intervallo di media,
- $I_{c}(y_i)$ è la funzione indicatrice, che restituisce $1$ se $y_i$ - il valore alfa-numerico della $i$-esima cella nell'intervallo condizionale - rispetta la *condizione* $c$, altrimenti $0$.

La formula per la funzione `MEDIA.SE` in Calc è:
  
  `=MEDIA.SE(intervallo condizionale; condizione; intervallo di media)`

**Esempio:**

Si consideri l'insieme di dati:

| Prodotto | Prezzo |
|----------|--------|
| A        | 15     |
| B        | 20     |
| A        | 25     |

La funzione `=MEDIA.SE(A1:A3;"A";B1:B3)` restituisce il valore:




  $$\frac{1 \cdot 15 + 0 \cdot 20 + 1  \cdot 25}{1 + 0 + 1} = 20$$


  

---

### 3.3. CONTA.SE  

La funzione `CONTA.SE` calcola il **numero** di valori alfa-numerici presenti in un intervallo di celle che rispettano una *condizione*:



  $$\text{CONTA.SE}(\text{intervallo}; \text{ condizione}) = \sum_{i=1}^{n} I_{c}(x_i)$$


  
dove:
- $\text{intervallo}$ è l’insieme di celle oggetto della funzione `CONTA.SE`,
- $\sum_{i=1}^{n} I_{c}(x_i) = I_{c}(x_1) + I_{c}(x_2) + \dots + I_{c}(x_n)$,
- $n$ è il numero di celle con valori alfa-numerici nell’intervallo,
- $x_i$ è il valore alfa-numerico della $i$-esima cella,
- $I_{c}(x_i)$ è la funzione indicatrice, che restituisce $1$ se $x_i$ rispetta la *condizione* $c$, altrimenti $0$.

La formula per la funzione `CONTA.SE` in Calc è:
  
  `=CONTA.SE(intervallo; condizione)`

**Esempio:**

Si consideri l'insieme di dati:

| Prodotto | Prezzo |
|----------|--------|
| A        | 15     |
| B        | 20     |
| A        | 25     |

La funzione `=CONTA.SE(A1:A3;"A")` restituisce il valore:




  $$1 + 0 + 1 = 2$$



  
---

## 4. Funzioni Condizionali Avanzate: SOMMA.PIÙ.SE, MEDIA.PIÙ.SE, CONTA.PIÙ.SE  

---

### 4.1. SOMMA.PIÙ.SE

La funzione `SOMMA.PIÙ.SE` calcola la **somma** dei valori numerici presenti in un intervallo di celle che rispettano *diverse condizioni* contemporaneamente:


  
$$\text{SOMMA.PIÙ.SE}(\text{intervallo di somma}; \text{ intervallo condizionale A}; \text{ condizione A}; \text{ intervallo condizionale B}; \text{ condizione B}; \dots) = \sum_{i=1}^{n} \left( \prod_{j=1}^{m} I_{c_j}(y_{ij}) \right) \cdot x_i$$


  

dove:  
- $\text{intervallo di somma}$ è l’insieme di celle oggetto della funzione `SOMMA`,  
- $\text{intervallo condizionale } j$ è l’insieme di celle oggetto della $j$-esima condizione `SE`,  
- $\text{condizione } j$ è il criterio per ciascun valore alfa-numerico dell'intervallo condizionale $j$,  
- $x_i$ è il valore numerico della $i$-esima cella nell'intervallo di somma,  
- $y_{ij}$ è il valore alfa-numerico della $i$-esima cella nell’intervallo condizionale $j$,  
- $I_{c_j}(y_{ij})$ è la funzione indicatrice che restituisce $1$ se $y_{ij}$ rispetta la *condizione* $c_j$, altrimenti $0$,  
- $n$ è il numero di celle con valori numerici nell’intervallo di somma,
- $m$ è il numero di condizioni.

La formula per la funzione `SOMMA.PIÙ.SE` in Calc è:

```
=SOMMA.PIÙ.SE(intervallo di somma; intervallo condizionale A; condizione A; intervallo condizionale B; condizione B; …)
```

**Esempio:**

Si consideri l’insieme di dati:

| Prodotto | Prezzo | Quantità |
|----------|--------|----------|
| A        | 15     | 5        |
| B        | 20     | 8        |
| A        | 25     | 3        |

La funzione `=SOMMA.PIÙ.SE(B1:B3; A1:A3; "A"; C1:C3; ">4")` calcola la somma dei valori presenti nell’intervallo `B1:B3` per i quali, corrispondentemente, il valore in `A1:A3` è `"A"` **e** il valore in `C1:C3` è maggiore di 4. In questo esempio, la funzione restituisce:




$$15$$



  
---

### 4.2. MEDIA.PIÙ.SE

La funzione `MEDIA.PIÙ.SE` calcola la **media aritmetica** dei valori numerici presenti in un intervallo di celle che rispettano *diverse condizioni* contemporaneamente:


  
$$\text{MEDIA.PIÙ.SE}(\text{intervallo di media}; \text{ intervallo condizionale A}; \text{ condizione A}; \text{ intervallo condizionale B}; \text{ condizione B}; \dots) = \frac{\sum_{i=1}^{n} \left( \prod_{j=1}^{m} I_{c_j}(y_{ij}) \right) \cdot x_i}{\sum_{i=1}^{n} \prod_{j=1}^{m} I_{c_j}(y_{ij})}$$


  

dove:  
- $\text{intervallo di media}$ è l’insieme di celle oggetto della funzione `MEDIA`,  
- $\text{intervallo condizionale } j$ è l’insieme di celle oggetto della $j$-esima condizione `SE`,  
- $\text{condizione } j$ è il criterio per ciascun valore alfa-numerico dell'intervallo condizionale $j$,  
- $x_i$ è il valore numerico della $i$-esima cella nell'intervallo di media,  
- $y_{ij}$ è il valore alfa-numerico della $i$-esima cella nell’intervallo condizionale $j$,  
- $I_{c_j}(y_{ij})$ è la funzione indicatrice che restituisce $1$ se $y_{ij}$ rispetta la *condizione* $c_j$, altrimenti $0$,  
- $n$ è il numero di celle con valori numerici nell’intervallo di media,
- $m$ è il numero di condizioni.

La formula per la funzione `MEDIA.PIÙ.SE` in Calc è:

```
=MEDIA.PIÙ.SE(intervallo di media; intervallo condizionale A; condizione A; intervallo condizionale B; condizione B; …)
```

**Esempio:**

Si consideri l’insieme di dati:

| Prodotto | Prezzo | Quantità |
|----------|--------|----------|
| A        | 15     | 5        |
| B        | 20     | 8        |
| A        | 25     | 3        |

La funzione `=MEDIA.PIÙ.SE(B1:B3; A1:A3; "A"; C1:C3; ">2")` calcola la media dei valori presenti nell’intervallo `B1:B3` per i quali, corrispondentemente, il valore in `A1:A3` è `"A"` **e** il valore in `C1:C3` è maggiore di 2. In questo esempio, la funzione restituisce:




$$20$$



  
---

### 4.3. CONTA.PIÙ.SE

La funzione `CONTA.PIÙ.SE` calcola il **numero** di valori alfa-numerici presenti in un intervallo di celle che rispettano *diverse condizioni* contemporaneamente:


  
$$\text{CONTA.PIÙ.SE}(\text{ intervallo condizionale A}; \text{ condizione A}; \text{ intervallo condizionale B}; \text{ condizione B}; \dots) = \sum_{i=1}^{n} \prod_{j=1}^{m} I_{c_j}(y_{ij})$$



dove:  
- $\text{intervallo condizionale } j$ è l’insieme di celle oggetto della $j$-esima condizione `SE`,  
- $\text{condizione } j$ è il criterio per ciascun valore alfa-numerico dell'intervallo condizionale $j$,  
- $y_{ij}$ è il valore alfa-numerico della $i$-esima cella nell’intervallo condizionale $j$,  
- $I_{c_j}(y_{ij})$ è la funzione indicatrice che restituisce $1$ se $y_{ij}$ rispetta la *condizione* $c_j$, altrimenti $0$,  
- $n$ è il numero di celle con valori alfa-numerici nell’intervallo,
- $m$ è il numero di condizioni.

La formula per la funzione `CONTA.PIÙ.SE` in Calc è:

```
=CONTA.PIÙ.SE(intervallo condizionale A; condizione A; intervallo condizionale B; condizione B; …)
```

**Esempio:**

Si consideri l’insieme di dati:

| Prodotto | Prezzo | Quantità |
|----------|--------|----------|
| A        | 15     | 5        |
| B        | 20     | 8        |
| A        | 25     | 3        |

La funzione `=CONTA.PIÙ.SE(A1:A3; "A"; C1:C3; ">4")` calcola il numero di celle per cui il valore in `A1:A3` è `"A"` **e** il valore in `C1:C3` è maggiore di 4. In questo esempio, la funzione restituisce:




$$1$$



  
---

## 5. Manuale

| Funzione            | Sintassi                                                                                      | Descrizione                                                                                           |
|---------------------|-----------------------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------|
| [SOMMA](#11-somma)  | `=SOMMA(intervallo di celle)`                                                                  | Calcola la **somma** dei valori numerici presenti nell'intervallo.                                     |
| [MEDIA](#12-media)  | `=MEDIA(intervallo di celle)`                                                                  | Calcola la **media aritmetica** dei valori numerici presenti nell'intervallo.                         |
| [CONTA](#13-conta)  | `=CONTA(intervallo di celle)`                                                                  | Calcola il **numero** di valori numerici presenti nell'intervallo.                                    |
| [MAX](#21-max)      | `=MAX(intervallo di celle)`                                                                    | Cerca il **valore massimo** tra i valori numerici presenti nell'intervallo.                           |
| [MIN](#22-min)      | `=MIN(intervallo di celle)`                                                                    | Cerca il **valore minimo** tra i valori numerici presenti nell'intervallo.                            |
| [SE](#23-se)        | `=SE(condizione; valore se vero; valore se falso)`                                             | Restituisce un valore in base ad una *condizione logica*.                                             |
| [SOMMA.SE](#31-sommase)  | `=SOMMA.SE(intervallo condizionale; condizione; intervallo di somma)`                        | Calcola la **somma** dei valori numerici presenti in un intervallo che rispettano una *condizione*.   |
| [MEDIA.SE](#32-mediase)  | `=MEDIA.SE(intervallo condizionale; condizione; intervallo di media)`                         | Calcola la **media aritmetica** dei valori numerici presenti in un intervallo che rispettano una *condizione*. |
| [CONTA.SE](#33-contase)  | `=CONTA.SE(intervallo; condizione)`                                                          | Calcola il **numero** di valori alfa-numerici presenti in un intervallo che rispettano una *condizione*. |
| [SOMMA.PIÙ.SE](#41-sommapiùse) | `=SOMMA.PIÙ.SE(intervallo di somma; intervallo condizionale A; condizione A; ...)`         | Calcola la **somma** dei valori numerici che rispettano *diverse condizioni* contemporaneamente.       |
| [MEDIA.PIÙ.SE](#42-mediapiùse) | `=MEDIA.PIÙ.SE(intervallo di media; intervallo condizionale A; condizione A; ...)`         | Calcola la **media aritmetica** dei valori numerici che rispettano *diverse condizioni* contemporaneamente. |
| [CONTA.PIÙ.SE](#43-contapiùse) | `=CONTA.PIÙ.SE(intervallo condizionale A; condizione A; ...)`                              | Calcola il **numero** di valori che rispettano *diverse condizioni* contemporaneamente.               |

---




