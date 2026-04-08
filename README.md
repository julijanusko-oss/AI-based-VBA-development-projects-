# AI-based-VBA-development-projects-
# Outlook Levél Továbbító és Kategorizáló Makró

Ez egy Outlookhoz készült makró, ami segít gyorsabban továbbítani a beérkező leveleket, és automatikusan kategória kártyákat tesz rájuk. Azért csináltam, hogy kevesebbet kelljen kattintgatni és gépelni a mindennapi munka során.

## Fő funkciók
* Tárgy másolása: A továbbított levél tárgya automatikusan az a szöveg lesz, amit épp kimásoltál (ami a vágólapodon van).
* Címzettek kiválasztása: Egy felugró ablakban gombokkal és egy listából lehet kiválasztani, hogy melyik osztálynak vagy kollégának menjen a levél.
* Automatikus szöveg: Magától beírja a megfelelő megszólítást (pl. Szia Palkó!) és a rövid kísérőszöveget a levélbe.
* Kategória kártyák: Amikor rányomsz a Küldés gombra, a makró a háttérben ráteszi a megfelelő Outlook kategória kártyát az eredeti levélre.

## Fájlok a mappában
* Module1.bas - Ez a kód csinálja a munka nagy részét (a levél összeállítását és a címzettek kezelését).
* KategoriaFigyelo.cls - Ez a kód csak azt figyeli, hogy mikor nyomod meg a Küldés gombot a kártyákhoz.
* UserForm_kod.txt - Ez a felugró panel (az ablak) kódja.
* README.md - Ez a leírás.

## Fontos lépés használat előtt
Mielőtt elkezded használni, nyisd meg a Module1.bas fájlt, és írd át benne a példa e-mail címeket a valós céges címekre. A kategóriák neveit is pontosan úgy írd be a kódba, ahogy az Outlookodban szerepelnek.
