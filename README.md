Fájl Átalakító Alkalmazás

Név: Bencze Ákos

Feladat rövid leírása

A projekt célja egy Python alapú, grafikus felhasználói felülettel rendelkező alkalmazás készítése, amely képes PDF fájlokból Word dokumentumot (DOCX), illetve Word dokumentumból PDF fájlt előállítani.  
Az alkalmazás lehetővé teszi egy tetszőleges mappa kiválasztását, a benne található *.pdf és *.docx kiterjesztésű fájlok listázását, majd a kijelölt fájl konvertálását a másik formátumba.

A program használata egyszerű, nem igényel technikai előképzettséget, a konverziót egyetlen gombnyomással végrehajtja, és automatikusan tájékoztatja a felhasználót az eredményről vagy az esetleges hibákról.



Felhasznált modulok:

tkinter - Grafikus felület létrehozása 
os - Fájlrendszer bejárása és fájlkezelés 
docx - DOCX dokumentum létrehozása PDF-ből 
docx2pdf - Word → PDF konverzió végrehajtása 
fitz (PyMuPDF) - PDF szöveg kinyerése DOCX generáláshoz 
platform - Operációs rendszer ellenőrzése 
subprocess - Microsoft Word folyamatainak lezárása Windows alatt 


Függvények felsorolása

__init__ - GUI felület inicializálása, stílus és widgetek létrehozása |
BÁ_refresh_file_list() - Mappa kiválasztása, fájlok beolvasása és kilistázása |
BÁ_on_convert_click() - Kijelölt fájl típusának ellenőrzése és a megfelelő konverzió indítása |
BÁ_pdf_to_word() - PDF dokumentum szövegének kinyerése és új DOCX fájl létrehozása |
BÁ_word_to_pdf() - DOCX dokumentum konvertálása PDF formátumba (Windows + Word szükséges) |
kill_word_processes() - Folyamatban lévő Microsoft Word lezárása a konverzió hibamentes futásához |


Összefoglalás

A program a dokumentumkezelés automatizálására alkalmas, egyszerű és gyors konverziós lehetőséget biztosít irodai, oktatási és személyes használatra. A GUI felület könnyen kezelhető, a függvények jól tagoltak, így a megoldás könnyen továbbfejleszthető.

