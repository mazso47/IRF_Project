HawaiiWeatherApp

Hawaii különböző városainak időjárását lehet elérni ezzel a programmal. Az adatok XML fájlokból kerülnek beolvasásra, maguk az XML fájlok pedig 
erről (https://w1.weather.gov/xml/current_obs/seek.php?state=hi&Find=Find) az oldalról kerülnek letöltésre. Van lehetőség ezen kívül az adatok frissítésére (XML fájlok újbóli letöltése), formázott excelbe exportálásra,
illetve csv-be írásra.

Ismert hiba:
Error: Couldn't process file abc.resx due to its being in the internet or Restricted zone or having the mark of the web on the file. Remove the mark of the web if you want to process these files
Fix: zip mappán jobb klikk -> tulajdonságok -> tiltás oldása jobb alul, majd kicsomagolás

Függvények:
setDecimalSeparator(): tizedesvesszőt pontra állítja
checkDir(): "xmlFiles" mappa létezését ellenőrzi, ha nem létezik, létrehozza
fillLists(): filenames és locations listákat tölti fel a letöltendő XML fájlnevekkel, valamint az azokhoz tartozó városnevekkel
updateData(): letölti az új XML fájlokat a filenames lista tartalma alapján
clearFiles(): üríti az xmlFiles mappa tartalmát
clearLinks(): üríti a links listát
getWeatherData(): a locationTextBox1 alapján kiválasztja a megfelelő XML fájlt, majd abból kiírja a megfelelő időjárás adatokat
getRecommendation(): bizonyos időjárásos adatok alapján eldönti, hogy ajánlott-e a kinti lét; itt jelenik meg az enumeráció használata
CreateExcel(), CreateTable(), GetCell(int x, int y): formázott excel létrehozása az összes XML fájl alapján
setTexts(): Resource1 alapján az egyes feliratok beállítása
textBoxAutoComplete(): a location lista alapján a textbox szövegkiegészítése


