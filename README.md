# Trello-organizing

A kód feladata, hogy megszámolja a Trello-ban lévő már elvégzett és elfogadott feladatokat, a megfelelő cég és feladat fontossági súlyozása alapján.

## Használat
A kód használata előtt ki kell tölteni az appsettings.json fájlt az adatbázis connection stringjével, a trello user secrettel, az API kulccsal és a táblázat azonosítójával.

```json
{
  "settings": {
    "connectionString": "CONNECTION_STRING_HELYE",
    "userSecret": "USER_SECRET_HELYE",
    "apiKey": "API_KEY_HELYE",
    "boardId": "TRELLO_BOARD_ID_HELYE"
  }
}
```
## Funkció
A kód működése során összehasonlítja az újonnan lekért adatokat a saját adatbázisával. Három lehetőség van:

- Ha a vizsgált kártya nem létezik az adatbázisban, létrehozza azt.
- Ha a kártya létezik az adatbázisban, de különbözik, frissíti az újonnan lekért adatokkal.
- Ha egy kártya prioritása megváltozik, frissíti az adatbázisban. Ha a kártya a legutóbbi ellenőrzés után el volt fogadva, és az aktuális ellenőrzéskor is el van fogadva, akkor a task adatbázisban szereplő legutóbbi módosítás dátuma szerint változtatja a számokat.
  > Ha egy kártyának frissítjük a labeljeit, akkor frissül a jelenlegi dátumra, tehát elveszik az információ, hogy melyik hónapban lett a task elfogadva, ezért az adatbázisban szereplő legutolsó dátumot figyeli a kód az elszámolásnál.
## Tesztesetek

### Elfogadott task>
- [x] Súlyozás frissítése (kisebből nagyobb, nagyobból kisebb)
- [x] Speciális label változások ellenőrzése (lekerül az összes label a taskról, új ismeretlen label kerül rá)
- [x] Task újranyitása (új súlyozással/súlyozás nélkül)
### Nem elfogadott task>
- [x] Súlyozás frissítése (kisebből nagyobb, nagyobból kisebb)
- [x] Task elfogadása (új súlyozással/súlyozás nélkül)
### Egyéb>
- [x] Több label esetén a legmagasabb prioritásút veszi
