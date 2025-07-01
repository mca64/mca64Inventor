# mca64Inventor AddIn Template

Ten projekt to minimalistyczny szablon dodatku (AddIn) dla Autodesk Inventor w C# (.NET Framework 4.8).

## Funkcjonalnoœæ
- Kompiluje siê do DLL ³adowanej przez Inventora.
- Dodaje zak³adkê (ribbon tab) o nazwie „mca64launcher”.
- Dodaje przycisk, który po klikniêciu wyœwietla prost¹ formê.

## Jak u¿ywaæ
1. Zbuduj projekt w Visual Studio (wymagany .NET Framework 4.8, referencje do bibliotek Autodesk Inventor).
2. Zarejestruj DLL poleceniem `regasm /codebase mca64Inventor.dll` (jako administrator).
3. Dodaj plik manifestu `.addin` do folderu AddIns Inventora (patrz poni¿ej).
4. Uruchom Inventora – AddIn pojawi siê na liœcie dodatków.

## Przyk³adowy plik manifestu `.addin`
Utwórz plik `mca64Inventor.addin` w folderze:
`C:\Users\<TwojaNazwa>\AppData\Roaming\Autodesk\Inventor <Wersja>\Addins\`
<AddIn>
  <ClassId>{963308E2-D850-466D-A1C5-503A2E171552}</ClassId>
  <ClientId>{963308E2-D850-466D-A1C5-503A2E171552}</ClientId>
  <DisplayName>mca64Inventor AddIn</DisplayName>
  <Description>Minimalistyczny szablon AddIn dla Inventora</Description>
  <Assembly>mca64Inventor.dll</Assembly>
  <LoadOnStartUp>1</LoadOnStartUp>
</AddIn>
## Rozwijanie
- Rozwijaj w³asn¹ logikê w klasach w przestrzeni nazw `mca64Inventor`.
- Dodawaj w³asne przyciski, panele, obs³ugê zdarzeñ.

---
**Szablon przygotowany na bazie oryginalnego kodu.**
