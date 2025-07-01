# mca64Inventor AddIn Template

Ten projekt to minimalistyczny szablon dodatku (AddIn) dla Autodesk Inventor w C# (.NET Framework 4.8).

## Funkcjonalno��
- Kompiluje si� do DLL �adowanej przez Inventora.
- Dodaje zak�adk� (ribbon tab) o nazwie �mca64launcher�.
- Dodaje przycisk, kt�ry po klikni�ciu wy�wietla prost� form�.

## Jak u�ywa�
1. Zbuduj projekt w Visual Studio (wymagany .NET Framework 4.8, referencje do bibliotek Autodesk Inventor).
2. Zarejestruj DLL poleceniem `regasm /codebase mca64Inventor.dll` (jako administrator).
3. Dodaj plik manifestu `.addin` do folderu AddIns Inventora (patrz poni�ej).
4. Uruchom Inventora � AddIn pojawi si� na li�cie dodatk�w.

## Przyk�adowy plik manifestu `.addin`
Utw�rz plik `mca64Inventor.addin` w folderze:
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
- Rozwijaj w�asn� logik� w klasach w przestrzeni nazw `mca64Inventor`.
- Dodawaj w�asne przyciski, panele, obs�ug� zdarze�.

---
**Szablon przygotowany na bazie oryginalnego kodu.**
