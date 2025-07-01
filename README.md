# mca64Inventor AddIn Template

This project is a minimalist AddIn template for Autodesk Inventor in C# (.NET Framework 4.8).

## Features
- Builds to a DLL loaded by Inventor.
- Adds a ribbon tab named "mca64Inventor".
- Adds a button that displays a simple form when clicked.

## How to use
1. Build the project in Visual Studio (requires .NET Framework 4.8, references to Autodesk Inventor libraries).
2. Register the DLL with `regasm /codebase mca64Inventor.dll` (as administrator).
3. Add a `.addin` manifest file to the Inventor AddIns folder (see below).
4. Start Inventor – the AddIn will appear in the AddIns list.

## Example `.addin` manifest file
Create a file `mca64Inventor.addin` in the folder:
`C:\Users\<YourName>\AppData\Roaming\Autodesk\Inventor <Version>\Addins\`
<AddIn>
  <ClassId>{963308E2-D850-466D-A1C5-503A2E171552}</ClassId>
  <ClientId>{963308E2-D850-466D-A1C5-503A2E171552}</ClientId>
  <DisplayName>mca64Inventor AddIn</DisplayName>
  <Description>Minimalist AddIn template for Inventor</Description>
  <Assembly>mca64Inventor.dll</Assembly>
  <LoadOnStartUp>1</LoadOnStartUp>
</AddIn>
## Development
- Add your own logic in classes in the `mca64Inventor` namespace.
- Add your own buttons, panels, event handling.

---

# mca64Inventor AddIn Template

Ten projekt to minimalistyczny szablon dodatku (AddIn) dla Autodesk Inventor w C# (.NET Framework 4.8).

## Funkcjonalność
- Kompiluje się do DLL ładowanej przez Inventora.
- Dodaje zakładkę (ribbon tab) o nazwie „mca64Inventor”.
- Dodaje przycisk, który po kliknięciu wyświetla prostą formę.

## Jak używać
1. Zbuduj projekt w Visual Studio (wymagany .NET Framework 4.8, referencje do bibliotek Autodesk Inventor).
2. Zarejestruj DLL poleceniem `regasm /codebase mca64Inventor.dll` (jako administrator).
3. Dodaj plik manifestu `.addin` do folderu AddIns Inventora (patrz poniżej).
4. Uruchom Inventora – AddIn pojawi się na liście dodatków.

## Przykładowy plik manifestu `.addin`
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
- Rozwijaj własną logikę w klasach w przestrzeni nazw `mca64Inventor`.
- Dodawaj własne przyciski, panele, obsługę zdarzeń.

---
**Szablon przygotowany na bazie oryginalnego kodu.**
