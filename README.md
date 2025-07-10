# mca64Inventor

![mca64Inventor](https://github.com/user-attachments/assets/d9cf2800-d9fb-4cc3-8b25-092c38701fec)

## English

AddIn for Autodesk Inventor.

### Functionality
- Compiles to a DLL loaded by Autodesk Inventor as an AddIn.
- Adds a ribbon tab named "Astrus" with a panel named "Grawerowanie" (Engraving) and a button "Uruchom Grawerowanie" (Run Engraving).
- When the button is clicked, a form is displayed, which allows:
  - **Part Listing:** Displays a list of parts from the active assembly, including their names, generated thumbnails, and current engraving text (from the "Grawer" iProperty).
  - **Thumbnail Generation:** Generates visual thumbnails (screenshots) for all `.ipt` parts within the assembly and displays them in the UI. This process opens each part, captures a view, and then closes it.
  - **Engraving:** Processes all parts in the assembly that have a defined "Grawer" iProperty. It creates an extrusion feature (cut operation) named "Grawerowanie64" on a work plane, applying the text from the "Grawer" iProperty. Users can select the font and font size directly from the form, which are then applied to the engraving.
  - **IGES Export:** Exports engraved parts to the IGES format into a subfolder named `IGES` within the assembly's directory. (Note: The IGES export logic is present but may require further implementation for full functionality).
  - **Part Closing Option:** An optional setting to close part documents after they have been engraved.
  - **Logging:** Provides a log area within the form to display messages and errors during operations.
- The AddIn works with Autodesk Inventor and requires .NET Framework 4.8.
- The button uses icons: `icon16` (16x16 px) and `icon32` (32x32 px) from the project resources. It is recommended to add your own icons to the project resources for best appearance.

### How to use
1.  Build the project in Visual Studio (requires .NET Framework 4.8 and Autodesk Inventor references).
2.  Register the DLL using `regasm /codebase mca64Inventor.dll` (as administrator).
3.  Add an `.addin` manifest file to Inventor's AddIns folder.
    Example location: `C:\Users\<YourUserName>\AppData\Roaming\Autodesk\Inventor <Version>\Addins\`

#### Example `.addin` manifest file
Create `mca64Inventor.addin` with the following content:
```xml
<AddIn>
  <ClassId>{a5f3c3e7-7c33-46a2-b16a-42a4a4c25a17}</ClassId>
  <ClientId>{a5f3c3e7-7c33-46a2-b16a-42a4a4c25a17}</ClientId>
  <DisplayName>mca64Inventor AddIn</DisplayName>
  <Description>Minimalistic AddIn template for Inventor with engraving and thumbnail generation features.</Description>
  <Assembly>mca64Inventor.dll</Assembly>
  <LoadOnStartUp>1</LoadOnStartUp>
</AddIn>
```

### Development
- Extend your logic in the `mca64Inventor` namespace.
- Add your own buttons, panels, and event handling as needed.
- To change the button icon, add your own `icon16` and `icon32` PNG files to the project resources (Properties/Resources.resx or Resource1.resx).

---

## Polish / Polski

AddIn dla Autodesk Inventor.

### Funkcjonalność
- Kompiluje się do pliku DLL ładowanego przez Autodesk Inventor jako AddIn.
- Dodaje zakładkę na wstążce o nazwie "Astrus" z panelem "Grawerowanie" i przyciskiem "Uruchom Grawerowanie".
- Po kliknięciu przycisku, wyświetlany jest formularz, który umożliwia:
  - **Listowanie Części:** Wyświetla listę części z aktywnego złożenia, w tym ich nazwy, wygenerowane miniaturki i aktualny tekst grawerowania (z właściwości iProperty "Grawer").
  - **Generowanie Miniaturek:** Generuje wizualne miniaturki (zrzuty ekranu) dla wszystkich części `.ipt` w złożeniu i wyświetla je w interfejsie użytkownika. Proces ten otwiera każdą część, przechwytuje widok, a następnie ją zamyka.
  - **Grawerowanie:** Przetwarza wszystkie części w złożeniu, które mają zdefiniowaną właściwość iProperty "Grawer". Tworzy operację wyciągnięcia (cięcie) o nazwie "Grawerowanie64" na płaszczyźnie roboczej, stosując tekst z właściwości "Grawer". Użytkownicy mogą wybrać czcionkę i jej rozmiar bezpośrednio z formularza, które są następnie stosowane do grawerowania.
  - **Eksport do IGES:** Eksportuje grawerowane części do formatu IGES do podfolderu o nazwie `IGES` w katalogu złożenia. (Uwaga: Logika eksportu IGES jest obecna, ale może wymagać dalszej implementacji dla pełnej funkcjonalności).
  - **Opcja Zamykania Części:** Opcjonalne ustawienie do zamykania dokumentów części po ich wygrawerowaniu.
  - **Logowanie:** Zapewnia obszar logowania w formularzu do wyświetlania wiadomości i błędów podczas operacji.
- AddIn współpracuje z Autodesk Inventor i wymaga .NET Framework 4.8.
- Przycisk wykorzystuje ikony: `icon16` (16x16 px) i `icon32` (32x32 px) z zasobów projektu. Zaleca się dodanie własnych ikon do zasobów projektu dla lepszego wyglądu.

### Jak używać
1.  Zbuduj projekt w Visual Studio (wymaga .NET Framework 4.8 i referencji do Autodesk Inventor).
2.  Zarejestruj plik DLL za pomocą `regasm /codebase mca64Inventor.dll` (jako administrator).
3.  Dodaj plik manifestu `.addin` do folderu AddIns Inventora.
    Przykład lokalizacji: `C:\Users\<TwojaNazwaUżytkownika>\AppData\Roaming\Autodesk\Inventor <Wersja>\Addins\`

#### Przykład pliku manifestu `.addin`
Utwórz plik `mca64Inventor.addin` o następującej zawartości:
```xml
<AddIn>
  <ClassId>{a5f3c3e7-7c33-46a2-b16a-42a4a4c25a17}</ClassId>
  <ClientId>{a5f3c3e7-7c33-46a2-b16a-42a4a4c25a17}</ClientId>
  <DisplayName>mca64Inventor AddIn</DisplayName>
  <Description>Minimalistyczny szablon AddIn dla Inventora z funkcjami grawerowania i generowania miniaturek.</Description>
  <Assembly>mca64Inventor.dll</Assembly>
  <LoadOnStartUp>1</LoadOnStartUp>
</AddIn>
```

### Rozwój
- Rozszerzaj logikę w przestrzeni nazw `mca64Inventor`.
- Dodawaj własne przyciski, panele i obsługę zdarzeń według potrzeb.
- Aby zmienić ikonę przycisku, dodaj własne pliki PNG `icon16` i `icon32` do zasobów projektu (Properties/Resources.resx lub Resource1.resx).

---

**Template based on original code.**

```