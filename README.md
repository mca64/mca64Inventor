# mca64Inventor AddIn

## Functionality
- Compiles to a DLL loaded by Autodesk Inventor as an AddIn.
- Adds a ribbon tab named "mca64Inventor" with a panel and a button "Grawerowanie".
- When the button is clicked, a form is shown. After confirmation, the add-in processes the active assembly:
  - Engraves all parts (with user iProperties property "Grawer" as text) in the assembly.
  - Exports engraved parts to IGES format in a subfolder named IGES in the assembly directory.
  - Generates thumbnails (screenshots) for all .ipt parts in the assembly and displays them in the UI.
- The add-in works with Autodesk Inventor and requires .NET Framework 4.8.
- The button uses icons: `icon16` (16x16 px) and `icon32` (32x32 px) from the project resources. You should add your own icons to the project resources for best appearance.

## How to use
1. Build the project in Visual Studio (requires .NET Framework 4.8 and Autodesk Inventor references).
2. Register the DLL using `regasm /codebase mca64Inventor.dll` (as administrator).
3. Add a `.addin` manifest file to Inventor's AddIns folder (see below).
4. Start Inventor – the AddIn will appear in the Add-Ins list.

### Example `.addin` manifest file
Create `mca64Inventor.addin` in:
`C:\Users\<YourName>\AppData\Roaming\Autodesk\Inventor <Version>\Addins\`<AddIn>
  <ClassId>{963308E2-D850-466D-A1C5-503A2E171552}</ClassId>
  <ClientId>{963308E2-D850-466D-A1C5-503A2E171552}</ClientId>
  <DisplayName>mca64Inventor AddIn</DisplayName>
  <Description>Minimalistic AddIn template for Inventor</Description>
  <Assembly>mca64Inventor.dll</Assembly>
  <LoadOnStartUp>1</LoadOnStartUp>
</AddIn>
## Development
- Extend your logic in the `mca64Inventor` namespace.
- Add your own buttons, panels, and event handling as needed.
- To change the button icon, add your own `icon16` and `icon32` PNG files to the project resources (Properties/Resources.resx or Resource1.resx).

---

**Template based on original code.**
