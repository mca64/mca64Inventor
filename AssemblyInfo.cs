using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// Ogólne informacje o zestawie są kontrolowane poprzez następujący
// zestaw atrybutów. Zmień wartości tych atrybutów, aby zmodyfikować informacje
// skojarzone z zestawem.
[assembly: AssemblyTitle("mca64Inventor AddIn Template")] // Tytuł zestawu (nazwa dodatku)
[assembly: AssemblyDescription("AddIn template for Autodesk Inventor.")] // Opis zestawu
[assembly: AssemblyConfiguration("")] // Konfiguracja zestawu (np. Debug/Release)
[assembly: AssemblyCompany("YourCompany")] // Nazwa firmy
[assembly: AssemblyProduct("mca64Inventor AddIn")] // Nazwa produktu
[assembly: AssemblyCopyright("Copyright © 2024")] // Informacje o prawach autorskich
[assembly: AssemblyTrademark("")] // Znak towarowy
[assembly: AssemblyCulture("")] // Kultura zestawu (np. en-US, pl-PL)

// Ustawienie ComVisible na false sprawia, że typy w tym zestawie nie są widoczne
// dla komponentów COM. Jeśli potrzebujesz dostępu do typu w tym zestawie z
// COM, ustaw atrybut ComVisible na true dla tego typu.
[assembly: ComVisible(true)]

// Poniższy GUID jest identyfikatorem typelib, jeśli ten projekt jest udostępniony dla COM
[assembly: Guid("D0A4FCF1-B4D0-4824-894A-7FA419AF92F8")] // Unikalny identyfikator GUID dla tego zestawu

// Informacje o wersji zestawu składają się z następujących czterech wartości:
//
//      Wersja główna (Major Version)
//      Wersja pomocnicza (Minor Version)
//      Numer kompilacji (Build Number)
//      Poprawka (Revision)
//
// Możesz określić wszystkie wartości lub pozostawić domyślne numery kompilacji i poprawki
// używając '*' jak pokazano poniżej:
[assembly: AssemblyVersion("1.0.0.*")] // Wersja zestawu (Major.Minor.Build.Revision)

// Aby podpisać zestaw, musisz określić klucz do użycia. Zapoznaj się z dokumentacją Microsoft .NET Framework, aby uzyskać więcej informacji na temat podpisywania zestawów.
// Użyj poniższych atrybutów, aby kontrolować, który klucz jest używany do podpisywania.
// Uwagi:
//   (*) Jeśli nie określono klucza, zestaw nie jest podpisany.
//   (*) KeyName odnosi się do klucza zainstalowanego w dostawcy usług kryptograficznych (CSP) na Twojej maszynie.
//   (*) KeyFile odnosi się do pliku zawierającego klucz.
//   (*) Jeśli określono wartości KeyFile i KeyName, dzieje się następujące:
//       (1) Jeśli KeyName można znaleźć w CSP, używany jest ten klucz.
//       (2) Jeśli KeyName nie istnieje, a KeyFile istnieje, klucz z KeyFile jest instalowany w CSP i używany.
//   (*) Aby utworzyć KeyFile, użyj narzędzia sn.exe (Strong Name).
//       Podczas określania KeyFile, lokalizacja pliku powinna być względna w stosunku do katalogu wyjściowego projektu,
//       np. %Project Directory%\obj\<configuration>. Jeśli KeyFile znajduje się w katalogu projektu,
//       użyj [assembly: AssemblyKeyFile("..\\..\\mykey.snk")]
//   (*) Opóźnione podpisywanie (Delay Signing) to zaawansowana opcja - więcej informacji znajdziesz w dokumentacji .NET Framework.
[assembly: AssemblyDelaySign(false)] // Czy podpisywanie zestawu ma być opóźnione
[assembly: AssemblyKeyFile("")] // Ścieżka do pliku klucza podpisywania
[assembly: AssemblyKeyName("")] // Nazwa klucza podpisywania w CSP
