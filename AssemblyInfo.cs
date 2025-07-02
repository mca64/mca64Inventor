using System.Reflection; // Importuje klasy do pobierania informacji o zestawie (assembly)
using System.Runtime.CompilerServices; // Importuje klasy do atrybutów kompilatora
using System.Runtime.InteropServices; // Importuje klasy do obs³ugi COM i GUID

//
// Informacje ogólne o zestawie s¹ kontrolowane przez poni¿szy zestaw atrybutów.
// Zmieñ wartoœci tych atrybutów, aby zmodyfikowaæ informacje powi¹zane z zestawem.
//
[assembly: AssemblyTitle("mca64Inventor AddIn Template")] // Tytu³ zestawu
[assembly: AssemblyDescription("Szablon AddIn dla Autodesk Inventor.")] // Opis zestawu
[assembly: AssemblyConfiguration("")] // Konfiguracja (np. Debug/Release)
[assembly: AssemblyCompany("YourCompany")] // Nazwa firmy
[assembly: AssemblyProduct("mca64Inventor AddIn")] // Nazwa produktu
[assembly: AssemblyCopyright("Copyright © 2024")] // Informacja o prawach autorskich
[assembly: AssemblyTrademark("")] // Znak towarowy
[assembly: AssemblyCulture("")] // Kultura (np. "pl-PL")

// Poni¿szy GUID jest identyfikatorem typbiblioteki, jeœli projekt jest ujawniany dla COM
[assembly: Guid("D0A4FCF1-B4D0-4824-894A-7FA419AF92F8")]

//
// Informacje o wersji zestawu sk³adaj¹ siê z nastêpuj¹cych czterech wartoœci:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// Mo¿esz okreœliæ wszystkie wartoœci lub u¿yæ domyœlnych numerów Build i Revision
// za pomoc¹ '*', jak pokazano poni¿ej:
[assembly: AssemblyVersion("1.0.*")]

//
// Aby podpisaæ zestaw, musisz okreœliæ klucz do u¿ycia. Zobacz dokumentacjê .NET Framework
// po wiêcej informacji o podpisywaniu zestawów.
//
// U¿yj poni¿szych atrybutów, aby kontrolowaæ, który klucz jest u¿ywany do podpisywania.
//
// Uwagi: 
//   (*) Jeœli nie okreœlono klucza, zestaw nie jest podpisany.
//   (*) KeyName odnosi siê do klucza zainstalowanego w Crypto Service Provider (CSP) na komputerze.
//   (*) KeyFile odnosi siê do pliku zawieraj¹cego klucz.
//   (*) Jeœli okreœlono zarówno KeyFile, jak i KeyName, wykonywane s¹ nastêpuj¹ce czynnoœci:
//       (1) Jeœli KeyName mo¿na znaleŸæ w CSP, ten klucz jest u¿ywany.
//       (2) Jeœli KeyName nie istnieje, a KeyFile istnieje, klucz z KeyFile jest instalowany w CSP i u¿ywany.
//   (*) Aby utworzyæ KeyFile, u¿yj narzêdzia sn.exe (Strong Name).
//       Podczas okreœlania KeyFile, lokalizacja powinna byæ wzglêdna wzglêdem katalogu wyjœciowego projektu,
//       np. %Project Directory%\obj\<configuration>. Jeœli plik KeyFile jest w katalogu projektu,
//       u¿yj [assembly: AssemblyKeyFile("..\\..\\mykey.snk")]
//   (*) Delay Signing to zaawansowana opcja - zobacz dokumentacjê .NET Framework po wiêcej informacji.
//
[assembly: AssemblyDelaySign(false)] // Czy opóŸniæ podpisywanie zestawu
[assembly: AssemblyKeyFile("")] // Œcie¿ka do pliku z kluczem (jeœli u¿ywasz podpisywania)
[assembly: AssemblyKeyName("")] // Nazwa klucza w CSP (jeœli u¿ywasz podpisywania)
