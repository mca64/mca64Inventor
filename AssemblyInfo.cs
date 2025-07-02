using System.Reflection; // Importuje klasy do pobierania informacji o zestawie (assembly)
using System.Runtime.CompilerServices; // Importuje klasy do atrybut�w kompilatora
using System.Runtime.InteropServices; // Importuje klasy do obs�ugi COM i GUID

//
// Informacje og�lne o zestawie s� kontrolowane przez poni�szy zestaw atrybut�w.
// Zmie� warto�ci tych atrybut�w, aby zmodyfikowa� informacje powi�zane z zestawem.
//
[assembly: AssemblyTitle("mca64Inventor AddIn Template")] // Tytu� zestawu
[assembly: AssemblyDescription("Szablon AddIn dla Autodesk Inventor.")] // Opis zestawu
[assembly: AssemblyConfiguration("")] // Konfiguracja (np. Debug/Release)
[assembly: AssemblyCompany("YourCompany")] // Nazwa firmy
[assembly: AssemblyProduct("mca64Inventor AddIn")] // Nazwa produktu
[assembly: AssemblyCopyright("Copyright � 2024")] // Informacja o prawach autorskich
[assembly: AssemblyTrademark("")] // Znak towarowy
[assembly: AssemblyCulture("")] // Kultura (np. "pl-PL")

// Poni�szy GUID jest identyfikatorem typbiblioteki, je�li projekt jest ujawniany dla COM
[assembly: Guid("D0A4FCF1-B4D0-4824-894A-7FA419AF92F8")]

//
// Informacje o wersji zestawu sk�adaj� si� z nast�puj�cych czterech warto�ci:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// Mo�esz okre�li� wszystkie warto�ci lub u�y� domy�lnych numer�w Build i Revision
// za pomoc� '*', jak pokazano poni�ej:
[assembly: AssemblyVersion("1.0.*")]

//
// Aby podpisa� zestaw, musisz okre�li� klucz do u�ycia. Zobacz dokumentacj� .NET Framework
// po wi�cej informacji o podpisywaniu zestaw�w.
//
// U�yj poni�szych atrybut�w, aby kontrolowa�, kt�ry klucz jest u�ywany do podpisywania.
//
// Uwagi: 
//   (*) Je�li nie okre�lono klucza, zestaw nie jest podpisany.
//   (*) KeyName odnosi si� do klucza zainstalowanego w Crypto Service Provider (CSP) na komputerze.
//   (*) KeyFile odnosi si� do pliku zawieraj�cego klucz.
//   (*) Je�li okre�lono zar�wno KeyFile, jak i KeyName, wykonywane s� nast�puj�ce czynno�ci:
//       (1) Je�li KeyName mo�na znale�� w CSP, ten klucz jest u�ywany.
//       (2) Je�li KeyName nie istnieje, a KeyFile istnieje, klucz z KeyFile jest instalowany w CSP i u�ywany.
//   (*) Aby utworzy� KeyFile, u�yj narz�dzia sn.exe (Strong Name).
//       Podczas okre�lania KeyFile, lokalizacja powinna by� wzgl�dna wzgl�dem katalogu wyj�ciowego projektu,
//       np. %Project Directory%\obj\<configuration>. Je�li plik KeyFile jest w katalogu projektu,
//       u�yj [assembly: AssemblyKeyFile("..\\..\\mykey.snk")]
//   (*) Delay Signing to zaawansowana opcja - zobacz dokumentacj� .NET Framework po wi�cej informacji.
//
[assembly: AssemblyDelaySign(false)] // Czy op�ni� podpisywanie zestawu
[assembly: AssemblyKeyFile("")] // �cie�ka do pliku z kluczem (je�li u�ywasz podpisywania)
[assembly: AssemblyKeyName("")] // Nazwa klucza w CSP (je�li u�ywasz podpisywania)
