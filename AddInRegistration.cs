using System; // Importuje podstawowe klasy .NET, np. do obs�ugi typ�w i wyj�tk�w
using Microsoft.Win32; // Importuje klasy do pracy z rejestrem systemu Windows

namespace mca64Inventor // Przestrze� nazw grupuj�ca powi�zane klasy projektu
{
    /// <summary>
    /// Klasa AddInRegistration zawiera metody do rejestracji i wyrejestrowania dodatku (AddIn) Inventora w rejestrze Windows.
    /// Dzi�ki temu Inventor wie, �e istnieje taki dodatek i mo�e go za�adowa�.
    /// </summary>
    public class AddInRegistration
    {
        /// <summary>
        /// Rejestruje dodatek Inventor AddIn w rejestrze Windows.
        /// To umo�liwia Inventorowi wykrycie i za�adowanie dodatku.
        /// </summary>
        /// <param name="t">Typ klasy dodatku (zwykle typeof(AddInServer))</param>
        public static void RegisterInventorAddIn(Type t)
        {
            // Tutaj mo�na doda� kod, kt�ry wpisuje odpowiednie klucze do rejestru Windows,
            // aby Inventor widzia� ten dodatek jako zarejestrowany COM AddIn.
            // Przyk�ad: ustawienie kluczy w HKEY_CLASSES_ROOT i HKEY_LOCAL_MACHINE.
        }

        /// <summary>
        /// Usuwa wpisy dodatku Inventor AddIn z rejestru Windows.
        /// Dzi�ki temu Inventor nie b�dzie ju� widzia� tego dodatku.
        /// </summary>
        /// <param name="t">Typ klasy dodatku (zwykle typeof(AddInServer))</param>
        public static void UnregisterInventorAddIn(Type t)
        {
            // Tutaj mo�na doda� kod, kt�ry usuwa odpowiednie klucze z rejestru Windows,
            // aby Inventor nie widzia� ju� tego dodatku jako zarejestrowanego COM AddIn.
        }
    }
}
