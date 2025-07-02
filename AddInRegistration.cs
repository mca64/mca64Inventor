using System; // Importuje podstawowe klasy .NET, np. do obs³ugi typów i wyj¹tków
using Microsoft.Win32; // Importuje klasy do pracy z rejestrem systemu Windows

namespace mca64Inventor // Przestrzeñ nazw grupuj¹ca powi¹zane klasy projektu
{
    /// <summary>
    /// Klasa AddInRegistration zawiera metody do rejestracji i wyrejestrowania dodatku (AddIn) Inventora w rejestrze Windows.
    /// Dziêki temu Inventor wie, ¿e istnieje taki dodatek i mo¿e go za³adowaæ.
    /// </summary>
    public class AddInRegistration
    {
        /// <summary>
        /// Rejestruje dodatek Inventor AddIn w rejestrze Windows.
        /// To umo¿liwia Inventorowi wykrycie i za³adowanie dodatku.
        /// </summary>
        /// <param name="t">Typ klasy dodatku (zwykle typeof(AddInServer))</param>
        public static void RegisterInventorAddIn(Type t)
        {
            // Tutaj mo¿na dodaæ kod, który wpisuje odpowiednie klucze do rejestru Windows,
            // aby Inventor widzia³ ten dodatek jako zarejestrowany COM AddIn.
            // Przyk³ad: ustawienie kluczy w HKEY_CLASSES_ROOT i HKEY_LOCAL_MACHINE.
        }

        /// <summary>
        /// Usuwa wpisy dodatku Inventor AddIn z rejestru Windows.
        /// Dziêki temu Inventor nie bêdzie ju¿ widzia³ tego dodatku.
        /// </summary>
        /// <param name="t">Typ klasy dodatku (zwykle typeof(AddInServer))</param>
        public static void UnregisterInventorAddIn(Type t)
        {
            // Tutaj mo¿na dodaæ kod, który usuwa odpowiednie klucze z rejestru Windows,
            // aby Inventor nie widzia³ ju¿ tego dodatku jako zarejestrowanego COM AddIn.
        }
    }
}
