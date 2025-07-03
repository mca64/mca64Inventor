using System; // Importuje podstawowe klasy .NET, np. do obs³ugi typów i wyj¹tków
using System.Drawing; // Importuje klasy do obs³ugi grafiki i obrazów
using System.Drawing.Imaging; // Importuje klasy do obs³ugi formatów obrazów
using System.Runtime.InteropServices; // Importuje klasy do pracy z niezarz¹dzanym kodem (interop)

namespace mca64Inventor // Przestrzeñ nazw grupuj¹ca klasy dodatku
{
    /// <summary>
    /// Klasa pomocnicza do konwersji System.Drawing.Image na stdole.IPictureDisp,
    /// wymagany przez API Autodesk Inventor do ustawiania ikon przycisków.
    /// </summary>
    public static class PictureDispConverter
    {
        // Importuje funkcjê z biblioteki systemowej do tworzenia obiektów IPictureDisp
        [DllImport("OleAut32.dll", EntryPoint = "OleCreatePictureIndirect", ExactSpelling = true, PreserveSig = false)]
        private static extern stdole.IPictureDisp OleCreatePictureIndirect([In] ref PICTDESC pictdesc, ref Guid iid, [MarshalAs(UnmanagedType.Bool)] bool fOwn);

        // Przechowuje identyfikator typu IPictureDisp
        private static Guid iPictureDispGuid = typeof(stdole.IPictureDisp).GUID;

        // Struktura opisuj¹ca bitmapê dla funkcji OleCreatePictureIndirect
        [StructLayout(LayoutKind.Sequential)]
        private struct PICTDESC
        {
            internal int cbSizeOfStruct; // Rozmiar struktury
            internal int picType;        // Typ obrazka (1 = bitmapa)
            internal IntPtr hbitmap;     // Uchwyt do bitmapy
            internal IntPtr hpal;        // Paleta (nieu¿ywana)
            internal int unused;         // Nie u¿ywane
        }

        /// <summary>
        /// Konwertuje System.Drawing.Image na stdole.IPictureDisp (bitmapa do u¿ycia w Inventor API).
        /// </summary>
        /// <param name="image">Obraz do konwersji</param>
        /// <returns>Obiekt IPictureDisp do ustawienia jako ikona przycisku</returns>
        public static stdole.IPictureDisp ToIPictureDisp(Image image)
        {
            // Pobierz uchwyt do bitmapy z obrazu
            IntPtr hBitmap = ((Bitmap)image).GetHbitmap();
            // Przygotuj strukturê z danymi bitmapy
            var pictdesc = new PICTDESC
            {
                cbSizeOfStruct = Marshal.SizeOf(typeof(PICTDESC)),
                picType = 1, // PICTYPE_BITMAP
                hbitmap = hBitmap,
                hpal = IntPtr.Zero,
                unused = 0
            };
            // Utwórz i zwróæ IPictureDisp
            return OleCreatePictureIndirect(ref pictdesc, ref iPictureDispGuid, true);
        }

        /// <summary>
        /// Konwertuje stdole.IPictureDisp na System.Drawing.Image (Bitmap).
        /// </summary>
        public static Image ToImage(stdole.IPictureDisp pictureDisp)
        {
            if (pictureDisp == null) return null;
            // IPictureDisp.Handle to IntPtr do HBITMAP
            IntPtr hBitmap = (IntPtr)pictureDisp.Handle;
            if (hBitmap == IntPtr.Zero) return null;
            return Image.FromHbitmap(hBitmap);
        }
    }
}
