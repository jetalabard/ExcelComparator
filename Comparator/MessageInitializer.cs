
namespace Comparator
{
    public static class MessageInitializer
    {
        public static string Result { get; private set; }

        public static void Init(int i, int columnIndexFile1, int columnIndexFile2, string CurrentColumn)
        {
            Result += "La cellule ligne " + i + " colonne " + columnIndexFile1 + " (" + CurrentColumn + ") est différente par rapport à la colonne " + columnIndexFile2;
        }

    }
}
