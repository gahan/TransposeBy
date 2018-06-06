namespace TransposeBy
{
    public static class MyArrayExtentions
    {
        public static void Fill<T>(this T[,] oSource, T oDefaultValue)
        {
            // If the source array is null then just exit

            if (oSource == null)
            {
                return;
            }

            // Loop thru both axis of the source array adding in the default value

            for (int iRow = 0; iRow < oSource.GetLength(0); iRow++)
            {
                for (int iCol = 0; iCol < oSource.GetLength(1); iCol++)
                {
                    oSource[iRow, iCol] = oDefaultValue;
                }
            }
        }
    }
}
