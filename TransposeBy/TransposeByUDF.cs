using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace TransposeBy
{
    public class TransposeByUDF : XlCall
    {
        [ExcelFunction(Description = "Transpose a range of values breaking every n number of rows/columns.")]
        public static object TransposeBy(
                                            [ExcelArgument(Name = "SourceData", Description = "The range of cells to be transposed.")] object oSource,
                                            [ExcelArgument(Name = "ByRow", Description = "Optional flag to force transposing vertically insted of the horizontal default.")] [Optional] bool bByRow
                                        )
        {
            try
            {
                // Create the reference to the output array of cells

                var oCaller = Excel(xlfCaller) as ExcelReference;
                if (oCaller == null)
                {
                    return new object[0, 0];
                }

                // Test that the source is a single column or row of values, the destination is an array function etc.

                if (oCaller.RowFirst == oCaller.RowLast && oCaller.ColumnFirst == oCaller.ColumnLast) { return ExcelError.ExcelErrorRef; }  // Formula has not been entered as an Array formula
                if (((System.Array)oSource).GetLength(0) > 1 && ((System.Array)oSource).GetLength(1) > 1) { return ExcelError.ExcelErrorValue; } // Source data is not a single column or row

                // Initialise the output result array

                object[,] oResult = new object[(oCaller.RowLast - oCaller.RowFirst) + 1, (oCaller.ColumnLast - oCaller.ColumnFirst) + 1];
                oResult.Fill("");

                // Fill the output array

                int iRow = 0;
                int iCol = 0;

                // Flag if the source data is a row instead of a column

                bool bIsRow = (((System.Array)oSource).GetLength(1) > 1 ? true : false);

                // Loop through the source range of cells

                for (int liLoop = 0; liLoop < (((System.Array)oSource).Length < ((System.Array)oResult).Length ? ((System.Array)oSource).Length : ((System.Array)oResult).Length); liLoop++)
                {
                    // Add the source value to the output array 

                    oResult[iRow, iCol] = ((object[,])oSource)[(bIsRow ? 0 : liLoop), (bIsRow ? liLoop : 0)];

                    // If transposing by row then increase the row counter.  If not, the column counter

                    if (bByRow) iRow++; else iCol++;

                    // If the last row or column is reached then move to the next row or column as appriopiate

                    if (iCol == (oCaller.ColumnLast - oCaller.ColumnFirst) + 1) { iRow++; iCol = 0; }
                    if (iRow == (oCaller.RowLast - oCaller.RowFirst) + 1) { iCol++; iRow = 0; }
                }

                // Return the object array to Excel

                return oResult;
            }
            catch (SystemException oError)
            {
                return ExcelError.ExcelErrorRef;
            }
        }
    }
}
