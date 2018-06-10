# TransposeBy - Extending Excel with C# and Excel-DNA
A small demonstration project on how to create an Excel User Defined Function in C# and Excel-DNA

**Eddie Gahan**

**June 2018**

**Intro**

At some point I think all of us have received an Excel workbook filled with badly formatted data that we somehow have to transform into a neatly formatted report.  Several times I&#39;ve gotten a column of data that contains records comprising multiple fields like this:

![Excel Data Column](https://github.com/gahan/TransposeBy/blob/master/images/Excel%20Data%20Column.png "Excel Data Column")

when what I really need is this:

![Excel Transposed Data](https://github.com/gahan/TransposeBy/blob/master/images/Excel%20Transposed%20Data.png "Excel Transposed Data")

Excel does contain an in-built function called Transpose but it&#39;s quite limited in that it just flips a row to a column and vice versa.  As I don&#39;t want to have to reformat the data by hand, preferring an automated solution, I have a couple of options open to me.  I can create a VBA based macro or an Office Add-In, but I don&#39;t really want to go that route.  Macros can be difficult to add to every workbook I use, and the Add-In will require me to add a button to the tool ribbon which I want to avoid.  The last option is to create a custom User Defined Function that will be available to me whenever I use Excel, which as it turns out, is very easy to do in a C# project thanks to an open source framework called Excel-DNA.

        [Excel-DNA](https://excel-dna.net) makes the creation of a .xll, the Excel add-in format (It&#39;s similar to a .dll but specific to Excel) extremely simple.  There are commercial solutions available, but you&#39;ll probably only ever need this framework.

**Set-up**

The function I ended up creating is called **TransposeBy** and I&#39;ll be using the project code to illustrate this article.  That source code can be found on GitHub in [this repository](https://github.com/gahan/TransposeBy).  Creation of the project is pretty easy so I&#39;m going to assume you can handle the following steps:

1. Create a C# class library project in Visual Studio.
2. Using nuget, install the following libraries:

        ExcelDNA.AddIn
        ExcelDNA.Integration
        ExcelDNA.Intellisense (Optional)

And now you&#39;re ready to go.

**The Solution**

So what does this function need to do to be able to accomplish the outcome I need.  It&#39;ll have to:

- Accept an Excel Range, either a column or a row, specifying the source data I wish to transpose.
- Optionally, accept a Boolean flag to switch the transposition from being by column to being by row.
- Exit if the function was not called as an Array Formula in Excel (If you&#39;re unfamiliar with Array Formula, you can find a primer on them [here](https://support.office.com/en-ie/article/guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7))
- Perform validation on the source data.
- Initialize and fill an output object array.

**The Code**

In my sample project you&#39;ll see a file called TransposeByUDF.cs.  This is just a standard class file which inherits the XlCall class in ExcelDNA.Integration and references the Namespace ExcelDna.Integration.  I&#39;ve also referenced the Namespace ExcelDna.Intellisense so that users can get some additional information on the function in Excel but it&#39;s not necessary for a simple function.

````csharp
using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace TransposeBy
{
    public class TransposeByUDF : XlCall
    {
````

The next step is to create a static function that returns an object and to attach the following directive to the function.  The ExcelFunction directive contains a number of options but I&#39;ve only used the description option to give Excel users some information on the function.  I&#39;m returning an object as my function will be filling in a cell range but you can return strings, numbers etc. as appropriate.  In the function declaration I&#39;ve also added a directive to the two parameters to give readable parameter names and descriptions to the Excel users.

````csharp
[ExcelFunction(Description = "Transpose a range of values breaking every n number of rows/columns.")]
public static object TransposeBy(
                                    [ExcelArgument(Name = "SourceData", Description = "The range of cells to be transposed.")] object oSource,
                                    [ExcelArgument(Name = "ByRow", Description = "Optional flag to force transposing vertically insted of the horizontal default.")] [Optional] bool bByRow
                                )
{
````

Now that the function is all set up, you can use the many objects that Excel-DNA exposes to interact with the parent Excel application.  Documentation in the Excel-DNA project isn&#39;t perfect but there are extensive example projects and an archive of the Google Groups forum that can be searched.  For my function I first needed to get the range of cells that was selected when the function was called and this was made available to me thru the XlCall.Excel object.  This object accepts an Enum as a parameter that allows you to define what type of information you want to retrieve.  In my case I&#39;ve used the value xlfCaller to reference the destination cell range.

````csharp
// Create the reference to the output array of cells

var oCaller = Excel(xlfCaller) as ExcelReference;
if (oCaller == null)
{
    return new object[0, 0];
}
````

If the range is null then I just return a blank object but I can also return Excel specific errors like #REF or #VALUE.  In the next couple of lines in the function I test to see if the destination range that was passed is just a single row or column of data and that the function was called as an array function.  If it fails either of those tests then I throw back some Excel specific errors.  I also wrapped the entire function in a try..catch block that if triggered throws back a #REF error to Excel.

````csharp
// Test that the source is a single column or row of values, the destination is an array function etc.

if (oCaller.RowFirst == oCaller.RowLast && oCaller.ColumnFirst == oCaller.ColumnLast) { return ExcelError.ExcelErrorRef; }  // Formula has not been entered as an Array formula
if (((System.Array)oSource).GetLength(0) > 1 && ((System.Array)oSource).GetLength(1) > 1) { return ExcelError.ExcelErrorValue; } // Source data is not a single column or row
````

As the function is called from Excel as an array formula I have to return an array object that matches the dimensions of the range of cells that were selected when the function was called.  Whilst the exact dimensions aren&#39;t available in the ExcelReference object I created, the top and bottom row and column refences are so a simple calculation can be done to get the required sizes.

````csharp
// Initialise the output result array

object[,] oResult = new object[(oCaller.RowLast - oCaller.RowFirst) + 1, (oCaller.ColumnLast - oCaller.ColumnFirst) + 1];
oResult.Fill("");
````

To make life easier for me I added a separate class file to the project so I could extend object arrays to have a method **.Fill()** that auto-initializes each element in the array with a specified parameter.  You should auto-initialize with a blank value to avoid null values being displayed in your cell range and you can find that extension in the MyArrayExtentions.cs file in the project.

The rest of the code in my sample function just deals with transposing the single row or column of data so that each element is filled in from left to right, line by line in to the output array.  It&#39;s pretty simple code and doesn&#39;t relate directly to Excel-DNA so I&#39;m not going to go into detail of it here.  The only important line is the one that returns, in my case, the object array at the end.  As mentioned, you can return a string and a numeric as appropriate to your function.

**Compiling and Testing**

Adding the ExcelDNA.AddIn library from NuGet to your project results in a few changes to the solution.  Firstly, you&#39;ll notice a file with the extension .dna in your project.  In my sample project it&#39;s called TransposeBy-AddIn.dna and it acts as a configuration file when distributing your custom function.  You probably won&#39;t need to make any changes to it unless you&#39;re using the ExcelDNA.Intellisense library in which case you&#39;ll have to add a line like I did at line 4 in the following code:

````xml
<DnaLibrary Name="TransposeBy Add-In" RuntimeVersion="v4.0">
  <ExternalLibrary Path="TransposeBy.dll" ExplicitExports="false" LoadFromBytes="true" Pack="true" />

  <Reference Path="ExcelDna.IntelliSense.dll" Pack="true" />
  
  <!-- 
       The RuntimeVersion attribute above allows two settings:
       * RuntimeVersion="v2.0" - for .NET 2.0, 3.0 and 3.5
       * RuntimeVersion="v4.0" - for .NET 4 and 4.5

       Additional referenced assemblies can be specified by adding 'Reference' tags. 
       These libraries will not be examined and registered with Excel as add-in libraries, 
       but will be packed into the -packed.xll file and loaded at runtime as needed.
       For example:
       
       <Reference Path="Another.Library.dll" Pack="true" />
  
       Excel-DNA also allows the xml for ribbon UI extensions to be specified in the .dna file.
       See the main Excel-DNA site at http://excel-dna.net for downloads of the full distribution.
  -->

</DnaLibrary>
````

The second major change you&#39;ll probably see is that the Debug property page in the projects properties now reference your installation of Excel (if present) with the .xll of your project referenced as a command line parameter.

![Debug Property Panel](https://github.com/gahan/TransposeBy/blob/master/images/Debug%20Property%20Page.PNG "Debug Property Panel")

If you have Excel installed, and I recommend that you do if you&#39;re developing add-ins, you&#39;ll be able to debug your custom function just by hitting F5.  Excel will start with your custom function loaded and by using breakpoints you can step thru your code as you develop and debug.

**Deploying**

Distributing your custom function can be done by creating an installer and the Excel-DNA project contains a template for creating a WiX-based installer.  This can be found [here](https://github.com/Excel-DNA/WiXInstaller).  If you just want to quickly install it into your own installation of Excel you can do this by opening the Excel Add-Ins section in Excel Options and browsing to where you stored the .xll file.

![Excel AddIn Options Panel](https://github.com/gahan/TransposeBy/blob/master/images/Excel%20AddIn%20Options%20Panel.PNG "Excel AddIn Options Panel")