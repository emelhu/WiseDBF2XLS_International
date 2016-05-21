using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Linq.Expressions;

namespace WiseDBF2XLS_International
{
  using eMeL;
  using ExcelLibrary.SpreadSheet;
  using NDbfReaderEx;

  class Program
  {
    static void Main(string[] args)
    {
      Console.BackgroundColor = ConsoleColor.DarkBlue;
      Console.ForegroundColor = ConsoleColor.Yellow;
      Console.WriteLine("**** Convert DBF to XLS with localization ****  (c) eMeL [www.emel.hu] **** FreeWare ****");
      Console.ResetColor();
      Console.WriteLine();

      if ((args.Length < 1) || (args.Length > 2))
      {
        Environment.ExitCode = 3;
        DispUsage();
      }
      else
      {
        try
        {
          Environment.ExitCode = 2;                                         // parameter error or file not DBF
          List<DbfItem> dbfList = GetDbfList(args);

          // DBF header multiple read, but guaranted all parameter/files suitable before start first conversion

          foreach (var dbfItem in dbfList)
          {
            Environment.ExitCode = 1;                                       // DBF read or XLS write error 
            ConvertDbf2Xls(dbfItem.fileName, dbfItem.codepage);
          }

          Environment.ExitCode = 0;                                         // Success
        }
        catch (Exception e)
        {
          Console.WriteLine();
          Console.BackgroundColor = ConsoleColor.Yellow;
          Console.ForegroundColor = ConsoleColor.Red;
          Console.WriteLine(e.Message);
          Console.ResetColor();
          Console.WriteLine();
          DispUsage();
        }        
      }

      #if DEBUG
      Console.ResetColor();
      Console.WriteLine("Press Enter to close...");
      Console.ReadLine();
      #endif
    }

    private static void ConvertDbf2Xls(string dbfFileName, int codepage)
    {
      string xlsFileName = Path.Combine(Path.ChangeExtension(dbfFileName, ".xls"));

      Console.Write("Convert '{0}' to XLS... ", dbfFileName);
      int left  = Console.CursorLeft;
      int top   = Console.CursorTop;
      int count = 0;

      Workbook  wb = new Workbook(); 
      Worksheet ws = new Worksheet("DBF data"); 

      using (DbfTable table = DbfTable.Open(dbfFileName, Encoding.GetEncoding(codepage)))
      {
        foreach (DbfRow row in table)
        {
          int columnIx = 0;

          if (count == 0)
          { // write excel header            
            foreach (var column in row.columns)
            {
              ws.Cells[0, columnIx] = new Cell(column.name); 
              columnIx++;
            }

            columnIx = 0;

            count++;
          }


          foreach (var column in row.columns)
          {
            switch (column.dbfType)
            {
              case NativeColumnType.Char:
                ws.Cells[count, columnIx] = new Cell(row.GetString(column)); 
                break;
              case NativeColumnType.Memo:
                ws.Cells[count, columnIx] = new Cell(row.GetString(column)); 
                break;
              case NativeColumnType.Date:
                ws.Cells[count, columnIx] = new Cell(row.GetDate(column)); 
                break;
              case NativeColumnType.Long:
                ws.Cells[count, columnIx] = new Cell(row.GetInt32(column)); 
                break;
              case NativeColumnType.Logical:
                ws.Cells[count, columnIx] = new Cell(row.GetBoolean(column)); 
                break;
              case NativeColumnType.Numeric:
                if ((column.size < 10) && (column.dec == 0))
                { // Reduce to integer (32 bit)
                 ws.Cells[count, columnIx] = new Cell((Int32)row.GetDecimal(column));
                }
                else if ((column.size < 19) && (column.dec == 0))
                { // Reduce to long (64 bit)
                  ws.Cells[count, columnIx] = new Cell((Int64)row.GetDecimal(column));
                }
                else
                {
                  ws.Cells[count, columnIx] = new Cell(row.GetDecimal(column));
                }

                break;
              case NativeColumnType.Float:
                ws.Cells[count, columnIx] = new Cell(row.GetDecimal(column)); 
                break;
              case NativeColumnType.Double:
                ws.Cells[count, columnIx] = new Cell(row.GetDouble(column));                
                break;
              default:
                throw new Exception("Invalid 'dbfType' at '" + column.name + "' column!");
            }

            columnIx++;
          }


          count++;
          Console.SetCursorPosition(left, top);
          Console.Write(count.ToString());
        }

        Console.SetCursorPosition(left, top);
        Console.WriteLine("-OK- ");

        wb.Worksheets.Add(ws); 
        wb.Save(xlsFileName);
        wb = null;
      }
    }

    static void DispUsage()
    {
      var encodings = Encoding.GetEncodings().Select(e => e.CodePage.ToString()).ToList();      

      Console.ResetColor();
      Console.WriteLine("Usage:");
      Console.WriteLine("  WiseDBF2XLS_International dbfFileName/searchPattern [codePageName|codepageNumber|DEFAULT] ");
      Console.WriteLine();
      Console.WriteLine("  Valid codePageNames:  \n  " + String.Join(",", Enum.GetNames(typeof(DbfSetCodePage.CodepageCodes))));
      Console.WriteLine("  Valid codepageNumbers:\n  " + String.Join(",", encodings));     
      Console.WriteLine("  DEFAULT (literal text): use as codepage code readed from DBF file");
      Console.WriteLine();

      
    }

    static List<DbfItem> GetDbfList(string[] args )
    {
      string[] fileNames = null;

      if (File.Exists(args[0]))
      {
        fileNames = new string[1];
        fileNames[0] = Path.GetFullPath(args[0]);
      }
      else
      {
        string dir     = Path.GetDirectoryName(args[0]);
        string pattern = Path.GetFileName(args[0]);

        if (String.IsNullOrWhiteSpace(dir))
        {
          dir = Directory.GetCurrentDirectory();
        }

        fileNames = Directory.GetFiles(dir, pattern);
      }

      if (fileNames.Length < 1)
      {
        throw new Exception(String.Format("'{0}' fileName/searchPattern: file(s) NOT found!", args[0])); 
      }

      //

      List<DbfItem> dbfList   = new List<DbfItem>();

      
      if ((args.Length == 1) || (String.Compare(args[1], "DEFAULT", true) == 0))
      { // Read codepage numbers from DBF head
        foreach (var fileName in fileNames)
        {
          int encodingCodePage = int.MaxValue;

          try
          {
            DbfSetCodePage dbfSetCodePage = new DbfSetCodePage(DbfSetCodePage.CodepageCodes.OEM, true);
            var            dbfCodePage = dbfSetCodePage.GetCodepageByte(fileName);

            encodingCodePage = DbfSetCodePage.GetEncodingCodePageFromCodepageCodes((DbfSetCodePage.CodepageCodes)dbfCodePage);          
          }
          catch (Exception e)
          {
            throw new Exception(String.Format("'{0}' file invalid!\n[{1}]", fileName, e.Message)); 
          }
        
          dbfList.Add(new DbfItem(fileName, encodingCodePage));
        }        
      }
      else
      {
        int encodingCodePage  = int.MinValue;                                // impossible value as Encoding.CodePage
        
        if (Int32.TryParse(args[1], out encodingCodePage))
        {
          try
          {
            var encoding = Encoding.GetEncoding(encodingCodePage);
          }
          catch (Exception e)
          {
            throw new Exception(String.Format("'{0}' codepageNumber invalid!\n[{1}]", encodingCodePage, e.Message)); 
          }
        }


        if (encodingCodePage < 0)
        {
          var dbfCodePagesText = Enum.GetNames(typeof(DbfSetCodePage.CodepageCodes));
          var dbfCodePageIndex = Array.IndexOf(dbfCodePagesText, args[1]);

          if (dbfCodePageIndex >= 0)
          {
            var dbfCodePage = EnumConverter<DbfSetCodePage.CodepageCodes>.LoadString(args[1]);

            encodingCodePage = DbfSetCodePage.GetEncodingCodePageFromCodepageCodes(dbfCodePage);
          }
        }

        if (encodingCodePage < 0)
        {
          throw new Exception(String.Format("'{0}' invalid codepage parameter!", args[1])); 
        }        

        foreach (var fileName in fileNames)
        {
          dbfList.Add(new DbfItem(fileName, encodingCodePage));
        }
      }      

      return dbfList;
    }
  }

  public struct DbfItem
  {
    public readonly string fileName; 
    public readonly int    codepage;

    public DbfItem(string fileName, int codepage)
    {
      this.fileName = fileName;
      this.codepage = codepage;

      try
      {
        var encoding = Encoding.GetEncoding(codepage);
      }
      catch (Exception e)
      {
        throw new Exception(String.Format("'{0}' codepageNumber invalid!\n[{1}]", codepage, e.Message)); 
      }

      try
      {
        DbfSetCodePage dbfSetCodePage = new DbfSetCodePage(DbfSetCodePage.CodepageCodes.OEM, true);
        var            dbfCodePage = dbfSetCodePage.GetCodepageByte(fileName);            
      }
      catch (Exception e)
      {
        throw new Exception(String.Format("'{0}' file invalid!\n[{1}]", fileName, e.Message)); 
      }
    }
  }

  static class EnumConverter<TEnum> where TEnum : struct, IConvertible
  {
    public static readonly Func<long, TEnum> Convert = GenerateConverter();

    static Func<long, TEnum> GenerateConverter()
    {
        var parameter = Expression.Parameter(typeof(long));
        var dynamicMethod = Expression.Lambda<Func<long, TEnum>>(
            Expression.Convert(parameter, typeof(TEnum)),
            parameter);
        return dynamicMethod.Compile();
    }

    //

    public static TEnum LoadString(string value, TEnum defaultValue = default(TEnum), bool ignoreCase = true) 
    {
      if (Enum.IsDefined(typeof(TEnum), value))
      {
        return (TEnum)Enum.Parse(typeof(TEnum), value, ignoreCase);
      }
      return defaultValue;
    }
  }
}
