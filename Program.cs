using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Linq.Expressions;

namespace WiseDBF2XLS_International
{
  using eMeL;
  using NDbfReaderEx;
  
  class Program
  {
    static void Main(string[] args)
    {
      Console.WriteLine("**** Convert DBF to XLS with localization ****  (c) eMeL [www.emel.hu] **** FreeWare ****");
      Console.WriteLine();

      if (args.Length != 2)
      {
        DispUsage();
      }
      else
      {
        var codePagesText = Enum.GetNames(typeof(DbfSetCodePage.CodepageCodes));
        var codePageIndex = Array.IndexOf(codePagesText, args[1]);

        if (codePageIndex < 0)
        {
          Console.WriteLine();
          Console.WriteLine("Error! Invalid codepage!\n\n");
          DispUsage();
        }
        else
        {
          DbfSetCodePage.CodepageCodes codePage = EnumConverter<DbfSetCodePage.CodepageCodes>.LoadString(args[1]);

          string[] fileNames = null;

          if (File.Exists(args[0]))
          {
            fileNames = new string[1];
            fileNames[0] = args[0];
          }
          else
          {
            fileNames = Directory.GetFiles(Directory.GetCurrentDirectory(), args[0]);
          }

          if (fileNames.Length > 0)
          {
            DbfSetCodePage setCP = new DbfSetCodePage(codePage, true);    

            foreach (var fileName in fileNames)
            {
              try
              {
                setCP.SetCodepageByte(fileName);
                Console.WriteLine("Success: '{0}' file : code page changed to {1}.", fileName, codePage.ToString());
              }
              catch (Exception e)
              {
                Console.WriteLine("Error: '{0}' file : code page NOT changed.", fileName);
                Console.WriteLine("      " + e.Message);
                Environment.ExitCode = 2;
              }
            }
          }
          else
          {
            Console.WriteLine("Error: '{0}' file(s) not exists!\n", args[0]);
            Environment.ExitCode = 1;
          }
        }
      }

      #if DEBUG
      Console.WriteLine("Press Enter to close...");
      Console.ReadLine();
      #endif
    }

    static void DispUsage()
    {
      Console.WriteLine("Usage:");
      Console.WriteLine("  WiseDBF2XLS_International fileName/searchPattern codePageName|codepageNumber|DEFAULT ");
      Console.WriteLine();
      Console.WriteLine("  Valid codePageNames:\n  " + String.Join(",", Enum.GetNames(typeof(DbfSetCodePage.CodepageCodes))));
      Console.WriteLine("  Valid codepageNumbers: as .Net Encoding.CodePage Property");
      Console.WriteLine("  DEFAULT (literal text): use as codepage code readed from DBF file");
      Console.WriteLine();
    }
  }

  //

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
