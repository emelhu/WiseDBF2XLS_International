using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace eMeL
{
  /// <summary>
  /// Set codepage identifier of a DBF datafile.
  /// There is not a value of code  page field in DBF file usually.
  /// For example: Clipper 5.2 leave it with zero always.
  /// Set valid codepage is important if you want read/write DBF file by OLE.
  /// </summary>
  public class DbfSetCodePage
  {
    private byte codePage;
    private bool throwException;

    private List<DbfFileTypes> enabledDbfFileTypes = null;          // default null mean that, all type is enabled

    //

    private const int codepageBytePosition = 29;                    // http://www.dbf2002.com/dbf-file-format.html

    /// <summary>
    /// Valid codepage bytes by standard codepage names for DBF file.
    /// information have got from http://forums.esri.com/Thread.asp?c=93&f=1170&t=197185#587982
    /// </summary>
    public enum CodepageCodes : byte                                       
    { 
      OEM         = 0x00,                                           // OEM = 0 
      CP437       = 0x01,                                           // Codepage_437_US_MSDOS = &H1 
      CP850       = 0x02,                                           // Codepage_850_International_MSDOS = &H2 
      CP1252      = 0x03,                                           // Codepage_1252_Windows_ANSI = &H3 
      ANSI        = 0x57,                                           // ANSI = &H57 
      CP737       = 0x6A,                                           // Codepage_737_Greek_MSDOS = &H6A 
      CP852       = 0x64,                                           // Codepage_852_EasernEuropean_MSDOS = &H64 
      CP857       = 0x6B,                                           // Codepage_857_Turkish_MSDOS = &H6B 
      CP861       = 0x67,                                           // Codepage_861_Icelandic_MSDOS = &H67 
      CP865       = 0x66,                                           // Codepage_865_Nordic_MSDOS = &H66 
      CP866       = 0x65,                                           // Codepage_866_Russian_MSDOS = &H65 
      CP950       = 0x78,                                           // Codepage_950_Chinese_Windows = &H78 
      CP936       = 0x7A,                                           // Codepage_936_Chinese_Windows = &H7A 
      CP932       = 0x7B,                                           // Codepage_932_Japanese_Windows = &H7B 
      CP1255      = 0x7D,                                           // Codepage_1255_Hebrew_Windows = &H7D 
      CP1256      = 0x7E,                                           // Codepage_1256_Arabic_Windows = &H7E 
      CP1250      = 0xC8,                                           // Codepage_1250_Eastern_European_Windows = &HC8 
      CP1251      = 0xC9,                                           // Codepage_1251_Russian_Windows = &HC9 
      CP1254      = 0xCA,                                           // Codepage_1254_Turkish_Windows = &HCA 
      CP1253      = 0xCB                                            // Codepage_1253_Greek_Windows = &HCB 
    };

    /// <summary>
    /// http://www.dbf2002.com/dbf-file-format.html
    /// </summary>
    public enum DbfFileTypes : byte
    {
      FoxBASE       = 0x02,                                         // FoxBASE
      DBase3        = 0x03,                                         // FoxBASE+/Dbase III plus, no memo
      BDE           = 0x04,                                         // Borland Database Engine *** NEW found eMeL
      VisualFoxPro  = 0x30,                                         // Visual FoxPro
      VisualFoxPro2 = 0x31,                                         // Visual FoxPro, autoincrement enabled
      VisualFoxPro3 = 0x32,                                         // Visual FoxPro with field type Varchar or Varbinary
      DBase4Sql     = 0x43,                                         // dBASE IV SQL table files, no memo
      DBase4SqlSys  = 0x63,                                         // dBASE IV SQL system files, no memo
      DBase3M       = 0x83,                                         // FoxBASE+/dBASE III PLUS, with memo
      DBase4M       = 0x8B,                                         // dBASE IV with memo
      DBase4SqlM    = 0xCB,                                         // dBASE IV SQL table files, with memo
      FoxPro2M      = 0xF5,                                         // FoxPro 2.x (or earlier) with memo
      HiperSix      = 0xE5,                                         // HiPer-Six format with SMT memo file
      FoxBASE2      = 0xFB                                          // FoxBASE
    };

    //

    public static int GetEncodingCodePageFromCodepageCodes(CodepageCodes codepageCode)
    {
      switch (codepageCode)
      {
        case CodepageCodes.OEM:
          return CultureInfo.CurrentCulture.TextInfo.OEMCodePage;   
        case CodepageCodes.CP437:
          return 437;
        case CodepageCodes.CP850:
          return 850;
        case CodepageCodes.CP1252:
          return 1252;
        case CodepageCodes.ANSI:
          return CultureInfo.CurrentCulture.TextInfo.ANSICodePage;  
        case CodepageCodes.CP737:
          return 737;
        case CodepageCodes.CP852:
          return 852;
        case CodepageCodes.CP857:
          return 857;
        case CodepageCodes.CP861:
          return 861;
        case CodepageCodes.CP865:
          return 865;
        case CodepageCodes.CP866:
          return 866;
        case CodepageCodes.CP950:
          return 950;
        case CodepageCodes.CP936:
          return 936;
        case CodepageCodes.CP932:
          return 932;
        case CodepageCodes.CP1255:
          return 1255;
        case CodepageCodes.CP1256:
          return 1256;
        case CodepageCodes.CP1250:
          return 1250;
        case CodepageCodes.CP1251:
          return 1251;
        case CodepageCodes.CP1254:
          return 1254;
        case CodepageCodes.CP1253:
          return 1253;
      }

      return int.MinValue;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="codePage"></param>
    /// <param name="throwException"></param>
    public DbfSetCodePage(byte codePage, bool throwException = false)
    {
      this.codePage       = codePage;
      this.throwException = throwException;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="codePage"></param>
    /// <param name="throwException"></param>
    public DbfSetCodePage(CodepageCodes codePage, bool throwException = false)
    {
      this.codePage       = (byte)codePage;
      this.throwException = throwException;
    }

    //

    public void AddEnabledDbfFileType(DbfFileTypes enabledDbfType, bool clearListBeforeAdd = false)
    {
      if (enabledDbfFileTypes == null)
      {
        enabledDbfFileTypes = new List<DbfFileTypes>();
      }

      if (clearListBeforeAdd)
      {
        enabledDbfFileTypes.Clear();
      }

      if (! enabledDbfFileTypes.Contains(enabledDbfType))
      {
        enabledDbfFileTypes.Add(enabledDbfType);
      }
    }

    //

    /// <summary>
    /// Retrieve codepage byte of DBF file.
    /// File is checked for content of valid dBase file and optionally check enabled types of DBF file.
    /// </summary>
    /// <param name="dbfFile">Name of DBF file.</param>
    /// <param name="enabledDbfFileTypes">Enabled type codes of DBF file.</param>
    /// <returns>Codepage code of DBF file.</returns>
    /// 
    /// <exception cref="System.IO.IOException"   
    /// Throw an exception if file not exist or not accessible or size or type invalid.
    /// </exception>
    static public byte GetCodepageByte(string dbfFile, List<DbfFileTypes> enabledDbfFileTypes = null)           
    {
      byte retCodepage;

      using (FileStream fs = new FileStream(dbfFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
      {
        {
          int dbfFileType     = fs.ReadByte();          // 0          -- DBF File type:

          int lastUpdateYear  = fs.ReadByte();          // 1,2,3      -- Last update (YYMMDD)
          int lastUpdateMonth = fs.ReadByte();
          int lastUpdateDay   = fs.ReadByte();

          if ((dbfFileType < 0) || (lastUpdateYear < 0) || (lastUpdateMonth < 0) || (lastUpdateDay < 0))
          {
            throw new EndOfStreamException();
          }

          int maxYear = (DateTime.Today.Year - 2000) + 1;                         // DBase III / Clipper date

          if (lastUpdateYear >= 100)
          { // If not a dBaseIII or Clipper file, correct to            
            // because dBaseIII+ (and earlier) has a Year 2000 bug. It stores the year as simply the last two digits of the actual year. 
            maxYear += 100;
          }

          if ((lastUpdateYear  < 0) || (lastUpdateYear  > maxYear) ||             // !WARNING: only next year is valid
              (lastUpdateMonth < 1) || (lastUpdateMonth > 12)      ||
              (lastUpdateDay   < 1) || (lastUpdateDay   > 31))
          {
            throw new IOException("Not a DBF file! ('Last update' (YYMMDD) error!)");
          }

          //

          if (Array.IndexOf(Enum.GetValues(typeof(DbfFileTypes)), (DbfFileTypes)dbfFileType) < 0)
          {
            throw new IOException("Not a DBF file! ('DBF File type' is not valid!)");
          }

          //

          if (enabledDbfFileTypes != null)
          {
            if (enabledDbfFileTypes.IndexOf((DbfFileTypes)dbfFileType) < 0)
            {
              throw new IOException("Not a DBF file! ('DBF File type' is not enabled!)");
            }
          }
        }

        //

        fs.Position = codepageBytePosition;

        int codepageByte    = fs.ReadByte();

        {
          int reservedZero30  = fs.ReadByte();
          int reservedZero31  = fs.ReadByte();

          if ((codepageByte < 0) || (reservedZero30 < 0) || (reservedZero31 < 0))
          {
            throw new EndOfStreamException();
          }

          if ((reservedZero30 != 0) || (reservedZero31 != 0))
          {
            throw new IOException("Not a DBF file! (Offset 30/31 byte not contains zero!)");
          }

          if (Array.IndexOf(Enum.GetValues(typeof(CodepageCodes)), (CodepageCodes)codepageByte) < 0)
          {
            throw new IOException("Not a DBF file! ('Code page mark' is not valid!)");
          }
        }

        retCodepage =(byte)codepageByte;
      }

      return retCodepage;
    }

    /// <summary>
    /// Set codepage byte of DBF file.
    /// File is checked for content of valid dBase file and optionally check enabled types of DBF file.
    /// </summary>
    /// <param name="dbfFile">Name of DBF file.</param>
    /// <param name="codePage">Codepage code for set in DBF file</param>
    /// <param name="enabledDbfFileTypes">Enabled type codes of DBF file.</param>
    /// 
    /// <exception cref="System.IO.IOException"   
    /// Throw an exception if file not exist or not accessible or size or type invalid.
    /// </exception>
    ///  
    static public void SetCodepageByte(string dbfFile, byte codePage, List<DbfFileTypes> enabledDbfFileTypes = null)            
    {
      byte actCodepage = GetCodepageByte(dbfFile, enabledDbfFileTypes);

      if (actCodepage == codePage)
      {
        return;                                                     // shortcut exit (access optimisation)
      }

      using (FileStream fs = new FileStream(dbfFile, FileMode.Open, FileAccess.Write, FileShare.None))
      {
        fs.Position = codepageBytePosition;

        fs.WriteByte(codePage);
      }
    }

    /// <summary>
    /// Retrieve codepage byte of DBF file.
    /// File is checked for content of valid dBase file and optionally check enabled types of DBF file.
    /// </summary>
    /// <param name="dbfFile">Name of DBF file.</param>
    /// <param name="enabledDbfFileTypes">Enabled type codes of DBF file.</param>
    /// <returns>
    /// Codepage code of DBF file.
    /// If throw exception not enabled, a 0xFF returned if error occured.
    /// </returns>
    /// 
    /// <exception cref="System.IO.IOException"   
    /// If throw exception enabled:
    /// Throw an exception if file not exist or not accessible or size or type invalid.
    /// </exception>
    public byte GetCodepageByte(string dbfFile)
    {
      byte ret = 0xFF;                                                    // invalid

      if (throwException)
      {
        ret = GetCodepageByte(dbfFile, enabledDbfFileTypes);              // static tag
      }
      else
      {
        try
        {
          ret = GetCodepageByte(dbfFile, enabledDbfFileTypes);            // static tag
        }
        catch 
        {
          ret = 0xFF;                                                     // invalid
        }
      }

      return ret;
    }

    public bool isCodepageByteCorrect(string dbfFile)
    {
      byte actual = GetCodepageByte(dbfFile);                                                     
      
      return (actual == codePage);
    }

    /// <summary>
    /// Set codepage byte of DBF file.
    /// File is checked for content of valid dBase file and optionally check enabled types of DBF file.
    /// </summary>
    /// <param name="dbfFile">Name of DBF file.</param>
    /// 
    /// <exception cref="System.IO.IOException"   
    /// If throw exception enabled:
    /// Throw an exception if file not exist or not accessible or size or type invalid.
    /// </exception>
    public bool SetCodepageByte(string dbfFile)            
    {
      bool ret = true;                                                      // default return value                                                    

      if (throwException)
      {
        SetCodepageByte(dbfFile, codePage, enabledDbfFileTypes);            // static tag
      }
      else
      {
        try
        {
          SetCodepageByte(dbfFile, codePage, enabledDbfFileTypes);          // static tag
        }
        catch  
        {
          ret = false;                                                      // don't success write
        }
      }

      return ret;
    }
  }
}
