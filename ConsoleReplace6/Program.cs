using System.Text;
using ClosedXML.Excel;
using ConsoleAppFramework;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Utf8StringInterpolation;
using ZLogger;
using ZLogger.Providers;

//==
var builder = ConsoleApp.CreateBuilder(args);
builder.ConfigureServices((ctx,services) =>
{
    // Register appconfig.json to IOption<MyConfig>
    services.Configure<MyConfig>(ctx.Configuration);

    // Using Cysharp/ZLogger for logging to file
    services.AddLogging(logging =>
    {
        logging.ClearProviders();
        logging.SetMinimumLevel(LogLevel.Trace);
        var jstTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time");
        var utcTimeZoneInfo = TimeZoneInfo.Utc;
        logging.AddZLoggerConsole(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });
        });
        logging.AddZLoggerRollingFile(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });

            // File name determined by parameters to be rotated
            options.FilePathSelector = (timestamp, sequenceNumber) => $"logs/{timestamp.ToLocalTime():yyyy-MM-dd}_{sequenceNumber:00}.log";
            
            // The period of time for which you want to rotate files at time intervals.
            options.RollingInterval = RollingInterval.Day;
            
            // Limit of size if you want to rotate by file size. (KB)
            options.RollingSizeKB = 1024;        
        });
    });
});

var app = builder.Build();
app.AddCommands<ConsoleReplaceApp>();
app.Run();


public class ConsoleReplaceApp : ConsoleAppBase
{
    bool isAllPass = true;

    readonly ILogger<ConsoleReplaceApp> logger;
    readonly IOptions<MyConfig> config;

    private List<Dictionary<string, MyCell>> myListDicCells = new List<Dictionary<string, MyCell>>();

    public ConsoleReplaceApp(ILogger<ConsoleReplaceApp> logger, IOptions<MyConfig> config)
    {
        this.logger = logger;
        this.config = config;
    }

//    [Command("")]
    public void replace(string configpath, string format, string outpath)
    {
//== start
        logger.ZLogInformation($"==== tool {getMyFileVersion()} ====");
        if (!File.Exists(configpath))
        {
            logger.ZLogError($"[NG] エクセルファイルが見つかりません({configpath})");
            return;
        }
        if (!File.Exists(format))
        {
            logger.ZLogError($"[NG] エクセルファイルが見つかりません({format})");
            return;
        }

        logger.ZLogTrace($"configpath:{configpath}");
        logger.ZLogDebug($"RackSelectSheetType:{config.Value.RackSelectSheetType} RackSelectSheetName:{config.Value.RackSelectSheetName}");


        readConfigFile(configpath, myListDicCells);
        printDicCells(myListDicCells);

        createExcelFile(myListDicCells, format, outpath);


        //== finish
        logger.ZLogInformation($"==== tool finish ====");
    }


    private void readConfigFile(string configpath, List<Dictionary<string, MyCell>> myListDicCells)
    {
        bool findTableName = false;
        FileStream fs = new FileStream(configpath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using XLWorkbook xlWorkbook = new XLWorkbook(fs);
        IXLWorksheet sheet_config;
        if (xlWorkbook.TryGetWorksheet("config", out sheet_config))
        {
            foreach (var table in sheet_config.Tables)
            {
                logger.ZLogTrace($"table name:{table.Name}");
                if (string.Compare("table_config", table.Name) == 0)
                {
                    findTableName = true;
                    logger.ZLogTrace($"シートの中にテーブル名({table.Name})が見つかりました!");
                    int columnFirst = 1;
                    int columnMax = table.RangeAddress.ColumnSpan + 1;
                    int rowFirst = 2;
                    int rowMax = table.RangeAddress.RowSpan + 1;
                    // header
                    var listHeaders = new List<string>();
                    for (int column = columnFirst; column < columnMax; column++)
                    {
                        var cellHeader = table.Cell(1, column);
//                        logger.ZLogTrace($"column:{column} row:{1} Value:{cellHeader.Value.ToString()} Type:{cellHeader.Value.Type.ToString()}");
                        listHeaders.Add(cellHeader.Value.ToString());
                    }
                    // data
                    for (int row = rowFirst; row < rowMax; row++)
                    {
                        var tmpDic = new Dictionary<string, MyCell>();
                        int i = 0;
                        for (int column = columnFirst; column < columnMax; column++)
                        {
                            var cell = table.Cell(row, column);
//                            logger.ZLogTrace($"column:{column} row:{row} Value:{cell.Value.ToString()} Type:{cell.Value.Type.ToString()}");
                            var tmpCell = new MyCell();
                            tmpCell.key = listHeaders[i];
                            tmpCell.value = cell.Value;
                            tmpCell.type = cell.Value.Type;
                            tmpDic.Add(tmpCell.key, tmpCell);
                            i++;
                        }
                        myListDicCells.Add(tmpDic);
                    }
                }
            }

            if (findTableName == false)
            {
                logger.ZLogError($"シートの中にテーブル名( )が見つかりませんでした");
                throw new Exception($"[Error]シートの中にテーブル名( )が見つかりませんでした");
            }
        }
        else
        {
            logger.ZLogError($"シート名( )が見つかりませんでした");
            throw new Exception($"[Error]シートの中にテーブル名( )が見つかりませんでした");
        }
    }

    private void printDicCells(List<Dictionary<string, MyCell>> myListDicCells)
    {
        foreach (var dic in myListDicCells)
        {
            foreach (var key in dic.Keys)
            {
                logger.ZLogTrace($"key:{key} value:{dic[key].value.ToString()} type:{dic[key].type.ToString()}");
            }
            
        }
    }

    private void createExcelFile(List<Dictionary<string, MyCell>> myListDicCells, string format, string outpath)
    {
        logger.ZLogInformation($"== start createExcelFile ==");
        string rackSelectSheetType = config.Value.RackSelectSheetType;
        string rackSelectSheetName = config.Value.RackSelectSheetName;
        try
        {
            File.Copy(format, outpath, true);
        }
        catch (System.Exception)
        {
            throw;
        }
        using FileStream fsExcel = new FileStream(outpath, FileMode.Open, FileAccess.ReadWrite, FileShare.Write);
        using XLWorkbook xlWorkbookExcel = new XLWorkbook(fsExcel);

        // memory sheet name
        List<string> listDeleteSheetName = new List<string>();
        foreach (var worksheet in xlWorkbookExcel.Worksheets)
        {
            listDeleteSheetName.Add(worksheet.Name);
        }

        // copy type->target sheet
        foreach (var myDic in myListDicCells)
        {
            if (!myDic.ContainsKey(rackSelectSheetType))
            {
                logger.ZLogError($"[ERROR] key({rackSelectSheetType}) is not Contains");
                return;
            }
            if (!myDic.ContainsKey(rackSelectSheetName))
            {
                logger.ZLogError($"[ERROR] key({rackSelectSheetName}) is not Contains");
                return;
            }
            var myCell_type = myDic[rackSelectSheetType];
            var myCell_name = myDic[rackSelectSheetName];
            string formatSheetName = myCell_type.value.ToString();
            IXLWorksheet formatSheet;
            if (!xlWorkbookExcel.TryGetWorksheet(formatSheetName, out formatSheet))
            {
                logger.ZLogError($"[ERROR] format worksheet is missing {formatSheetName}");
                return;
            }
            if (myCell_name.value.IsBlank)
            {
                logger.ZLogError($"[ERROR] target worksheet name is Blank");
                return;
            }
            string replaceSheetName = myCell_name.value.ToString();
            var replaceSheet = formatSheet.CopyTo(replaceSheetName);
            if (replaceSheet == null || replaceSheet.IsEmpty())
            {
                logger.ZLogError($"[ERROR] target worksheet is missing {replaceSheetName}");
                return;
            }

            foreach (var key in myDic.Keys.Reverse())
            {
                var myCell = myDic[key];
                string replaceWord = myCell.key.ToString();
                logger.ZLogTrace($"key:{key} replaceWord:{replaceWord} type:{myCell.type}");

                IXLCells cells = replaceSheet.Search(replaceWord, System.Globalization.CompareOptions.IgnoreNonSpace, false);
                if (cells == null || cells.Count<IXLCell>() == 0)
                {
                    logger.ZLogWarning($"[ERROR] Search result == NULL or 0");
                    continue;
                }
                foreach (IXLCell? cell in cells)
                {
                    if (cell != null)
                    {
                        string targetWord = myCell.value.ToString();
                        string tmpString = cell.Value.ToString();
                        string newString = tmpString.Replace(replaceWord, targetWord);
                        cell.SetValue(newString);
                        logger.ZLogTrace($"tmpString:{tmpString} replaceWord:{replaceWord} newString:{newString}");
                    }
                }
            }
        }

        // delete format sheet
        foreach (var name in listDeleteSheetName)
        {
            xlWorkbookExcel.Worksheet(name).Delete();
        }

        xlWorkbookExcel.Save();
    }

    private string getMyFileVersion()
    {
        System.Diagnostics.FileVersionInfo ver = System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location);
        return ver.InternalName + "(" + ver.FileVersion + ")";
    }
}

//==
public class MyConfig
{
    public string RackSelectSheetType { get; set; } = "DEFAULT";
    public string RackSelectSheetName { get; set; } = "DEFAULT";
}

public class MyCell
{
    public string key = "DEFAULT";
    public XLCellValue value;
    public XLDataType type = XLDataType.Blank;
}
