using Microsoft.Extensions.Configuration;

using System.Diagnostics;

using SzarotkaWZ;
using SzarotkaWZ.Models;

try
{
    var builder = new ConfigurationBuilder()
        .SetBasePath(Directory.GetCurrentDirectory())
        .AddJsonFile("appsettings.json", optional: false);

    IConfiguration config = builder.Build();


    var csvPath = config.GetSection("Lokalizacja_pliku_CSV").Value;
    var docxPath = config.GetSection("Lokalizacja_pliku_docx").Value;
    var maxSheet = int.Parse(config.GetSection("Ilość_wz_do_pobrania").Value);
    var wordApp = config.GetSection("Lokalizacja_pliku_word.exe").Value;

    int StartSheet = 0;
    maxSheet += StartSheet + 1;
    List<Wz> wzs = [];

    for (int i = StartSheet; i < maxSheet; i++)
    {
        var wz = ClosedXMLGetData.GetModel(i, csvPath);
        if (wz.Nabywca != "")
            wzs.Add(wz);
    }
    string formattedDate = DateTime.Today.ToString("O");

    var extension = Path.GetExtension(docxPath);
    var name = Path.GetFileNameWithoutExtension(docxPath);
    var path = Path.GetDirectoryName(docxPath);

    if(!System.IO.Directory.Exists(docxPath))
    {
        var directory = new DirectoryInfo(path);
        System.IO.Directory.CreateDirectory(directory.FullName);
    }

    for (int i = 0; ; i++)
    {
        if (i == 0)
        {
            if (!System.IO.File.Exists(Path.Combine(path, $"{name}-{Date()}{extension}")))
            {
                docxPath = Path.Combine(path, $"{name}-{Date()}{extension}");
                break;
            }
        }

        if (!System.IO.File.Exists(Path.Combine(path, $"{name}-{Date()}({i}){extension}")))
        {
            docxPath = Path.Combine(path, $"{name}-{Date()}({i}){extension}");
            break;
        }
    }

    OpenXmlCreatedDocx.CreateDocxWithTables(docxPath, wzs);


    var p = new Process();
    p.StartInfo.FileName = wordApp;
    p.StartInfo.ArgumentList.Add(docxPath);
    p.Start();

    Environment.Exit(0);
}
catch (Exception ex)
{
    Console.WriteLine("Error wiadomość{0}{1}{2}Error Stack Trace{3}{4}"
        , Environment.NewLine
        , ex.Message
        , Environment.NewLine
        , Environment.NewLine
        , ex.StackTrace);
    Console.ReadKey();
}

static string Date()
{
    var today = DateTime.Today;
    string forrmated = $"{today.Day}-{today.Month}-{today.Year}";
    return forrmated;
}