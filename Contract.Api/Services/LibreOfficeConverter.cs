using System.Diagnostics;

namespace Contract.Api.Services;

public class LibreOfficeConverter
{
    private readonly string _sofficePath;

    public LibreOfficeConverter(string sofficePath)
    {
        _sofficePath = sofficePath;
    }

    public string ConvertToPdf(string docxPath, string outputDir)
    {
        Directory.CreateDirectory(outputDir);

        var psi = new ProcessStartInfo
        {
            FileName = _sofficePath,
            Arguments = $"--headless --convert-to pdf --outdir \"{outputDir}\" \"{docxPath}\"",
            CreateNoWindow = true,
            UseShellExecute = false
        };

        using var process = Process.Start(psi);
        process.WaitForExit();

        var pdfPath = Path.ChangeExtension(
            Path.Combine(outputDir, Path.GetFileName(docxPath)),
            ".pdf"
        );

        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("LibreOffice не создал PDF", pdfPath);

        return pdfPath;
    }
}
