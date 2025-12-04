using System.Globalization;
using Contract.Api.Models;
using Xceed.Words.NET;
using Xceed.Document.NET;

namespace Contract.Api.Services;

public class AgreementDocumentService
{
    private readonly string _templatePath;
    private const decimal VatPercent = 12m;

    public AgreementDocumentService(string templatePath)
    {
        _templatePath = templatePath;
    }

    public string GenerateDocx(AgreementRequest request, string outputDir)
    {
        var ru = new CultureInfo("ru-RU");
        Directory.CreateDirectory(outputDir);

        var docxPath = Path.Combine(
            outputDir,
            $"dogovor_{request.AgreementNumber}_{DateTime.Now:yyyyMMddHHmmss}.docx"
        );

        // ЧТОБЫ НЕ ЛОВИТЬ "file is used by another process"
        using var fs = new FileStream(_templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var doc = DocX.Load(fs);

        // номер и дата
        doc.ReplaceText("{AGREEMENT_NUMBER}", request.AgreementNumber);
        doc.ReplaceText("{AGREEMENT_DATE}", DateTime.Now.ToString("dd MMMM yyyy года", ru));

        // расчёт строк
        var items = request.Items.Select((x, i) =>
        {
            var sum = x.PriceNoVat * x.Quantity;
            var vat = Math.Round(sum * VatPercent / 100m, 2);
            var total = sum + vat;

            return new AgreementItemCalculated
            {
                Index = i + 1,
                Name = x.Name,
                Unit = x.Unit,
                Quantity = x.Quantity,
                PriceNoVat = x.PriceNoVat,
                SumNoVat = sum,
                Vat = vat,
                Total = total
            };
        }).ToList();

        InsertItemsTable(doc, items);

        var totalSum = items.Sum(x => x.Total);
        doc.ReplaceText("{TOTAL_SUM_WITH_VAT}", totalSum.ToString("N2", ru));
        doc.ReplaceText("{TOTAL_SUM_WORDS}", totalSum.ToString("N2", ru)); // потом пропишем словами

        doc.SaveAs(docxPath);
        return docxPath;
    }

    // ======== САМЫЙ ПРОСТОЙ И НАДЁЖНЫЙ ВАРИАНТ ТАБЛИЦЫ БЕЗ MERGE =========
    private static void InsertItemsTable(DocX doc, List<AgreementItemCalculated> items)
    {
        var ph = doc.Paragraphs.FirstOrDefault(p => p.Text.Contains("{ITEMS_TABLE}"));
        if (ph == null)
            throw new Exception("Плейсхолдер {ITEMS_TABLE} не найден в шаблоне");

        // если вдруг пусто — просто выведем заглушку
        if (items == null || items.Count == 0)
        {
            ph.ReplaceText("{ITEMS_TABLE}", "ПОЗИЦИИ ОТСУТСТВУЮТ");
            return;
        }

        int cols = 9;
        int rows = items.Count + 2; // заголовок + строки + итог

        var table = doc.AddTable(rows, cols);
        table.Design = TableDesign.TableGrid;
        table.Alignment = Alignment.center;

        string[] header = {
            "№", "Наименование товара", "Ед. изм", "Кол-во",
            "Цена за ед. (без НДС)", "Стоимость (без НДС)",
            "Ставка НДС", "Сумма НДС", "Всего с НДС"
        };

        // ШАПКА
        for (int c = 0; c < cols; c++)
        {
            table.Rows[0].Cells[c].Paragraphs[0]
                .Append(header[c])
                .Bold()
                .FontSize(10)
                .Alignment = Alignment.center;
        }

        // ТОВАРЫ
        for (int i = 0; i < items.Count; i++)
        {
            var x = items[i];
            var r = table.Rows[i + 1];

            r.Cells[0].Paragraphs[0].Append(x.Index.ToString()).Alignment = Alignment.center;
            r.Cells[1].Paragraphs[0].Append(x.Name);
            r.Cells[2].Paragraphs[0].Append(x.Unit).Alignment = Alignment.center;
            r.Cells[3].Paragraphs[0].Append(x.Quantity.ToString("0.##")).Alignment = Alignment.right;
            r.Cells[4].Paragraphs[0].Append(x.PriceNoVat.ToString("N2")).Alignment = Alignment.right;
            r.Cells[5].Paragraphs[0].Append(x.SumNoVat.ToString("N2")).Alignment = Alignment.right;
            r.Cells[6].Paragraphs[0].Append("12%").Alignment = Alignment.center;
            r.Cells[7].Paragraphs[0].Append(x.Vat.ToString("N2")).Alignment = Alignment.right;
            r.Cells[8].Paragraphs[0].Append(x.Total.ToString("N2")).Alignment = Alignment.right;
        }

        // ИТОГО — БЕЗ MERGE, ПРОСТО ОТДЕЛЬНАЯ СТРОКА
        var totalRow = table.Rows[rows - 1];

        totalRow.Cells[0].Paragraphs[0].Append("ИТОГО:").Bold();
        totalRow.Cells[1].Paragraphs[0].Append("");
        totalRow.Cells[2].Paragraphs[0].Append("");
        totalRow.Cells[3].Paragraphs[0].Append("");
        totalRow.Cells[4].Paragraphs[0].Append("");

        var totalNoVat = items.Sum(x => x.SumNoVat);
        var totalVat = items.Sum(x => x.Vat);
        var totalWithVat = items.Sum(x => x.Total);

        totalRow.Cells[5].Paragraphs[0].Append(totalNoVat.ToString("N2")).Bold().Alignment = Alignment.right;
        totalRow.Cells[6].Paragraphs[0].Append("").Alignment = Alignment.center; // можно "12%"
        totalRow.Cells[7].Paragraphs[0].Append(totalVat.ToString("N2")).Bold().Alignment = Alignment.right;
        totalRow.Cells[8].Paragraphs[0].Append(totalWithVat.ToString("N2")).Bold().Alignment = Alignment.right;

        ph.InsertTableAfterSelf(table);
        ph.Remove(false);
    }

    private class AgreementItemCalculated
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        public string Unit { get; set; } = "";
        public decimal Quantity { get; set; }
        public decimal PriceNoVat { get; set; }
        public decimal SumNoVat { get; set; }
        public decimal Vat { get; set; }
        public decimal Total { get; set; }
    }
}
