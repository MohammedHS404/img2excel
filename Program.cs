using ClosedXML.Excel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Reflection;

using var workbook = new XLWorkbook();

var ws
    = workbook.Worksheets.Add("Duck");

string fileName = "sisi";
string extension = "jpg";

string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).Substring(0,
    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).IndexOf("bin"));

Bitmap bitmap = new Bitmap(Path.Combine(path, "Images", $"{fileName}.{extension}"));

int currentWidth = bitmap.Width;

int currentHeight = bitmap.Height;

int newHeight = 60;

int newWidth = (currentWidth * newHeight) / currentHeight;

bitmap = ResizeImage(bitmap, newWidth, newHeight);

for (int x = 0; x < bitmap.Width; x++)
{
    for (int y = 0; y < bitmap.Height; y++)
    {
        Console.WriteLine($"(${x},${y})");
        Color pixel = bitmap.GetPixel(x, y);


        var cell = ws.Cell(y + 1, x + 1);
        cell.Value = "";
        cell.Style.Fill.BackgroundColor = XLColor.FromColor(pixel);
    }
}

Console.WriteLine("Saving");
workbook.SaveAs(
    Path.Combine(path, "ExcelImages", $"{fileName}-{Guid.NewGuid().ToString().Substring(0, 4)}.xlsx"));
Console.WriteLine("Done");

Bitmap ResizeImage(Bitmap image, int width, int height)
{
    var destRect = new Rectangle(0, 0, width, height);
    var destImage = new Bitmap(width, height);

    destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

    using (var graphics = Graphics.FromImage(destImage))
    {
        graphics.CompositingMode = CompositingMode.SourceCopy;
        graphics.CompositingQuality = CompositingQuality.HighQuality;
        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        graphics.SmoothingMode = SmoothingMode.HighQuality;
        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

        using (var wrapMode = new ImageAttributes())
        {
            wrapMode.SetWrapMode(WrapMode.TileFlipXY);
            graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
        }
    }

    return destImage;
}