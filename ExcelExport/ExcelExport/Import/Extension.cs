namespace ExcelExport.Import;

public static class Extension
{
    public static MemoryFormFile ToMemoryFormFile(this string dataUrl)
    {
        string contentType = dataUrl.Split(';')[0].Split(':')[1];
        string base64 = dataUrl.Split(',')[1];
        byte[] data = Convert.FromBase64String(base64);

        using MemoryStream stream = new(data);
        MemoryFormFile formFile = new(stream, contentType, "your-file-name.png");

        return formFile;
    }
}
