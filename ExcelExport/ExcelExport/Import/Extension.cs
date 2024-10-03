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

    public static Stream ToStream(this string dataUrl)
    {
        string base64 = dataUrl.Split(',')[1];

        byte[] bytes = Convert.FromBase64String(base64);
        MemoryStream stream = new(bytes);
        stream.Seek(0, SeekOrigin.End);
        stream.Close();
        return stream;
    }

    public static (string contentType, string extension) GetContentType(this string dataUrl)
    {
        string contentType = dataUrl.Split(';')[0].Split(':')[1];
        return (contentType, contentType.Split("/").Last());
    }

    public static (string contentType, string extension, byte[] buffer) ReadDataUrl(this string dataUrl)
    {
        string contentType = dataUrl.Split(';')[0].Split(':')[1];
        string extension = contentType.Split("/").Last();
        string base64 = dataUrl.Split(',')[1];
        byte[] data = Convert.FromBase64String(base64);

        return (contentType, extension, data);
    }
}
