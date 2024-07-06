﻿using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Internal;

namespace ExcelExport.Import;

public class MemoryFormFile(MemoryStream stream,
                      string contentType,
                      string fileName)
{
    private readonly MemoryStream _stream = stream;
    private readonly string _contentType = contentType;

    public string ContentType => _contentType;
    public string ContentDisposition => $"form-data; name=\"{Name}\"; filename=\"{FileName}\"";
    public IHeaderDictionary Headers => new HeaderDictionary();
    public long Length { get; } = stream.Length;
    public string Name { get; } = fileName;
    public string FileName { get; } = fileName;

    public void CopyTo(Stream target)
    {
        _stream.CopyTo(target);
    }

    public async Task CopyToAsync(Stream target, CancellationToken cancellationToken = default)
    {
        await _stream.CopyToAsync(target, cancellationToken);
    }

    public Stream OpenReadStream()
    {
        return _stream;
    }

    public IFormFile ToFormFile()
    {
        FormFile formFile = new(OpenReadStream(), 0, Length, Name, FileName)
        {
            Headers = new HeaderDictionary(),
            ContentType = ContentType
        };
        return formFile;
    }
}
