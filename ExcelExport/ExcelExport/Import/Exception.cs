namespace ExcelExport.Import;

public class ValidationException : Exception
{
    public List<string> Errors { get; }

    public ValidationException(List<string> errors, string message) : base(message) => Errors = errors;

    public static void Throw(List<string> errors)
        => throw new ValidationException(errors, "Validation failed");

    public static void Throw(List<string> errors, string message)
        => throw new ValidationException(errors, message);
}