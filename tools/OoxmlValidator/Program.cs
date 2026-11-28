using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

if (args.Length != 1 || !File.Exists(args[0]))
{
    Console.Error.WriteLine("Usage: OoxmlValidator <document.docx>");
    return 2;
}

using var document = WordprocessingDocument.Open(args[0], false);
var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
var errors = validator.Validate(document).ToList();
if (errors.Count == 0)
{
    Console.WriteLine($"ISO/IEC 29500 (Office 2019) validation passed: {args[0]}");
    return 0;
}

foreach (var error in errors)
{
    Console.Error.WriteLine($"{error.ErrorType}: part={error.Part?.Uri}; path={error.Path?.XPath}; {error.Description}");
}
Console.Error.WriteLine($"ISO/IEC 29500 validation failed: {errors.Count} schema error(s)");
return 1;
