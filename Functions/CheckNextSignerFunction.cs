using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Threading.Tasks;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using DocumentFormat.OpenXml.Packaging;
using System.Text.Json;
using System.Text.Json.Serialization;

public class CheckNextSignerFunction
{
    private readonly ILogger _logger;

    public CheckNextSignerFunction(ILogger<CheckNextSignerFunction> logger)
    {
        _logger = logger;
    }

    [Function("CheckNextSigner")]
    public async Task<HttpResponseData> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post")] HttpRequestData req)
    {
        _logger.LogInformation("Received base64-encoded .docx for signature check.");

        try
        {
            // Read and parse the body
            using var reader = new StreamReader(req.Body);
            var bodyText = await reader.ReadToEndAsync();
            var payload = JsonSerializer.Deserialize<BinaryPayload>(bodyText);

            if (string.IsNullOrWhiteSpace(payload?.Content))
            {
                var errorResponse = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
                await errorResponse.WriteAsJsonAsync(new { Error = "Missing base64 document content." });
                return errorResponse;
            }

            // Decode base64 content
            byte[] fileBytes = Convert.FromBase64String(payload.Content);
            using var stream = new MemoryStream(fileBytes);
            using var doc = WordprocessingDocument.Open(stream, false);

            var mainPart = doc.MainDocumentPart;
            bool isDigitallySigned = doc.DigitalSignatureOriginPart != null;

            XDocument xDoc = XDocument.Parse(mainPart.Document.InnerXml);
            XNamespace o = "urn:schemas-microsoft-com:office:office";
            var signatureLines = xDoc.Descendants(o + "signatureline").ToList();

            var firstUnsignedSigner = signatureLines
                .Select(sig =>
                {
                    string signerName = sig.Attribute(o + "suggestedsigner")?.Value ?? "Unknown";
                    string email = sig.Attribute(o + "suggestedsigneremail")?.Value ?? "Unknown";
                    bool isSigned = isDigitallySigned && VerifySignature(doc, signerName);
                    return new { SuggestedSigner = signerName, Email = email, Signed = isSigned };
                })
                .FirstOrDefault(s => !s.Signed);

            var response = req.CreateResponse(System.Net.HttpStatusCode.OK);
            await response.WriteAsJsonAsync(new { FirstUnsignedSigner = firstUnsignedSigner });

            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error processing document: {ex.Message}");
            var errorResponse = req.CreateResponse(System.Net.HttpStatusCode.InternalServerError);
            await errorResponse.WriteAsJsonAsync(new { Error = "Failed to process .docx content." });
            return errorResponse;
        }
    }

    private static bool VerifySignature(WordprocessingDocument doc, string signerName)
    {
        var signatureOrigin = doc.DigitalSignatureOriginPart;
        if (signatureOrigin == null) return false;

        foreach (var part in signatureOrigin.Parts)
        {
            using var sigStream = part.OpenXmlPart.GetStream();
            XDocument sigXml = XDocument.Load(sigStream);

            var actualSigner = sigXml.Descendants()
                .FirstOrDefault(e => e.Name.LocalName == "X509IssuerName")?.Value;

            Console.WriteLine($"Hmmmm: {actualSigner}");

            if (actualSigner != null && actualSigner.IndexOf(signerName, StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
        }
        return false;
    }
    
    private class BinaryPayload
    {
        [JsonPropertyName("$content-type")]
        public string ContentType { get; set; }

        [JsonPropertyName("$content")]
        public string Content { get; set; }
    }
}
