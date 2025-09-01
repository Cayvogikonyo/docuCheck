using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Threading.Tasks;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using DocumentFormat.OpenXml.Packaging;
using System.Text.Json.Serialization;
using System.Text.Json;
using System.Net;

public class DocxSignatureCheckFunction
{
    private readonly ILogger _logger;

    public DocxSignatureCheckFunction(ILogger<DocxSignatureCheckFunction> logger)
    {
        _logger = logger;
    }
    [Function("CheckSignatures")]
    public async Task<HttpResponseData> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequestData req)
    {
        _logger.LogInformation("Received HTTP request with enriched JSON body.");

        try
        {
            // Read the entire JSON payload
            using var reader = new StreamReader(req.Body);
            string requestBody = await reader.ReadToEndAsync();

            // Define the structure for deserialization
            var envelope = JsonSerializer.Deserialize<EnrichedPayload>(requestBody);
            _logger.LogInformation($"JSON body:{envelope}");

            if (envelope?.Body?.Content == null)
            {
                var error = req.CreateResponse(HttpStatusCode.BadRequest);
                await error.WriteAsJsonAsync(new { Error = "Missing or invalid document content." });
                return error;
            }

            // Decode the base64-encoded document
            byte[] docBytes = Convert.FromBase64String(envelope.Body.Content);
            using var doc = WordprocessingDocument.Open(new MemoryStream(docBytes), false);

            var mainPart = doc.MainDocumentPart;
            bool isDigitallySigned = doc.DigitalSignatureOriginPart != null;

            XDocument xDoc = XDocument.Parse(mainPart.Document.InnerXml);
            XNamespace o = "urn:schemas-microsoft-com:office:office";
            var signatureLines = xDoc.Descendants(o + "signatureline").ToList();


            var resultList = signatureLines.Select(signature =>
            {
                string signerName = signature.Attribute(o+"suggestedsigner")?.Value ?? "Unknown";
                string signerEmail = signature.Attribute(o+"suggestedsigneremail")?.Value ?? "Unknown";
                bool hasSignatureLine = signature.Attribute("issignatureline")?.Value == "t";
                bool isSigned = isDigitallySigned && VerifySignature(doc, signerName);

                return new
                {
                    SuggestedSigner = signerName,
                    Email = signerEmail,
                    Signed = isSigned
                };
            }).ToList();


            int signatureCount = GetDigitalSignatureCount(doc);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteAsJsonAsync(new
            {
                SignatureCount = signatureCount,
                signatures = resultList
            });

            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error processing document: {ex.Message}");
            var fail = req.CreateResponse(HttpStatusCode.InternalServerError);
            await fail.WriteAsJsonAsync(new { Error = "Failed to process document." });
            return fail;
        }
    }

    private class EnrichedPayload
    {
        [JsonPropertyName("statusCode")]
        public int StatusCode { get; set; }

        [JsonPropertyName("headers")]
        public Dictionary<string, string> Headers { get; set; }

        [JsonPropertyName("body")]
        public BodyWrapper Body { get; set; }

        public class BodyWrapper
        {
            [JsonPropertyName("$content-type")]
            public string ContentType { get; set; }

            [JsonPropertyName("$content")]
            public string Content { get; set; }
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

            if (actualSigner != null && actualSigner.IndexOf(signerName, StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
        }
        return false;
    }

    private static int GetDigitalSignatureCount(WordprocessingDocument doc)
    {
        var signatureOrigin = doc.DigitalSignatureOriginPart;
        if (signatureOrigin == null) return 0;

        int count = 0;
        foreach (var part in signatureOrigin.Parts)
        {
            using var sigStream = part.OpenXmlPart.GetStream();
            XDocument sigXml = XDocument.Load(sigStream);
            count += sigXml.Descendants().Count(e => e.Name.LocalName == "X509IssuerName");
        }
        return count;
    }
}