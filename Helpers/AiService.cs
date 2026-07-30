using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSheetManager.Helpers
{
    internal static class AiService
    {
        static AiService()
        {
            try
            {
                // Enable TLS 1.2 & 1.3 and bypass self-signed SSL certificate errors for local/internal AI servers
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | (SecurityProtocolType)3072;
                ServicePointManager.ServerCertificateValidationCallback = (sender, cert, chain, sslPolicyErrors) => true;
            }
            catch { }
        }

        public static string GetEndpointUrl()
        {
            string baseUrl = SettingsHelper.GetAiBaseUrl().Trim();
            if (string.IsNullOrEmpty(baseUrl)) baseUrl = "http://localhost:11434/v1";

            baseUrl = baseUrl.TrimEnd('/');

            // If user enters a full path containing /chat/completions, /completions, /api/chat, use exact URL as provided!
            if (baseUrl.IndexOf("/chat/completions", StringComparison.OrdinalIgnoreCase) >= 0 ||
                baseUrl.IndexOf("/completions", StringComparison.OrdinalIgnoreCase) >= 0 ||
                baseUrl.IndexOf("/api/chat", StringComparison.OrdinalIgnoreCase) >= 0 ||
                baseUrl.IndexOf("/api/generate", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return baseUrl;
            }

            if (baseUrl.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
            {
                return baseUrl + "/chat/completions";
            }

            return baseUrl + "/v1/chat/completions";
        }

        public static async Task<(bool Success, string Message)> TestConnectionDiagnosticsAsync()
        {
            try
            {
                string endpoint = GetEndpointUrl();
                string modelName = SettingsHelper.GetAiModelName();
                string apiKey = SettingsHelper.GetAiApiKey();

                string jsonPayload = $"{{\"model\":\"{modelName}\",\"messages\":[{{\"role\":\"user\",\"content\":\"Hi\"}}],\"max_tokens\":5}}";

                var request = (HttpWebRequest)WebRequest.Create(endpoint);
                request.Method = "POST";
                request.ContentType = "application/json; charset=utf-8";
                request.Timeout = 15000;

                if (!string.IsNullOrEmpty(apiKey) && !apiKey.Equals("local", StringComparison.OrdinalIgnoreCase))
                {
                    request.Headers["Authorization"] = "Bearer " + apiKey;
                }

                byte[] bodyBytes = Encoding.UTF8.GetBytes(jsonPayload);
                request.ContentLength = bodyBytes.Length;

                using (var reqStream = await request.GetRequestStreamAsync())
                {
                    await reqStream.WriteAsync(bodyBytes, 0, bodyBytes.Length);
                }

                using (var response = (HttpWebResponse)await request.GetResponseAsync())
                using (var resStream = response.GetResponseStream())
                using (var reader = new StreamReader(resStream, Encoding.UTF8))
                {
                    string responseBody = await reader.ReadToEndAsync();
                    return (true, $"✓ Connected successfully!\nEndpoint: {endpoint}\nModel: {modelName}");
                }
            }
            catch (WebException webEx)
            {
                string detail = webEx.Message;
                if (webEx.Response is HttpWebResponse httpRes)
                {
                    try
                    {
                        using var resStream = httpRes.GetResponseStream();
                        using var reader = new StreamReader(resStream, Encoding.UTF8);
                        detail = $"HTTP {(int)httpRes.StatusCode} ({httpRes.StatusDescription}): {reader.ReadToEnd()}";
                    }
                    catch { }
                }
                return (false, $"✗ Connection Error:\nEndpoint: {GetEndpointUrl()}\nDetail: {detail}");
            }
            catch (Exception ex)
            {
                return (false, $"✗ Error: {ex.Message}\nEndpoint: {GetEndpointUrl()}");
            }
        }

        public static async Task<string> GetCompletionAsync(string prompt, string? systemPrompt = null)
        {
            return await Task.Run(() =>
            {
                try
                {
                    string endpoint = GetEndpointUrl();
                    string modelName = SettingsHelper.GetAiModelName();
                    string apiKey = SettingsHelper.GetAiApiKey();

                    systemPrompt ??= "You are an intelligent AI assistant integrated inside Microsoft Excel. Provide helpful, accurate, and concise answers.";

                    string escapedSystem = EscapeJsonString(systemPrompt);
                    string escapedUser = EscapeJsonString(prompt);

                    string jsonPayload = $"{{\"model\":\"{modelName}\",\"messages\":[{{\"role\":\"system\",\"content\":\"{escapedSystem}\"}},{{\"role\":\"user\",\"content\":\"{escapedUser}\"}}],\"temperature\":0.3}}";

                    var request = (HttpWebRequest)WebRequest.Create(endpoint);
                    request.Method = "POST";
                    request.ContentType = "application/json; charset=utf-8";
                    request.Timeout = 60000;

                    if (!string.IsNullOrEmpty(apiKey) && !apiKey.Equals("local", StringComparison.OrdinalIgnoreCase))
                    {
                        request.Headers["Authorization"] = "Bearer " + apiKey;
                    }

                    byte[] bodyBytes = Encoding.UTF8.GetBytes(jsonPayload);
                    request.ContentLength = bodyBytes.Length;

                    using (var reqStream = request.GetRequestStream())
                    {
                        reqStream.Write(bodyBytes, 0, bodyBytes.Length);
                    }

                    using (var response = (HttpWebResponse)request.GetResponse())
                    using (var resStream = response.GetResponseStream())
                    using (var reader = new StreamReader(resStream, Encoding.UTF8))
                    {
                        string responseBody = reader.ReadToEnd();
                        return ExtractContentFromOpenAiJson(responseBody);
                    }
                }
                catch (WebException webEx)
                {
                    if (webEx.Response is HttpWebResponse httpRes)
                    {
                        using var resStream = httpRes.GetResponseStream();
                        using var reader = new StreamReader(resStream, Encoding.UTF8);
                        return $"AI Server Error [{(int)httpRes.StatusCode}]: {reader.ReadToEnd()}";
                    }
                    return $"Error connecting to Local AI: {webEx.Message}";
                }
                catch (Exception ex)
                {
                    return $"Error connecting to Local AI: {ex.Message}";
                }
            });
        }

        private static string EscapeJsonString(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Replace("\\", "\\\\")
                       .Replace("\"", "\\\"")
                       .Replace("\r", "\\r")
                       .Replace("\n", "\\n")
                       .Replace("\t", "\\t");
        }

        private static string ExtractContentFromOpenAiJson(string json)
        {
            try
            {
                int idx = json.IndexOf("\"content\"", StringComparison.OrdinalIgnoreCase);
                if (idx < 0) return json;

                int colonIdx = json.IndexOf(':', idx);
                if (colonIdx < 0) return json;

                int firstQuote = json.IndexOf('"', colonIdx + 1);
                if (firstQuote < 0) return json;

                StringBuilder sb = new StringBuilder();
                for (int i = firstQuote + 1; i < json.Length; i++)
                {
                    char ch = json[i];
                    if (ch == '\\' && i + 1 < json.Length)
                    {
                        char next = json[i + 1];
                        if (next == '"') sb.Append('"');
                        else if (next == '\\') sb.Append('\\');
                        else if (next == 'n') sb.Append('\n');
                        else if (next == 'r') sb.Append('\r');
                        else if (next == 't') sb.Append('\t');
                        else sb.Append(next);
                        i++;
                    }
                    else if (ch == '"')
                    {
                        break;
                    }
                    else
                    {
                        sb.Append(ch);
                    }
                }

                string result = sb.ToString().Trim();
                return string.IsNullOrEmpty(result) ? json : result;
            }
            catch
            {
                return json;
            }
        }
    }
}
