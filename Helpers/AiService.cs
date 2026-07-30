using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSheetManager.Helpers
{
    public static class AiService
    {
        public static string GetEndpointUrl()
        {
            string baseUrl = SettingsHelper.GetAiBaseUrl().Trim();
            if (string.IsNullOrEmpty(baseUrl)) baseUrl = "http://localhost:11434/v1";

            baseUrl = baseUrl.TrimEnd('/');

            if (baseUrl.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
            {
                return baseUrl;
            }
            if (baseUrl.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
            {
                return baseUrl + "/chat/completions";
            }

            return baseUrl + "/v1/chat/completions";
        }

        public static async Task<bool> TestConnectionAsync()
        {
            try
            {
                string result = await GetCompletionAsync("ping", "Respond with pong only.");
                return !string.IsNullOrEmpty(result) && !result.StartsWith("Error", StringComparison.OrdinalIgnoreCase) && !result.StartsWith("AI Server Error", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
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
                    if (webEx.Response != null)
                    {
                        using (var resStream = webEx.Response.GetResponseStream())
                        using (var reader = new StreamReader(resStream, Encoding.UTF8))
                        {
                            return $"AI Server Error: {reader.ReadToEnd()}";
                        }
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
