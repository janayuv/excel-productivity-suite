using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Engine.Updates
{
    /// <summary>
    /// Lightweight, zero-dependency parser for GitHub Releases API JSON payloads.
    /// </summary>
    public static class GitHubReleaseParser
    {
        public static UpdateInfo Parse(string json, string currentVersion)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return UpdateInfo.CreateError(currentVersion, "Empty response from GitHub API.");
            }

            try
            {
                string tagName = ExtractJsonStringField(json, "tag_name");
                string releaseName = ExtractJsonStringField(json, "name");
                string htmlUrl = ExtractJsonStringField(json, "html_url");
                string publishedAtRaw = ExtractJsonStringField(json, "published_at");
                string body = ExtractJsonStringField(json, "body");

                DateTime? publishedAt = null;
                if (!string.IsNullOrEmpty(publishedAtRaw) &&
                    DateTime.TryParse(publishedAtRaw, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal, out var dt))
                {
                    publishedAt = dt;
                }

                string latestVersion = NormalizeVersionTag(tagName);
                bool isNewer = IsNewerVersion(currentVersion, latestVersion);

                var assets = ExtractAssets(json);

                string? downloadUrl = null;
                string? checksumUrl = null;
                string assetName = string.Empty;
                long assetSize = 0;

                foreach (var asset in assets)
                {
                    if (asset.Name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                    {
                        downloadUrl = asset.DownloadUrl;
                        assetName = asset.Name;
                        assetSize = asset.Size;
                    }
                    else if (asset.Name.Equals("SHA256SUMS.txt", StringComparison.OrdinalIgnoreCase))
                    {
                        checksumUrl = asset.DownloadUrl;
                    }
                }

                return new UpdateInfo
                {
                    CurrentVersion = currentVersion,
                    LatestVersion = latestVersion,
                    ReleaseName = string.IsNullOrWhiteSpace(releaseName) ? tagName : releaseName,
                    ReleaseNotes = body,
                    HtmlUrl = htmlUrl,
                    DownloadUrl = downloadUrl,
                    ChecksumUrl = checksumUrl,
                    AssetName = assetName,
                    AssetSizeBytes = assetSize,
                    PublishedAt = publishedAt,
                    IsUpdateAvailable = isNewer
                };
            }
            catch (Exception ex)
            {
                return UpdateInfo.CreateError(currentVersion, $"Failed to parse GitHub release metadata: {ex.Message}");
            }
        }

        public static string NormalizeVersionTag(string tag)
        {
            if (string.IsNullOrWhiteSpace(tag)) return string.Empty;
            string t = tag.Trim();
            if (t.StartsWith("v", StringComparison.OrdinalIgnoreCase))
            {
                t = t.Substring(1);
            }
            return t.Trim();
        }

        public static bool IsNewerVersion(string currentVersion, string latestVersion)
        {
            string cur = NormalizeVersionTag(currentVersion);
            string lat = NormalizeVersionTag(latestVersion);

            if (string.IsNullOrWhiteSpace(cur) || string.IsNullOrWhiteSpace(lat))
                return false;

            if (Version.TryParse(cur, out var vCur) && Version.TryParse(lat, out var vLat))
            {
                return vLat > vCur;
            }

            string[] curParts = cur.Split('.');
            string[] latParts = lat.Split('.');
            int maxLen = Math.Max(curParts.Length, latParts.Length);

            for (int i = 0; i < maxLen; i++)
            {
                int cVal = i < curParts.Length && int.TryParse(curParts[i], NumberStyles.Integer, CultureInfo.InvariantCulture, out int cv) ? cv : 0;
                int lVal = i < latParts.Length && int.TryParse(latParts[i], NumberStyles.Integer, CultureInfo.InvariantCulture, out int lv) ? lv : 0;
                if (lVal > cVal) return true;
                if (lVal < cVal) return false;
            }

            return false;
        }

        public static string ExtractJsonStringField(string json, string fieldName)
        {
            var match = Regex.Match(json, $"\"{Regex.Escape(fieldName)}\"\\s*:\\s*\"((?:\\\\\"|[^\"])*)\"");
            if (match.Success)
            {
                return UnescapeJsonString(match.Groups[1].Value);
            }
            return string.Empty;
        }

        public static long ExtractJsonLongField(string json, string fieldName)
        {
            var match = Regex.Match(json, $"\"{Regex.Escape(fieldName)}\"\\s*:\\s*(\\d+)");
            if (match.Success && long.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long val))
            {
                return val;
            }
            return 0;
        }

        public static List<(string Name, long Size, string DownloadUrl)> ExtractAssets(string json)
        {
            var list = new List<(string Name, long Size, string DownloadUrl)>();
            if (string.IsNullOrEmpty(json)) return list;

            int assetsKeyIdx = json.IndexOf("\"assets\"", StringComparison.OrdinalIgnoreCase);
            if (assetsKeyIdx < 0) return list;

            int openBracket = json.IndexOf('[', assetsKeyIdx);
            if (openBracket < 0) return list;

            int depth = 0;
            int closeBracket = -1;
            bool inQuote = false;
            for (int i = openBracket; i < json.Length; i++)
            {
                char c = json[i];
                if (c == '\\' && inQuote)
                {
                    i++;
                    continue;
                }
                if (c == '"')
                {
                    inQuote = !inQuote;
                    continue;
                }
                if (!inQuote)
                {
                    if (c == '[') depth++;
                    else if (c == ']')
                    {
                        depth--;
                        if (depth == 0)
                        {
                            closeBracket = i;
                            break;
                        }
                    }
                }
            }

            if (closeBracket < 0) return list;

            string assetsBlock = json.Substring(openBracket + 1, closeBracket - openBracket - 1);

            int objDepth = 0;
            int objStart = -1;
            bool objInQuote = false;
            for (int i = 0; i < assetsBlock.Length; i++)
            {
                char c = assetsBlock[i];
                if (c == '\\' && objInQuote)
                {
                    i++;
                    continue;
                }
                if (c == '"')
                {
                    objInQuote = !objInQuote;
                    continue;
                }
                if (!objInQuote)
                {
                    if (c == '{')
                    {
                        if (objDepth == 0) objStart = i;
                        objDepth++;
                    }
                    else if (c == '}')
                    {
                        objDepth--;
                        if (objDepth == 0 && objStart >= 0)
                        {
                            string itemJson = assetsBlock.Substring(objStart, i - objStart + 1);
                            string name = ExtractJsonStringField(itemJson, "name");
                            long size = ExtractJsonLongField(itemJson, "size");
                            string url = ExtractJsonStringField(itemJson, "browser_download_url");
                            if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(url))
                            {
                                list.Add((name, size, url));
                            }
                            objStart = -1;
                        }
                    }
                }
            }

            return list;
        }

        public static string UnescapeJsonString(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return string.Empty;
            var sb = new StringBuilder(raw.Length);
            for (int i = 0; i < raw.Length; i++)
            {
                if (raw[i] == '\\' && i + 1 < raw.Length)
                {
                    char next = raw[i + 1];
                    switch (next)
                    {
                        case 'n': sb.Append('\n'); i++; break;
                        case 'r': sb.Append('\r'); i++; break;
                        case 't': sb.Append('\t'); i++; break;
                        case '"': sb.Append('"'); i++; break;
                        case '\\': sb.Append('\\'); i++; break;
                        case '/': sb.Append('/'); i++; break;
                        case 'u':
                            if (i + 5 < raw.Length && int.TryParse(raw.Substring(i + 2, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int unicodeChar))
                            {
                                sb.Append((char)unicodeChar);
                                i += 5;
                            }
                            else
                            {
                                sb.Append('\\');
                            }
                            break;
                        default:
                            sb.Append(next);
                            i++;
                            break;
                    }
                }
                else
                {
                    sb.Append(raw[i]);
                }
            }
            return sb.ToString();
        }
    }
}
