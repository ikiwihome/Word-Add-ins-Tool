using System;
using UtfUnknown;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;

namespace Word_AddIns
{

    public static class EncodingUtility
    {
        /// <summary>
        /// 默认的块大小(10KB)，用于分块读取文件
        /// </summary>
        public const int DefaultChunkSize = 10000;

        public delegate void ChunkHandler(string chunk, double progress);

        private static Encoding _systemDefaultANSIEncoding;

        private static Encoding _currentCultureANSIEncoding;

        public static string ReadTextFileWithEncoding(string filePath, Encoding encoding = null)
        {
            return ReadTextFileWithEncodingAsync(filePath, encoding, null).GetAwaiter().GetResult();
        }

        public static async Task<string> ReadTextFileWithEncodingAsync(string filePath, Encoding encoding = null, IProgress<double> progress = null)
        {
            string text;
            var bom = new byte[4];

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                await stream.ReadAsync(bom, 0, 4).ConfigureAwait(false); // Read BOM values
                stream.Position = 0; // Reset stream position

                var reader = CreateStreamReader(stream, bom, encoding);

                async Task<string> PeekAndReadAsync()
                {
                    if (encoding == null)
                    {
                        await Task.Run(() => reader.Peek()).ConfigureAwait(false);
                        encoding = reader.CurrentEncoding;
                        Debug.WriteLine($"Detected encoding: {encoding.EncodingName}");
                    }

                    var buffer = new char[4096];
                    var result = new StringBuilder();
                    var totalBytes = stream.Length;
                    var readBytes = 0L;

                    while (true)
                    {
                        var readCount = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
                        if (readCount == 0) break;

                        result.Append(buffer, 0, readCount);
                        readBytes += readCount;

                        var percentage = (double)readBytes / totalBytes * 100;
                        progress?.Report(percentage);
                        Debug.WriteLine($"Read progress: {percentage:F2}% (Bytes: {readBytes}/{totalBytes})");
                    }

                    reader.Close();
                    return result.ToString();
                }

                try
                {
                    text = await PeekAndReadAsync().ConfigureAwait(false);
                }
                catch (DecoderFallbackException)
                {
                    Debug.WriteLine("DecoderFallbackException occurred, trying fallback encoding");
                    stream.Position = 0; // Reset stream position
                    encoding = GetFallBackEncoding();
                    reader = new StreamReader(stream, encoding);
                    text = await PeekAndReadAsync().ConfigureAwait(false);
                }
            }

            Debug.WriteLine("File read completed successfully");
            return text;
        }

        public static void ReadTextFileWithEncodingChunked(string filePath, ChunkHandler chunkHandler, Encoding encoding = null, int chunkSize = DefaultChunkSize)
        {
            ReadTextFileWithEncodingChunkedAsync(filePath, chunkHandler, encoding, null, chunkSize).GetAwaiter().GetResult();
        }

        public static async Task ReadTextFileWithEncodingChunkedAsync(string filePath, ChunkHandler chunkHandler, Encoding encoding = null, IProgress<double> progress = null, int chunkSize = DefaultChunkSize)
        {
            var bom = new byte[4];
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                await stream.ReadAsync(bom, 0, 4).ConfigureAwait(false);
                stream.Position = 0;

                var reader = CreateStreamReader(stream, bom, encoding);
                if (encoding == null)
                {
                    await Task.Run(() => reader.Peek()).ConfigureAwait(false);
                    encoding = reader.CurrentEncoding;
                    Debug.WriteLine($"Detected encoding: {encoding.EncodingName}");
                }

                var totalBytes = stream.Length;
                var readBytes = 0L;
                var buffer = new char[chunkSize];
                var leftover = new StringBuilder();

                while (true)
                {
                    var readCount = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
                    if (readCount == 0) break;

                    var chunk = new StringBuilder(leftover.ToString());
                    leftover.Clear();

                    // Manually find last line break to avoid Array.FindLastIndex issues
                    int lastLineBreak = -1;
                    if (readCount > 0 && buffer.Length > 0)
                    {
                        int endPos = Math.Min(readCount, buffer.Length) - 1;
                        for (int i = endPos; i >= 0; i--)
                        {
                            if (buffer[i] == '\n' || buffer[i] == '\r')
                            {
                                lastLineBreak = i;
                                break;
                            }
                        }
                    }

                    if (lastLineBreak >= 0 && lastLineBreak < buffer.Length)
                    {
                        int chunkLength = Math.Min(lastLineBreak + 1, buffer.Length);
                        int leftoverLength = Math.Max(0, Math.Min(readCount - chunkLength, buffer.Length - chunkLength));

                        if (chunkLength > 0 && chunkLength <= buffer.Length)
                        {
                            chunk.Append(buffer, 0, chunkLength);
                        }
                        if (leftoverLength > 0 && leftoverLength <= buffer.Length - chunkLength)
                        {
                            leftover.Append(buffer, chunkLength, leftoverLength);
                        }
                    }
                    else if (readCount > 0 && buffer.Length > 0)
                    {
                        int appendLength = Math.Min(readCount, buffer.Length);
                        if (appendLength > 0)
                        {
                            chunk.Append(buffer, 0, appendLength);
                        }
                    }

                    readBytes += readCount;
                    var percentage = (double)readBytes / totalBytes * 100;

                    if (chunk.Length > 0)
                    {
                        chunkHandler?.Invoke(chunk.ToString(), percentage);
                    }

                    progress?.Report(percentage);
                    Debug.WriteLine($"Processed chunk: {percentage:F2}% (Bytes: {readBytes}/{totalBytes})");
                }

                // Process remaining content
                if (leftover.Length > 0)
                {
                    chunkHandler?.Invoke(leftover.ToString(), 100);
                }

                reader.Close();
                Debug.WriteLine("File read completed successfully");
            }
        }

        private static bool TryGetSystemDefaultANSIEncoding(out Encoding encoding)
        {
            if (_systemDefaultANSIEncoding != null)
            {
                encoding = _systemDefaultANSIEncoding;
                return true;
            }
            encoding = Encoding.GetEncoding(0);
            _systemDefaultANSIEncoding = encoding;
            return true;

        }
        private static bool TryGetCurrentCultureANSIEncoding(out Encoding encoding)
        {
            if (_currentCultureANSIEncoding != null)
            {
                encoding = _currentCultureANSIEncoding;
                return true;
            }
            encoding = Encoding.GetEncoding(Thread.CurrentThread.CurrentCulture.TextInfo.ANSICodePage);
            _currentCultureANSIEncoding = encoding;
            return true;
        }

        private static Encoding GetFallBackEncoding()
        {
            if (TryGetSystemDefaultANSIEncoding(out var systemDefaultEncoding))
            {
                return systemDefaultEncoding;
            }
            else if (TryGetCurrentCultureANSIEncoding(out var currentCultureEncoding))
            {
                return currentCultureEncoding;
            }
            else
            {
                return new UTF8Encoding(false);
            }
        }

        private static StreamReader CreateStreamReader(Stream stream, byte[] bom, Encoding encoding = null)
        {
            StreamReader reader;
            if (encoding != null)
            {
                reader = new StreamReader(stream, encoding);
            }
            else
            {
                if (HasBom(bom))
                {
                    reader = new StreamReader(stream);
                }
                else // No BOM, need to guess or use default decoding set by user
                {
                    var success = TryGuessEncoding(stream, out var autoEncoding);
                    stream.Position = 0; // Reset stream position
                    reader = success ?
                        new StreamReader(stream, autoEncoding) :
                        new StreamReader(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true));
                }
            }
            return reader;
        }

        private static bool TryGuessEncoding(Stream stream, out Encoding encoding)
        {
            encoding = null;

            var result = CharsetDetector.DetectFromStream(stream);
            if (result.Detected?.Encoding != null) // Detected can be null
            {
                encoding = AnalyzeAndGuessEncoding(result);
                return true;
            }
            return false;
        }

        private static Encoding AnalyzeAndGuessEncoding(DetectionResult result)
        {
            Encoding encoding = result.Detected.Encoding;
            var confidence = result.Detected.Confidence;
            var foundBetterMatch = false;

            // Let's treat ASCII as UTF-8 for better accuracy
            if (EncodingEquals(encoding, Encoding.ASCII)) encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

            // If confidence is above 80%, we should just use it
            if (confidence > 0.80f && result.Details.Count == 1) return encoding;

            // Try find a better match based on User's current Windows ANSI code page
            // Priority: UTF-8 > SystemDefaultANSIEncoding (Codepage: 0) > CurrentCultureANSIEncoding
            if (!(encoding is UTF8Encoding))
            {
                foreach (var detail in result.Details)
                {
                    if (detail.Confidence <= 0.5f)
                    {
                        continue;
                    }
                    if (detail.Encoding is UTF8Encoding)
                    {
                        foundBetterMatch = true;
                    }
                    else if (TryGetSystemDefaultANSIEncoding(out var systemDefaultEncoding)
                             && EncodingEquals(systemDefaultEncoding, detail.Encoding))
                    {
                        foundBetterMatch = true;
                    }
                    else if (TryGetCurrentCultureANSIEncoding(out var currentCultureEncoding)
                             && EncodingEquals(currentCultureEncoding, detail.Encoding))
                    {
                        foundBetterMatch = true;
                    }

                    if (foundBetterMatch)
                    {
                        encoding = detail.Encoding;
                        confidence = detail.Confidence;
                        break;
                    }
                }
            }

            // We should fall back to UTF-8 and give it a try if:
            // 1. Detected Encoding is not UTF-8
            // 2. Detected Encoding is not SystemDefaultANSIEncoding (Codepage: 0)
            // 3. Detected Encoding is not CurrentCultureANSIEncoding
            // 4. Confidence of detected Encoding is below 50%
            if (!foundBetterMatch && confidence < 0.5f)
            {
                encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
            }

            return encoding;
        }

        private static bool HasBom(byte[] bom)
        {
            // Analyze the BOM
            if (bom[0] == 0x2b && bom[1] == 0x2f && bom[2] == 0x76) return true; // Encoding.UTF7
            if (bom[0] == 0xef && bom[1] == 0xbb && bom[2] == 0xbf) return true; // Encoding.UTF8
            if (bom[0] == 0xff && bom[1] == 0xfe) return true; // Encoding.Unicode
            if (bom[0] == 0xfe && bom[1] == 0xff) return true; // Encoding.BigEndianUnicode
            if (bom[0] == 0 && bom[1] == 0 && bom[2] == 0xfe && bom[3] == 0xff) return true; // Encoding.UTF32
            return false;
        }

        private static bool EncodingEquals(Encoding p, Encoding q)
        {
            if (p.CodePage == q.CodePage)
            {
                if (q is UTF7Encoding ||
                    q is UTF8Encoding ||
                    q is UnicodeEncoding ||
                    q is UTF32Encoding)
                {
                    return Encoding.Equals(p, q); // To make sure we compare bigEndian and byteOrderMark flags
                }
                else
                {
                    return true;
                }
            }

            return false;
        }
    }
}
