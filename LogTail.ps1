# Reads only newly appended bytes. Decoder state survives partial Unicode writes.
if (-not ('WordPdfLogTail' -as [type])) {
 Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
public sealed class WordPdfLogTail {
    readonly string path;
    long offset;
    Decoder decoder;
    public bool HasMore { get; private set; }
    public long BytesRead { get; private set; }
    public WordPdfLogTail(string path) { this.path = path; }
    public string ReadNew(bool finished) {
        if (!File.Exists(path)) { HasMore = false; return ""; }
        using (var file = new FileStream(path, FileMode.Open, FileAccess.Read,
               FileShare.ReadWrite | FileShare.Delete)) {
            if (file.Length < offset) { offset = 0; decoder = null; }
            file.Position = offset;
            long available = file.Length - offset;
            // Wait for enough bytes to recognize a byte-order mark.
            if (decoder == null && available < 4 && !finished) {
                HasMore = available > 0; return "";
            }
            byte[] bytes = new byte[(int)Math.Min(65536, available)];
            int count = file.Read(bytes, 0, bytes.Length);
            offset += count; BytesRead += count;
            int skip = 0;
            if (decoder == null) {
                Encoding encoding = new UTF8Encoding(false);
                if (count >= 4 && bytes[0]==0xff && bytes[1]==0xfe && bytes[2]==0 && bytes[3]==0) {
                    encoding = new UTF32Encoding(false, true); skip=4;
                } else if (count >= 3 && bytes[0]==0xef && bytes[1]==0xbb && bytes[2]==0xbf) {
                    skip=3;
                } else if (count >= 2 && bytes[0]==0xff && bytes[1]==0xfe) {
                    encoding=Encoding.Unicode; skip=2;
                } else if (count >= 2 && bytes[0]==0xfe && bytes[1]==0xff) {
                    encoding=Encoding.BigEndianUnicode; skip=2;
                }
                decoder=encoding.GetDecoder();
            }
            HasMore = file.Length > offset;
            char[] chars = new char[(count + 4) * 2];
            int used = decoder.GetChars(bytes, skip, count-skip, chars, 0, finished && !HasMore);
            return new string(chars, 0, used);
        }
    }
}
'@
}
