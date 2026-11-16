using System;
using System.Text;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // PowerPoint records Presentation.Tags in its native Undo history. Unlike
    // CustomXMLParts, this data can share one transaction with shape geometry.
    public static class NativeConstraintStore
    {
        public const string TagName = "PPTGCS_CONSTRAINTS";
        public const int MaxXmlBytes = 1024 * 1024;
        private const string LegacyNamespace = "urn:ppt-gcs:v2";
        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false, true);

        public static ConstraintDocument Read(object presentationObject)
        {
            if (presentationObject == null) throw new ArgumentNullException("presentationObject");
            dynamic presentation = presentationObject;
            string encoded = (string)presentation.Tags[TagName];
            if (String.IsNullOrEmpty(encoded))
            {
                if ((int)presentation.CustomXMLParts.SelectByNamespace(LegacyNamespace).Count != 0)
                    throw new InvalidOperationException("此文稿使用旧的约束存储格式，当前版本不直接读取；请先保存备份并重新建立约束。");
                return new ConstraintDocument(new LayoutConstraint[0]);
            }
            return ConstraintDocument.Parse(DecodeXml(encoded));
        }

        public static void Write(object presentationObject, ConstraintDocument document)
        {
            if (presentationObject == null) throw new ArgumentNullException("presentationObject");
            if (document == null) throw new ArgumentNullException("document");
            string encoded = EncodeXml(document.ToXml());
            dynamic tags = ((dynamic)presentationObject).Tags;
            string before = (string)tags[TagName];
            if (String.Equals(before, encoded, StringComparison.Ordinal)) return;
            try
            {
                tags.Add(TagName, encoded);
                if (!String.Equals((string)tags[TagName], encoded, StringComparison.Ordinal))
                    throw new InvalidOperationException("PowerPoint 未完整保存约束数据，操作已取消。");
            }
            catch (Exception writeError)
            {
                try
                {
                    if (String.IsNullOrEmpty(before)) tags.Delete(TagName);
                    else tags.Add(TagName, before);
                    if (!String.Equals((string)tags[TagName], before ?? String.Empty, StringComparison.Ordinal))
                        throw new InvalidOperationException("PowerPoint 无法恢复操作前的约束标签。");
                }
                catch (Exception restoreError)
                {
                    throw new AggregateException("约束保存失败，且原标签未能恢复。", writeError, restoreError);
                }
                throw;
            }
        }

        // Tag values are uppercased by Office, so XML and ordinary Base64 are
        // unsuitable. Uppercase hexadecimal preserves the exact UTF-8 bytes.
        public static string EncodeXml(string xml)
        {
            if (xml == null) throw new ArgumentNullException("xml");
            if (xml.Length == 0 || xml.Length > MaxXmlBytes)
                throw new FormatException("约束 XML 为空或超过 1 MiB 限制。");
            byte[] bytes;
            try { bytes = Utf8.GetBytes(xml); }
            catch (EncoderFallbackException error) { throw new FormatException("约束 XML 包含无效的 Unicode 字符。", error); }
            if (bytes.Length > MaxXmlBytes)
                throw new FormatException("约束 XML 超过 1 MiB UTF-8 限制。");
            return BitConverter.ToString(bytes).Replace("-", String.Empty);
        }

        public static string DecodeXml(string encoded)
        {
            if (encoded == null) throw new ArgumentNullException("encoded");
            if (encoded.Length == 0 || (encoded.Length & 1) != 0 || encoded.Length > MaxXmlBytes * 2)
                throw new FormatException("约束标签为空、长度无效或超过 1 MiB 限制。");
            var bytes = new byte[encoded.Length / 2];
            for (int i = 0; i < bytes.Length; i++)
                bytes[i] = (byte)((HexDigit(encoded[i * 2]) << 4) | HexDigit(encoded[i * 2 + 1]));
            try { return Utf8.GetString(bytes); }
            catch (DecoderFallbackException error) { throw new FormatException("约束标签不是有效的 UTF-8 数据。", error); }
        }

        private static int HexDigit(char value)
        {
            if (value >= '0' && value <= '9') return value - '0';
            if (value >= 'A' && value <= 'F') return value - 'A' + 10;
            throw new FormatException("约束标签包含无效的十六进制字符。");
        }
    }
}
