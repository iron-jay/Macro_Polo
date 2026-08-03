using System;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Macro_Polo.Core;
using Xunit;

namespace Macro_Polo.Core.Tests
{
    /// <summary>
    /// Covers the trust decision. The file parsing itself is exercised against real signed
    /// documents by hand, since a committed fixture would mean committing a signing certificate.
    /// </summary>
    public class VbaSignatureReaderTests
    {
        [Fact]
        public void A_certificate_that_is_not_in_trusted_publishers_is_not_trusted()
        {
            using (X509Certificate2 certificate = CreateSelfSigned("CN=Macro Polo Test, O=Nowhere"))
            {
                VbaSignature signature = VbaSignatureReader.Evaluate(certificate);

                Assert.Equal(PublisherTrust.NotTrusted, signature.Trust);
                Assert.Equal(certificate.Thumbprint, signature.Thumbprint);
                Assert.Contains("Trusted Publishers", signature.UntrustedReason);
            }
        }

        /// <summary>
        /// The reason is shown to the user, so it has to say something they can act on rather than
        /// just repeating that it did not work.
        /// </summary>
        [Fact]
        public void An_untrusted_certificate_explains_itself()
        {
            using (X509Certificate2 certificate = CreateSelfSigned("CN=Macro Polo Test"))
            {
                VbaSignature signature = VbaSignatureReader.Evaluate(certificate);

                Assert.False(string.IsNullOrWhiteSpace(signature.UntrustedReason));
                Assert.Equal("Macro Polo Test", signature.SignerName);
            }
        }

        [Fact]
        public void A_missing_file_yields_unknown_rather_than_throwing()
        {
            VbaSignature signature = VbaSignatureReader.Read(@"C:\does\not\exist\nothing.docm");

            Assert.Equal(PublisherTrust.Unknown, signature.Trust);
            Assert.Null(signature.SignerName);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void An_absent_path_yields_unknown(string path)
        {
            Assert.Equal(PublisherTrust.Unknown, VbaSignatureReader.Read(path).Trust);
        }

        /// <summary>
        /// A document that is not an OOXML package - a legacy .doc, say - is reported as unknown
        /// rather than guessed at, because the signature lives somewhere this does not read.
        /// </summary>
        [Fact]
        public void A_file_that_is_not_a_package_yields_unknown()
        {
            string path = System.IO.Path.GetTempFileName();

            try
            {
                System.IO.File.WriteAllBytes(path, new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 });

                Assert.Equal(PublisherTrust.Unknown, VbaSignatureReader.Read(path).Trust);
            }
            finally
            {
                System.IO.File.Delete(path);
            }
        }

        private static X509Certificate2 CreateSelfSigned(string subject)
        {
            using (RSA key = RSA.Create(2048))
            {
                var request = new CertificateRequest(subject, key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

                return request.CreateSelfSigned(
                    DateTimeOffset.UtcNow.AddDays(-1),
                    DateTimeOffset.UtcNow.AddDays(1));
            }
        }
    }
}
