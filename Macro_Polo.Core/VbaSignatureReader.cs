using System;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;

namespace Macro_Polo.Core
{
    /// <summary>
    /// Reads the certificate a document's VBA project was signed with, and decides whether this
    /// machine trusts it.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Office tells add-ins only that a signature exists, never who signed it. Without this, the
    /// add-in cannot tell a macro that will sit behind the trust bar from one that has already run
    /// - and it reported the latter as the former, which is the wrong way round for a tool whose
    /// job is to say whether code executed.
    /// </para>
    /// <para>
    /// For the OOXML formats the signature is not inside the VBA project at all: it sits alongside
    /// it as separate package parts, so reading it needs nothing more than the zip reader and
    /// <see cref="SignedCms"/>. Word writes up to three of them - legacy, agile and V3 - all
    /// carrying the same signer, so the first one that decodes is enough.
    /// </para>
    /// <para>
    /// The legacy binary formats keep the signature inside the compound file instead, which needs
    /// a reader this does not have. Those come back <see cref="PublisherTrust.Unknown"/>, which
    /// reports exactly what the add-in did before and never overstates.
    /// </para>
    /// </remarks>
    public static class VbaSignatureReader
    {
        /// <summary>Package parts Word and Excel write the VBA signature to, best first.</summary>
        private static readonly string[] SignatureParts =
        {
            "word/vbaProjectSignatureV3.bin",
            "xl/vbaProjectSignatureV3.bin",
            "word/vbaProjectSignatureAgile.bin",
            "xl/vbaProjectSignatureAgile.bin",
            "word/vbaProjectSignature.bin",
            "xl/vbaProjectSignature.bin"
        };

        /// <summary>
        /// Resolves the signature on the document at <paramref name="fullPath"/>. Never throws:
        /// anything it cannot establish comes back as <see cref="VbaSignature.Unknown"/>.
        /// </summary>
        public static VbaSignature Read(string fullPath)
        {
            X509Certificate2 certificate = null;

            try
            {
                certificate = ReadCertificate(fullPath);
                if (certificate == null)
                {
                    return VbaSignature.Unknown;
                }

                return Evaluate(certificate);
            }
            catch (Exception ex)
            {
                Log.Warn("Could not read the VBA signature from " + fullPath, ex);
                return VbaSignature.Unknown;
            }
            finally
            {
                if (certificate != null)
                {
                    certificate.Reset();
                }
            }
        }

        /// <summary>
        /// Decides whether Office would accept <paramref name="certificate"/> without asking.
        /// </summary>
        /// <remarks>
        /// Membership of Trusted Publishers is necessary but not sufficient: the chain has to
        /// validate as well. A self-signed certificate added to Trusted Publishers, but whose root
        /// is not trusted, is still refused - a state worth naming, because it looks like it should
        /// work and does not.
        /// </remarks>
        internal static VbaSignature Evaluate(X509Certificate2 certificate)
        {
            string signer = GetCommonName(certificate);
            string thumbprint = certificate.Thumbprint;

            if (!IsInTrustedPublishers(thumbprint))
            {
                return new VbaSignature(PublisherTrust.NotTrusted, signer, thumbprint,
                    "the certificate is not in Trusted Publishers");
            }

            using (var chain = new X509Chain())
            {
                chain.ChainPolicy.RevocationMode = X509RevocationMode.NoCheck;

                if (!chain.Build(certificate))
                {
                    string reason = chain.ChainStatus.Length > 0
                        ? "the certificate chain is not valid (" + chain.ChainStatus[0].Status + ")"
                        : "the certificate chain is not valid";

                    return new VbaSignature(PublisherTrust.NotTrusted, signer, thumbprint, reason);
                }
            }

            return new VbaSignature(PublisherTrust.Trusted, signer, thumbprint, null);
        }

        private static bool IsInTrustedPublishers(string thumbprint)
        {
            // Office honours both the per-user and the machine-wide store.
            return IsInStore(StoreLocation.CurrentUser, thumbprint)
                || IsInStore(StoreLocation.LocalMachine, thumbprint);
        }

        private static bool IsInStore(StoreLocation location, string thumbprint)
        {
            var store = new X509Store(StoreName.TrustedPublisher, location);

            try
            {
                store.Open(OpenFlags.ReadOnly);

                foreach (X509Certificate2 candidate in store.Certificates)
                {
                    if (string.Equals(candidate.Thumbprint, thumbprint, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warn("Could not read the " + location + " Trusted Publishers store", ex);
            }
            finally
            {
                store.Close();
            }

            return false;
        }

        private static X509Certificate2 ReadCertificate(string fullPath)
        {
            if (string.IsNullOrEmpty(fullPath) || !File.Exists(fullPath))
            {
                return null;
            }

            // The host has the document open, so the file has to be shared on the way in.
            using (var file = new FileStream(fullPath, FileMode.Open, FileAccess.Read,
                       FileShare.ReadWrite | FileShare.Delete))
            {
                if (!LooksLikeZip(file))
                {
                    // A legacy compound-file document. Not parsed; see the remarks above.
                    return null;
                }

                using (var archive = new ZipArchive(file, ZipArchiveMode.Read, leaveOpen: true))
                {
                    foreach (string partName in SignatureParts)
                    {
                        ZipArchiveEntry entry = archive.GetEntry(partName);
                        if (entry == null)
                        {
                            continue;
                        }

                        X509Certificate2 certificate = ReadCertificateFromPart(entry);
                        if (certificate != null)
                        {
                            return certificate;
                        }
                    }
                }
            }

            return null;
        }

        private static X509Certificate2 ReadCertificateFromPart(ZipArchiveEntry entry)
        {
            byte[] part;
            using (Stream stream = entry.Open())
            using (var buffer = new MemoryStream())
            {
                stream.CopyTo(buffer);
                part = buffer.ToArray();
            }

            // The part begins with a DigSigInfoSerialized header and the PKCS#7 blob follows it.
            // Rather than walk that header, find the DER SEQUENCE that starts the blob and let
            // SignedCms confirm it: a wrong offset simply fails to decode.
            for (int offset = 0; offset < Math.Min(part.Length - 4, 512); offset++)
            {
                if (part[offset] != 0x30 || (part[offset + 1] != 0x82 && part[offset + 1] != 0x83))
                {
                    continue;
                }

                try
                {
                    var blob = new byte[part.Length - offset];
                    Buffer.BlockCopy(part, offset, blob, 0, blob.Length);

                    var signed = new SignedCms();
                    signed.Decode(blob);

                    if (signed.SignerInfos.Count > 0 && signed.SignerInfos[0].Certificate != null)
                    {
                        return signed.SignerInfos[0].Certificate;
                    }
                }
                catch (Exception)
                {
                    // Not the start of the blob; keep looking.
                }
            }

            Log.Warn("Found " + entry.FullName + " but could not decode a signer from it", null);
            return null;
        }

        /// <summary>True for the OOXML formats, which are zip packages.</summary>
        private static bool LooksLikeZip(Stream stream)
        {
            long position = stream.Position;

            try
            {
                int first = stream.ReadByte();
                int second = stream.ReadByte();
                return first == 'P' && second == 'K';
            }
            finally
            {
                stream.Position = position;
            }
        }

        /// <summary>The signer's common name, falling back to the whole subject.</summary>
        private static string GetCommonName(X509Certificate2 certificate)
        {
            try
            {
                string name = certificate.GetNameInfo(X509NameType.SimpleName, false);
                return string.IsNullOrEmpty(name) ? certificate.Subject : name;
            }
            catch (Exception ex)
            {
                Log.Warn("Could not read the signer name", ex);
                return certificate.Subject;
            }
        }
    }
}
