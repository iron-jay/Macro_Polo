namespace Macro_Polo.Core
{
    /// <summary>Whether the certificate a macro was signed with is one this machine trusts.</summary>
    public enum PublisherTrust
    {
        /// <summary>
        /// Could not be determined: the file is not on disk, is a format we do not parse, was
        /// locked, or carried no signature we could read. Reported honestly rather than guessed.
        /// </summary>
        Unknown,

        /// <summary>
        /// The signing certificate is in Trusted Publishers and its chain validates. Office will
        /// run macros signed with it without prompting.
        /// </summary>
        Trusted,

        /// <summary>
        /// The signature was read, but the certificate is not one Office will accept without the
        /// user first agreeing to trust it.
        /// </summary>
        NotTrusted
    }

    /// <summary>What could be established about the signature on a document's VBA project.</summary>
    public sealed class VbaSignature
    {
        /// <summary>Nothing could be established.</summary>
        public static readonly VbaSignature Unknown = new VbaSignature(PublisherTrust.Unknown, null, null, null);

        public VbaSignature(PublisherTrust trust, string signerName, string thumbprint, string untrustedReason)
        {
            Trust = trust;
            SignerName = signerName;
            Thumbprint = thumbprint;
            UntrustedReason = untrustedReason;
        }

        public PublisherTrust Trust { get; private set; }

        /// <summary>Common name of the signer, for display. Null when unknown.</summary>
        public string SignerName { get; private set; }

        public string Thumbprint { get; private set; }

        /// <summary>
        /// Why the certificate was rejected, when it was. A self-signed certificate that has been
        /// added to Trusted Publishers but not to Trusted Roots still fails, and saying which of
        /// the two is missing saves a lot of guessing.
        /// </summary>
        public string UntrustedReason { get; private set; }
    }
}
