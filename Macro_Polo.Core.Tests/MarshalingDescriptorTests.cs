using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
using Macro_Polo.Core;
using Xunit;

namespace Macro_Polo.Core.Tests
{
    /// <summary>
    /// Compares the marshalling of our hand-written <see cref="IDTExtensibility2"/> against the
    /// real one, parameter by parameter.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the check that was missing. Ordinary reflection cannot see <c>MarshalAs</c> - it is
    /// a pseudo-custom-attribute stored in the FieldMarshal metadata table - so a wrong or absent
    /// descriptor is invisible to the compiler, to unit tests, and to a COM interface probe. It
    /// only shows up as the host process dying inside a marshalling stub.
    /// </para>
    /// <para>
    /// Omitting the SAFEARRAY descriptor on <c>custom</c> made the CLR marshal it as a COM
    /// interface instead, which crashed Word and Excel on every launch. Reading the FieldMarshal
    /// blobs out of both assemblies and comparing the bytes catches that class of defect exactly.
    /// </para>
    /// </remarks>
    public class MarshalingDescriptorTests
    {
        private const string ExtensibilityPia =
            @"C:\Windows\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll";

        private const string SafeArrayOfVariant = "SAFEARRAY(vt=0x0C)";
        private const string IDispatch = "0x1A";

        [Fact]
        public void The_custom_parameter_is_marshalled_as_a_safearray_of_variant_on_every_member()
        {
            Dictionary<string, List<ParameterMarshalling>> ours = ReadMarshalling(
                typeof(IDTExtensibility2).Assembly.Location, "IDTExtensibility2");

            Assert.NotEmpty(ours);

            foreach (KeyValuePair<string, List<ParameterMarshalling>> method in ours)
            {
                ParameterMarshalling custom = method.Value.SingleOrDefault(p => p.Name == "custom");

                Assert.True(custom != null, method.Key + " has no 'custom' parameter");
                Assert.Equal(SafeArrayOfVariant, Describe(custom.Descriptor));
            }
        }

        [Fact]
        public void The_object_parameters_are_marshalled_as_idispatch()
        {
            List<ParameterMarshalling> onConnection = ReadMarshalling(
                typeof(IDTExtensibility2).Assembly.Location, "IDTExtensibility2")["OnConnection"];

            foreach (string name in new[] { "application", "addInInst" })
            {
                ParameterMarshalling parameter = onConnection.Single(p => p.Name == name);

                Assert.Equal(IDispatch, Describe(parameter.Descriptor));
            }
        }

        /// <summary>
        /// Where the primary interop assembly is installed, compare against it directly rather
        /// than against constants. Skipped when it is absent.
        /// </summary>
        [Fact]
        public void Our_marshalling_matches_the_primary_interop_assembly_when_present()
        {
            if (!File.Exists(ExtensibilityPia))
            {
                return;
            }

            Dictionary<string, List<ParameterMarshalling>> official = ReadMarshalling(ExtensibilityPia, "IDTExtensibility2");
            Dictionary<string, List<ParameterMarshalling>> ours = ReadMarshalling(typeof(IDTExtensibility2).Assembly.Location, "IDTExtensibility2");

            Assert.Equal(official.Keys.OrderBy(k => k), ours.Keys.OrderBy(k => k));

            foreach (string method in official.Keys)
            {
                List<ParameterMarshalling> expected = official[method];
                List<ParameterMarshalling> actual = ours[method];

                Assert.Equal(expected.Count, actual.Count);

                for (int i = 0; i < expected.Count; i++)
                {
                    Assert.True(
                        Describe(expected[i].Descriptor) == Describe(actual[i].Descriptor),
                        method + " parameter " + (i + 1) + " marshalling differs. Expected "
                            + Describe(expected[i].Descriptor) + ", got " + Describe(actual[i].Descriptor));
                }
            }
        }

        /// <summary>
        /// Reduces a FieldMarshal blob to what it actually means.
        /// </summary>
        /// <remarks>
        /// Comparing the raw bytes is too strict: for a SAFEARRAY the blob ends with an optional
        /// user-defined subtype name, and the primary interop assembly emits it as an empty string
        /// (a trailing 0x00) where the C# compiler omits it entirely. Those encode the same thing,
        /// so the comparison is on the native type and element type instead.
        /// </remarks>
        private static string Describe(byte[] descriptor)
        {
            if (descriptor == null || descriptor.Length == 0)
            {
                return "(none)";
            }

            const byte NativeTypeSafeArray = 0x1D;

            if (descriptor[0] == NativeTypeSafeArray)
            {
                byte elementType = descriptor.Length > 1 ? descriptor[1] : (byte)0;
                return "SAFEARRAY(vt=0x" + elementType.ToString("X2") + ")";
            }

            return "0x" + descriptor[0].ToString("X2");
        }

        private sealed class ParameterMarshalling
        {
            public string Name;
            public byte[] Descriptor;
        }

        /// <summary>
        /// Reads the FieldMarshal blob for every parameter of every method on the named interface.
        /// </summary>
        private static Dictionary<string, List<ParameterMarshalling>> ReadMarshalling(string assemblyPath, string typeName)
        {
            var result = new Dictionary<string, List<ParameterMarshalling>>(StringComparer.Ordinal);

            using (var stream = File.OpenRead(assemblyPath))
            using (var pe = new PEReader(stream))
            {
                MetadataReader md = pe.GetMetadataReader();

                foreach (TypeDefinitionHandle typeHandle in md.TypeDefinitions)
                {
                    TypeDefinition type = md.GetTypeDefinition(typeHandle);
                    if (md.GetString(type.Name) != typeName)
                    {
                        continue;
                    }

                    foreach (MethodDefinitionHandle methodHandle in type.GetMethods())
                    {
                        MethodDefinition method = md.GetMethodDefinition(methodHandle);
                        var parameters = new List<ParameterMarshalling>();

                        foreach (ParameterHandle parameterHandle in method.GetParameters())
                        {
                            Parameter parameter = md.GetParameter(parameterHandle);
                            if (parameter.SequenceNumber == 0)
                            {
                                continue; // the return value
                            }

                            BlobHandle blob = parameter.GetMarshallingDescriptor();

                            parameters.Add(new ParameterMarshalling
                            {
                                Name = parameter.Name.IsNil ? null : md.GetString(parameter.Name),
                                Descriptor = blob.IsNil ? null : md.GetBlobBytes(blob)
                            });
                        }

                        result[md.GetString(method.Name)] = parameters;
                    }
                }
            }

            return result;
        }
    }
}
