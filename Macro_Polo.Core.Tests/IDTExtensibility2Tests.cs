using System;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Macro_Polo.Core;
using Xunit;

namespace Macro_Polo.Core.Tests
{
    /// <summary>
    /// Guards the hand-written <see cref="IDTExtensibility2"/> declaration against the real one.
    /// </summary>
    /// <remarks>
    /// Every one of these assertions is invisible to the compiler and fatal at runtime. Declaring
    /// the interface as IDispatch-only rather than dual exposed just the IDispatch vtable, so
    /// Office's first call landed past the end of the exposed slots and killed the host process
    /// outright - an access violation inside the CLR, with no managed exception and nothing in the
    /// add-in's own log. Reordering or renaming a member would do the same thing. None of it can be
    /// caught by building, so it is caught here.
    /// </remarks>
    public class IDTExtensibility2Tests
    {
        /// <summary>Path to the primary interop assembly, when the machine happens to have it.</summary>
        private const string ExtensibilityPia =
            @"C:\Windows\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll";

        private static readonly Guid ExpectedIid = new Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744");

        private static readonly string[] ExpectedMembers =
        {
            "OnConnection",
            "OnDisconnection",
            "OnAddInsUpdate",
            "OnStartupComplete",
            "OnBeginShutdown"
        };

        [Fact]
        public void The_interface_is_imported_from_com()
        {
            Assert.NotEmpty(typeof(IDTExtensibility2).GetCustomAttributes(typeof(ComImportAttribute), false));
        }

        [Fact]
        public void The_interface_id_matches_the_published_one()
        {
            Assert.Equal(ExpectedIid, typeof(IDTExtensibility2).GUID);
        }

        /// <summary>
        /// The bug that took down Word and Excel. Office calls through the vtable, which only a
        /// dual interface exposes.
        /// </summary>
        [Fact]
        public void The_interface_is_dual_rather_than_dispatch_only()
        {
            var attribute = (InterfaceTypeAttribute)typeof(IDTExtensibility2)
                .GetCustomAttributes(typeof(InterfaceTypeAttribute), false)
                .SingleOrDefault();

            // No attribute at all also means dual, which is how the real declaration is written.
            ComInterfaceType actual = attribute == null
                ? ComInterfaceType.InterfaceIsDual
                : attribute.Value;

            Assert.Equal(ComInterfaceType.InterfaceIsDual, actual);
        }

        /// <summary>Vtable slots are assigned in declaration order, so the order is the contract.</summary>
        [Fact]
        public void The_members_are_declared_in_the_published_order()
        {
            string[] actual = typeof(IDTExtensibility2).GetMethods().Select(m => m.Name).ToArray();

            Assert.Equal(ExpectedMembers, actual);
        }

        [Fact]
        public void Every_member_takes_the_custom_argument_by_reference()
        {
            foreach (MethodInfo method in typeof(IDTExtensibility2).GetMethods())
            {
                ParameterInfo last = method.GetParameters().Last();

                Assert.Equal("custom", last.Name);
                Assert.True(last.ParameterType.IsByRef, method.Name + " must take custom by reference");
                Assert.Equal(typeof(Array), last.ParameterType.GetElementType());
            }
        }

        /// <summary>
        /// Where the primary interop assembly is installed, compare against it directly. This is
        /// the check that actually caught the defect; it is skipped rather than failed on machines
        /// without the assembly, since it is a developer-machine convenience and not a requirement.
        /// </summary>
        [Fact]
        public void The_declaration_matches_the_primary_interop_assembly_when_present()
        {
            if (!System.IO.File.Exists(ExtensibilityPia))
            {
                return;
            }

            Type official = Assembly.LoadFrom(ExtensibilityPia)
                .GetTypes()
                .SingleOrDefault(t => t.IsInterface && t.GUID == ExpectedIid);

            if (official == null)
            {
                return;
            }

            Assert.Equal(
                official.GetMethods().Select(m => m.Name).ToArray(),
                typeof(IDTExtensibility2).GetMethods().Select(m => m.Name).ToArray());

            Assert.Equal(InterfaceTypeOf(official), InterfaceTypeOf(typeof(IDTExtensibility2)));

            for (int i = 0; i < ExpectedMembers.Length; i++)
            {
                Type[] officialParameters = official.GetMethods()[i].GetParameters().Select(p => p.ParameterType).ToArray();
                Type[] ourParameters = typeof(IDTExtensibility2).GetMethods()[i].GetParameters().Select(p => p.ParameterType).ToArray();

                Assert.Equal(officialParameters.Length, ourParameters.Length);

                for (int p = 0; p < officialParameters.Length; p++)
                {
                    // Enum parameters are declared in our own namespace, so compare the underlying
                    // shape rather than the identity: both marshal as a 32-bit integer.
                    Assert.Equal(Shape(officialParameters[p]), Shape(ourParameters[p]));
                }
            }
        }

        private static ComInterfaceType InterfaceTypeOf(Type type)
        {
            var attribute = (InterfaceTypeAttribute)type
                .GetCustomAttributes(typeof(InterfaceTypeAttribute), false)
                .SingleOrDefault();

            return attribute == null ? ComInterfaceType.InterfaceIsDual : attribute.Value;
        }

        private static string Shape(Type type)
        {
            bool byRef = type.IsByRef;
            Type core = byRef ? type.GetElementType() : type;
            string name = core.IsEnum ? "enum:" + Enum.GetUnderlyingType(core).Name : core.FullName;

            return (byRef ? "ref " : string.Empty) + name;
        }
    }
}
