using Macro_Polo.Core;
using Xunit;

namespace Macro_Polo.Core.Tests
{
    public class OfficeVersionTests
    {
        [Theory]
        [InlineData("16.0", "16.0")]
        [InlineData("16.0.17928.20114", "16.0")]
        [InlineData("15.0", "15.0")]
        [InlineData("17", "17.0")]
        public void A_reported_version_is_reduced_to_its_major_form(string reported, string expected)
        {
            Assert.Equal(expected, OfficeVersion.Normalize(reported));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not a version")]
        [InlineData("0.0")]
        public void An_unusable_version_falls_back_rather_than_producing_a_broken_registry_path(string reported)
        {
            Assert.Equal(OfficeVersion.Fallback, OfficeVersion.Normalize(reported));
        }
    }
}
