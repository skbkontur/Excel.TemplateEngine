using FluentAssertions;

using NUnit.Framework;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;

namespace SkbKontur.Excel.TemplateEngine.Tests.ObjectPrintingTests
{
    [TestFixture]
    public class TemplateDescriptionHelperTests
    {
        [TestCase("Template:TemplateName:A1:B2", true)]
        [TestCase("Template:TemplateName:ABCD122:BER22", true)]
        [TestCase("Template:TemplateName:ABCDEFGHIJKLMNOPQRSTUVWXYZ1:B1234567890", true)]
        [TestCase("TemplateTemplateName:A1:B2", false)]
        [TestCase(":TemplateName::", false)]
        [TestCase("::", false)]
        [TestCase("Template::A2:BCD33", false)]
        [TestCase("Template:Name::", false)]
        [TestCase("Template:Name:B:33", false)]
        [TestCase("Template:Name:33:33", false)]
        [TestCase("Template:Name:QWE:EWQ", false)]
        [TestCase("Template::A1:A2", false)]
        [TestCase("Template:Name:A0:A2", false)]
        [TestCase("Template:Name:A02:A2", false)]
        [TestCase("Template:Name:A02:a2", false)]
        [TestCase("", false)]
        public void TemplateDescriptionCorrectnessTest(string expression, bool isCorrect)
        {
            TemplateDescriptionHelper.IsCorrectTemplateDescription(expression).Should().Be(isCorrect);
        }

        [TestCase("Value::Property", true)]
        [TestCase("Value::P", true)]
        [TestCase("Value::A.B", true)]
        [TestCase("Value::aa.bb", true)]
        [TestCase("Value::Property[]", true)]
        [TestCase("Value::A[].B[].Cc.ddd[]", true)]
        [TestCase("Value::B52", true)]
        [TestCase("Value::", false)]
        [TestCase("Value::9das", false)]
        [TestCase("Value::[]", false)]
        [TestCase("Value::Asdf[", false)]
        [TestCase("Value::asdf]", false)]
        [TestCase("Value::A[]B[]", false)]
        [TestCase("Value::sdf.[]", false)]
        [TestCase("Value::128[]", false)]
        [TestCase("Value::asd,vcd", false)]
        [TestCase("Value:Template:", false)]
        [TestCase("Val::A", false)]
        public void ValueDescriptionCorrectnessTest(string expression, bool isCorrect)
        {
            TemplateDescriptionHelper.IsCorrectValueDescription(expression).Should().Be(isCorrect);
        }

        [TestCase("CheckBox:CheckBoxName:Path")]
        [TestCase("DropDown:DropDownName:Path")]
        [TestCase("CheckBox:abc:A.SubItem.Array[10].Element")]
        [TestCase("DropDown:cde:Test.Dict[\"aaa\"]")]
        [TestCase("CheckBox:very[strange] name\"with quote:Path")]
        public void TestIsCorrectFormValueDescriptionReturnsTrue(string description)
        {
            TemplateDescriptionHelper.IsCorrectFormValueDescription(description).Should().BeTrue();
        }

        [TestCase("Lalala:CheckBoxName:Path")]
        [TestCase("CheckBox::Path")]
        [TestCase("CheckBox::")]
        [TestCase(":Name:Path")]
        [TestCase("::Path")]
        [TestCase("::")]
        [TestCase("abracadabra")]
        [TestCase("CheckBox:b")]
        [TestCase("CheckBox:abc:Path:Test")]
        public void TestIsCorrectFormValueDescriptionReturnsFalse(string description)
        {
            TemplateDescriptionHelper.IsCorrectFormValueDescription(description).Should().BeFalse();
        }

        [TestCase("Test[123]")]
        [TestCase("Test[\"lalala\"]")]
        [TestCase("Test[abc]")]
        [TestCase("Test['abc']")]
        [TestCase("Test[1.5]")]
        public void TestIsCollectionAccessPathPartReturnsTrue(string pathPart)
        {
            TemplateDescriptionHelper.IsCollectionAccessPathPart(pathPart).Should().BeTrue();
        }

        [TestCase("Test")]
        [TestCase("Test[]")]
        [TestCase("[]")]
        [TestCase("[1]")]
        [TestCase("[\"abc\"]")]
        [TestCase("Test[123]abc")]
        public void TestIsCollectionAccessPathPartReturnsFalse(string pathPart)
        {
            TemplateDescriptionHelper.IsCollectionAccessPathPart(pathPart).Should().BeFalse();
        }

        [TestCase("Test[]")]
        public void TestIsIsArrayPathPartReturnsTrue(string pathPart)
        {
            TemplateDescriptionHelper.IsArrayPathPart(pathPart).Should().BeTrue();
        }

        [TestCase("Test")]
        [TestCase("Test[123]")]
        [TestCase("[]")]
        [TestCase("[1]")]
        [TestCase("[\"abc\"]")]
        [TestCase("Test[123]lalala")]
        [TestCase("Test[][]")]
        public void TestIsArrayPathPartReturnsFalse(string pathPart)
        {
            TemplateDescriptionHelper.IsArrayPathPart(pathPart).Should().BeFalse();
        }

        [Test]
        public void TemplateCoordinatesTest()
        {
            TemplateDescriptionHelper.TryExtractCoordinates("Template:qwe:QWE123:ASD987", out var range).Should().BeTrue();
            range.UpperLeft.ColumnIndex.Should().Be(26 * 26 + 19 * 26 + 4);
            range.UpperLeft.RowIndex.Should().Be(123);
            range.LowerRight.ColumnIndex.Should().Be(17 * 26 * 26 + 23 * 26 + 5);
            range.LowerRight.RowIndex.Should().Be(987);
        }
    }
}