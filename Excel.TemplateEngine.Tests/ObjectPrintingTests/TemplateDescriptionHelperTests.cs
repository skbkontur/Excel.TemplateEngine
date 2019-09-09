using NUnit.Framework;

namespace Excel.TemplateEngine.Tests.ExcelObjectPrinterTests
{
    [TestFixture]
    public class TemplateDescriptionHelperTests
    {
        [Test]
        public void TemplateDescriptionCorrectnessTest()
        {
            Assert.IsTrue(TemplateDescriptionHelper.IsCorrectTemplateDescription("Template:TemplateName:A1:B2"));
            Assert.IsTrue(TemplateDescriptionHelper.IsCorrectTemplateDescription("Template:TemplateName:ABCD122:BER22"));
            Assert.IsTrue(TemplateDescriptionHelper.IsCorrectTemplateDescription("Template:TemplateName:ABCDEFGHIJKLMNOPQRSTUVWXYZ1:B1234567890"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectTemplateDescription("TemplateTemplateName:A1:B2"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectTemplateDescription(":TemplateName::"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectTemplateDescription("::"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectTemplateDescription("Template::A2:BCD33"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectTemplateDescription("Template:Name::"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectTemplateDescription("Template:Name:B:33"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectTemplateDescription("Template:Name:33:33"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectTemplateDescription("Template:Name:QWE:EWQ"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectTemplateDescription("Template::A1:A2"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectTemplateDescription("Template:Name:A0:A2"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectTemplateDescription("Template:Name:A02:A2"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectTemplateDescription("Template:Name:A02:a2"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectTemplateDescription(""));
        }

        [Test]
        public void ValueDescriptionCorrectnessTest()
        {
            Assert.IsTrue(TemplateDescriptionHelper.IsCorrectValueDescription("Value::Property"));
            Assert.IsTrue(TemplateDescriptionHelper.IsCorrectValueDescription("Value::P"));
            Assert.IsTrue(TemplateDescriptionHelper.IsCorrectValueDescription("Value::A.B"));
            Assert.IsTrue(TemplateDescriptionHelper.IsCorrectValueDescription("Value::aa.bb"));
            Assert.IsTrue(TemplateDescriptionHelper.IsCorrectValueDescription("Value::Property[]"));
            Assert.IsTrue(TemplateDescriptionHelper.IsCorrectValueDescription("Value::A[].B[].Cc.ddd[]"));
            Assert.IsTrue(TemplateDescriptionHelper.IsCorrectValueDescription("Value::B52"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectValueDescription("Value::"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectValueDescription("Value::9das"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectValueDescription("Value::[]"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectValueDescription("Value::Asdf["));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectValueDescription("Value::asdf]"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectValueDescription("Value::A[]B[]"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectValueDescription("Value::sdf.[]"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectValueDescription("Value::128[]"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectValueDescription("Value::asd,vcd"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectValueDescription("Value:Template:"));
            Assert.IsFalse(TemplateDescriptionHelper.IsCorrectValueDescription("Val::A"));
        }

        [TestCase("CheckBox:CheckBoxName:Path")]
        [TestCase("DropDown:DropDownName:Path")]
        [TestCase("CheckBox:abc:A.SubItem.Array[10].Element")]
        [TestCase("DropDown:cde:Test.Dict[\"aaa\"]")]
        [TestCase("CheckBox:very[strange] name\"with quote:Path")]
        public void TestIsCorrectFormValueDescriptionReturnsTrue(string description)
        {
            Assert.True(TemplateDescriptionHelper.IsCorrectFormValueDescription(description));
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
            Assert.False(TemplateDescriptionHelper.IsCorrectFormValueDescription(description));
        }

        [TestCase("Test[123]")]
        [TestCase("Test[\"lalala\"]")]
        [TestCase("Test[abc]")]
        [TestCase("Test['abc']")]
        [TestCase("Test[1.5]")]
        public void TestIsCollectionAccessPathPartReturnsTrue(string pathPart)
        {
            Assert.True(TemplateDescriptionHelper.IsCollectionAccessPathPart(pathPart));
        }

        [TestCase("Test")]
        [TestCase("Test[]")]
        [TestCase("[]")]
        [TestCase("[1]")]
        [TestCase("[\"abc\"]")]
        [TestCase("Test[123]abc")]
        public void TestIsCollectionAccessPathPartReturnsFalse(string pathPart)
        {
            Assert.False(TemplateDescriptionHelper.IsCollectionAccessPathPart(pathPart));
        }

        [TestCase("Test[]")]
        public void TestIsIsArrayPathPartReturnsTrue(string pathPart)
        {
            Assert.True(TemplateDescriptionHelper.IsArrayPathPart(pathPart));
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
            Assert.False(TemplateDescriptionHelper.IsArrayPathPart(pathPart));
        }

        [Test]
        public void TemplateCoordinatesTest()
        {
            IRectangle range;
            Assert.IsTrue(TemplateDescriptionHelper.TryExtractCoordinates("Template:qwe:QWE123:ASD987", out range));
            Assert.AreEqual(26 * 26 + 19 * 26 + 4, range.UpperLeft.ColumnIndex);
            Assert.AreEqual(123, range.UpperLeft.RowIndex);
            Assert.AreEqual(17 * 26 * 26 + 23 * 26 + 5, range.LowerRight.ColumnIndex);
            Assert.AreEqual(987, range.LowerRight.RowIndex);
        }
    }
}