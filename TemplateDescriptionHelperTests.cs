using System;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.Core.Tests.ExcelObjectPrinterTests
{
    [TestFixture]
    public class TemplateDescriptionHelperTests
    {
        [Test]
        public void TemplateDescriptionCorrectnessTest()
        {
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:TemplateName:A1:B2"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:TemplateName:ABCD122:BER22"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:TemplateName:ABCDEFGHIJKLMNOPQRSTUVWXYZ1:B1234567890"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("TemplateTemplateName:A1:B2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription(":TemplateName::"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("::"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template::A2:BCD33"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name::"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name:B:33"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name:33:33"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name:QWE:EWQ"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template::A1:A2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name:A0:A2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name:A02:A2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name:A02:a2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription(""));
        }

        [Test]
        public void ValueDescriptionCorrectnessTest()
        {
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::Property"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::P"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::A.B"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::aa.bb"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::Property[]"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::A[].B[].Cc.ddd[]"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::B52"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::9das"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::[]"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::Asdf["));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::asdf]"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::A[]B[]"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::sdf.[]"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::128[]"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::asd,vcd"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value:Template:"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Val::A"));
        }

        [Test]
        public void CellsMergingCommandCorrectnessTest()
        {
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCells:A1:B2"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCells:ABCD122:BER22"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCells:ABCDEFGHIJKLMNOPQRSTUVWXYZ1:B1234567890"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCellsMergeCells:A1:B2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCells::"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("::"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCells::A2:BCD33"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCells::"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCells:B:33"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCells:33:33"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCells:QWE:EWQ"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCells::A1:A2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCells:A0:A2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCells:A02:A2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand("MergeCells:A02:a2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectCellsMergingCommand(""));
        }

        [Test]
        public void TemplateCoordinatesTest()
        {
            Tuple<ICellPosition, ICellPosition> range;
            Assert.IsTrue(TemplateDescriptionHelper.Instance.TryExtractCoordinates("Template:qwe:QWE123:ASD987", out range));
            Assert.AreEqual(17 * 26 * 26 + 23 * 26 + 5, range.Item1.ColumnIndex);
            Assert.AreEqual(123, range.Item1.RowIndex);
            Assert.AreEqual(26 * 26 + 19 * 26 + 4, range.Item2.ColumnIndex);
            Assert.AreEqual(987, range.Item2.RowIndex);

            Assert.IsTrue(TemplateDescriptionHelper.Instance.TryExtractCoordinates("MergeCells:QWE123:ASD987", out range));
            Assert.AreEqual(17 * 26 * 26 + 23 * 26 + 5, range.Item1.ColumnIndex);
            Assert.AreEqual(123, range.Item1.RowIndex);
            Assert.AreEqual(26 * 26 + 19 * 26 + 4, range.Item2.ColumnIndex);
            Assert.AreEqual(987, range.Item2.RowIndex);
        }
    }
}