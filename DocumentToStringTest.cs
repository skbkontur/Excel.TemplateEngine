using System;
using System.IO;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelFileGenerator;

namespace SKBKontur.Catalogue.Core.Tests.ExcelFileGeneratorTests
{
    [TestFixture]
    public class DocumentToStringTest
    {
        [Test]
        public void Test()
        {
            var templateDoucument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(templateFileName));
            var s = templateDoucument.ToString();
            Console.Write(s);

            Assert.AreEqual(@"A1:123.45
	NumberingFormat = {FormatCode = 0.0000}
	FillStyle = {FillColor = {R = 128, G = 129, B = 200, A = 255}}
	FontStyle = {
			Size = 20
			Color = R = 122, G = 200, B = 20, A = 255
			Underlined
			Bold
		}
	BordersStyle = {
			LeftBorder = {
				Color = R = 255, G = 0, B = 255, A = 255
				BorderType = Double
			}
			RightBorder = {
				Color = R = 255, G = 0, B = 255, A = 255
				BorderType = Double
			}
			LeftBorder = {
				Color = R = 255, G = 0, B = 255, A = 255
				BorderType = Double
			}
			RightBorder = {
				Color = R = 255, G = 0, B = 255, A = 255
				BorderType = Double
			}
		}
	Alignment = {
			HorizontalAlignment = Left
			VerticalAlignment = Default
			WrapText
		}
B9:Model:Document:B10:C11

B10:Покупатель:
	FontStyle = {
			Size = 11
		}
C10:Поставщик:
	FontStyle = {
			Size = 11
		}
B11:Value:Organization:Buyer
	FontStyle = {
			Size = 11
		}
C11:Value:Organization:Supplier
	NumberingFormat = {FormatCode = #,##0.00""р.""}
	FontStyle = {
			Size = 11
		}
D11:Value:String:Name
	FontStyle = {
			Size = 11
		}
A22:Template:Organization:A23:A24

A23:Value:String:Name

A24:Value:String:Address

", s);
        }

        private const string templateFileName = @"ExcelFileGeneratorTests\Files\template.xlsx";
    }
}