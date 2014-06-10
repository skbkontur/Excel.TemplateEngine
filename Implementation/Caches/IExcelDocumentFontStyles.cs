﻿using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    internal interface IExcelDocumentFontStyles
    {
        uint AddFont(ExcelCellFontStyle style);
    }
}