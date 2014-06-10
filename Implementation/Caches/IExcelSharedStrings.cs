﻿using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    internal interface IExcelSharedStrings
    {
        uint AddSharedString(FormattedStringValue value);
        void Save();
    }
}