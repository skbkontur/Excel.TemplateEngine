﻿using System;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public interface IFormValueParser
    {
        // todo (mpivko, 15.12.2017): why not generic?
        bool TryParse([NotNull] ITableParser tableParser, [NotNull] string name, [NotNull] Type modelType, out object result);
    }
}
