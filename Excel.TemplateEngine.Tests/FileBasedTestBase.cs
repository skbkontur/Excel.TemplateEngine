using System;
using System.IO;

using NUnit.Framework;

namespace SkbKontur.Excel.TemplateEngine.Tests
{
    public class FileBasedTestBase
    {
        protected string GetFilePath(string path)
        {
            var baseNamespace = typeof(FileBasedTestBase).Namespace;
            var currentNamespace = GetType().Namespace;
            if (string.IsNullOrEmpty(baseNamespace))
                throw new Exception("Base namespace should not be null");
            if (string.IsNullOrEmpty(currentNamespace))
                throw new Exception("Current namespace should not be null");
            if (string.IsNullOrEmpty(path))
                throw new Exception("Path should not be null");

            if (!currentNamespace.StartsWith(baseNamespace))
                throw new Exception($"Unable to get path for namespace: '{currentNamespace}' with base namespace: '{baseNamespace}'");

            var namespaceFolders = currentNamespace.Substring(baseNamespace.Length).Split('.');
            return Path.Combine(TestContext.CurrentContext.TestDirectory, Path.Combine(namespaceFolders), "Files", path);
        }
    }
}