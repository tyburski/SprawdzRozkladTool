using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SprawdzRozklad
{
    public class ScriptGenerator
    {
        public string Generate(string resource, List<string> insertStatements)
        {
            string resourceName = $"SprawdzRozklad.Skrypty.import_{resource}.sql";

            string sqlContent;
            var assembly = Assembly.GetExecutingAssembly();

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream, Encoding.GetEncoding(1250)))
            {
                sqlContent = reader.ReadToEnd();
            }

            int startIndex = sqlContent.IndexOf("--START");
            int endIndex = sqlContent.IndexOf("--KONIEC");

            if (startIndex == -1 || endIndex == -1 || endIndex <= startIndex)
            {
                throw new Exception("Nie znaleziono komentarzy --START i --KONIEC w pliku sql");
            }

            startIndex += "--START".Length;

            string newSqlContent = sqlContent.Substring(0, startIndex) + Environment.NewLine
                                 + string.Join(Environment.NewLine, insertStatements) + Environment.NewLine
                                 + sqlContent.Substring(endIndex);

            string tempFilePath = Path.Combine(Path.GetTempPath(), $"temp_{resource}.sql");

            File.WriteAllText(tempFilePath, newSqlContent, Encoding.GetEncoding(1250));

            return tempFilePath;
        }
    }
}
