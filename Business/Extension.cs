using GroupDocs.Conversion;
using System.Text;

namespace Business
{
    public static class Extension
    {
        public static string ToJson(this string excelFilePath)
        {
            string jsonValue = "", tempFile = "";

            try
            {
                tempFile = Path.GetTempFileName();

                var converter = new Converter(excelFilePath);
                var convertOptions = converter.GetPossibleConversions()["json"].ConvertOptions;
                converter.Convert(tempFile, convertOptions);

                jsonValue = File.ReadAllText(tempFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (File.Exists(tempFile))
                    File.Delete(tempFile);
            }

            return jsonValue;
        }

        public static string ToJson(this Stream excelFile)
        {
            string jsonValue = "";
            try
            {
                MemoryStream stream = new MemoryStream();
                var converter = new Converter(() => excelFile);
                var convertOptions = converter.GetPossibleConversions()["json"].ConvertOptions;
                converter.Convert(() => stream, convertOptions);

                var jsonValueByteArray = stream.ToArray();
                jsonValue = Encoding.UTF8.GetString(jsonValueByteArray, 0, jsonValueByteArray.Length);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return jsonValue;
        }
    }
}
