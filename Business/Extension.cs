using GroupDocs.Conversion;

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
    }
}
