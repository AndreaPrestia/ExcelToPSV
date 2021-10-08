using ExcelDataReader;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelToPSV
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Count() < 2)
            {
                Console.WriteLine("You must provide a source and destination path! Press a key to close...");
                Console.ReadLine();
                return;
            }

            try
            {
                Console.WriteLine($"Converting {args[0]} to {args[1]}");
                ConvertToPSV(args[0], args[1]);
                Console.WriteLine("Done!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                Console.WriteLine("Press a key to close...");
                Console.ReadLine();
            }
        }

        private static void ConvertToPSV(string source, string target)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            if (target is null)
            {
                throw new ArgumentNullException(nameof(target));
            }

            if (!File.Exists(source))
            {
                throw new InvalidOperationException($"Path {source} does not exists.");
            }

            FileInfo fileInfo = new FileInfo(source);

            StringBuilder csvContent = new StringBuilder();

            using (Stream stream = new FileStream(fileInfo.FullName, FileMode.Open, FileAccess.Read))
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                IExcelDataReader reader = fileInfo.Extension.Equals("xls") ? ExcelReaderFactory.CreateBinaryReader(stream) : ExcelReaderFactory.CreateOpenXmlReader(stream);

                while (reader.Read())
                {
                    for (int c = 0; c < reader.FieldCount; c++)
                    {
                        csvContent.Append(reader[c]).Append('|');
                    }

                    csvContent.AppendLine();
                }

                reader.Close();
            }

            File.WriteAllText(target, csvContent.ToString());
        }
    }
}
