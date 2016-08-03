using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Data;
using System.IO;

namespace CCure_MAS_Data_Reader
{
    public static class Program
    {
        static void Main(string[] args)
        {
            // Gives you A date time to add to csv name so you always generate a new unique file
            DateTime dateTimeVariable = DateTime.Now;

            //Get Desktop of Current User to dump CSV with Timestamp there.
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string output =  path + "\\ActivePersonnel-" + dateTimeVariable.Year + "-" +dateTimeVariable.Month + "-" +dateTimeVariable.Day + ".csv";

            // Setup the Connection to the MAS to pull the data we need
            SqlConnection MASDatabase = new SqlConnection("REDACTED");
            SqlCommand MasQuery = new SqlCommand(); 
            SqlDataReader reader = null; 

            // Setup the Query we want to use to pull data, the type of data we are sending (text), and link the query to the MAS so we can authenticate and run it
            MasQuery.CommandText = "REDACTED";
            MasQuery.CommandType = CommandType.Text;
            MasQuery.Connection = MASDatabase;

            // Open the MAS Database and read all the data from the query
            MASDatabase.Open();
            reader = MasQuery.ExecuteReader();

            ToCsv(reader, output, true);
            MASDatabase.Close();

        }

        // This was fun to write, and frustrating. TODO: Upgrade from CSV to XLSX, or support both.
        public static void ToCsv(this IDataReader dataReader, string fileName, bool includeHeaderAsFirstRow)
        {

            const string Separator = ",";

            StreamWriter streamWriter = new StreamWriter(fileName);

            StringBuilder sb = null;

            if (includeHeaderAsFirstRow)
            {
                sb = new StringBuilder();
                for (int index = 0; index < dataReader.FieldCount; index++)
                {
                    if (dataReader.GetName(index) != null)
                        sb.Append(dataReader.GetName(index));

                    if (index < dataReader.FieldCount - 1)
                        sb.Append(Separator);
                }
                streamWriter.WriteLine(sb.ToString());
            }

            while (dataReader.Read())
            {
                sb = new StringBuilder();
                for (int index = 0; index < dataReader.FieldCount - 1; index++)
                {
                    if (!dataReader.IsDBNull(index))
                    {
                        string value = dataReader.GetValue(index).ToString();
                        if (dataReader.GetFieldType(index) == typeof(String))
                        {
                            if (value.IndexOf("\"") >= 0)
                                value = value.Replace("\"", "\"\"");

                            if (value.IndexOf(Separator) >= 0)
                                value = "\"" + value + "\"";
                        }
                        sb.Append(value);
                    }

                    if (index < dataReader.FieldCount - 1)
                        sb.Append(Separator);
                }

                if (!dataReader.IsDBNull(dataReader.FieldCount - 1))
                    sb.Append(dataReader.GetValue(dataReader.FieldCount - 1).ToString().Replace(Separator, " "));

                streamWriter.WriteLine(sb.ToString());
            }
            dataReader.Close();
            streamWriter.Close();
        }
    }
}

