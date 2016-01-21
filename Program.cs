/*******************************************
 * ExcelParser tool for KNIME Analytics    *
 * Written by Corey Noble (noblecorey.com) *
 * Winter 2016                             *
 *******************************************/

// Libraries
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.VisualBasic;

// Declare namespace
namespace ExcelParser
{
    // Declare class
    class Program
    {
        // Global variables
        static DataTable dataTable = new DataTable();

        // Main method
        // Create console, ask for input, execute necessary functions
        static void Main(string[] args)
        {
            // Prompt for and store input
            Console.Write("Please enter the file that you want to read (include the extension): ");
            string inputFile = Console.ReadLine();

            // Parse and clean data
            DataParser(inputFile);
            DataCleaner();

            // Prompt for and write to output, clearing file beforehand
            Console.Write("Please enter the file that you want the data written to: ");
            string outputFile = Console.ReadLine();
            File.WriteAllText(@outputFile, string.Empty);
            CreateCSVFile(outputFile);
        }

        // DataParser method
        static void DataParser(string file)
        {
            // Specify connection to file
            string connection = string.Format("provider=Microsoft.Jet.OLEDB.4.0; data source={0};Extended Properties=Excel 8.0;", file);

            // Use specified connection
            using (OleDbConnection con = new OleDbConnection(connection))
            {
                // Specify sheet name
                var sheet = "Sheet1$";

                // Create query for database
                string query = string.Format("SELECT * FROM [{0}]", sheet);

                // Open connection
                con.Open();

                // Create adapter and fill dataTable
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                adapter.Fill(dataTable);

                // Close connection
                con.Close();
            }
        }

        // CreateCSVFile method
        static void CreateCSVFile(string file)
        {
            // Create stream
            StreamWriter sw = new StreamWriter(file, false);

            // Iterate through each row in dataTable
            foreach (DataRow dr in dataTable.Rows)
            {
                // Check for illegal characters
                if (dr[0].ToString() != "000")
                {
                    if (dr[0].ToString() != " ")
                    {
                        if (dr[0].ToString() != "")
                        {
                            if (!Convert.IsDBNull(dr[0]))
                            {
                                // Write to stream
                                sw.Write(dr[0].ToString());
                            }

                            // Add enter after word
                            sw.Write(sw.NewLine);
                        }
                    }
                }
            }

            // Close stream
            sw.Close();
        }

        // DataCleaner method
        static void DataCleaner()
        {
            // Iterate through each row in dataTable
            foreach (DataRow dr in dataTable.Rows)
            {
                // Create string version of the word in the current row
                string rowString = dr[0].ToString();

                // Create StringBuilder
                StringBuilder sb = new StringBuilder(rowString);

                // Iterate through each character of the word
                for (int i = 0; i < rowString.Length; i++ )
                {
                    // Variable for cleaning purposes
                    int j = 0;

                    // Replace commas with newlines if first comma encountered
                    if (j == 0)
                    {
                        if (rowString[i].ToString() == ",")
                        {
                            sb[i] = Convert.ToChar("\n");
                            j++;
                        }
                    }

                    // Replace commas with nothing if not first comma encountered
                    else
                    {
                        if (rowString[i].ToString() == ",")
                        {
                            sb[i] = Convert.ToChar("");
                        }
                    }
                }

                // Rebuild string
                rowString = sb.ToString();

                // Remove quotation marks
                rowString = rowString.Replace("\"", "");

                // Remove all hyperlinks
                if (rowString.Length > 3)
                {
                    for (int i = 0; i < rowString.Length; i++)
                    {
                        if (rowString[0].ToString() == "h")
                        {
                            if (rowString[1].ToString() == "t")
                            {
                                if (rowString[2].ToString() == "t")
                                {
                                    if (rowString[3].ToString() == "p")
                                    {
                                        rowString = "";
                                    }
                                }
                            }
                        }
                    }
                }
                
                // Replace all illegal characters with empty strings
                Regex rgx = new Regex("[^a-zA-Z0-9 -'@]");
                rowString = rgx.Replace(rowString, "");

                // Set rowString to all lowercase
                rowString = rowString.ToLower();

                // Assign to row
                dr[0] = rowString;
            }
        }
    }
}
