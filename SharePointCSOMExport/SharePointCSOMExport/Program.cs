using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Security;
using System.IO;

namespace SharePointCSOMExport
{
    class Program
    {
        static void Main(string[] args)
        {

            #region Site Details - Read the details from app.config
            string URL = ConfigurationManager.AppSettings["URL"];
            string username = ConfigurationManager.AppSettings["username"];
            string password = ConfigurationManager.AppSettings["password"];
            string listName = ConfigurationManager.AppSettings["listName"];
            string viewName = ConfigurationManager.AppSettings["viewName"];
            string path = ConfigurationManager.AppSettings["path"];
            #endregion

            using (ClientContext context = new ClientContext(URL))
            {
                SecureString securePassword = GetSecureString(password);
                context.Credentials = new SharePointOnlineCredentials(username, securePassword);

                List list = context.Web.Lists.GetByTitle(listName);
                context.Load(list);
                context.ExecuteQuery();

                View view = list.Views.GetByTitle(viewName);
                context.Load(view);
                context.ExecuteQuery();


                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                query.ViewXml = "<View><Query>" + view.ViewQuery + "</Query></View>";
                ListItemCollection items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQuery();

                //Console.WriteLine("Total Count: " + items.Count);

                /* foreach (ListItem item in items)
                   {
                       Console.WriteLine("Title " + item["Title"]);
                   } 

                    */
                #region -- Datatable
                //new datatable
                DataTable data = new DataTable();

                //add column names (internal names from URL)
                data.Columns.Add("Title", typeof(string));
                data.Columns.Add("Actual_Units_Scanned", typeof(Int32));
                data.Columns.Add("Attainment", typeof(double));
                data.Columns.Add("Created_By", typeof(string));

                //add each row to datarow
                foreach (ListItem item in items)
                {
                    data.Rows.Add(item["LineLead"], item["Actual_x0020_Units_x0020_Scanned"], item["Attainment"], item["Created"]);
                }

                //Just to display the data table
                foreach (DataRow row in data.Rows)
                {
                    Console.WriteLine();
                    for (int x = 0; x < data.Columns.Count; x++)
                    {
                        Console.Write(row[x].ToString() + " ");
                    }
                }
                #endregion


                #region --generate csv

                PutDataTableToCsv(path,data,true);

                #endregion   


                Console.ReadKey();

            }

        }

        static void PutDataTableToCsv(string path, DataTable table, bool isFirstRowHeader)
        {
            var lines = new List<string>(); // create a list of strings to hold the file rows

            // if there are headers add them to the file first
            if (isFirstRowHeader)
            {
                string[] colnames = table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
                var header = string.Join(",", colnames);
                lines.Add(header);
            }

            // Place commas between the elements of each row
            var valueLines = table.AsEnumerable().Cast<DataRow>().Select(row => string.Join(",", row.ItemArray.Select(o => o.ToString()).ToArray()));

            // Stuff the rows into a string joined by new line characters
            var allLines = string.Join(Environment.NewLine, valueLines.ToArray<string>());
            lines.Add(allLines);

            // put that file to bed
            System.IO.File.WriteAllLines(path, lines.ToArray());
        }

        public static SecureString GetSecureString(string userPassword)
        {
            SecureString securePassword = new SecureString();
            foreach (char c in userPassword.ToCharArray())
            {
                securePassword.AppendChar(c);
            }
            return securePassword;
        }
    }
}
