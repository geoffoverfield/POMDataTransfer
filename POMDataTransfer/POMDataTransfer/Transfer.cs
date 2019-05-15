// Geoff Overfield
// 05/14/2019
// Scripts to transfer and merge client data from a collection
// of query-based CSV files to a single file for upload to new client
// Completed at request of Management of Peace of Mind Massage

#region Namespaces
using System;
using System.Collections.Generic;
#endregion

namespace POMDataTransfer
{
    class Transfer
    {
        static void Main(string[] args)
        {
            Dictionary<string, Dictionary<string, string>> dUserData = new Dictionary<string, Dictionary<string, string>>();
            ExcelReader pUsersReader = new ExcelReader(POM_Files.UsersFile, 1);
            //ExcelReader pEmailReader = new ExcelReader(POM_Files.MailingList, 1);

            int i = 0;
            for (i = 0; i < 227; i++)
            {
                string sFirstName = pUsersReader.ReadCell(i, 0);
                string sLastName = pUsersReader.ReadCell(i, 1);
                if (!dUserData.ContainsKey(sLastName))
                    dUserData.Add(sLastName, new Dictionary<string, string>() { { sFirstName, string.Empty } });
                else
                {
                    if (!dUserData[sLastName].ContainsKey(sFirstName))
                        dUserData[sLastName].Add(sFirstName, string.Empty);
                }
            }

            //transferEmailAddresses(pEmailReader, pUsersReader, dUserEmails);

            pUsersReader.Save();
            pUsersReader.SaveAs(POM_Files.UsersSaveDirectoryCSV);
            //pEmailReader.Dispose();
            pUsersReader.Dispose();

            Console.ReadLine();
        }

        private static void transferEmailAddresses(ExcelReader pEmailReader, ExcelReader pUsersReader, Dictionary<string, Dictionary<string, string>> dUserEmails)
        {
            /// Get and input users emails
            for (int i = 0; i < 18868; i++)
            {
                var sEmail = pEmailReader.ReadCell(i, 4);
                if (string.IsNullOrWhiteSpace(sEmail)) continue;

                var sLastName = pEmailReader.ReadCell(i, 0);
                if (dUserEmails.ContainsKey(sLastName))
                {
                    var sFirstName = pEmailReader.ReadCell(i, 1);
                    if (dUserEmails.ContainsKey(sLastName))
                    {
                        if (dUserEmails[sLastName].ContainsKey(sFirstName))
                            dUserEmails[sLastName][sFirstName] = sEmail;
                    }
                }
            }

            for (int i = 0; i < dUserEmails.Count; i++)
            {
                string sFirst = pUsersReader.ReadCell(i, 0);
                string sLast = pUsersReader.ReadCell(i, 1);
                pUsersReader.WriteToCell(i, 2, dUserEmails[sLast][sFirst]);
            }
        }
    }
}
