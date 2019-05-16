// Geoff Overfield
// 05/14/2019
// Scripts to transfer and merge client data from a collection
// of query-based CSV files to a single file for upload to new client
// Completed at request of Management of Peace of Mind Massage

#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
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
            //ExcelReader pPhoneReader = new ExcelReader(POM_Files.PhoneNumbers, 1);

            int i = 0;
            for (i = 1; i < 228; i++)
            {
                string sFirstName = pUsersReader.ReadCell(i, 0);
                string sLastName = pUsersReader.ReadCell(i, 1);
                if (string.IsNullOrWhiteSpace(sLastName)) continue;

                if (!dUserData.ContainsKey(sLastName))
                    dUserData.Add(sLastName, new Dictionary<string, string>() { { sFirstName, string.Empty } });
                else
                {
                    if (!dUserData[sLastName].ContainsKey(sFirstName))
                        dUserData[sLastName].Add(sFirstName, string.Empty);
                }
            }

            //transferEmailAddresses(pEmailReader, pUsersReader, dUserData, true);
            //transferPhoneNumbers(pPhoneReader, pUsersReader, dUserData);

            pUsersReader.Save();
            pUsersReader.SaveAs(POM_Files.UsersSaveDirectoryCSV);
            //pEmailReader.Dispose();
            //pPhoneReader.Dispose();
            pUsersReader.Dispose();

            Console.ReadLine();
        }

        private static void transferEmailAddresses(ExcelReader pEmailReader, ExcelReader pUsersReader, Dictionary<string, Dictionary<string, string>> dUserData, bool bTakePhoneNumber)
        {
            /// Get and input users emails
            for (int i = 0; i < 18868; i++)
            {
                var sEmail = pEmailReader.ReadCell(i, 4);
                string sHomeNumber = pEmailReader.ReadCell(i, 6);
                if (string.IsNullOrWhiteSpace(sEmail) ||
                    string.IsNullOrWhiteSpace(sHomeNumber)) continue;

                var sLastName = pEmailReader.ReadCell(i, 0);
                if (dUserData.ContainsKey(sLastName))
                {
                    var sFirstName = pEmailReader.ReadCell(i, 1);
                    if (dUserData.ContainsKey(sLastName))
                    {
                        if (dUserData[sLastName].ContainsKey(sFirstName))
                            dUserData[sLastName][sFirstName] = bTakePhoneNumber ? sHomeNumber : sEmail;
                    }
                }
            }

            for (int i = 0; i < dUserData.Count; i++)
            {
                string sFirst = pUsersReader.ReadCell(i, 0);
                string sLast = pUsersReader.ReadCell(i, 1);
                if (!dUserData.ContainsKey(sLast) ||
                    !dUserData[sLast].ContainsKey(sFirst)) continue;

                if (bTakePhoneNumber)
                    pUsersReader.WriteToCell(i, 5, dUserData[sLast][sFirst]);
                else
                pUsersReader.WriteToCell(i, 2, dUserData[sLast][sFirst]);
            }
        }

        private static void transferPhoneNumbers(ExcelReader pPhoneReader, ExcelReader pUsersReader, Dictionary<string, Dictionary<string, string>> dUserData)
        {
            /// Get and input users phone numbers
            for (int i = 0; i < 415; i++)
            {
                var sPhoneNumber = pPhoneReader.ReadCell(i, 2);
                if (string.IsNullOrWhiteSpace(sPhoneNumber)) continue;

                var sCompoundName = pPhoneReader.ReadCell(i, 1);
                var sNames = sCompoundName.Split(',');
                var sLastName = sNames[0].Trim();
                if (dUserData.ContainsKey(sLastName))
                {
                    var sFirstName = sNames[1].Trim();
                    if (dUserData.ContainsKey(sLastName))
                    {
                        if (dUserData[sLastName].ContainsKey(sFirstName))
                            dUserData[sLastName][sFirstName] = sPhoneNumber;
                    }
                }
            }

            for (int i = 0; i < dUserData.Count; i++)
            {
                string sFirst = pUsersReader.ReadCell(i, 0);
                string sLast = pUsersReader.ReadCell(i, 1);
                if (!dUserData.ContainsKey(sLast) ||
                    !dUserData[sLast].ContainsKey(sFirst)) continue;

                pUsersReader.WriteToCell(i, 7, dUserData[sLast][sFirst]);
            }
        }
    }
}
