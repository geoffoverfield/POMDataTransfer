// Geoff Overfield
// 05/14/2019
// Scripts to transfer and merge client data from a collection
// of query-based CSV files to a single file for upload to new client
// Completed at request of Management of Peace of Mind Massage

namespace POMDataTransfer
{
    public class POM_Files
    {
        public static string UsersFile = @"D:\$\POMDataTransfer\POMDataTransfer\POMDataTransfer\Resources\PeaceOfMind_Clients_SoapVault.xlsx";
        public static string MailingList = @"D:\$\POMDataTransfer\POMDataTransfer\POMDataTransfer\Resources\Mailing_List.xlsx";
        public static string PhoneNumbers = @"D:\$\POMDataTransfer\POMDataTransfer\POMDataTransfer\Resources\Member_PhoneNumbers.xlsx";

        public static string UsersSaveDirectoryCSV = @"D:\$\POMDataTransfer\POMDataTransfer\POMDataTransfer\Resources\PeaceOfMind_Clients_SoapVault.csv";
    }
}
