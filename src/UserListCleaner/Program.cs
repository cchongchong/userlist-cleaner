using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using CommandLine;
using CommandLine.Text;
using NLog;
using NPOI.XSSF.UserModel;

namespace UserListCleaner
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var options = new Options();
            if (Parser.Default.ParseArguments(args, options))
            {
                log.Info(options.InputFile);
                log.Info(options.OutputFile);
                if (!string.IsNullOrEmpty(options.CompareToInputFile))
                {
                    log.Info("Compare-to list: {0}", options.CompareToInputFile);
                    log.Info("Auto correct username based on compare-to list: {0}", options.AutoCorrect ? "Yes" : "No");
                }
                if (!string.IsNullOrEmpty(options.ReplaceToInputFile))
                {
                    log.Info("Replace-to list: {0}", options.ReplaceToInputFile);
                    log.Info("Auto correct values based on replace-to list: {0}", options.AutoCorrect ? "Yes" : "No");
                }
                if (options.EmailAsAccountName)
                {
                    log.Info("Auto replace AccountName with EmailAddress: {0}", options.EmailAsAccountName ? "Yes" : "No");
                }

                var users = new List<User>();
                var compareTo = new List<Tuple<string, string, string>>();
                var replaceTo = new List<Tuple<string, string>>();

                #region load users
                try
                {
                    log.Debug("start loading {0}.", options.InputFile);
                    var inputWorkbook = new XSSFWorkbook(new MemoryStream(File.ReadAllBytes(options.InputFile)));
                    
                    for (int i = 0; i < inputWorkbook.NumberOfSheets; i++)
                    {
                        var sheet = inputWorkbook.GetSheetAt(i);
                        for (int j = sheet.FirstRowNum + 1; j <= sheet.LastRowNum; j++)
                        {
                            var row = sheet.GetRow(j);
                            if (row == null)
                            {
                                continue;
                            }
                            var user = new User();
                            for (int k = row.FirstCellNum; k <= row.LastCellNum; k++)
                            {
                                var cell = row.GetCell(k);
                                var value = cell?.ToString().Trim();
                                value = !string.Equals("NULL", value, StringComparison.OrdinalIgnoreCase) ? value : null;
                                switch (k)
                                {
                                    case 0:
                                        user.FirstName = value;
                                        break;
                                    case 1:
                                        user.LastName = value;
                                        break;
                                    case 2:
                                        user.BusinessTitle = !string.IsNullOrEmpty(value) ? value : "Other";
                                        break;
                                    case 3:
                                        user.OrganizationUnit = value;
                                        break;
                                    case 4:
                                        user.PhoneType = !string.IsNullOrEmpty(value) ? value : "Other";
                                        break;
                                    case 5:
                                        user.PhoneNumber = !string.IsNullOrEmpty(value) ? value : "000-000-0000";
                                        break;
                                    case 6:
                                        user.UserRoles = value;
                                        break;
                                    case 7:
                                        user.AccountName = value;
                                        break;
                                    case 8:
                                        if (value != null)
                                        {
                                            if (!EmailRegex.IsMatch(value))
                                            {
                                                log.Warn("{0} has invalid format, skipped this row. Current Sheet/Row [{1}]/[{2}]", value, sheet.SheetName, row.RowNum);
                                                continue;
                                            }
                                            if (options.AutoCorrect)
                                            {
                                                if (value.Contains(","))
                                                {
                                                    log.Info("{0} auto corrected. Current Sheet/Row [{1}]/[{2}]", value, sheet.SheetName, row.RowNum);
                                                    value = value?.Replace(",", ".");
                                                }
                                            }
                                        }
                                        user.EmailAddress = value;
                                        break;
                                }
                            }
                            user.AccountName = user.AccountName ?? user.EmailAddress;//set email address as default account name
                            if (!string.IsNullOrEmpty(user.EmailAddress))
                            {
                                var loadedUser = users.FirstOrDefault(x => x.EmailAddress == user.EmailAddress);
                                if (loadedUser != null)
                                {
                                    log.Info("{0} already loaded. Current Sheet/Row [{1}]/[{2}]", user.EmailAddress, sheet.SheetName, row.RowNum);
                                }
                                else
                                {
                                    users.Add(user);
                                }
                            }
                        }
                    }
                }
                catch (Exception exception)
                {
                    log.Error(exception, "Cannot load users from input file {0}", options.InputFile);
                }
                log.Info("{0} users loaded.", users.Count);
                #endregion

                if (!string.IsNullOrEmpty(options.CompareToInputFile))
                {
                    #region load compare-to users
                    try
                    {
                        log.Debug("start loading {0}.", options.CompareToInputFile);
                        var inputWorkbook = new XSSFWorkbook(new MemoryStream(File.ReadAllBytes(options.CompareToInputFile)));

                        //original spreadsheet has 1 sheet
                        var sheet = inputWorkbook.GetSheetAt(0);

                        for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                        {
                            var row = sheet.GetRow(i);
                            if (row == null)
                            {
                                continue;
                            }
                            var key = row.GetCell(0)?.ToString()?.Trim();
                            var username = row.GetCell(1)?.ToString()?.Trim();
                            var email = row.GetCell(2)?.ToString()?.Trim();
                            if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(email))
                            {
                                compareTo.Add(new Tuple<string, string, string>(key, username, email));
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        log.Error(exception, "Cannot load users from compare to input file {0}", options.CompareToInputFile);
                    }
                    log.Info("{0} compare-to users loaded.", compareTo.Count);
                    #endregion

                    if (options.AutoCorrect)
                    {
                        #region correct username by following compre-to users
                        foreach (var user in users)
                        {
                            var existUser = compareTo.FirstOrDefault(x => string.Equals(x.Item3, user.EmailAddress, StringComparison.OrdinalIgnoreCase));
                            if (existUser != null && !string.Equals(existUser.Item2, user.AccountName))
                            {
                                log.Info("Account name replaced with {0} from {1} for {2}", existUser.Item2, user.AccountName, user.EmailAddress);
                                user.AccountName = existUser.Item2;
                            }
                        }
                        #endregion
                    }

                    var usersNotExist = users.Where(x => compareTo.All(y => !string.Equals(y.Item3, x.EmailAddress, StringComparison.OrdinalIgnoreCase))).ToList();
                    log.Info("Users cannot be found in our database: {0}", string.Join(",", usersNotExist.Select(x => x.EmailAddress)));

                    var usersNotFromClient =
                        compareTo.Where(
                            x =>
                                users.All(y => !string.Equals(x.Item3, y.EmailAddress, StringComparison.OrdinalIgnoreCase))).ToList();
                    log.Info("Users cannot be found from input file (keys): {0}", string.Join(",", usersNotFromClient.Select(x => x.Item1)));
                    log.Info("Users cannot be found from input file: {0}", string.Join(",", usersNotFromClient.Select(x => x.Item3)));
                }

                if (!string.IsNullOrEmpty(options.ReplaceToInputFile))
                {
                    #region load replace-to values
                    try
                    {
                        log.Debug("start loading {0}.", options.ReplaceToInputFile);
                        var inputWorkbook = new XSSFWorkbook(new MemoryStream(File.ReadAllBytes(options.ReplaceToInputFile)));

                        //original spreadsheet has 1 sheet
                        var sheet = inputWorkbook.GetSheetAt(0);

                        for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                        {
                            var row = sheet.GetRow(i);
                            if (row == null)
                            {
                                continue;
                            }
                            var originalValue = row.GetCell(0)?.ToString()?.Trim();
                            var newValue = row.GetCell(1)?.ToString()?.Trim();
                            if (!string.IsNullOrEmpty(originalValue) && !string.IsNullOrEmpty(newValue))
                            {
                                replaceTo.Add(new Tuple<string, string>(originalValue, newValue));
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        log.Error(exception, "Cannot load users from replace to input file {0}", options.ReplaceToInputFile);
                    }
                    #endregion

                    if (options.AutoCorrect)
                    {
                        foreach (var user in users)
                        {
                            #region correct OrganizationUnit by following replace-to values
                            var organizationUnit = replaceTo.FirstOrDefault(x => string.Equals(x.Item1, user.OrganizationUnit, StringComparison.OrdinalIgnoreCase));
                            if (organizationUnit != null)
                            {
                                log.Info("OrganizationUnit replaced with {0} from {1} for {2}", organizationUnit.Item2, user.OrganizationUnit, user.EmailAddress);
                                user.OrganizationUnit = organizationUnit.Item2;
                            }
                            #endregion

                            #region correct PhoneType by following replace-to values
                            var phoneType = replaceTo.FirstOrDefault(x => string.Equals(x.Item1, user.PhoneType, StringComparison.OrdinalIgnoreCase));
                            if (phoneType != null)
                            {
                                log.Info("PhoneType replaced with {0} from {1} for {2}", phoneType.Item2, user.PhoneType, user.EmailAddress);
                                user.PhoneType = phoneType.Item2;
                            }
                            #endregion

                            #region correct UserRoles by following replace-to values
                            var userRoles = replaceTo.FirstOrDefault(x => string.Equals(x.Item1, user.UserRoles, StringComparison.OrdinalIgnoreCase));
                            if (userRoles != null)
                            {
                                log.Info("UserRoles replaced with {0} from {1} for {2}", userRoles.Item2, user.UserRoles, user.EmailAddress);
                                user.UserRoles = userRoles.Item2;
                            }
                            #endregion
                        }
                    }
                }

                if (options.EmailAsAccountName)
                {
                    #region Repalce AccountName with EmailAddress
                    foreach (var user in users)
                    {
                        log.Info("Account name replaced with email address for {0}", user.EmailAddress);
                        user.AccountName = user.EmailAddress;
                    }
                    #endregion
                }

                #region write users into files
                if (users.Any())
                {
                    try
                    {
                        WriteWorkbook(users, options.OutputFile);
                        log.Info("New file generated: {0}", options.OutputFile);
                    }
                    catch (Exception exception)
                    {
                        log.Error(exception, "Cannot write users to PA output file {0}", options.OutputFile);
                    }
                }
                #endregion
            }
            else
            {
                Console.WriteLine(options.GetUsage());
            }
        }

        private static void WriteWorkbook(IList<User> users, string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                var outputWorkbook = new XSSFWorkbook();
                var staffSheet = outputWorkbook.CreateSheet("Staff");

                var headerRow = staffSheet.CreateRow(0);
                var headerCell0 = headerRow.CreateCell(0);
                headerCell0.SetCellValue("First Name");
                var headerCell1 = headerRow.CreateCell(1);
                headerCell1.SetCellValue("Last Name");
                var headerCell2 = headerRow.CreateCell(2);
                headerCell2.SetCellValue("Business Title");
                var headerCell3 = headerRow.CreateCell(3);
                headerCell3.SetCellValue("Organization Unit");
                var headerCell4 = headerRow.CreateCell(4);
                headerCell4.SetCellValue("Phone Type");
                var headerCell5 = headerRow.CreateCell(5);
                headerCell5.SetCellValue("Phone Number");
                var headerCell6 = headerRow.CreateCell(6);
                headerCell6.SetCellValue("User Roles");
                var headerCell7 = headerRow.CreateCell(7);
                headerCell7.SetCellValue("Account Name");
                var headerCell8 = headerRow.CreateCell(8);
                headerCell8.SetCellValue("Email Address");

                for (int i = 0; i < users.Count; i++)
                {
                    var user = users[i];

                    var userRow = staffSheet.CreateRow(i + 1);

                    var userCell0 = userRow.CreateCell(0);
                    userCell0.SetCellValue(user.FirstName);
                    var userCell1 = userRow.CreateCell(1);
                    userCell1.SetCellValue(user.LastName);
                    var userCell2 = userRow.CreateCell(2);
                    userCell2.SetCellValue(user.BusinessTitle);
                    var userCell3 = userRow.CreateCell(3);
                    userCell3.SetCellValue(user.OrganizationUnit);
                    var userCell4 = userRow.CreateCell(4);
                    userCell4.SetCellValue(user.PhoneType);
                    var userCell5 = userRow.CreateCell(5);
                    userCell5.SetCellValue(user.PhoneNumber);
                    var userCell6 = userRow.CreateCell(6);
                    userCell6.SetCellValue(user.UserRoles);
                    var userCell7 = userRow.CreateCell(7);
                    userCell7.SetCellValue(user.AccountName);
                    var userCell8 = userRow.CreateCell(8);
                    userCell8.SetCellValue(user.EmailAddress);
                }

                outputWorkbook.Write(stream);
            }
        }

        private static Logger log = LogManager.GetCurrentClassLogger();
        private static readonly Regex EmailRegex = new Regex(@"^([a-zA-Z0-9_\-\.\']+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$");

        private class User
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string BusinessTitle { get; set; }
            public string OrganizationUnit { get; set; }
            public string PhoneType { get; set; }
            public string PhoneNumber { get; set; }
            public string UserRoles { get; set; }
            public string AccountName { get; set; }
            public string EmailAddress { get; set; }
        }

        private class Options
        {
            [Option('i', "input", Required = true, HelpText = "Input file to read. XLSX file that contains n sheet and 9 columns.")]
            public string InputFile { get; set; }

            [Option('o', "output", Required = true, HelpText = "Output file to write. XLSX file.")]
            public string OutputFile { get; set; }

            [Option('c', "compare", Required = false, HelpText = "Compare to input file to read. XLSX file that contains 1 sheet and 3 columns.")]
            public string CompareToInputFile { get; set; }

            [Option('r', "replace", Required = false, HelpText = "Compare to input file to read. XLSX file that contains 1 sheet and 3 columns.")]
            public string ReplaceToInputFile { get; set; }

            [Option('a', "auto-correct", Required = false, DefaultValue = false, HelpText = "Auto correct values based on compare-to list and replace-to list.")]
            public bool AutoCorrect { get; set; }

            [Option('e', "email-as-account", Required = false, DefaultValue = false, HelpText = "Auto correct username to email address.")]
            public bool EmailAsAccountName { get; set; }

            [HelpOption]
            public string GetUsage()
            {
                return HelpText.AutoBuild(this);
            }
        }
    }
}