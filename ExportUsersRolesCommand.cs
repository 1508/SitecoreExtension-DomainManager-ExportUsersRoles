using System.Globalization;
using System.IO;
using System.Linq;
using Sitecore;
using Sitecore.Data.Serialization;
using Sitecore.Diagnostics;
using Sitecore.IO;
using Sitecore.Security.Accounts;
using Sitecore.Security.Domains;
using Sitecore.Shell.Applications.Dialogs.ProgressBoxes;
using Sitecore.Shell.Framework.Commands;
using Sitecore.Shell.Framework.Commands.Serialization;
using Sitecore.Web.UI.Sheer;
using System;
using System.Collections.Generic;
using SpreadsheetGear;

namespace SitecoreExtension.DomainManager.ExportUsersRoles
{
    public class ExportUsersRolesCommand : Command
    {
        /// <summary>
        /// Executes the command in the specified context.
        /// 
        /// </summary>
        /// <param name="context">The context.</param><contract><requires name="context" condition="not null"/></contract>
        public override void Execute(CommandContext context)
        {
            Assert.ArgumentNotNull((object)context, "context");
            string str = context.Parameters["domainname"];
            if (string.IsNullOrEmpty(str))
                SheerResponse.Alert("Please select a domain first.", new string[0]);
            else
            {
                Domain domain = Domain.GetDomain(str);
                if (!(domain != (Domain)null))
                {
                    SheerResponse.Alert(string.Format("Domain name '{0}' could not be resolved", str), new string[0]);
                    return;
                }

                var reportFolder = Sitecore.Configuration.Settings.DataFolder + "/reports";
                var reportFolderPath = FileUtil.MapPath(reportFolder);
                if (!Directory.Exists(reportFolderPath))
                {
                    // Create folder for temporary reports if not present
                    Log.Info(
                        string.Format("SitecoreExtension.DomainManager.ExportUsersRoles: Creating reports folder {0}",
                            reportFolderPath), this);
                    Directory.CreateDirectory(reportFolderPath);
                }

                Log.Audit(
                    string.Format("SitecoreExtension.DomainManager.ExportUsersRoles: Generating report for {0}", str),
                    this);

                // Save to FilePath
                var filename = string.Format("UserReport-{0}-{1}.xlsx", str,
                    DateTime.Now.ToUniversalTime().ToString("yyyy-MM-dd-hh-mm"));

                try
                {
                    // Generate UserData Listning
                    var userDatas = domain.GetUsers().Select(user => new UserData(user)).ToList();
                    if (!userDatas.Any())
                    {
                        SheerResponse.Alert(string.Format("Domain name '{0}' contains no user accounts.", str), new string[0]);
                        return;
                    }

                    Log.Audit(
                        string.Format("SitecoreExtension.DomainManager.ExportUsersRoles: Exporting {1} users to {0} ", filename, userDatas.Count()),
                        this);

                    // Dump to temporary file
                    GenerateExcel(userDatas, Path.Combine(reportFolderPath, filename));
                    
                    // Send Path to shell
                    SheerResponse.Download(string.Format("{0}/{1}", reportFolder, filename));
                }
                catch (Exception exception)
                {
                    Log.Error("Could not generate report due to exception", exception, this);
                    throw;
                }
            }
        }

        private void GenerateExcel(List<UserData> userdatas, string filename)
        {
            var workbook = SpreadsheetGear.Factory.GetWorkbook();
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Users";
            // Make the headers
            worksheet.Cells["A1"].Value = "Name";
            worksheet.Cells["A1"].Font.Bold = true;
            worksheet.Cells["B1"].Value = "Email";
            worksheet.Cells["B1"].Font.Bold = true;
            worksheet.Cells["C1"].Value = "Login";
            worksheet.Cells["C1"].Font.Bold = true;
            worksheet.Cells["D1"].Value = "Domain";
            worksheet.Cells["D1"].Font.Bold = true;
            worksheet.Cells["E1"].Value = "Description";
            worksheet.Cells["E1"].Font.Bold = true;
            worksheet.Cells["F1"].Value = "State";
            worksheet.Cells["F1"].Font.Bold = true;
            worksheet.Cells["G1"].Value = "IsAdministrator";
            worksheet.Cells["G1"].Font.Bold = true;
            worksheet.Cells["H1"].Value = "Roles";
            worksheet.Cells["H1"].Font.Bold = true;
            
            // start from row 2
            int i = 2;

            foreach (var userData in userdatas)
            {
                worksheet.Cells["A" + i].Value = userData.Name;
                worksheet.Cells["B" + i].Value = userData.Email;
                worksheet.Cells["C" + i].Value = userData.Login;
                worksheet.Cells["D" + i].Value = userData.DomainName;
                worksheet.Cells["E" + i].Value = userData.Description;
                worksheet.Cells["F" + i].Value = userData.State;
                worksheet.Cells["G" + i].Value = userData.IsAdministrator;
                worksheet.Cells["H" + i].Value = string.Join(";", userData.Roles);
                i++;
            }

            workbook.SaveAs(filename, FileFormat.OpenXMLWorkbook);
        }

        [Serializable]
        private class UserData
        {
            public string Name { get; set; }
            public string Email { get; set; }
            public string Login { get; set; }
            public string Password { get; set; }
            public string DomainName { get; set; }
            public string Description { get; set; }
            public List<string> Roles { get; set; }
            public string State { get; set; }
            public bool IsAdministrator { get; set; }
            public string RawLine { get; set; }

            public UserData()
            {
                this.Name = this.Email = this.Login = this.RawLine = string.Empty;
                this.Roles = new List<string>();
            }

            public UserData(User user)
                : base()
            {
                this.Name = this.Email = this.Login = this.RawLine = string.Empty;
                this.Roles = new List<string>();

                this.Name = user.Profile.FullName;
                this.Email = user.Profile.Email;
                this.Login = user.LocalName;
                this.DomainName = user.GetDomainName();
                this.Description = user.Description;
                this.State = user.Profile.State;
                this.IsAdministrator = user.IsAdministrator;
                foreach (Role role in user.Roles)
                {
                    this.Roles.Add(role.LocalName);
                }
            }
        }
    }
}