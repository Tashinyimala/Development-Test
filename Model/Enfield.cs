using Dev_Test.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Dev_Test.Model
{
    class Enfield : Utils
    {
        //OleDBConnection for MS Access Database
        private OleDbConnection connection = new OleDbConnection();
        private string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\tnyima\Downloads\Development Test\Sample\Enfield_Xml.accdb;Persist Security Info=False;";

        public string AccountNumber = "", SequenceNo = "", PrevYearDebt = "", AccountBalance = "", AdminAreaName = "", AdminArea = "", BillingIndicator = "",
                          PersonReference1 = "", citizenNameInitials1 = "", CitizenNameTitle1 = "", CitizenNameForename1 = "", CitizenNameSurname1 = "", AssociateUKAddressLine1_1 = "",
                          AssociateUKAddressLine2_1 = "", AssociateUKAddressPostCode1 = "",
                          PersonReference2 = "", citizenNameInitials2 = "", CitizenNameTitle2 = "", CitizenNameForename2 = "", CitizenNameSurname2 = "", AssociateUKAddressLine1_2 = "",
                          AssociateUKAddressLine2_2 = "", AssociateUKAddressPostCode2 = "", BillDocumentInd = "",
                          AcountYearDescription = "", AccountYearCurrFinYearFlag = "", AccountYearCountRec = "",
                          YearPeriodStartDate = "", YearPeriodEndDate = "", AdjustmentDescription = "", AdjustmentDetailDescription = "", DetailDescriptionAmount = "",
                          AccountYearMOP = "", AccountYearTypeOfPayment = "", AccountYearRecoveryFlag = "", AccountYearRecovery = "", AccountYearYearBalance = "", AccountYearBenefitInd = "",
                          AccountYearSparFlag = "", AccountYearPropertyLAPropertyReference = "", AccountYearPropertyAddressLine1 = "", AccountYearPropertyAddressLine2 = "",
                          AccountYearPropertyAddressPostCode = "",
                          AccountYearChargeComponent1Description = "", AccountYearChargeComponent1AnnualAmount = "", AccountYearChargeComponent1PercentageChange = "", AccountYearChargeComponent1AmountChange = "",
                          AccountYearChargeComponent2Description = "", AccountYearChargeComponent2AnnualAmount = "", AccountYearChargeComponent2PercentageChange = "", AccountYearChargeComponent2AmountChange = "",
                          AccountYearChargeComponent3Description = "", AccountYearChargeComponent3AnnualAmount = "", AccountYearChargeComponent3PercentageChange = "", AccountYearChargeComponent3AmountChange = "",
                          AccountYearChargeComponent4Description = "", AccountYearChargeComponent4AnnualAmount = "", AccountYearChargeComponent4PercentageChange = "", AccountYearChargeComponent4AmountChange = "",
                          AccountYearChargeComponent4ParishTotalSpend = "",
                          AccountYearChargeComponent5Description = "", AccountYearChargeComponent5AnnualAmount = "", AccountYearChargeComponent5PercentageChange = "", AccountYearChargeComponent5AmountChange = "",
                          AccountYeaLiabilityPeriodGrossAmount = "", AccountYearLiabilityPeriodNetAmount = "", AccountYearLiabilityPeriodPeriodStartDate = "", AccountYearLiabilityPeriodPeriodEndDate = "",
                          AccountYearLiabilityPeriodNumberofDays = "", AccountYearPropertyCurrentBand = "", AccountYearPropertyCurrentParish = "", AccountYearPropertyBillReasonCode = "", AccountYearPropertyBillReasonDesc = "",
                          AccountYearPropertyProvisional = "",
                          Barcode = "", OCRLine = "";

        public void XMLFileReader(string fileLocation)
        {
            XmlDocument xmlDocument = new XmlDocument();
            try
            {
                xmlDocument.Load(fileLocation);
                XmlNodeList nodes = xmlDocument.DocumentElement.SelectNodes("/Bill/Account");

                UpdateHeader();

                foreach (XmlNode node in nodes)
                {
                    AccountNumber = node.SelectSingleNode("AccountNumber").InnerText;
                    SequenceNo = node.SelectSingleNode("SequenceNo").InnerText;
                    PrevYearDebt = node.SelectSingleNode("PrevYearDebt").InnerText;
                    AccountBalance = node.SelectSingleNode("AccountBalance").InnerText;
                    AdminAreaName = node.SelectSingleNode("LASignatory").InnerText;
                    AdminArea = node.SelectSingleNode("AdminArea").InnerText;
                    BillDocumentInd = node.SelectSingleNode("BillDocumentInd").InnerText;
                    Barcode = node.SelectSingleNode("Barcode").InnerText;
                    OCRLine = node.SelectSingleNode("OCRLine").InnerText;

                    // AssociateList/Associate Nodes
                    XmlNodeList associateNodes = node.SelectNodes("AssociateList/Associate");
                    XmlNode associateNode = associateNodes[0];

                    BillingIndicator = associateNode.Attributes["BillingIndicator"].Value;
                    string[] associateNodeInnerText = GetInnerTextFromXML(associateNode.InnerXml);
                    PersonReference1 = $"{associateNodeInnerText[1]}";
                    citizenNameInitials1 = $"{associateNodeInnerText[4]} {associateNodeInnerText[5]} {associateNodeInnerText[6]}";
                    CitizenNameTitle1 = $"{associateNodeInnerText[7]}";
                    CitizenNameForename1 = $"{associateNodeInnerText[9]} {associateNodeInnerText[10]}";
                    CitizenNameSurname1 = $"{associateNodeInnerText[12]}";
                    AssociateUKAddressLine1_1 = $"{associateNodeInnerText[17]} {associateNodeInnerText[18]} {associateNodeInnerText[19]}";
                    AssociateUKAddressLine2_1 = $"{associateNodeInnerText[21]}";
                    AssociateUKAddressPostCode1 = $"{associateNodeInnerText[23]} {associateNodeInnerText[24]} ";

                    if (associateNodes.Count > 1)
                    {
                        XmlNode associateNode1 = associateNodes[1];
                        string[] associate1NodeInnerText = GetInnerTextFromXML(associateNode1.InnerXml);
                        PersonReference2 = $"{associate1NodeInnerText[1]}";
                        citizenNameInitials2 = $"{associate1NodeInnerText[4]} {associate1NodeInnerText[5]} {associate1NodeInnerText[6]}";
                        CitizenNameTitle2 = $"{associate1NodeInnerText[7]}";
                        CitizenNameForename2 = $"{associate1NodeInnerText[9]} {associate1NodeInnerText[10]}";
                        CitizenNameSurname2 = $"{associate1NodeInnerText[12]}";
                        AssociateUKAddressLine1_2 = $"{associate1NodeInnerText[17]} {associate1NodeInnerText[18]} {associate1NodeInnerText[19]}";
                        AssociateUKAddressLine2_2 = $"{associate1NodeInnerText[21]}";
                        AssociateUKAddressPostCode2 = $"{associate1NodeInnerText[23]} {associate1NodeInnerText[24]} ";
                    }


                    XmlNodeList AccountYearNodes = node.SelectNodes("AccountYear");
                    XmlNode accountYearNode1 = AccountYearNodes[0];

                    //AccountYear Node
                    AcountYearDescription = accountYearNode1.SelectSingleNode("YearDescription").InnerText;
                    AccountYearCurrFinYearFlag = accountYearNode1.SelectSingleNode("CurrFinYearFlag").InnerText;
                    AccountYearCountRec = accountYearNode1.SelectSingleNode("CountRec").InnerText;

                    String[] YearPeriodText = GetInnerTextFromXML(accountYearNode1.SelectSingleNode("YearPeriod").InnerXml);
                    YearPeriodStartDate = YearPeriodText[1];
                    YearPeriodEndDate = YearPeriodText[3];

                    string[] AccountYearAdjustments = GetInnerTextFromXML(accountYearNode1.SelectSingleNode("Adjustment").InnerXml);
                    AdjustmentDescription = AccountYearAdjustments[1];
                    AdjustmentDetailDescription = AccountYearAdjustments[3];
                    DetailDescriptionAmount = AccountYearAdjustments[5];

                    AccountYearMOP = accountYearNode1.SelectSingleNode("MOP").InnerText;
                    AccountYearRecoveryFlag = accountYearNode1.SelectSingleNode("Recovery/RecoveryFlag").InnerText;
                    AccountYearTypeOfPayment = accountYearNode1.SelectSingleNode("TypeOfPayment").InnerText;
                    AccountYearRecovery = accountYearNode1.SelectSingleNode("Recovery").InnerText;
                    AccountYearYearBalance = accountYearNode1.SelectSingleNode("YearBalance").InnerText;
                    AccountYearBenefitInd = accountYearNode1.SelectSingleNode("BenefitInd").InnerText;
                    AccountYearSparFlag = accountYearNode1.SelectSingleNode("SparFlag").InnerText;

                    // AccountYear Property Nodes
                    AccountYearPropertyLAPropertyReference = accountYearNode1.SelectSingleNode("Property/LAPropertyReference").InnerText;

                    XmlNodeList PropertyAddressNodes = accountYearNode1.SelectNodes("Property/PropertyAddress");
                    XmlNode PropertyAddressNode = PropertyAddressNodes[0];
                    string[] AccountYearPropertyPropertyAddressInnerText = GetInnerTextFromXML(PropertyAddressNode.InnerXml);
                    AccountYearPropertyAddressLine1 = $"{AccountYearPropertyPropertyAddressInnerText[2]} {AccountYearPropertyPropertyAddressInnerText[3]} {AccountYearPropertyPropertyAddressInnerText[4]}";
                    AccountYearPropertyAddressLine2 = $"{AccountYearPropertyPropertyAddressInnerText[6]}";
                    AccountYearPropertyAddressPostCode = $"{AccountYearPropertyPropertyAddressInnerText[8]} {AccountYearPropertyPropertyAddressInnerText[9]}";

                    // ChargeComponent Nodes
                    XmlNodeList PropertyChargeComponentNodes = accountYearNode1.SelectNodes("Property/ChargeComponent");

                    if (PropertyChargeComponentNodes.Count <= 5)
                    {
                        // 1st ChargeComponent node
                        XmlNode PropertyChargeComponent1 = PropertyChargeComponentNodes[0];
                        AccountYearChargeComponent1Description = PropertyChargeComponent1.SelectSingleNode("Description").InnerText;
                        AccountYearChargeComponent1AnnualAmount = PropertyChargeComponent1.SelectSingleNode("AnnualAmount").InnerText;
                        AccountYearChargeComponent1PercentageChange = PropertyChargeComponent1.SelectSingleNode("PercentageChange").InnerText;
                        AccountYearChargeComponent1AmountChange = PropertyChargeComponent1.SelectSingleNode("AmountChange").InnerText;


                        XmlNode PropertyChargeComponent2 = PropertyChargeComponentNodes[1];
                        AccountYearChargeComponent2Description = PropertyChargeComponent2.SelectSingleNode("Description").InnerText;
                        AccountYearChargeComponent2AnnualAmount = PropertyChargeComponent2.SelectSingleNode("AnnualAmount").InnerText;
                        AccountYearChargeComponent2PercentageChange = PropertyChargeComponent2.SelectSingleNode("PercentageChange").InnerText;
                        AccountYearChargeComponent2AmountChange = PropertyChargeComponent2.SelectSingleNode("AmountChange").InnerText;

                        XmlNode PropertyChargeComponent3 = PropertyChargeComponentNodes[2];
                        AccountYearChargeComponent3Description = PropertyChargeComponent3.SelectSingleNode("Description").InnerText;
                        AccountYearChargeComponent3AnnualAmount = PropertyChargeComponent3.SelectSingleNode("AnnualAmount").InnerText;
                        AccountYearChargeComponent3PercentageChange = PropertyChargeComponent3.SelectSingleNode("PercentageChange").InnerText;
                        AccountYearChargeComponent3AmountChange = PropertyChargeComponent3.SelectSingleNode("AmountChange").InnerText;

                        XmlNode PropertyChargeComponent4 = PropertyChargeComponentNodes[3];
                        AccountYearChargeComponent4Description = PropertyChargeComponent4.SelectSingleNode("Description").InnerText;
                        AccountYearChargeComponent4AnnualAmount = PropertyChargeComponent4.SelectSingleNode("AnnualAmount").InnerText;
                        AccountYearChargeComponent4PercentageChange = PropertyChargeComponent4.SelectSingleNode("PercentageChange").InnerText;
                        AccountYearChargeComponent4ParishTotalSpend = PropertyChargeComponent4.SelectSingleNode("ParishTotalSpend").InnerText;
                        AccountYearChargeComponent4AmountChange = PropertyChargeComponent4.SelectSingleNode("AmountChange").InnerText;

                        XmlNode PropertyChargeComponent5 = PropertyChargeComponentNodes[4];
                        AccountYearChargeComponent5Description = PropertyChargeComponent5.SelectSingleNode("Description").InnerText;
                        AccountYearChargeComponent5AnnualAmount = PropertyChargeComponent5.SelectSingleNode("AnnualAmount").InnerText;
                        AccountYearChargeComponent5PercentageChange = PropertyChargeComponent5.SelectSingleNode("PercentageChange").InnerText;
                        AccountYearChargeComponent5AmountChange = PropertyChargeComponent5.SelectSingleNode("AmountChange").InnerText;
                    }
                    else
                    {
                        MessageBox.Show("Charge");
                    }

                    AccountYeaLiabilityPeriodGrossAmount = accountYearNode1.SelectSingleNode("Property/LiabilityPeriod/GrossAmount").InnerText;
                    AccountYearLiabilityPeriodNetAmount = accountYearNode1.SelectSingleNode("Property/LiabilityPeriod/NetAmount").InnerText;
                    AccountYearLiabilityPeriodPeriodStartDate = accountYearNode1.SelectSingleNode("Property/LiabilityPeriod/Period/StartDate").InnerText;
                    AccountYearLiabilityPeriodPeriodEndDate = accountYearNode1.SelectSingleNode("Property/LiabilityPeriod/Period/EndDate").InnerText;
                    AccountYearLiabilityPeriodNumberofDays = accountYearNode1.SelectSingleNode("Property/LiabilityPeriod/Period/NumberofDays").InnerText;

                    AccountYearPropertyCurrentBand = accountYearNode1.SelectSingleNode("Property/CurrentBand").InnerText;
                    AccountYearPropertyCurrentParish = accountYearNode1.SelectSingleNode("Property/CurrentParish").InnerText;
                    AccountYearPropertyBillReasonCode = accountYearNode1.SelectSingleNode("Property/BillReasonCode").InnerText;
                    AccountYearPropertyBillReasonDesc = accountYearNode1.SelectSingleNode("Property/BillReasonDesc").InnerText;
                    AccountYearPropertyProvisional = accountYearNode1.SelectSingleNode("Property/Provisional").InnerText;

                    Console.WriteLine(AccountYearPropertyBillReasonDesc);

                    string[] items = { AccountNumber, SequenceNo, PrevYearDebt, AccountBalance, AdminAreaName, AdminArea, BillingIndicator,
                                       PersonReference1, citizenNameInitials1, CitizenNameTitle1, CitizenNameForename1, CitizenNameSurname1, AssociateUKAddressLine1_1,
                                       AssociateUKAddressLine2_1, AssociateUKAddressPostCode1,
                                       PersonReference2, citizenNameInitials2, CitizenNameTitle2, CitizenNameForename2, CitizenNameSurname2, AssociateUKAddressLine1_2,
                                       AssociateUKAddressLine2_2, AssociateUKAddressPostCode2, BillDocumentInd,
                                       AcountYearDescription, AccountYearCurrFinYearFlag, AccountYearCountRec,
                                       YearPeriodStartDate, YearPeriodEndDate, AdjustmentDescription, AdjustmentDetailDescription, DetailDescriptionAmount,
                                       AccountYearMOP, AccountYearTypeOfPayment, AccountYearRecoveryFlag, AccountYearRecovery, AccountYearYearBalance, AccountYearBenefitInd,
                                       AccountYearSparFlag, AccountYearPropertyLAPropertyReference, AccountYearPropertyAddressLine1, AccountYearPropertyAddressLine2,
                                       AccountYearPropertyAddressPostCode,
                                       AccountYearChargeComponent1Description, AccountYearChargeComponent1AnnualAmount, AccountYearChargeComponent1PercentageChange, AccountYearChargeComponent1AmountChange,
                                       AccountYearChargeComponent2Description, AccountYearChargeComponent2AnnualAmount, AccountYearChargeComponent2PercentageChange, AccountYearChargeComponent2AmountChange,
                                       AccountYearChargeComponent3Description, AccountYearChargeComponent3AnnualAmount, AccountYearChargeComponent3PercentageChange, AccountYearChargeComponent3AmountChange,
                                       AccountYearChargeComponent4Description, AccountYearChargeComponent4AnnualAmount, AccountYearChargeComponent4PercentageChange, AccountYearChargeComponent4AmountChange,
                                       AccountYearChargeComponent4ParishTotalSpend,
                                       AccountYearChargeComponent5Description, AccountYearChargeComponent5AnnualAmount, AccountYearChargeComponent5PercentageChange, AccountYearChargeComponent5AmountChange,
                                       AccountYeaLiabilityPeriodGrossAmount, AccountYearLiabilityPeriodNetAmount, AccountYearLiabilityPeriodPeriodStartDate, AccountYearLiabilityPeriodPeriodEndDate,
                                       AccountYearLiabilityPeriodNumberofDays, AccountYearPropertyCurrentBand, AccountYearPropertyCurrentParish, AccountYearPropertyBillReasonCode, AccountYearPropertyBillReasonDesc,
                                       AccountYearPropertyProvisional,Barcode, OCRLine
                                     };

                    // Adding data to datatable
                    myTable.Rows.Add(items);

                    SaveDataToMSAccessDB();
                }
            }
            finally
            {
                connection.Close();
            }
        }

        public void UpdateHeader()
        {
            for (int col = 1; col <= 100; col++)
            {
                myTable.Columns.Add($"Field{col}");
            }
        }

        public string[] GetInnerTextFromXML(string xmlText)
        {
            // Match any XML tag (opening or closing tags)
            // follwed by any successive whitespace
            Regex regex = new Regex(@"<[^>].+?>", RegexOptions.Singleline);
            string resultText = regex.Replace(xmlText, @" ");

            // Text to String Array split by space
            char[] whitespace = new char[] { ' ', '\t' };
            string[] resultTexts = resultText.Split(whitespace);

            return resultTexts;
        }


        // Saving Datatable to MS Access Database
        private void SaveDataToMSAccessDB()
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                var dataAdapter = new OleDbDataAdapter("SELECT * FROM Import_Table", connection);

                using (OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(dataAdapter))
                {
                    try
                    {
                        connection.Open();
                        commandBuilder.ConflictOption = ConflictOption.CompareRowVersion;
                        dataAdapter.Update(myTable);
                    }
                    catch (OleDbException ex)
                    {
                        MessageBox.Show(ex.Message, "OleDbException Error");
                    }
                    catch (Exception x)
                    {
                        MessageBox.Show(x.Message, "Exception Error");
                    }
                }

                connection.Close();
                dataAdapter.Dispose();
            }
        }
    }
}
