using System;
using System.IO;  //Needed to get folder path
using System.Drawing; //Needed to make use of Insight
using System.Threading; //used for the Thread.Sleep
using Excel = Microsoft.Office.Interop.Excel;   //Add a Reference to the EXCEL namespace
using NUnit.Framework;
using HP.LFT.SDK;
using HP.LFT.SDK.Web;
using HP.LFT.SDK.Insight;

namespace AOB_LeanFTv1
{
    [TestFixture]
    public class LeanFtTest : UnitTestClassBase
    {
        private IBrowser browser;
        private AOBModel appModel;
        private BrowserType browserType;
        //switch baseURI for if you have internet connection or if 
        //you are accessing local alm-aob image
        private string baseUri = "http://15.126.221.115:47001/advantage/";
        //private string baseUri = "http://alm-aob:47001/advantage/";

        [TestFixtureSetUp]
        public void TestFixtureSetUp()
        {
            // Setup once per fixture
            //Turn on screenshot capture
            Reporter.SnapshotCaptureLevel = HP.LFT.Report.CaptureLevel.All;
        }

        [SetUp]
        public void SetUp()
        {
            #region Create and start browser
            browserType = BrowserType.Chrome;
            browser = BrowserFactory.Launch(browserType);
            browser.Navigate(baseUri);
            appModel = new AOBModel(browser);
            #endregion
        }
        public void login ()
        {
            #region Login example using descriptive
            /*
            browser.Describe<IEditField>(new EditFieldDescription
            {
                Type = @"text",
                TagName = @"INPUT",
                Name = @"j_username"
            }).SetValue("jojo");

            //The following lines will be manually added but are here for reference
            browser.Describe<IEditField>(new EditFieldDescription
            {
                Type = @"password",
                TagName = @"INPUT",
                Name = @"j_password"
            }).SetSecure("566b5022398f18e63353f718002a474753645c3a5e90e4219266bf134698");

            //The below line should be shown how to add through the spy
            browser.Describe<IButton>(new ButtonDescription
            {
                ButtonType = @"submit",
                TagName = @"INPUT",
                Name = @"Login"
            }).Click();
            */
            #endregion

            #region Login using the application model
            appModel.AdvantageOnlineBankingPage.UserName.SetValue("jojo");
            appModel.AdvantageOnlineBankingPage.UserPassword.SetSecure("566b5022398f18e63353f718002a474753645c3a5e90e4219266bf134698");
            appModel.AdvantageOnlineBankingPage.LoginButton.Click();
            #endregion
        }

        [Test]
        public void TestLogin()
        {
            login();

            #region validate successful login
            // In this section you can see the usage of Asserts
            // along with the LeanFT reporting to have an audit trail
            // and to be able to easily view results since these tests
            // may be part of a Continuous Test or Continuous Integration
            // process
            try
            {
                Assert.AreEqual("Account", appModel.AdvantageOnlineBankingPage.AccountsLink.InnerText.TrimEnd());
                Reporter.ReportEvent("Login check", "validate good login", HP.LFT.Report.Status.Passed);
            } catch(Exception e)
            {
                Reporter.ReportEvent("Login check", "login fail", HP.LFT.Report.Status.Failed, e, appModel.AdvantageOnlineBankingPage.GetSnapshot());
                Assert.Fail();
            }
            appModel.AdvantageOnlineBankingPage.LogoutLink.Click();
            #endregion
        }

        [Test]
        public void TestLoginInsight()
        {
            string imageFolder;
            
            // Not using the login() function here as this test is
            // to show the use of Insight image for objects
            appModel.AdvantageOnlineBankingPage.UserName.SetValue("jojo");
            appModel.AdvantageOnlineBankingPage.UserPassword.SetSecure("566b5022398f18e63353f718002a474753645c3a5e90e4219266bf134698");

            #region get image to use
            //using the current execution directory to find the image(s) to be used
            switch (browserType.ToString())
            {
                case "Chrome":
                    imageFolder = Directory.GetCurrentDirectory() + @"\..\..\InsightImages\imgLogin.PNG";
                    break;
                case "InternetExplorer":
                    imageFolder = Directory.GetCurrentDirectory() + @"\..\..\InsightImages\imgLoginIE.PNG";
                    break;
                case "Firefox":
                    imageFolder = Directory.GetCurrentDirectory() + @"\..\..\InsightImages\imgLoginFF.PNG";
                    break;
                default:
                    imageFolder = Directory.GetCurrentDirectory() + @"\..\..\InsightImages\imgLogin.PNG";
                    break;
            }
            #endregion

            try
            {
                Image image = Image.FromFile(imageFolder);
                #region NOTES HERE
                //As per the documentation, you can add a sensitivity %.  I have found that 99
                //seems to be very good when I have changed my resolution from 1024x768 to 2880x1800
                //and the image is still found.
                //NOTE: This was found to be true when working with Chrome
                #endregion
                var loginButton = browser.Describe<IInsightObject>(new InsightDescription(image, 99));
                loginButton.Click();
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("Get Insight Image", "Object image not found in "+browserType.ToString(), HP.LFT.Report.Status.Failed, e, imageFolder);
                Assert.Fail("Insight Object Not Found");
            }
            //appModel.AdvantageOnlineBankingPage.LogoutLink.Click();
        }

        // Declare an array to hold the various screens we are going to validate are available
        static string[] strMenuItems = new string[] { "Accounts", "Bill Pay", "Credit Cards" };

        [Test, TestCaseSource("strMenuItems")]
        public void TestScreenExists(string strScreenName)
        {
            login();

            #region select menu link from left panel of web page
            // Click on the link in the left hand menu based on the passed in screen name
            browser.Describe<ILink>(new LinkDescription
            {
                TagName = @"A",
                InnerText = @strScreenName
            }).Click();
            #endregion

            #region get the center banner object
            // Set a variable to hold the object for the screen banner.  Note, the banner is in uppercase, therefore
            // the screenname must be converted to uppercase.
            var objScreenBanner = browser.Describe<IWebElement>(new WebElementDescription
            {
                ClassName = @"center-name-text",
                TagName = @"SPAN",
            });
            #endregion

            if (objScreenBanner.Exists()){
                Reporter.ReportEvent("Availablity", strScreenName + " screen is Available", HP.LFT.Report.Status.Passed, browser.GetSnapshot());
            } else {
                Reporter.ReportEvent("Availablity", strScreenName + " screen is NOT Available", HP.LFT.Report.Status.Failed, browser.GetSnapshot());
                Assert.Fail();
            }
        }

        // Test to transfer money from one account to another.
        [Test]
        public void TransferMoney()
        {
            // The assumption is the file will be located with the project
            string strDataSheet = Directory.GetCurrentDirectory() + @"\..\..\DataFiles\MoneyTransfer.xlsx";

            #region define Excell assets
            // Set up acccess to various EXCEL assets
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range DataRange;
            #endregion

            #region define variable for looping
            // Declare the strings to hold the from account, to account and the amount to be transferred
            string strFromAccount, strToAccount, strAmount;

            // Define the Row Count, the columns in the spreadsheet that hold the 'From Account', 'To Account' and 'Amount' respectively.

            int rCnt = 0;
            int FrmAccountCol = 1;
            int ToAccountCol = 2;
            int AmountCol = 3;
            #endregion

            login();

            // Open a new instance EXCEL.  Then open the data sheet that's defined the the global variable strDataSheet
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(strDataSheet, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            // Get the work sheet named 'transfers'
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["transfers"];

            // Find the 'used range' of the excel sheet
            DataRange = xlWorkSheet.UsedRange;

            #region loop thu Excel for transferring funds
            // Starting at row 2 (row 1 is the header information) loop through each row that has data
            for (rCnt = 2; rCnt <= DataRange.Rows.Count; rCnt++)
            {
                // First click on the 'Money Transfer' link in the left menu
                appModel.AdvantageOnlineBankingPage.MoneyTransferLink.Click();

                // Read each column text for the current loop item
                strFromAccount = (string)(DataRange.Cells[rCnt, FrmAccountCol] as Excel.Range).Text;
                strToAccount = (string)(DataRange.Cells[rCnt, ToAccountCol] as Excel.Range).Text;
                strAmount = (string)(DataRange.Cells[rCnt, AmountCol] as Excel.Range).Text;

                // As the data in the EXCEL sheet does not contain ALL the text in the selection boxes 'From Account'
                // and 'To account' (i.e. the selection box items contain the Account type, the account number AND
                // the amount currently in that account [which obviously changes]) we have to look through
                // each of the items in the selection box and see if it 'contains' the data in the EXCEL sheet 
                // (which is only the Account Type and the Account Number).  Therefore, first loop through all
                // of the available selections in the list.

                foreach (IListItem MyItems in appModel.AdvantageOnlineBankingPage.FromAccountListBox.Items)
                {
                    // If the current list item contains the text from the spreadsheet - Then reset the
                    // string holding the account number to be the full text and then select it from the list

                    if (MyItems.Text.Contains(strFromAccount))
                    {
                        strFromAccount = MyItems.Text;
                        appModel.AdvantageOnlineBankingPage.FromAccountListBox.Select(strFromAccount);

                    }
                    else if (MyItems.Text.Contains(strToAccount))
                    {

                        strToAccount = MyItems.Text;
                        appModel.AdvantageOnlineBankingPage.ToAccountListBox.Select(strToAccount);
                    }
                }

                // Click the 'Next' button.
                appModel.AdvantageOnlineBankingPage.NextButton.Click();

                // Set the amount value, the transfer date to todays date (converted appropriately) then
                // click the 'Next' button.
                appModel.AdvantageOnlineBankingPage.AmountEditField.SetValue(strAmount);
                appModel.AdvantageOnlineBankingPage.TransferDateEditField.SetValue(DateTime.Now.ToString("MM/dd/yyyy"));
                appModel.AdvantageOnlineBankingPage.NextButton.Click();

                // Click Ok to complete the transfer

                appModel.AdvantageOnlineBankingPage.OKButton.Click();

                // Validate the message that the money has been transfer.  Note, we have to find the amount via
                // a regular expression (See Application Model) as the inner text contains the ACTUAL account.
                
                if (appModel.AdvantageOnlineBankingPage.ReferenceNumber.Exists())//appModel.AmountTransfered.InnerText.Contains(strAmount))
                {
                    Reporter.ReportEvent("MoneyTransfer", "Money Transfered Successfully. "+appModel.AdvantageOnlineBankingPage.ReferenceNumber.InnerText, HP.LFT.Report.Status.Passed, appModel.AdvantageOnlineBankingPage.GetSnapshot());

                }
                else
                {
                    Reporter.ReportEvent("MoneyTransfer", "Money Not Transfered", HP.LFT.Report.Status.Failed, appModel.AdvantageOnlineBankingPage.GetSnapshot());
                    Assert.Fail();
                }
            }
            #endregion
        }
        [TearDown]
        public void TearDown()
        {
            // Clean up after each test
            browser.Close();
        }

        //Xpath for the following would be:
        // Flowers - //TABLE[@id="dlProducts"]/TBODY[1]/TR[2]/TD[1]/DIV[1]/DIV[3]/DIV[1]/INPUT[1]
        // Other   - //TABLE[@id="dlProducts"]/TBODY[1]/TR[2]/TD[2]/DIV[1]/DIV[3]/DIV[1]/INPUT[1]
        //NOTE: with LeanFT, if the product shows up in a differnt order LeanFT will still
        //      find the correct product.  You can not easily guarantee that with using Xpath
        //      ALSO there are ways to walk through this in Selenium without needing to know Xpath
        String[] loopStr = new String [] {"Flowers - Rose Quantity must range from 1 to 10 ",
                                      "Other - Buildings Quantity must range from 1 to 10 ",
                                      "Sports - Yacht Quantity must range from 1 to 10 "};
        [Test, TestCaseSource("loopStr")]
        public void OrderChecks(string checkName)
        {
            login();

            #region access products based on objects and not Xpath
            // Click on Order Checkbooks link
            browser.Describe<ILink>(new LinkDescription
		    {
			    TagName = @"A",
			    InnerText = @"Order Checkbooks "
		    }).Click();

            // Select the desired checks
            browser.Describe<IFrame>(new FrameDescription()).Describe<IWebElement>(new WebElementDescription
            {
                TagName = @"TD",
                InnerText = checkName
            }).Describe<ICheckBox>(new CheckBoxDescription
                        {
                            Type = @"checkbox",
                            TagName = @"INPUT"
                        }).Click();
            #endregion

            //Add the sleep if you want to see it being checked
            //Thread.Sleep(2000);
        }

        [TestFixtureTearDown]
        public void TestFixtureTearDown()
        {
            // Clean up once per fixture
        }
    }
}
