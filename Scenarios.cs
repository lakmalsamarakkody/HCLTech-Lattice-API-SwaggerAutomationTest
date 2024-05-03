using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using OpenQA.Selenium;
using SwaggerWebAPI.Libs;
using SwaggerWebAPI.Tests;
using System.Collections.Generic;
using System.Threading;

namespace SwaggerWebAPI
{
    public class Scenarios
    {
        private ExtentReports ExtReport;
        private ExcelAPI InputExcelApi, ValidationExcelApi, OutputExcelApi;
        private Dictionary<string, ExcelAPI> ExcelApiList;
        private IWebDriver Driver;

        [OneTimeSetUp]
        public void OneTimeSetup()
        {
            //Import Settings
            Settings settings = new Settings();
            DBConnection.SetConnection(settings.GetDataSource(),settings.GetInitialCatalog(),settings.GetUsername(),settings.GetPassword());

            //Initialise Driver
            Driver driver = new Driver(settings.GetDriverType(), settings.GetDriverVersion());
            Driver = driver.GetDriver();
            Driver.Manage().Window.Maximize();
            Driver.Navigate().GoToUrl(settings.GetDriverHomeUrl());

            //Initiating Input Excel
            string InputExcelPath = settings.GetInputExcelPath();
            InputExcelApi = new ExcelAPI(InputExcelPath);

            //Initiating Validation Excel
            string ValidationExcelPath = settings.GetValidationExcelPath();
            ValidationExcelApi = new ExcelAPI(ValidationExcelPath);

            //Initiating Output Excel
            string OutputExcelPath = settings.GetOutputExcelPath();
            OutputExcelApi = new ExcelAPI(OutputExcelPath);

            //ExcelAPIList
            ExcelApiList = new Dictionary<string, ExcelAPI>();
            ExcelApiList.Add("InputExcelApi", InputExcelApi);
            ExcelApiList.Add("ValidationExcelApi", ValidationExcelApi);
            ExcelApiList.Add("OutputExcelApi", OutputExcelApi);

            //Initiating Report
            string ExtentReportPath = settings.GetExtentReportPath();
            ExtReport = new ExtentReports();
            var htmlReporter = new ExtentHtmlReporter(ExtentReportPath);
            ExtReport.AttachReporter(htmlReporter);
        }

        [SetUp]
        public void Setup()
        {
            //nothing here
        }

      
        [Test]
        public void TS_001_CcyPairs_GetCcyPairs()
        {
            CcyPairs ccyPair = new CcyPairs(Driver, ExcelApiList, ExtReport);
            ccyPair.GetCcyPairs();
            Thread.Sleep(3000);

        }

        [Test]
        public void TS_002_Connection_Getconnection()
        {
            Getconnection connection = new Getconnection(Driver, ExcelApiList, ExtReport);
            connection.GetConnection();
            Thread.Sleep(3000);
        }
        
        [Test]
        public void TS_003_Coversion_GetConversionRate()
        {
            ConversionRate conv = new ConversionRate(Driver, ExcelApiList, ExtReport);
            conv.GetConversionRate();
            Thread.Sleep(2000);
        }
        
        [Test]
        public void TS_004_GetCounterparty()
        {
             Counterparty c = new Counterparty(Driver, ExcelApiList, ExtReport);
            c.GetCounterparty();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_005_GetClients()
        {
            GetClients c = new GetClients(Driver, ExcelApiList, ExtReport);
            c.GetTradeBroker();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_006_ExcutingBrokers()
        {
            ExecutingBrokers c = new ExecutingBrokers(Driver, ExcelApiList, ExtReport);
            c.GetExcBroker();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_007_BrokerCodes()
        {
            BrokerCodes c = new BrokerCodes(Driver, ExcelApiList, ExtReport);
            c.GetBrokerCodes();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_008_GetCutoff()
        {
            Cutoff c = new Cutoff(Driver, ExcelApiList, ExtReport);
            c.GetCutoff();
        }
        
        [Test]
        public void TS_009_GetCanDealStatus()
        {
            CanDealStatus c = new CanDealStatus(Driver, ExcelApiList, ExtReport);
            c.GetDealstatus();
        }
        
        [Test]
        public void TS_010_GetPrimeBrokerIds()
        {
            PrimeBrokerIds c = new PrimeBrokerIds(Driver, ExcelApiList, ExtReport);
            c.GetPrimeBrokerIds();
        }
        
        [Test]
        public void TS_011_GetFixingSource()
        {
            FixingSource c = new FixingSource(Driver, ExcelApiList, ExtReport);
            c.GetFixSource();
            Thread.Sleep(1000);
        }

        [Test]
        public void TS_012_GetLocalRecieved()
        {
            LocalRecieved c = new LocalRecieved(Driver, ExcelApiList, ExtReport);
            c.GetLocalRecieved();
        }
        
        [Test]
        public void TS_013_GetDealCreatedandLastUpdate()
        {
            DealCreate c = new DealCreate(Driver, ExcelApiList, ExtReport);
            c.GetDealCreate();
            Thread.Sleep(1000);
        }

        [Test]
        public void TS_014_GetPrimeBrokerCounterparty()
        {
            PrimeBrokerCounterparty c = new PrimeBrokerCounterparty(Driver, ExcelApiList, ExtReport);
            c.GetPrimeBrokerCounter();
        }
        
        [Test]
        public void TS_015_GetLastchangeIDofCounterparty()
        {
            LastchangeIDofCounterparty c = new LastchangeIDofCounterparty(Driver, ExcelApiList, ExtReport);
            c.GetLastchange();
        }
        
        [Test]
        public void TS_016_GetMaxGCDId()
        {
            MaxGCD c = new MaxGCD(Driver, ExcelApiList, ExtReport);
            c.GetMaxGCD();
        }
        
        [Test]
        public void TS_017_GetAddedDatefromAuditGCDUpdation()
        {
            AddedDatefromAuditGCDUpdation c = new AddedDatefromAuditGCDUpdation(Driver, ExcelApiList, ExtReport);
            c.GetAuditGCD();
            Thread.Sleep(1000);
        }

        [Test]
        public void TS_018_GetAddedByFromAuditGCDUpdation()
        {
            AddedByFromAuditGCDUpdation c = new AddedByFromAuditGCDUpdation(Driver, ExcelApiList, ExtReport);
            c.GetAddeByFrom();
            Thread.Sleep(1000);
        }

        [Test]
        public void TS_019_GetCounterFromAuditpartyGCDUpdation()
        {
            CounterFromAuditGCDUpdation c = new CounterFromAuditGCDUpdation(Driver, ExcelApiList, ExtReport);
            c.GetCounterFromAudit();
        }
        
        [Test]
        public void TS_020_GetEditGCDIdRoleUser()
        {
            EditGCDIdRole c = new EditGCDIdRole(Driver, ExcelApiList, ExtReport);
            c.GetEditGCDIdRole();
            Thread.Sleep(3000);
        }

        [Test]
        public void TS_021_GetRoleDetails()
        {
            GetRoleDetails c = new GetRoleDetails(Driver, ExcelApiList, ExtReport);
            c.RoleDetails();
            Thread.Sleep(2000);

        }

        [Test]
        public void TS_022_GetDuplicateCodeforAdd()
        {
            DuplicateCode c = new DuplicateCode(Driver, ExcelApiList, ExtReport);
            c.GetDuplicateCode();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_023_GetGetGCDid()
        {
            GetGCDid c = new GetGCDid(Driver, ExcelApiList, ExtReport);
            c.GetGCDID();
        }
        
        [Test]
        public void TS_024_GetBrokerCodeResults()
        {
            GetBrokerCodeResult c = new GetBrokerCodeResult(Driver, ExcelApiList, ExtReport);
            c.Get_BrokerCodeResult();
        }
        
        [Test]
        public void TS_025_GetBrokerCodeforAdmin()
        {
            BrokerCodeforAdmin c = new BrokerCodeforAdmin(Driver, ExcelApiList, ExtReport);
            c.GetBrokerCodeforAdmin();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_026_GetConversionRateResults()
        {
            ConversionRateResult c = new ConversionRateResult(Driver, ExcelApiList, ExtReport);
            c.Get_ConversionRateResult();
        }
        
        [Test]
        public void TS_027_GetDealStrategy()
        {
            DealStrategy c = new DealStrategy(Driver, ExcelApiList, ExtReport);
            c.GetDealStrategy();
        }
        
        [Test]
        public void TS_028_GetBrokerCodeIdResult()
        {
            BrokerCodeIdResult c = new BrokerCodeIdResult(Driver, ExcelApiList, ExtReport);
            c.GetBrokerCodeIdResult();
        }
        
        [Test]
        public void TS_029_GetConversionRateResult()
        {
            GetConversionRateResult c = new GetConversionRateResult(Driver, ExcelApiList, ExtReport);
            c.ConversionRateResult();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_030_GetBrokerCodeResults()
        {
            GetBrokerCodeResults c = new GetBrokerCodeResults(Driver, ExcelApiList, ExtReport);
            c.Get_BrokerCodeResult();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_031_GetTraderBrokerCodeResult()
        {
            TraderBrokerCodeResult c = new TraderBrokerCodeResult(Driver, ExcelApiList, ExtReport);
            c.GetTraderBrokerCodeIdResult();
            Thread.Sleep(2000);
        }
        
        [Test]
        public void TS_032_GetCounterparty()
        {
            GetCounterparty c = new GetCounterparty(Driver, ExcelApiList, ExtReport);
            c.Get_Counterparty();
        }
        
        [Test]
        public void TS_033_GetLattiseCounterTradeDetails()
        {
            LattiseCounterTradeDetails c = new LattiseCounterTradeDetails(Driver, ExcelApiList, ExtReport);
            c.GetLattiseCounterTradeDetails();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_034_GetLattiseCounterpartyTradeDetailsonTraderCode()
        {
            LattiseCounterpartyTradeDetailsonTraderCode c = new LattiseCounterpartyTradeDetailsonTraderCode(Driver, ExcelApiList, ExtReport);
            c.GetCounterpartyTradeDetailsonTraderCode();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_035_GetFixingSourceDetails()
        {
            FixingSourceDetails c = new FixingSourceDetails(Driver, ExcelApiList, ExtReport);
            c.GetFixingSourceDetails();
            Thread.Sleep(3000);
        }

        [Test]
        public void TS_036_GetLattiseTraderDetails()
        {
            GetLattiseTraderDetails c = new GetLattiseTraderDetails(Driver, ExcelApiList, ExtReport);
            c.GetTradeDetails();
            Thread.Sleep(3000);
        }
        
        [Test]
        public void TS_037_GetDuplicateFixingSource()
        {
            DuplicateFixingSource c = new DuplicateFixingSource(Driver, ExcelApiList, ExtReport);
            c.Get_DuplicateFixingSource();
            Thread.Sleep(3000);

        }

        [Test]
        public void TS_038_GetDupicateDomainLoginforAdd()
        {
            DupicateDomainLoginforAdd c = new DupicateDomainLoginforAdd(Driver, ExcelApiList, ExtReport);
            c.Get_DupicateDomainLoginforAdd();
            Thread.Sleep(3000);

        }

        [Test]
        public void TS_039_GetAgentContactName()
        {
            AgentContactName c = new AgentContactName(Driver, ExcelApiList, ExtReport);
            c.GetAgentContactName();
            Thread.Sleep(3000);
        }

        [Test]
        public void TS_040_GetFwdScaling()
        {
            FwdScaling c = new FwdScaling(Driver, ExcelApiList, ExtReport);
            c.GetFwdScaling();
            Thread.Sleep(3000);
        }


        [Test]
        public void TS_041_GetGCDIDForEnteringOrUpdatingUserContactRef()
        {
            GCDIDForEnteringOrUpdatingUserContactRef c = new GCDIDForEnteringOrUpdatingUserContactRef(Driver, ExcelApiList, ExtReport);
            c.GetUpdatingUserContactRef();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_042_GetCutoffData(){ 

            CutOffData c = new CutOffData(Driver, ExcelApiList, ExtReport);
            c.GetCutOffData();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_043_GetUSPerson(){

            USPerson c = new USPerson(Driver, ExcelApiList, ExtReport);
            c.GetUSPersonValue();
            Thread.Sleep(3000);
        }

         [Test]
         public void TS_044_GetCommenetforTrade(){

             CommenetforTrade c = new CommenetforTrade(Driver, ExcelApiList, ExtReport);
             c.GetCommenetforTrade();
            Thread.Sleep(2000);
        }

        [Test]
         public void TS_045_GetCountofRecords(){

             CountofRecords c = new CountofRecords(Driver, ExcelApiList, ExtReport);
             c.GetCountofRecords();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_046_GetDupicateDomainLoginforUpdate(){

            DupicateDomainLoginforUpdate c = new DupicateDomainLoginforUpdate(Driver, ExcelApiList, ExtReport);
            c.Get_DupicateDomainLoginforUpdate();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_047_GetCurrencyPairResult(){

            CurrencyPairResult c = new CurrencyPairResult(Driver, ExcelApiList, ExtReport);
            c.GetCurrencyPairResult();
            Thread.Sleep(3000);
        }

         [Test]
         public void TS_048_GetAllPrimeBorkers(){

             GetAllPrimeBorkers c = new GetAllPrimeBorkers(Driver, ExcelApiList, ExtReport);
             c.Get_AllPrimeBorkers();
            Thread.Sleep(2000);
        }

        [Test]
         public void TS_049_GetAllCcypairs(){

             GetAllCcypairs c = new GetAllCcypairs(Driver, ExcelApiList, ExtReport);
             c.Get_AllCcypairs();
            Thread.Sleep(2000);
        }

        [Test]
         public void TS_050_GetAllDirections()
         {
             GetAllDirections c = new GetAllDirections(Driver, ExcelApiList, ExtReport);
             c.Get_AllDirections();
            Thread.Sleep(2000);
        }

        [Test]
         public void TS_051_GetAllCallPuts(){

             GetAllCallPuts c = new GetAllCallPuts(Driver, ExcelApiList, ExtReport);
             c.Get_AllCallPuts();
            Thread.Sleep(2000);
        }


        [Test]
        public void TS_052_GetMaxIdOfNewlyaddedCCyFixing()
        {

            MaxIdOfNewlyaddedCCyFixing c = new MaxIdOfNewlyaddedCCyFixing(Driver, ExcelApiList, ExtReport);
            c.GetMaxIdOfNewlyaddedCCyFixing();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_053_GetCurrenciesCombo()
        {

            CurrenciesCombo c = new CurrenciesCombo(Driver, ExcelApiList, ExtReport);
            c.Get_CurrenciesCombo();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_054_GetCurrencyCombination()
        {
            CurrencyCombination c = new CurrencyCombination(Driver, ExcelApiList, ExtReport);
            c.GetCurrencyCombination();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_055_GetMessageInTrade()
        {
            MessageInTrade c = new MessageInTrade(Driver, ExcelApiList, ExtReport);
            c.GetMessageInTrade();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_056_GetPrimeBrokers()
        {
            GetPrimeBrokers c = new GetPrimeBrokers(Driver, ExcelApiList, ExtReport);
            c.Get_PrimeBrokers();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_057_GetCounterparties()
        {
            GetCounterparties c = new GetCounterparties(Driver, ExcelApiList, ExtReport);
            c.Get_Counterparties();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_058_GetNameofExecutingBroker()
        {
            NameofExecutingBroker c = new NameofExecutingBroker(Driver, ExcelApiList, ExtReport);
            c.GetNameofExecutingBroker();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_059_GetGCDIDofExecutingBroker()
        {
            GCDIDofExecutingBroker c = new GCDIDofExecutingBroker(Driver, ExcelApiList, ExtReport);
            c.GetGCDIDofExecutingBroker();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_060_GetExecutingBrokerName()
        {
            ExecutingBrokerName c = new ExecutingBrokerName(Driver, ExcelApiList, ExtReport);
            c.GetExecutingBrokerName();
            Thread.Sleep(3000);
        }

        [Test]
        public void TS_061_GetIsActiveBroker()
        {
            IsActiveBroker c = new IsActiveBroker(Driver, ExcelApiList, ExtReport);
            c.GetIsActiveBroker();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_062_GetPostingNameAndCode()
        {
            PostingNameAndCode c = new PostingNameAndCode(Driver, ExcelApiList, ExtReport);
            c.Get_PostingNameAndCode();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_063_GetUSPersonLEIvalues()
        {
            USPersonLEIvalues c = new USPersonLEIvalues(Driver, ExcelApiList, ExtReport);
            c.GetUSPersonLEIvalues();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_064_GetExistingCurrencies()
        {
            ExistingCurrencies c = new ExistingCurrencies(Driver, ExcelApiList, ExtReport);
            c.GetExistingCurrencies();
            Thread.Sleep(3000);
        }

        [Test]
        public void TS_065_GetLastchangedIDofCCyPair()
        {
            LastchangedIDofCCyPair c = new LastchangedIDofCCyPair(Driver, ExcelApiList, ExtReport);
            c.GetLastchangedIDofCCyPair();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_066_GetCutOffValues()
        {
            GetCutOffValues c = new GetCutOffValues(Driver, ExcelApiList, ExtReport);
            c.Get_CutOffValues();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_067_GetLastchangedIDofBroker()
        {
            LastchangedIDofBroker c = new LastchangedIDofBroker(Driver, ExcelApiList, ExtReport);
            c.Get_LastchangedIDofBroker();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_068_GetCcypairs()
        {
            GetCcypairs c = new GetCcypairs(Driver, ExcelApiList, ExtReport);
            c.Get_Ccypairs();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_069_GetDeal()
        {
            GetDeal c = new GetDeal(Driver, ExcelApiList, ExtReport);
            c.Get_Deal();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_070_GetDealStatus()
        {
            GetDealStatus c = new GetDealStatus(Driver, ExcelApiList, ExtReport);
            c.Deal_Status();
            Thread.Sleep(2000);
        }
        
        [Test]
        public void TS_071_GetLatestDealChange()
        {
            GetLatestDealChange c = new GetLatestDealChange(Driver, ExcelApiList, ExtReport);
            c.LatestDeal_Change();
            Thread.Sleep(2000);
        }
        
        [Test]
        public void TS_072_GetLatestDealChangeId()
        {
            GetLatestDealChangeId c = new GetLatestDealChangeId(Driver, ExcelApiList, ExtReport);
            c.LatestDeal_ChangeId();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_073_GetDealChanges()
        {
            DealChanges c = new DealChanges(Driver, ExcelApiList, ExtReport);
            c.GetDealChanges();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_074_GetValidateUser()
        {
            ValidateUser c = new ValidateUser(Driver, ExcelApiList, ExtReport);
            c.GetValidateUser();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_075_GetPrimeBrokers()
        {
            PrimeBrokers c = new PrimeBrokers(Driver, ExcelApiList, ExtReport);
            c.GetPrimeBrokers();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_076_GetLatestStatusChange()
        {
            LatestStatusChange c = new LatestStatusChange(Driver, ExcelApiList, ExtReport);
            c.GetLatestStatusChange();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_077_GetSettings()
        {
            GetSetting c = new GetSetting(Driver, ExcelApiList, ExtReport);
            c.GetSystemSetting();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_078_GetCitiPBID()
        {
            CitiPBID c = new CitiPBID(Driver, ExcelApiList, ExtReport);
            c.GetCitiPBID();
            Thread.Sleep(2000);
        }
        
        [Test]
        public void TS_079_GetRangeOfTrades()
        {
            RangeOfTrades c = new RangeOfTrades(Driver, ExcelApiList, ExtReport);
            c.GetRangeOfTrades();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_080_GetExpireTrades()
        {
            ExpireTrades c = new ExpireTrades(Driver, ExcelApiList, ExtReport);
            c.GetExpireTrades();
        }
        
        [Test]
        public void TS_081_GetTradeChanges()
        {
            TradeChanges c = new TradeChanges(Driver, ExcelApiList, ExtReport);
            c.GetTradeChanges();
            Thread.Sleep(2000);
        }
        
        [Test]
        public void TS_082_GetTradesCount()
        {
            TradesCount c = new TradesCount(Driver, ExcelApiList, ExtReport);
            c.GetTradesCount();
            Thread.Sleep(2000);
        }

        [Test]
        public void TS_083_GetLastTradeChangedID()
        {
            LastTradeChangedID c = new LastTradeChangedID(Driver, ExcelApiList, ExtReport);
            c.GetLastTradeChangedID();
        }

        [TearDown]
        public void TearDown()
        {
            //nothing here
        }

        [OneTimeTearDown]

        public void OneTimeTearDown()
        {
            Driver.Quit();
            InputExcelApi.CloseExcel();
            ValidationExcelApi.CloseExcel();
            OutputExcelApi.CloseExcel();
            ExtReport.Flush();
        }


        
    }
}