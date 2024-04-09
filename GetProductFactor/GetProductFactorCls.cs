using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;

namespace GetProductFactor
{

    [ComVisible(true)]
    [ProgId("GetProductFactor.GetProductFactorCls")]
    public class GetProductFactorCls : IWorkflowExtension
    {

        INautilusServiceProvider sp;
        private IDataLayer dal;
        private Sample sample;

        public void Execute(ref LSExtensionParameters Parameters)
        {

            try
            {


                string tableName = Parameters["TABLE_NAME"];
                //   string  tableName = tableName.ToString();
                sp = Parameters["SERVICE_PROVIDER"];

                var rs = Parameters["RECORDS"];

                //Get entity id
                var entityId = rs.Fields[tableName + "_ID"].Value;

                var ntlsCon = Utils.GetNtlsCon(sp);

                Utils.CreateConstring(ntlsCon);
                dal = new DataLayer();
                dal.Connect();

                long id = long.Parse(entityId.ToString());

                //Get sample of Product
                sample = GetSpecifiedSample(tableName, id);
                Product product;
                if (sample == null)
                {
                    writeNullToLog("sample");
                    return;
                }

                else
                {
                    product = sample.Product;

                    if (product == null)
                    {
                        writeNullToLog("product");
                        return;
                    }
                    if (sample.FactorNameIn == null)
                    {
                        writeNullToLog("Factor Name In");
                        return;
                    }
                    var calc = dal.GetCalculationFactor("PRODUCT", product.ProductId.ToString(), sample.FactorNameIn);
                    if (calc == null)
                    {
                        writeNullToLog("Calculation Factor");
                        return;
                    }
                    switch (calc.CalculationFactorType)
                    {
                        case "B":
                            sample.FactorValueOut = calc.FactorTExtValue;
                            break;
                        case "T":
                            sample.FactorValueOut = calc.FactorTExtValue;
                            break;
                        case "N":
                            sample.FactorValueOut = calc.FactorValue.ToString();
                            break;

                    }
                }

            }
            catch (Exception ex)
            {

                Logger.WriteLogFile(ex);
            }
            finally
            {
                if (dal != null) { dal.SaveChanges(); dal.Close(); }
            }

        }

        private Sample GetSpecifiedSample(string tableName, long entityId)
        {
            switch (tableName)
            {
                case "SAMPLE":
                    return dal.GetSampleByKey(entityId);
                case "ALIQOUT":
                    var aliqout = dal.GetAliquotById(entityId);
                    if (aliqout != null) return aliqout.Sample;
                    break;
                case "TEST":
                    var test = dal.GetTestById(entityId);
                    if (test != null && test.Aliquot != null)
                        return test.Aliquot.Sample;
                    break;
                case "RESULT":
                    var result = dal.GetResultById(entityId);
                    if (result != null)
                        if (result.Test != null && result.Test.Aliquot != null)
                            return result.Test.Aliquot.Sample;
                    break;
                default:
                    return null;
            }
            return null;
        }
        public void writeNullToLog(string item)
        {
            if (sample != null)
            {
                sample.FactorValueOut = "does not exists";
            }
            Logger.WriteLogFile(item + " is null", false);
        }
    }
}
