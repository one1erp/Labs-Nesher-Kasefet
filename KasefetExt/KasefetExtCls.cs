using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ADODB;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using One1.Controls;
using System.Diagnostics;

namespace KasefetExt
{

    [ComVisible(true)]
    [ProgId("KasefetExt.KasefetExt_Cls")]
    public class KasefetExt_Cls : IWorkflowExtension //, IEntityExtension
    {

        INautilusServiceProvider sp;

        public void Execute(ref LSExtensionParameters Parameters)
        {

            try
            {
                Logger.WriteLogFile("start", false);
                #region param
                //used for debug
                //Debugger.Launch();
                sp = Parameters["SERVICE_PROVIDER"];

                //recrdset declaration send an error when going to fields
                //we use the dynamic object as var , but use the recordset declaration in intellisense
                var rs = Parameters["RECORDS"];
                //Recordset rs = Parameters["RECORDS"];

                rs.MoveLast();

                var sdgId = rs.Fields["SDG_ID"].Value;

                long id = long.Parse(sdgId.ToString());
                //CustomMessageBox.Show(id.ToString());

                ////////////יוצר קונקשן//////////////////////////
                var ntlCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlCon);
                /////////////////////////////
                
                var dal = new DataLayer();
                dal.Connect();
                #endregion
                //   var wnId = Parameters["WORKFLOW_NODE_ID"];




                Sdg sdg = dal.GetSdgById(id);
                if (sdg != null && sdg.LabInfo != null && sdg.LabInfo.Name != "Water")
                {
                   CustomMessageBox.Show("לא התקימו התנאים ליצירת המסמך");
                    return;
                }

                var b = new CreateFile(dal, sdg);
                var success = b.Success;
                string s;
                if (success)
                {
                  CustomMessageBox.Show("המסמך נוצר.");

                }

                else
                {
                    CustomMessageBox.Show("נכשלה יצירת המסמך לכספת.");

                }
            }


            catch (Exception e)
            {
                Logger.WriteLogFile("error in Execute " + e.Message, false);

                Logger.WriteLogFile(e);
                CustomMessageBox.Show("נכשלה יצירת המסמך לכספת.");

            }


        }




        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            return ExecuteExtension.exEnabled;
        }
    }
}
