using ADODB;
using LSEXT;
using LSSERVICEPROVIDERLib;
using Oracle.ManagedDataAccess.Client;
using Patholab_Common;
using Patholab_DAL_V1;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;


namespace Print_VEGA_Cassete
{
    [ComVisible(true)]
    [ProgId("Print_VEGA_Cassete.PrintVEGACasseteCls")]
    public class PrintCasseteCls : IWorkflowExtension
    {
        INautilusServiceProvider sp;
        private DataLayer dal;
        OracleConnection connection = null;

        public bool DEBUG;
        public void Execute(ref LSExtensionParameters Parameters)
        {
            string sampleDigit = "";
            string AliqDigit="";//
            long tableID = 0;
            string ext = "";
            string aliqPathoName="";
            string aliqNautilusName="";
            string pcol="";
            string priority = "";
            try
            {
                if (DEBUG)
                {
                    sampleDigit = "1";
                    AliqDigit = "1";
                    tableID = 0;
                    ext = ".1.1";
                    aliqPathoName = "P_    1-19";
                    aliqNautilusName = "B000001/19.1.1";
                    pcol = "1";
                    priority = "0";
                }
                #region param
                if (!DEBUG)
                {
                    string tableName = Parameters["TABLE_NAME"];



                    int i = 1;
                    sp = Parameters["SERVICE_PROVIDER"];

                    Recordset rs = Parameters["RECORDS"];

                    rs.MoveLast();


                    try
                    {
                        long.TryParse(rs.Fields["ALIQUOT_ID"].Value.ToString(), out tableID);
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteLogFile(ex);
                        MessageBox.Show("This program works on ALIQUOT only.");
                        return;
                    }
                }
                #endregion



                #region Data
                if (!DEBUG)
                {
                    var ntlCon = Utils.GetNtlsCon(sp);
                    Utils.CreateConstring(ntlCon);

                    
                    if (ntlCon != null)
                    {
                        connection = GetConnection(ntlCon);
                        

                        dal = new DataLayer();
                        dal.Connect(ntlCon);
                    }
                    else 
                    {
                        throw new Exception("can't get nautilus connection.");
                    }


                    var aliq = (from item in dal.GetAll<ALIQUOT>()
                                where item.ALIQUOT_ID == tableID
                                select
                                new
                                {
                                    PRIORITY = item.SAMPLE.SDG.SDG_USER.U_PRIORITY,
                                    NautilusName = item.NAME,
                                    PatholabName = item.SAMPLE.SDG.SDG_USER.U_PATHOLAB_NUMBER,
                                    printerCol = item.ALIQUOT_USER.U_PRINTER_COL
                                }).SingleOrDefault();

                    if (aliq == null)
                    {
                        MessageBox.Show("Can't find the aliquot for the id");
                        return;
                    }

                    //    char limitSmpAliq = ',';

                    int index = 10;
                    aliqNautilusName = aliq.NautilusName;
                    //      aliqNautilusName = aliqNautilusName.Replace('/', '-');
                    //    aliqNautilusName = aliqNautilusName.Replace('.', limitSmpAliq);
                    aliqPathoName = CreateShortName(aliq.NautilusName, aliq.PatholabName);
                    ext = aliq.NautilusName.Substring(index, aliq.NautilusName.Length - index);
                    var dig = ext.Split('.');
                    sampleDigit = dig[1];
                    AliqDigit = dig[2];

                    pcol = aliq.printerCol;
                    if (false)//aliq.PRIORITY.HasValue&&aliq.PRIORITY.Value=999f) ask ziv
                    {
                        priority = "1";
                    }

                }
                string[] arr = new string[5];
                arr[0] = string.Format("DT;{0};", pcol);
                arr[1] = "S1;#1#;#2#;";
                arr[2] = string.Format("S2;#1#{0};#2#.{1};#3#.{2};#4#;#5#{3};", aliqPathoName, sampleDigit, AliqDigit, aliqNautilusName);
                arr[3] = "S3;#1#;#2#;";
                arr[4] = string.Format("PR;{0};", priority);


                long wnid = Parameters["WORKFLOW_NODE_ID"]; // wnid = Workflow Node ID
                string path = getPrinterPath(dal, wnid); //dal.GetPhraseByName("System Parameters").PhraseEntriesDictonary["Vega Path"];
                string dt = DateTime.Now.ToString("yyyyMMddHHmmssFFF");
                if (!Directory.Exists(path))
                {
                    MessageBox.Show(path + "Doesn't Exists");
                    return;
                }

                File.AppendAllLines(Path.Combine(path, dt + ".txt"), arr);


                #endregion



            }
            catch (Exception ex)
            {
                MessageBox.Show("נכשלה הדפסת קסטה.");
                Logger.WriteLogFile(ex);
            }
            finally
            {
                if (dal != null) dal.Close();
            }
        }


        private string getPrinterPath(DataLayer dal, long wnid)
        {
            OracleCommand cmd = null;
            var sql = string.Format("select parent_id from lims_sys.workflow_node where workflow_node_id={0}", wnid);

            cmd = new OracleCommand(sql, connection);
            var parentNodeId = cmd.ExecuteScalar();

            string printerEventName = string.Empty;
            if (parentNodeId != null)
            {
                sql = string.Format("select LONG_NAME from lims_sys.workflow_node where workflow_node_id={0}", parentNodeId);

                cmd.CommandText = sql;
                printerEventName = Convert.ToString(cmd.ExecuteScalar());

                if (!string.IsNullOrEmpty(printerEventName))
                {
                    PHRASE_ENTRY entry = dal.FindBy<PHRASE_ENTRY>(pe => pe.PHRASE_HEADER.NAME.Equals("Vega Printer") && pe.PHRASE_INFO.Equals(printerEventName, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();

                    if (entry != null) return entry.PHRASE_DESCRIPTION;
                    else MessageBox.Show(string.Format("Phrase: Vega Printer.{0}Can't find phrase entry information that match{0}the event name: {1}", Environment.NewLine, printerEventName));
                }
            }

            return null;
        }

        public OracleConnection GetConnection(INautilusDBConnection ntlsCon)
        {
            OracleConnection connection = null;
            if (ntlsCon != null)
            {
                //initialize variables
                string rolecommand;
                //try catch block
                try
                {

                    string connectionString;
                    string server = ntlsCon.GetServerDetails();
                    string user = ntlsCon.GetUsername();
                    string password = ntlsCon.GetPassword();

                    connectionString =
                        string.Format("Data Source={0};User ID={1};Password={2};", server, user, password);

                    if (string.IsNullOrEmpty(user))
                    {
                        connectionString = "User Id=/;Data Source=" + server + ";Connection Timeout=60;";
                    }

                    //create connection
                    connection = new OracleConnection(connectionString);

                    //open the connection
                    connection.Open();

                    //get lims user password
                    string limsUserPassword = ntlsCon.GetLimsUserPwd();

                    //set role lims user
                    if (limsUserPassword == "")
                    {
                        //lims_user is not password protected 
                        rolecommand = "set role lims_user";
                    }
                    else
                    {
                        //lims_user is password protected
                        rolecommand = "set role lims_user identified by " + limsUserPassword;
                    }

                    //set the oracle user for this connection
                    OracleCommand command = new OracleCommand(rolecommand, connection);

                    //try/catch block
                    try
                    {
                        //execute the command
                        command.ExecuteNonQuery();
                    }
                    catch (Exception f)
                    {
                        //throw the exeption
                    }

                    //get session id
                    double sessionId = ntlsCon.GetSessionId();

                    //connect to the same session 
                    string sSql = string.Format("call lims.lims_env.connect_same_session({0})", sessionId);

                    //Build the command 
                    command = new OracleCommand(sSql, connection);

                    //execute the command
                    command.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    //throw the exeption
                }
            }
            return connection;
        }

        private string CreateShortName(string aliqoutName, string sdgPathoName)
        {


            StringBuilder sb = new StringBuilder();
            sb.Append(sdgPathoName[0]);
            sb.Append(sdgPathoName[1]);

            bool StopRemove = true;
            for (int i = 2; i < sdgPathoName.Length; i++)
            {
                if (StopRemove)
                {


                    if (sdgPathoName[i] == '0')
                    {
                        sb.Append(' ');
                    }
                    else
                    {
                        StopRemove = false;
                        sb.Append(sdgPathoName[i]);
                    }
                }
                else
                {
                    sb.Append(sdgPathoName[i]);
                }
            }


            var aliqPathoName = sb.ToString();// +ext;
            aliqPathoName = aliqPathoName.Replace('/', '-');
            aliqPathoName = aliqPathoName.Replace('_', ' ');


            return aliqPathoName;
        }





    }
}
