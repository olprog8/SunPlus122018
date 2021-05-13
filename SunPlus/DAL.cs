using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Collections; //Подключаем для работы с ArrayList

using System.Data.SqlClient; //Подключаем для работы с ADO.NET
using System.Data.Common; //Подключаем для работы с ADO.NET

using System.Threading; //Подключаем для работы с Потоками

using System.Globalization; //ПАРАМЕТРЫ СТРАНЫ (создаем для замены , на . в полях decimal)

using System.IO; //в данном Пространстве Имен находится StreamWriter

using System.Data.Odbc;


/// <summary>
/// 1111
/// <summary>

namespace SunPlus
{
    class DAL
    {
        //string connectionString;
        readonly static string connectionString = @"Data Source='RUMOWWSX12032';Initial Catalog='DHL_Interfaces';Integrated Security = false; User ID = 'RequestSunUser'; PWD = 'RequestXXXUser@2019'; Persist Security Info = False";
        //string connectionStringConf = @"Data Source='RUMOWWS20030020';Initial Catalog='SUNDB';User ID='SUN';Password='sunpas';Persist Security Info=False";
        string connectionStringConf = @"Data Source='RUMOWWSX12032';Initial Catalog='DHL_Interfaces';Integrated Security=SSPI;Persist Security Info=False";
        //string conStringTradeAO = @"Data Source='RUMOWWSX23084';Initial Catalog='RUTRIDB';Integrated Security=true;User ID='olesnits';Persist Security Info=False";
        string conStringTradeAO = @"Data Source='RUMOWWSX12006';Initial Catalog='RUTRIDB';Integrated Security=false;User ID='trd_user'; PWD='rfghjk_trd';Persist Security Info=False";
        string conStringTradeOOO = @"Data Source='RUMOWWSX12006';Initial Catalog='RUTREDB';Integrated Security=false;User ID='trd_user'; PWD='rfghjk_trd';Persist Security Info=False";

        SqlConnection con;

        string period;
        string sunuser;
        string curol;//текущая роль
        string bunit;//бизнес юнит
        string djmax;//максимум журналов
        string rallocmax;//максимум разаллокирования
        string output;

        DataSet DSJrnal;
        SqlDataAdapter DAJrnal, DAJrnalLad;
        SqlCommandBuilder CBldg;


        
        public DataSet UpdateJrnalLDG()
        {
            CBldg = new SqlCommandBuilder(DAJrnal);
            DAJrnal.Update(DSJrnal, "TabSunJrnalLDG");

            return DSJrnal;

        }

        public DataSet GetSunJrnal(string jrnalNoDB, string jrnalRefDB, int itempNum, string busUnit)
        {

            string queryLayout = "";
            //itempNum = 2;
            switch (busUnit)
            {
                case "A11":
                    {
                        connectionStringConf = @"Data Source='RUMOWWS20030020';Initial Catalog='SS5DB';User ID='SUN';Password='sunpas';Persist Security Info=False";
                        break;
                    }
            }

            switch (itempNum)
            { 
                case 1: 
                { 
                    if (jrnalNoDB == "%")
                    {
                        queryLayout = "select JRNAL_NO, JRNAL_LINE, TREFERENCE, PERIOD, TRANS_DATETIME, JRNAL_SRCE, JRNAL_TYPE, AMOUNT, OTHER_AMT, D_C, ACCNT_CODE, ANAL_T6, DESCRIPTN, " +
                                        "CONV_CODE, CONV_RATE, ROUGH_FLAG, ALLOCATION, ALLOC_IN_PROGRESS, IN_USE_FLAG, ANAL_T0, ANAL_T1, ANAL_T2, ANAL_T3, ANAL_T4, ANAL_T5, ANAL_T7, ANAL_T8, ANAL_T9 " +
                                        "from [SunSystemsData].." + busUnit + "_A_SALFLDG where TREFERENCE = '" + jrnalRefDB + "'";
                    }
                    else if (jrnalRefDB == "%")
                    {
                        queryLayout = "select JRNAL_NO, JRNAL_LINE, TREFERENCE, PERIOD, TRANS_DATETIME, JRNAL_SRCE, JRNAL_TYPE, D_C, AMOUNT, OTHER_AMT, ACCNT_CODE, ANAL_T6, DESCRIPTN, " +
                                        "CONV_CODE, CONV_RATE, ROUGH_FLAG, ALLOCATION, ALLOC_IN_PROGRESS, IN_USE_FLAG, ANAL_T0, ANAL_T1, ANAL_T2, ANAL_T3, ANAL_T4, ANAL_T5, ANAL_T7, ANAL_T8, ANAL_T9 " +
                                        "from [SunSystemsData].." + busUnit + "_A_SALFLDG where JRNAL_NO = " + jrnalNoDB;

                    }
                    else
                        queryLayout = "select JRNAL_NO, JRNAL_LINE, TREFERENCE, PERIOD, TRANS_DATETIME, JRNAL_SRCE, JRNAL_TYPE, D_C, AMOUNT, OTHER_AMT, ACCNT_CODE, ANAL_T6, DESCRIPTN, " +
                                        "CONV_CODE, CONV_RATE, ROUGH_FLAG, ALLOCATION, ALLOC_IN_PROGRESS, IN_USE_FLAG , ANAL_T0, ANAL_T1, ANAL_T2, ANAL_T3, ANAL_T4, ANAL_T5, ANAL_T7, ANAL_T8, ANAL_T9 " +
                                        "from [SunSystemsData].." + busUnit + "_A_SALFLDG where JRNAL_NO = " + jrnalNoDB + " and TREFERENCE = '" + jrnalRefDB + "'";
                    break; 
                }
                case 2:
                {
                    if (jrnalNoDB == "%")
                    {
                        queryLayout = "select JRNAL_NO, JRNAL_LINE, TREFERENCE, PERIOD, TRANS_DATETIME, JRNAL_SRCE, JRNAL_TYPE, D_C, AMOUNT, OTHER_AMT, ACCNT_CODE, ANAL_T6, DESCRIPTN, " +
                    "CONV_CODE, CONV_RATE, ROUGH_FLAG, ALLOCATION, ALLOC_IN_PROGRESS, IN_USE_FLAG, ANAL_T0, ANAL_T1, ANAL_T2, ANAL_T3, ANAL_T4, ANAL_T5, ANAL_T7, ANAL_T8, ANAL_T9 " +
                    "from [SunSystemsData].." + busUnit + "_A_SALFLDG where TREFERENCE = '" + jrnalRefDB + "'";
                    }
                    else if (jrnalRefDB == "%")
                    {
                        queryLayout = "select JRNAL_NO, JRNAL_LINE, TREFERENCE, PERIOD, TRANS_DATETIME, JRNAL_SRCE, JRNAL_TYPE, D_C, AMOUNT, OTHER_AMT, ACCNT_CODE, ANAL_T6, DESCRIPTN, " +
                    "CONV_CODE, CONV_RATE, ROUGH_FLAG, ALLOCATION, ALLOC_IN_PROGRESS, IN_USE_FLAG, ANAL_T0, ANAL_T1, ANAL_T2, ANAL_T3, ANAL_T4, ANAL_T5, ANAL_T7, ANAL_T8, ANAL_T9 " +
                    "from [SunSystemsData].." + busUnit + "_A_SALFLDG where JRNAL_NO = " + jrnalNoDB;

                    }
                    else
                        queryLayout = "select JRNAL_NO, JRNAL_LINE, TREFERENCE, PERIOD, TRANS_DATETIME, JRNAL_SRCE, JRNAL_TYPE, D_C, AMOUNT, OTHER_AMT, ACCNT_CODE, ANAL_T6, DESCRIPTN, " +
                                            "CONV_CODE, CONV_RATE, ROUGH_FLAG, ALLOCATION, ALLOC_IN_PROGRESS, IN_USE_FLAG, ANAL_T0, ANAL_T1, ANAL_T2, ANAL_T3, ANAL_T4, ANAL_T5, ANAL_T7, ANAL_T8, ANAL_T9 " +
                                            "from [SunSystemsData].." + busUnit + "_A_SALFLDG where JRNAL_NO = " + jrnalNoDB + " and TREFERENCE = '" + jrnalRefDB + "'";
                    break;
                }
                case 3:
                {
                    queryLayout =
                    "select a1s.JRNAL_NO, a1s.JRNAL_LINE, a1s.TREFERENCE, a1s.PERIOD, a1s.TRANS_DATETIME, a1s.JRNAL_SRCE, a1s.JRNAL_TYPE, a1s.AMOUNT, a1s.D_C, a1s.ACCNT_CODE, a1s.ANAL_T6, a1s.DESCRIPTN, " +
                    "a2s.GNRL_DESCR_01, a2s.GNRL_DESCR_02 AS GD02_NOMER_FAKTURA, a2s.GNRL_DESCR_03, a2s.GNRL_DESCR_05 AS GD5_OLD_JNL_TYP, " +
                    "a2s.GNRL_3_DATETIME AS GDT3_DATE_FAKTURA, a2s.GNRL_1_DATETIME AS GDT1_DATE_OPLATA, a2s.GNRL_2_DATETIME AS GDT2_DATE_PRIN_UCHET, " +
                    "a2s.GNRL_DESCR_11 AS GD11_KODOPLATY, a2s.GNRL_DESCR_12 AS GD12, a2s.GNRL_DESCR_13 AS GD13, a2s.GNRL_DESCR_14 AS GD14, " +
                    "a2s.GNRL_DESCR_15 AS GD15, a2s.GNRL_DESCR_16 AS GD16, a2s.GNRL_DESCR_17 AS GD17, a2s.GNRL_DESCR_18 AS GD18, a2s.GNRL_DESCR_19 AS GD19 " +
                    "FROM [SunSystemsData].." + busUnit + "_A_SALFLDG  as a1s with (NOLOCK) INNER JOIN [SunSystemsData].." + busUnit + "_A_SALFLDG_LAD  AS a2s with (NOLOCK) " +
                    "ON a1s.JRNAL_NO=a2s.JRNAL_NO AND a1s.JRNAL_LINE=a2s.JRNAL_LINE WHERE a1s.JRNAL_NO = " + jrnalNoDB + " and a1s.TREFERENCE like '" + jrnalRefDB + "'" +
                    "order by a1s.JRNAL_NO, a1s.TREFERENCE, a1s.JRNAL_LINE";

                    break;
                }
                case 4: 
                { 
                    queryLayout = 
                    "select a1s.JRNAL_NO, a1s.TREFERENCE, a1s.JRNAL_LINE, a1s.PERIOD, a1s.TRANS_DATETIME, a1s.JRNAL_SRCE, a1s.JRNAL_TYPE, a1s.AMOUNT, a1s.D_C, a1s.ACCNT_CODE, a1s.ANAL_T6, a1s.DESCRIPTN, " +
                    "a2s.GNRL_DESCR_01, a2s.GNRL_DESCR_02 AS GD02_NOMER_FAKTURA, a2s.GNRL_DESCR_03, a2s.GNRL_DESCR_05 AS GD5_OLD_JNL_TYP, " +
                    "a2s.GNRL_3_DATETIME AS GDT3_DATE_FAKTURA, a2s.GNRL_1_DATETIME AS GDT1_DATE_OPLATA, a2s.GNRL_2_DATETIME AS GDT2_DATE_PRIN_UCHET, " +
                    "a2s.GNRL_DESCR_11 AS GD11_KODOPLATY, a2s.GNRL_DESCR_12 AS GD12, a2s.GNRL_DESCR_13 AS GD13, a2s.GNRL_DESCR_14 AS GD14, " +
                    "a2s.GNRL_DESCR_15 AS GD15, a2s.GNRL_DESCR_16 AS GD16, a2s.GNRL_DESCR_17 AS GD17, a2s.GNRL_DESCR_18 AS GD18, a2s.GNRL_DESCR_19 AS GD19 " +
                    "FROM [SunSystemsData].." + busUnit + "_A_SALFLDG  as a1s with (NOLOCK) INNER JOIN [SunSystemsData].." + busUnit + "_A_SALFLDG_LAD  AS a2s with (NOLOCK) " +
                    "ON a1s.JRNAL_NO=a2s.JRNAL_NO AND a1s.JRNAL_LINE=a2s.JRNAL_LINE WHERE a1s.JRNAL_NO = " + jrnalNoDB + " and a1s.TREFERENCE like '" + jrnalRefDB + "'" +
                    "order by a1s.JRNAL_NO, a1s.TREFERENCE, a1s.JRNAL_LINE";
                    
                    break; 
                }
                case 5:
                {
                    queryLayout =
                    "select a2s.PERIOD, a2s.JRNAL_NO, a2s.JRNAL_LINE, " +
                    "a2s.GNRL_DESCR_01, a2s.GNRL_DESCR_02 AS GD02_NOMER_FAKTURA, a2s.GNRL_DESCR_03, a2s.GNRL_DESCR_05 AS GD5_OLD_JNL_TYP, " +
                    "a2s.GNRL_3_DATETIME AS GDT3_DATE_FAKTURA, a2s.GNRL_1_DATETIME AS GDT1_DATE_OPLATA, a2s.GNRL_2_DATETIME AS GDT2_DATE_PRIN_UCHET, " +
                    "a2s.GNRL_DESCR_11 AS GD11_KODOPLATY, a2s.GNRL_DESCR_12 AS GD12, a2s.GNRL_DESCR_13 AS GD13, a2s.GNRL_DESCR_14 AS GD14, " +
                    "a2s.GNRL_DESCR_15 AS GD15, a2s.GNRL_DESCR_16 AS GD16, a2s.GNRL_DESCR_17 AS GD17, a2s.GNRL_DESCR_18 AS GD18, a2s.GNRL_DESCR_19 AS GD19 " +
                    "FROM [SunSystemsData].." + busUnit + "_A_SALFLDG_LAD as a2s with (NOLOCK) " +
                    "WHERE a2s.JRNAL_NO = " + jrnalNoDB + 
                    " order by a2s.JRNAL_NO, a2s.JRNAL_LINE";

                    break;
                } 

                case 6:
                {
                    queryLayout =
                    "select * " +
                    "FROM [SunSystemsData].." + busUnit + "_A_SALFLDG as a1s with (NOLOCK) " +
                    "WHERE a1s.JRNAL_NO = " + jrnalNoDB + " and a1s.TREFERENCE like '" + jrnalRefDB + "'" +
                    "order by a1s.JRNAL_NO, a1s.TREFERENCE, a1s.JRNAL_LINE";

                    break;
                } 
            }

            DAJrnal = new SqlDataAdapter(queryLayout, connectionStringConf);
            DSJrnal = new DataSet();

            DAJrnal.Fill(DSJrnal, "TabSunJrnalLDG");

            return DSJrnal;
        }


        public string GetCurrentPeriod(string businessUnit)
        {

            using (con = new SqlConnection(connectionString))
            {

//                SqlCommand com = new SqlCommand("SELECT SUBSTRING(SUN_DATA, 31, 7) AS CURRENT_PERIOD FROM " + businessUnit + "_SSRFMSC WHERE SUN_TB = 'LDG'", con);
//                SqlCommand com = new SqlCommand("SELECT SUBSTRING(SUN_DATA, 42, 8) AS OPEN_PERIOD FROM " + businessUnit + "_SSRFMSC WHERE SUN_TB = 'LDG'", con);                
                SqlCommand com = new SqlCommand("SELECT RTRIM(CURRENT_PERD) FROM SunSystemsData.." + businessUnit + "_LDG_DEFN", con);
                try
                {
                    con.Open();

                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);

                    if (dr.HasRows)
                        foreach (DbDataRecord result in dr)
                        {
                            period = result[0].ToString();
                        }

                    if (period.Trim() == "2017013")
                    {
                        period = "2018001";
                    }

                }
                catch
                {
                    period = "none";
                }
            }
            return period.Trim();
        }

        public string[] GetSunProfile(string winuser)
        {
            string[] sunProf = new string[5];

            using (con = new SqlConnection(connectionStringConf))
            {

                //OdbcCommand com = new OdbcCommand("SELECT TOP 1 OPR_CODE FROM OPRB WHERE PWIN_USER = '"+winuser+"'", con);
                SqlCommand com = new SqlCommand("SELECT TOP 1 OPR_CODE, PJDEL_MAX, PROL, PJRAL_MAX, BUNIT FROM OPRB WHERE PWIN_USER = '" + winuser + "'", con);

                try
                {
                   
                    con.Open();

                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);

                    if (dr.HasRows)
                        foreach (DbDataRecord result in dr)
                        {
                            sunuser = result[0].ToString();
                            djmax = result[1].ToString();
                            curol = result[2].ToString();
                            rallocmax = result[3].ToString();
                            bunit = result[4].ToString();
                        }
                }
                catch
                {
                    sunuser = "none";
                    curol = "R01";
                    djmax = "0";
                    rallocmax = "0";
                    bunit = "none";
                }
            }

            sunProf[0] = sunuser;
            sunProf[1] = djmax;
            sunProf[2] = curol;
            sunProf[3] = rallocmax;
            sunProf[4] = bunit;

            return sunProf;
        }

        public int RecordLog(string strAPPLIC, string strCODOPER, string strWINUSER, string JrnalNo, string SunUser, string busUnit, int cntSALF, int cntAlloc, int cntLAD, string resAct, string JrnalRef)
        {

            using (con = new SqlConnection(connectionString))
            {

                SqlCommand com = new SqlCommand("INSERT INTO [RU0_4T_DJLOG] (APPLIC, CODOPER, USRNAME, DJ_NO_INT, DJ_SRCE_CH, DJ_UNIT, QTN_SLDG_INT, QTN_SLDGA_INT, QTN_SLDGLAD_INT, DJ_STATUS, DJ_DATETIME, DJ_REFS)" +
                "VALUES ('" + strAPPLIC + "', '" + strCODOPER + "', '" + strWINUSER + "'," + int.Parse(JrnalNo) +", '" + SunUser.Trim() + "', '" + busUnit + "', " + cntSALF + ", " + cntAlloc + ", " + cntLAD + ", '" + resAct + "', GETDATE(), '" + JrnalRef + "')", con);

                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
                }
                catch
                {
                    return 0;
                }
            }

            return 1;
        }


        internal string journalDelLayout(string busUnit, string jNumber, string period, string sunUser, string djRol, string winUser)
        {

            string resJrnalNo = "", resPeriod = "", resSunUser = "";
            int logRes;

            ArrayList LedgerRows = new ArrayList();

            using (con = new SqlConnection(connectionString))
            {

                SqlCommand com = new SqlCommand("SELECT DISTINCT JRNAL_NO, JRNAL_SRCE, PERIOD FROM [SunSystemsData].." + busUnit + "_A_SALFLDG WHERE JRNAL_NO =" + jNumber, con);

                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
                  

                    //Если запрос вернул 1 и более строк
                    if (dr.HasRows)
                    {
                        int k = 0;
                        foreach (DbDataRecord result in dr)
                        //                            LedgerRows1.Add(result.GetValue(1));
                        {
                            //Получаем реальные значения JRNAL_SRCE, PERIOD с учетом JRNAL_NO и BusUnit
                            resJrnalNo = result[0].ToString();
                            resSunUser = result[1].ToString();
                            resPeriod = result[2].ToString();
                            k++;
                        }

                        //Проверяем количество возвращенных строк строк
//                        if (k > 1 && djRol != "R03")
                        if (k > 1 && djRol != "R04" && djRol != "R05")
                        {   
                            output = "Не удалён: Журнал № <" + jNumber + "> обнаружен в базе " + busUnit + ", имеет " + k.ToString() + " сочетаний JRANL_SRCE, PERIOD. Обратитесь к Администратору;";
                            return output;
                        }
                    }
                    else
                    {
                        output = "Не удалён: Журнал № <" + jNumber + "> не обнаружен в базе " + busUnit + ";";
                        return output;
                    }
                }
                catch
                { }

                //Проверяем совпадение Пользователя
                //OTE R01
                if (resSunUser.Trim() != sunUser.Trim() && djRol != "R02" && djRol != "R04" && djRol != "R05")
                //if (resSunUser.Trim() != "AS" && djRol != "R02" && djRol != "R04" && djRol != "R05")
                {
                    output = "Не удалён: Владелец ["+ resSunUser.Trim() +"] удаляемого журнала <" + jNumber + "> не соответствует текущему пользователю {" + sunUser + "};";
                    //output = "Владелец удаляемого журнала <" + resSunUser.Trim() + ">, удаляющий пользователь " + sunUser.Trim() + ";";
                    return output;
                }

                //Проверяем INT, чтобы период был больше или равно Периода С 
                if (int.Parse(resPeriod.Trim()) != int.Parse(period.Trim()))
                {
                    if (int.Parse(resPeriod.Trim()) < int.Parse(period.Trim()) && djRol != "R03" && djRol != "R04" && djRol != "R05")
                    {
                        output = "Не удалён: Период " + resPeriod + " удаляемого журнала <" + jNumber + "> не соответствует открытому в Sun;";
                        return output;
                    }
                }
                int[] counts = new int[4];
                //Получаем состав журнала журнала
                counts = checkJournalCounts(resJrnalNo, resSunUser, busUnit);
                output = "V1=" + counts[0] + "; V2=" + counts[1] + "; V3="+ counts[2] +"; V4="+ counts[3]+";";

                int[] delResults = new int[5];

                if (counts[1] > 0)
                {
                    logRes = RecordLog("SUNPLUS_01", "JRNALDEL10", winUser, resJrnalNo, resSunUser, busUnit, 0, counts[1], 0, "FLAG_A_OPER_CANCELED", string.Empty);
                    output = "Не удалён: Удаление журнала <" + jNumber + "> не возможно, т.к. в журнале имеется " + counts[2] + " аллокаций ["+ logRes +"];";
                    return output;
                }
                else
                {

                    delResults = JournalDelete(resJrnalNo, busUnit, counts[2], counts[3]);
                    if (delResults[0] == 1)
                    {
                        logRes = RecordLog("SUNPLUS_01", "JRNALDEL10", winUser, resJrnalNo, resSunUser, busUnit, delResults[2], 0, delResults[3], "JRNAL_DELETED", string.Empty);
                        output = "Журнал № <" + jNumber + "> пользователя '" + resSunUser.Trim() + "' в количестве " + delResults[2] + " линий удален [" + logRes + "];";
                    }
                    else if (delResults[0] == 0)
                    {
                        logRes = RecordLog("SUNPLUS_01", "JRNALDEL10", winUser, resJrnalNo, resSunUser, busUnit, 0, 0, 0, "DELETE_UNSUCCESSFUL", string.Empty);
                        output = "Журнал <" + jNumber + "> не удален [" + logRes + "];;";
                    }
                
                }

            }
            return output;
        }

        //FOR DELETE
        private int[] checkJournalCounts(string jNumber, string sunUser, string busUnit)
        {
            int[] cntLines = new int[4];

            using (con = new SqlConnection(connectionString))
            {
                //Получаем количество строк журнала
                SqlCommand com = new SqlCommand(
                "SELECT COUNT(*) FROM [SunSystemsData]..[" + busUnit + "_A_SALFLDG] WHERE JRNAL_NO = "+ jNumber +" AND (JRNAL_SRCE = '"+ sunUser +"' AND ACCNT_CODE <> '99999');" +
                "SELECT COUNT(*) FROM [SunSystemsData]..[" + busUnit + "_A_SALFLDG] WHERE JRNAL_NO = " + jNumber + " AND (JRNAL_SRCE = '" + sunUser + "' AND ACCNT_CODE <> '99999') AND ALLOCATION LIKE 'A';" +
                "SELECT COUNT(*) FROM [SunSystemsData]..[" + busUnit + "_A_SALFLDG] WHERE JRNAL_NO = " + jNumber + ";" +
                "SELECT COUNT(*) FROM [SunSystemsData]..[" + busUnit + "_A_SALFLDG_LAD] WHERE JRNAL_NO = " + jNumber +";"
                , con);

                try
                {
                    con.Open();

                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
                    int i = 0;

                    if (dr.HasRows)
                    do
                    {
                        
                        while (dr.Read())
                        {
                            if (i == 0)
                            {
                                cntLines[0] = int.Parse(dr[0].ToString());
                            }
                            else if (i == 1)
                            {
                                cntLines[1] = int.Parse(dr[0].ToString());
                            }
                            else if (i == 2)
                            {
                                cntLines[2] = int.Parse(dr[0].ToString());
                            }
                            else if (i == 3)
                            {
                                cntLines[3] = int.Parse(dr[0].ToString());
                            }

                        }
                        i++;

                    } while (dr.NextResult());

                }
                catch
                { }
            }
            return cntLines;


        }
        //for RAZALLOC
        private int[] checkJournalCounts(string jNumber, string sunUser, string busUnit, string jRefer)
        {
            int[] cntLines = new int[3];

            using (con = new SqlConnection(connectionString))
            {
                //Получаем количество строк журнала
                SqlCommand com = new SqlCommand(
                "SELECT COUNT(*) FROM [SunSystemsData]..[" + busUnit + "_A_SALFLDG] WHERE JRNAL_NO = " + jNumber + " AND TREFERENCE = '" + jRefer + "' AND (JRNAL_SRCE = '" + sunUser + "' AND ACCNT_CODE <> '99999');" +
                "SELECT COUNT(*) FROM [SunSystemsData]..[" + busUnit + "_A_SALFLDG] WHERE JRNAL_NO = " + jNumber + " AND TREFERENCE = '" + jRefer + "' AND (JRNAL_SRCE = '" + sunUser + "' AND ACCNT_CODE <> '99999') AND ALLOCATION LIKE 'A';" +
                "SELECT COUNT(ALLOC_REF) FROM [SunSystemsData]..[" + busUnit + "_A_SALFLDG] WHERE JRNAL_NO = " + jNumber + " AND TREFERENCE = '" + jRefer + "' AND (JRNAL_SRCE = '" + sunUser + "' AND ACCNT_CODE <> '99999') AND ALLOCATION LIKE 'A' AND ALLOC_REF<>0;"
                , con);
                try
                {
                    con.Open();

                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
                    int i = 0;

                    if (dr.HasRows)
                        do
                        {

                            while (dr.Read())
                            {
                                if (i == 0)
                                {
                                    cntLines[0] = int.Parse(dr[0].ToString());
                                }
                                else if (i == 1)
                                {
                                    cntLines[1] = int.Parse(dr[0].ToString());
                                }
                                else if (i == 2)
                                {
                                    cntLines[2] = int.Parse(dr[0].ToString());
                                }
                            }
                            i++;

                        } while (dr.NextResult());

                }
                catch
                { }
            }
            return cntLines;
        }

        //1 - удаление, 2- сдвиг периода

        internal int JrnalPerDayModif(string businessUnit, string gv_winuser, int typmod)
        {
            int CountDeleted = -1;
            SqlCommand com = null;

            using (con = new SqlConnection(connectionString))
            {
                if (typmod == 1)
                {
                    com = new SqlCommand("SELECT COUNT(*) FROM RU0_4T_DJLOG WHERE USRNAME = '" + gv_winuser + "' AND CONVERT(nvarchar(10), DJ_DATETIME, 103) = CONVERT(nvarchar(10), GETDATE(), 103) AND APPLIC = 'SUNPLUS_02' AND CODOPER = 'REFRAZLK10' AND DJ_STATUS = 'FLAG_A_ISOUT_CANCELD'", con);
                }
                else if (typmod == 2)
                {
                    com = new SqlCommand("SELECT COUNT(*) FROM RU0_4T_DJLOG WHERE USRNAME = '" + gv_winuser + "' AND CONVERT(nvarchar(10), DJ_DATETIME, 103) = CONVERT(nvarchar(10), GETDATE(), 103) AND APPLIC = 'SUNPLUS_02' AND CODOPER = 'JRNLSHIF12' AND DJ_STATUS = 'PER_IS_SHIFT'", con);
                }
 

                try
                {
                    con.Open();

                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);

                    if (dr.HasRows)
                        foreach (DbDataRecord result in dr)
                        {
                             CountDeleted = int.Parse(result[0].ToString());
                        }

                }
                catch
                {
                    CountDeleted = 0;
                }

            }

            return CountDeleted;
        }

        internal int RefPerDayRazalloc(string businessUnit, string gv_winuser)
        {

            int CountRazalloc = -1;

            using (con = new SqlConnection(connectionString))
            {

                SqlCommand com = new SqlCommand("SELECT COUNT(*) FROM RU0_4T_DJLOG WHERE USRNAME = '" + gv_winuser + "' AND CONVERT(nvarchar(10), DJ_DATETIME, 103) = CONVERT(nvarchar(10), GETDATE(), 103) AND APPLIC = 'SUNPLUS_01' AND CODOPER = 'REFRAZLK10' AND DJ_STATUS = 'JNREF_RALLOC_SUCCESF'", con);

                try
                {
                    con.Open();

                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);

                    if (dr.HasRows)
                        foreach (DbDataRecord result in dr)
                        {
                            CountRazalloc = int.Parse(result[0].ToString());
                        }

                }
                catch
                {
                    CountRazalloc = 0;
                }

            }

            return CountRazalloc;
        }

        internal int CourierPerDayIns(string businessUnit, string gv_winuser)
        {
            {
                int CountAddCourier = -1;

                using (con = new SqlConnection(connectionString))
                {

                    SqlCommand com = new SqlCommand("SELECT COUNT(*) FROM RU0_4T_DJLOG WHERE USRNAME = '" + gv_winuser + "' AND CONVERT(nvarchar(10), DJ_DATETIME, 103) = CONVERT(nvarchar(10), GETDATE(), 103) AND APPLIC = 'SUNPLUS_01' AND CODOPER = 'COURADD12' AND DJ_STATUS = 'COUR_ADDED_SUCSF'", con);

                    try
                    {
                        con.Open();

                        SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);

                        if (dr.HasRows)
                            foreach (DbDataRecord result in dr)
                            {
                                CountAddCourier = int.Parse(result[0].ToString());
                            }

                    }
                    catch
                    {
                        CountAddCourier = 0;
                    }

                }

                return CountAddCourier;
            }
        }


        private int[] JournalDelete(string resJrnalNo, string busUnit, int countJrnal, int countJrnalLad)
        {

            int[] result = new int[5];

            int delFlag = 1;

            SqlCommand insLAD;
            SqlCommand delLAD;
            SqlCommand insLDG;
            SqlCommand delLDG;

            int varInsLAD = 0, varDelLAD = 0;
            int varInsLDG = 0, varDelLDG = 0;



            using (con = new SqlConnection(connectionString))
            {

                con.Open();
                SqlTransaction sqlTransact = con.BeginTransaction();

                //INSERT в BACKUP таблицу
                //DELETE из SALFLDG таблицы
                insLAD = new SqlCommand("INSERT INTO [" + busUnit + "_4T_DJ_LAD] SELECT * FROM [" + busUnit + "_A_SALFLDG_LAD] WHERE JRNAL_NO = " + resJrnalNo, con, sqlTransact);
                delLAD = new SqlCommand("DELETE FROM [SunSystemsData]..[" + busUnit + "_A_SALFLDG_LAD] WHERE JRNAL_NO = " + resJrnalNo, con, sqlTransact);

                insLDG = new SqlCommand("INSERT INTO [" + busUnit + "_4T_DJ] SELECT * FROM [" + busUnit + "_A_SALFLDG] WHERE JRNAL_NO = " + resJrnalNo, con, sqlTransact);
                delLDG = new SqlCommand("DELETE FROM [SunSystemsData]..[" + busUnit + "_A_SALFLDG] WHERE JRNAL_NO = " + resJrnalNo, con, sqlTransact);

                try
                {

                    if (countJrnalLad > 0)
                    {
                        varInsLAD = insLAD.ExecuteNonQuery();
                        varDelLAD = delLAD.ExecuteNonQuery();
                    }

                    varInsLDG = insLDG.ExecuteNonQuery();
                    varDelLDG = delLDG.ExecuteNonQuery();

                    result[1] = varInsLDG;
                    result[2] = varDelLDG;
                    result[3] = varInsLAD;
                    result[4] = varDelLAD;

                }
                catch
                {
                    sqlTransact.Rollback();
                    delFlag = 0;
                }

                sqlTransact.Commit();
                result[0] = delFlag;
            }

            return result;

        }

        //string resJrnalNo, string busUnit, int countJrnal, int countJrnalLad
        //cbbx_busUnit.SelectedItem.ToString(), txbx_journalNumber.Text.ToString(), txbx_reference.Text.ToString(), txbx_sunUser.Text, lbl_role.Text
        internal string RazAllocLayout(string busUnit, string resJrnalNum, string resJrnalRef, string sunUser, string resRole, string winUser)
        {
            //1 определяем, какому пользователю принадлежит журнал+референс
            //2 если он принадлежит допустимому пользователю и допустимой роли, определяем, есть ли аллокации в этом референсе
            //3 если есть аллокации в этом референсе, разаллокируем
            //4 сообщаем, что разаллокировали

            string resJrnalNo = "", resPeriod = "", resSunUser = "";
            int logRes;

            using (con = new SqlConnection(connectionString))
            {

                SqlCommand com = new SqlCommand("SELECT DISTINCT JRNAL_NO, JRNAL_SRCE, PERIOD FROM [SunSystemsData].." + busUnit + "_A_SALFLDG WHERE JRNAL_NO =" + resJrnalNum + " AND TREFERENCE = '" + resJrnalRef + "'", con);

                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);



                    //Если запрос вернул 1 и более строк
                    if (dr.HasRows)
                    {
                        int k = 0;
                        foreach (DbDataRecord result in dr)
                        //                            LedgerRows1.Add(result.GetValue(1));
                        {
                            //Получаем реальные значения JRNAL_SRCE, PERIOD с учетом JRNAL_NO и BusUnit
                            resJrnalNo = result[0].ToString();
                            resSunUser = result[1].ToString();
                            resPeriod = result[2].ToString();
                            k++;
                        }

                        //Проверяем количество возвращенных строк строк (for Accruals in two periods)
                        //                        if (k > 1 && djRol != "R03")
                        if (k > 1 && resRole != "R04" && resRole != "R05")
                        {
                            output = "Не разаллокирован: Журнал № <" + resJrnalNum + "> обнаружен в базе " + busUnit + ", имеет " + k.ToString() + " сочетаний JRANL_SRCE, PERIOD;";
                            return output;
                        }
                    }
                    else
                    {
                        output = "Не разаллокирован: Журнал № <" + resJrnalNum + "> не обнаружен в базе " + busUnit + ";";
                        return output;
                    }
                }
                catch
                { }

                //Проверяем совпадение Пользователя
                //OTE R01
                if (resSunUser.Trim() != sunUser.Trim() && resRole != "R02" && resRole != "R04" && resRole != "R05")
                //if (resSunUser.Trim() != "AS" && resRole != "R02" && resRole != "R04" && resRole != "R05")
                {
                    output = "Не разаллокирован: Владелец [" + resSunUser.Trim() + "] референса журнала <" + resJrnalNum + "> не соответствует текущему пользователю {" + sunUser.Trim() + "};";
                    //output = "Владелец удаляемого журнала <" + resSunUser.Trim() + ">, удаляющий пользователь " + sunUser.Trim() + ";";
                    return output;
                }

                //Проверяем совпадение Периода
                if (resPeriod.Trim() != period.Trim() && resRole != "R03" && resRole != "R04" && resRole != "R05")
                {
                    output = "Не разаллокирован: Период " + resPeriod + " референса журнала <" + resJrnalNum + "> не соответствует открытому в Sun;";
                    return output;
                }

                int[] counts = new int[2];
                //Получаем состав журнала журнала
                counts = checkJournalCounts(resJrnalNo, resSunUser, busUnit, resJrnalRef);
                output = "V1=" + counts[0] + "; V2=" + counts[1] + "; V3=" + counts[2] + ";";

                int[] RazallocResults = new int[3];

                if (counts[1] == 0)
                {
                    //'REFRAZLK10' AND DJ_STATUS = 'REFER_RAZALLK'
                    logRes = RecordLog("SUNPLUS_01", "REFRAZLK10", winUser, resJrnalNo, resSunUser, busUnit, 0, counts[1], 0, "FLAG_A_ISOUT_CANCELD", resJrnalRef);
                    output = "Не разаллокировано: В журнале <" + resJrnalNo + "> по операциям с данным референсом отсутствует Признак аллокирования [" + logRes + "];";
                    return output;
                }
                else if (counts[2] == 0)
                {
                    //'REFRAZLK10' AND DJ_STATUS = 'REFER_RAZALLK'
                    logRes = RecordLog("SUNPLUS_01", "REFRAZLK10", winUser, resJrnalNo, resSunUser, busUnit, 0, counts[1], 0, "NUM_AL_ISOUT_CANCELD", resJrnalRef);
                    output = "Не разаллокировано: В журнале <" + resJrnalNo + "> по операциям с данным референсом отсутствует Номер аллокации. Обратитесь к администратору Sun. [" + logRes + "];";
                    return output;
                }
                else
                {
//resJrnalNo, resSunUser, busUnit, resJrnalRef
                    RazallocResults = JournalRazalloc(resJrnalNo, resJrnalRef, busUnit);
                    if (RazallocResults[0] == 1)
                    {
                        logRes = RecordLog("SUNPLUS_01", "REFRAZLK10", winUser, resJrnalNo, resSunUser, busUnit, 0, RazallocResults[1], 0, "JNREF_RALLOC_SUCCESF", resJrnalRef);
                        output = "Референс <" + resJrnalRef + "> журнала № <" + resJrnalNo + "> пользователя '" + resSunUser.Trim() + "' разаллокирован  [" + logRes + "];";
                    }
                    else if (RazallocResults[0] == 0)
                    {
                        logRes = RecordLog("SUNPLUS_01", "REFRAZLK10", winUser, resJrnalNo, resSunUser, busUnit, 0, 0, 0, "DELETE_UNSUCCESSFUL", resJrnalRef);
                        output = "Референс <" + resJrnalRef + "> журнала № <" + resJrnalNo + "> не разаллокирован [" + logRes + "];";
                    }
                }
            }
            return output;
        }

        //Разаллокирование
        private int[] JournalRazalloc(string resJrnalNo, string resJrnalRef, string busUnit)
        {

            int[] result = new int[2];

            int razallocFlag = 1;

            SqlCommand RazAlloc;

            int varRazalloc = 0;

            using (con = new SqlConnection(connectionString))
            {

                con.Open();
                SqlTransaction sqlTransact = con.BeginTransaction();

                RazAlloc = new SqlCommand("UPDATE [SunSystemsData]..[" + busUnit + "_A_SALFLDG] SET ALLOCATION = '', ALLOC_REF = 0, ALLOC_DATETIME = null, ALLOC_PERIOD = 0, ALLOCN_CODE = null, SPLIT_ORIG_LINE = 0" +
                    " where ALLOC_REF in " +
                    "(select DISTINCT(ALLOC_REF) from [SunSystemsData]..[" + busUnit + "_A_SALFLDG] with(Nolock)" + 
                    " WHERE JRNAL_NO=" + resJrnalNo + " and ALLOC_REF!=0 " +
                    " and TREFERENCE IN ('" + resJrnalRef + "'))", con, sqlTransact);
                /*
                update M11_A_SALFLDG
                set ALLOCATION = '', ALLOC_REF = 0, ALLOC_DATETIME = null, ALLOC_PERIOD = 0, ALLOCN_CODE = null, SPLIT_ORIG_LINE = 0
                where ALLOC_REF in (
                 select DISTINCT(ALLOC_REF)
                 from M11_A_SALFLDG with(Nolock)
                where 
                JRNAL_NO IN (223276) and ALLOC_REF!=0 
                --
                and TREFERENCE IN ('157999')
                ) 
                */

                try
                {
                    RazAlloc.CommandTimeout = 300;
                    varRazalloc = RazAlloc.ExecuteNonQuery();
                    result[1] = varRazalloc;
                }
                catch
                {
                    sqlTransact.Rollback();
                    razallocFlag = 0;
                }

                sqlTransact.Commit();
                result[0] = razallocFlag;
            }

            return result;
           
        }

        
        //TRANSACTION MATCHING
        internal int TmPerDay(string businessUnit, string gv_winuser)
        {

            int CountTmAlloc = -1;

            using (con = new SqlConnection(connectionString))
            {

                SqlCommand com = new SqlCommand("SELECT COUNT(*) FROM RU0_4T_DJLOG WHERE CONVERT(nvarchar(10), DJ_DATETIME, 103) = CONVERT(nvarchar(10), GETDATE(), 103) AND APPLIC = 'SUNPLUS_01' AND CODOPER = 'TMALLOCA11' AND DJ_STATUS = 'TMA11_ALLOCC_SUCCESF'", con);

                try
                {
                    con.Open();

                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);

                    if (dr.HasRows)
                        foreach (DbDataRecord result in dr)
                        {
                            CountTmAlloc = int.Parse(result[0].ToString());
                        }

                }
                catch
                {
                    CountTmAlloc = 0;
                }

            }

            return CountTmAlloc;

        }

        internal string TmAllocLayout(string busUnit, string sunUser, string resRole, string gv_winuser)
        {

            //1 Запускаем процедуру
            //2 получаем результат
            //3 Пишем в лог

            string resCodTm = "", resTmStart = "", resTmFinsh = "";
            int resTmAlloc, resCntAlloc, logRes;

            resCntAlloc = 0;

            using (con = new SqlConnection(connectionString))
            {

                SqlCommand com = new SqlCommand("EXEC _TM_A11LINK", con);
                com.CommandTimeout = 0;

                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);


                    //Если запрос вернул 1 и более строк
                    if (dr.HasRows)
                    {
                        int k = 0;
                        foreach (DbDataRecord result in dr)
                        //                            LedgerRows1.Add(result.GetValue(1));
                        {
                            //Получаем Результат TM
                            resTmAlloc = Int32.Parse(result[0].ToString());
                            resCntAlloc = Int32.Parse(result[1].ToString());
                            resCodTm = result[2].ToString();
                            resTmStart = result[3].ToString();
                            resTmFinsh = result[4].ToString();
                            k++;
                        }

                        //Проверяем количество возвращенных строк строк
                        if (k > 1)
                        {
                            logRes = RecordLog("SUNPLUS_01", "TMALLOCA11", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "TMA11_ALLOCC_ERROR2", "TM Error2;");
                            output = "TM [TM Error 2]: Не отработал [" + logRes + "];";
                        }
                    }
                    else
                    {
                        logRes = RecordLog("SUNPLUS_01", "TMALLOCA11", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "TMA11_ALLOCC_ERROR1", "TM Error1;");
                        output = "TM [TM Error 1]: Не отработал [" + logRes + "];";

                        return output;
                    }
                }
                catch
                { }

                logRes = RecordLog("SUNPLUS_01", "TMALLOCA11", gv_winuser, "0", sunUser, busUnit, resCntAlloc, 0, 0, "TMA11_ALLOCC_SUCCESF", "StartTime: " + resTmStart + "; FinishTime: " + resTmFinsh + "; Code TM:" + resCodTm + ";");
                output = "TM отработал: пользователь '" + gv_winuser + "', старт [" + resTmStart + "], окончание [" + resTmFinsh + "], код: " + resCodTm + " [" + logRes + "];";

            }

            return output;

        }


        internal string courierInsLayout(string busUnit, string resCourierFIO, string resPeriod, string sunUser, string resRole, string gv_winuser)
        {
            //cbbx_busUnit.SelectedItem.ToString(), txbx_fiocourier.ToString(), lbl_period.Text, txbx_sunUser.Text, lbl_role.Text, gv_winuser
            //string busUnit, string resJrnalNum, string resJrnalRef, string sunUser, string resRole, string winUser
            string logUnit, output;
            int logRes;

            if (busUnit != "RU0")
            {
                logUnit = "RU4";
            }
            else
            {
                logUnit = "RU0";
            }

            using (con = new SqlConnection(connectionString))
            {

                //SqlCommand com = new SqlCommand("INSERT INTO [" + logUnit + "_ANL_CODE]" +// (APPLIC, CODOPER, USRNAME, DJ_NO_INT, DJ_SRCE_CH, DJ_UNIT, QTN_SLDG_INT, QTN_SLDGA_INT, QTN_SLDGLAD_INT, DJ_STATUS, DJ_DATETIME, DJ_REFS)" +
                                                //"VALUES('13', left('" + resCourierFIO + "',15), 0, 'FSA', getdate(), 0, left('" + resCourierFIO + "',15),'" + resCourierFIO +"', null, 0, 0, 0, 99, 0)", con);

                SqlCommand com = new SqlCommand("select COUNT(*) from RU0_ANL_CODE where ANL_CAT_ID=34 AND UPPER(NAME) = '" + resCourierFIO +"'", con);

                int CountCour = -1;
                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);

                    if (dr.HasRows)
                    {
                        foreach (DbDataRecord result in dr)
                        {
                            CountCour = int.Parse(result[0].ToString());
                        }
                    }

                    if (CountCour > 0)
                    {
                        logRes = RecordLog("SUNPLUS_01", "COURADD12", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "COUR_DUBLR_ERROR", resCourierFIO);
                        output = "Курьер не добавлен. Курьер с ФИО <" + resCourierFIO + "> уже существует в SUN [" + logRes + "];";
                        return output;
                    }

                    com.CommandText = "INSERT INTO [" + logUnit + "_ANL_CODE] VALUES('34', left('" + resCourierFIO + "',15), 0, '" + sunUser + "', getdate(), 0, left('" + resCourierFIO + "',15),'" + resCourierFIO +"', null, 0, 0, 0, 99, 0)";
                    con.Open();
                    dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);

                }
                catch
                {
                    logRes = RecordLog("SUNPLUS_01", "COURADD12", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "COUR_ADDED_ERROR", resCourierFIO);
                    output = "Курьер <" + resCourierFIO + "> не добавлен в SUN [" + logRes + "];";
                }

                logRes = RecordLog("SUNPLUS_01", "COURADD12", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "COUR_ADDED_SUCSF", resCourierFIO);
                output = "Курьер <" + resCourierFIO + "> успешно добавлен в SUN [" + logRes + "];";
            }

            //'SUNPLUS_01' AND CODOPER = 'COURADD10' AND DJ_STATUS = 'COURR_ADDED_SUCCESF'"

            return output;

        }

        // dal.OpenTradePerLayout(this.cbbx_region.SelectedItem.ToString(), this.cbbx_TradePeriod.Text.ToString(), this.txbx_ComntPeriod.Text.ToString(), "TRZ", "M14", gv_winuser) + "\n" + rtxbx_info.Text;
        internal string OpenTradePerdLayout(string strRegion, string strPeriod, string strComntOPer, string sunUser, string strRole, string gv_winuser)
        {

            int[] cntLines = new int[1];
            string strMonth, strYear, output;
            int logRes;
            strMonth = strPeriod.Substring(0, 2);
            strYear = strPeriod.Substring(2, 4);
            output = "";

            using (con = new SqlConnection(conStringTradeAO))
            {
                //Изменить период
                //Записать лог
                SqlCommand com = new SqlCommand("UPDATE dbo._Reference1 SET _Fld10 = DATEADD(dd,-1,CAST((CAST('"+ strYear + "' AS NVARCHAR(4))+'-'+RIGHT('0'+CAST('"+ strMonth +"' AS NVARCHAR(2)),2)+'-'+'01') AS datetime)) WHERE _Code IN  ('" + strRegion + "')", con);
                
                try
                {
                    con.Open();

                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);


                        logRes = RecordLog("SUNPLUS_01", "TRPERD11", gv_winuser, "0", sunUser, strRole, 0, 0, 0, "TRAO_OPN_SUCCS", "ОТКРЫТ " + strRegion + " " + strPeriod + " " + strComntOPer);
                        output = "Период <" + strPeriod + "> ОТКРЫТ в TRADE для региона <" + strRegion + "> [" + logRes + "];";


                }
                catch
                { 
                    logRes = RecordLog("SUNPLUS_01", "TRPERD11", gv_winuser, "0", sunUser, strRole, 0, 0, 0, "TRAO_OPN_ERROR", "НЕ ОТКРЫТ " + strRegion + " " + strPeriod + " "+ strComntOPer);
                    output = "Период " + strPeriod + " НЕ ОТКРЫТ в TRADE для региона <" +  strRegion + ">. Обратитесь к Администратоту [" + logRes + "];";
                }
            }

            return output;

        }

        internal string GetOpenTradePerd(string strRegion, string strUnit, string gv_winuser)
        {
            string[] cntLines = new string[1];

            using (con = new SqlConnection(conStringTradeAO))
            {
                //Изменить период
                //Записать лог
//                SqlCommand com = new SqlCommand("select RIGHT('0'+CAST(MONTH(_Fld10) as nvarchar(2)),2) + CAST(YEAR(_Fld10) as nvarchar(4)) as PERIOD from dbo._Reference1 WHERE _Code IN ('" + strRegion + "')", con);
                SqlCommand com = new SqlCommand("select RIGHT('0'+CAST(MONTH(DATEADD(mm,1,_Fld10)) as nvarchar(2)),2) + CAST(YEAR(_Fld10) as nvarchar(4)) as PERIOD from dbo._Reference1 WHERE _Code IN ('" + strRegion + "')", con);
                try
                {
                    con.Open();

                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
                    int i = 0;

                    if (dr.HasRows)
                        do
                        {

                            while (dr.Read())
                            {
                                if (i == 0)
                                {
                                    cntLines[0] = dr[0].ToString();
                                }
                                else if (i == 1)
                                {
                                    cntLines[1] = dr[0].ToString();
                                }
                                else if (i == 2)
                                {
                                    cntLines[2] = dr[0].ToString();
                                }
                                else if (i == 3)
                                {
                                    cntLines[3] = dr[0].ToString();
                                }

                            }
                            i++;

                        } while (dr.NextResult());

                }
                catch
                { }
            }
            return cntLines[0].ToString();
        }

        //Close period
        internal string OpenTradePerdLayout(string strRegion, string sunUser, string strRole, string gv_winuser)
        {

            int[] cntLines = new int[1];
            string output;
            int logRes;

            string curdate;
            DateTime TheDate = DateTime.Today.AddDays(-5);
            curdate = TheDate.Year.ToString() + "-" + TheDate.Month.ToString() + "-" + TheDate.Day.ToString();


            output = "";

            using (con = new SqlConnection(conStringTradeAO))

            {
                //Изменить период
                //Записать лог
                SqlCommand com = new SqlCommand("UPDATE dbo._Reference1  SET _Fld10 = DATEADD(dd,-DAY('" + curdate + "')," +
                                                "CAST((CAST(YEAR('" + curdate + "') AS NVARCHAR(4))+'-'+RIGHT('0'+CAST(MONTH('" + curdate + "') AS NVARCHAR(2)),2)+" +
                                                "'-'+RIGHT('0'+CAST(DAY('" + curdate + "') AS NVARCHAR(2)),2)) AS datetime)) " + 
                                                " WHERE _Code IN ('" + strRegion + "')", con);

                try
                {
                    con.Open();

                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);


                    logRes = RecordLog("SUNPLUS_01", "TRPERD11", gv_winuser, "0", sunUser, strRole, 0, 0, 0, "TRAO_CLS_SUCCS", "ЗАКРЫТ " + strRegion);
                    output = "Регион <" + strRegion + "> ЗАКРЫТ в TRADE [" + logRes + "];";


                }
                catch
                {
                    logRes = RecordLog("SUNPLUS_01", "TRPERD11", gv_winuser, "0", sunUser, strRole, 0, 0, 0, "TRAO_CLS_ERROR", "НЕ ЗАКРЫТ " + strRegion);
                    output = "Регион <" + strRegion + "> НЕ ЗАКРЫТ в TRADE. Обратитесь к Администратору [" + logRes + "];";
                }
            }

            return output;

        }

        internal string courierUnblockLayout(string busUnit, string resCourierFIO, string resPeriod, string sunUser, string resRole, string gv_winuser)
        {
            //cbbx_busUnit.SelectedItem.ToString(), txbx_fiocourier.ToString(), lbl_period.Text, txbx_sunUser.Text, lbl_role.Text, gv_winuser
            //string busUnit, string resJrnalNum, string resJrnalRef, string sunUser, string resRole, string winUser
            string logUnit, output;
            int logRes;

            if (busUnit != "A11")
            {
                logUnit = "M11";
            }
            else
            {
                logUnit = "RU0";
            }

            using (con = new SqlConnection(connectionString))
            {

                //SqlCommand com = new SqlCommand("INSERT INTO [" + logUnit + "_ANL_CODE]" +// (APPLIC, CODOPER, USRNAME, DJ_NO_INT, DJ_SRCE_CH, DJ_UNIT, QTN_SLDG_INT, QTN_SLDGA_INT, QTN_SLDGLAD_INT, DJ_STATUS, DJ_DATETIME, DJ_REFS)" +
                //"VALUES('13', left('" + resCourierFIO + "',15), 0, 'FSA', getdate(), 0, left('" + resCourierFIO + "',15),'" + resCourierFIO +"', null, 0, 0, 0, 99, 0)", con);

                SqlCommand com = new SqlCommand("select COUNT(*) from SunSystemsData..RU0_ANL_CODE where ANL_CAT_ID=13 AND STATUS = 3 AND UPPER(NAME) = '" + resCourierFIO + "'", con);

                int CountCour = -1;
                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);

                    if (dr.HasRows)
                    {
                        foreach (DbDataRecord result in dr)
                        {
                            CountCour = int.Parse(result[0].ToString());
                        }
                    }

                    if (CountCour == 0)
                    {
                        logRes = RecordLog("SUNPLUS_01", "COURUNB12", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "COUR_UNBL_ERROR", resCourierFIO);
                        output = "Курьер с ФИО <" + resCourierFIO + "> не существует или не заблокирован в SUN [" + logRes + "];";
                        return output;
                    }

                    com.CommandText = "UPDATE [" + logUnit + "_ANL_CODE] SET STATUS = 0, LAST_CHANGE_DATETIME = getdate(), LAST_CHANGE_USER_ID = '" + sunUser + "' WHERE UPPER(NAME) = '" + resCourierFIO + "' AND ANL_CAT_ID=13 AND STATUS = 3";
                    con.Open();
                    dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);

                }
                catch
                {
                    logRes = RecordLog("SUNPLUS_01", "COURUNB12", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "COUR_UNBLC_ERROR", resCourierFIO);
                    output = "Курьер <" + resCourierFIO + "> не добавлен в SUN [" + logRes + "];";
                }

                logRes = RecordLog("SUNPLUS_01", "COURUNB12", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "COUR_UNBLC_SUCSF", resCourierFIO);
                output = "Курьер <" + resCourierFIO + "> успешно РАЗБЛОКИРОВАН в SUN [" + logRes + "];";
            }

            //'SUNPLUS_01' AND CODOPER = 'COURADD10' AND DJ_STATUS = 'COURR_ADDED_SUCCESF'"

            return output;

        }

        internal string LoadRateToSun(DateTime RateDate, string gv_winuser, string sunUser)
        {
            String RateDateStr;
            int logRes;

            /*'Public Const SunCnnTst = "ODBC;UID=SUN6;DSN=SUN6;"*/

            RateDateStr = RateDate.Year.ToString() + "-" + RateDate.Month.ToString() + "-" + RateDate.Day.ToString();


            using (con = new SqlConnection(connectionString))
            {

                //                SqlCommand com = new SqlCommand("SELECT SUBSTRING(SUN_DATA, 31, 7) AS CURRENT_PERIOD FROM " + businessUnit + "_SSRFMSC WHERE SUN_TB = 'LDG'", con);
                SqlCommand com = new SqlCommand("EXEC sp_RU1_IMPORT_RATE_TRIDB '" + RateDateStr + "'", con);
                try
                {
                    con.Open();

                    if (com.ExecuteNonQuery() >= 1)
                    {
                        logRes = RecordLog("SUNPLUS_02", "RATELOAD2", gv_winuser, "0", sunUser, "M15", 0, 0, 0, "RATE_LOAD_SUCSF", RateDateStr);
                        output = "Процедура загрузки курсов в SUN RU1RU0 отработала [" + logRes + "];";
                    }

                }
                catch
                {
                        logRes = RecordLog("SUNPLUS_01", "RATELOAD1", gv_winuser, "0", sunUser, "RU0", 0, 0, 0, "RATE_LOAD_ERROR", RateDateStr);
                        output = "Процедура загрузки курсов в SUN RU1RU0 отработала НЕ КОРРЕКТНО... [" + logRes + "];";
                }
            }
            return output;
        }

        internal string JournalInUseOff(string busUnit, string jrnalUseNum, string period, string sunUser, string role, string gv_winuser)
        {
            //cbbx_busUnit.SelectedItem.ToString(), this.txbx_journaUselNumber.Text.ToString(), lbl_period.Text, txbx_sunUser.Text, lbl_role.Text, gv_winuser
            int logRes;
            string logUnit;

            if (busUnit != "M11")
            {
                logUnit = "M11";
            }
            else
            {
                logUnit = "M11";
            }


            using (con = new SqlConnection(connectionString))
            {

                // SqlCommand com = new SqlCommand("SELECT SUBSTRING(SUN_DATA, 31, 7) AS CURRENT_PERIOD FROM " + businessUnit + "_SSRFMSC WHERE SUN_TB = 'LDG'", con);

                /*
                                 RazAlloc = new SqlCommand("UPDATE [" + busUnit + "_A_SALFLDG] SET ALLOCATION = '', ALLOC_REF = 0, ALLOC_DATETIME = null, ALLOC_PERIOD = 0, ALLOCN_CODE = null, SPLIT_ORIG_LINE = 0" +
                    " where ALLOC_REF in " + 
                    "(select DISTINCT(ALLOC_REF) from [" + busUnit + "_A_SALFLDG] with(Nolock)" + 
                    " WHERE JRNAL_NO=" + resJrnalNo + " and ALLOC_REF!=0 " +
                    " and TREFERENCE IN ('" + resJrnalRef + "'))", con, sqlTransact);
                 */
                //SqlCommand com = new SqlCommand("UPDATE " + logUnit + "_A_SALFLDG SET IN_USE_FLAG = '' WHERE JRNAL_NO = '" + jrnalUseNum + "' AND PERIOD = " + period + " AND IN_USE_FLAG <> ''", con);
                SqlCommand com = new SqlCommand("UPDATE [SunSystemsData].." + busUnit + "_A_SALFLDG SET IN_USE_FLAG = '' WHERE JRNAL_NO = " + jrnalUseNum + " AND IN_USE_FLAG <> ''", con);
                try
                {
                    con.Open();

                    if (com.ExecuteNonQuery() >= 1)
                    {
                        logRes = RecordLog("SUNPLUS_01", "JRNINUSE1", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "JRNL_USOF_SUCSF", jrnalUseNum);
                        output = "Статус <IN USE> по журналу <" + jrnalUseNum + "> отключен. [" + logRes + "];";
                    }
                    else
                    {
                        logRes = RecordLog("SUNPLUS_01", "JRNINUSE1", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "JRNL_USOF_NFIND", jrnalUseNum);
                        output = "Снятие Статуса <IN USE> по журналу <" + jrnalUseNum + "> НЕ произведено, т.к. данный статус НЕ обнаружен. [" + logRes + "];";
                    }

                }
                catch
                {
                    logRes = RecordLog("SUNPLUS_01", "JRNINUSE1", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "JRNL_USOF_ERROR", jrnalUseNum);
                    output = "Снятие Статуса <IN USE> по журналу <" + jrnalUseNum + "> НЕ ПРОИЗВЕДЕНО. [" + logRes + "];";
                }
            }
            return output;
        }

        //cbbx_busUnit.SelectedItem.ToString(), this.txbx_supCode2.Text.ToString(), txbx_sunUser.Text, lbl_role.Text, gv_winuser, false
        internal string supUnblock(string busUnit, string supCode, string sunUser, string role, string gv_winuser)
        {
            //cbbx_busUnit.SelectedItem.ToString(), this.txbx_journaUselNumber.Text.ToString(), lbl_period.Text, txbx_sunUser.Text, lbl_role.Text, gv_winuser
            int logRes;
            string logUnit;

            if (busUnit == "RU0")
            {
                logUnit = "RU0";
            }
            else
            {
                logUnit = "RU4";
            }


            using (con = new SqlConnection(connectionString))
            {

                SqlCommand com = new SqlCommand("UPDATE [SunSystemsData].." + busUnit + "_ACNT SET STATUS = '0', LAST_CHANGE_DATETIME = Getdate(), LAST_CHANGE_USER_ID = '" + sunUser.TrimEnd() + "' WHERE ACNT_CODE = '" + supCode + "'", con);

                try
                {
                    con.Open();

                    if (com.ExecuteNonQuery() >= 1)
                    {
                        logRes = RecordLog("SUNPLUS_01", "SUPPUNBL1", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "SUPP_UNBL_SUCSF", supCode);
                        output = "Поставщик <" + supCode + "> Разблокирован в SUN [" + logRes + "];";

                        try
                        {
                        string esender = MainForm.GetSender(gv_winuser);
                        MainForm.SendMail(esender, "Инфо: ПОСТАВЩИК " + supCode + " разблокирован в SUN " + busUnit + ". ","Коллеги, добрый день!\n\nПоставщик: " + supCode + " \n разблокирован в SUN " + busUnit + ". \n\nЭто служебное сообщение (и адрес эл.почты), просьба на него не отвечать. \nСпасибо. \n\nС Уважением, Олег Лесницкий.\nmailto:oleg.lesnitsky@dhl.com", busUnit, 0);
                        }
                        catch
                        {
                            output = "Поставщик <" + supCode + "> Разблокирован в 1C Trade. [" + logRes + " unmess];";
                        }
                    }
                    else
                    {
                        logRes = RecordLog("SUNPLUS_01", "SUPPUNBL1", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "SUPP_UNBL_NFIND", supCode);
                        output = "Поставщик <" + supCode + "> НЕ Разблокирован в SUN. Обратитесь к администратору [" + logRes + "];";
                    }

                }
                catch
                {
                    logRes = RecordLog("SUNPLUS_01", "SUPPUNBL1", gv_winuser, "0", sunUser, busUnit, 0, 0, 0, "SUPP_UNBL_ERROR", supCode);
                    output = "Разблокировка поставщика <" + supCode + "> в SUN НЕ ПРОИЗВЕДЕНА. [" + logRes + "];";
                }
            }
            return output;
        }

        internal string supUnblockTrade(string busUnit, string supCode, string sunUser, string role, string gv_winuser, string reasUnbl)
        {

            //cbbx_busUnit.SelectedItem.ToString(), this.txbx_journaUselNumber.Text.ToString(), lbl_period.Text, txbx_sunUser.Text, lbl_role.Text, gv_winuser
            int logRes;
            string logUnit;
            string curntConString;
            string nmColumn;

            logUnit = "RU0";
            nmColumn = "_Fld994";
            curntConString = conStringTradeAO;

            if (busUnit == "RU4")
            {

                logUnit = "RU4";
                nmColumn = "_Fld1113";
                curntConString = conStringTradeOOO;
            }

            using (con = new SqlConnection(curntConString))
            {
                //Изменить период
                //Записать лог
                SqlCommand com = new SqlCommand("UPDATE dbo._Reference16 SET " + nmColumn + "=0x00 WHERE _Fld144='" + supCode + "'", con);

                try
                {
                    con.Open();

                    if (com.ExecuteNonQuery() >= 1)
                    {
                        logRes = RecordLog("SUNPLUS_01", "SUPPUNBL2", gv_winuser, "0", sunUser, (busUnit == "RU0" ? "TRI" : "TRE"), 0, 0, 0, "SUP2_UNBL_SUCSF", supCode + ": " + reasUnbl);
                        output = "Поставщик <" + supCode + "> Разблокирован в 1C Trade " + (busUnit == "RU0"?"АО":"ООО") + " [" + logRes + "];";
                        
                        try
                        {
                            string esender = MainForm.GetSender(gv_winuser);
                            MainForm.SendMail(esender, "Инфо: ПОСТАВЩИК " + supCode + " разблокирован в 1C Trade " + (busUnit == "RU0" ? "АО" : "ООО"), "Коллеги, добрый день!\n\nПоставщик: " + supCode + " \nразблокирован в 1C Trade " + (logUnit == "RU0" ? "АО" : "ООО") + " (" + reasUnbl +  "). \n\nЭто служебное сообщение (и адрес эл.почты), просьба на него не отвечать. \nСпасибо. \n\nС Уважением, Олег Лесницкий.\nmailto:oleg.lesnitsky@dhl.com", logUnit, 0);
                            MainForm.SendMail("oleg.lesnitsky@dhl.com", "Инфо: ПОСТАВЩИК " + supCode + " разблокирован в 1C Trade " + (busUnit == "RU0" ? "АО" : "ООО"), "Коллеги, добрый день!\n\nПоставщик: " + supCode + " \nразблокирован в 1C Trade " + (logUnit == "RU0" ? "АО" : "ООО") + " (" + reasUnbl + "). \n\nЭто служебное сообщение (и адрес эл.почты), просьба на него не отвечать. \nСпасибо. \n\nС Уважением, Олег Лесницкий.\nmailto:oleg.lesnitsky@dhl.com", logUnit, 0);
                        }
                        catch
                        {
                            output = "Поставщик <" + supCode + "> Разблокирован в 1C Trade " + (logUnit == "RU0" ? "АО" : "ООО") + ". [" + logRes + " unmess];";
                        }
                    }
                    else
                    {
                        logRes = RecordLog("SUNPLUS_01", "SUPPUNBL2", gv_winuser, "0", sunUser, (busUnit == "RU0" ? "TRI" : "TRE"), 0, 0, 0, "SUP2_UNBL_NFIND", supCode + ": " + reasUnbl);
                        output = "Поставщик <" + supCode + "> НЕ Разблокирован в 1C Trade " + (logUnit == "RU0" ? "АО" : "ООО") + ". Обратитесь к администратору [" + logRes + "];";
                    }

                }
                catch
                {
                    logRes = RecordLog("SUNPLUS_01", "SUPPUNBL2", gv_winuser, "0", sunUser, (busUnit == "RU0" ? "TRI" : "TRE"), 0, 0, 0, "SUP2_UNBL_ERROR", supCode);
                    output = "Разблокировка поставщика <" + supCode + "> в 1C Trade " + (logUnit == "RU0" ? "" : " Express") + "НЕ ПРОИЗВЕДЕНА [" + logRes + "];";
                }
            }

            return output;

        }

        internal string VisionUserManage(int Mode, string busUnit, string gv_winuser)
        {
            string logUnit;
            string resLine = "";
            int logRes = 0;

            if (busUnit != "A11")
            {
                logUnit = "M11";
            }
            else
            {
                logUnit = "M11";
            }

            using (con = new SqlConnection(connectionString))
            {

                SqlCommand com = new SqlCommand("EXEC M11_4SP_VICLEAR " + Mode, con);

                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);

                    if (dr.HasRows)
                    {
                        foreach (DbDataRecord result in dr)
                        {
                            resLine = result[0].ToString();
                        }
                        logRes = RecordLog("SUNPLUS_01", "VISNOVRD1", gv_winuser, "0", "", "M11", 0, 0, 0, "VIS_OVERD_SUCSF", resLine);
                    }

                }
                catch
                {
                    logRes = RecordLog("SUNPLUS_01", "VISNOVRD1", gv_winuser, "0", "", "M11", 0, 0, 0, "VIS_OVERD_ERROR", resLine);
                    return "Error during process [" + logRes + "];";
                }
            }

            return resLine + "[" + logRes + "];";
        }

        internal string journalShiftLayout(string busUnit, string jNumber, string period, string sunUser, string djRol, string winUser)
        {

            int logRes = 0;
            string resLine = "";

            using (con = new SqlConnection(connectionString))
            {
                SqlCommand com;
                //                SqlCommand com = new SqlCommand("SELECT SUBSTRING(SUN_DATA, 31, 7) AS CURRENT_PERIOD FROM " + businessUnit + "_SSRFMSC WHERE SUN_TB = 'LDG'", con);
                //SqlCommand com = new SqlCommand("EXEC " + busUnit + "_4SP_PSHIFT1 " + jNumber, con);
                if (busUnit == "RU1" || busUnit == "RU0" || busUnit == "RU5" || busUnit == "RU4")
                {
                    com = new SqlCommand("EXEC " + busUnit + "_4SP_PSHIFT1 " + jNumber, con);
                }
                else
                {
                    com = new SqlCommand("EXEC UUU_4SP_PSHIFT1 " + jNumber, con);                   
                }

                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);

                    if (dr.HasRows)
                    {
                        foreach (DbDataRecord result in dr)
                        {
                            resLine = result[0].ToString();
                        }
                        logRes = RecordLog("SUNPLUS_01", "JRNLSHIF11", winUser, "0", "", busUnit, 0, 0, 0, "PER_IS_SHIFT", resLine);
                    }
                }
                catch (Exception e)
                {
                    logRes = RecordLog("SUNPLUS_01", "JRNLSHIF11", winUser, "0", "", busUnit, 0, 0, 0, "PER_ISNOT_SHIFT", resLine);
                    return "Error during process [" + logRes + "]; "+e.Message;
                }

            }
            return resLine + "[" + logRes + "];";
        }


        internal string journalShiftLayout(string busUnit, string jNumber, string period, string sunUser, string djRol, string winUser, string newPer, string reasMov, int adMail)
        {

            int logRes = 0;
            string resLine = "";

            using (con = new SqlConnection(connectionString))
            {
                SqlCommand com;
                //                SqlCommand com = new SqlCommand("SELECT SUBSTRING(SUN_DATA, 31, 7) AS CURRENT_PERIOD FROM " + businessUnit + "_SSRFMSC WHERE SUN_TB = 'LDG'", con);
                //SqlCommand com = new SqlCommand("EXEC " + busUnit + "_4SP_PSHIFT1 " + jNumber, con);
                if (busUnit == "RU1" || busUnit == "RU0" || busUnit == "RU5" || busUnit == "RU4")
                {
                    com = new SqlCommand("EXEC " + busUnit + "_4SP_PSHIFT2 " + jNumber + ", " + newPer, con);
                }
                else
                {
                    com = new SqlCommand("EXEC UUU_4SP_PSHIFT1 " + jNumber, con);
                }

                try
                {
                    con.Open();
                    SqlDataReader dr = com.ExecuteReader(System.Data.CommandBehavior.CloseConnection);

                    if (dr.HasRows)
                    {
                        foreach (DbDataRecord result in dr)
                        {
                            resLine = result[0].ToString();
                        }
                        logRes = RecordLog("SUNPLUS_02", "JRNLSHIF2", winUser, "0", "", busUnit, 0, 0, 0, "PER_IS_SHIFT", resLine);

                        try
                        {
                            string esender = MainForm.GetSender(winUser);
                            MainForm.SendMail(esender, "Инфо: Журнал num. " + jNumber + " " + busUnit + " перенесён из " + period + "->" + newPer + " период.", "Коллеги, добрый день!\n\n Журнал num. " + jNumber + " " + busUnit + " перенесён из " + period + " в " + newPer + " период. \nПричина(" + reasMov + "). \n\nЭто служебное сообщение (и адрес эл.почты), просьба на него не отвечать. \nСпасибо. \n\nС Уважением, Олег Лесницкий.\nmailto:oleg.lesnitsky@dhl.com", busUnit, adMail);
                            MainForm.SendMail("oleg.lesnitsky@dhl.com", "Инфо: Журнал num. " + jNumber + " " + busUnit + " перенесён из " + period + "->" + newPer + " период.", "Коллеги, добрый день!\n\n Журнал num. " + jNumber + " " + busUnit + " перенесён из " + period + " в " + newPer + " период. \nПричина(" + reasMov + "). \n\nЭто служебное сообщение (и адрес эл.почты), просьба на него не отвечать. \nСпасибо. \n\nС Уважением, Олег Лесницкий.\nmailto:oleg.lesnitsky@dhl.com", busUnit, adMail);
                        }
                        catch
                        {
                            output = "Поставщик <" + " " + "> Разблокирован в 1C Trade " + (" " == "RU0" ? "АО" : "ООО") + ". [" + logRes + " unmess];";
                        }

                    }
                }
                catch(Exception e)
                {
                    logRes = RecordLog("SUNPLUS_02", "JRNLSHIF2", winUser, "0", "", busUnit, 0, 0, 0, "PER_ISNOT_SHIFT", resLine);
                    return "Error during process [" + logRes + "]; " + e.Message;
                }

            }
            return resLine + "[" + logRes + "];";
        }
    }
}


