using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Net.Mail;

namespace MaterialProcurement
{
    struct Values
    {
        public string vqtycal;
        public double vqtysum;
    }
    internal class Program
    {
        static String asp_id = "";                     //產品編號(材料名)
        static String asp_type = "";                   //產品類型
        static String asp_name = "";                   //材料單
        static String asp_um = "";                     //單位
        static Double asp_purprice = 0;                //單價
        static Double asp_standprice = 0;              //
        static String asp_vendorid = "";               //廠商
        static String asp_currency = "";               //幣種
        static String asp_czf = "";                    //參照法
        static Double asp_tjjz = 0;                   //
        static String asp_area = "";                   //
        static Double asp_safeqty = 0;                 //一年內已成交數量
        static Double asp_weight = 0;                  //
        static Double asp_purleadtime = 0;             //
        static Double asp_makeleadtime = 0;            //
        static String asp_location = "";               //
        static Double asp_purchprice = 0;              //數量計算式
        static String asp_purcurrency = "";           //
        static int asp_dummyflag = 0;                  //控管材料
        static String asp_pricecal = "";               //火車頭單價計算式
        static String asp_vendormaterialno = "";       //品號
        static String asp_spec = "";                   //規格
        static String asp_line = "";                   //越南運費check
        static String asp_od = "";                     //審核+越南材料check
        static String asp_multinum = "";               //同品號
        static Double asp_vnweight = 0;                //越南運費-重量
        static Double asp_vnpcs = 0;                   //越南運費-數量
        static String asp_lengum = "";                 //安規線材check
        static String asp_oddate = "";                 //審核日期
        static String asp_oduser = "";                 //審核者
        static String oddate = "";                      //審核日期
        static String oduser = "";                      //審核者

        //static String strSQLConnection = "Data Source =192.168.10.119; Initial Catalog = Test; Persist Security Info=false; User ID = sa; Password = yzf; Max Pool Size=30000;Connection Timeout=1200";//本機資料庫連接
        //static String strSQLConnection = "Data Source =192.168.10.22; Initial Catalog = Test; Persist Security Info=false; User ID = sa; Password = yzf; Max Pool Size=30000;Connection Timeout=1200";//資料庫連接測試區
        static String strSQLConnection = "Data Source =192.168.10.22; Initial Catalog = Price; Persist Security Info=false; User ID = sa; Password = yzf; Max Pool Size=30000;Connection Timeout=1200";//資料庫連接正式區
        static void Main(string[] args)
        {
            try
            {
                //先設定登錄系統名稱=系統更新 FOR 火車頭 輸入者
                string strSQL = $@"update wus set wus_name = '系統更新' from wus where wus_computername = host_name()";
                DoExecuteNonQuery(strSQL);
                //20221107將"填充劑活性鈣 HX-CCR 3000""A5DD34%"變更"中性碳酸鈣XD-2""A5DD34%"
                string[] strID = new string[] { "銅桿OD2.6mm/kg", "PVC粉S-60", "PVC粉S-65", "芯線料/HDPE 9007(3364)/kg", "可塑劑 DOTP", "可塑劑 TOTM", "填充劑中性碳酸鈣XD-2" };
                //string[] strERPID = new string[] { "A6HA01-8888", "A5CB02-5555", "A5CB01-5555", "A5BA09-7777", "A5DA07-5555", "A5DA03-5555", "A5DD34-5555" };
                //20220802 modify by Thomas 品號改成搜前6碼
                string[] strERPID = new string[] { "A6HA01%", "A5CB02%' or PURTD.TD004 like N'A5CB08%", "A5CB01%", "A5BA09%' or PURTD.TD004 like N'A5BA18%", "A5DA07%", "A5DA03%", "A5DB23%" };
                for (int i = 1; i < strID.Length; i++)
                {
                    Double dblSettingPrice = 0;
                    if (i==0)
                    {
                        dblSettingPrice = GetCopperPrice(strID[i], strERPID[i]);    //取得銅設定價格 
                    }
                    else
                    {
                        dblSettingPrice = GetSettingPrice(strID[i], strERPID[i]);   //取得設定價格
                    }
                    
                    GetAsp(strID[i]);       //取得asp參數
                    if (asp_od.Substring(0, 1) == "Y")       //審核日期.審核者
                    {
                        oddate = asp_oddate;
                        oduser = asp_oduser;
                    }
                    else
                    {
                        oddate = "";
                        oduser = "";
                    }
                    asp_purprice = Convert.ToDouble(dblSettingPrice.ToString("#0.##"));      //輸入火車頭的銅價取6位 //20220601修改
                    asp_pricecal = asp_purprice.ToString();                                  //銅價火車頭單價計算式=單價
                    DoUpdate_asp();                     //更新asp資料
                    DoUpdate_asp_od();                  //更新審核+越南材料check
                    DoUpdate_asp_line();                //更新越南運費計重
                    DoUpdate_asp_lengum();              //線材材料UL標記
                    DoCheck_asp_vendormaterialno();     //檢查是否有品號,若有則檢查多品號設定
                    DoCheck_pri_newcostchk();           //檢查材料單是否存在,若不存在則把標記去除
                } 
                string strResult = "材料採購價自動輸入成功 ";
                Mail(strResult);
            }
            catch (Exception ex)
            {
                string strResult = "材料採購價自動輸入失敗 " + ex;
                Mail(strResult);
            }
        }
        private static Double GetCopperPrice(string strID, string strERPID)     //取得銅設定價格
        {
            Double dblAvgPrice = 0;     //下月均價/kg
            Double dblLME = 0;          //LME/NTD
            Double dblSHFE = 0;         //SHFE/NTD
            Double dblAvgMonth = 0;     //當月採購均價
            Double dblAvgTotal = 0;     //(LME+SHFE+當月採購均價)/3
            dblLME = GetCopper("LME");
            dblSHFE = GetCopper("SHFE");
            dblAvgPrice = (dblLME + dblSHFE) / 2;
            dblAvgMonth = GetCopperAvgMonth();

            if (dblAvgMonth == 0)          //如果沒有當月採購均價;(LME+SHFE+當月採購均價)/3=火車頭設定價
            {
                dblAvgTotal = GetAspPrice(strID);
            }
            else
            {
                dblAvgTotal = (dblLME + dblSHFE + dblAvgMonth) / 3;
            }
            if (dblAvgTotal > dblAvgPrice)          //比較得出最終價格
            {
                return dblAvgTotal;
            }
            else
            {
                return dblAvgPrice;
            }
        }
        private static Double GetSettingPrice(string strID, string strERPID)     //取得設定價格
        {
            Double dblAvgMonth = 0;     //當月採購均價
            dblAvgMonth = GetAvgMonth(strERPID);

            if (dblAvgMonth == 0)          //如果沒有當月採購均價;當月採購均價=火車頭設定價
            {
                dblAvgMonth = GetAspPrice(strID);
            }
            return dblAvgMonth;
        }
        private static Double GetCopper(string strCopper)     //取得銅價
        {
            Double dblLME = 0;            //LME價格
            Double dblSHFE = 0;           //SHFE價格
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            SqlDataReader dr;
            string strSQL = $@"select t1.login_date_from as login_date,
                                       t1.lme             as lme_copper,
                                       t1.shfe            as shfe_copper,
                                       t2.working_USD     as working_USD,
                                       t2.working_RMB     as working_RMB
                                from   (select login_date_from,
                                               Avg(lme_copper)  lme,
                                               Avg(shfe_copper) shfe
                                        from   copper_price_detail
                                        where  login_date_from = '{DateTime.Now.ToString("yyyy/MM")}'
                                        group  by login_date_from) t1
                                       left join copper_price t2
                                              on t1.login_date_from = t2.login_date ";
            cmd = conn.CreateCommand();
            cmd.CommandText = strSQL;
            dr = cmd.ExecuteReader();
            do
            {
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        dblLME = (Double)dr["lme_copper"] + (Double)dr["working_USD"];
                        dblSHFE = (Double)dr["shfe_copper"] + (Double)dr["working_RMB"];
                    }
                }
            } while (dr.NextResult());
            conn.Close(); //關閉資料庫連接
            switch (strCopper)
            {
                case "LME":
                    return (Double)(dblLME / 1000 * GetExchangeRate("美金"));
                case "SHFE":
                    return (Double)(dblSHFE / 1000 / 1.11 * GetExchangeRate("人民幣"));
                default:
                    return 0;
            }
        }
        private static Double GetExchangeRate(string strCurrency)     //取得匯率
        {
            Double dblExchangeRate = 0;
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            SqlDataReader dr;
            string strSQL = $@"select cum_convert
                                from   cum
                                where  cum_code='{strCurrency}'
                                and    cum_adddate= (　select max(cum_adddate) from cum where cum_code='{strCurrency}' and　format(cum_adddate,'yyyyMM')<='{ DateTime.Now.ToString("yyyyMM")}') ";
            cmd = conn.CreateCommand();
            cmd.CommandText = strSQL;
            dr = cmd.ExecuteReader();
            do
            {
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        dblExchangeRate = (Double)dr["cum_convert"];
                    }
                }
            } while (dr.NextResult());
            conn.Close(); //關閉資料庫連接
            return dblExchangeRate;
        }
        private static Double GetCopperAvgMonth()     //取得當月採購均價
        {
            Double dblNTD = 0;      //金額(NTD)
            Double dblWeight = 0;   //數量/kg
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            SqlDataReader dr;
            string strSQL = $@"select     PURTC.TC001                                           as 採購單別,
                                           PURTC.TC002                                          as 採購單號,
                                           PURTC.TC003                                          as 採購日期,
                                           PURTC.TC004 + ' ' + PURMA.MA002                      as 廠商名稱,
                                           PURTD.TD004                                          as 品號,
                                           PURTD.TD006                                          as 規格,
                                           PURTC.TC009                                          as 備註,
                                           PURTD.TD010* cum_convert/1.11                        as 採購單價,
                                           PURTD.TD008                                          as 數量合計,
                                           PURTD.TD010 * PURTD.TD008 / 1.11                     as RMBAMT,
                                           (PURTD.TD010 * PURTD.TD008 / 1.11 ) * cum_convert    as NTAMT
                                from       [ERPDB].[MSLCN].dbo.PURTC
                                inner join [ERPDB].[MSLCN].dbo.PURTD
                                on         PURTC.TC001 = PURTD.TD001
                                and        PURTC.TC002 = PURTD.TD002
                                inner join [ERPDB].[MSLCN].dbo.PURMA
                                on         PURMA.MA001 = PURTC.TC004,
                                           cum
                                where      (
                                                      PURTC.TC001 = N'C330' )
                                and        (
                                                      PURTD.TD004 = N'A6HA01-8888' )
                                and        (
                                                      cum_code = '人民幣' )
                                and        (
                                                      PURTC.TC014 = 'Y' )
                                and        cum_adddate= (　select MAX(cum_adddate) from cum where cum_code='人民幣' and　format(cum_adddate,'yyyyMM')<='{DateTime.Now.ToString("yyyyMM")}')
                                and        (
                                                      substring(PURTC.TC003,1,6) = '{DateTime.Now.ToString("yyyyMM")}' )
                                union all
                                select     PURTC.TC001                                      as 採購單別,
                                           PURTC.TC002                                      as 採購單號,
                                           PURTC.TC003                                      as 採購日期,
                                           PURTC.TC004 + ' ' + PURMA.MA002                  as 廠商名稱,
                                           PURTD.TD004                                      as 品號,
                                           PURTD.TD006                                      as 規格,
                                           PURTC.TC009                                      as 備註,
                                           (PURTD.TD010 / 1000)* cum_convert                as 採購單價,
                                           PURTD.TD008*1000                                 as 數量合計,
                                           PURTD.TD010 * PURTD.TD008                        as USAMT,
                                           (PURTD.TD010 * PURTD.TD008) * cum_convert        as NTAMT
                                from       [ERPDB].[TESTMVE1].dbo.PURTC
                                inner join [ERPDB].[TESTMVE1].dbo.PURTD
                                on         PURTC.TC001 = PURTD.TD001
                                and        PURTC.TC002 = PURTD.TD002
                                inner join [ERPDB].[TESTMVE1].dbo.PURMA
                                on         PURMA.MA001 = PURTC.TC004,
                                           cum
                                where      (
                                                      PURTC.TC001 = N'M330' )
                                and        (
                                                      PURTD.TD004 = N'A6HA01-8888' )
                                and        (
                                                      cum_code = '美金' )
                                and        (
                                                      PURTC.TC014 = 'Y' )
                                and        cum_adddate= (　select MAX(cum_adddate) from cum where cum_code='美金' and　format(cum_adddate,'yyyyMM')<='{DateTime.Now.ToString("yyyyMM")}')
                                and        (
                                                      substring(PURTC.TC003,1,6) = '{DateTime.Now.ToString("yyyyMM")}' )
                                order by   採購單號";
            cmd = conn.CreateCommand();
            cmd.CommandText = strSQL;
            dr = cmd.ExecuteReader();
            do
            {
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        dblNTD = dblNTD + (Double)dr["NTAMT"];
                        dblWeight = dblWeight + (Double)(Decimal)dr["數量合計"];
                    }
                }
            } while (dr.NextResult());
            conn.Close(); //關閉資料庫連接
            if (dblWeight == 0)
            {
                return 0;
            }
            else
            {
                return dblNTD / dblWeight;
            }
        }
        private static Double GetAvgMonth(string strERPID)     //取得當月採購均價
        {
            Double dblNTD = 0;      //金額(NTD)
            Double dblWeight = 0;   //數量/kg
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            SqlDataReader dr;
            //20220802 modify by Thomas 品號改成搜前6碼
            string strSQL = $@"select     PURTC.TC001                                           as 採購單別,
                                           PURTC.TC002                                          as 採購單號,
                                           PURTC.TC003                                          as 採購日期,
                                           PURTC.TC004 + ' ' + PURMA.MA002                      as 廠商名稱,
                                           PURTD.TD004                                          as 品號,
                                           PURTD.TD006                                          as 規格,
                                           PURTC.TC009                                          as 備註,
                                           PURTD.TD010* cum_convert/1.11                        as 採購單價,
                                           PURTD.TD008                                          as 數量合計,
                                           PURTD.TD010 * PURTD.TD008 / 1.11                     as RMBAMT,
                                           (PURTD.TD010 * PURTD.TD008 / 1.11 ) * cum_convert    as NTAMT
                                from       [ERPDB].[MSLCN].dbo.PURTC
                                inner join [ERPDB].[MSLCN].dbo.PURTD
                                on         PURTC.TC001 = PURTD.TD001
                                and        PURTC.TC002 = PURTD.TD002
                                inner join [ERPDB].[MSLCN].dbo.PURMA
                                on         PURMA.MA001 = PURTC.TC004,
                                           cum
                                where      (
                                                      PURTC.TC001 = N'C330' )
                                and        (
                                                      PURTD.TD004 like N'{strERPID}' )
                                and        (
                                                      cum_code = '人民幣' )
                                and        (
                                                      PURTC.TC014 = 'Y' )
                                and        cum_adddate= (　select MAX(cum_adddate) from cum where cum_code='人民幣' and　format(cum_adddate,'yyyyMM')<='{DateTime.Now.ToString("yyyyMM")}')
                                and        (
                                                      PURTC.TC003 between '{DateTime.Now.AddMonths(-1).ToString("yyyyMM26")}' and '{DateTime.Now.AddMonths(0).ToString("yyyyMM25")}' )
                                union all
                                select     PURTC.TC001                                      as 採購單別,
                                           PURTC.TC002                                      as 採購單號,
                                           PURTC.TC003                                      as 採購日期,
                                           PURTC.TC004 + ' ' + PURMA.MA002                  as 廠商名稱,
                                           PURTD.TD004                                      as 品號,
                                           PURTD.TD006                                      as 規格,
                                           PURTC.TC009                                      as 備註,
                                           (PURTD.TD010 )* cum_convert                as 採購單價,
                                           PURTD.TD008                                as 數量合計,
                                           PURTD.TD010 * PURTD.TD008                        as USAMT,
                                           (PURTD.TD010 * PURTD.TD008) * cum_convert        as NTAMT
                                from       [ERPDB].[TESTMVE1].dbo.PURTC
                                inner join [ERPDB].[TESTMVE1].dbo.PURTD
                                on         PURTC.TC001 = PURTD.TD001
                                and        PURTC.TC002 = PURTD.TD002
                                inner join [ERPDB].[TESTMVE1].dbo.PURMA
                                on         PURMA.MA001 = PURTC.TC004,
                                           cum
                                where      (
                                                      PURTC.TC001 = N'M330' )
                                and        (
                                                      PURTD.TD004 like N'{strERPID}' )
                                and        (
                                                      cum_code = '美金' )
                                and        (
                                                      PURTC.TC014 = 'Y' )
                                and        cum_adddate= (　select MAX(cum_adddate) from cum where cum_code='美金' and　format(cum_adddate,'yyyyMM')<='{DateTime.Now.ToString("yyyyMM")}')
                                and        (
                                                      PURTC.TC003 between '{DateTime.Now.AddMonths(-1).ToString("yyyyMM26")}' and '{DateTime.Now.AddMonths(0).ToString("yyyyMM25")}' )
                                order by   採購單號";
            cmd = conn.CreateCommand();
            cmd.CommandText = strSQL;
            dr = cmd.ExecuteReader();
            do
            {
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        dblNTD = dblNTD + (Double)dr["NTAMT"];
                        dblWeight = dblWeight + (Double)(Decimal)dr["數量合計"];
                    }
                }
            } while (dr.NextResult());
            conn.Close(); //關閉資料庫連接
            if (dblWeight == 0)
            {
                return 0;
            }
            else
            {
                return dblNTD / dblWeight;
            }
        }
        private static Double GetAspPrice(string strID)     //取得火車頭設定價
        {
            Double dblAsbPrice = 0;
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            SqlDataReader dr;
            string strSQL = $@"select asb_price
                                from   asb
                                where  asb_id = '{strID}'
                                       and asb_changedate = (select Max(asb_changedate)
                                                             from   asb
                                                             where  asb_id = '{strID}'
                                                                    and Format(asb_changedate, 'yyyyMM') <= '{DateTime.Now.ToString("yyyyMM")}')";
            cmd = conn.CreateCommand();
            cmd.CommandText = strSQL;
            dr = cmd.ExecuteReader();
            do
            {
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        dblAsbPrice = (Double)dr["asb_price"];
                    }
                }
            } while (dr.NextResult());
            conn.Close(); //關閉資料庫連接
            if (dblAsbPrice == 0)
            {
                return 0;
            }
            else
            {
                return dblAsbPrice;
            }
        }
        private static void GetAsp(string strID)     //取得asp參數
        {
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            SqlDataReader dr;
            string strSQL = $@"select * from asp where asp_id= '{strID}'";
            cmd = conn.CreateCommand();
            cmd.CommandText = strSQL;
            dr = cmd.ExecuteReader();
            do
            {
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        asp_id = (string)dr["asp_id"];
                        asp_type = (string)dr["asp_type"];
                        asp_name = (string)dr["asp_name"];
                        asp_um = (string)dr["asp_um"];
                        asp_purprice = (Double)dr["asp_purprice"];
                        asp_standprice = (Double)dr["asp_standprice"];
                        asp_vendorid = (string)dr["asp_vendorid"];
                        asp_currency = (string)dr["asp_currency"];
                        asp_czf = (string)dr["asp_czf"];
                        asp_tjjz = (Double)dr["asp_tjjz"];
                        asp_area = (string)dr["asp_area"];
                        asp_safeqty = (Double)dr["asp_safeqty"];
                        asp_weight = (Double)dr["asp_weight"];
                        asp_purleadtime = (Double)dr["asp_purleadtime"];
                        asp_makeleadtime = (Double)dr["asp_makeleadtime"];
                        asp_location = (string)dr["asp_location"];
                        asp_purchprice = (Double)dr["asp_purchprice"];
                        asp_purcurrency = (string)dr["asp_purcurrency"];
                        asp_dummyflag = (int)dr["asp_dummyflag"];
                        asp_pricecal = (string)dr["asp_pricecal"];
                        asp_vendormaterialno = (string)dr["asp_vendormaterialno"];
                        asp_spec = (string)dr["asp_spec"];
                        asp_line = (string)dr["asp_line"];
                        asp_od = (string)dr["asp_od"];
                        asp_multinum = (string)dr["asp_multinum"];
                        asp_vnweight = (Double)dr["asp_vnweight"];
                        asp_vnpcs = (Double)dr["asp_vnpcs"];
                        asp_lengum = (string)dr["asp_lengum"];
                        asp_oddate = ((DateTime)dr["asp_oddate"]).ToString("yyyy/MM/dd");
                        asp_oduser = (string)dr["asp_oduser"];
                    }
                }
            } while (dr.NextResult());
            conn.Close(); //關閉資料庫連接

        }
        private static void DoUpdate_asp()     //更新asp資料
        {
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            string strSQL = $@"exec asp_save1 N'{ asp_id }',
                                            '{ asp_type }',
                                            '{ asp_name }',
                                            '{ asp_um }',
                                            { asp_purprice },
                                            { asp_standprice },
                                            '{ asp_vendorid }',
                                            '{ asp_currency }',
                                            N'{ asp_czf }',
                                            '{ asp_tjjz }',
                                            '{ asp_area }',
                                            { asp_safeqty },
                                            { asp_weight },
                                            { asp_purleadtime },
                                            { asp_makeleadtime },
                                            '',
                                            { asp_purchprice },
                                            '{ asp_purcurrency }',
                                            { asp_dummyflag },
                                            '{ asp_pricecal }',
                                            '{ asp_vendormaterialno }',
                                            N'{ asp_spec }'";
            cmd = new SqlCommand(strSQL, conn);
            //cmd.ResetCommandTimeout();
            //CommandTimeout 重設為30秒
            //怕下列指令執行較長,將他延長設為1200秒
            cmd.CommandTimeout = 1200;
            cmd.ExecuteNonQuery();
            conn.Close(); //關閉資料庫連接
        }
        private static void DoUpdate_asp_od()     //更新審核+越南材料check
        {
            string odtmp;
            if (asp_od.Substring(0, 1) == "Y")
            {
                odtmp = "Y";
            }
            else
            {
                odtmp = " ";
            }
            if (asp_od.Substring(1, 1) == "Y")
            {
                odtmp = odtmp + "Y";
            }
            else
            {
                odtmp = odtmp + " ";
            }
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            string strSQL = $@"update asp 
                                set   asp_od = '{ odtmp }', 
                                      asp_oddate = '{ oddate }', 
                                      asp_oduser = '{ oduser }', 
                                      asp_multinum = N'{ asp_multinum }' 
                                where asp_id = '{ asp_id }'";
            cmd = new SqlCommand(strSQL, conn);
            cmd.ExecuteNonQuery();

            conn.Close(); //關閉資料庫連接
        }
        private static void DoUpdate_asp_line()     //更新越南運費計重
        {
            if (asp_line == "Y")
            {
                //使用asp_line欄位來記錄越南運費計重,asp_vnweight記錄重量,asp_vnpcs記錄數量
                DoUpdate_asp_line_Y();
                //更新計重重量
                DoUpdateWeight();

            }
            else
            {
                //使用asp_line欄位來記錄越南運費計重,asp_vnweight記錄重量,asp_vnpcs記錄數量
                DoUpdate_asp_line_N();
                //更新計重重量
                DoUpdateWeight();
            }
        }
        private static void DoUpdate_asp_line_Y()     //使用asp_line_Y欄位來記錄越南運費計重,asp_vnweight記錄重量,asp_vnpcs記錄數量
        {
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            string strSQL = $@"update asp set asp_line='Y',asp_vnweight= { asp_vnweight} ,asp_vnpcs={asp_vnpcs} where asp_id='{asp_id}'";
            cmd = new SqlCommand(strSQL, conn);
            cmd.ExecuteNonQuery();

            conn.Close(); //關閉資料庫連接
        }
        private static void DoUpdate_asp_line_N()     //使用asp_line_Y欄位來記錄越南運費計重,asp_vnweight記錄重量,asp_vnpcs記錄數量
        {
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            string strSQL = $@"update asp set asp_line='',asp_vnweight= 0 ,asp_vnpcs=0 where asp_id='{asp_id}' and asp_line='Y'";
            cmd = new SqlCommand(strSQL, conn);
            cmd.ExecuteNonQuery();

            conn.Close(); //關閉資料庫連接
        }
        private static void DoUpdateWeight()     //更新計重重量
        {
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            SqlDataReader dr;

            string vid;

            string strSQL = $@"SELECT distinct pri_customerid from pri with (nolock) where pri_part in ( select pub_vnfreight from pub) and pri_customerid in (select pri_customerid from pri where pri_part = '{asp_id}')";

            cmd = conn.CreateCommand();
            cmd.CommandText = strSQL;
            dr = cmd.ExecuteReader();
            do
            {
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        vid = (string)dr["pri_customerid"];

                        strSQL = $@"SELECT pri_perqty,isnull(asp_vnweight,0) 'asp_vnweight',isnull(asp_vnpcs,0) 'asp_vnpcs' from pri with (nolock) left join asp on asp_id=pri_part where pri_customerid ='{ vid }' and asp_line='Y'";
                        Values V = GetValues(strSQL);       //取得vqtycal;vqtysum;


                        strSQL = $@"update pri set pri_perqty=round({ V.vqtysum },6),pri_perqtycal='{ V.vqtycal }',pri_cost=round(pri_tbprice*round({ V.vqtysum},6),6),pri_costflag='Y' where pri_customerid='{ vid}' and pri_part in (select pub_vnfreight from pub)";
                        DoExecuteNonQuery(strSQL);       //執行SQL命令

                    }
                }
            } while (dr.NextResult());
            conn.Close(); //關閉資料庫連接
        }
        private static Values GetValues(string strSQL)      //取得vqtycal;vqtysum;
        {
            Values V = new Values();
            V.vqtycal = "";
            V.vqtysum = 0;
            SqlCommand cmd;
            SqlDataReader dr;
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            cmd = conn.CreateCommand();
            cmd.CommandText = strSQL;
            dr = cmd.ExecuteReader();
            do
            {
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        if (V.vqtycal == "")
                        {
                            if (dr["asp_vnweight"].ToString() == "0")
                            {
                                V.vqtycal = ((double)dr["pri_perqty"]).ToString("0." + new string('#', 339));
                                V.vqtysum = (double)dr["pri_perqty"];
                            }
                            else
                            {
                                V.vqtycal = ((double)dr["asp_vnweight"]).ToString("0." + new string('#', 339)) + "/" + ((double)dr["asp_vnpcs"]).ToString("0." + new string('#', 339)) + "*" + ((double)dr["pri_perqty"]).ToString("0." + new string('#', 339));
                                V.vqtysum = (double)dr["asp_vnweight"] / (double)dr["asp_vnpcs"] * (double)dr["pri_perqty"];
                            }
                        }
                        else
                        {
                            if (dr["asp_vnweight"].ToString() == "0")
                            {
                                V.vqtycal = V.vqtycal + "+" + ((double)dr["pri_perqty"]).ToString("0." + new string('#', 339));
                                V.vqtysum = V.vqtysum + (double)dr["pri_perqty"];
                            }
                            else
                            {
                                V.vqtycal = V.vqtycal + "+" + ((double)dr["asp_vnweight"]).ToString("0." + new string('#', 339)) + "/" + ((double)dr["asp_vnpcs"]).ToString("0." + new string('#', 339)) + "*" + ((double)dr["pri_perqty"]).ToString("0." + new string('#', 339));
                                V.vqtysum = V.vqtysum + (double)dr["asp_vnweight"] / (double)dr["asp_vnpcs"] * (double)dr["pri_perqty"];
                            }
                        }
                    }
                }
            } while (dr.NextResult());
            conn.Close(); //關閉資料庫連接
            return V;
        }
        private static void DoExecuteNonQuery(string strSQL)     //執行SQL命令
        {
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            cmd = new SqlCommand(strSQL, conn);
            cmd.ExecuteNonQuery();
            conn.Close(); //關閉資料庫連接
        }
        private static void DoUpdate_asp_lengum()     //線材材料UL標記
        {
            if (asp_lengum == "Y")
            {
                asp_lengum = "Y";
            }
            else
            {
                asp_lengum = "";
            }
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            string strSQL = $@"update asp set asp_lengum='{asp_lengum}' where asp_id='{asp_id }'";
            cmd = new SqlCommand(strSQL, conn);
            cmd.ExecuteNonQuery();

            conn.Close(); //關閉資料庫連接
        }
        private static void DoCheck_asp_vendormaterialno()     //檢查是否有品號,若有則檢查多品號設定
        {
            Boolean chknum = false;
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            SqlDataReader dr;
            string strSQL = "";
            if (asp_vendormaterialno != "")
            {
                strSQL = $@"SELECT  [aspnum_id] as 材料名,[aspnum_num] AS 品號 From aspnum  where aspnum_id ='{ asp_id }' and aspnum_num='{ asp_vendormaterialno }'";
                cmd = conn.CreateCommand();
                cmd.CommandText = strSQL;
                dr = cmd.ExecuteReader();
                do
                {
                    if (dr.HasRows)
                    {
                        chknum = true;
                    }
                } while (dr.NextResult());
                conn.Close(); //關閉資料庫連接
                if (chknum != true)//'檢查多品號中是否已有資料,若有才處理,若沒有則要先建立好品號資料才能儲存
                {
                    strSQL = $@"INSERT INTO [dbo].[aspnum] 
                                ([aspnum_id],[aspnum_num],[aspnum_modifydate],[aspnum_price],[aspnum_currency],[aspnum_memo],[aspnum_pricecal],[aspnum_vendorid],[aspnum_spec],[aspnum_um]) 
                                VALUES 
                                ('{asp_id}','{ asp_vendormaterialno }','{ DateTime.Now.ToString("yyyy/MM/dd") }','{ asp_purprice } ','{ asp_currency } ','{ asp_czf } ','{ asp_pricecal } ','{ asp_vendorid} ','{ asp_spec } ','{ asp_um } ')";
                    DoExecuteNonQuery(strSQL);       //執行SQL命令
                    //cmd = new SqlCommand(strSQL, conn);
                    //cmd.ExecuteNonQuery();
                }
            }
            else
            {
                strSQL = $@"delete from aspnum  where aspnum_id ='{ asp_id  }'";
                DoExecuteNonQuery(strSQL);       //執行SQL命令
                //cmd = new SqlCommand(strSQL, conn);
                //cmd.ExecuteNonQuery();
            }
            conn.Close(); //關閉資料庫連接
        }
        private static void DoCheck_pri_newcostchk()     //檢查材料單是否存在,若不存在則把標記去除
        {
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            SqlDataReader dr;
            string strSQL = "";
            strSQL = $@"select * from pri where pri_customerid='{asp_id}' and pri_newcostchk like 'Y%'";
            cmd = conn.CreateCommand();
            cmd.CommandText = strSQL;
            dr = cmd.ExecuteReader();

            if (dr.HasRows == false)
            {
                strSQL = $@"update asp set asp_name='' where asp_id='{asp_id}'";
                DoExecuteNonQuery(strSQL);
                //cmd = new SqlCommand(strSQL, conn);
                //cmd.ExecuteNonQuery();
            }
            conn.Close(); //關閉資料庫連接
        }
        private static void Mail(string strResult)   //發送MAIL
        {
            MailMessage MyMail = new MailMessage();
            MyMail.From = new MailAddress("sqluser@msl.com.tw");
            //MyMail.To.Add("收件者Email");加入收件者Email
            //MyMail.To.Add("thomas@msl.com.tw"); //加入收件者Email
            MyMail.To.Add("peggy@msl.com.tw"); //加入收件者Email
            //MyMail.CC.Add("副本的Mail"); //加入副本的Mail
            //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            MyMail.Subject = "材料採購價設定資訊";
            MyMail.Body = strResult + " 結束時間:" + DateTime.Now.ToString(); //設定信件內容
            MyMail.IsBodyHtml = false; //是否使用html格式
            //Attachment attdata = new Attachment(@"D:\" + sDate + @"生產日報.xlsx", MediaTypeNames.Application.Octet);
            //MyMail.Attachments.Add(attdata);
            SmtpClient MySMTP = new SmtpClient("webmail.msl.com.tw", 25);
            MySMTP.Credentials = new System.Net.NetworkCredential("sqluser@msl.com.tw", "msl22995234");
            MySMTP.Send(MyMail);
            MyMail.Dispose(); //釋放資源
        }
    }
}
