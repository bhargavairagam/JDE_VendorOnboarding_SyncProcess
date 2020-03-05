using LinqToExcel;
using System;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Collections.Generic;
using Ryan.VendorOnboarding.Domain.Entities;
using Ryan.VendorOnboarding.Domain.Concrete;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;
using System.Diagnostics;
using CsvHelper;

namespace JDE_VendorOnboarding_SyncProcess
{
    class Program
    {
        static void Main(string[] args)
        {


          string pathToExcelFile = ConfigurationManager.AppSettings["RTEFilepath"].ToString();
            EventLog v = new EventLog("Application");

            v.Source = "JDE_SYNC_PROCESS";


          //  processCustomers();



            try
            {

                v.WriteEntry("JDE SYNC Process Started", EventLogEntryType.Information, 101, 1);

                foreach (string fi in Directory.GetFiles(pathToExcelFile))
                {
                    // string filePath = pathToExcelFile + fi.
                    string jdeid = "";
                    string name = "";


                    //  bool go = false;
                    using (StreamReader sr = new StreamReader(fi))
                    {


                        string headerLine = sr.ReadLine();
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            List<string> lineValues = line.Split(',').ToList();

                            string type = lineValues[2];
                            string dbaname = lineValues[0];

                            if (lineValues.Count >= 19)
                            {
                                jdeid = lineValues[19];


                                if (!string.IsNullOrEmpty(jdeid))
                                {
                                    jdeid = jdeid.Replace("\"", "");

                                    if (jdeid.Trim() == "")
                                    {
                                        jdeid = lineValues[20];
                                    }
                                }

                                name = lineValues[0];

                                //if (name.Contains("ARC") || name.Contains("Bonnies") || name.Contains("Pullman QOZB") || name.Contains("Z&N Properties ") || name.Contains("Banner Health"))
                                //   // if (name.Contains("ARC"))
                                //    {
                                //        // go = true;
                                //        Console.WriteLine("DBANAME : " + name + " ; JDEID : " + jdeid);
                                //    }
                                //    else
                                //    {
                                //        continue;
                                //    }
                            }
                            else
                            {
                                sr.Close();

                                break;
                            }

                            if (!type.Contains("Suppliers") || string.IsNullOrEmpty(jdeid) || (!string.IsNullOrEmpty(name) && name.Contains("DO NOT")))
                            {
                                sr.Close();

                                break;
                            }
                            else
                            {
                                try
                                {
                                    Console.WriteLine("Processing JDEID : " + jdeid);
                                    v.WriteEntry(" JDEID : " + jdeid, EventLogEntryType.Information, 101, 1);
                                     Starttheprocess(jdeid);


                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("JDE SYNC Process Exception : " + ex.Message);
                                    Console.WriteLine("JDE SYNC Process Exception : " + ex.StackTrace);
                                    v.WriteEntry("JDE SYNC Process Exception : " + ex.Message + " -- " + ex.InnerException, EventLogEntryType.Error, 101, 1);
                                    continue;
                                }

                            }

                        }
                    }

                    // delete files after reading
                    //if (!go)
                    //{
                    //    File.Delete(fi);
                    //}
                    File.Delete(fi);

                    Console.WriteLine(" File Deleted");

                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                v.WriteEntry("Error Desc : " + ex.InnerException, EventLogEntryType.Error, 101, 1);

            }

            Console.WriteLine(" End of Program. Close now.!!! ");
          //  Console.ReadLine();
        }

        private static void processCustomers()
        {

            string pathToExcelFile = ConfigurationManager.AppSettings["RTEFilepath"].ToString();
            EventLog v = new EventLog("Application");

            v.Source = "JDE_SYNC_PROCESS";

            try
            {

                v.WriteEntry("JDE SYNC Process Started", EventLogEntryType.Information, 101, 1);

                foreach (string fi in Directory.GetFiles(pathToExcelFile))
                {
                    // string filePath = pathToExcelFile + fi.
                    string jdeid = "";
                    string name = "";


                    //  bool go = false;
                    using (StreamReader sr = new StreamReader(fi))
                    {


                        string headerLine = sr.ReadLine();
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            List<string> lineValues = line.Split(',').ToList();

                            string type = lineValues[2];
                            string dbaname = lineValues[0];

                            if (lineValues.Count >= 19)
                            {
                                jdeid = lineValues[19];


                                if (!string.IsNullOrEmpty(jdeid))
                                {
                                    jdeid = jdeid.Replace("\"", "");

                                    if (jdeid.Trim() == "")
                                    {
                                        jdeid = lineValues[20];
                                    }
                                }

                                name = lineValues[0];

                          
                            }
                            else
                            {
                                sr.Close();

                                break;
                            }

                            if (!type.Contains("Suppliers") || string.IsNullOrEmpty(jdeid) || (string.IsNullOrEmpty(name) && name.Contains("DO NOT")))
                            {
                                                sr.Close();

                                             break;
                           }



                                try
                                {
                                Console.WriteLine("Processing JDEID : " + jdeid);
                                v.WriteEntry(" JDEID : " + jdeid, EventLogEntryType.Information, 101, 1);
                                Starttheprocess(jdeid);


                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("JDE SYNC Process Exception : " + ex.Message);
                                Console.WriteLine("JDE SYNC Process Exception : " + ex.StackTrace);
                                v.WriteEntry("JDE SYNC Process Exception : " + ex.Message + " -- " + ex.InnerException, EventLogEntryType.Error, 101, 1);
                                continue;
                            }

                            

                        }
                    }

                  
                    File.Delete(fi);

                    Console.WriteLine(" File Deleted");

                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                v.WriteEntry("Error Desc : " + ex.InnerException, EventLogEntryType.Error, 101, 1);

            }


           
        }

       

        private static void Starttheprocess( string vnumber)
        {
            EventLog v = new EventLog("Application");
            //  v.Source = "JDE SYNC PROCESS";
            v.Source = "JDE_SYNC_PROCESS";

            VendorProfile vp = new VendorProfile();
            // string  connec = ConfigurationManager.
            SqlConnection con;
            SqlDataReader reader;
            bool isvendornew = false;
            EFVendorProfileRepositary efrep = new EFVendorProfileRepositary();

            try
            {
                string connec = ConfigurationManager.ConnectionStrings["JdeReplication"].ConnectionString;

                vnumber = vnumber.Replace("\"", "");

                // vp =   efrep.GetVendorDetialsByJdeID(vnumber).Result;

                IEnumerable<VendorProfile> xyx = efrep.GetAllVendorDetails();
                vp = xyx.Where(b => b.JDEVendorID == vnumber).FirstOrDefault();


                if(vp == null)
                {
                    vp = new VendorProfile();
                    vp.VendorGuid = Guid.NewGuid().ToString();
                    vp.SubmittedTime = DateTime.Now;
                    vp.VendorType = "V";
                    vp.SourceType = "JDE_SYNC_PROCESS";
                    vp.STATUSINJDE = "Active";
                    vp.VendorStatus = "Approved";
                    vp.IsEinVerified = "Y";
                    isvendornew = true;
                }

                // get id from RTE generated by JDE.
                vp.JDEVendorID = vnumber;
               

              

                v.WriteEntry("JDEVendorID: " + vnumber , EventLogEntryType.Information, 101, 1);

                string vrnumber = string.Empty;
                int eftb = 0;
                string taxid = string.Empty;
                string personcorp = string.Empty;

                string plinenumber = string.Empty;
                string emailadd = string.Empty;
                if (vp.VRVENDORNUMBER == "0")
                {
                    vp.VRVENDORNUMBER = string.Empty;
                }

                con = new SqlConnection(connec);
                con.Open();


                //  Get VRNumber , EFTB , TAXID , PERSONCORP
                string queryname = BuildNameSqlStringForJDE(vnumber);

                reader = new SqlCommand(queryname, con).ExecuteReader();
                int i = 0;
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        i++;
                        if (i == 1)
                        {
                            vp.VendorDBAName = Convert.ToString(reader["DBANAME"]);
                            if (string.IsNullOrEmpty(vp.VendorDBAName))
                            {
                                return;
                            }
                            vp.VendorLegalName = Convert.ToString(reader["VENDORNAME"].ToString());
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No DBA NAME Found .");
                }

                //  Get VRNumber , EFTB , TAXID , PERSONCORP
                string query1 = BuildSqlStringForJDE(vnumber);

                reader = new SqlCommand(query1, con).ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        eftb = reader.GetInt32(reader.GetOrdinal("ABEFTB"));
                        vrnumber = Convert.ToString(reader.GetInt32(reader.GetOrdinal("VRNUMBER")));

                        vp.VRVENDORNUMBER = vrnumber;
                        //if (vp.VRVENDORNUMBER == "0")
                        //{
                        //    vp.VRVENDORNUMBER = string.Empty;
                        //    vrnumber= string.Empty;
                        //}

                        // vp.VRVENDORNUMBER = vrnumber.ToString();
                        vp.VendorEIN = reader["TAXID"].ToString();
                        vp.PersonCorpCode = reader["PERSONCORP"].ToString();
                        string paym = reader["PayTerm"].ToString().TrimStart('0');
                        if (string.IsNullOrWhiteSpace(vp.PaymentTerm) && paym.Length ==3)
                        {
                            vp.PaymentTerm = paym.Substring(1);
                        }
                        else
                        {
                            vp.PaymentTerm = paym;
                        }

                        v.WriteEntry("EIN: " + vp.VendorEIN , EventLogEntryType.Information, 101, 1);
                    }
                }
                else
                {
                   
                }

                // get address of the company.
                //  Get VRNumber , EFTB , TAXID , PERSONCORP
                StringBuilder getaddress = new StringBuilder();
                getaddress.Append("select ALADD1 as Address1, ALADD2 Address2 ,ALCTY1 City , ");
                getaddress.Append("ALADDS ST , ALADDZ ZIPCODE , ALCTR COUNTRYCODE   from JDE.F0116 where ALAN8 = " + vp.JDEVendorID);
                getaddress.Append("and ALEFTB =  " + eftb.ToString());

                reader = new SqlCommand(getaddress.ToString(), con).ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        vp.VendorAddress1 = reader["Address1"].ToString();
                        vp.VendorAddress2 = reader["Address2"].ToString();
                        vp.VendorCity = reader["City"].ToString();
                        vp.VendorState = reader["ST"].ToString();
                        vp.VendorZipCode = reader["ZIPCODE"].ToString();
                        vp.VendorCountry = reader["COUNTRYCODE"].ToString();

                        v.WriteEntry("EIN: " + vp.VendorAddress1, EventLogEntryType.Information, 101, 1);
                    }
                }
                else
                {
                    Console.WriteLine("No rows found for  company address");
                }

                if (vp.JDEVendorID != vp.VRVENDORNUMBER && (!string.IsNullOrEmpty(vp.VRVENDORNUMBER)))
                {
                   // vp.VRVENDORNUMBER = vrnumber;
                    // get payment address of the company.
                    //  Get address1 , adress2 , city , zipcode and payment address
                    StringBuilder getpayaddress = new StringBuilder();
                    getpayaddress.Append("select ALADD1 as Address1, ALADD2 Address2 ,ALCTY1 City , ");
                    getpayaddress.Append("ALADDS ST , ALADDZ ZIPCODE , ALCTR COUNTRYCODE   from JDE.F0116 where ALAN8 = " + vp.VRVENDORNUMBER.ToString());
                    getpayaddress.Append("and ALEFTB =  " + eftb.ToString());

                    reader = new SqlCommand(getpayaddress.ToString(), con).ExecuteReader();

                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            vp.AIAddress = reader["Address1"].ToString();
                            vp.AIAddress2 = reader["Address2"].ToString();
                            vp.AICity = reader["City"].ToString();
                            vp.AIState = reader["ST"].ToString();
                            vp.AIZip = reader["ZIPCODE"].ToString();
                            vp.AICountry = reader["COUNTRYCODE"].ToString();

                        }
                    }
                    else
                    {
                        Console.WriteLine("No rows found for  payment address");
                    }

                }


                // get phone numbers and line 0 is for company phone details.
                StringBuilder pnumber = new StringBuilder();
                pnumber.Append("select WPPHTP as PTYPE,  WPAR1 as AREACODE , WPPH1 as PHONE from JDE.F0115 where WPIDLN = 0 and  WPAN8 = " + vp.JDEVendorID);

                reader = new SqlCommand(pnumber.ToString(), con).ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        string ptype = reader["PTYPE"].ToString();

                        if (ptype == "FAX")
                        {
                            vp.VFax = reader["AREACODE"].ToString() + "-" + reader["PHONE"].ToString();

                        }
                        else
                        {
                            vp.VPhone = reader["AREACODE"].ToString() + "-" + reader["PHONE"].ToString();
                        }


                    }
                }
                else
                {
                    Console.WriteLine("No rows found for  Phone");
                }


                // get payment address of the company.
                //  Get address1 , adress2 , city , zipcode and payment address
                StringBuilder enumber = new StringBuilder();
                enumber.Append("select EAEMAL    from JDE.F01151 where   EAIDLN = 1   AND EAAN8 = " + vp.JDEVendorID);


                reader = new SqlCommand(enumber.ToString(), con).ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        vp.VendorEmail = reader["EAEMAL"].ToString();

                    }
                }
                else
                {
                   
                }


               

                // save vendor
                var x = efrep.SaveVendorDetials(vp);


                // send email 
                if (isvendornew)
                {
                    string AdminEmails = ConfigurationManager.AppSettings["VOBAdminEmail"].ToString();
                    string[] Aemails = AdminEmails.Split(',');
                    List<string> emails = new List<string>();
                    foreach (string s in Aemails)
                    {
                        emails.Add(s);
                    }

                    string mbody = "New Vendor : " + vp.VendorDBAName + " added to VendorOnboarding by VOB JDE SYNC Process";

                    EmailNotificationUtility.SendEmailToClient(emails, "", vp.VendorDBAName, mbody, mbody);
                }
              


                reader.Close();
                con.Close();
                con.Dispose();


                // save it to vendor profile




            }
            catch (Exception ex)
            {
                v.WriteEntry("Error Desc : " + ex.InnerException, EventLogEntryType.Error, 101, 1);
                v.WriteEntry("Error Desc : " + ex.Message, EventLogEntryType.Error, 101, 1);
                Console.WriteLine(ex.Message);
            }
            finally
            {
               
            }

        }

        private static string BuildSqlStringForJDE(string jDEVendorID)
        {


            StringBuilder sbSql = new StringBuilder();

            sbSql.Append("SELECT ABALPH as DBANAME,ABDC as VENDORNAME, ABAN8 as VENDORNUMBER, ABEFTB , ABAN85 as VRNUMBER , ");
            sbSql.Append("ABTAX As TAXID , ABTAXC as PERSONCORP , A6TRAP as PayTerm from JDE.F0101 ");
            // sbSql.Append("left join JDE.F0111 on WWAN8 = ABAN8 ");
            sbSql.Append("left join JDE.F0401 on A6AN8 = ABAN8 ");
            sbSql.Append("where ABAN8 = " + jDEVendorID + "");


            return sbSql.ToString();
        }


        private static string BuildNameSqlStringForJDE(string jDEVendorID)
        {


            StringBuilder sbSql = new StringBuilder();

            sbSql.Append("SELECT WWALPH as DBANAME, WWMLNM as VENDORNAME ,WWGNNM as FN , WWSRNM as LN , WWATTL as DESIG  ");
            sbSql.Append(" from JDE.F0111 ");
            // sbSql.Append("left join JDE.F0111 on WWAN8 = ABAN8 ");
         
            sbSql.Append(" where CAST( WWAN8  as int) = " + jDEVendorID + "");


            return sbSql.ToString();
        }









    }
     
}
