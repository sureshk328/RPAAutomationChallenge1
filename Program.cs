using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SHDocVw;
using mshtml;
using System.Diagnostics;
using System.Threading;
using System.Data;
using System.Data.OleDb;

namespace GoogleSearch
{
    class Program
    {
        static void Main(string[] args)
        {
            
            //Start IE with google url
            //Process.Start("iexplore", "www.google.com");

            //Wait for page load. This can be avoided but that will extra code
            Thread.Sleep(5000);

            InternetExplorer objInternetExplorer=null;
            HTMLDocument htmlDocument;
            DataTable dtRecords = new DataTable();
            dtRecords = ReadExcelToDT(@"\challenge.xlsx", "Sheet1");
            for (int j = dtRecords.Rows.Count - 1; j >= 0; j--)
            {
                if (dtRecords.Rows[j][1] == DBNull.Value)
                {
                    dtRecords.Rows[j].Delete();
                }
            }
            dtRecords.AcceptChanges();

            int i = 0;
            //Getting InternetExplorer from all open windows
            SHDocVw.ShellWindows shellWindows = new SHDocVw.ShellWindows(); shellWindows = new SHDocVw.ShellWindows();
            foreach (InternetExplorer browser in shellWindows )
            {
                if(browser.LocationURL.ToLower().Contains("www.rpachallenge.com"))
                {
                    objInternetExplorer = browser;
                    break;
                }
            }
            foreach (DataRow dr in dtRecords.Rows)
            {
                if (objInternetExplorer != null)
                {
                    //Get HTML content in HTMLDocument from InternetExplorer object
                    htmlDocument = objInternetExplorer.Document;
                    if(i==0)
                    {
                        HTMLButtonElement startBtn = (HTMLButtonElement)htmlDocument.getElementById("start");
                        startBtn.click();
                        i++;
                    }
                    
                    var divElementsCollection = htmlDocument.getElementsByTagName("div");
                    //js-inputContainer input-group
                    //js-inputContainer input-group
                    if (divElementsCollection != null)
                    {
                        foreach (HTMLDivElement mainDiv in divElementsCollection)
                        {
                            if (mainDiv != null)
                            {
                                if (mainDiv.innerText != null && mainDiv.innerText.Trim().Equals("First Name"))
                                {
                                    var inputCollection = mainDiv.getElementsByTagName("input");
                                    foreach (HTMLInputElement input in inputCollection)
                                    {
                                        if (input != null)
                                        {
                                            input.value = dr["First Name"].ToString();
                                            continue;
                                        }
                                    }

                                }
                                if (mainDiv.innerText != null && mainDiv.innerText.Trim().Equals("Last Name"))
                                {
                                    var inputCollection = mainDiv.getElementsByTagName("input");
                                    foreach (HTMLInputElement input in inputCollection)
                                    {
                                        if (input != null)
                                        {
                                            input.value = dr["Last Name"].ToString(); ;
                                            continue;
                                        }
                                    }

                                }
                                if (mainDiv.innerText != null && mainDiv.innerText.Trim().Equals("Company Name"))
                                {
                                    var inputCollection = mainDiv.getElementsByTagName("input");
                                    foreach (HTMLInputElement input in inputCollection)
                                    {
                                        if (input != null)
                                        {
                                            input.value = dr["Company Name"].ToString(); ;
                                            continue;
                                        }
                                    }

                                }
                                if (mainDiv.innerText != null && mainDiv.innerText.Trim().Equals("Role in Company"))
                                {
                                    var inputCollection = mainDiv.getElementsByTagName("input");
                                    foreach (HTMLInputElement input in inputCollection)
                                    {
                                        if (input != null)
                                        {
                                            input.value = dr["Role in Company"].ToString(); ;
                                            continue;
                                        }
                                    }

                                }
                                if (mainDiv.innerText != null && mainDiv.innerText.Trim().Equals("Address"))
                                {
                                    var inputCollection = mainDiv.getElementsByTagName("input");
                                    foreach (HTMLInputElement input in inputCollection)
                                    {
                                        if (input != null)
                                        {
                                            input.value = dr["Address"].ToString(); ;
                                            continue;
                                        }
                                    }

                                }
                                if (mainDiv.innerText != null && mainDiv.innerText.Trim().Equals("Email"))
                                {
                                    var inputCollection = mainDiv.getElementsByTagName("input");
                                    foreach (HTMLInputElement input in inputCollection)
                                    {
                                        if (input != null)
                                        {
                                            input.value = dr["Email"].ToString(); ;
                                            continue;
                                        }
                                    }

                                }
                                if (mainDiv.innerText != null && mainDiv.innerText.Trim().Equals("Phone Number"))
                                {
                                    var inputCollection = mainDiv.getElementsByTagName("input");
                                    foreach (HTMLInputElement input in inputCollection)
                                    {
                                        if (input != null)
                                        {
                                            input.value = dr["Phone Number"].ToString(); ;
                                            continue;
                                        }
                                    }

                                }
                            }

                        }
                    }
                    var btnSumbitCollection = htmlDocument.getElementsByTagName("input");
                    foreach (HTMLInputElement btnSubmit in btnSumbitCollection)
                    {
                        if (btnSubmit != null && btnSubmit.value.Trim().Equals("Submit"))
                        {
                            btnSubmit.click();
                            
                            break;
                        }
                    }
                    

                }
            }

            Console.ReadLine();


        }

        static DataTable ReadExcelToDT(string filePath, string sheetName)
        {

            string sqlquery = "Select * From ["+ sheetName + "$]";
            DataSet ds = new DataSet();
            string constring = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=YES;\"";
            OleDbConnection con = new OleDbConnection(constring + "");
            OleDbDataAdapter da = new OleDbDataAdapter(sqlquery, con);
            da.Fill(ds);
            DataTable dt = ds.Tables[0];
            return dt;
        }



        
    }
}
