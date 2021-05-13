using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Directory = System.IO.Directory;
using DirectoryInfo = System.IO.DirectoryInfo;
using File = System.IO.File;
using FileInfo = System.IO.FileInfo;

using System.Collections;

using System.Xml;
using System.Xml.XPath;

using System.Data;

namespace SunPlus
{
    class Workxml
    {

        //Идем по списку файлов.xml
        //Проверяем Дата = Текущий день?
        //Если равна, проверяем Тег User и протокол содержит Ошибки
        //Если Нужный User протокол
        //Выборка последнего за текущий день



        public DataTable GetXMLDataTableProtocol(string path, string userSun)
        {

            //MessageBox.Show("02 Вход в процедуру " + path + ", " + userSun);

            DataTable XMLErrors = new DataTable();

            //            userSun = "EKZ";
            //            DateTime ActDate = new DateTime(2017, 11, 13);  //текущая дата
            //
            
            //            
            DateTime ActDate = DateTime.Today;  //текущая дата

            //MessageBox.Show("03 Работа с процедурой, массив, аррэйлист");

            FileInfo[] fi; //массив типа FileInfo работа с файловой системой
            ArrayList al = new ArrayList();  //коллекция

            DirectoryInfo di = new DirectoryInfo(path); //работа с каталогами
            //
            fi = di.GetFiles("*.xml"); //получаем список файлов с расширением xml

            //fi = di.GetFiles("SSCLog83019400495310102.xml"); //получаем список файлов с расширением xml
            //MessageBox.Show("1 Кол-во XML " + fi.Length.ToString() + ", пользователь " + userSun.Length.ToString() + ", дата " + ActDate.ToString());

            if (fi.Length != 0) //если длина массива не равна 0
            {
                DataColumn Column1 = XMLErrors.Columns.Add("File Name", typeof(String));
                DataColumn Column2 = XMLErrors.Columns.Add("Journal Line", typeof(String));
                DataColumn Column3 = XMLErrors.Columns.Add("Account Code", typeof(String));
                DataColumn Column4 = XMLErrors.Columns.Add("Transaction Reference", typeof(String));
                //DataColumn Column5 = XMLErrors.Columns.Add("Description", typeof(String));
                DataColumn Column6 = XMLErrors.Columns.Add("Ошибка Error", typeof(String));


                //ищем файл с условиями текущая дата, пользователь
                foreach (FileInfo f in fi)
                {
                    if (f.CreationTime.Date == ActDate.Date)  //текущая дата
                    //if (f.CreationTime.Date != null)
                    {
                        //MessageBox.Show("2 Дата CR совпала " + f.CreationTime.Date.ToString() + "с текущей " + ActDate.ToString() + " по файлу: " + f.Name);
                        //MessageBox.Show("3a Проверка на Fail дала: " + CheckXMLprotocol(f.FullName).ToString() + " по файлу: " + f.Name);

                        //MessageBox.Show("3b Проверка на Fail дала: " + CheckJrnalSrc(f.FullName, userSun).ToString() + " по файлу: " + f.Name);

                        if (CheckXMLprotocol(f.FullName) && CheckJrnalSrc(f.FullName, userSun))  //Есть Ошибка в протоколе (fail) и Есть Journal Srce (JSRCE)
                        {

                            XMLErrors = GetXMLData(f, ref XMLErrors);
                            //MessageBox.Show("3 Найден файл " + f.Name + " " + XMLErrors.Rows.Count);

                        }

                    }

                }

                //string[] files = (string[])al.ToArray(typeof(string));
                //MessageBox.Show("04a завершение работы с процедурой, возврат: " + XMLErrors.Rows.Count);

                return XMLErrors;

            }

            //MessageBox.Show("04b завершение работы с процедурой");

            return XMLErrors;

        }


        public static DataTable GetXMLData(FileInfo curXMLFile, ref DataTable XMLDataErrors)
        {
            
            //MessageBox.Show("Вошли в GetXMLData");

            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(curXMLFile.FullName);
            
            //MessageBox.Show("Загрузили файл в XDoc");
            
            XmlElement xRoot = xDoc.DocumentElement;

            XmlNodeList childnodes = xRoot.SelectNodes("//Line[@status='fail']");

            XmlNodeList fieldBusUnitText = xRoot.SelectNodes("//BusinessUnit");

            int specError = 0;
            DataRow newRow = null;

            //MessageBox.Show("Загрузили файл в XDoc" + childnodes.Count.ToString());
            
            //загрузили в файл XDoc 498

            if (childnodes.Count > 0)
            {

                foreach (XmlNode n in childnodes)
                {
                    XmlNode fieldUserText = n.SelectSingleNode("Messages/Message/UserText");


                    if (!fieldUserText.InnerText.Contains("rejected"))
                    {

                        newRow = XMLDataErrors.NewRow();

                        XmlNode field13 = n.SelectSingleNode("JournalLineNumber");
                        XmlNode field1 = n.SelectSingleNode("AccountCode");
                        XmlNode field2 = n.SelectSingleNode("TransactionReference");
                        //XmlNode field10 = n.SelectSingleNode("Description");

                        specError = 0;

                        if (fieldUserText.InnerText.Contains("Abort() method called from SASI script"))
                        {
                            specError = 1;

                            newRow[0] = curXMLFile.Name;
                            newRow[1] = "All";
                            newRow[2] = "All";
                            newRow[3] = "All";
                            newRow[4] = "Единичные ошибки не обнаружены. Загружаемый файл не в балансе";


                        }
                        else
                        {
                            newRow[0] = curXMLFile.Name;
                            newRow[1] = field13.InnerText;

                            try
                            {
                                newRow[2] = field1.InnerText;
                            }
                            catch
                            {
                                newRow[2] = "<Не указан>";
                            }

                            newRow[3] = field2.InnerText;
                            newRow[4] = replaceErrorDescriptn(fieldUserText.InnerText.TrimEnd());

                        }

                        if (specError == 0)
                        XMLDataErrors.Rows.Add(newRow);


                    }
                }

            }

            else
             {

                newRow = XMLDataErrors.NewRow();

                newRow[0] = "Файлы не найдены!";
                newRow[1] = "";
                newRow[2] = "";
                newRow[3] = "";
                newRow[4] = "Error(s) have(s) not found...";
                XMLDataErrors.Rows.Add(newRow);

            }
                
                
                if (specError == 1 && XMLDataErrors.Rows.Count == 0)
                    XMLDataErrors.Rows.Add(newRow);


            return XMLDataErrors;
        }


        private static string replaceErrorDescriptn(string strErrorMessage)
        {
            strErrorMessage = strErrorMessage.Replace("Missing", "Не указано");
            strErrorMessage = strErrorMessage.Replace("Unknown", "Неизвестен");
            strErrorMessage = strErrorMessage.Replace("Must give a value for", "Необходимо указать");

            return strErrorMessage;
        }


        private bool CheckXMLprotocol(string prmFullFileName)
        {
            XmlDocument xDoc = new XmlDocument();

            try { 
                xDoc.Load(prmFullFileName);
                XmlElement xRoot = xDoc.DocumentElement;

                XmlNodeList childnodes = xRoot.SelectNodes("//Line[@status='fail']");

                return childnodes.Count > 0 ? true : false;
            }
            catch
            { 
                return false;
            }

        }


        private bool CheckJrnalSrc(string prmFullFileName, string prmUserSun)
        {
            bool result = false;

            //MessageBox.Show("Вошли в процедуру CheckJrnalSrc");
            

            // Создание XPath документа.
            var document = new XPathDocument(prmFullFileName);
            XPathNavigator navigator = document.CreateNavigator();

            // Прямой запрос XPath.
            XPathNodeIterator iterator1 = navigator.Select("SSC/User/Name");

            //MessageBox.Show("InnerXml " + iterator1.Current.InnerXml + ", параметр пользователя" + prmUserSun);

            while (iterator1.MoveNext())
                result = iterator1.Current.InnerXml == prmUserSun ? true : false;
            //result = iterator1.Current.InnerXml == "IKO" ? true : false;

            //MessageBox.Show("Результат CheckJrnalSrc: " + result);

            return result;
        }



        //Тестовый вывод Ошибок (неактивно!!!)
        public string GetXMLFilesInfo(string path, string userSun)
        {
            string strPoolErrors = "";

            //DateTime ActDate = DateTime.Today;  //текущая дата
            DateTime ActDate = new DateTime(2017, 10, 4);  //текущая дата

            FileInfo[] fi; //массив типа FileInfo работа с файловой системой
            ArrayList al = new ArrayList();  //коллекция

            DirectoryInfo di = new DirectoryInfo(path); //работа с каталогами
            fi = di.GetFiles("*.xml"); //получаем список файлов с расширением xml

            if (fi.Length != 0) //если длина массива не равна 0
            {
                //ищем файл с условиями текущая дата, пользователь
                foreach (FileInfo f in fi)
                {
                    if (f.CreationTime.Date == ActDate.Date)  //текущая дата
                    //if (f.CreationTime.Date != null)
                    {
                        if (CheckXMLprotocol(f.FullName) && CheckJrnalSrc(f.FullName, userSun))  //Есть Ошибка в протоколе (fail) и Есть Journal Srce (JSRCE)
                        {
                            strPoolErrors = strPoolErrors + getSunXMLErrors(f);
                            //al.Add(f.Name);                        
                        }

                    }

                }

                //string[] files = (string[])al.ToArray(typeof(string));

                return strPoolErrors;

            }

            return "Журналы не найдены!";

        }

        //Тестовый вывод Ошибок (неактивно!!!)
        private string getSunXMLErrors(FileInfo curXMLFile)
        {

            //параметр полное имя файла
            //Console.WriteLine(new string('-', 20));

            string strErrorLayout = "";

            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(curXMLFile.FullName);
            XmlElement xRoot = xDoc.DocumentElement;

            XmlNodeList childnodes = xRoot.SelectNodes("//Line[@status='fail']");

            XmlNodeList fieldBusUnitText = xRoot.SelectNodes("//BusinessUnit");

            if (childnodes.Count > 0)
            {
                strErrorLayout += "## SourceFile: " + curXMLFile.Name + "\n\n";

                foreach (XmlNode n in childnodes)
                {
                    XmlNode fieldUserText = n.SelectSingleNode("Messages/Message/UserText");

                    //if (fieldtest.InnerText == "185")
                    //    strErrorLayout = strErrorLayout;

                    if (!fieldUserText.InnerText.Contains("rejected"))
                    {
                        XmlNode field13 = n.SelectSingleNode("JournalLineNumber");
                        strErrorLayout += "[Journal Line]: " + field13.InnerText;

                        foreach (XmlNode b in fieldBusUnitText)
                            strErrorLayout += " [Business Unit]:" + b.InnerText;

                        XmlNode field1 = n.SelectSingleNode("AccountCode");
                        strErrorLayout += "  [AccountCode]: " + field1.InnerText;

                        XmlNode field2 = n.SelectSingleNode("TransactionReference");
                        strErrorLayout += "  [Transaction Reference]: " + field2.InnerText + "\n";

                        XmlNode field10 = n.SelectSingleNode("Description");

                        if (fieldUserText.InnerText.Contains("SASI") || fieldUserText.InnerText.Contains("imbalance"))
                        {
                            strErrorLayout += "[Description]: " + "The journal imbalanced!              <<< ОШИБКА >>>: " + "\n";
                            break;
                        }
                        else
                        {
                            strErrorLayout += "[Description]: " + field10.InnerText + "\n";
                            strErrorLayout += "              <<< ОШИБКА >>>: " + replaceErrorDescriptn(fieldUserText.InnerText.TrimEnd()) + "\n";
                            strErrorLayout += "\n";
                        }

                    }
                }

            }

            else
                strErrorLayout += "Error(s) have(s) not found... \n";

            return strErrorLayout += new string('-', 185) + "\n"; ;

        }


    }


}
