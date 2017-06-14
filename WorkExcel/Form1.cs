using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace WorkExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private ISheet sheet;
        private string pathFile;
        private string pPathFile;
        private void button1_Click(object sender, EventArgs e)
        {
            XSSFWorkbook hsswb;
            List<string> lStr = new List<string>();
            pathFile = @"D:\Kaa_info\ЗАДАНИЯ_ООО_ОНИТ\Тестовые задания\Приложение 1.xlsx";
            if (File.Exists(pathFile))
            {
                pPathFile = pathFile;
            }
            else
            {
                UploadPathExcel upe = new UploadPathExcel();
                pPathFile = upe.getPathtoExcelFile();
            }
            using (FileStream file = new FileStream(pPathFile, FileMode.Open, FileAccess.Read)) 
            {
                hsswb = new XSSFWorkbook(file);
            }
            sheet = hsswb.GetSheet("Лист1");
            ListData ld = new ListData(sheet);
            for (int i = 0; i < ld.getCountRow(); i++)
            {
                // listBox1.Items.Add(ld.getRowfromExcel(i));
                 listBox1.Items.Add(ld.getFullRowfromExcel(i));
            }
            sheet = hsswb.GetSheet("Лист2");
            ListData ld2 = new ListData(sheet);
            listBox2.Items.Add(ld2.getZag(sheet));
            for (int i = 0; i < ld2.getCountRow(); i++)
            {
                //listBox2.Items.Add(ld2.getRowfromExcel(i);
                listBox2.Items.Add(ld2.getFullRowfromExcel(i));
            }
            obrabotkaTwoTable ott = new obrabotkaTwoTable(ld, ld2);
            if (ott.getCountnoEqual() > 0)
            {
                for (int k = 0; k < ott.getCountnoEqual(); k++)
                {
                    listBox3.Items.Add(ott.getdataNoEquals(k));
                }
            } 
        }
    }
    public class ListData 
    {
        private string[] number;
        private string[] data;
        private string[] gaz;
        private string[] neft;
        private string[] condesat;
        private string[] fullstr;
        private int countRow;
        private string numb;
        private string dat;
        private string gz;
        private string nft;
        private string cond;
        public ListData(ISheet sheet)
        {
            countRow = sheet.LastRowNum + 1;
            number = new string[countRow];
            data = new string[countRow];
            gaz = new string[countRow];
            neft = new string[countRow];
            condesat = new string[countRow];
            fullstr = new string[countRow];
            for (int row = 0; row < countRow; row++)
            {
                if (sheet.GetRow(row) != null)
                {

                    numb = String.Format("{0}",sheet.GetRow(row).Cells[0]);
                    dat = String.Format("{0}", sheet.GetRow(row).Cells[1]);
                    gz = String.Format("{0}", sheet.GetRow(row).Cells[2]);
                    nft = String.Format("{0}", sheet.GetRow(row).Cells[3]);
                    cond = String.Format("{0}", sheet.GetRow(row).Cells[4]);
                    number[row] = numb;
                    data[row] = dat;
                    gaz[row] = gz;
                    neft[row] = nft;
                    condesat[row] = cond;
                    fullstr[row] = String.Format("{0};{1};{2};{3};{4}",numb,dat,gz,nft,cond);
                }
            }
        }
        public int getCountRow()
        {
            return countRow;
        }
        public string getRowfromExcel(int n)
        {
            string str = String.Format("{0};{1};{2};{3};{4}", number[n], data[n], gaz[n], neft[n], condesat[n]);
            return str;
        }
        public string getFullRowfromExcel(int n)
        {
            string str = String.Format("{0}", fullstr[n]);
            return str;
        }
        public string getZag(ISheet sheet)
        {
            string str = String.Format("{0};{1};{2};{3};{4}", sheet.GetRow(0).Cells[0], sheet.GetRow(0).Cells[1], sheet.GetRow(0).Cells[2], sheet.GetRow(0).Cells[3], sheet.GetRow(0).Cells[4]);
            return str;
        }
    }
    public class obrabotkaTwoTable
    {
        private List <string> noEqualstr;
        private int countnoEqualstr;
        public obrabotkaTwoTable(ListData obj1, ListData obj2)
        {
            noEqualstr = new List<string>();
            compareTwoTable(obj1,obj2);
        }
        private void compareTwoTable(ListData obj1, ListData obj2)
        {
            int f = 0;
            for (int i = 0; i < obj2.getCountRow(); i++)
            {
                f = 0;
                for (int j = 0; j < obj1.getCountRow(); j++)
                {
                    if (obj2.getFullRowfromExcel(i).Equals(obj1.getFullRowfromExcel(j)))
                    {
                        f++;
                    }
                }
                if (f == 0)
                {
                    int l = i + 1;
                    string itog = String.Format("{0} number stroki - {1}", l, obj2.getFullRowfromExcel(i));
                    noEqualstr.Add(itog);
                }    
            }
            countnoEqualstr = noEqualstr.Count;
        }
        public string getdataNoEquals(int n)
        {
            return noEqualstr.ElementAt<string>(n);
        }
        public int getCountnoEqual()
        {
            return countnoEqualstr;
        }
    }
    public class UploadPathExcel
    {
        private string fullpath;
        private string fileNameExcel;
        private String lineTxt;
        private string [] arstr;
        public UploadPathExcel()
        {
            openfiles();
        }
        private void openfiles()
        {
            fullpath = Application.StartupPath;
            string patt=@"(.*bin).*";
            workRegex wR = new workRegex(patt, fullpath);
            fullpath = wR.getRezStroka();
            fullpath = String.Format(@"{0}\Release\initial.cfg",fullpath);
            using (StreamReader sr = new StreamReader(fullpath, Encoding.Default))
            {
                lineTxt = sr.ReadToEnd();  
            }
            lineTxt = lineTxt.Replace("\r", "");
            arstr = lineTxt.Split('\n');
            patt = "path\\s+=\\s+\"(.*)\"";
            workRegex wr = new workRegex(patt,arstr[0]);
            fileNameExcel = wr.getRezStroka();
        }
        public string getPathtoExcelFile()
        {
            return fileNameExcel;
        }

    }
    public class workRegex
    {
        private string wPatt;
        private string sStr;
        private string rezStr;
        public workRegex(string patt,string str) 
        {
            wPatt = patt;
            sStr = str;
            findValuecCompare();
        }
        private void findValuecCompare()
        {
            Regex rg = new Regex(wPatt);
            Match match = rg.Match(sStr);
            for (int i = 1; i < match.Groups.Count; i++)
            {
                if (match.Groups.Count == 2)
                {
                    rezStr = match.Groups[i].Value;
                    break;
                }
                else 
                {
                    rezStr = match.Groups[i].Value;
                }
            }      
        }
        public string getRezStroka()
        {
            return rezStr;
        }
    }
}
