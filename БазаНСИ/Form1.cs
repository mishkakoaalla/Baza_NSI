﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KompasAPI7;
using Kompas6Constants;
using Kompas6API5;
using KAPITypes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using reference = System.Int32;
using System.Text.RegularExpressions;





namespace БазаНСИ
{

    public partial class Form1 : Form
    {
        
        private KompasObject kompas;
        private IApplication appl;         // Интерфейс приложения

        private bool tip_obj_base = false;  // Базовая ли спецификация или нет (по умолчание "нет")
        public int num ; // Количеств острок в спецификации
        public int nomer_razdela = 0;
        public int nomer_razdela_base = 0;
        public string obrez_SB;
        public int stroka_base;
        //public Excel.Worksheet sheet;
        




        public ISpecificationCommentObject obj; //Объект спецификации
        public ISpecificationBaseObject obj_base; //Базовый Объект спецификации
        public ISpecificationColumns oC ;

        // public ыекш OBJ;


        string aa, bb, cc;

        List<string> path = new List<string>();
        List<string> path_name = new List<string>();
        List<string> spisok_sb = new List<string>();
        List<object> spisok_doc = new List<object>();
        spec_stroka[] Sps = new spec_stroka[1500];
        Baza[] Base = new Baza[1500];



        void close_file()
        {
            Console.WriteLine("ЗАКРЫТИЕ ДОКУМЕНТЫ");
            ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();
            if (doc != null)
                doc.ksCloseDocument();
        }



        public Form1()
        {
            //label1.Text = "Перенеси файлы сюда";
            InitializeComponent();
            GetKompas();
            //START();
            Console.WindowWidth = 200;
            //Console.BufferWidth = 1000;

            

        }


        public string ObrezFileName(string pp)
        {
                string obrez_do_naimen = Path.GetFileName(pp).Remove(Path.GetFileName(pp).IndexOf(" "));
                string subString = "СБ";
                int indexOfSubstring = obrez_do_naimen.IndexOf(subString);
                if (indexOfSubstring > 0)
                {
                    obrez_SB = obrez_do_naimen.Substring(0, obrez_do_naimen.Length - 2);
                }
                else
                {
                    obrez_SB = obrez_do_naimen;
                }

                return obrez_SB;           
        }


        public string ObrezSB(string ss)
        {
            string obrez_do_naimen = ss;
            string subString = "СБ";
            int indexOfSubstring = obrez_do_naimen.IndexOf(subString);
            if (indexOfSubstring > 0)
            {
                obrez_SB = obrez_do_naimen.Substring(0, obrez_do_naimen.Length - 2);
            }
            else
            {
                obrez_SB = obrez_do_naimen;
            }

            return obrez_SB;

        }

        public void START()
        {
            

            


            int stolb = 1;//   A - B - C - D     //Cells(5, 3) = C5
            int stroka = 1; //  1-2-3  
            int nomer_Sps = 1;




            




            spec_stroka[] Spisok_dok = new spec_stroka[1500];



            //for (int i1 = 0; i1 <1000;i1++)
            //{
            //    Sps[i1] = new spec_stroka();
            //}

            



            Console.WriteLine("Количество документов = " + path.Count);
            for (int i = 0; i < path.Count; i++)
            {


                IKompasDocument doc = appl.Documents.Open(path[i], false, false);// Получаем интерфейс активного документа 2D в API7
                Console.WriteLine("Получение спецификации из документа № - " + Convert.ToInt32(i + 1));
                
                SpecificationDescription Specification_Descriptions = doc.SpecificationDescriptions.Active;




                if (Specification_Descriptions != null)
                {

                    ISpecificationCommentObjects SpcObjects = Specification_Descriptions.CommentObjects;
                    ISpecificationBaseObjects SpcObjectsBase = Specification_Descriptions.BaseObjects;

                    
                    //Console.WriteLine(" ВСПОМОГАТЕЛЬНЫЕ объекты " + SpcObjects.Count);
                    //Console.WriteLine(" Базовые объекты " + SpcObjectsBase.Count);
                    int kol_com = SpcObjects.Count;
                    int kol_base = SpcObjectsBase.Count;

                    if (kol_com == 0 & kol_base > 0)
                    {
                        tip_obj_base = true;
                        num = kol_base;
                    }
                    if (kol_base == 0 & kol_com > 0)
                    {
                        tip_obj_base = false;
                        num = kol_com;
                    }




                    ISpecificationObject Specification_Object;
                    ISpecificationColumns Specification_Columns;
                    ISpecificationColumn Specification_Column;

                    // Начало блока вспомагательных объктов 
                    for (int SD = 0; SD < num; SD++)
                    {
                        Console.WriteLine("----- Строка " + (SD + 1) + "   ---- ");
                        Console.WriteLine("");
                        //var ww = SpcObjects[SD];
                        if (!tip_obj_base)
                        {
                            ISpecificationCommentObject obj = SpcObjects[SD];
                            var OBJ = obj;
                            ISpecificationColumns oC = obj.Columns;
                            int qq = obj.Section;
                            Console.WriteLine("!!!!СЕКЦИЯ " + (qq) + "   !!!!!");
                            Specification_Object = obj;
                            nomer_razdela = qq;
                            //5 - документация   10 - комплексы    15 - сборочные единицы   20 - детали    25 - стандартные изделия   30 - прочие изделия    35 - материалы      40 - комплекты
                        }
                        else
                        {
                            ISpecificationBaseObject obj_base = SpcObjectsBase[SD];
                            var OBJ = obj_base;
                            ISpecificationColumns oC = obj_base.Columns;
                            int qq = obj_base.Section;
                            Console.WriteLine("!!!!СЕКЦИЯ " + (qq) + "   !!!!!");
                            Specification_Object = obj_base;
                            nomer_razdela_base = qq;
                        }

                        if(((nomer_razdela == 5) | (nomer_razdela_base == 5)) & (i!=0))
                        {
                            Console.WriteLine("----------------Пропуск раздела документация-------------------");
                                continue;                              
                        }
                        else
                        {

                            Sps[nomer_Sps] = new spec_stroka();
                            Sps[nomer_Sps].doc_name = ObrezFileName(path[i]);
                            Sps[nomer_Sps].sortir = false;


                            if (nomer_razdela == 5 | nomer_razdela_base == 5)
                            {
                                Sps[nomer_Sps].tip_stroki = "CБ";
                                Console.WriteLine("----------------Запись раздела документация  ГЛАВНОГО ДОКУМЕНТА-------------------");
                            }

                            if (nomer_razdela == 15 | nomer_razdela_base == 15 )
                            {
                                Sps[nomer_Sps].tip_stroki = "CБ";
                            }
                            if (nomer_razdela == 25 | nomer_razdela_base == 25)
                            {
                                Sps[nomer_Sps].tip_stroki = "СТ";
                            }
                            if (nomer_razdela == 30 | nomer_razdela_base == 30)
                            {
                                Sps[nomer_Sps].tip_stroki = "П";
                            }
                            if (nomer_razdela == 35 | nomer_razdela_base == 35)
                            {
                                Sps[nomer_Sps].tip_stroki = "М";
                            }






                            Specification_Columns = Specification_Object.Columns;

                            for (int bCol = 0; bCol < Specification_Columns.Count; bCol++)
                            {
                                Specification_Column = Specification_Columns[bCol];
                                var st = Specification_Column.Text.Str;
                                Console.WriteLine("Столбец " + (bCol + 1) + " - " + st);

                                //Заполнение      /////////////////////////////////////                            ///////////////////////////
                                //sheet.Cells[stroka, stolb] = st;
                                if (nomer_razdela == 20 | nomer_razdela_base == 20)
                                {
                                    if (bCol == 0)
                                    {
                                        if (st != "БЧ")
                                        {
                                            Sps[nomer_Sps].tip_stroki = "Д";
                                        }
                                        else
                                        {
                                            Sps[nomer_Sps].tip_stroki = "БЧ";
                                        }
                                    }
                                    
                                }

                                switch (bCol)
                                {
                                    case 0:
                                        Sps[nomer_Sps].format = st;


                                        break;
                                    case 2:
                                        Sps[nomer_Sps].poz = st;
                                        break;
                                    case 3:
                                        //Sps[nomer_Sps].obozn = st;
                                        Sps[nomer_Sps].obozn = ObrezSB(st);
                                        break;
                                    case 4:
                                        //Sps[nomer_Sps].naimen = st;
                                        if (nomer_razdela == 5 | nomer_razdela_base == 5)
                                        {
                                            Sps[nomer_Sps].naimen = textBox4.Text;
                                        }
                                        else
                                        {

                                            Sps[nomer_Sps].naimen = Regex.Replace(st, @"[ \n]", " ");
                                        }

                                        break;
                                    case 5:
                                        Sps[nomer_Sps].kol = st;
                                        break;
                                    case 6:
                                        Sps[nomer_Sps].prim = st;
                                        break;

                                }

                                //stolb += 1;
                            }

                            Console.WriteLine("----- Конец cтроки ---- ");
                           // stolb = 1;
                            //stroka += 1;
                            nomer_Sps += 1;
                            nomer_razdela = 0;
                            nomer_razdela_base = 0;

                        }
                    }
                
                    

                                                                                                                       
                    if (doc != null)
                    {
                        doc.Close(0); //Закрыть документ
                    }
                }
                else
                {
                    Console.WriteLine("Пропущен документ (документ не спецификации, и не на чертеже)");
                    if (doc != null)
                    {

                        ksDocument2D docD = (ksDocument2D)kompas.ActiveDocument2D();
                        ksStamp stamp = (ksStamp)docD.GetStamp();


                        LayoutSheets _ls = doc.LayoutSheets;
                        LayoutSheet LS = _ls.ItemByNumber[1];
                        IStamp isamp = LS.Stamp;                     
                        IText qq = isamp.Text[3];
                        IText ww = isamp.Text[2];

                        Console.WriteLine("ШТАМП Материал -------------  " + qq.Str);
                        Console.WriteLine("ШТАМП Обозначение -------------  " + ww.Str);

                        Sps[nomer_Sps] = new spec_stroka();
                        Sps[nomer_Sps].doc_name = ObrezFileName(path[i]);
                        Sps[nomer_Sps].obozn = ww.Str;
                        Sps[nomer_Sps].tip_stroki = "Материал из детали";
                        Sps[nomer_Sps].material = qq.Str;
                        Sps[nomer_Sps].sortir = false;

                        nomer_Sps += 1;
                        doc.Close(0); //Закрыть документ

                    }
                }

                doc.Close(0);
                Console.WriteLine("");
                Console.WriteLine("-Проверка-");






            }



            ////////////////////////////////////////////////////////////////////////////////////////

            //sheet.Columns["D:D"].ColumnWidth = 16.0;
            // sheet.Columns["E:E"].ColumnWidth = 25.0;


            try
            {

                    //ex.Application.ActiveWorkbook.SaveAs("D:\\1111111111111111111.xlsx", Type.Missing,
                   // Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                    //Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                
            }
            catch 
            {

                MessageBox.Show("Сперва закройте БазуНСИ");
            }




            ///////////////////////////////////////////
            /*

            workBook = null;
            sheet = null;
            //ex.Quit();
            ex = null;
            GC.Collect();

             */
            


    
            for (int i2 = 0; i2 < 1500; i2++)
            {
                if (Sps[i2] != null)
                {
                    if (Sps[i2].sortir != null)
                    {
                        
                        Sps[i2].GetInfoSst();
                        
                    }
                }
            }

            


        }
    





    







        void GetKompas()
        {
            try
            {            

                kompas = (KompasObject)System.Runtime.InteropServices.Marshal.GetActiveObject("kompas.application.5");
                appl = (IApplication)kompas.ksGetApplication7();
                //MessageBox.Show("Подключение установлено");
                Console.WriteLine("Подключение установлено");
                appl.KompasError.Clear();

            }
            catch
            {
                //MessageBox.Show("Компас не запущен - ЗАПУСКАЕМ ");
                Console.WriteLine("Компас не запущен - ЗАПУСКАЕМ");
                Type t = Type.GetTypeFromProgID("KOMPAS.Application.5");
                kompas = (KompasObject)Activator.CreateInstance(t);
                kompas = (KompasObject)System.Runtime.InteropServices.Marshal.GetActiveObject("kompas.application.5");
                appl = (IApplication)kompas.ksGetApplication7();
                kompas.Visible = true;  //  
                appl.KompasError.Clear();
                //kompas.ActivateControllerAPI();
            }
        }

        void button1_Click(object sender, EventArgs e)
        {
            ksDocument2D doc = (ksDocument2D)kompas.Document2D();
            doc.ksOpenDocument(path[0], false);
            {
                ksStamp stamp = (ksStamp)doc.GetStamp();
                if (stamp != null && stamp.ksOpenStamp() == 1)
                {
                }
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            START();  
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            path.Clear();
            path_name.Clear();
            label1.Text = "Перенеси файлы сюда";
        }


        void panel1_DragEnter(object sender, DragEventArgs e)
        {
           if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            int ex_stroka = 1;
            int ex_stolb = 1;

            Excel.Application ex = new Excel.Application();
            ex.Visible = true;
            ex.SheetsInNewWorkbook = 2;
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
            ex.DisplayAlerts = false;
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
            sheet.Name = "База НСИ";

            for (int i2 = 0; i2 < Base.Length; i2++)
            {
                if (Base[i2] != null)
                {
                    if (Base[i2].ispolnitel != null)
                    {
                        for (int i3 = 0; i3 < 11; i3++)
                        {
                            switch (i3)
                            {
                                case 2:
                                    sheet.Cells[ex_stroka, ex_stolb] = Base[i2].vhodimost;
                                    ex_stolb += 1;
                                    break;

                                case 3:
                                    sheet.Cells[ex_stroka, ex_stolb] = Base[i2].naimen;
                                    ex_stolb += 1;
                                    break;
                                case 4:
                                    sheet.Cells[ex_stroka, ex_stolb] = Base[i2].obozn;
                                    ex_stolb += 1;
                                    break;
                                case 5:
                                    sheet.Cells[ex_stroka, ex_stolb] = Base[i2].type_cher;
                                    ex_stolb += 1;
                                    break;
                                case 6:
                                    sheet.Cells[ex_stroka, ex_stolb] = Base[i2].ispolnitel;
                                    ex_stolb += 1;
                                    break;
                                case 7:
                                    sheet.Cells[ex_stroka, ex_stolb] = Base[i2].gotovnost;
                                    ex_stolb += 1;
                                    break;
                                case 8:
                                    sheet.Cells[ex_stroka, ex_stolb] = Base[i2].kuda_vhodit;
                                    ex_stolb += 4;
                                    break;
                                case 9:
                                    sheet.Cells[ex_stroka, ex_stolb] = Base[i2].material;
                                    ex_stolb += 6;
                                    break;
                                case 10:
                                    sheet.Cells[ex_stroka, ex_stolb] = Base[i2].kol;
                                    ex_stolb += 1;
                                    break;
                            }















                            //Base[i2].DrawBase();
                            
                            
                        }
                        ex_stolb = 1;




                    }
                    ex_stroka += 1;
                }
                
            }
            sheet.Columns["B:B"].ColumnWidth = 45.0;
            sheet.Columns["C:C"].ColumnWidth = 20.0;
            sheet.Columns["K:K"].ColumnWidth = 40.0;


        }

        private void button5_Click(object sender, EventArgs e)
        {
            for (int i2 = 0; i2 < 1500; i2++)
            {
                if (Sps[i2] != null)
                {
                    if (Sps[i2].sortir != null)
                    {

                        Sps[i2].GetInfoSst();

                    }
                }
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        void panel1_DragDrop(object sender, DragEventArgs e)
        {
            var allowedExtensions = new[] { ".cdw", ".spw"};


            foreach (string obj in (string[])e.Data.GetData(DataFormats.FileDrop))
                if (Directory.Exists(obj))
                {
                    // path.AddRange(Directory.GetFiles(obj, "*.*", SearchOption.AllDirectories)
                    //.Where(f=> f.EndsWith(".cdw")|| f.EndsWith(".spw")).ToArray()                

                   // );
                    //MessageBox.Show("Не вабраны файлы с расширением  .cdw или .spw");

                }
                else
                {
                    string q = Path.GetFileName(obj);
                    string w = Path.GetExtension(obj);

                    if (w == ".cdw" || w == ".spw")
                    {
                        path.Add(obj);
                        path_name.Add(q);
                        Console.WriteLine("Докkkkkkkkkkkkkkkkkkkkkkkk " + w);
                    }
                }
            label1.Text = string.Join("\r\n", path_name);

            
           // label1.Text += file + "\n";
            

            
        }





        public void button3_Click(object sender, EventArgs e)
        {

            Console.WriteLine("-------------------------\n-------------------------\n-------------------------");

            for (int i = 0; i < Sps.Length; i++)
            {

                if (Sps[i] != null)
                {
                    Base[0] = new Baza();
                    Base[0].naimen = Sps[i].naimen;
                    Base[0].obozn = Sps[i].obozn;
                    Base[0].ispolnitel = textBox1.Text;
                    Base[0].vhodimost = textBox3.Text;
                    Base[0].gotovnost = textBox2.Text;
                    Base[0].type_cher = "СБ";
                    Sps[i].sortir = true;
                    stroka_base = 1;
                    break;
                    
                }
            }



            for (int i=0; i< Sps.Length; i++)
            {

                if (Sps[i] != null)
                {
                    if ((Sps[i].tip_stroki == "CБ") & (Sps[i].sortir == false))
                    {
                        Base[stroka_base] = new Baza();
                        Base[stroka_base].obozn = Sps[i].obozn;
                        Base[stroka_base].naimen = Sps[i].naimen;
                        Base[stroka_base].ispolnitel = textBox1.Text;
                        Base[stroka_base].type_cher = "СБ";
                        Base[stroka_base].vhodimost = textBox3.Text;
                        Base[stroka_base].gotovnost = textBox2.Text;
                        Base[stroka_base].kol = Sps[i].kol;


                        Sps[i].sortir = true;


                        stroka_base += 1;
                        Console.WriteLine("Это сборка " + i);
                        for (int i2 = 0; i2 < Sps.Length; i2++)
                        {
                            if (Sps[i2] != null)
                            {
                                if (Sps[i2].doc_name == Sps[i].obozn)
                                {
                                    Base[stroka_base] = new Baza();


                                    Base[stroka_base].ispolnitel = textBox1.Text;
                                    Base[stroka_base].type_cher = Sps[i2].tip_stroki;

                                    Base[stroka_base].kol = Sps[i2].kol;
                                    Base[stroka_base].gotovnost= textBox2.Text;


                                    if ((Sps[i2].tip_stroki != "М") & (Sps[i2].tip_stroki != "СТ") & (Sps[i2].tip_stroki != "П"))
                                         {
                                        Base[stroka_base].obozn = Sps[i2].obozn;
                                        Base[stroka_base].naimen = Sps[i2].naimen;

                                    }
                                    else
                                    {
                                        if (Sps[i2].tip_stroki != "М")
                                        {
                                            Base[stroka_base].material = Sps[i2].naimen;
                                        }
                                        if (Sps[i2].tip_stroki != "СТ")
                                        {
                                            Base[stroka_base].material = Sps[i2].naimen;
                                        }
                                    }



                                    Sps[i2].sortir = true;
                                    stroka_base += 1;

                                }

                            }
                        }
                    }
                    if ((Sps[i].tip_stroki == "Д") & (Sps[i].sortir == false))
                    {
                        Base[stroka_base] = new Baza();
                        Base[stroka_base].obozn = Sps[i].obozn;
                        Base[stroka_base].naimen = Sps[i].naimen;
                        Base[stroka_base].type_cher = "Д";
                        Base[stroka_base].ispolnitel = textBox1.Text;
                        Base[stroka_base].material = "";
                        Sps[i].sortir = true;

                        Base[stroka_base].kol = Sps[i].kol;
                        Base[stroka_base].gotovnost = textBox2.Text;


                        for (int i3 = 0; i3 < Sps.Length; i3++)
                        {
                            if (Sps[i3] != null)
                            {
                                if (Sps[i3].doc_name == Sps[i].obozn)
                                {
                                    if (Sps[i3].sortir == false)
                                    {
                                        Base[stroka_base].material = Sps[i3].material;
                                        Sps[i3].sortir = true;

                                    }
                                }

                            }

                        }
                        stroka_base += 1;




                    }
                    
                    if ((Sps[i].tip_stroki == "БЧ") & (Sps[i].sortir == false))
                    {
                        Base[stroka_base] = new Baza();
                        Base[stroka_base].obozn = Sps[i].obozn;
                        Base[stroka_base].naimen = Sps[i].naimen;
                        Base[stroka_base].kol = Sps[i].kol;


                        Base[stroka_base].type_cher = "БЧ";
                        Base[stroka_base].ispolnitel = textBox1.Text;
                        Base[stroka_base].gotovnost = textBox2.Text;

                        Sps[i].sortir = true;
                        stroka_base += 1;
                    }
                    if ((Sps[i].tip_stroki == "СТ") & (Sps[i].sortir == false))
                    {
                        Base[stroka_base] = new Baza();
                        Base[stroka_base].kol = Sps[i].kol;

                        Base[stroka_base].type_cher = "СТ";
                        Base[stroka_base].material = Sps[i].naimen;
                        Base[stroka_base].ispolnitel = textBox1.Text;
                        Base[stroka_base].gotovnost = textBox2.Text;
                        Sps[i].sortir = true;
                        stroka_base += 1;
                    }
                    if ((Sps[i].tip_stroki == "П") & (Sps[i].sortir == false))
                    {
                        Base[stroka_base] = new Baza();
                        Base[stroka_base].kol = Sps[i].kol;
                        Base[stroka_base].obozn = Sps[i].obozn;
                        Base[stroka_base].material = Sps[i].naimen;
                        Base[stroka_base].type_cher = "П";
                        Base[stroka_base].ispolnitel = textBox1.Text;
                        Base[stroka_base].gotovnost = textBox2.Text;
                        Sps[i].sortir = true;
                        stroka_base += 1;
                    }
                    if ((Sps[i].tip_stroki == "М") & (Sps[i].sortir == false))
                    {
                        Base[stroka_base] = new Baza();

                        Base[stroka_base].kol = Sps[i].kol;

                        Base[stroka_base].material = Sps[i].naimen;
                        Base[stroka_base].type_cher = "М";
                        Base[stroka_base].ispolnitel = textBox1.Text;
                        Base[stroka_base].gotovnost = textBox2.Text;
                        Sps[i].sortir = true;
                        stroka_base += 1;
                    }









                }

            }







        }

    }
}
