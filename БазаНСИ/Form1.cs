using System;
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


        public ISpecificationCommentObject obj; //Объект спецификации
        public ISpecificationBaseObject obj_base; //Базовый Объект спецификации
        public ISpecificationColumns oC ;

        // public ыекш OBJ;


        string aa, bb, cc;

        List<string> path = new List<string>();
        List<string> path_name = new List<string>();
        List<string> spisok_sb = new List<string>();
        List<object> spisok_doc = new List<object>();


        

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
            
        }





        private void START()
        {
            /*
            Excel.Application ex = new Excel.Application();
            ex.Visible = true;
            ex.SheetsInNewWorkbook = 2;
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
            ex.DisplayAlerts = false;
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
            sheet.Name = "База НСИ";
            */


            int stolb = 1;//   A - B - C - D     //Cells(5, 3) = C5
            int stroka = 1; //  1-2-3  
            int nomer_Sps = 1;




            spec_stroka[] Sps = new spec_stroka[1500];

            spec_stroka[] Sp_l1 = new spec_stroka[1500];
            spec_stroka[] Sp_l2 = new spec_stroka[1500];
            spec_stroka[] Sp_l3 = new spec_stroka[1500];
            spec_stroka[] Sp_l4 = new spec_stroka[1500];
            spec_stroka[] Sp_l5 = new spec_stroka[1500];


            spec_stroka[] Spisok_dok = new spec_stroka[1500];



            //for (int i1 = 0; i1 <1000;i1++)
            //{
            //    Sps[i1] = new spec_stroka();
            //}

            for (int ip = 0; ip < path.Count; ip++)
            {
                Spisok_dok[ip] = new spec_stroka();
                Spisok_dok[ip].poz = ip.ToString();


                string obrez_do_naimen = Path.GetFileName(path[ip]).Remove(Path.GetFileName(path[ip]).IndexOf(" "));
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

                Spisok_dok[ip].obozn = obrez_SB;
                Spisok_dok[ip].GetNameFiles();
                //Console.WriteLine("Tessssssssssssssssssssssst "+Spisok_dok[ip].NameDoc());
            }



            Console.WriteLine("Количество документов = " + path.Count);
            for (int i = 0; i < path.Count; i++)
            {


                IKompasDocument doc = appl.Documents.Open(path[i], true, false);// Получаем интерфейс активного документа 2D в API7
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

                        if ((nomer_razdela == 5) | (nomer_razdela_base == 5))
                        {
                            continue;
                        }
                        else
                        {

                            Sps[nomer_Sps] = new spec_stroka();
                            Specification_Columns = Specification_Object.Columns;

                            for (int bCol = 0; bCol < Specification_Columns.Count; bCol++)
                            {
                                Specification_Column = Specification_Columns[bCol];
                                var st = Specification_Column.Text.Str;
                                Console.WriteLine("Столбец " + (bCol + 1) + " - " + st);

                                //Заполнение      /////////////////////////////////////                            ///////////////////////////
                                //sheet.Cells[stroka, stolb] = st;

                                switch (bCol)
                                {
                                    case 0:
                                        Sps[nomer_Sps].format = st;
                                        break;
                                    case 2:
                                        Sps[nomer_Sps].poz = st;
                                        break;
                                    case 3:
                                        Sps[nomer_Sps].obozn = st;
                                        break;
                                    case 4:
                                        Sps[nomer_Sps].naimen = st;
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
                    if (Sps[i2].format != null)
                    {
                        Console.WriteLine("ПРОВЕРКА длина " + Sps.Length + "   " );
                        Sps[i2].GetInfoSst();
                        
                    }
                }
            }

            


        }
    


        string ObrezName()
        {


            return Name;
        }


    







        void GetKompas()
        {
            try
            {            

                kompas = (KompasObject)System.Runtime.InteropServices.Marshal.GetActiveObject("kompas.application.5");
                appl = (IApplication)kompas.ksGetApplication7();
                MessageBox.Show("Подключение установлено");
                appl.KompasError.Clear();

            }
            catch
            {
                MessageBox.Show("Компас не запущен - ЗАПУСКАЕМ ");
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

            short t = 4;
            int nn = 0;
            
            

            


        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            path.Clear();
            path_name.Clear();
            label1.Text = "Перенеси файлы сюда";
        }

        private void button3_Click(object sender, EventArgs e)
        {


        }

        void panel1_DragEnter(object sender, DragEventArgs e)
        {
           if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }

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




        

    }
}
