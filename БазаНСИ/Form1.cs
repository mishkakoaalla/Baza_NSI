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
//тестирование коммита 
erroererererer


namespace БазаНСИ
{
    public partial class Form1 : Form
    {
        private KompasObject kompas;
        private IApplication appl;         // Интерфейс приложения
       // private IKompasDocument2D doc;          // Интерфейс документа 2D в API7 
        //private ksDocument2D doc2D;        // Интерфейс документа 2D в API5

        string aa, bb, cc;

        List<string> path = new List<string>();

        void close_file()
        {
            Console.WriteLine("ЗАКРЫТИЕ ДОКУМЕНТЫ");
            ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();
            if (doc != null)
                doc.ksCloseDocument();
        }



        public Form1()
        {
            InitializeComponent();
            GetKompas();   
            //START();
        }





        private void START()
        {
           
            Console.WriteLine("Количество документов = " + path.Count);
            for (int i = 0; i < path.Count; i++)
            {

                IKompasDocument doc = appl.Documents.Open(path[i], true, true);// Получаем интерфейс активного документа 2D в API7




                Console.WriteLine("Получение спецификации из документа № - " + Convert.ToInt32(i+1) );
                SpecificationDescription Specification_Descriptions = doc.SpecificationDescriptions.Active;

                if (Specification_Descriptions != null)
                {

                    ISpecificationCommentObjects SpcObjects = Specification_Descriptions.CommentObjects;
                    ISpecificationBaseObjects SpcObjectsBase = Specification_Descriptions.BaseObjects;
                    

                    var num = SpcObjects.Count;
                    ISpecificationObject Specification_Object;
                    ISpecificationColumns Specification_Columns;
                    ISpecificationColumns Specification_Additional_Columns;
                    ISpecificationColumn Specification_Column;


                    for (int SD = 0; SD < num; SD++)
                    {
                        Console.WriteLine("----- Строка " + (SD + 1) + "   ---- ");
                        Console.WriteLine("");
                        //var ww = SpcObjects[SD];
                        

                        ISpecificationCommentObject obj = SpcObjects[SD];

                        var q = obj.Subsection;
                        var q1 = obj.UniqueNumber;
                        var q2 = obj.BaseObject;
                        var q3 = obj.Columns;
                        var q4 = obj.Parent;
                        var q5 = obj.State;
                        var q6 = obj.BlockNumber;
                        var q7 = obj.BlockNumberByIndex[0];



                        Console.WriteLine("!!!! ТИП ОБЬЕКТА " + (q,q1,q2,q3,q4,q5,q6,q7) + "   !!!!!");


                        ISpecificationColumns oC = obj.Columns;
                        Specification_Object = obj;

                        long qq = obj.Section;
                        Console.WriteLine("!!!!СЕКЦИЯ " + (qq) + "   !!!!!");



                        Specification_Columns = Specification_Object.Columns;
                        Specification_Additional_Columns = Specification_Object.AdditionalColumns;
                        for (int bCol = 0; bCol < Specification_Columns.Count; bCol++)
                        {
                            



                            Specification_Column = Specification_Columns[bCol];
                            var st = Specification_Column.Text.Str;
                            Console.WriteLine("Столбец " + (bCol + 1) + " - " + st);
                        }
                        Console.WriteLine("----- Конец cтроки ---- ");




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
                        doc.Close(0); //Закрыть документ
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
                MessageBox.Show("Подключение установлено");

            }
            catch
            {
                MessageBox.Show("Компас не запущен - ЗАПУСКАЕМ ");
                Type t = Type.GetTypeFromProgID("KOMPAS.Application.5");
                kompas = (KompasObject)Activator.CreateInstance(t);
                kompas = (KompasObject)System.Runtime.InteropServices.Marshal.GetActiveObject("kompas.application.5");
                appl = (IApplication)kompas.ksGetApplication7();
                kompas.Visible = true;  //  
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
                    path.AddRange(Directory.GetFiles(obj, "*.*", SearchOption.AllDirectories)
                        .Where(f=> f.EndsWith(".cdw")|| f.EndsWith(".spw")).ToArray()
                        
                        
                        );
                }
                else
                {

                    path.Add(obj);
                    Console.WriteLine("Док " + obj);
                }
            label1.Text = string.Join("\r\n", path);

            
           // label1.Text += file + "\n";
            

            
        }


    }
}
