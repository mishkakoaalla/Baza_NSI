using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace БазаНСИ
{
    class spec_stroka
    {
        public int nomer { get; set; }
        public string format { get; set; }
        public string poz { get; set; }
        public string obozn { get; set; }
        public string naimen { get; set; }
        public string kol { get; set; }
        public string prim { get; set; }


        public string tip_stroki { get; set; } //Сборка, деталь, материал, прочие изделия, БЧ ....
        public string material { get; set; }
        public string doc_name { get; set; }
        public bool sortir { get; set; }

        public spec_stroka()
        {
            
            nomer = nomer;
            format = format;
            poz = poz;
            obozn = obozn;
            naimen = naimen;
            kol = kol;
            prim = prim;

            material = material;
            doc_name = doc_name;
            tip_stroki = tip_stroki;
            sortir = sortir;
        }

        public void GetInfoSst()
        {
            if (tip_stroki != "Материал из детали")
            {
                Console.WriteLine($" Имя документа : {doc_name} Сортирован: {sortir} Тип строки: {tip_stroki}  Обозначение: {obozn}  Наименование: {naimen}  Количество : {kol}  ");
            }
            else
            {
                Console.WriteLine($" Имя документа : {doc_name} Сортирован: {sortir} Тип строки: {tip_stroki}  Обозначение: {obozn}  Наименование: {naimen}  Количество : {kol}  Материал: {material} ");
            }
        }

        public void GetNameFiles()
        {
            Console.WriteLine($" Позиция: {poz}  Обозначение: {obozn} ");
            
        }

         public string NameDoc()
        {
            string dd = obozn;
            return dd;
        }

    }
}
