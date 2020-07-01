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

        public spec_stroka()
        {
            nomer = nomer;
            format = format;
            poz = poz;
            obozn = obozn;
            naimen = naimen;
            kol = kol;
            prim = prim;
        }

        public void GetInfoSst()
        {
            Console.WriteLine($" Формат : {format} Позиция: {poz}  Обозначение: {obozn}  Наименование: {naimen}  Количество : {kol}  Примечание {prim} ");
        }



    }
}
