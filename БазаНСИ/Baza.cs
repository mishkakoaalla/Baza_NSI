using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace БазаНСИ
{
    class Baza
    {
        public string name_doc { get; set; }
        public bool sort { get; set; }
        public string vhodimost { get; set; }
        public string naimen { get; set; }
        public string obozn { get; set; }
        public string type_cher { get; set; }
        public string ispolnitel { get; set; }
        public string gotovnost { get; set; }
        public string kuda_vhodit { get; set; }
        public string material { get; set; }
        public string kol { get; set; }
        
        
        public Baza()
        {
            name_doc = name_doc;
            sort = sort;
            vhodimost = vhodimost;
            naimen = naimen;
            obozn = obozn;
            type_cher = type_cher;
            ispolnitel = ispolnitel;
            gotovnost = gotovnost;
            kuda_vhodit = kuda_vhodit;
            material = material;
            kol = kol;
        }

        public void DrawBase()
        {
            Console.WriteLine($" : {vhodimost} : {naimen} : {obozn} : {type_cher} : {ispolnitel} : {gotovnost} : {kuda_vhodit} : {material} : {kol} ");
        }


    }
}
