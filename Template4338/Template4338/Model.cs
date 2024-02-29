using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template4338
{
    public class Model
        {
            public int Id { get; set; }
            public string ФИО { get; set; }
            public long Код_клиента { get; set; }
            public DateTime Дата_рождения { get; set; }
            public int Индекс { get; set; }
            public string Город { get; set; }
            public string Улица { get; set; }
            public int Дом { get; set; }
            public int Квартира { get; set; }
            public string Mail { get; set; }

    }
}
