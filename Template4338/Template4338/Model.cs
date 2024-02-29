using Newtonsoft.Json;
using System;

public class Model
{
    [JsonIgnore]
    public int Id { get; set; }

    public string FullName { get; set; }
    public string CodeClient { get; set; }
    public string BirthDate { get; set; }
    public string Index { get; set; }
    public string City { get; set; }
    public string Street { get; set; }
    public int Home { get; set; }
    public int Kvartira { get; set; }
    public string E_mail { get; set; }
}
