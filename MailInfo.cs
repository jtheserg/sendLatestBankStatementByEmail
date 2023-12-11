namespace sendLatestBankStatementByEmail;

internal class MailInfo
{
    public Month Month { get; set; }
    public string Year { get; set; }
    public string EmailDate { get; set; }
}

public class Month
{
    public Month(string number)
    {
        Number = number;
    }

    public string Number { get; private set; }
    public string Word
    {
        get
        {
            var monthDictionary = new Dictionary<string, string>
            {
                ["01"] = "Leden",
                ["02"] = "Únor",
                ["03"] = "Březen",
                ["04"] = "Duben",
                ["05"] = "Květen",
                ["06"] = "Červen",
                ["07"] = "Červenec",
                ["08"] = "Srpen",
                ["09"] = "Září",
                ["10"] = "Říjen",
                ["11"] = "Listopad",
                ["12"] = "Prosinec"
            };
            return monthDictionary[Number];
        }
    }
}