using System;
using ClosedXML.Excel;
using HtmlAgilityPack;



namespace CsharpWebScraping
{   


    class Program
    {
        public struct Ads // The structure that holds the information about the add
        {
            public string title;
            public int price;
            public int metri;
            public string url;
        };
    
        static void Main(string[] args)
        {
            Ads[] addList = new Ads[1000]; 
            string urlB = "https://www.olx.ro/imobiliare/apartamente-garsoniere-de-inchiriat/oradea/?currency=EUR&page=";
            string toAdd = "&search%5Bfilter_float_price:from%5D=150&search%5Bfilter_float_price:to%5D=300&search%5Bfilter_float_m:from%5D=10&search%5Bfilter_float_m:to%5D=100";
            //URL has 2 parts, first is the actual URL and the second are the filters
            var httpClient = new HttpClient();

            var html = httpClient.GetStringAsync(urlB+toAdd).Result;
            HtmlDocument htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(html);


            var pagenumber = htmlDocument.DocumentNode.Descendants("li").Where(node =>node.GetAttributeValue("class", "").Equals("  pagination-item  css-ps94ux")).ToList();
            int max = int.Parse(pagenumber[pagenumber.Count-1].InnerText);//get the number of pagers

            List <float> perPrice = new List<float>(new float[2000]);//a list for the prices
            int n=0;
            for(int a = 1;a<=max;a++)//itterate through the pages
            {
                string url = urlB;
                url = url+a+toAdd;// create the url with the curent page
                html = httpClient.GetStringAsync(url).Result;
                htmlDocument.LoadHtml(html);
                var adsTitle = htmlDocument.DocumentNode.Descendants("h6").Where(node => node.GetAttributeValue("class", "").Equals("css-16v5mdi er34gjf0")).ToList();
                var adsPrice = htmlDocument.DocumentNode.Descendants("p").Where(node =>node.GetAttributeValue("class", "").Equals("css-10b0gli er34gjf0")).ToList();
                var adsSpace = htmlDocument.DocumentNode.Descendants("span").Where(node =>node.GetAttributeValue("class", "").Equals("css-643j0o")).ToList();
                //get all the info that I need
                Thread.Sleep(1000);

                int preN = n;
                for(int i=0;i<adsTitle.Count;i++)//Adding the adds to my struct
                {
                    addList[i+preN].title = adsTitle[i].InnerText;
                }
                int m=0;
                for(int i=0;i<adsPrice.Count;i++)
                {
                    char[] word = new char[200];
                    int index=0;
                    for(int j=0;j<adsPrice[i].InnerText.Length;j++)
                    {
                        if(Char.IsDigit(adsPrice[i].InnerText[j]))//the prices string have $ and some text and I don't need that
                        {
                            word[index] = adsPrice[i].InnerText[j];
                            index++;
                        }
                    }

                    addList[i+preN].price = int.Parse(word);
                    m=i;
                }
                for(int i=0;i<adsSpace.Count;i++)
                {
                    int index = 0;
                    char[] word = new char[200];
                    for(int j=0;j<adsSpace[i].InnerText.Length;j++)
                    {
                        if(Char.IsDigit(adsSpace[i].InnerText[j]))//the same as the price
                        {
                            word[index]=adsSpace[i].InnerText[j];
                            index++;
                        }
                        else j = adsSpace[i].InnerText.Length;
                    }
                    addList[i+preN].metri = int.Parse(word);
                }





                List <string> links = new List<string>();
                foreach (HtmlNode link in htmlDocument.DocumentNode.SelectNodes("//a[@href]"))//I'm getting all the links from that page
                {
                    HtmlAttribute att = link.Attributes["href"];
                    
                    if (att.Value.Contains("a"))
                    {
                        // showing output
                        links.Add(att.Value);
                    }
                }
                int ind=0;
                for(int i=0;i<links.Count;i++)
                {
                    if(links[i].Contains("oferta"))//I'm chosing the links that have "oferta" in them
                    {
                        string word = links[i];
                        string nWord = "";
                        if(word[0] == '/')
                        {
                            nWord = "https://www.olx.ro"+word;// most of the url don't have the www.

                            addList[ind+preN].url = nWord;
                            ind++;
                        }
                        else
                        {
                            addList[ind+preN].url = links[i];
                            ind++;


                        }
                    
                    }
                }
                preN+=m;
                n=preN;
            }
            for(int i=0;i<=n;i++)
            {
                perPrice[i] = (float)addList[i].price/addList[i].metri;//I calculate the price per square meter 
            }
            for(int i=0;i<=n;i++)
            {
                for(int j=0;j<=n;j++)
                {
                    if(perPrice[i] < perPrice[j])//sorting the list based on the price per square meter
                    {
                        var temp = perPrice[i];
                        perPrice[i] = perPrice[j];
                        perPrice[j] = temp;
                        var t = addList[i];
                        addList[i] = addList[j];
                        addList[j] = t;
                    }
                }
            }

            Ads[] tem = new Ads[2000];
            tem = addList.Distinct().ToArray();
            addList = tem; 
            n=addList.Count()-1;
            using (var workbook = new XLWorkbook())//exporting data to an excel file
                {
                    var workSheet = workbook.Worksheets.Add("Anunturi");//Creating a worksheet  and giving cells a name
                    workSheet.Cell("A1").Value = "TITLU";
                    workSheet.Cell("B1").Value = "PRET";
                    workSheet.Cell("C1").Value = "SUPRAFATA";
                    workSheet.Cell("D1").Value = "URL";
                    
                    int a=2;
                    for(int i=0;i<=n;i++)//adding all the items in the list to the excel sheet
                    {
                        if(can(addList[i]))//this is a function that verifies if the title or URL are null
                        {
                            workSheet.Cell("A"+a).Value = addList[i].title;
                            workSheet.Cell("B"+a).Value = addList[i].price;
                            workSheet.Cell("C"+a).Value = addList[i].metri;
                            workSheet.Cell("D"+a).Value = addList[i].url;
                            a++;
                        }
                    }
                    workSheet.Columns().AdjustToContents();
                    workSheet.Rows().AdjustToContents();
                    workbook.SaveAs("Anunturi.xlsx");
                }


                
                
        }
        public static bool can(Ads temp)//verify if the title or URL are null
        {
            if(temp.title == null)
                return false;
            if(temp.url == null)
                return false;
            return true;
        }

        
    }
}