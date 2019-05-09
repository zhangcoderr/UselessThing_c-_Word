using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
//using MSWord = Microsoft.Office.Interop.Word;

//using Spire.Doc.Documents;
using Aspose.Words;
namespace ConsoleApplication1
{

    class Program
    {
        enum ProjectType
        {
             市政,
             建筑,
             other
        }
        enum EnumBookMark
        {
            招标编号, 
            工程名称, 
            招标人, 
            日期, 
            投标费,
            安全文明施工费,
            担保金值,
            担保金百分比 ,
            工期,
            负责人,
            负责人专业,
            负责人证书,
            误期违约金额,
            预付款金额

        }

        static Dictionary<string, string> dataDic = new Dictionary<string, string>();
        string GetFirstDicValue(string key)
        {
            if (dataDic.ContainsKey(key))
            {
                return dataDic[key];
            }
            else
            {
                return null;
            }
        }
        
        static void Main(string[] args)
        {
            Document doc = new Document(@"C:\Users\Administrator\Desktop\Test\test.docx");
            Document resultDoc=new Document(@"C:\Users\Administrator\Desktop\Test\Base\BaseDoc.docx");
            DocumentBuilder builder = new DocumentBuilder(resultDoc);
            ProjectType projectType=ProjectType.other;
            

            List<string> paraList=new List<string>();

            foreach (Paragraph item in doc.FirstSection.Body.Paragraphs)
            {
                if (item.GetText().Contains("目录")) break;
                //Console.WriteLine(item.GetText());
                paraList.Add(item.GetText());
            }
            bool end = false;
            foreach (string paraStr in paraList)
            {
                if (!paraStr.Contains(":") && !paraStr.Contains("：")) continue;
                if (end) break;
                string key = "";
                string value = "";
                bool inKey = true;
                foreach (char c in paraStr)
                {
                    if (inKey)
                    {
                        if (c != ':' && c != '：'&&c!=' ')
                            key += c;
                    }
                    else
                    {
                        if (c != ':' && c != '：' && c != ' '&&c!='\r')
                            value += c;
                    }
                    if (c == ':' || c == '：')
                    {
                        inKey = false;
                    }
                    if (!string.IsNullOrWhiteSpace(value) &&( c == ' '||c=='\r'))
                    {
                        key.Trim();
                        value.Trim();
                        if (dataDic.ContainsKey(key))
                        {
                            key += "0";

                        }
                        try
                        {
                            if (key == "日期"||key=="时间")
                            {
                                end = true;
                                break;
                            }
                            if (value.Contains("盖章"))
                            {
                                value=  value.Replace("(盖章)", "");
                            }
                            dataDic.Add(key, value);
                        }
                        catch (Exception)
                        {
                            
                            throw;
                        }
                        
                        

                        inKey = true;

                        key = "";
                        value = "";
                    }

                }
            }
            #region find 施工投标文件
            //int paraStartIndex = 0;
            //bool start = false;
            //foreach (Paragraph item in doc.FirstSection.Body.Paragraphs)
            //{
            //    paraStartIndex++;

            //    if (item.GetText().Contains("施工投标文件"))
            //    {

            //        Console.WriteLine(paraStartIndex);
            //        builder.MoveToParagraph(paraStartIndex, 0);
            //        builder.Writeln("eheheheheheh");
            //        start = true;
            //    }
            //    if (start)
            //    {
            //        if (item.GetText().Contains("招标编号"))
            //        {
                      
            //        }
            //    }
                
            //}


            #endregion

            string projectTypeStr = GetStringValue(@"本次招标要求投标人的项目负责人须具备(?<key>.*?)类证书）", doc.Range.Text);

            if (projectTypeStr.Contains("市政公用工程"))
            {
                projectType = ProjectType.市政;
            }
            else if (projectTypeStr.Contains("建筑工程专业"))
            {
                projectType = ProjectType.建筑;
            }
            
            string 中标价= GetStringValue(@"投标报价（中标价）为(?<key>.*?)万元", doc.Range.Text);
            string temp = "";
            foreach (var item in 中标价)
            {
                if (item != '.')
                {
                    temp += item;
                }
            }
            中标价 = temp;

            string 安全文明施工费 = GetStringValue(@"安全文明施工费(?<key>.*?)万元", doc.Range.Text);
            string 工期 = GetStringValue(@"工期要求\a(?<key>.*?)日历天", doc.Range.Text);//
            string 担保金百分比=GetStringValue(@"交纳中标价(?<key>.*?)%的履约保证金",doc.Range.Text);
            string 担保金值=GetStringValue(@"投标担保金额\a(?<key>.*?)元",doc.Range.Text);
            string 日期 = GetStringValue(@"投标文件递交截止时间.*?2019年(?<key>.*?)日", doc.Range.Text);
            string 误期违约金额 = GetStringValue(@"误期违约金额.*?(?<key>.*?)元", doc.Range.Text);
            string 预付款金额 = GetStringValue(@"预付款金额(:|：).*?(?<key>.*?)预付款保函金额", doc.Range.Text);

            if (!dataDic.ContainsKey("工程名称"))
            {
                dataDic.Add("工程名称","桐庐县合村乡合村村高标准基本农田建设项目");
                Console.WriteLine("工程名称不存在");
            }
             if (!dataDic.ContainsKey("招标编号"))
            {
                dataDic.Add("招标编号", "TLGT20190006");
                Console.WriteLine("招标编号不存在");
            }
            dataDic.Add(EnumBookMark.安全文明施工费.ToString(), 安全文明施工费);
            dataDic.Add(EnumBookMark.担保金百分比.ToString(), 担保金百分比);
            dataDic.Add(EnumBookMark.担保金值.ToString(), 担保金值);
            dataDic.Add(EnumBookMark.工期.ToString(), 工期);
            dataDic.Add(EnumBookMark.日期.ToString(), 日期);
            dataDic.Add(EnumBookMark.投标费.ToString(), 中标价);
            //dataDic.Add(EnumBookMark.误期违约金额.ToString(), 误期违约金额);
            //dataDic.Add(EnumBookMark.预付款金额.ToString(), 预付款金额);
            
            switch (projectType)
            {
                case ProjectType.市政:
                    {
                        dataDic.Add(EnumBookMark.负责人.ToString(), "林小东");
                        dataDic.Add(EnumBookMark.负责人证书.ToString(), " 二级建造师证/浙233131283587");
                        dataDic.Add(EnumBookMark.负责人专业.ToString(), "市政公用工程");
                        


                    }
                    break;
                case ProjectType.建筑:
                    {
                        dataDic.Add(EnumBookMark.负责人.ToString(), "吴敏");
                        dataDic.Add(EnumBookMark.负责人证书.ToString(), " 二级建造师证/浙233181802351 ");
                        dataDic.Add(EnumBookMark.负责人专业.ToString(), "建筑工程");
                        

                    }
                    break;
                case ProjectType.other:
                    break;
                default:
                    break;
            }




            Write(builder, EnumBookMark.招标编号.ToString(), 2);
            Write(builder, EnumBookMark.工程名称.ToString(), 10);
            Write(builder, EnumBookMark.招标人.ToString(), 2);
            Write(builder, EnumBookMark.日期.ToString(), 5);

            Write(builder, EnumBookMark.投标费.ToString());

            Write(builder, EnumBookMark.安全文明施工费.ToString());
            Write(builder, EnumBookMark.担保金值.ToString());
            Write(builder, EnumBookMark.担保金百分比.ToString());
            Write(builder, EnumBookMark.工期.ToString(), 1);
            if (projectType != ProjectType.other)
            {
                Write(builder, EnumBookMark.负责人.ToString(), 2);
                Write(builder, EnumBookMark.负责人证书.ToString());
                Write(builder, EnumBookMark.负责人专业.ToString());
            }
            
            //Write(builder, EnumBookMark.预付款金额.ToString());
            //Write(builder, EnumBookMark.误期违约金额.ToString());






            resultDoc.Save(@"C:\Users\Administrator\Desktop\test\ResultDoc(" + DateTime.Now.Day + "号" + DateTime.Now.Hour + "时" + DateTime.Now.Minute +"分"+DateTime.Now.Second+ "秒).docx");

            Console.WriteLine("success!  "+projectType.ToString());

            Console.Read();
        }
        static string GetStringValue(string reStr, string docText)
        {
            Regex re = new Regex(reStr);
            Match match = re.Match(docText);
            return match.Groups["key"].Value;
        }
        static void Write(DocumentBuilder builder, string bookMark,int count=0)
        {
            if (count > 0)
            {
                for (int i = 0; i <= count; i++)
                {
                    builder.MoveToBookmark(bookMark + i.ToString());
                   
                    if(dataDic.ContainsKey(bookMark))
                        builder.Write(dataDic[bookMark]);
                    else
                        Console.WriteLine(bookMark);

                }
         
                
            }
            else
            {
                builder.MoveToBookmark(bookMark);
                builder.Write(dataDic[bookMark]);
            }
           
        }
    }
}
