using System;
using System.Net;
using System.IO;
using System.Net.Http;
using Aspose.Cells;
using GroupDocs.Conversion.Options.Convert;
using GroupDocs.Conversion;
using GroupDocs.Conversion.FileTypes;
using System.Threading.Tasks;
using Newtonsoft.Json;
using IronPython;
using IronPython.Hosting;

namespace enervus
{
    class Program{
        static async Task Main(String[] args){
            string url = "https://bakerhughesrigcount.gcs-web.com/static-files/b562fc2a-b229-41eb-8407-54dda5dc7295";
            string savePath = @"Worldwide Rig Count Dec 2022.xlsx";
            
            var workbook = new Workbook(@"Worldwide Rig Count Dec 2022.xlsx");
            workbook.Save(@"Worldwide Rig Count Dec 2022.csv");

            string filePath = "Worldwide Rig Count Dec 2022.csv";
            StreamReader reader = null;
            int currentYear =  DateTime.Now.Year;
            if (File.Exists(filePath)){
                reader = new StreamReader(File.OpenRead(filePath));
                reader.ReadLine();
                int cnt = 0;
                List<string> listA = new List<string>();
                while (!reader.EndOfStream){
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    foreach (var item in values){
                            if(item.Length >0 && (int)item[0]>=48 && (int)item[0]<=57 && cnt<=0){
                                int curr_item = Int32.Parse(item);
                                if(curr_item == currentYear-1 || curr_item == currentYear-2){
                                cnt = 168;
                                }
                            }
                            if(cnt>0){
                            listA.Add(item);
                            cnt--;
                            }
                    }
                }

                string output_path = "output.csv";

                using (var file = File.CreateText(output_path)){
                    cnt = 0;
                    for(int i = 0;i<listA.Count;i++){
                        cnt++;
                        if(cnt == 12){
                            cnt = 0;
                            file.WriteLine();
                            if( i == 167){
                                file.Write(',');
                                file.Write(',');
                                file.Write(',');
                                file.Write(',');
                                file.Write(',');
                                file.WriteLine();
                            }
                            continue;
                        }
                        else if(cnt == 11)continue;
                        else {
                        file.Write(',');
                        file.Write(listA[i]);
                        }
                    }
                }
            } else {
                Console.WriteLine("File doesn't exist");
            }
             Console.ReadLine();
        }
    }
}
