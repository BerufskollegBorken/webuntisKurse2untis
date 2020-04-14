using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace webuntisKurse2untis
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("webuntisKurse2untis");
            Console.WriteLine("===================");
            Console.WriteLine("");

            Studentgroups studentgroups = new Studentgroups();
            ExportLessons exportlessons = new ExportLessons();
            
            string datei = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\StudentgroupStudents.csv";
            
            if (!File.Exists(datei))
            {
                Console.WriteLine("Die Datei " + datei + " existiert nicht.");
                Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
                Console.WriteLine("1. sich als admin anmelden");
                Console.WriteLine("2. auf Administration > Export klicken");
                Console.WriteLine("3. Schülerguppen als CSV exportieren nach " + datei);
                Console.WriteLine("ENTER beendet das Programm.");
                Console.ReadKey();
                Environment.Exit(0);
            }
            else
            {
                if (System.IO.File.GetLastWriteTime(datei).Date != DateTime.Now.Date)
                {
                    Console.WriteLine("Die Datei " + datei + " ist nicht von heute.");
                    Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
                    Console.WriteLine("1. sich als admin anmelden");
                    Console.WriteLine("2. auf Administration > Export klicken");
                    Console.WriteLine("3. Schülerguppen als CSV exportieren nach " + datei);
                    Console.WriteLine("ENTER beendet das Programm.");
                    Console.ReadKey();
                    Environment.Exit(0);
                }
            }
            
            using (StreamReader reader = new StreamReader(datei))
            {
                Console.Write("Schülergruppen aus Webuntis ".PadRight(30, '.'));
                
                while (true)
                {                    
                    string line = reader.ReadLine();
                    try
                    {
                        Studentgroup studentgroup = new Studentgroup();
                        var x = line.Split('\t');
                        studentgroup.StudentId = Convert.ToInt32(x[0]);
                        studentgroup.Name = x[1];
                        studentgroup.Forename = x[2];
                        studentgroup.StudentgroupName = x[3];
                        studentgroup.Subject = x[4];
                        studentgroup.StartDate = x[5];
                        studentgroup.EndDate = x[6];
                        studentgroup.Kurzname = studentgroup.generateKurzname();
                        studentgroups.Add(studentgroup);                       
                    }
                    catch (Exception)
                    {                        
                    }

                    if (line == null)
                    {
                        break;
                    }                    
                }
                Console.WriteLine((" " + studentgroups.Count.ToString()).PadLeft(30, '.'));
            }

            datei = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\ExportLessons.csv";

            if (!File.Exists(datei))
            {
                Console.WriteLine("Die Datei " + datei + " existiert nicht.");
                Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
                Console.WriteLine("1. sich als admin anmelden");
                Console.WriteLine("2. auf Administration > Export klicken");
                Console.WriteLine("3. Unterricht als CSV exportieren nach " + datei);
                Console.WriteLine("ENTER beendet das Programm.");
                Console.ReadKey();
                Environment.Exit(0);
            }
            else
            {
                if (System.IO.File.GetLastWriteTime(datei).Date != DateTime.Now.Date)
                {
                    Console.WriteLine("Die Datei " + datei + " ist nicht von heute.");
                    Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
                    Console.WriteLine("1. sich als admin anmelden");
                    Console.WriteLine("2. auf Administration > Export klicken");
                    Console.WriteLine("3. Unterricht als CSV exportieren nach " + datei);
                    Console.WriteLine("ENTER beendet das Programm.");
                    Console.ReadKey();
                    Environment.Exit(0);
                }
            }
            using (StreamReader reader = new StreamReader(datei))
            {
                Console.Write("Unterrichte aus Webuntis ".PadRight(30, '.'));

                while (true)
                {
                    string line = reader.ReadLine();
                    try
                    {
                        ExportLesson exportLesson = new ExportLesson();
                        var x = line.Split('\t');
                        exportLesson.LessonId = Convert.ToInt32(x[0]);
                        exportLesson.LessonNumber = Convert.ToInt32(x[1])/100;
                        exportLesson.Subject = x[2];
                        exportLesson.Teacher = x[3];
                        exportLesson.Klassen = x[4];
                        exportLesson.Studentgroup = x[5];
                        exportLesson.Periods = x[6];
                        exportLesson.Startdate = x[7];
                        exportLesson.EndDate = x[8];
                        exportLesson.Room = x[9];
                        exportLesson.Foreignkey = x[9];                        
                        exportlessons.Add(exportLesson);                                                
                    }
                    catch (Exception)
                    {
                    }

                    if (line == null)
                    {
                        break;
                    }
                }
                Console.WriteLine((" " + exportlessons.Count.ToString()).PadLeft(30, '.'));
            }

            datei = Directory.GetCurrentDirectory() + "\\GPU015_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".TXT";

            using (StreamWriter writer = new StreamWriter(datei))
            {
                foreach (var studentgroup in studentgroups)
                {
                    writer.Write("\"" + studentgroup.Kurzname + "\"|");  // Kurzname
                    writer.Write((from e in exportlessons where e.Studentgroup == studentgroup.StudentgroupName select e.LessonNumber).FirstOrDefault() + "|"); // Unterrichtsnummer
                    writer.Write("\"" + studentgroup.Subject + "\"|"); // Fachkürzel
                    writer.Write("|");
                    writer.Write("\"" + (from e in exportlessons where e.Studentgroup == studentgroup.StudentgroupName select e.Klassen).FirstOrDefault() + "\"|"); // Klasse
                    writer.Write("|");
                    writer.Write("\"0\"|");
                    writer.Write("|");
                    writer.Write("|");
                    writer.Write("\"" + (from e in exportlessons where e.Studentgroup == studentgroup.StudentgroupName select e.LessonNumber).FirstOrDefault() + "\"|"); // Unterrichtsnummer
                    writer.Write("\"" + studentgroup.Subject + "\"|"); // Fachkürzel
                    writer.Write("|");
                    writer.WriteLine("\"1\"|");
                }
            }

            try
            {
                Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            try
            {
                System.Diagnostics.Process.Start(datei);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            Console.WriteLine("Jetzt zuerst die Stundenten aus Webuntis importieren.");
            Console.WriteLine("Dann die exportierte GPU015 in Untis einlesen.");
            Console.WriteLine("ENTER schließt die Anwendung.");
            Console.ReadKey();
        }
    }
}
