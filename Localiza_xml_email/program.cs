using Limilabs.Client.IMAP;
using Limilabs.Mail;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;

namespace Localiza_xml_email
{
    class Program
    {
        [DllImport("User32.dll", CharSet = CharSet.Unicode)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("User32.dll", CharSet = CharSet.Unicode)]
        public static extern int MessageBox(IntPtr h, string m, string c, int type);



        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;


        static void Main(string[] args)
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);

            Imap imap = new Imap();
            imap.Connect("imap.server.com");
            imap.UseBestLogin("login", "senha");

            imap.SelectInbox();
            List<long> uidList = imap.Search(Flag.Unseen);
            foreach (long uid in uidList)
            {
                IMail email = new MailBuilder().CreateFromEml(imap.GetMessageByUID(uid));
                Classe_email.Quantidade = email.Attachments.Count;
                if (Classe_email.Quantidade > 0)
                {
                    int i = 0;
                    while (i < Classe_email.Quantidade)
                    {
                        Classe_email.Nome = email.Attachments[i].FileName;
                        string caminho_temp = Path.GetTempPath() + Classe_email.Nome;
                        string caminho_servidor = @"\\servidor\base\xml_compras\" + DateTime.Now.ToString("MM_yyyy");

                        email.Attachments[i].Save(caminho_temp);
                        FileInfo file = new FileInfo(caminho_temp);

                        if (file.Extension == ".xml")
                        {
                            if (Directory.Exists(caminho_servidor))
                                Copia_arquivo(caminho_temp, caminho_servidor + @"\" + Classe_email.Nome + "");
                            else
                            {
                                Directory.CreateDirectory(caminho_servidor);
                                Copia_arquivo(caminho_temp, caminho_servidor + @"\" + Classe_email.Nome + "");
                            }
                        }
                        i++;
                    }
                }

            }
            imap.Close();
            Environment.Exit(0);
        }
        static void Copia_arquivo(string arquivo_origem, string arquivo_destino)
        {
            File.Copy(arquivo_origem, arquivo_destino, true);


            try
            {
                using (XmlReader reader = XmlReader.Create(arquivo_destino))
                {
                    while (reader.Read())
                    {
                        if (reader.IsStartElement())
                        {
                            switch (reader.Name.ToString())
                            {

                                case "xNome":
                                    Classe_email.Transp = reader.ReadString();
                                    break;
                                case "xFant":
                                    Classe_email.Nome_xml = reader.ReadString();
                                    break;
                                case "natOp":
                                    Classe_email.Natureza_op = reader.ReadString();
                                    break;
                            }
                        }
                    }
                }
                MessageBox((IntPtr)0, "Atenção, novo xml copiado\n" + arquivo_destino + "\nTransportadora:" + Classe_email.Transp + "\nNatureza da op:" + Classe_email.Natureza_op + "\nEmpresa:" + Classe_email.Nome_xml + "", "Novo Xml", 0);
            }
            catch (XmlException e)
            {
                MessageBox((IntPtr)0, "" + e.ToString() + "", "Novo Xml", 0);
            }
        }

        private static class Classe_email
        {
            public static int Quantidade { get; set; }
            public static string Nome { get; set; }
            public static string Transp { get; set; }
            public static string Natureza_op { get; set; }
            public static string Nome_xml { get; set; }


        }
    }
}
