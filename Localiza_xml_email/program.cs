//Criador:Felipe Machado Ignacio dos santos
//ultima atualizaçao 25/08/2020
//Biblioteca usada no nuget: mail.dll

/*este projeto tem como principal objetivo receber anexos de e-mails não lidos
recebe os anexos na pasta temp, e depois guarda-os em um servidor na rede.
logo após avisa o usuario que existe um xml novo no servidor*/

using Limilabs.Client.IMAP;
using Limilabs.Mail;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;


namespace Localiza_xml_email
{
    class Program
    {
        //strings de configuração
        const string Imap_Config = "imap.server.com";
        const string login_Config = "email@email.com.br";
        const string Pass_Config = "";

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
            //usando este metodo para esconder o console
            IntPtr handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);
            //usando este metodo para esconder o console

            //inicializa as configurações de imap
            Imap imap = new Imap();
            imap.Connect(Imap_Config);
            imap.UseBestLogin(login_Config, Pass_Config);
            imap.SelectInbox();
            List<long> uidList = imap.Search(Flag.Unseen);// recebe a lista de e-mails não lidos
            foreach (long uid in uidList)
            {
                IMail email = new MailBuilder().CreateFromEml(imap.GetMessageByUID(uid));//variavel recebe o e-mail não lido com todos os atributos
                Classe_email.Quantidade = email.Attachments.Count;//recebe a quantidade de anexos do e-mail
                if (Classe_email.Quantidade > 0)
                {
                    int i = 0;
                    while (i < Classe_email.Quantidade)//enquanto o contador i for menor que a quantidade de anexos ele executa o que tem dentro do while
                    {
                        Classe_email.Nome = email.Attachments[i].FileName;//recebe o nome do arquivo
                        Classe_email.Caminho_Temp = Path.GetTempPath() + Classe_email.Nome; //path da pasta temp do windows + nome do arquivo
                        Classe_email.Caminho_Servidor = @"\\servidor\base\xml_compras\" + DateTime.Now.ToString("MM_yyyy"); // caminho no servidor local, onde serão salvos o xml
                        email.Attachments[i].Save(Classe_email.Caminho_Temp);//salva o anexo na pasta temp
                        FileInfo file = new FileInfo(Classe_email.Caminho_Temp);// file info do arquivo temporario

                        if (file.Extension == ".xml")
                        {
                            try
                            {   //Lê o arquivo xml e pega a chave dele 
                                using (XmlReader reader = XmlReader.Create(file.FullName))
                                {
                                    while (reader.Read())
                                    {
                                        if (reader.IsStartElement())
                                        {
                                            switch (reader.Name.ToString())
                                            {
                                                case "chNFe":
                                                    Classe_email.Chave = reader.ReadString();
                                                    break;
                                            }
                                        }
                                    }

                                }
                            }
                            catch (XmlException e)
                            {
                                MessageBox((IntPtr)0, "" + e.ToString() + "", "Erro", 0);
                            }
                            if (Directory.Exists(Classe_email.Caminho_Servidor))
                                Copia_arquivo(Classe_email.Caminho_Temp, Classe_email.Caminho_Servidor + @"\" + Classe_email.Chave + ".xml");//metodo que copia mensagem para o servidor e mostra mensagem
                            else
                            {
                                Directory.CreateDirectory(Classe_email.Caminho_Servidor);//cria diretorio caso não exista
                                Copia_arquivo(Classe_email.Caminho_Temp, Classe_email.Caminho_Servidor + @"\" + Classe_email.Chave + ".xml");//metodo que copia mensagem para o servidor e mostra mensagem
                            }
                        }
                        try
                        {
                            file.Delete();//exclui o arquivo quando terminar de executar
                        }
                        catch (IOException e)
                        {
                            MessageBox((IntPtr)0, "" + e.ToString() + "", "Erro", 0);
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
            File.Copy(arquivo_origem, arquivo_destino, true);//copia o xml e sobrepoe caso exista, as vezes o e-mail vem duplicado do fornecedor
            int count = 0;
            try
            {
                using (XmlReader reader = XmlReader.Create(arquivo_destino))
                {
                    while (reader.Read())
                    {
                        if (reader.IsStartElement())
                        {
                            if (reader.Name.ToString() == "xNome")
                            {
                                if (count == 0)//no xml o primeiro elemento que contem xNome é onde está localizado o nome do fornecedor, empresa que emitiu o xml
                                {
                                    Classe_email.Fornecedor = reader.ReadString();
                                    count++;
                                }
                            }
                            if (reader.Name.ToString() == "natOp") //natureza da operação no xml
                            {
                                Classe_email.Natureza_op = reader.ReadString();
                            }
                        }
                    }
                }
                MessageBox((IntPtr)0, "Atenção, novo xml copiado\n" + arquivo_destino + "\n\nFORNCEDOR:" + Classe_email.Fornecedor.ToUpper() + "\nNatureza da op:" + Classe_email.Natureza_op, " Novo Xml", 0);
            }
            catch (XmlException e)
            {
                MessageBox((IntPtr)0, "" + e.ToString() + "", "Erro", 0);
            }
        }

        private static class Classe_email
        {
            public static int Quantidade { get; set; }
            public static string Nome { get; set; }
            public static string Transp { get; set; }
            public static string Natureza_op { get; set; }
            public static string Fornecedor { get; set; }
            public static string Chave { get; set; }
            public static string Caminho_Servidor { get; set; }
            public static string Caminho_Temp { get; set; }

        }
    }
}
