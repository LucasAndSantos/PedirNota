using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Diagnostics;

namespace PedirNota
{
    public partial class Form1 : Form
    {
        private MailMessage Email;
        Stopwatch Stop = new Stopwatch();
        string verifica;

        public Form1()
        {
            InitializeComponent();
            txt_nome.Text = (Environment.UserName);           

            string caminho = @"C:\DP-Field\Configuracao_de_Email.txt";
            if (!File.Exists(caminho))
            {
                txt_enviarPara.Text = "raquel.nogueira@dpaschoal.com.br";
                txt_copiaPara.Text = "ti.laboratorio@dpaschoal.com.br";
            }
            else
            {
                string[] envioPara = File.ReadAllLines(@"C:\DP-Field\Configuracao_de_Email.txt");
                string[] copiaPara = File.ReadAllLines(@"C:\DP-Field\Configuracao_de_Email.txt");               
                txt_enviarPara.Text = envioPara[0];
                if (txt_copiaPara.Text.Length == 0)
                {
                    txt_copiaPara.Text = "ti.laboratorio@dpaschoal.com.br";
                }
                else
                {
                    txt_copiaPara.Text = copiaPara[1];
                }                
            }           
        }
        public void Form1_Load(object sender, EventArgs e)
        {
            if (Environment.UserName == "Administrador" || Environment.UserName == "Admin")
            {
                MessageBox.Show(" Por questões de segurança e controle usuário 'Administrador' não" +
                    " pode solicitar nota \r\n Favor utilizar outro usuário para realizar a solicitação",
                    "Trocar de usuário", MessageBoxButtons.OK, MessageBoxIcon.Error);

                Application.Exit();
            }
        }
        public void txt_destino_TextChanged(object sender, EventArgs e)
        {
            string cdCentroDecusto = "";

            if (txt_destino.Text == "01")            
                cdCentroDecusto = "169";
            
            if (txt_destino.Text == "02")            
                cdCentroDecusto = "";

            if (txt_destino.Text == "03")            
                cdCentroDecusto = "";
            
            if (txt_destino.Text == "04")            
                cdCentroDecusto = "170";
            
            if (txt_destino.Text == "05")            
                cdCentroDecusto = "171";
            
            if (txt_destino.Text == "06")            
                cdCentroDecusto = "";
            
            if (txt_destino.Text == "07")            
                cdCentroDecusto = "176";
            
            if (txt_destino.Text == "08")            
                cdCentroDecusto = "182";
           
            if (txt_destino.Text == "09")            
                cdCentroDecusto = "185";
            
            if (txt_destino.Text == "10")            
                cdCentroDecusto = "193";
            
            if (txt_destino.Text == "11")            
                cdCentroDecusto = "202";
            
            if (txt_destino.Text == "12")            
                cdCentroDecusto = "203";
            
            if (txt_destino.Text == "13")            
                cdCentroDecusto = "220";
            
            if (txt_destino.Text == "14")            
                cdCentroDecusto = "218";
            
            if (txt_destino.Text == "15")            
                cdCentroDecusto = "226";
            
            if (txt_destino.Text == "16")            
                cdCentroDecusto = "239";
            
            if (txt_destino.Text == "17")            
                cdCentroDecusto = "";
            
            if (txt_destino.Text == "18")            
                cdCentroDecusto = "247";
            
            if (txt_destino.Text == "19")            
                cdCentroDecusto = "251";
            
            if (txt_destino.Text == "20")            
                cdCentroDecusto = "237";
            
            if (txt_destino.Text == "21")            
                cdCentroDecusto = "265";
            
            if (txt_destino.Text == "22")            
                cdCentroDecusto = "265";
            
            if (txt_destino.Text == "23")            
                cdCentroDecusto = "283";

            if (txt_destino.Text == "24")
                cdCentroDecusto = "";

            if (txt_destino.Text == "25")
                cdCentroDecusto = "";

            if (txt_destino.Text == "26")
                cdCentroDecusto = "";

            if (txt_destino.Text == "27")
                cdCentroDecusto = "";

            if (txt_destino.Text == "28")
                cdCentroDecusto = "";

            if (txt_destino.Text == "29")
                cdCentroDecusto = "";

            if (txt_destino.Text == "30")
                cdCentroDecusto = "";

            if (txt_destino.Text.Length == 3)
            {
                txt_solicitante.Text = "adminstrativo" + txt_destino.Text + "@dpaschoal.com.br";
                txt_centroDeCusto.Text = "030" + txt_destino.Text;
            }
            if (txt_destino.Text.Length == 2)
            {
                txt_solicitante.Text = "@dpaschoal.com.br";
                txt_centroDeCusto.Text = "038" + cdCentroDecusto;
            }
            if (txt_destino.Text.Length == 0)
            {
                txt_solicitante.Text = "";
                txt_centroDeCusto.Text = "";
            }
        }
        public void Limpar()
        {
            txt_destino.Text = "";
            txt_peso.Text = "";
            txt_valor.Text = "";
            cmb_equipamento.Text = "";
            txt_codSap.Text = "";
            txt_marca.Text = "";
            txt_modelo.Text = "";
            txt_hostName.Text = "";
            txt_seNovo.Text = "";
            txt_seAntigo.Text = "";
            txt_numeroDeSerie.Text = "";
            txt_chamado.Text = "";
            txt_formaDeEnvio.Text = "Sedex";
            txt_solicitante.Text = "";
            txt_centroDeCusto.Text = "";
            txt_centroDeCustoTi.Text = "110300";
        }
        public void btn_limpar_Click(object sender, EventArgs e)
        {
            Limpar();
        }
        public void btn_enviar_Click(object sender, EventArgs e)
        {           
            if (txt_destino.Text.Length == 0 || cmb_equipamento.Text.Length == 0 || txt_volume.Text.Length == 0)
            {
                MessageBox.Show("Preencha os campos obrigatorios \r\n \r\n Destino\r\n Equipamento\r\n Volume", 
                    "Dados incompletos!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                if (txt_numeroDeSerie.Text == verifica)
                {
                    MessageBox.Show("Já foi solicitado nota para este número de série", "Atenção!",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    string[] envioPara = File.ReadAllLines(@"C:DPaschoal\DP-Field\Configuracao_de_Email.txt");
                    string[] copiaPara = File.ReadAllLines(@"C:DPaschoal\DP-Field\Configuracao_de_Email.txt");
                    txt_enviarPara.Text = envioPara[0];
                    txt_copiaPara.Text = copiaPara[1];

                    string destino = ""; //Numero da loja destino para titulo do email
                    string emailCopia = copiaPara[1];
                    string lojaCdDestino = ""; //Numero da loja Corpo do email

                    if (txt_destino.Text.Length == 2)
                    {
                        destino = "CD " + txt_destino.Text + " - " + cmb_equipamento.Text;
                        lojaCdDestino = "CD " + txt_destino.Text;
                    }
                    if (txt_destino.Text.Length == 3)
                    {
                        destino = "Loja " + txt_destino.Text + " - " + cmb_equipamento.Text;
                        lojaCdDestino = "loja " + txt_destino.Text;
                    }

                    string destinatario = envioPara[0];
                    string remetente = "suporte@dpaschoal.com.br";

                    Email = new MailMessage();
                    Email.To.Add(new MailAddress(destinatario));                   
                    Email.From = (new MailAddress(remetente));
                    Email.CC.Add(new MailAddress(emailCopia));
                    Email.Subject = "Nota Fiscal " + destino;
                    Email.IsBodyHtml = true;
                    Email.Body = "Bom dia," +
                        "<br />" + "Favor gerar nota fiscal para o seguinte pedido:" +
                        "<br /><br />" + "<b>Destino:</b> " + lojaCdDestino +
                        "<br />" + "<b>Equipamento:</b> " + cmb_equipamento.Text +
                        "<br />" + "<b>Codigo SAP:</b> " + txt_codSap.Text +
                        "<br />" + "<b>Peso:</b> " + txt_peso.Text +
                        "<br />" + "<b>Valor:</b> " + txt_valor.Text +
                        "<br />" + "<b>Marca:</b> " + txt_marca.Text +
                        "<br />" + "<b>Modelo:</b> " + txt_modelo.Text +
                        "<br />" + "<b>HostName:</b> " + txt_hostName.Text +
                        "<br />" + "<b>SE Novo:</b> " + txt_seNovo.Text +
                        "<br />" + "<b>SE Antigo:</b> " + txt_seAntigo.Text +
                        "<br />" + "<b>Numero de Serie:</b> " + txt_numeroDeSerie.Text +
                        "<br />" + "<b>Chamado:</b> " + txt_chamado.Text +
                        "<br />" + "<b>Forma de Envio:</b> " + txt_formaDeEnvio.Text +
                        "<br /><br />" + "<b>Solicitante:</b> " + txt_solicitante.Text +
                        "<br />" + "<b>Centro de custo TI:</b> " + txt_centroDeCustoTi.Text +
                        "<br />" + "<b>Centro de Custo:</b> " + txt_centroDeCusto.Text +
                        "<br />" + "<b>Volume:</b> " + txt_volume.Text +
                        "<br /><br /><b>att,</b><br />" + "<b>" + txt_nome.Text + "</b>";

                    SmtpClient cliente = new SmtpClient("smtp.office365.com", 587);
                    using (cliente)
                    {
                        cliente.Credentials = new NetworkCredential(remetente, "comercial");
                        cliente.EnableSsl = true;
                        cliente.Send(Email);
                    }
                    MessageBox.Show("Email enviado com sucesso", "E-mail Enviado! ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    verifica = txt_numeroDeSerie.Text;
                } 
            }                
        }        
        public void cmb_equipamento_TextChanged(object sender, EventArgs e)
        {
            if (cmb_equipamento.Text.Length == 0)
            {
                txt_codSap.Text = "";
            }
            if (cmb_equipamento.Text == "DESKTOP")
            {
                txt_codSap.Text = "5105790";
                txt_peso.Text = "5 Kg";
                txt_valor.Text = "R$ 2.000,00";
                txt_marca.Text = "";
                txt_modelo.Text = "";
            }
            if (cmb_equipamento.Text == "MONITOR")
            {
                txt_codSap.Text = "5105536";
                txt_peso.Text = "2 Kg";
                txt_valor.Text = "R$ 500,00";
                txt_marca.Text = "";
                txt_modelo.Text = "";
            }
            if (cmb_equipamento.Text == "NOTEBOOK")
            {
                txt_codSap.Text = "5111358";
                txt_peso.Text = "4 Kg";
                txt_valor.Text = "R$ 3.500,00";
                txt_marca.Text = "";
                txt_modelo.Text = "";
            }
            if (cmb_equipamento.Text == "NOBREAK")
            {
                txt_codSap.Text = "9999724";
                txt_peso.Text = "40 Kg";
                txt_valor.Text = "R$ 1.200,00";
                txt_marca.Text = "";
                txt_modelo.Text = "";
            }
            if (cmb_equipamento.Text == "TELEFONE FIXO")
            {
                txt_codSap.Text = "9999949";
                txt_peso.Text = "1 Kg";
                txt_valor.Text = "R$ 300,00";
                txt_marca.Text = "";
                txt_modelo.Text = "";
            }
            if (cmb_equipamento.Text == "TELEFONE SEM FIO")
            {
                txt_codSap.Text = "5106672";
                txt_peso.Text = "1 Kg";
                txt_valor.Text = "R$ 150,00";
                txt_marca.Text = "";
                txt_modelo.Text = "";
            }
            if (cmb_equipamento.Text == "HEADSET")
            {
                txt_codSap.Text = "5101735";
                txt_peso.Text = "1 Kg";
                txt_valor.Text = "R$ 100,00";
                txt_marca.Text = "";
                txt_modelo.Text = "";
            }
            if (cmb_equipamento.Text == "ESTABILIZADOR")
            {
                txt_codSap.Text = "9999447";
                txt_peso.Text = "5 Kg";
                txt_valor.Text = "R$ 200,00";
                txt_marca.Text = "";
                txt_modelo.Text = "";
            }
            if (cmb_equipamento.Text == "ECF IMPRESSORA")
            {
                txt_codSap.Text = "5107083";
                txt_peso.Text = "5 Kg";
                txt_valor.Text = "R$ 800,00";
                txt_marca.Text = "";
                txt_modelo.Text = "";
            }
            if (cmb_equipamento.Text == "M521 IMPRESSORA")
            {
                txt_codSap.Text = "9999611";
                txt_peso.Text = "20 Kg";
                txt_valor.Text = "R$ 1.070,00";
                txt_marca.Text = "";
                txt_modelo.Text = "";
            }
            if (cmb_equipamento.Text == "HP3015 IMPRESSORA")
            {
                txt_codSap.Text = "9999611";
                txt_peso.Text = "20 Kg";
                txt_valor.Text = "R$ 1.070,00";
                txt_marca.Text = "";
                txt_modelo.Text = "";
            }
            if (cmb_equipamento.Text == "BANDEJA HP3015")
            {
                txt_codSap.Text = "5121761";
                txt_peso.Text = "2 Kg";
                txt_valor.Text = "R$ 1000,00";
                txt_marca.Text = "";
                txt_modelo.Text = "";
            }
            if (cmb_equipamento.Text == "HD")
            {
                txt_codSap.Text = "9999599";
                txt_peso.Text = "1 Kg";
                txt_valor.Text = "R$ 300,00";
                txt_marca.Text = "SEAGATE";
                txt_modelo.Text = "BARRACUDA";
            }
            if (cmb_equipamento.Text == "INTERFACE")
            {
                txt_codSap.Text = "5106699";
                txt_peso.Text = "1 Kg";
                txt_valor.Text = "R$ 400,00";
                txt_marca.Text = "INTELBRAS";
                txt_modelo.Text = "ITC4100";
            }
        }
        public void btn_configuracoes_Click(object sender, EventArgs e)
        {
            Hide();

            Directory.CreateDirectory(@"C:\DPaschoal\DP-Field");
            string caminho = @"C:\DPaschoal\DP-Field\Configuracao_de_Email.txt";
            if (!File.Exists(caminho))
            {
                using (StreamWriter padrao = new StreamWriter(caminho))
                {
                    padrao.WriteLine("raquel.nogueira@dpaschoal.com.br");
                    padrao.WriteLine("ti.laboratorio@dpaschoal.com.br");
                }
                MessageBox.Show("Aplicando configurações Padrao! \r\n \r\n" +
                    "Criando Diretorio em: \r\nC:\\DPaschoal\\DP-Field", "Configurando!");                             
            }
            Configuracoes config = new Configuracoes();
            config.Show();
        }       
        public void Configuracoes_FormClosed(object sender, FormClosingEventArgs e)
        {
            this.Show();
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
