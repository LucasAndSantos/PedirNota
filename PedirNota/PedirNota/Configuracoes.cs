using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace PedirNota
{
    public partial class Configuracoes : Form
    {
        public Configuracoes()
        {
            InitializeComponent();
        }
        public void Configuracoes_Load(object sender, EventArgs e)
        {
            string[] envioPadrao = File.ReadAllLines(@"C:\DPaschoal\DP-Field\Configuracao_de_Email.txt");
            string[] copiaPadrao = File.ReadAllLines(@"C:\DPaschoal\DP-Field\Configuracao_de_Email.txt");

            txt_emailEnvio.Text = envioPadrao[0];
            txt_emailCopia.Text = copiaPadrao[1];           
        }
        public void button1_Click(object sender, EventArgs e)
        {
            if (txt_emailEnvioAlterar.Text.Length == 0)
            {
                MessageBox.Show("Sem endereço de envio?\r\n \r\n " +
                    "Defina um novo email antes de salvar as modificações!",
                    "Informações incompletas!", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
            else
            {
                MessageBox.Show("Salvar alterações em: \r\n \r\n " +
                    "C:\\DP-Field\\Configuracao_de_Email.txt");
            
                SaveFileDialog salvar = new SaveFileDialog();                
                salvar.FileName = "Configuracao_de_Email.txt";
                salvar.DefaultExt = "txt";
                DialogResult salvou = salvar.ShowDialog();
                if (salvou == DialogResult.OK)
                {
                    File.WriteAllText(salvar.FileName, txt_emailEnvioAlterar.Text + 
                        "\r\n" + txt_emailCopiaAlterar.Text);
                }
            }
        }
        public void button2_Click(object sender, EventArgs e)
        {            
            Close();
        }
        public void button3_Click(object sender, EventArgs e)
        {               
            txt_emailEnvioAlterar.Text = "raquel.nogueira@dpaschoal.com.br";
            txt_emailCopiaAlterar.Text = "ti.laboratorio@dpaschoal.com.br";

            MessageBox.Show("Não Esqueça de salvar as alterações!", "Atenção!",
                MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
        }
        public void Configuracoes_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form1 f = new Form1();
            f.Show();
        }
        public void Configuracoes_FormClosing(object sender, FormClosingEventArgs e)
        {}        
    }
}
