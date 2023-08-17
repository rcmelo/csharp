using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace XmlNfeLerCsharp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void lblNfValorTotal_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            LerXTR();
        }
        private void LerXTR() 
        {
            var arquivo = @"C:\xml\Estrutura_XML.xml";
            //var identificacaoTransacao = "";
            var tipoTransacao = "";
            var numeroLote = "";
            var competenciaLote = "";
            var registroANS = "";
            //var vProd = "";


            using (XmlReader meuXml = XmlReader.Create(arquivo)) 
            {
                var fimItens = false;

                while (meuXml.Read()) 
                {
                    //--- Cabeçalho
                    if(meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "tipoTransacao")
                        lblTipoTransacao.Text = meuXml.ReadElementString();

                    if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "numeroLote")
                        lblNumLote.Text = meuXml.ReadElementString();
                    
                    if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "competenciaLote")
                        lblCompLote.Text = meuXml.ReadElementString();

                    if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "registroANS")
                        lblRegistroAns.Text = meuXml.ReadElementString();

                    //-- itens XTR
                    if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "registroANS")
                    {
                        //lendo atributos da tag <det>
                        item = meuXml.GetAttribute("nItem");

                        //variáveis para guardar o conteúdo do item                   
                        cProd = "";
                        xProd = "";
                        qCom = "";
                        vUnCom = "";
                        vProd = "";
                    }
                    else if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "total")
                    {
                        fimItens = true;
                    }

                    if(!fimItens) 
                    {
                        if(meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "cProd")
                            cProd = meuXml.ReadElementString();

                        //if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "cEan")
                        //    cEan = meuXml.ReadElementString();

                        if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "xProd")
                            xProd = meuXml.ReadElementString();

                        if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "qCom")
                            qCom = meuXml.ReadElementString().Replace(".", ",");

                        if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "vUnCom")
                            vUnCom = meuXml.ReadElementString().Replace(".", ",");

                        if (meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "vProd")
                        {
                            vProd = meuXml.ReadElementString().Replace(".", ",");
                            listView1.Items.Add(new ListViewItem(new[] { item, cProd, xProd, qCom, vUnCom.ToString(), vProd.ToString().Replace(".", ",") }));
                        }
                    }

                    //--Rodape
                    if(meuXml.NodeType == XmlNodeType.Element && meuXml.Name == "vNF")
                        lblNfValorTotal.Text = meuXml.ReadElementString().Replace(".",",");

                }
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }
    }
}
