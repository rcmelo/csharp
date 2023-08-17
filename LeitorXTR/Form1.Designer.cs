namespace XmlNfeLerCsharp
{
    partial class Form1
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.lblNumLote = new System.Windows.Forms.TextBox();
            this.lblCompLote = new System.Windows.Forms.TextBox();
            this.lblTipoTransacao = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblNfValorTotal = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.listView1 = new System.Windows.Forms.ListView();
            this.Item = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.id = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.decricao = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.qtde = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.unitario = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.valor = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.label6 = new System.Windows.Forms.Label();
            this.lblRegistroAns = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.label1.Location = new System.Drawing.Point(273, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(368, 50);
            this.label1.TabIndex = 0;
            this.label1.Text = "Leitor XTR e XTE";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(32, 116);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 16);
            this.label2.TabIndex = 1;
            this.label2.Text = "Nº Lote";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(363, 450);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(127, 45);
            this.button1.TabIndex = 2;
            this.button1.Text = "Abrir";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lblNumLote
            // 
            this.lblNumLote.Location = new System.Drawing.Point(142, 113);
            this.lblNumLote.Name = "lblNumLote";
            this.lblNumLote.Size = new System.Drawing.Size(211, 22);
            this.lblNumLote.TabIndex = 3;
            this.lblNumLote.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // lblCompLote
            // 
            this.lblCompLote.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblCompLote.Location = new System.Drawing.Point(507, 78);
            this.lblCompLote.Name = "lblCompLote";
            this.lblCompLote.Size = new System.Drawing.Size(217, 22);
            this.lblCompLote.TabIndex = 4;
            this.lblCompLote.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // lblTipoTransacao
            // 
            this.lblTipoTransacao.Location = new System.Drawing.Point(142, 77);
            this.lblTipoTransacao.Name = "lblTipoTransacao";
            this.lblTipoTransacao.Size = new System.Drawing.Size(211, 22);
            this.lblTipoTransacao.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(385, 80);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(116, 16);
            this.label3.TabIndex = 6;
            this.label3.Text = "Competencia Lote";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(32, 80);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(104, 16);
            this.label4.TabIndex = 7;
            this.label4.Text = "Tipo Transação";
            this.label4.Click += new System.EventHandler(this.label4_Click);
            // 
            // lblNfValorTotal
            // 
            this.lblNfValorTotal.Location = new System.Drawing.Point(168, 408);
            this.lblNfValorTotal.Name = "lblNfValorTotal";
            this.lblNfValorTotal.Size = new System.Drawing.Size(185, 22);
            this.lblNfValorTotal.TabIndex = 8;
            this.lblNfValorTotal.TextChanged += new System.EventHandler(this.lblNfValorTotal_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(47, 411);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(115, 16);
            this.label5.TabIndex = 9;
            this.label5.Text = "Valor total da NFe";
            // 
            // listView1
            // 
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Item,
            this.id,
            this.decricao,
            this.qtde,
            this.unitario,
            this.valor});
            this.listView1.FullRowSelect = true;
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(12, 188);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(844, 211);
            this.listView1.TabIndex = 10;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // Item
            // 
            this.Item.Text = "Item";
            this.Item.Width = 40;
            // 
            // id
            // 
            this.id.Text = "ID";
            this.id.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // decricao
            // 
            this.decricao.Text = "Descrição do produto";
            this.decricao.Width = 270;
            // 
            // qtde
            // 
            this.qtde.Text = "Quantidade";
            this.qtde.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.qtde.Width = 100;
            // 
            // unitario
            // 
            this.unitario.Text = "Unitário";
            this.unitario.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.unitario.Width = 100;
            // 
            // valor
            // 
            this.valor.Text = "valor";
            this.valor.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.valor.Width = 120;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(385, 119);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(108, 20);
            this.label6.TabIndex = 11;
            this.label6.Text = "RegistroANS";
            this.label6.Click += new System.EventHandler(this.label6_Click);
            // 
            // lblRegistroAns
            // 
            this.lblRegistroAns.Location = new System.Drawing.Point(507, 116);
            this.lblRegistroAns.Name = "lblRegistroAns";
            this.lblRegistroAns.Size = new System.Drawing.Size(217, 22);
            this.lblRegistroAns.TabIndex = 12;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(868, 507);
            this.Controls.Add(this.lblRegistroAns);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.lblNfValorTotal);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblTipoTransacao);
            this.Controls.Add(this.lblCompLote);
            this.Controls.Add(this.lblNumLote);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Leitor XML NFe SEFAZ";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox lblNumLote;
        private System.Windows.Forms.TextBox lblCompLote;
        private System.Windows.Forms.TextBox lblTipoTransacao;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox lblNfValorTotal;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader Item;
        private System.Windows.Forms.ColumnHeader id;
        private System.Windows.Forms.ColumnHeader decricao;
        private System.Windows.Forms.ColumnHeader qtde;
        private System.Windows.Forms.ColumnHeader unitario;
        private System.Windows.Forms.ColumnHeader valor;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox lblRegistroAns;
    }
}

