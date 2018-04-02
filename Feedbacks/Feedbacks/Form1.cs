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
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
namespace Feedbacks
{
    public partial class Form1 : Form
    {
        string origen, destino;
        string[] lstArchivosOriginales;
        List<string> lstArchivos = new List<string>();

        List<string> lstNombres = new List<string>();

        int columnas = 20;

        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public Form1()
        {
            InitializeComponent();
        }

        private void Ejecutar_Click(object sender, EventArgs e)
        {
            leerArchivos();
            abrirArchivo();

            label3.Text = lstArchivos[0];
            //label3.Text = lstNombres[0];

        }

        public void leerArchivos()
        {
            origen = txtOrigen.Text;
            destino = txtDestino.Text;

            lstArchivosOriginales = Directory.GetFiles(txtOrigen.Text, "*.csv").Select(Path.GetFileName).ToArray(); ;
            if (lstArchivosOriginales.Length != 0)
            {
                for (int i = 0; i < lstArchivosOriginales.Length; i++)
                    lstArchivos.Add(lstArchivosOriginales[i].ToString());
            }
            else
            {
                return;
            }
                             
        }

        public void abrirArchivo()
        {
            List<string> varCol = new List<string>();
            origen = txtOrigen.Text;
            string nombre;
            int flg_copia;

            varCol.Add("ICOMMKT_Date");
            varCol.Add("Email");
            varCol.Add("ICOMMKT_Transport");
            varCol.Add("ICOMMKT_Profile");
            varCol.Add("id_ripley");
            varCol.Add("Nombre_Archivo");

            foreach (string archivo in lstArchivos)
            {
                nombre = archivo.Substring(0, (archivo.IndexOf("."))) + "_FB";
                wb = excel.Workbooks.Open(origen + "\\" + archivo);
                lstNombres.Add(nombre);



            }
        }
    }
}
