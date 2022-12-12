//using Azure;
using IPBSSPWebServices.Layouts.WebServices;
using IPBSSPWebServices.Utilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Hosting;
using System.Windows.Forms;

namespace CollateFilesToPDFUtility
{
    public partial class Form1 : Form
    {
        string root=string.Empty;
        public Form1()
        {
            InitializeComponent();

        }

        private void AddFiles(object sender, EventArgs e)
        {
            try
            {
                Form1 f1 = new Form1();
                openFileDialog1.Multiselect = true;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    foreach (String file in openFileDialog1.FileNames)
                    {
                        int rowcount = dataGridView1.RowCount;
                        dataGridView1.Rows.Add(file, rowcount, true);
                        
                    }
                    button3.Show();
                    button4.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in Add Files : "+ex.Message.ToString());
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                FolderBrowserDialog folderDlg = new FolderBrowserDialog();
                folderDlg.ShowNewFolderButton = true;
                // Show the FolderBrowserDialog.  
                DialogResult result = folderDlg.ShowDialog();
                if (result == DialogResult.OK)
                {
                    label2.Show();
                    textBox1.Show();
                    textBox1.Text = folderDlg.SelectedPath;
                    root = folderDlg.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in button1_Click : " + ex.Message.ToString());
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(root))
                {
                    MessageBox.Show("Please Select Destination Path .");
                    return;

                }
                string FileName = textBox2.Text;
                if (string.IsNullOrEmpty(FileName))
                {
                    MessageBox.Show("Please Enter Name For Merged File .");
                    return;
                }
                Int32 TableCount = dataGridView1.Rows.GetRowCount(DataGridViewElementStates.Visible);
                if (TableCount > 0)
                {
                    var checkedcount = 0;
                    List<int> sequence=new List<int>(); ;
                    List<DocumentEntity> DocList = new List<DocumentEntity>();
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.Cells[2].Value != null && (bool)(row.Cells[2].Value) == true)
                        {
                            if (row.Cells[0].Value == null || row.Cells[0].Value.ToString() == "")
                            {
                                MessageBox.Show("Please Select File For Checked Rows");
                                return;
                            }

                            if (row.Cells[1].Value == null || row.Cells[1].Value.ToString() == "")
                            {
                                MessageBox.Show("Please Enter Sequence Number For Checked Rows");
                                return;
                            }
                            if (sequence.Contains(Convert.ToInt32(row.Cells[1].Value)))
                            {
                                MessageBox.Show("Duplicate Sequence Found !.");
                                return;
                            }
                            else
                            {
                                sequence.Add(Convert.ToInt32(row.Cells[1].Value));
                            }
                            checkedcount++;

                        }
                    }
                    if (checkedcount > 0)
                    {
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.Cells[2].Value != null && (bool)(row.Cells[2].Value) == true)
                            {
                                DocList.Add(new DocumentEntity() { docUrl = row.Cells[0].Value.ToString(), sequence = Convert.ToInt32(row.Cells[1].Value) });
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please Select Checkbox Against Rows To Collate");
                        return;
                    }
                    Form2 f2 = new Form2();
                    f2.Show();
                    IPBSSPDocumentService docser = new IPBSSPDocumentService();
                    string merged = docser.CollateFilesInPdf(DocList, root, FileName);
                    f2.Hide();
                    this.Show();
                    if (merged != null && merged == "Files Merged and Saved To Destination Folder!.")
                    {
                        MessageBox.Show(merged);
                        dataGridView1.DataSource = null;
                        dataGridView1.Rows.Clear();
                        dataGridView1.Refresh();
                        root = string.Empty;
                        textBox1.Clear();
                        textBox2.Clear();
                        textBox1.Hide();
                        label2.Hide();

                    }
                    else
                    {
                        MessageBox.Show(merged);
                    }
                }
                else
                {
                    MessageBox.Show("Please Select 1 or More files to Collate To PDF !.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in button3_Click : " + ex.Message.ToString());
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                var confirmResult = MessageBox.Show("Are you sure to Clear All Files ?",
                                         "Confirm Clear!!",
                                         MessageBoxButtons.YesNo);
                if (confirmResult == DialogResult.Yes)
                {
                    dataGridView1.DataSource = null;
                    dataGridView1.Rows.Clear();
                    dataGridView1.Refresh();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in button4_Click : " + ex.Message.ToString());
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
