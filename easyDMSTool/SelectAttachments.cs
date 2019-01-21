using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace easyDMSTool
{
    public partial class SelectAttachments : Form
    {
        public List<string> selectedAttachments = new List<string>();
        private bool isChecked = true;
        public bool withMailBody = true;
        public SelectAttachments(List<string> attachments)
        {
            InitializeComponent();
            checkedListBox1.Items.Clear();
            foreach (string item in attachments)
            {
                //if(!item.Contains("MOVED___"))
                    checkedListBox1.Items.Add(item);
            }
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if(!checkedListBox1.Items[i].ToString().Contains("MOVED_"))
                    checkedListBox1.SetItemChecked(i, true);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count > 0)
            {
                foreach (string item in checkedListBox1.CheckedItems)
                {
                    selectedAttachments.Add(item);
                    this.DialogResult = System.Windows.Forms.DialogResult.OK;
                    this.Close();
                }
            }
            else
            {
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, isChecked);
                }
            isChecked = !isChecked;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxMailBody.Checked)
                this.withMailBody = true;
            else
                this.withMailBody = false;
         }
    }
}
