using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OPCCommunication
{
    public partial class Form1 : Form
    {
        private Opc.Da.Server server = null;
        private OpcCom.Factory factor = new OpcCom.Factory();
        private Opc.Da.Item[] items;
        private Opc.Da.Subscription group;
        private Opc.IRequest req;

        public Form1()
        {
            InitializeComponent();

            this.Size = new Size(455, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                OpcCom.ServerEnumerator serverEnumerator = new OpcCom.ServerEnumerator();
                Opc.Server[] servers = serverEnumerator.GetAvailableServers(Opc.Specification.COM_DA_30);

                listBox1.Items.Clear();
                foreach (Opc.Server tmp in servers)
                    listBox1.Items.Add(tmp.Url.ToString());
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if ((server != null) && (server.IsConnected))
                server.Disconnect();

            if (listBox1.SelectedItems.Count != 0)
            {
                Opc.URL url = new Opc.URL(listBox1.SelectedItem.ToString());
                server = new Opc.Da.Server(factor, null);

                try
                {
                    server.Connect(url, new Opc.ConnectData(new System.Net.NetworkCredential()));
                }
                catch (Exception Ex)
                {
                    MessageBox.Show("Connection Error:" + Ex.Message);
                    return;
                }

                Opc.Da.SubscriptionState subscriptionState = new Opc.Da.SubscriptionState();
                group = (Opc.Da.Subscription)server.CreateSubscription(subscriptionState);
                items = new Opc.Da.Item[2];
                items[0] = new Opc.Da.Item();
                items[0].ItemName = "Bucket Brigade.Int1";
                items[1] = new Opc.Da.Item();
                items[1].ItemName = "Bucket Brigade.Int2";
                items = group.AddItems(items);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Opc.Da.ItemValue[] itemValues = new Opc.Da.ItemValue[group.Items.Count()];
            Random random = new Random();

            for (int i = 0; i < group.Items.Count(); i++)
            {
                Opc.Da.ItemValue itemValue = new Opc.Da.ItemValue(group.Items[i].ItemName);
                itemValue.ItemPath = group.Items[i].ItemPath;
                itemValue.ServerHandle = group.Items[i].ServerHandle;
                itemValue.Value = random.Next(0, 100);
                itemValues[i] = itemValue;
            }
            group.Write(itemValues, 1, WriteCompleteCallback, out req);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            group.Read(group.Items, 0, ReadCompleteCallback, out req);
        }

        void ReadCompleteCallback(object clientHandle, Opc.Da.ItemValueResult[] results)
        {
            foreach (Opc.Da.ItemValueResult readResult in results)
            {
                listBox2.BeginInvoke((MethodInvoker)delegate
                {
                    listBox2.Items.Insert(0, readResult.Value);
                });
            }
        }

        void WriteCompleteCallback(object clientHandle, Opc.IdentifiedResult[] results)
        {
            foreach (Opc.IdentifiedResult writeResult in results)
            {
                listBox2.BeginInvoke((MethodInvoker)delegate
                {
                    listBox2.Items.Insert(0, writeResult.ResultID);
                });
            }
        }
    }
}
