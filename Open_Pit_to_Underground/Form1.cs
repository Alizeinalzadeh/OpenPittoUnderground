using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Open_Pit_to_Underground
{
    public partial class frmopenpittoindergroundtransition : Form
    {
        public frmopenpittoindergroundtransition()
        {
            InitializeComponent();
        }
        double woey, tax, r, tov, tdevo, tdevu, cdevo, cov, ceo, ceu, cwo, cso, cp, cs, ct, co, ro, ru, rp, rs, i, ddo, ddu, cy, g, p, wmetaly;
        double cpoy, csy, ctpy, coty, cpuy, io, iu;
        double[] ww = new double[100];
        double[] woo = new double[100];
        double[] wou = new double[100];
        double[] npvt = new double[100];
        double[] ws = new double[100];
        double[] uc = new double[100];
        double[] to = new double[100];
        double[] tu = new double[100];
        double[] cwoy = new double[100];
        double[] ceoy = new double[100];
        double[] wwy = new double[100];
        double[] ctu = new double[100];
        double[] ceuy = new double[100];
        double[] cto = new double[100];
        double[] npvo = new double[100];
        double[] nvu = new double[100];
        double[] npvu = new double[100];
        double[] dcfo = new double[100];
        double[] dcfu = new double[100];
        double[] cfu = new double[100];
        double[] cfo = new double[100];
        double[] taxu = new double[100];
        double[] taxo = new double[100];
        double[] wovy = new double[100];
        double[] covy = new double[100];
        double[] cdoy = new double[100];
        double[] cduy = new double[100];
        double[] cdsy = new double[100];

        double cdo, cdu, cds;




        //اجرا شدن اولیه برنامه
        private void Form1_Load(object sender, EventArgs e)
        {
             
            btnapplyalternative.Enabled = false;
            btnapplyore.Enabled = false;
            btnapplydevelopment.Enabled = false;
            btnapplycost.Enabled = false;
            btnapplyeffi.Enabled = false;
            btnapplyother.Enabled = false;

            btnclearalternative.Enabled = false;
            btnclearore.Enabled = false;
            btncleardevelopment.Enabled = false;
            btnclearcost.Enabled = false;
            btncleareffi.Enabled = false;
            btnclearore.Enabled = false;
            
        }

        /////////////////////////////clear/////////////////////////
        //clear alternative
        private void btnclearalternative_Click(object sender, EventArgs e)
        {
            DialogResult ans;
            ans = MessageBox.Show("Are You Sure?", "Attention", MessageBoxButtons.YesNo);
            if (ans == DialogResult.Yes)
            {
                foreach (Control control in this.tabalternatives.Controls)
                    if (control is TextBox)
                        (control as TextBox).Clear();
            }

            else
            {
                txtalternative1.Focus();
                txtalternative1.SelectAll();
            }
        }

        //clear ore details
        private void btnclearore_Click(object sender, EventArgs e)
        {
            DialogResult ans;
            ans = MessageBox.Show("Are You Sure?", "Attention", MessageBoxButtons.YesNo);
            if (ans == DialogResult.Yes)
            {
                foreach (Control control in this.tabore.Controls)
                    if (control is TextBox)
                        (control as TextBox).Clear();
            }

            else
            {
                txtreserve.Focus();
                txtreserve.SelectAll();
            }
        }

        //clear development
        private void btncleardevelopment_Click(object sender, EventArgs e)
        {
            DialogResult ans;
            ans = MessageBox.Show("Are You Sure?", "Attention", MessageBoxButtons.YesNo);
            if (ans == DialogResult.Yes)
            {
                foreach (Control control in this.tabore.Controls)
                    if (control is TextBox)
                        (control as TextBox).Clear();
            }

            else
            {
                nmbroverburden.Focus();
            }
        }

        //clear costs
        private void btnclearcost_Click(object sender, EventArgs e)
        {
            DialogResult ans;
            ans = MessageBox.Show("Are You Sure?", "Attention", MessageBoxButtons.YesNo);
            if (ans == DialogResult.Yes)
            {
                foreach (Control control in this.tabcosts.Controls)
                    if (control is TextBox)
                        (control as TextBox).Clear();
            }

            else
            {
                txtopenpitcost.Focus();
                txtopenpitcost.SelectAll();
            }
        }

        //clear efficiency
        private void btncleareffi_Click(object sender, EventArgs e)
        {
            DialogResult ans;
            ans = MessageBox.Show("Are You Sure?", "Attention", MessageBoxButtons.YesNo);
            if (ans == DialogResult.Yes)
            {
                foreach (Control control in this.tabefficiency.Controls)
                    if (control is TextBox)
                        (control as TextBox).Clear();
            }

            else
            {
                txtopenpitefficiency.Focus();
                txtopenpitefficiency.SelectAll();
            }
        }

        private void label21_Click(object sender, EventArgs e)
        {
        }

        //clear other
        private void btnclearother_Click(object sender, EventArgs e)
        {
            DialogResult ans;
            ans = MessageBox.Show("Are You Sure?", "Attention", MessageBoxButtons.YesNo);
            if (ans == DialogResult.Yes)
            {
                foreach (Control control in this.tabother.Controls)
                    if (control is TextBox)
                        (control as TextBox).Clear();
            }

            else
            {
                txti.Focus();
                txti.SelectAll();
            }
        }

        //clear result
        private void btnclearallresult_Click(object sender, EventArgs e)
        {
            DialogResult ans;
            ans = MessageBox.Show("Are You Sure? \nAll Your Data Will Be Lost!!!", "Attention", MessageBoxButtons.YesNo);
            if (ans == DialogResult.Yes)
            {
                foreach (Control control in this.tabalternatives.Controls)
                    if (control is TextBox)
                        (control as TextBox).Clear();

                foreach (Control control in this.tabore.Controls)
                    if (control is TextBox)
                        (control as TextBox).Clear();

                foreach (Control control in this.tabcosts.Controls)
                    if (control is TextBox)
                        (control as TextBox).Clear();

                foreach (Control control in this.tabefficiency.Controls)
                    if (control is TextBox)
                        (control as TextBox).Clear();

                foreach (Control control in this.tabother.Controls)
                    if (control is TextBox)
                        (control as TextBox).Clear();

                btnprint.Enabled = false;
                dataGridView1.Rows.Clear();
                dataGridView2.Rows.Clear();

                lblnpvo1result.Text = "No Data";
                lblnpvo2result.Text = "No Data";
                lblnpvu1result.Text = "No Data";
                lblnpvu2result.Text = "No Data";
                lblnpvt1result.Text = "No Data";
                lblnpvt2result.Text = "No Data";
                lblfinalresult.Text = "No Data";
            }
        }

        //////////////////////////////Exit//////////////////////////////
        //exit alternative
        private void btnexitalternative_Click(object sender, EventArgs e)
        {
            
            DialogResult ans;
            ans = MessageBox.Show("Do You Wanna Exit?", "Exit", MessageBoxButtons.YesNo);
            if (ans == DialogResult.Yes)
                Application.Exit();

            
            


        }

        //exit ore detail
        private void btnexitore_Click(object sender, EventArgs e)
        {
            btnexitalternative_Click(null, null);
        }

        //exit costs
        private void btnexitcost_Click(object sender, EventArgs e)
        {
            btnexitalternative_Click(null, null);
        }

        //exit development
        private void btnexitdevelopment_Click(object sender, EventArgs e)
        {
            btnexitalternative_Click(null, null);
        }

        //exit efficiency
        private void btnexiteffi_Click(object sender, EventArgs e)
        {
            btnexitalternative_Click(null, null);
        }

        //exit other
        private void btnexitother_Click(object sender, EventArgs e)
        {
            btnexitalternative_Click(null, null);
        }

        //exit result
        private void btnexitresult_Click(object sender, EventArgs e)
        {
            btnexitalternative_Click(null, null);
        }

        ////////////////////////////// calculate/////////////////////////////////////////
        //calculate result
        private void btncalculate_Click(object sender, EventArgs e)
        {

            //داده های ورودی برای انجام محاسبات
            tov = Convert.ToDouble(nmbroverburden.Value);
            cdevo = Convert.ToDouble(txtopenpitdevelopment.Text);
            tdevo = Convert.ToDouble(nmbropenpit.Value);
            tdevu = Convert.ToDouble(nmbrunderground.Value);
            cov = Convert.ToDouble(txtoverburdencost.Text);
            tax = Convert.ToDouble(txttax.Text);
            

            if (nmbr.Value == 2)
            {
                ww[1] = Convert.ToDouble(txtwaste1.Text);
                ww[2] = Convert.ToDouble(txtwaste2.Text);
                woo[1] = Convert.ToDouble(txtopenpit1.Text);
                woo[2] = Convert.ToDouble(txtopenpit2.Text);
                ws[1] = Convert.ToDouble(txtoverburden1.Text);
                ws[2] = Convert.ToDouble(txtoverburden2.Text);
                uc[1] = Convert.ToDouble(txtundergrounddevelopment1.Text);
                uc[2] = Convert.ToDouble(txtundergrounddevelopment2.Text);


            }
            else if (nmbr.Value == 3)
            {
                ww[1] = Convert.ToDouble(txtwaste1.Text);
                ww[2] = Convert.ToDouble(txtwaste2.Text);
                ww[3] = Convert.ToDouble(txtwaste3.Text);
                woo[1] = Convert.ToDouble(txtopenpit1.Text);
                woo[2] = Convert.ToDouble(txtopenpit2.Text);
                woo[3] = Convert.ToDouble(txtopenpit3.Text);
                ws[1] = Convert.ToDouble(txtoverburden1.Text);
                ws[2] = Convert.ToDouble(txtoverburden2.Text);
                ws[3] = Convert.ToDouble(txtoverburden3.Text);
                uc[1] = Convert.ToDouble(txtundergrounddevelopment1.Text);
                uc[2] = Convert.ToDouble(txtundergrounddevelopment2.Text);
                uc[3] = Convert.ToDouble(txtundergrounddevelopment3.Text);

            }
            else if (nmbr.Value == 4)
            {
                ww[1] = Convert.ToDouble(txtwaste1.Text);
                ww[2] = Convert.ToDouble(txtwaste2.Text);
                ww[3] = Convert.ToDouble(txtwaste3.Text);
                ww[4] = Convert.ToDouble(txtwaste4.Text);
                woo[1] = Convert.ToDouble(txtopenpit1.Text);
                woo[2] = Convert.ToDouble(txtopenpit2.Text);
                woo[3] = Convert.ToDouble(txtopenpit3.Text);
                woo[4] = Convert.ToDouble(txtopenpit4.Text);
                ws[1] = Convert.ToDouble(txtoverburden1.Text);
                ws[2] = Convert.ToDouble(txtoverburden2.Text);
                ws[3] = Convert.ToDouble(txtoverburden3.Text);
                ws[4] = Convert.ToDouble(txtoverburden4.Text);
                uc[1] = Convert.ToDouble(txtundergrounddevelopment1.Text);
                uc[2] = Convert.ToDouble(txtundergrounddevelopment2.Text);
                uc[3] = Convert.ToDouble(txtundergrounddevelopment3.Text);
                uc[4] = Convert.ToDouble(txtundergrounddevelopment4.Text);


            }
            else if (nmbr.Value == 5)
            {
                ww[1] = Convert.ToDouble(txtwaste1.Text);
                ww[2] = Convert.ToDouble(txtwaste2.Text);
                ww[3] = Convert.ToDouble(txtwaste3.Text);
                ww[4] = Convert.ToDouble(txtwaste4.Text);
                ww[5] = Convert.ToDouble(txtwaste5.Text);
                woo[1] = Convert.ToDouble(txtopenpit1.Text);
                woo[2] = Convert.ToDouble(txtopenpit2.Text);
                woo[3] = Convert.ToDouble(txtopenpit3.Text);
                woo[4] = Convert.ToDouble(txtopenpit4.Text);
                woo[5] = Convert.ToDouble(txtopenpit5.Text);
                ws[1] = Convert.ToDouble(txtoverburden1.Text);
                ws[2] = Convert.ToDouble(txtoverburden2.Text);
                ws[3] = Convert.ToDouble(txtoverburden3.Text);
                ws[4] = Convert.ToDouble(txtoverburden4.Text);
                ws[4] = Convert.ToDouble(txtoverburden4.Text);
                ws[5] = Convert.ToDouble(txtoverburden5.Text);
                uc[1] = Convert.ToDouble(txtundergrounddevelopment1.Text);
                uc[2] = Convert.ToDouble(txtundergrounddevelopment2.Text);
                uc[3] = Convert.ToDouble(txtundergrounddevelopment3.Text);
                uc[4] = Convert.ToDouble(txtundergrounddevelopment4.Text);
                uc[5] = Convert.ToDouble(txtundergrounddevelopment5.Text);

            }

            r = Convert.ToDouble(txtreserve.Text);

            ceo = Convert.ToDouble(txtopenpitcost.Text);
            ceu = Convert.ToDouble(txtundergroundcost.Text);
            cwo = Convert.ToDouble(txtstrippingcost.Text);
            cso = Convert.ToDouble(txtoverburdencost.Text);
            cp = Convert.ToDouble(txtprocessingcost.Text);
            cs = Convert.ToDouble(txtsmeltingcost.Text);
            ct = Convert.ToDouble(txttrasnportationcost.Text);
            co = Convert.ToDouble(txtothercosts.Text);
            ro = Convert.ToDouble(txtopenpitefficiency.Text);
            ru = Convert.ToDouble(txtundergroundefficiency.Text);
            rp = Convert.ToDouble(txtprocessingefficiency.Text);
            rs = Convert.ToDouble(txtsmeltingefficiency.Text);
            i = Convert.ToDouble(txti.Text);
            ddo = Convert.ToDouble(txtopenpitdilution.Text);
            ddu = Convert.ToDouble(txtundergrounddilution.Text);
            cy = Convert.ToDouble(txtfactorycapacity.Text);
            g = Convert.ToDouble(txtgrade.Text);
            p = Convert.ToDouble(txtprice.Text);

            cdo = Convert.ToDouble(txtopenpitinvest.Text);
            cdu = Convert.ToDouble(txtundergroundinvest.Text);
            cds = Convert.ToDouble(txtmutualinvest.Text);

            double tdo = Convert.ToDouble(nmbropenpit.Value);
            double tdu = Convert.ToDouble(nmbrunderground.Value);
            double tdw = Convert.ToDouble(nmbroverburden.Value);


            wmetaly = Convert.ToDouble(txtfactorycapacity.Text);



            //انجام محاسبات با فرض ثابت بودن ظرفیت سالیانه ورودی کارخانه
            if (chkbxconstantcapacity.Checked == true)
            {
                //وزن کانسنگ ورودی روباز
                double wooey = cy / (1 + (ddo / 100));

                //وزن کانسنگ ورودی زیرزمینی
                double wuoey = cy / (1 + (ddu / 100));

                //وزن استخراجی روباز
                double wopoey = wooey * 100 / ro;

                //وزن استخراجی زیرزمینی
                double wupoey = wuoey * 100 / ru;

                //وزن کانسنگ زیرزمینی هر گزینه
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    wou[counter] = r - woo[counter];
                }

                //عمر روباز هر گزینه
                double[] to = new double[100];
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    to[counter] = woo[counter] / wopoey;

                }

                //عمر زیرزمینی هر گزینه
                double[] tu = new double[100];
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    tu[counter] = wou[counter] / wupoey;
                }

                //باطله برداری سالیانه هر گزینه
                double[] wwy = new double[100];
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    wwy[counter] = ww[counter] / to[counter];
                }

                //وزن فلز محصول روباز
                double wometaly = cy * g / 100 * (1 - ddo / 100) * rp / 100 * rs / 100;

                //وزن فلز محصول زیرزمینی
                double wumetaly = cy * g / 100 * (1 - ddu / 100) * rp / 100 * rs / 100;

                //هزینه استخراج روباز سالیانه
                double[] ceoy = new double[100];
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    ceoy[counter] = cy * ceo;
                }

                //هزینه باطله برداری روباز سالیانه
                double[] cwoy = new double[100];
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    cwoy[counter] = cwo * ww[counter];
                }

                //هزینه فرآوری روباز سالیانه
                double cpy = cy * cp;

                //هزینه ذوب روباز سالیانه
                double csoy = wometaly * cs;

                //هزینه حمل و نقل روباز سالیانه
                double ctpoy = ct * wometaly;

                //هزینه های غیره 
                double cotoy = co * wometaly;


                //کل هزینه های عملیاتی روباز
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    cto[counter] = ceoy[counter] + cwoy[counter] + cpoy + csy + ctpy + coty;
                }

                //هزینه استخراج زیرزمینی سالیانه

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    ceuy[counter] = cy * ceu;
                }

                //هزینه فرآوری زیرزمینی سالیانه
                cpuy = cy * cp;

                //کل هزینه های عملیاتی زیرزمینی

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    ctu[counter] = ceuy[counter] + cpuy + csy + ctpy + coty;
                }

                //درآمد سالیانه روباز
                io = wmetaly * p;

                //درآمد سالیانه زیرزمینی
                iu = wmetaly * p;


                //هزینه های سرمایه ای روباز (روباره برداری)

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    wovy[counter] = ws[counter] / tov;
                    covy[counter] = wovy[counter] * cov;
                }

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    cdoy[counter] = cdo / to[counter];
                    cduy[counter] = cdu / tu[counter];
                    cdsy[counter] = cds / to[counter] + tu[counter];
                }

                //محاسبه مالیات روباز

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    taxo[counter] = (io - cto[counter]-cdoy[counter]) * tax / 100;
                }

                //محاسبه مالیات زیرزمینی

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    taxu[counter] = (iu - ctu[counter]-cduy[counter]) * tax / 100;
                }

                //سود ناخالص روباز 

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    cfo[counter] = io - cto[counter] - taxo[counter];
                }

                //سود ناخالص زیرزمینی 

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    cfu[counter] = io - ctu[counter] - taxu[counter];
                }

                //ارزش سرمایه ای روباز هر گزینه

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    dcfo[counter] = cfo[counter] * (Math.Pow((1 + (i / 100)), to[counter]) - 1) / ((i / 100) * (Math.Pow((1 + (i / 100)), to[counter])));
                }

                // ارزش خالص سرمایه ای فعلی روباز

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    npvo[counter] = dcfo[counter] - covy[counter];
                }

                lblnpvo1result.Text = Convert.ToString(npvo[1]);
                lblnpvo2result.Text = Convert.ToString(npvo[2]);
                lblnpvo3result.Text = Convert.ToString(npvo[3]);
                lblnpvo4result.Text = Convert.ToString(npvo[4]);
                lblnpvo5result.Text = Convert.ToString(npvo[5]);


                //ارزش سرمایه ای زیرزمینی هر گزینه
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    dcfu[counter] = cfu[counter] * (Math.Pow((1 + (i / 100)), tu[counter]) - 1) / ((i / 100) * (Math.Pow((1 + (i / 100)), tu[counter])));
                }

                // ارزش خالص سرمایه ای فعلی زیرزمینی
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    nvu[counter] = dcfu[counter] - uc[counter];
                }



                //ارزش خالص سرمایه ای فعلی زیرزمینی در سال صفر پروژه 
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    npvu[counter] = nvu[counter] * (1 / (Math.Pow((1 + (i / 100)), to[counter])));
                }

                lblnpvu1result.Text = Convert.ToString(npvu[1]);
                lblnpvu2result.Text = Convert.ToString(npvu[2]);
                lblnpvu3result.Text = Convert.ToString(npvu[3]);
                lblnpvu4result.Text = Convert.ToString(npvu[4]);
                lblnpvu5result.Text = Convert.ToString(npvu[5]);

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    if (npvu[counter] <= 0)
                    {
                        npvu[counter] = 0;
                    }

                    npvt[counter] = npvo[counter] + npvu[counter];
                }
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    npvt[counter] = Math.Round(npvt[counter], 2);
                }

                //چاپ نتیجه نهایی ارزش خالص فعلی
                lblnpvt1result.Text = Convert.ToString(npvt[1]);
                lblnpvt2result.Text = Convert.ToString(npvt[2]);
                lblnpvt3result.Text = Convert.ToString(npvt[3]);
                lblnpvt4result.Text = Convert.ToString(npvt[4]);
                lblnpvt5result.Text = Convert.ToString(npvt[5]);

                //محاسبه ماکزیمم ارزش خالص فعلی
                double max, min;
                max = min = npvt[1];
                if (max < npvt[1])
                    max = npvt[1];
                else if (max < npvt[2])
                    max = npvt[2];
                else if (max < npvt[3])
                    max = npvt[3];
                else if (max < npvt[4])
                    max = npvt[4];
                else if (max < npvt[5])
                    max = npvt[5];

                // چاپ نهایی نتایج محاسبات 
                if (max == npvt[1])
                    lblfinalresult.Text = "The Best Alternative is Alternative 1 With " + txtalternative1.Text + "m" + " Depth";
                else if (max == npvt[2])
                    lblfinalresult.Text = "The Best Alternative is Alternative 2 With " + txtalternative2.Text + "m" + " Depth";
                else if (max == npvt[3])
                    lblfinalresult.Text = "The Best Alternative is Alternative 3 With " + txtalternative3.Text + "m" + " Depth";
                else if (max == npvt[4])
                    lblfinalresult.Text = "The Best Alternative is Alternative 4 With " + txtalternative4.Text + "m" + " Depth";
                else if (max == npvt[5])
                    lblfinalresult.Text = "The Best Alternative is Alternative 5 With " + txtalternative5.Text + "m" + " Depth";

                ////کل هزینه های عملیاتی روباز
                //double[] cto = new double[100];
                //for (int counter = 1; counter <= nmbr.Value; counter++)
                //{
                //    cto[counter] = ceoy[counter] + cwoy[counter] + cpy + csoy + ctpoy + cotoy;
                //}

                ////هزینه استخراج زیرزمینی سالیانه
                //double[] ceuy = new double[100];
                //for (int counter = 1; counter <= nmbr.Value; counter++)
                //{
                //    ceuy[counter] = cy * ceu;
                //}

                ////هزینه ذوب زیرزمینی سالیانه
                //double csuy = wumetaly * cs;

                ////هزینه حمل و نقل روباز سالیانه
                //double ctpuy = ct * wumetaly;

                ////هزینه های غیره 
                //double cotuy = co * wumetaly;

                ////کل هزینه های عملیاتی زیرزمینی
                //double[] ctu = new double[100];
                //for (int counter = 1; counter <= nmbr.Value; counter++)
                //{
                //    ctu[counter] = ceuy[counter] + cpy + csuy + ctpuy + cotuy;
                //}

                ////درآمد سالیانه روباز
                //double io = wometaly * p;

                ////درآمد سالیانه زیرزمینی
                //double iu = wumetaly * p;

                ////هزینه های سرمایه ای
                //double wovy = wov / tov;
                //double covy = wov * cov;

                ////محاسبه مالیات روباز
                //double[] taxo = new double[100];
                //for (int counter = 1; counter <= nmbr.Value; counter++)
                //{
                //    taxo[counter] = (io - ctu[counter]) * 0 / 100;
                //}

                ////محاسبه مالیات زیرزمینی
                //double[] taxu = new double[100];
                //for (int counter = 1; counter <= nmbr.Value; counter++)
                //{
                //    taxu[counter] = (iu - ctu[counter]) * 0 / 100;
                //}

                ////محاسبه جریان نقدینگی هر سال روباز
                //double[] cfo = new double[100];
                //double[] dcfo = new double[100];
                //for (int counter1 = 1; counter1 <= nmbr.Value; counter1++)
                //{
                //    for (int counter2 = 1; counter2 < to[counter1]; counter2++)
                //    {
                //        cfo[counter2] = io - cto[counter2] - taxo[counter2];

                //        for (int counter3 = Convert.ToInt32(nmbropenpit.Value - 1); counter3 < to[counter1]; counter3++)
                //        {
                //            dcfo[counter1] = (cfo[counter2] / Math.Pow((1 + i / 100), counter3)) + dcfo[counter1];
                //        }
                //    }
                //}

                ////هزینه آماده سازی سالیانه روباز
                //double cdevoy = cdevo / tdevo;

                ////محاسبه هزینه های سرمایه ای روباز
                //double cco = cdevoy + covy;

                //// محاسبه جریان نفدینگی تنزیل شده کل روباز
                //double[] cfdevo = new double[100];
                //double[] dcfdevo = new double[100];
                //for (int counter1 = 1; counter1 <= nmbr.Value; counter1++)
                //{
                //    for (int counter2 = 0; counter2 < nmbropenpit.Value; counter2++)
                //    {
                //        cfdevo[counter2] = (cco);
                //        dcfdevo[counter1] = cfdevo[counter2] / (Math.Pow(i / 100 + 1, counter2)) + dcfdevo[counter1];
                //    }
                //}

                ////محاسبه نهایی ارزش خالص فعلی روباز
                //double[] npvto = new double[100];
                //for (int counter = 1; counter <= nmbr.Value; counter++)
                //{
                //    npvto[counter] = dcfo[counter] + dcfdevo[counter];
                //}

                ////چاپ نتایج روباز
                //for (int counter = 1; counter <= nmbr.Value; counter++)
                //{
                //    npvto[counter] = Math.Round(npvto[counter], 2);
                //}
                //lblnpvo1result.Text = Convert.ToString(npvto[1]);
                //lblnpvo2result.Text = Convert.ToString(npvto[2]);
                //lblnpvo3result.Text = Convert.ToString(npvto[3]);
                //lblnpvo4result.Text = Convert.ToString(npvto[4]);
                //lblnpvo5result.Text = Convert.ToString(npvto[5]);

                ////محاسبه جریان نقدینگی هر سال زیرزمینی
                //double[] cfu = new double[100];
                //double[] dcfu = new double[100];
                //for (int counter1 = 1; counter1 <= nmbr.Value; counter1++)
                //{
                //    for (int counter2 = 1; counter2 < tu[counter1]; counter2++)
                //    {
                //        cfu[counter2] = iu - ctu[counter2] - taxu[counter2];

                //        for (int counter3 = Convert.ToInt32(nmbrunderground.Value - 1); counter3 < tu[counter1]; counter3++)
                //        {
                //            dcfu[counter1] = (cfu[counter2] / Math.Pow((1 + i / 100), counter3)) + dcfu[counter1];
                //        }
                //    }
                //}

                ////هزینه آماده سازی سالیانه زیرزمینی
                //double cdevuy = cdevu / tdevu;

                ////محاسبه هزینه های سرمایه ای زیرزمینی
                //double ccu = cdevuy;

                //// محاسبه جریان نفدینگی تنزیل شده کل زیرزمینی
                //double[] cfdevu = new double[100];
                //double[] dcfdevu = new double[100];
                //for (int counter1 = 1; counter1 <= nmbr.Value; counter1++)
                //{
                //    for (int counter2 = 0; counter2 < nmbrunderground.Value; counter2++)
                //    {
                //        cfdevo[counter2] = (ccu);
                //        dcfdevu[counter1] = cfdevu[counter2] / (Math.Pow(i / 100 + 1, counter2)) + dcfdevu[counter1];
                //    }
                //}

                ////محاسبه نهایی ارزش خالص فعلی روباز
                //double[] npvtu = new double[100];
                //for (int counter = 1; counter <= nmbr.Value; counter++)
                //{
                //    npvtu[counter] = dcfu[counter] + dcfdevu[counter];
                //}

                ////چاپ نتایج زیرزمینی
                //for (int counter = 1; counter <= nmbr.Value; counter++)
                //{
                //    npvtu[counter] = Math.Round(npvtu[counter], 2);
                //    npvto[counter] = Math.Round(npvto[counter], 2);
                //}
                //lblnpvu1result.Text = Convert.ToString(npvtu[1]);
                //lblnpvu2result.Text = Convert.ToString(npvtu[2]);
                //lblnpvu3result.Text = Convert.ToString(npvtu[3]);
                //lblnpvu4result.Text = Convert.ToString(npvtu[4]);
                //lblnpvu5result.Text = Convert.ToString(npvtu[5]);


                ////محاسبه نهایی ارزش خالص فعلی دو گزینه 
                //double[] npvt = new double[100];
                //for (int counter = 1; counter <= nmbr.Value; counter++)
                //{
                //    npvt[counter] = npvto[counter] + npvtu[counter];
                //}
                //for (int counter = 1; counter <= nmbr.Value; counter++)
                //{
                //    npvt[counter] = Math.Round(npvt[counter], 2);
                //}

                ////چاپ نتیجه نهایی ارزش خالص فعلی
                //lblnpvt1result.Text = Convert.ToString(npvt[1]);
                //lblnpvt2result.Text = Convert.ToString(npvt[2]);
                //lblnpvt3result.Text = Convert.ToString(npvt[3]);
                //lblnpvt4result.Text = Convert.ToString(npvt[4]);
                //lblnpvt5result.Text = Convert.ToString(npvt[5]);

                ////محاسبه ماکزیمم ارزش خالص فعلی
                //double max, min;
                //max = min = npvt[1];
                //if (max < npvt[1])
                //    max = npvt[1];
                //else if (max < npvt[2])
                //    max = npvt[2];
                //else if (max < npvt[3])
                //    max = npvt[3];
                //else if (max < npvt[4])
                //    max = npvt[4];
                //else if (max < npvt[5])
                //    max = npvt[5];

                //// چاپ نهایی نتایج محاسبات 
                //if (max == npvt[1])
                //    lblfinalresult.Text = "The Best Alternative is Alternative 1 With " + txtalternative1.Text + "m" + " Depth";
                //else if (max == npvt[2])
                //    lblfinalresult.Text = "The Best Alternative is Alternative 2 With " + txtalternative2.Text + "m" + " Depth";
                //else if (max == npvt[3])
                //    lblfinalresult.Text = "The Best Alternative is Alternative 3 With " + txtalternative3.Text + "m" + " Depth";
                //else if (max == npvt[4])
                //    lblfinalresult.Text = "The Best Alternative is Alternative 4 With " + txtalternative4.Text + "m" + " Depth";
                //else if (max == npvt[5])
                //    lblfinalresult.Text = "The Best Alternative is Alternative 5 With " + txtalternative5.Text + "m" + " Depth";



            }


            //انجام محاسبات با فرض ثابت بودن فلز استحصالی سالیانه
            if (chkbxconstantcapacity.Checked == false)
            {
                //محاسبه کانسنگ سالیانه مورد نیاز
                woey = wmetaly / (100-g) / rp / rs * 100 * 100 * 100;

                //محاسبه وزن کانسنگ سالیانه ورودی به کارخانه روباز
                double coy = woey * ((1 - (ddo / 100)));

                //محاسبه وزن کانسنگ سالیانه ورودی به کارخانه زیرزمینی
                double cuy = woey * ((1 + (ddu / 100)));

                //کانسنگ سالیانه قابل استخراج روباز

                //کانسنگ سالیانه قابل استخراج زیرزمینی

                //وزن کانسنگ زیرزمینی هر گزینه
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    wou[counter] = r - woo[counter];
                }

                //عمر روباز هر گزینه 
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    to[counter] = woo[counter]*ro/100 / woey;
                }

                //عمر زیرزمینی هر گزینه
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    tu[counter] = wou[counter]*ru/100 / woey;
                }


                //باطله برداری سالیانه هر گزینه
                
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    wwy[counter] = ww[counter] / to[counter];
                }
                 

                //هزینه استخراج روباز سالیانه
               
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    ceoy[counter] = coy * ceo;
                }

                //هزینه باطله برداری روباز سالیانه
                
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    cwoy[counter] = cwo * wwy[counter];
                }

                //هزینه فرآوری روباز سالیانه
                 cpoy = coy * cp;

                //هزینه ذوب روباز سالیانه
                 csy = wmetaly * cs;

                //هزینه حمل و نقل 
                 ctpy = ct * wmetaly;

                //هزینه های غیره 
                 coty = co;

                //کل هزینه های عملیاتی روباز
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    cto[counter] = ceoy[counter] + cwoy[counter] + cpoy + csy + ctpy + coty;
                }

                //هزینه استخراج زیرزمینی سالیانه

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    ceuy[counter] = cuy * ceu;
                }

                //هزینه فرآوری زیرزمینی سالیانه
                 cpuy = cuy * cp;

                //کل هزینه های عملیاتی زیرزمینی

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    ctu[counter] = ceuy[counter] + cpuy + csy + ctpy + coty;
                }

                //درآمد سالیانه روباز
                 io = wmetaly * p;

                //درآمد سالیانه زیرزمینی
                 iu = wmetaly * p;


                //هزینه های سرمایه ای روباز (روباره برداری)
                
                for (int counter=1;counter<=nmbr.Value;counter++)
                {
                    wovy[counter]= ws[counter] / tov;
                    covy[counter] = wovy[counter] * cov - cdevo;
                }
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    cdoy[counter] = cdo / to[counter];
                    cduy[counter] = cdu / tu[counter];
                    cdsy[counter] = cds / to[counter] + tu[counter];
                }

                //محاسبه مالیات روباز
                
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    taxo[counter] = (io - cto[counter] - cdoy[counter]) *  tax/ 100;
                }

                //محاسبه مالیات زیرزمینی
                
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    taxu[counter] = (iu - ctu[counter] - cduy[counter]) * tax / 100;
                }

                //سود ناخالص روباز 
                
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    cfo[counter] = io - cto[counter] - taxo[counter];
                }

                //سود ناخالص زیرزمینی 
                
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    cfu[counter] = io - ctu[counter] - taxu[counter];
                }

                //ارزش سرمایه ای روباز هر گزینه
                
                for (int counter=1;counter<=nmbr.Value;counter++)
                {
                    dcfo[counter] = cfo[counter] * (Math.Pow((1 + (i/100)), to[counter]) - 1) / ((i/100) * (Math.Pow((1 + (i/100)), to[counter])));
                }

                // ارزش خالص سرمایه ای فعلی روباز
                
                for (int counter=1;counter<=nmbr.Value;counter++)
                {
                    npvo[counter] = dcfo[counter] - covy[counter];
                }

                lblnpvo1result.Text = Convert.ToString(npvo[1]);
                lblnpvo2result.Text = Convert.ToString(npvo[2]);
                lblnpvo3result.Text = Convert.ToString(npvo[3]);
                lblnpvo4result.Text = Convert.ToString(npvo[4]);
                lblnpvo5result.Text = Convert.ToString(npvo[5]);


                //ارزش سرمایه ای زیرزمینی هر گزینه
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    dcfu[counter] = cfu[counter] * (Math.Pow((1 + (i / 100)), tu[counter]) - 1) / ((i / 100) * (Math.Pow((1 + (i / 100)), tu[counter])));
                }

                // ارزش خالص سرمایه ای فعلی زیرزمینی
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    nvu[counter] = dcfu[counter] - uc[counter];
                }

                

                //ارزش خالص سرمایه ای فعلی زیرزمینی در سال صفر پروژه 
                for (int counter=1;counter<=nmbr.Value;counter++)
                {
                    npvu[counter] = nvu[counter] * (1 / (Math.Pow((1 + (i / 100)), to[counter])));
                }

                lblnpvu1result.Text = Convert.ToString(npvu[1]);
                lblnpvu2result.Text = Convert.ToString(npvu[2]);
                lblnpvu3result.Text = Convert.ToString(npvu[3]);
                lblnpvu4result.Text = Convert.ToString(npvu[4]);
                lblnpvu5result.Text = Convert.ToString(npvu[5]);

                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    if (npvu[counter]<=0)
                    {
                        npvu[counter] = 0;
                    }
                    npvt[counter] = npvo[counter] + npvu[counter];
                }
                for (int counter = 1; counter <= nmbr.Value; counter++)
                {
                    npvt[counter] = Math.Round(npvt[counter], 2);
                }

                //چاپ نتیجه نهایی ارزش خالص فعلی
                lblnpvt1result.Text = Convert.ToString(npvt[1]);
                lblnpvt2result.Text = Convert.ToString(npvt[2]);
                lblnpvt3result.Text = Convert.ToString(npvt[3]);
                lblnpvt4result.Text = Convert.ToString(npvt[4]);
                lblnpvt5result.Text = Convert.ToString(npvt[5]);

                //محاسبه ماکزیمم ارزش خالص فعلی
                double max, min;
                max = min = npvt[1];
                if (max < npvt[1])
                    max = npvt[1];
                else if (max < npvt[2])
                    max = npvt[2];
                else if (max < npvt[3])
                    max = npvt[3];
                else if (max < npvt[4])
                    max = npvt[4];
                else if (max < npvt[5])
                    max = npvt[5];

                // چاپ نهایی نتایج محاسبات 
                if (max == npvt[1])
                    lblfinalresult.Text = "The Best Alternative is Alternative 1 With " + txtalternative1.Text + "m" + " Depth";
                else if (max == npvt[2])
                    lblfinalresult.Text = "The Best Alternative is Alternative 2 With " + txtalternative2.Text + "m" + " Depth";
                else if (max == npvt[3])
                    lblfinalresult.Text = "The Best Alternative is Alternative 3 With " + txtalternative3.Text + "m" + " Depth";
                else if (max == npvt[4])
                    lblfinalresult.Text = "The Best Alternative is Alternative 4 With " + txtalternative4.Text + "m" + " Depth";
                else if (max == npvt[5])
                    lblfinalresult.Text = "The Best Alternative is Alternative 5 With " + txtalternative5.Text + "m" + " Depth";
            }
        }

        ////////////////////////////////apply///////////////////////////////////////////
        //apply alternative
        private void btnapplyalternative_Click(object sender, EventArgs e)
        {
            if (txtalternative1.Text.Length != 0 && txtalternative2.Text.Length != 0)
                MessageBox.Show("Your Data Has Been Saved", "Done", MessageBoxButtons.OK);
            else
                MessageBox.Show("Please Check Your Data", "Error", MessageBoxButtons.OK);
        }

        //apply ore details
        private void btnapplyore_Click(object sender, EventArgs e)
        {
            if (txtreserve.Text.Length != 0 && txtoverburden1.Text.Length != 0 && txtopenpit1.Text.Length != 0 && txtopenpit2.Text.Length != 0 && txtwaste1.Text.Length != 0 && txtwaste2.Text.Length != 0)
                MessageBox.Show("Your Data Has Been Saved", "Done", MessageBoxButtons.OK);
            else
                MessageBox.Show("Please Check Your Data", "Error", MessageBoxButtons.OK);
        }

        //apply develop 
        private void btnapplydevelopment_Click(object sender, EventArgs e)
        {
            if (txtopenpitdevelopment.Text.Length != 0 && txtundergrounddevelopment1.Text.Length != 0)
                MessageBox.Show("Your Data Has Been Saved", "Done", MessageBoxButtons.OK);
            else
                MessageBox.Show("Please Check Your Data", "Error", MessageBoxButtons.OK);
        }

        //apply costs
        private void btnapplycost_Click(object sender, EventArgs e)
        {
            if (txtopenpitcost.Text.Length != 0 && txtundergroundcost.Text.Length != 0 && txtoverburdencost.Text.Length != 0 && txtprocessingcost.Text.Length != 0 && txtsmeltingcost.Text.Length != 0 && txttrasnportationcost.Text.Length != 0 && txtothercosts.Text.Length != 0)
                MessageBox.Show("Your Data Has Been Saved", "Done", MessageBoxButtons.OK);
            else
                MessageBox.Show("Please Check Your Data", "Error", MessageBoxButtons.OK);

        }

        //apply efficiency
        private void btnapplyeffi_Click(object sender, EventArgs e)
        {
            if (txtopenpitefficiency.Text.Length != 0 && txtundergroundefficiency.Text.Length != 0 && txtprocessingefficiency.Text.Length != 0 && txtsmeltingefficiency.Text.Length != 0)
                MessageBox.Show("Your Data Has Been Saved", "Done", MessageBoxButtons.OK);
            else
                MessageBox.Show("Please Check Your Data", "Error", MessageBoxButtons.OK);

        }

        //apply other
        private void btnapplyother_Click(object sender, EventArgs e)
        {
            if (txtopenpitdilution.Text.Length != 0 && txtundergrounddilution.Text.Length != 0 && txti.Text.Length != 0 && txtfactorycapacity.Text.Length != 0 && txtgrade.Text.Length != 0 && txtprice.Text.Length != 0)
                MessageBox.Show("Your Data Has Been Saved", "Done", MessageBoxButtons.OK);
            else
                MessageBox.Show("Please Check Your Data", "Error", MessageBoxButtons.OK);
        }

        //اضافه کردن گزینه 
        private void btnaddalternative_Click(object sender, EventArgs e)
        {
            if (nmbr.Value == 2)
            {
                txtalternative3.Visible = false;
                lblalternative3.Visible = false;
                txtalternative4.Visible = false;
                lblalternative4.Visible = false;
                txtalternative5.Visible = false;
                lblalternative5.Visible = false;
                txtalternative6.Visible = false;
                lblalternative6.Visible = false;
                txtalternative7.Visible = false;
                lblalternative7.Visible = false;
                txtalternative8.Visible = false;
                lblalternative8.Visible = false;
                txtalternative9.Visible = false;
                lblalternative9.Visible = false;

                lblopenpit3.Visible = false;
                txtopenpit3.Visible = false;
                lblwaste3.Visible = false;
                txtwaste3.Visible = false;
                lblopenpit4.Visible = false;
                txtopenpit4.Visible = false;
                lblwaste4.Visible = false;
                txtwaste4.Visible = false;
                lblopenpit5.Visible = false;
                txtopenpit5.Visible = false;
                lblwaste5.Visible = false;
                txtwaste5.Visible = false;
                lbloverburden3.Visible = false;
                lbloverburden4.Visible = false;
                lbloverburden5.Visible = false;
                txtoverburden3.Visible = false;
                txtoverburden4.Visible = false;
                txtoverburden5.Visible = false;


                lblalternative3result.Visible = false;
                lblnpvo3.Visible = false;
                lblnpvu3.Visible = false;
                lblnpvt3.Visible = false;
                lblnpvo3result.Visible = false;
                lblnpvu3result.Visible = false;
                lblnpvt3result.Visible = false;

                lblalternative4result.Visible = false;
                lblnpvo4.Visible = false;
                lblnpvu4.Visible = false;
                lblnpvt4.Visible = false;
                lblnpvo4result.Visible = false;
                lblnpvu4result.Visible = false;
                lblnpvt4result.Visible = false;

                lblalternative3result.Visible = false;
                lblnpvo5.Visible = false;
                lblnpvu5.Visible = false;
                lblnpvt5.Visible = false;
                lblnpvo5result.Visible = false;
                lblnpvu5result.Visible = false;
                lblnpvt5result.Visible = false;

                lblundergrounddevelopmentcost1.Visible = true;
                lblundergrounddevelopmentcost2.Visible = true;
                lblundergrounddevelopmentcost3.Visible = false;
                lblundergrounddevelopmentcost4.Visible = false;
                lblundergrounddevelopmentcost5.Visible = false;
                txtundergrounddevelopment1.Visible = true;
                txtundergrounddevelopment2.Visible = true;
                txtundergrounddevelopment3.Visible = false;
                txtundergrounddevelopment4.Visible = false;
                txtundergrounddevelopment5.Visible = false;


            }
            else if (nmbr.Value == 3)
            {
                txtalternative3.Visible = true;
                lblalternative3.Visible = true;
                txtalternative4.Visible = false;
                lblalternative4.Visible = false;
                txtalternative5.Visible = false;
                lblalternative5.Visible = false;
                txtalternative6.Visible = false;
                lblalternative6.Visible = false;
                txtalternative7.Visible = false;
                lblalternative7.Visible = false;
                txtalternative8.Visible = false;
                lblalternative8.Visible = false;
                txtalternative9.Visible = false;
                lblalternative9.Visible = false;


                lblopenpit3.Visible = true;
                lblwaste3.Visible = true;
                txtopenpit3.Visible = true;
                txtwaste3.Visible = true;

                lblopenpit4.Visible = false;
                lblwaste4.Visible = false;
                txtopenpit4.Visible = false;
                txtwaste4.Visible = false;

                lblopenpit5.Visible = false;
                lblwaste5.Visible = false;
                txtopenpit5.Visible = false;
                txtwaste5.Visible = false;

                lbloverburden3.Visible = true;
                lbloverburden4.Visible = false;
                lbloverburden5.Visible = false;
                txtoverburden3.Visible = true;
                txtoverburden4.Visible = false;
                txtoverburden5.Visible = false;

                lblalternative3result.Visible = true;
                lblnpvo3.Visible = true;
                lblnpvu3.Visible = true;
                lblnpvt3.Visible = true;
                lblnpvo3result.Visible = true;
                lblnpvu3result.Visible = true;
                lblnpvt3result.Visible = true;

                lblalternative4result.Visible = false;
                lblnpvo4.Visible = false;
                lblnpvu4.Visible = false;
                lblnpvt4.Visible = false;
                lblnpvo4result.Visible = false;
                lblnpvu4result.Visible = false;
                lblnpvt4result.Visible = false;

                lblalternative5result.Visible = false;
                lblnpvo5.Visible = false;
                lblnpvu5.Visible = false;
                lblnpvt5.Visible = false;
                lblnpvo5result.Visible = false;
                lblnpvu5result.Visible = false;
                lblnpvt5result.Visible = false;

                lblundergrounddevelopmentcost1.Visible = true;
                lblundergrounddevelopmentcost2.Visible = true;
                lblundergrounddevelopmentcost3.Visible = true;
                lblundergrounddevelopmentcost4.Visible = false;
                lblundergrounddevelopmentcost5.Visible = false;
                txtundergrounddevelopment1.Visible = true;
                txtundergrounddevelopment2.Visible = true;
                txtundergrounddevelopment3.Visible = true;
                txtundergrounddevelopment4.Visible = false;
                txtundergrounddevelopment5.Visible = false;

            }
            else if (nmbr.Value == 4)
            {

                txtalternative3.Visible = true;
                lblalternative3.Visible = true;
                txtalternative4.Visible = true;
                lblalternative4.Visible = true;
                txtalternative5.Visible = false;
                lblalternative5.Visible = false;
                txtalternative6.Visible = false;
                lblalternative6.Visible = false;
                txtalternative7.Visible = false;
                lblalternative7.Visible = false;
                txtalternative8.Visible = false;
                lblalternative8.Visible = false;
                txtalternative9.Visible = false;
                lblalternative9.Visible = false;

                lblopenpit3.Visible = true;
                lblwaste3.Visible = true;
                txtopenpit3.Visible = true;
                txtwaste3.Visible = true;

                lblopenpit4.Visible = true;
                lblwaste4.Visible = true;
                txtopenpit4.Visible = true;
                txtwaste4.Visible = true;

                lblopenpit5.Visible = false;
                lblwaste5.Visible = false;
                txtopenpit5.Visible = false;
                txtwaste5.Visible = false;

                lbloverburden3.Visible = true;
                lbloverburden4.Visible = true;
                lbloverburden5.Visible = false;
                txtoverburden3.Visible = true;
                txtoverburden4.Visible = true;
                txtoverburden5.Visible = false;


                lblalternative3result.Visible = true;
                lblnpvo3.Visible = true;
                lblnpvu3.Visible = true;
                lblnpvt3.Visible = true;
                lblnpvo3result.Visible = true;
                lblnpvu3result.Visible = true;
                lblnpvt3result.Visible = true;

                lblalternative4result.Visible = true;
                lblnpvo4.Visible = true;
                lblnpvu4.Visible = true;
                lblnpvt4.Visible = true;
                lblnpvo4result.Visible = true;
                lblnpvu4result.Visible = true;
                lblnpvt4result.Visible = true;

                lblalternative5result.Visible = false;
                lblnpvo5.Visible = false;
                lblnpvu5.Visible = false;
                lblnpvt5.Visible = false;
                lblnpvo5result.Visible = false;
                lblnpvu5result.Visible = false;
                lblnpvt5result.Visible = false;

                lblundergrounddevelopmentcost1.Visible = true;
                lblundergrounddevelopmentcost2.Visible = true;
                lblundergrounddevelopmentcost3.Visible = true;
                lblundergrounddevelopmentcost4.Visible = true;
                lblundergrounddevelopmentcost5.Visible = false;
                txtundergrounddevelopment1.Visible = true;
                txtundergrounddevelopment2.Visible = true;
                txtundergrounddevelopment3.Visible = true;
                txtundergrounddevelopment4.Visible = true;
                txtundergrounddevelopment5.Visible = false;
            }
            else if (nmbr.Value == 5)
            {
                txtalternative3.Visible = true;
                lblalternative3.Visible = true;
                txtalternative4.Visible = true;
                lblalternative4.Visible = true;
                txtalternative5.Visible = true;
                lblalternative5.Visible = true;
                txtalternative6.Visible = false;
                lblalternative6.Visible = false;
                txtalternative7.Visible = false;
                lblalternative7.Visible = false;
                txtalternative8.Visible = false;
                lblalternative8.Visible = false;
                txtalternative9.Visible = false;
                lblalternative9.Visible = false;

                lblopenpit3.Visible = true;
                lblwaste3.Visible = true;
                txtopenpit3.Visible = true;
                txtwaste3.Visible = true;

                lblopenpit4.Visible = true;
                lblwaste4.Visible = true;
                txtopenpit4.Visible = true;
                txtwaste4.Visible = true;

                lblopenpit5.Visible = true;
                lblwaste5.Visible = true;
                txtopenpit5.Visible = true;
                txtwaste5.Visible = true;


                lblalternative3result.Visible = true;
                lblnpvo3.Visible = true;
                lblnpvu3.Visible = true;
                lblnpvt3.Visible = true;
                lblnpvo3result.Visible = true;
                lblnpvu3result.Visible = true;
                lblnpvt3result.Visible = true;

                lblalternative4result.Visible = true;
                lblnpvo4.Visible = true;
                lblnpvu4.Visible = true;
                lblnpvt4.Visible = true;
                lblnpvo4result.Visible = true;
                lblnpvu4result.Visible = true;
                lblnpvt4result.Visible = true;

                lblalternative5result.Visible = true;
                lblnpvo5.Visible = true;
                lblnpvu5.Visible = true;
                lblnpvt5.Visible = true;
                lblnpvo5result.Visible = true;
                lblnpvu5result.Visible = true;
                lblnpvt5result.Visible = true;

                lbloverburden3.Visible = true;
                lbloverburden4.Visible = true;
                lbloverburden5.Visible = true;
                txtoverburden3.Visible = true;
                txtoverburden4.Visible = true;
                txtoverburden5.Visible = true;

                lblundergrounddevelopmentcost1.Visible = true;
                lblundergrounddevelopmentcost2.Visible = true;
                lblundergrounddevelopmentcost3.Visible = true;
                lblundergrounddevelopmentcost4.Visible = true;
                lblundergrounddevelopmentcost5.Visible = true;
                txtundergrounddevelopment1.Visible = true;
                txtundergrounddevelopment2.Visible = true;
                txtundergrounddevelopment3.Visible = true;
                txtundergrounddevelopment4.Visible = true;
                txtundergrounddevelopment5.Visible = true;

            }
            else if (nmbr.Value == 6)
            {
                txtalternative3.Visible = true;
                lblalternative3.Visible = true;
                txtalternative4.Visible = true;
                lblalternative4.Visible = true;
                txtalternative5.Visible = true;
                lblalternative5.Visible = true;
                txtalternative6.Visible = true;
                lblalternative6.Visible = true;
                txtalternative7.Visible = false;
                lblalternative7.Visible = false;
                txtalternative8.Visible = false;
                lblalternative8.Visible = false;
                txtalternative9.Visible = false;
                lblalternative9.Visible = false;
            }
            else if (nmbr.Value == 7)
            {
                txtalternative3.Visible = true;
                lblalternative3.Visible = true;
                txtalternative4.Visible = true;
                lblalternative4.Visible = true;
                txtalternative5.Visible = true;
                lblalternative5.Visible = true;
                txtalternative6.Visible = true;
                lblalternative6.Visible = true;
                txtalternative7.Visible = true;
                lblalternative7.Visible = true;
                txtalternative8.Visible = false;
                lblalternative8.Visible = false;
                txtalternative9.Visible = false;
                lblalternative9.Visible = false;
            }
            else if (nmbr.Value == 8)
            {
                txtalternative3.Visible = true;
                lblalternative3.Visible = true;
                txtalternative4.Visible = true;
                lblalternative4.Visible = true;
                txtalternative5.Visible = true;
                lblalternative5.Visible = true;
                txtalternative6.Visible = true;
                lblalternative6.Visible = true;
                txtalternative7.Visible = true;
                lblalternative7.Visible = true;
                txtalternative8.Visible = true;
                lblalternative8.Visible = true;
                txtalternative9.Visible = false;
                lblalternative9.Visible = false;
            }
            else if (nmbr.Value == 9)
            {
                txtalternative3.Visible = true;
                lblalternative3.Visible = true;
                txtalternative4.Visible = true;
                lblalternative4.Visible = true;
                txtalternative5.Visible = true;
                lblalternative5.Visible = true;
                txtalternative6.Visible = true;
                lblalternative6.Visible = true;
                txtalternative7.Visible = true;
                lblalternative7.Visible = true;
                txtalternative8.Visible = true;
                lblalternative8.Visible = true;
                txtalternative9.Visible = true;
                lblalternative9.Visible = true;
            }
        }

        //تب ها 
        private void tabalternatives_Click(object sender, EventArgs e)
        {
            txtalternative1.Focus();
        }

        private void tabore_Click(object sender, EventArgs e)
        {
            txtreserve.Focus();
        }

        private void tabcosts_Click(object sender, EventArgs e)
        {
            txtopenpitcost.Focus();
        }

        private void tabefficiency_Click(object sender, EventArgs e)
        {
            txtopenpitefficiency.Focus();
        }

        private void tabother_Click(object sender, EventArgs e)
        {
            txti.Focus();
        }


        //کد های اضافه 
        private void txtalternative1_TextChanged(object sender, EventArgs e)
        {
            if (txtalternative1.Text.Length == 0 || txtalternative2.Text.Length == 0)
            {
                btnapplyalternative.Enabled = false;
            }

            else
            {
                btnapplyalternative.Enabled = true;
            }

            if (txtalternative1.Text.Length != 0 || txtalternative2.Text.Length != 0)
            {
                btnclearalternative.Enabled = true;
            }
            else
                btnapplyalternative.Enabled = false;

        }

        private void txtalternative2_TextChanged(object sender, EventArgs e)
        {
            txtalternative1_TextChanged(null, null);
        }

        private void txtreserve_TextChanged(object sender, EventArgs e)
        {
            if (txtreserve.Text.Length == 0 || txtoverburden1.Text.Length == 0 || txtoverburden2.Text.Length == 0 || txtopenpit1.Text.Length == 0 || txtopenpit2.Text.Length == 0 || txtwaste1.Text.Length == 0 || txtwaste2.Text.Length == 0)
            {
                btnapplyalternative.Enabled = false;
            }
            else
            {
                btnapplyore.Enabled = true;
            }

            if (txtreserve.Text.Length != 0 || txtoverburden1.Text.Length != 0 || txtoverburden2.Text.Length != 0 || txtopenpit1.Text.Length != 0 || txtopenpit2.Text.Length != 0 || txtwaste1.Text.Length != 0 || txtwaste2.Text.Length != 0)
            {
                btnclearore.Enabled = true;
            }
            else
                btnclearore.Enabled = false;



        }

        private void txtoverburden2_TextChanged(object sender, EventArgs e)
        {
            txtreserve_TextChanged(null, null);

        }

        private void txtoverburden_TextChanged(object sender, EventArgs e)
        {
           txtreserve_TextChanged(null, null);

        }

        private void txtopenpitdevelopment_TextChanged(object sender, EventArgs e)
        {
            if (txtopenpitdevelopment.Text.Length == 0 || txtundergrounddevelopment1.Text.Length == 0 || txtundergrounddevelopment2.Text.Length == 0)
            {
                btnapplydevelopment.Enabled = false;
            }
            else
                btnapplydevelopment.Enabled = true;

            if (txtopenpitdevelopment.Text.Length == 0 || txtundergrounddevelopment1.Text.Length == 0 || txtundergrounddevelopment2.Text.Length == 0)
            {
                btncleardevelopment.Enabled = true;
            }
            else
                btncleardevelopment.Enabled = false;
        }

        private void txtundergrounddevelopment1_TextChanged(object sender, EventArgs e)
        {
            txtopenpitdevelopment_TextChanged(null, null);
        }

        private void txtundergrounddevelopment2_TextChanged(object sender, EventArgs e)
        {
            txtopenpitdevelopment_TextChanged(null, null);

        }

        private void txtopenpit1_TextChanged(object sender, EventArgs e)
        {
            txtreserve_TextChanged(null, null);
        }

        private void txtopenpit2_TextChanged(object sender, EventArgs e)
        {
            txtreserve_TextChanged(null, null);

        }

        private void txtwaste1_TextChanged(object sender, EventArgs e)
        {
            txtreserve_TextChanged(null, null);

        }

        private void txtwaste2_TextChanged(object sender, EventArgs e)
        {
            txtreserve_TextChanged(null, null);

        }

        private void txtopenpitcost_TextChanged(object sender, EventArgs e)
        {
            if (txtopenpitcost.Text.Length == 0 || txtundergroundcost.Text.Length == 0 || txtstrippingcost.Text.Length == 0 || txtoverburdencost.Text.Length == 0 || txtprocessingcost.Text.Length == 0 || txtsmeltingcost.Text.Length == 0 || txttrasnportationcost.Text.Length == 0 || txtothercosts.Text.Length == 0 )
                btnapplycost.Enabled = false;
            else
                btnapplycost.Enabled = true;

            if (txtopenpitcost.Text.Length == 0 || txtundergroundcost.Text.Length == 0 || txtstrippingcost.Text.Length == 0 || txtoverburdencost.Text.Length == 0 || txtprocessingcost.Text.Length == 0 || txtsmeltingcost.Text.Length == 0 || txttrasnportationcost.Text.Length == 0 || txtothercosts.Text.Length == 0)
            {
                btnclearcost.Enabled = true;

            }
            else
                btnclearcost.Enabled = true;
        }

        private void txtundergroundcost_TextChanged(object sender, EventArgs e)
        {
            txtopenpitcost_TextChanged(null, null);
        }

        private void txtstrippingcost_TextChanged(object sender, EventArgs e)
        {
            txtopenpitcost_TextChanged(null, null);
        }

        private void txtoverburdencost_TextChanged(object sender, EventArgs e)
        {
            txtopenpitcost_TextChanged(null, null);
        }

        private void txtprocessingcost_TextChanged(object sender, EventArgs e)
        {
            txtopenpitcost_TextChanged(null, null);
        }

        private void txtsmeltingcost_TextChanged(object sender, EventArgs e)
        {
            txtopenpitcost_TextChanged(null, null);
        }

        private void txttrasnportationcost_TextChanged(object sender, EventArgs e)
        {
            txtopenpitcost_TextChanged(null, null);
        }

        private void txtothercosts_TextChanged(object sender, EventArgs e)
        {
            txtopenpitcost_TextChanged(null, null);
        }

        private void nmbroverburden_ValueChanged(object sender, EventArgs e)
        {
        }

        private void txtopenpitefficiency_TextChanged(object sender, EventArgs e)
        {
            if (txtopenpitefficiency.Text.Length == 0 || txtundergroundefficiency.Text.Length == 0 || txtprocessingefficiency.Text.Length == 0 || txtsmeltingefficiency.Text.Length == 0 )
            {
                btnapplyeffi.Enabled = false;
            }
            else
                btnapplyeffi.Enabled = true;

            if (txtopenpitefficiency.Text.Length == 0 || txtundergroundefficiency.Text.Length == 0 || txtprocessingefficiency.Text.Length == 0 || txtsmeltingefficiency.Text.Length == 0)
            {
                btncleareffi.Enabled = true;
            }
            else
                btncleareffi.Enabled = false;
        }

        private void txtundergroundefficiency_TextChanged(object sender, EventArgs e)
        {
            txtopenpitefficiency_TextChanged(null, null);
        }

        private void txtprocessingefficiency_TextChanged(object sender, EventArgs e)
        {
            txtopenpitefficiency_TextChanged(null, null);

        }

        private void txtsmeltingefficiency_TextChanged(object sender, EventArgs e)
        {
            txtopenpitefficiency_TextChanged(null, null);

        }

        private void txti_TextChanged(object sender, EventArgs e)
        {
            if (txti.Text.Length == 0 || txtopenpitdilution.Text.Length == 0 || txtundergrounddilution.Text.Length == 0 || txtfactorycapacity.Text.Length == 0 || txtprice.Text.Length == 0 || txttax.Text.Length == 0)
            {
                btnapplyother.Enabled = false;
            }
            else
                btnapplyother.Enabled = true;

            if (txti.Text.Length == 0 || txtopenpitdilution.Text.Length == 0 || txtundergrounddilution.Text.Length == 0 || txtfactorycapacity.Text.Length == 0 || txtprice.Text.Length == 0 || txttax.Text.Length == 0)
            {
                btnapplyother.Enabled = true;
            }
            else
                btnapplyother.Enabled = false;
        }

        private void txtopenpitdilution_TextChanged(object sender, EventArgs e)
        {
            txti_TextChanged(null, null);
        }

        private void txtundergrounddilution_TextChanged(object sender, EventArgs e)
        {
            txti_TextChanged(null, null);
        }

        private void txtfactorycapacity_TextChanged(object sender, EventArgs e)
        {
            txti_TextChanged(null, null);
        }

        private void txtgrade_TextChanged(object sender, EventArgs e)
        {
            txti_TextChanged(null, null);
        }

        private void txtprice_TextChanged(object sender, EventArgs e)
        {
            txti_TextChanged(null, null);
        }
        private void txttax_TextChanged(object sender, EventArgs e)
        {
            txti_TextChanged(null, null);
        }
        private void nmbr_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {
        }

        private void btncleardevelopment_Click_1(object sender, EventArgs e)
        {

        }

        private void chkbxconstantcapacity_CheckedChanged(object sender, EventArgs e)
        {
            if (chkbxconstantcapacity.Checked == false)
            {
                lblfactorycapacity.Text = "Annual Capacity (MT/Year)";
            }
            else if (chkbxconstantcapacity.Checked == true)
            {
                lblfactorycapacity.Text = "Annual Metal (MT/Year)";
            }
        }

        private void btnprint_Click(object sender, EventArgs e)
        {
            this.tabcontrol.SelectedIndex = 7;

            this.dataGridView1.Rows.Add(lblalternative1.Text, txtalternative1.Text);
            this.dataGridView1.Rows.Add(lblalternative2.Text, txtalternative2.Text);
            this.dataGridView1.Rows.Add(lblalternative3.Text, txtalternative3.Text);
            this.dataGridView1.Rows.Add(lblalternative4.Text, txtalternative4.Text);
            this.dataGridView1.Rows.Add(lblalternative5.Text, txtalternative5.Text);

            this.dataGridView1.Rows.Add(lblreserve.Text, txtreserve.Text);

            this.dataGridView1.Rows.Add(lbloverburden1.Text, txtoverburden1.Text);
            this.dataGridView1.Rows.Add(lblopenpit1.Text, txtopenpit1.Text);
            this.dataGridView1.Rows.Add(lblwaste1.Text, txtwaste1.Text);


            this.dataGridView1.Rows.Add(lbloverburden2.Text, txtoverburden2.Text);
            this.dataGridView1.Rows.Add(lblopenpit2.Text, txtopenpit2.Text);
            this.dataGridView1.Rows.Add(lblwaste2.Text, txtwaste2.Text);


            this.dataGridView1.Rows.Add(lbloverburden3.Text, txtoverburden3.Text);
            this.dataGridView1.Rows.Add(lblopenpit3.Text, txtopenpit3.Text);
            this.dataGridView1.Rows.Add(lblwaste3.Text, txtwaste3.Text);


            this.dataGridView1.Rows.Add(lbloverburden4.Text, txtoverburden4.Text);
            this.dataGridView1.Rows.Add(lblopenpit4.Text, txtopenpit4.Text);
            this.dataGridView1.Rows.Add(lblwaste4.Text, txtwaste4.Text);


            this.dataGridView1.Rows.Add(lbloverburden5.Text, txtoverburden5.Text);
            this.dataGridView1.Rows.Add(lblopenpit5.Text, txtopenpit5.Text);
            this.dataGridView1.Rows.Add(lblwaste5.Text, txtwaste5.Text);


            this.dataGridView1.Rows.Add(lbloverburdenyear.Text, nmbroverburden.Value);
            this.dataGridView1.Rows.Add(lblopenpityear.Text, nmbropenpit.Value);
            this.dataGridView1.Rows.Add(lblundergroundyear.Text, nmbrunderground.Value);
            this.dataGridView1.Rows.Add(lblopenpitdevelopmentcost.Text, txtopenpitdevelopment.Text);
            this.dataGridView1.Rows.Add(lblundergrounddevelopmentcost1.Text, txtundergrounddevelopment1.Text);
            this.dataGridView1.Rows.Add(lblundergrounddevelopmentcost2.Text, txtundergrounddevelopment2.Text);
            this.dataGridView1.Rows.Add(lblundergrounddevelopmentcost3.Text, txtundergrounddevelopment3.Text);
            this.dataGridView1.Rows.Add(lblundergrounddevelopmentcost4.Text, txtundergrounddevelopment4.Text);
            this.dataGridView1.Rows.Add(lblundergrounddevelopmentcost5.Text, txtundergrounddevelopment5.Text);


            this.dataGridView1.Rows.Add(lblopenpitcost.Text, txtopenpitcost.Text);
            this.dataGridView1.Rows.Add(lblundergroundcost.Text, txtundergroundcost.Text);
            this.dataGridView1.Rows.Add(lblstrippingcost.Text, txtstrippingcost.Text);
            this.dataGridView1.Rows.Add(lbloverburdencost.Text, txtoverburdencost.Text);
            this.dataGridView1.Rows.Add(lblprocessingcost.Text, txtprocessingcost.Text);
            this.dataGridView1.Rows.Add(lblsmeltingcost.Text, txtsmeltingcost.Text);
            this.dataGridView1.Rows.Add(lbltrasportationcost.Text, txttrasnportationcost.Text);
            this.dataGridView1.Rows.Add(lblothercosts.Text, txtothercosts.Text);


            this.dataGridView1.Rows.Add(lblopenpitefficiency.Text, txtopenpitefficiency.Text);
            this.dataGridView1.Rows.Add(lblundergroundefficiency.Text, txtundergroundefficiency.Text);
            this.dataGridView1.Rows.Add(lblprocessingefficiency.Text, txtprocessingefficiency.Text);
            this.dataGridView1.Rows.Add(lblsmeltingefficiency.Text, txtsmeltingefficiency.Text);


            this.dataGridView1.Rows.Add(lbli.Text, txti.Text);
            this.dataGridView1.Rows.Add(lblopenpitdilution.Text, txtopenpitdilution.Text);
            this.dataGridView1.Rows.Add(lblundergrounddilution.Text, txtundergrounddilution.Text);
            this.dataGridView1.Rows.Add(lblfactorycapacity.Text, txtfactorycapacity.Text);
            this.dataGridView1.Rows.Add(lblgrade.Text, txtgrade.Text);
            this.dataGridView1.Rows.Add(lblprice.Text, txtprice.Text);
                                                    

            this.dataGridView2.Rows.Add(lblnpvo1.Text, lblnpvo1result.Text);
            this.dataGridView2.Rows.Add(lblnpvu1.Text, lblnpvu1result.Text);
            this.dataGridView2.Rows.Add(lblnpvt1.Text, lblnpvt1result.Text);


            this.dataGridView2.Rows.Add(lblnpvo2.Text, lblnpvo2result.Text);
            this.dataGridView2.Rows.Add(lblnpvu2.Text, lblnpvu2result.Text);
            this.dataGridView2.Rows.Add(lblnpvt2.Text, lblnpvt2result.Text);


            this.dataGridView2.Rows.Add(lblnpvo3.Text, lblnpvo3result.Text);
            this.dataGridView2.Rows.Add(lblnpvu3.Text, lblnpvu3result.Text);
            this.dataGridView2.Rows.Add(lblnpvt3.Text, lblnpvt3result.Text);


            this.dataGridView2.Rows.Add(lblnpvo4.Text, lblnpvo4result.Text);
            this.dataGridView2.Rows.Add(lblnpvu4.Text, lblnpvu4result.Text);
            this.dataGridView2.Rows.Add(lblnpvt4.Text, lblnpvt4result.Text);


            this.dataGridView2.Rows.Add(lblnpvo5.Text, lblnpvo5result.Text);
            this.dataGridView2.Rows.Add(lblnpvu5.Text, lblnpvu5result.Text);
            this.dataGridView2.Rows.Add(lblnpvt5.Text, lblnpvt5result.Text);


            btnprint.Enabled = false;
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Bitmap bm = new Bitmap(this.dataGridView1.Width, this.dataGridView1.Height);

            dataGridView1.DrawToBitmap(bm, new Rectangle(0, 0, this.dataGridView1.Width, this.dataGridView1.Height));
            e.Graphics.DrawImage(bm, 0, 0);
        }

        private void btnfinalprint_Click(object sender, EventArgs e)
        {
            printDocument1.Print();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("The Excel File Has Been Created", "Done", MessageBoxButtons.OK);

            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            app.Visible = true;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Result";

            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {

                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;

            }

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }

            ///////////////////////////////////////////////////////////////////

            for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
            {

                worksheet.Cells[1, i + 7] = dataGridView2.Columns[i - 1].HeaderText;

            }

            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView2.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1 + 7] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                }
            }

            btnprint.Enabled = true;
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();

        }
    }
}



