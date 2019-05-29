using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using MaterialSkin.Controls;
using MaterialSkin;
using System.Data.OleDb;
using System.IO;
using System.Net.Mail;
using System.Net;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace material_design
{
    public partial class Form1 : MaterialForm
    {       
        public Form1()
        {
            InitializeComponent();
        }
        private const int KaynakKlasorSatirNo = 0;
        private const int KopyaKlasorSatirNo = 1;
        private string[] _ayarlarDosyasiIcerigi;
        int saat = 00, dakika = 00, saniye = 00;
        OleDbConnection baglantı = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\star_\\OneDrive\\Masaüstü\\punchdkn.mdb");
         public static void pdfKaydet(DataGridView veriTablosu, string filename)
            {
            string Tahoma = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "Tahoma.TTF");
            BaseFont bf = BaseFont.CreateFont(Tahoma, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            //  BaseFont bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1250, BaseFont.NOT_EMBEDDED);
                PdfPTable pdftable = new PdfPTable(veriTablosu.Columns.Count);
                pdftable.DefaultCell.Padding = 2;
                pdftable.WidthPercentage = 100;
                pdftable.HorizontalAlignment = Element.ALIGN_CENTER;
                pdftable.DefaultCell.BorderWidth = 1;
                iTextSharp.text.Font text = new iTextSharp.text.Font(bf, 10, iTextSharp.text.Font.NORMAL);
                foreach (DataGridViewColumn column in veriTablosu.Columns)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText, text));
                    cell.BackgroundColor = new iTextSharp.text.BaseColor(128, 255, 255);
                    pdftable.AddCell(cell);
                }
                foreach (DataGridViewRow row in veriTablosu.Rows)
                {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    try
                    {

                        pdftable.AddCell(new Phrase(cell.Value.ToString(), text));
                    }
                    catch { }
                }
                }
                var savefiledialog = new SaveFileDialog();
                savefiledialog.FileName = filename;
                savefiledialog.DefaultExt = ".pdf";
                if (savefiledialog.ShowDialog() == DialogResult.OK)
                {
                    using (FileStream stream = new FileStream(savefiledialog.FileName, FileMode.Create))
                    {
                        Document pdfdoc = new Document(PageSize.A4, 10f, 10f, 10f, 10f);
                        PdfWriter.GetInstance(pdfdoc, stream);
                        pdfdoc.Open();
                        pdfdoc.Add(pdftable);
                        pdfdoc.Close();
                        stream.Close();
                    }
                }
            }        
        private void verilerigörüntüle()
        {
            ListViewKod.Items.Clear();
            baglantı.Open();
            OleDbCommand komut = new OleDbCommand();
            komut.Connection = baglantı;
            komut.CommandText = ("Select * From punch");
            OleDbDataReader oku = komut.ExecuteReader();
            while (oku.Read())
            {
                ListViewItem ekle = new ListViewItem();
                ekle.Text = oku["Panel_kod"].ToString();
                ekle.SubItems.Add(oku["Tarih"].ToString());
                ekle.SubItems.Add(oku["Panel_Izo"].ToString());
                ekle.SubItems.Add(oku["Sac_Tipi"].ToString());
                ekle.SubItems.Add(oku["Panel_Tipi"].ToString());
                ekle.SubItems.Add(oku["Dıs_Sac_Olcu"].ToString());
                ekle.SubItems.Add(oku["Ic_Sac_Olcu"].ToString());
                ekle.SubItems.Add(oku["Operatör_ismi"].ToString());
                ekle.SubItems.Add(oku["Vardiya_saati"].ToString());
                ekle.SubItems.Add(oku["islendigi_saat"].ToString());
                ekle.SubItems.Add(oku["Makina"].ToString());
                ListViewKod.Items.Add(ekle);
            }
            baglantı.Close();
        }
        private void AyarlariOku()
        {
            var ayarlarDosyasiKayitliMi = File.Exists("Ayarlar.txt");
            if (!ayarlarDosyasiKayitliMi)
            {
                _ayarlarDosyasiIcerigi = new string[3];
                _ayarlarDosyasiIcerigi[KaynakKlasorSatirNo] = "(KAYITLI DEĞİL)";
                _ayarlarDosyasiIcerigi[KopyaKlasorSatirNo] = "(KAYITLI DEĞİL)";
                kaynakKlasorTextBox.Text = "(KAYITLI DEĞİL)";
                hedefKlasorTextBox.Text = "(KAYITLI DEĞİL)";
                kopyaKlasorTextBox.Text = "(KAYITLI DEĞİL)";
                return;
            }
            _ayarlarDosyasiIcerigi = File.ReadAllLines("Ayarlar.txt");
            var kaynakKlasor = _ayarlarDosyasiIcerigi[KaynakKlasorSatirNo];
            var kopyaKlasor = _ayarlarDosyasiIcerigi[KopyaKlasorSatirNo];
            kaynakKlasorTextBox.Text = kaynakKlasor;
            kopyaKlasorTextBox.Text = kopyaKlasor;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: Bu kod satırı 'punchdknDataSet7.saat' tablosuna veri yükler. Bunu gerektiği şekilde taşıyabilir, veya kaldırabilirsiniz.
            this.saatTableAdapter3.Fill(this.punchdknDataSet7.saat);
            // TODO: Bu kod satırı 'punchdknDataSet6.saat' tablosuna veri yükler. Bunu gerektiği şekilde taşıyabilir, veya kaldırabilirsiniz.
            this.saatTableAdapter2.Fill(this.punchdknDataSet6.saat);
            // TODO: Bu kod satırı 'punchdknDataSet5.saat' tablosuna veri yükler. Bunu gerektiği şekilde taşıyabilir, veya kaldırabilirsiniz.
            this.saatTableAdapter1.Fill(this.punchdknDataSet5.saat);
            // TODO: Bu kod satırı 'punchdknDataSet4.saat' tablosuna veri yükler. Bunu gerektiği şekilde taşıyabilir, veya kaldırabilirsiniz.
            this.saatTableAdapter.Fill(this.punchdknDataSet4.saat);
            // TODO: Bu kod satırı 'punchdknDataSet3.Vardiya' tablosuna veri yükler. Bunu gerektiği şekilde taşıyabilir, veya kaldırabilirsiniz.
            this.vardiyaTableAdapter.Fill(this.punchdknDataSet3.Vardiya);
            // TODO: Bu kod satırı 'punchdknDataSet2.Makina' tablosuna veri yükler. Bunu gerektiği şekilde taşıyabilir, veya kaldırabilirsiniz.
            this.makinaTableAdapter.Fill(this.punchdknDataSet2.Makina);
            // TODO: Bu kod satırı 'punchdknDataSet1.personel' tablosuna veri yükler. Bunu gerektiği şekilde taşıyabilir, veya kaldırabilirsiniz.
            this.personelTableAdapter.Fill(this.punchdknDataSet1.personel);
            // TODO: Bu kod satırı 'punchdknDataSet.durus_kodu' tablosuna veri yükler. Bunu gerektiği şekilde taşıyabilir, veya kaldırabilirsiniz.
            this.durus_koduTableAdapter.Fill(this.punchdknDataSet.durus_kodu);
            MaterialSkinManager msm = MaterialSkinManager.Instance;
            msm.AddFormToManage(this);
            msm.Theme = MaterialSkinManager.Themes.DARK;
            AyarlariOku();
            barkodTxt.Focus();
            this.ActiveControl = barkodTxt;
            this.Location = new System.Drawing.Point(0, 0);
            this.Size = Screen.PrimaryScreen.WorkingArea.Size;
        }
        private void izoOly(object sender, EventArgs e)
        {
            if (IzoTxt.Text == "40")
                IzoTxt.Clear();
            else if (IzoTxt.Text == "50")
                IzoTxt.Clear();
            else if (IzoTxt.Text == "60")
                IzoTxt.Clear();
            System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
            IzoTxt.Text += btn.Text;
            if (IzoTxt.Text == "40")
                IzoLbl.BackColor = Color.PaleGreen;
            else if (IzoTxt.Text == "50")
                IzoLbl.BackColor = Color.Aquamarine;
            else if (IzoTxt.Text == "60")
                IzoLbl.BackColor = Color.YellowGreen;
        }
        private void IzoTxt_TextChanged(object sender, EventArgs e)
        {
            PnlKodTxt.Text = IzoTxt.Text + "_" + SacTxt.Text + "_" + DısOlcu.Text + "x" + IcOlcu.Text + "_" + PnlTxt.Text;
            if (IzoTxt.Text == "40")
                IzoTxt.BackColor = Color.PaleGreen;
            else if (IzoTxt.Text == "50")
                IzoTxt.BackColor = Color.Aquamarine;
            else if (IzoTxt.Text == "60")
                IzoTxt.BackColor = Color.YellowGreen;
        }
        private void SacOly(object sender, EventArgs e)
        {
            if (SacTxt.Text == "D")
                SacTxt.Clear();
            else if (SacTxt.Text == "I")
                SacTxt.Clear();
            System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
            SacTxt.Text += btn.Text;
            if (SacTxt.Text == "D")
                SacLbl.BackColor = Color.Khaki;
            else if (SacTxt.Text == "I")
                SacLbl.BackColor = Color.Pink;
        }
        private void SacTxt_TextChanged(object sender, EventArgs e)
        {
            PnlKodTxt.Text = IzoTxt.Text + "_" + SacTxt.Text + "_" + DısOlcu.Text + "x" + IcOlcu.Text + "_" + PnlTxt.Text;
            if (SacTxt.Text == "D")
                SacTxt.BackColor = Color.Khaki;
            else if (SacTxt.Text == "I")
                SacTxt.BackColor = Color.Pink;
        }
        private void PnlTipOly(object sender, EventArgs e)
        {
            if (PnlTxt.Text == "NP")
                PnlTxt.Clear();
            else if (PnlTxt.Text == "MDP")
                PnlTxt.Clear();
            else if (PnlTxt.Text == "GP")
                PnlTxt.Clear();
            else if (PnlTxt.Text == "NK")
                PnlTxt.Clear();
            else if (PnlTxt.Text == "GK")
                PnlTxt.Clear();
            else if (PnlTxt.Text == "KNK")
                PnlTxt.Clear();
            else if (PnlTxt.Text == "KGK")
                PnlTxt.Clear();
            else if (PnlTxt.Text == "NICK")
                PnlTxt.Clear();
            else if (PnlTxt.Text == "GICK")
                PnlTxt.Clear();
            else if (PnlTxt.Text == "GCKK")
                PnlTxt.Clear();
            else if (PnlTxt.Text == "PTP")
                PnlTxt.Clear();
            else if (PnlTxt.Text == "PK")
                PnlTxt.Clear();
            System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
            PnlTxt.Text += btn.Text;
            if (PnlTxt.Text == "NP")
                PanelLbl.BackColor = Color.DeepSkyBlue;
            else if (PnlTxt.Text == "MDP")
                PanelLbl.BackColor = Color.Bisque;
            else if (PnlTxt.Text == "GP")
                PanelLbl.BackColor = Color.MediumAquamarine;
            else if (PnlTxt.Text == "NK")
                PanelLbl.BackColor = Color.BurlyWood;
            else if (PnlTxt.Text == "GK")
                PanelLbl.BackColor = Color.SteelBlue;
            else if (PnlTxt.Text == "KNK")
                PanelLbl.BackColor = Color.LightCoral;
            else if (PnlTxt.Text == "KGK")
                PanelLbl.BackColor = Color.DarkSeaGreen;
            else if (PnlTxt.Text == "NICK")
                PanelLbl.BackColor = Color.SpringGreen;
            else if (PnlTxt.Text == "GICK")
                PanelLbl.BackColor = Color.LightSteelBlue;
            else if (PnlTxt.Text == "GCKK")
                PanelLbl.BackColor = Color.MediumSpringGreen;
            else if (PnlTxt.Text == "PTP")
                PanelLbl.BackColor = Color.PowderBlue;
        }
        private void PnlTxt_TextChanged(object sender, EventArgs e)
        {
            PnlKodTxt.Text = IzoTxt.Text + "_" + SacTxt.Text + "_" + DısOlcu.Text + "x" + IcOlcu.Text + "_" + PnlTxt.Text;
            if (PnlTxt.Text == "NP")
                PnlTxt.BackColor = Color.DeepSkyBlue;
            else if (PnlTxt.Text == "MDP")
                PnlTxt.BackColor = Color.Bisque;
            else if (PnlTxt.Text == "GP")
                PnlTxt.BackColor = Color.MediumAquamarine;
            else if (PnlTxt.Text == "NK")
                PnlTxt.BackColor = Color.BurlyWood;
            else if (PnlTxt.Text == "GK")
                PnlTxt.BackColor = Color.SteelBlue;
            else if (PnlTxt.Text == "KNK")
                PnlTxt.BackColor = Color.LightCoral;
            else if (PnlTxt.Text == "KGK")
                PnlTxt.BackColor = Color.DarkSeaGreen;
            else if (PnlTxt.Text == "NICK")
                PnlTxt.BackColor = Color.SpringGreen;
            else if (PnlTxt.Text == "GICK")
                PnlTxt.BackColor = Color.LightSteelBlue;
            else if (PnlTxt.Text == "GCKK")
                PnlTxt.BackColor = Color.MediumSpringGreen;
            else if (PnlTxt.Text == "PTP")
                PnlTxt.BackColor = Color.PowderBlue;
            else if (PnlTxt.Text == "PK")
                PnlTxt.BackColor = Color.YellowGreen;
        }
        private void DısOlcu_TextChanged(object sender, EventArgs e)
        {
            PnlKodTxt.Text = IzoTxt.Text + "_" + SacTxt.Text + "_" + DısOlcu.Text + "x" + IcOlcu.Text + "_" + PnlTxt.Text;
        }
        private void IcOlcu_TextChanged(object sender, EventArgs e)
        {
            PnlKodTxt.Text = IzoTxt.Text + "_" + SacTxt.Text + "_" + DısOlcu.Text + "x" + IcOlcu.Text + "_" + PnlTxt.Text;
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            ZmnLbl.Text = DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss");
        }
        private void pictureBox8_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            barkodTxt.Clear();
            barkodTxt.Focus();
        }
        private void barkodTxt_TextChanged(object sender, EventArgs e)
        {
            string barkod = Clipboard.GetText();
        }
        private void KynkKlsrBtn_Click(object sender, EventArgs e)
        {
            if (pswrdTxt.Text == "123456")
            {
                using (var kaynakKlasorDialog = new FolderBrowserDialog())
                {
                    var dialogResult = kaynakKlasorDialog.ShowDialog();
                    if (dialogResult != DialogResult.OK) return;

                    _ayarlarDosyasiIcerigi[KaynakKlasorSatirNo] = kaynakKlasorDialog.SelectedPath;
                    kaynakKlasorTextBox.Text = kaynakKlasorDialog.SelectedPath;
                }
            }
            else
            {
                MessageBox.Show("Lütfen " +
                    "ilk sayfadaki Şifre kısmını doğru giriniz!!!", "UYARI!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                pswrdTxt.Focus();
            }

            AyarlariKaydet();
        }
        private void AyarlariKaydet()
        {
            File.WriteAllLines("Ayarlar.txt", _ayarlarDosyasiIcerigi);
        }
        private void KpyaKlsrBtn_Click(object sender, EventArgs e)
        {
            if (pswrdTxt.Text == "123456")
            {
                using (var kopyaKlasorDialog = new FolderBrowserDialog())
                {
                    var dialogResult = kopyaKlasorDialog.ShowDialog();
                    if (dialogResult != DialogResult.OK) return;

                    _ayarlarDosyasiIcerigi[KopyaKlasorSatirNo] = kopyaKlasorDialog.SelectedPath;
                    kopyaKlasorTextBox.Text = kopyaKlasorDialog.SelectedPath;
                }
            }
            else
            {
                MessageBox.Show("Lütfen ilk sayfadaki Şifre kısmını doğru giriniz!!!", "UYARI!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                pswrdTxt.Focus();
            }

            AyarlariKaydet();
        }
        private void VeriGstrBtn_Click(object sender, EventArgs e)
        {
            if (pswrdTxt.Text == "123456")
            {
                verilerigörüntüle();
                int kayitsayisi = ListViewKod.Items.Count;
                sayiTxt.Text = Convert.ToString(kayitsayisi);
            }
            else
            {
                MessageBox.Show("Lütfen yan taraftaki Şifre kısmını doğru giriniz!!!", "UYARI!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                pswrdTxt.Focus();
            }
        }
        private void MailBtn_Click(object sender, EventArgs e)
        {
            MailMessage mesajım = new MailMessage();
            SmtpClient istemci = new SmtpClient();
            istemci.UseDefaultCredentials = false;
            istemci.Credentials = new System.Net.NetworkCredential("daikin_panel@hotmail.com","Panel123");//parantez içindeki 2. tırnak içine mailin şifresini yazman gerek
            istemci.Port = 587;
            istemci.Host = "smtp.live.com";
            istemci.EnableSsl = true;
            mesajım.To.Add("star_wars11@hotmail.com");
            mesajım.From = new MailAddress("daikin_panel@hotmail.com");
            mesajım.Subject = "Uyarı!";
            mesajım.Body = PnlKodTxt.Text + " kodlu punch dosyası bulunamadı.";
            object userState = mesajım;
            

            istemci.Send(message: mesajım);
        }
        private void metroTile21_Click(object sender, EventArgs e)
        {
            PnlTxt.Text = ("NP");
            PnlTxt.BackColor = Color.MediumAquamarine;
            PanelLbl.BackColor = Color.MediumAquamarine;
        }
        private void metroTile11_Click(object sender, EventArgs e)
        {
            PnlTxt.Text = ("NP");
            PnlTxt.BackColor = Color.PowderBlue;
            PanelLbl.BackColor = Color.PowderBlue;
        }
        private void TmzleBtn_Click(object sender, EventArgs e)
        {
            ListViewKod.Items.Clear();
            pswrdTxt.Clear();
            sayiTxt.Clear();
        }
        private void timer2_Tick(object sender, EventArgs e)
        {
            krnmtre.Text = ((Convert.ToString(saat) + " : ") + (Convert.ToString(dakika) + " : ") + Convert.ToString(saniye));
            if ((saniye == 59))
            {
                saniye = 00;
                dakika = dakika + 1;
                if (dakika == 60)
                {
                    saniye = 00;
                    dakika = 00;
                    saat = saat + 1;
                }
            }
            saniye = saniye + 1;
        }
        private void timer3_Tick(object sender, EventArgs e)
        {
            saatLbl.Text = DateTime.Now.ToString("HH:mm:ss");
            tarihLbl.Text = DateTime.Now.ToString("dd-MM-yyyy");
        }
        private void barkodAta_Click(object sender, EventArgs e)
        {
            DirectoryInfo df = new DirectoryInfo(kopyaKlasorTextBox.Text);
            foreach (FileInfo item in df.GetFiles("*.*"))
            {
                item.Delete();
                break;
            }

            {
                string[] files = Directory.GetFiles(kaynakKlasorTextBox.Text, barkodTxt.Text + ".nc", SearchOption.AllDirectories);

                if (files.Length == 0)
                {
                    MessageBox.Show("Punch Dosyası Bulunamadı.", "Başarısız", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    barkodTxt.Clear();
                    barkodTxt.Focus();
                }
                else
                {
                    MessageBox.Show("Punch Dosyası Bulundu.İşlemeye başlayabilirsiniz.", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    var dosyalar = new DirectoryInfo(kaynakKlasorTextBox.Text).GetFiles(barkodTxt.Text + ".nc", SearchOption.AllDirectories);
                    foreach (var d in dosyalar)
                    {
                        hedefKlasorTextBox.Text = d.FullName;
                    }
                    baglantı.Open();
                    OleDbCommand komut = new OleDbCommand("insert into punch (Panel_kod,Tarih,Panel_Izo,Sac_Tipi,Panel_Tipi,Dıs_Sac_Olcu,Ic_Sac_Olcu,Operatör_ismi,Vardiya_saati,islendigi_saat,Makina) values ('" + PnlKodTxt.Text.ToString() + "','" + ZmnLbl.Text.ToString() + "','" + IzoTxt.Text.ToString() + "','" + SacTxt.Text.ToString() + "','" + PnlTxt.Text.ToString() + "','" + DısOlcu.Text.ToString() + "','" + IcOlcu.Text.ToString() + "','" + IsimCombo.Text.ToString() + "','" + VardiyaCombo.Text.ToString() + "','" + saatLbl.Text.ToString() + "','" + MakinaCombo.Text.ToString() + "')", baglantı);
                    komut.ExecuteNonQuery();
                    baglantı.Close();
                    DirectoryInfo dm = new DirectoryInfo(kopyaKlasorTextBox.Text);
                    if (pswrdTxt.Text == "123456")
                    {
                        verilerigörüntüle();
                    }
                    foreach (FileInfo item in dm.GetFiles("*.*"))
                    {
                        item.Delete();
                        break;
                    }
                    string fileName = barkodTxt.Text + ".nc";
                    string targetPath = kopyaKlasorTextBox.Text;
                    string sourceFile = hedefKlasorTextBox.Text;//Path.Combine(sourcePath, fileName);
                    string destFile = Path.Combine(targetPath, fileName);
                    File.Copy(sourceFile, destFile, true);
                    string myPath = kopyaKlasorTextBox.Text;
                    System.Diagnostics.Process prc = new System.Diagnostics.Process();
                    prc.StartInfo.FileName = myPath;
                    prc.Start();
                    DısOlcu.Clear();
                    IcOlcu.Clear();
                    PnlKodTxt.Clear();
                    barkodTxt.Clear();
                    barkodTxt.Focus();
                }
            }
        }
        private void KntrolBtn_Click(object sender, EventArgs e)
        {
            {

                DirectoryInfo df = new DirectoryInfo(kopyaKlasorTextBox.Text);
                foreach (FileInfo item in df.GetFiles("*.*"))
                {
                    item.Delete();
                    break;
                }

                {
                    string[] files = Directory.GetFiles(kaynakKlasorTextBox.Text, PnlKodTxt.Text + ".nc", SearchOption.AllDirectories);

                    if (files.Length == 0)
                    {
                        MessageBox.Show("Punch Dosyası Bulunamadı.", "Başarısız", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        barkodTxt.Clear();
                        barkodTxt.Focus();
                    }
                    else
                    {
                        MessageBox.Show("Punch Dosyası Bulundu.İşlemeye başlayabilirsiniz.", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        var dosyalar = new DirectoryInfo(kaynakKlasorTextBox.Text).GetFiles(PnlKodTxt.Text + ".nc", SearchOption.AllDirectories);
                        foreach (var d in dosyalar)
                        {
                            hedefKlasorTextBox.Text = d.FullName;
                        }
                        baglantı.Open();
                        OleDbCommand komut = new OleDbCommand("insert into punch (Panel_kod,Tarih,Panel_Izo,Sac_Tipi,Panel_Tipi,Dıs_Sac_Olcu,Ic_Sac_Olcu,Operatör_ismi,Vardiya_saati,islendigi_saat,Makina) values ('" + PnlKodTxt.Text.ToString() + "','" + ZmnLbl.Text.ToString() + "','" + IzoTxt.Text.ToString() + "','" + SacTxt.Text.ToString() + "','" + PnlTxt.Text.ToString() + "','" + DısOlcu.Text.ToString() + "','" + IcOlcu.Text.ToString() + "','" + IsimCombo.Text.ToString() + "','" + VardiyaCombo.Text.ToString() + "','" + saatLbl.Text.ToString() + "','" + MakinaCombo.Text.ToString() + "')", baglantı);
                        komut.ExecuteNonQuery();
                        baglantı.Close();
                        DirectoryInfo dm = new DirectoryInfo(kopyaKlasorTextBox.Text);
                        if (barkodTxt.Text == "123456")
                        {
                            verilerigörüntüle();
                        }
                        foreach (FileInfo item in dm.GetFiles("*.*"))
                        {
                            item.Delete();
                            break;
                        }
                        string fileName = PnlKodTxt.Text + ".nc";
                        string targetPath = kopyaKlasorTextBox.Text;
                        string sourceFile = hedefKlasorTextBox.Text;//Path.Combine(sourcePath, fileName);
                        string destFile = Path.Combine(targetPath, fileName);
                        File.Copy(sourceFile, destFile, true);
                        string myPath = kopyaKlasorTextBox.Text;
                        System.Diagnostics.Process prc = new System.Diagnostics.Process();
                        prc.StartInfo.FileName = myPath;
                        prc.Start();
                        DısOlcu.Clear();
                        IcOlcu.Clear();
                        PnlKodTxt.Clear();
                        barkodTxt.Clear();
                        barkodTxt.Focus();

                    }
                }
            }
        }
        private void DurBtn_Click(object sender, EventArgs e)
        {           
                timer2.Enabled = false;
                DrdSaatLbl.Text = DateTime.Now.ToString("HH:mm:ss");
                baglantı.Open();
                OleDbCommand tt = new OleDbCommand("insert into duruslar(Durus_Baslama_Saati,Durus_Bitis_Saati,Durus_Suresi,Durus_Kodu,Durus_Tarih,Operatör_ismi) values ('" + BslaSaatLbl.Text.ToString() + "','" + DrdSaatLbl.Text.ToString() + "','" + krnmtre.Text.ToString() + "','" + DurusCombo.Text.ToString() + "','" + ZmnLbl.Text.ToString() + "','" + IsimCombo.Text.ToString() + "')", baglantı);
                tt.ExecuteNonQuery();
                baglantı.Close();                           
        }
        private void SıfırBtn_Click(object sender, EventArgs e)
        {
            dakika = 00;
            saat = 00;
            saniye = 00;
            timer2.Enabled = false;
        }
        private void metroDateTime1_ValueChanged(object sender, EventArgs e)
        {
            textBox1.Text = bslaSrguDt.Text + " " + BslaSaat.Text;
        }
        private void BslaSaat_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = bslaSrguDt.Text + " " + BslaSaat.Text;
        }
        private void btrSrguDt_ValueChanged(object sender, EventArgs e)
        {
            textBox2.Text = btrSrguDt.Text + " " + BtrSaat.Text;
        }
        private void TrhSorgu_Click(object sender, EventArgs e)
        {           
            string zmn = textBox1.Text;
            string mbd = textBox2.Text;
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = baglantı;
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            baglantı.Open();
            cmd.CommandText = "select * from punch where Tarih between @Tarih1 and @Tarih2";
            cmd.CommandType = CommandType.Text;
            cmd.Parameters.AddWithValue("@Tarih1", zmn);
            cmd.Parameters.AddWithValue("@Tarih2", mbd);
            da.SelectCommand = cmd;
            da.Fill(ds);
            TarihGrid.DataSource = ds.Tables[0];
            baglantı.Close();
            TarihGrid.Style = MetroFramework.MetroColorStyle.Green;
        }
        private void BtrSaat_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox2.Text = btrSrguDt.Text + " " + BtrSaat.Text;
        }
        private void barkodTxt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                barkodTxt.Text = barkodTxt.Text.Replace(" ", "_");
                barkodTxt.Text = barkodTxt.Text.Replace("X", "x");
                IzoTxt.Text = barkodTxt.Text.Substring(0, 2);
                SacTxt.Text = barkodTxt.Text.Substring(3, 1);
                PnlTxt.Text = barkodTxt.Text.Substring(barkodTxt.Text.LastIndexOf('_') + 1);
                int kel = barkodTxt.Text.IndexOf('x');
                string sel = barkodTxt.Text.Substring(0, kel);
                string al = sel.Substring(sel.IndexOf('_') + 1);
                string ver = al.Substring(al.IndexOf('_') + 1);
                DısOlcu.Text = ver;
                int kil = barkodTxt.Text.LastIndexOf('_');
                string sil = barkodTxt.Text.Substring(0, kil);
                string dil = sil.Substring(sil.IndexOf('x') + 1);
                IcOlcu.Text = dil;

            }
        }
        private void KydetBtn_Click(object sender, EventArgs e)
        {
            baglantı.Open();
            OleDbCommand kmd = new OleDbCommand("insert into personel(Operatör_ismi)values ('" + IsimKayıtTxt.Text.ToString() + "')", baglantı);
            kmd.ExecuteNonQuery();            
            baglantı.Close();           
            IsimKayıtTxt.Clear();
        }
        private void KpyaBtn_Click(object sender, EventArgs e)
        {
            string myPath = kopyaKlasorTextBox.Text;
            System.Diagnostics.Process prc = new System.Diagnostics.Process();
            prc.StartInfo.FileName = myPath;
            prc.Start();
        }
        private void access_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"C:\Users\star_\OneDrive\Masaüstü\punchdkn.mdb");
        }
        private void metroDateTime1_ValueChanged_1(object sender, EventArgs e)
        {
            textBox3.Text = metroDateTime1.Text + " " + metroComboBox1.Text;
        }
        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox3.Text = metroDateTime1.Text + " " + metroComboBox1.Text;
        }
        private void metroDateTime2_ValueChanged(object sender, EventArgs e)
        {
            textBox4.Text = metroDateTime2.Text + " " + metroComboBox2.Text;
        }
        private void metroComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox4.Text = metroDateTime2.Text + " " + metroComboBox2.Text;
        }
        private void metroButton1_Click(object sender, EventArgs e)
        {
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = baglantı;
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            baglantı.Open();
            cmd.CommandText = "select* from Duruslar where Durus_Tarih between @Tarih1 and @Tarih2";
            cmd.CommandType = CommandType.Text;
            cmd.Parameters.AddWithValue("@Tarih1", textBox3.Text);
            cmd.Parameters.AddWithValue("@Tarih2", textBox4.Text);
            da.SelectCommand = cmd;
            da.Fill(ds);
            TarihGrid.DataSource = ds.Tables[0];
            baglantı.Close();
            TarihGrid.Style = MetroFramework.MetroColorStyle.Red;           
        }
        private void ExclTile_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j < TarihGrid.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = TarihGrid.Columns[j].HeaderText;
            }
            StartRow++;
            for (int i = 0; i < TarihGrid.Rows.Count; i++)
            {
                for (int j = 0; j < TarihGrid.Columns.Count; j++)
                {
                    Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 = TarihGrid[j, i].Value == null ? "" : TarihGrid[j, i].Value;
                    myRange.Select();
                }
            }
        }
        private void metroButton2_Click(object sender, EventArgs e)
        {
            TarihGrid.Columns.Clear();
        }
        private void PdfTile_Click(object sender, EventArgs e)
        {
            pdfKaydet(TarihGrid, "text");
        }
        private void IsimKayıtTxt_TextChanged(object sender, EventArgs e)
        {
            IsimKayıtTxt.Text = IsimKayıtTxt.Text.ToUpper();
            IsimKayıtTxt.SelectionStart = IsimKayıtTxt.Text.Length;
        }

        private void MetroTile19_Click(object sender, EventArgs e)
        {
            PnlTxt.Clear();
            System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
            PnlTxt.Text += btn.Text;
            if (PnlTxt.Text == "PK")
            PanelLbl.BackColor = Color.YellowGreen;
        }

        private void pswrdTxt_TextChanged(object sender, EventArgs e)
        {
            if (pswrdTxt.Text == "123456")
            {
                userPic.BackColor = Color.Lime;
                kilitpic.BackColor = Color.Lime;
                label2.Visible = false;
                TrhSorgu.Enabled = true;
                metroButton1.Enabled = true;
                ExclTile.Enabled = true;
                access.Enabled = true;
                PdfTile.Enabled = true;
                KynkKlsrBtn.Enabled = true;
                KpyaKlsrBtn.Enabled = true;
            }
            else
            {
                userPic.BackColor = Color.DeepSkyBlue;
                kilitpic.BackColor = Color.DeepSkyBlue;
                label2.Visible = true;
                TrhSorgu.Enabled = false;
                metroButton1.Enabled = false;
                ExclTile.Enabled = false;
                access.Enabled = false;
                PdfTile.Enabled = false;
                KynkKlsrBtn.Enabled = false;
                KpyaKlsrBtn.Enabled = false;
            }
        }
        private void baslaBtn_Click(object sender, EventArgs e)
        {
            DialogResult kayit = new DialogResult();
            kayit = MessageBox.Show("Duruş kaydı başlatılsın mı?", "UYARI", MessageBoxButtons.YesNo);
            if (kayit == DialogResult.Yes)
            {
                BslaSaatLbl.Text = DateTime.Now.ToString("HH:mm:ss");
                timer2.Enabled = true;
            }
        }
    }
}
