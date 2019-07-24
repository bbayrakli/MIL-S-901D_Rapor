using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using System.Drawing.Drawing2D;

namespace _901DD
{
    public partial class Form1 : Form
    {
        #region

        double beam160, beam200, beam250, beam_160, beam_200, beam_250;
        public void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                comboBox1.Visible = true;
                beam160 = (34 * Convert.ToDouble(comboBox1.SelectedItem));
                
            }
            else
            {
                comboBox1.Visible = false;
                beam160 = 0;
            }
        }

        public void CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                
                comboBox2.Visible = true;
                beam200 = (34 / 1.6 * 2 * Convert.ToDouble(comboBox2.SelectedItem));
            }
            else
            {
                comboBox2.Visible = false;
                beam200 = 0;
            }
        }

        public void CheckBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                comboBox3.Visible = true;
                beam250 = (34 / 1.6 * 2.5 * Convert.ToDouble(comboBox3.SelectedItem));
            }
            else
            {
                comboBox3.Visible = false;
                beam250 = 0;

            }

        }

        public void CheckBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                comboBox4.Visible = true;
                beam_160 = (34 * Convert.ToDouble(comboBox4.SelectedItem));
            }
            else
            {
                comboBox4.Visible = false;
            }
        }
        public void CheckBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                comboBox5.Visible = true;
                beam_200 = (34 / 1.6 * 2 * Convert.ToDouble(comboBox5.SelectedItem));
            }
            else
            {
                comboBox5.Visible = false;
                beam_200 = 0;
            }
        }
        public void CheckBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                comboBox6.Visible = true;
                beam_250 = (34 / 1.6 * 2.5 * Convert.ToDouble(comboBox6.SelectedItem));
            }
            else
            {
                comboBox6.Visible = false;
                beam_250 = 0;
            }
        }

        #endregion

        public Form1()
        {
            InitializeComponent();
            
            
        }
        public void Button2_Click_1(object sender, EventArgs e)
        {

            WrdFile.ShowDialog();
            FPath.Text = WrdFile.FileName;
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            textBox1.Text = openFileDialog1.FileName;
        }

        public void Button3_Click(object sender, EventArgs e)
        {
            
            jpgFolder.ShowDialog();
            picPath.Text = jpgFolder.SelectedPath;
        }
        public void Button4_Click(object sender, EventArgs e)
        {
            
            graphFolder.ShowDialog();
            graPath.Text = graphFolder.SelectedPath;
        }
        public void Button5_Click(object sender, EventArgs e)
        {
            SaveAs.ShowDialog();
            SaveDir.Text = SaveAs.SelectedPath;
        }

        public void Button1_Click(object sender, EventArgs e)
        {
            info.Text = "...değişken tanımlamaları yapılıyor.";
            string ekipman_adi, musteri_adi, kanal_1_seri, kanal_2_seri, kanal_3_seri, kanal_4_seri, seri_no, hazirlayan_1, musteri_adres, rfq_1, test_tar,
                   ekipman_tar, rapor_tar, grade_1, type_1, class_1, en_01, boy_01, yukseklik_01;

            double ekipman_kutle, fikstur_kutle, toplam_kutle, egikfiks_kutle, egik_kutle_toplam, clamps, egik_clamp, ekpmfkst,
                    dcy_1, dcy_2, dcy_3, ecy_1, ecy_2, ecy_3, dmy_1_m, dmy_2_m, dmy_3_m, emy_1_m, emy_2_m, emy_3_m;

            progressBar1.Value = 10;
            progressBar1.Update();
            



            //DEĞİŞKEN TANIMLARI
            //========================================================================================================================
            //STRING                                                                                                                  
            //========================================================================================================================
            //ekipman_adi   = Ekipman Adı            = ek_1.Textbox
            //musteri_adi   = Müşteri Adı            = ma_1.Textbox
            //kanal_n_seri  = n. Kanal Seri Numarası = chn.Textbox
            //seri_no       = Ekipman Seri Numarası  = eseri1.Textbox
            //hazirlayan    = Testi Yapan Kişi       = hazr_1.Textbox
            //musteri_adres = Müşteri Adresi         = madres_1.Textbox
            //rfq_1         = RFQ No                 = rfq1.Textbox
            //test_tar      = Test Tarihi            = ttarih.Textbox
            //ekipman_tar   = Ekipman Kabul Tarihi   = etarih.Textbox
            //rapor_tar     = Rapor Tarihi           = raptar.Textbox
            //grade_1       = Grade                  = grade_1.Combobox
            //type_1        = Type                   = type_1.Combobox
            //class_1       = Class                  = class_1.Combobox
            //en_01         = Ekipman Eni            = en_1.Textbox
            //boy_01        = Ekpman Boyu            = boy_1.Textbox
            //yukseklik_01  = Ekipman Yüksekliği     = yuk_1.Textbox

            //=========================================================================================================================
            //DOUBLE
            //=========================================================================================================================
            //ekipman_kutle  = Ekipman Kütlesi                 = ek_1.Textbox
            //fikstur_kutle  = Fikstür Kütlesi                 = fk_1.Textbox
            //toplam_kutle   = Toplam Kütle (Dikey)            = ekipman_kutle + fikstur_kutle
            //pound          = Ingiliz Ağırlık Birimi          = toplam_kutle * (2.204)
            //efikstu_kutle  = Eğik Fikstür Kütlesi            = efk_1.TExtbox
            //egik_kutle_toplam   = Eğik Fikstür Toplam Kütle       = efikstu_kutle + ekipman_kutle
            //egiklbs_1        = Ingiliz Ağırlık Birimi          = egik_kutle_toplam * (2.2)


            //=========================================================================================================================
            //DOUBLE
            //=========================================================================================================================
            //dcy_n          = Dikey Çekiç Yüksekliği n. Vuruş = Tablodan Hesaplanacak
            //ecy_n          = Eğik Çekiç Yüksekliği n. Vuruş  = Tablodan Hesaplanacak
            //dmy_n_m        = Dikey Çekiç Yüksekliği Metrk    = dcy_n * (304.8)
            //emy_n_m        = Eğik Çekiç Yüksekliği Metrik    = ecy_n * (304.8)
            //=========================================================================================================================
            //=========================================================================================================================


            //=========================================================================================================================
            //DEĞİŞKENLERE DEĞER ATAMA
            //=========================================================================================================================
            info.Text = "...değişken değerleri atanıyor.";
            ekipman_adi = ea_1.Text;
            musteri_adi = ma_1.Text;
            kanal_1_seri = ch1.Text;
            kanal_2_seri = ch2.Text;
            kanal_3_seri = ch3.Text;
            kanal_4_seri = ch4.Text;
            seri_no = eseri1.Text;
            hazirlayan_1 = hzr_1.Text;
            musteri_adres = musteri_adresi.Text;
            rfq_1 = rfq1.Text;
            test_tar = ttarih.Text;
            ekipman_tar = etarih.Text;
            rapor_tar = raptar.Text;
            grade_1 = grade_01.Text;
            type_1 = Type_01.Text;
            class_1 = class_01.Text;
            en_01 = en_1.Text;
            boy_01 = boy_1.Text;
            yukseklik_01 = yuk_1.Text;
            ekipman_kutle = Convert.ToDouble(ek_1.Text);
            fikstur_kutle = Convert.ToDouble(fk_1.Text);
            clamps = beam160 + beam200 + beam250 + 40;
            egik_clamp = beam_160 + beam_200 + beam_250 + 50;
            toplam_kutle = ekipman_kutle + fikstur_kutle + clamps;
            double pound = toplam_kutle * 2.2;
            egikfiks_kutle = 771;
            egik_kutle_toplam = egikfiks_kutle + ekipman_kutle + fikstur_kutle + egik_clamp;
            double egiklbs_1 = egik_kutle_toplam * 2.2;
            ekpmfkst = ekipman_kutle + fikstur_kutle;

            progressBar1.Value = 20;
            progressBar1.Update();
            //=========================================================================================================================
            //BOŞ ALAN UYARILARI
            //=========================================================================================================================

            #region

            if (ea_1.Text == "")
            {
                MessageBox.Show("Ekipman adı giriniz");
                System.Windows.Forms.Application.Exit();
                return;
                
            }
            if (ma_1.Text == "")
            {
                MessageBox.Show("Müşteri Adı giriniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (eseri1.Text == "")
            {
                MessageBox.Show("Seri no giriniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (hzr_1.Text == "")
            {
                MessageBox.Show("Hazırlayanı giriniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (musteri_adresi.Text == "")
            {
                MessageBox.Show("Müşteri Adresi giriniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (rfq1.Text == "")
            {
                MessageBox.Show("RFQ no giriniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (ttarih.Text == "")
            {
                MessageBox.Show("Test Tarihi giriniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (raptar.Text == "")
            {
                MessageBox.Show("Rapor Tarihi giriniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (etarih.Text == "")
            {
                MessageBox.Show("Ekipman Kabul Tarihi giriniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (class_01.Text == "")
            {
                MessageBox.Show("Class seçiniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (Type_01.Text == "")
            {
                MessageBox.Show("Type seçiniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (grade_01.Text == "")
            {
                MessageBox.Show("Grade seçiniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (en_1.Text == "")
            {
                MessageBox.Show("Ekipman Enini giriniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (boy_1.Text == "")
            {
                MessageBox.Show("Ekipman Boyunu giriniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (yuk_1.Text == "")
            {
                MessageBox.Show("Ekipman Yüksekliğini giriniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (ek_1.Text == "")
            {
                MessageBox.Show("Ekipman Kütlesini giriniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            if (fk_1.Text == "")
            {
                MessageBox.Show("Fikstür Kütlesini giriniz");
                System.Windows.Forms.Application.Exit();
                return;
            }
            
            #endregion


            //=========================================================================================================================
            //901D TABLO İŞLEMLERİ
            //=========================================================================================================================
            info.Text = "...çekiç yükseklikleri hesaplanıyor.";

            #region

            if (pound < 1000)
            {

                dcy_1 = 0.75;
                dcy_2 = 1.75;
                dcy_3 = 1.75;
                dmy_1_m = dcy_1 * 304.8;
                dmy_2_m = dcy_2 * 304.8;
                dmy_3_m = dcy_3 * 304.8;
            }

            else
            {
                if (1000 <= pound && pound < 2000)
                {

                    dcy_1 = 1.00;
                    dcy_2 = 2.00;
                    dcy_3 = 2.00;
                    dmy_1_m = dcy_1 * 304.8;
                    dmy_2_m = dcy_2 * 304.8;
                    dmy_3_m = dcy_3 * 304.8;

                }

                else
                {
                    if (2000 <= pound && pound < 3000)
                    {
                        dcy_1 = 1.25;
                        dcy_2 = 2.25;
                        dcy_3 = 2.25;
                        dmy_1_m = dcy_1 * 304.8;
                        dmy_2_m = dcy_2 * 304.8;
                        dmy_3_m = dcy_3 * 304.8;
                    }

                    else
                    {
                        if (3000 <= pound && pound < 3500)
                        {

                            dcy_1 = 1.50;
                            dcy_2 = 2.50;
                            dcy_3 = 2.50;
                            dmy_1_m = dcy_1 * 304.8;
                            dmy_2_m = dcy_2 * 304.8;
                            dmy_3_m = dcy_3 * 304.8;
                        }

                        else
                        {
                            if (3500 <= pound && pound < 4000)
                            {
                                dcy_1 = 1.75;
                                dcy_2 = 2.75;
                                dcy_3 = 2.75;
                                dmy_1_m = dcy_1 * 304.8;
                                dmy_2_m = dcy_2 * 304.8;
                                dmy_3_m = dcy_3 * 304.8;
                            }

                            else
                            {
                                if (4000 <= pound && pound < 4200)
                                {
                                    dcy_1 = 2.00;
                                    dcy_2 = 3.00;
                                    dcy_3 = 3.00;
                                    dmy_1_m = dcy_1 * 304.8;
                                    dmy_2_m = dcy_2 * 304.8;
                                    dmy_3_m = dcy_3 * 304.8;
                                }

                                else
                                {
                                    if (4200 <= pound && pound < 4400)
                                    {
                                        dcy_1 = 2.00;
                                        dcy_2 = 3.25;
                                        dcy_3 = 3.25;
                                        dmy_1_m = dcy_1 * 304.8;
                                        dmy_2_m = dcy_2 * 304.8;
                                        dmy_3_m = dcy_3 * 304.8;
                                    }

                                    else
                                    {
                                        if (4400 <= pound && pound < 4600)
                                        {
                                            dcy_1 = 2.00;
                                            dcy_2 = 3.50;
                                            dcy_3 = 3.50;
                                            dmy_1_m = dcy_1 * 304.8;
                                            dmy_2_m = dcy_2 * 304.8;
                                            dmy_3_m = dcy_3 * 304.8;
                                        }

                                        else
                                        {
                                            if (4600 <= pound && pound < 4800)
                                            {
                                                dcy_1 = 2.25;
                                                dcy_2 = 3.75;
                                                dcy_3 = 3.75;
                                                dmy_1_m = dcy_1 * 304.8;
                                                dmy_2_m = dcy_2 * 304.8;
                                                dmy_3_m = dcy_3 * 304.8;
                                            }
                                            else
                                            {
                                                if (4800 <= pound && pound < 5000)
                                                {
                                                    dcy_1 = 2.25;
                                                    dcy_2 = 4.00;
                                                    dcy_3 = 4.00;
                                                    dmy_1_m = dcy_1 * 304.8;
                                                    dmy_2_m = dcy_2 * 304.8;
                                                    dmy_3_m = dcy_3 * 304.8;
                                                }

                                                else
                                                {
                                                    if (5000 <= pound && pound < 5200)
                                                    {
                                                        dcy_1 = 2.50;
                                                        dcy_2 = 4.50;
                                                        dcy_3 = 4.50;
                                                        dmy_1_m = dcy_1 * 304.8;
                                                        dmy_2_m = dcy_2 * 304.8;
                                                        dmy_3_m = dcy_3 * 304.8;

                                                    }

                                                    else
                                                    {
                                                        if (5200 <= pound && pound < 5400)
                                                        {
                                                            dcy_1 = 2.50;
                                                            dcy_2 = 5.00;
                                                            dcy_3 = 5.00;
                                                            dmy_1_m = dcy_1 * 304.8;
                                                            dmy_2_m = dcy_2 * 304.8;
                                                            dmy_3_m = dcy_3 * 304.8;
                                                        }
                                                        else
                                                        {
                                                            if (5400 <= pound && pound < 5600)
                                                            {
                                                                dcy_1 = 2.50;
                                                                dcy_2 = 5.50;
                                                                dcy_3 = 5.50;
                                                                dmy_1_m = dcy_1 * 304.8;
                                                                dmy_2_m = dcy_2 * 304.8;
                                                                dmy_3_m = dcy_3 * 304.8;
                                                            }

                                                            else
                                                            {
                                                                if (5600 <= pound && pound < 6200)

                                                                {
                                                                    dcy_1 = 2.75;
                                                                    dcy_2 = 5.50;
                                                                    dcy_3 = 5.50;
                                                                    dmy_1_m = dcy_1 * 304.8;
                                                                    dmy_2_m = dcy_2 * 304.8;
                                                                    dmy_3_m = dcy_3 * 304.8;
                                                                }
                                                                else
                                                                {
                                                                    if (6200 <= pound && pound < 6800)
                                                                    {
                                                                        dcy_1 = 3.00;
                                                                        dcy_2 = 5.50;
                                                                        dcy_3 = 5.50;
                                                                        dmy_1_m = dcy_1 * 304.8;
                                                                        dmy_2_m = dcy_2 * 304.8;
                                                                        dmy_3_m = dcy_3 * 304.8;
                                                                    }

                                                                    else
                                                                    {
                                                                        if (6800 <= pound && pound < 7400)
                                                                        {
                                                                            dcy_1 = 3.25;
                                                                            dcy_2 = 5.50;
                                                                            dcy_3 = 5.50;
                                                                            dmy_1_m = dcy_1 * 304.8;
                                                                            dmy_2_m = dcy_2 * 304.8;
                                                                            dmy_3_m = dcy_3 * 304.8;
                                                                        }
                                                                        else
                                                                        {
                                                                            dcy_1 = 3.25;
                                                                            dcy_2 = 5.50;
                                                                            dcy_3 = 5.50;
                                                                            dmy_1_m = dcy_1 * 304.8;
                                                                            dmy_2_m = dcy_2 * 304.8;
                                                                            dmy_3_m = dcy_3 * 304.8;
                                                                        }

                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }


                                }

                            }

                        }
                    }
                }
            }


            if (egiklbs_1 < 1000)
            {

                ecy_1 = 0.75;
                ecy_2 = 1.75;
                ecy_3 = 1.75;
                emy_1_m = ecy_1 * 304.8;
                emy_2_m = ecy_2 * 304.8;
                emy_3_m = ecy_3 * 304.8;
            }

            else
            {
                if (1000 <= egiklbs_1 && pound < 2000)
                {

                    ecy_1 = 1.00;
                    ecy_2 = 2.00;
                    ecy_3 = 2.00;
                    emy_1_m = ecy_1 * 304.8;
                    emy_2_m = ecy_2 * 304.8;
                    emy_3_m = ecy_3 * 304.8;

                }

                else
                {
                    if (2000 <= egiklbs_1 && egiklbs_1 < 3000)
                    {
                        ecy_1 = 1.25;
                        ecy_2 = 2.25;
                        ecy_3 = 2.25;
                        emy_1_m = ecy_1 * 304.8;
                        emy_2_m = ecy_2 * 304.8;
                        emy_3_m = ecy_3 * 304.8;
                    }

                    else
                    {
                        if (3000 <= egiklbs_1 && egiklbs_1 < 3500)
                        {

                            ecy_1 = 1.50;
                            ecy_2 = 2.50;
                            ecy_3 = 2.50;
                            emy_1_m = ecy_1 * 304.8;
                            emy_2_m = ecy_2 * 304.8;
                            emy_3_m = ecy_3 * 304.8;
                        }

                        else
                        {
                            if (3500 <= egiklbs_1 && egiklbs_1 < 4000)
                            {
                                ecy_1 = 1.75;
                                ecy_2 = 2.75;
                                ecy_3 = 2.75;
                                emy_1_m = ecy_1 * 304.8;
                                emy_2_m = ecy_2 * 304.8;
                                emy_3_m = ecy_3 * 304.8;
                            }

                            else
                            {
                                if (4000 <= egiklbs_1 && egiklbs_1 < 4200)
                                {
                                    ecy_1 = 2.00;
                                    ecy_2 = 3.00;
                                    ecy_3 = 3.00;
                                    emy_1_m = ecy_1 * 304.8;
                                    emy_2_m = ecy_2 * 304.8;
                                    emy_3_m = ecy_3 * 304.8;
                                }

                                else
                                {
                                    if (4200 <= egiklbs_1 && egiklbs_1 < 4400)
                                    {
                                        ecy_1 = 2.00;
                                        ecy_2 = 3.25;
                                        ecy_3 = 3.25;
                                        emy_1_m = ecy_1 * 304.8;
                                        emy_2_m = ecy_2 * 304.8;
                                        emy_3_m = ecy_3 * 304.8;
                                    }

                                    else
                                    {
                                        if (4400 <= egiklbs_1 && egiklbs_1 < 4600)
                                        {
                                            ecy_1 = 2.00;
                                            ecy_2 = 3.50;
                                            ecy_3 = 3.50;
                                            emy_1_m = ecy_1 * 304.8;
                                            emy_2_m = ecy_2 * 304.8;
                                            emy_3_m = ecy_3 * 304.8;
                                        }

                                        else
                                        {
                                            if (4600 <= egiklbs_1 && egiklbs_1 < 4800)
                                            {
                                                ecy_1 = 2.25;
                                                ecy_2 = 3.75;
                                                ecy_3 = 3.75;
                                                emy_1_m = ecy_1 * 304.8;
                                                emy_2_m = ecy_2 * 304.8;
                                                emy_3_m = ecy_3 * 304.8;
                                            }
                                            else
                                            {
                                                if (4800 <= egiklbs_1 && egiklbs_1 < 5000)
                                                {
                                                    ecy_1 = 2.25;
                                                    ecy_2 = 4.00;
                                                    ecy_3 = 4.00;
                                                    emy_1_m = ecy_1 * 304.8;
                                                    emy_2_m = ecy_2 * 304.8;
                                                    emy_3_m = ecy_3 * 304.8;
                                                }

                                                else
                                                {
                                                    if (5000 <= egiklbs_1 && egiklbs_1 < 5200)
                                                    {
                                                        ecy_1 = 2.50;
                                                        ecy_2 = 4.50;
                                                        ecy_3 = 4.50;
                                                        emy_1_m = ecy_1 * 304.8;
                                                        emy_2_m = ecy_2 * 304.8;
                                                        emy_3_m = ecy_3 * 304.8;

                                                    }

                                                    else
                                                    {
                                                        if (5200 <= egiklbs_1 && egiklbs_1 < 5400)
                                                        {
                                                            ecy_1 = 2.50;
                                                            ecy_2 = 5.00;
                                                            ecy_3 = 5.00;
                                                            emy_1_m = ecy_1 * 304.8;
                                                            emy_2_m = ecy_2 * 304.8;
                                                            emy_3_m = ecy_3 * 304.8;
                                                        }
                                                        else
                                                        {
                                                            if (5400 <= egiklbs_1 && egiklbs_1 < 5600)
                                                            {
                                                                ecy_1 = 2.50;
                                                                ecy_2 = 5.50;
                                                                ecy_3 = 5.50;
                                                                emy_1_m = ecy_1 * 304.8;
                                                                emy_2_m = ecy_2 * 304.8;
                                                                emy_3_m = ecy_3 * 304.8;
                                                            }

                                                            else
                                                            {
                                                                if (5600 <= egiklbs_1 && egiklbs_1 < 6200)

                                                                {
                                                                    ecy_1 = 2.75;
                                                                    ecy_2 = 5.50;
                                                                    ecy_3 = 5.50;
                                                                    emy_1_m = ecy_1 * 304.8;
                                                                    emy_2_m = ecy_2 * 304.8;
                                                                    emy_3_m = ecy_3 * 304.8;
                                                                }
                                                                else
                                                                {
                                                                    if (6200 <= egiklbs_1 && egiklbs_1 < 6800)
                                                                    {
                                                                        ecy_1 = 3.00;
                                                                        ecy_2 = 5.50;
                                                                        ecy_3 = 5.50;
                                                                        emy_1_m = ecy_1 * 304.8;
                                                                        emy_2_m = ecy_2 * 304.8;
                                                                        emy_3_m = ecy_3 * 304.8;
                                                                    }

                                                                    else
                                                                    {
                                                                        if (6800 <= egiklbs_1 && egiklbs_1 < 7400)
                                                                        {
                                                                            ecy_1 = 3.25;
                                                                            ecy_2 = 5.50;
                                                                            ecy_3 = 5.50;
                                                                            emy_1_m = ecy_1 * 304.8;
                                                                            emy_2_m = ecy_2 * 304.8;
                                                                            emy_3_m = ecy_3 * 304.8;

                                                                        }
                                                                        else
                                                                        {
                                                                            ecy_1 = 3.25;
                                                                            ecy_2 = 5.50;
                                                                            ecy_3 = 5.50;
                                                                            emy_1_m = ecy_1 * 304.8;
                                                                            emy_2_m = ecy_2 * 304.8;
                                                                            emy_3_m = ecy_3 * 304.8;
                                                                        }

                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }


                                }

                            }

                        }
                    }
                }
            }
            #endregion

            //=========================================================================================================================
            //DEĞİŞKENLERİN DEĞERLERİ ATANDI
            //=========================================================================================================================

            #region

            Microsoft.Office.Interop.Word.Application wrd = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = wrd.Documents.Open(WrdFile.FileName);
            wrd.Visible = false;
            //=========================================================================================================================
            //WORD ÜZERİNDEKİ DEĞİŞKENLERİ DEĞİŞTİRME
            //=========================================================================================================================
            //Ekipman Adını Değiştir
            //=========================================================================================================================
            info.Text = "...gerekli bilgiler giriliyor.";
            var find = doc.Range().Find;
            find.Text = "<ekipman_adi>";
            find.Format = true;
            find.Replacement.Text = ekipman_adi;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Müşteri Adını Değiştir
            //=========================================================================================================================
            find.Text = "<musteri_adi>";
            find.Format = true;
            find.Replacement.Text = musteri_adi;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Kanal 1 Seri No Değiştir
            //=========================================================================================================================
            find.Text = "<kanal_1_seri>";
            find.Format = true;
            find.Replacement.Text = kanal_1_seri;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Kanal 2 Seri No Değiştir
            //=========================================================================================================================
            find.Text = "<kanal_2_seri>";
            find.Format = true;
            find.Replacement.Text = kanal_2_seri;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Kanal 3 Seri No Değiştir
            //=========================================================================================================================
            find.Text = "<kanal_3_seri>";
            find.Format = true;
            find.Replacement.Text = kanal_3_seri;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Kanal 4 Seri No Değiştir
            //=========================================================================================================================
            find.Text = "<kanal_4_seri>";
            find.Format = true;
            find.Replacement.Text = kanal_4_seri;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Seri No Değiştir
            //=========================================================================================================================
            find.Text = "<seri_no>";
            find.Format = true;
            find.Replacement.Text = seri_no;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Hazırlayanı Değiştir
            //=========================================================================================================================
            find.Text = "<hazirlayan>";
            find.Format = true;
            find.Replacement.Text = hazirlayan_1;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Müşteri Adresi Değiştir
            //=========================================================================================================================
            find.Text = "<musteri_adres>";
            find.Format = true;
            find.Replacement.Text = musteri_adres;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //RFQ No Değiştir
            //=========================================================================================================================
            find.Text = "<rfq_1>";
            find.Format = true;
            find.Replacement.Text = rfq_1;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Test Tarihi Değiştir
            //=========================================================================================================================
            find.Text = "<test_tar>";
            find.Format = true;
            find.Replacement.Text = test_tar;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Ekipman Tarihi Değiştir
            //=========================================================================================================================
            find.Text = "<ekipman_tar>";
            find.Format = true;
            find.Replacement.Text = ekipman_tar;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Rapor Tarihi Değiştir
            //=========================================================================================================================
            find.Text = "<rapor_tar>";
            find.Format = true;
            find.Replacement.Text = rapor_tar;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Grade Değiştir
            //=========================================================================================================================
            find.Text = "<grade_1>";
            find.Format = true;
            find.Replacement.Text = grade_1;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Type Değiştir
            //=========================================================================================================================
            find.Text = "<type_1>";
            find.Format = true;
            find.Replacement.Text = type_1;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Class Değiştir
            //=========================================================================================================================
            find.Text = "<class_1>";
            find.Format = true;
            find.Replacement.Text = class_1;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //En Boy Yükseklik Değiştir
            //=========================================================================================================================
            find.Text = "<en_01>";
            find.Format = true;
            find.Replacement.Text = en_01;
            find.Execute(Replace: WdReplace.wdReplaceAll);
            find.Text = "<boy_01>";
            find.Format = true;
            find.Replacement.Text = boy_01;
            find.Execute(Replace: WdReplace.wdReplaceAll);
            find.Text = "<yukseklik_01>";
            find.Format = true;
            find.Replacement.Text = yukseklik_01;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            progressBar1.Value = 30;
            progressBar1.Update();
            

            //=========================================================================================================================
            //Ekipman Kütlesi Değiştir
            //=========================================================================================================================
            find.Text = "<ekipman_kutle>";
            find.Format = true;
            find.Replacement.Text = ekipman_kutle.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Fikstür Kütlesi Değiştir
            //=========================================================================================================================
            find.Text = "<fikstur_kutle>"; //özel imal edilmiş test jigi
            find.Format = true;
            find.Replacement.Text = fikstur_kutle.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            find.Text = "<clamps_kutle>"; //dikey test fikstürü beamler, civatalar ve clampler toplamı
            find.Format = true;
            find.Replacement.Text = Convert.ToString(clamps);
            find.Execute(Replace: WdReplace.wdReplaceAll);

            find.Text = "<ekpmfkst_kutle>";
            find.Format = true;
            find.Replacement.Text = Convert.ToString(ekpmfkst);
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Toplam Kütle Değiştir
            //=========================================================================================================================
            find.Text = "<toplam_kutle>";
            find.Format = true;
            find.Replacement.Text = toplam_kutle.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Pound Değeri Değiştir
            //=========================================================================================================================
            find.Text = "<pound_1>";
            find.Format = true;
            find.Replacement.Text = pound.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Eğik Fikstür Kütlesi Değiştir
            //=========================================================================================================================
            find.Text = "<efikstu_kutle>";
            find.Format = true;
            find.Replacement.Text = Convert.ToString(egikfiks_kutle + egik_clamp);
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Eğik Fikstür Toplam Kütle Değiştir
            //=========================================================================================================================
            find.Text = "<egik_kutle_toplam>";
            find.Format = true;
            find.Replacement.Text = egik_kutle_toplam.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //=========================================================================================================================
            //Eğik Fikstür Pound Değeri Değiştir
            //=========================================================================================================================
            find.Text = "<egklbs_1>";
            find.Format = true;
            find.Replacement.Text = egiklbs_1.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            progressBar1.Value = 40;
            

            //=========================================================================================================================
            //Çekiç Yüksekliklerini Değiştir
            //=========================================================================================================================
            find.Text = "<dcy_1>";
            find.Format = true;
            find.Replacement.Text = dcy_1.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);
            find.Text = "<dcy_2>";
            find.Format = true;
            find.Replacement.Text = dcy_2.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);
            find.Text = "<dcy_3>";
            find.Format = true;
            find.Replacement.Text = dcy_3.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);
            find.Text = "<ecy_1>";
            find.Format = true;
            find.Replacement.Text = ecy_1.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);
            find.Text = "<ecy_2>";
            find.Format = true;
            find.Replacement.Text = ecy_2.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);
            find.Text = "<ecy_3>";
            find.Format = true;
            find.Replacement.Text = ecy_3.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);
            find.Text = "<dmy_1_m>";
            find.Format = true;
            find.Replacement.Text = dmy_1_m.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);
            find.Text = "<dmy_2_m>";
            find.Format = true;
            find.Replacement.Text = dmy_2_m.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);
            find.Text = "<dmy_3_m>";
            find.Format = true;
            find.Replacement.Text = dmy_3_m.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);
            find.Text = "<emy_1_m>";
            find.Format = true;
            find.Replacement.Text = emy_1_m.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);
            find.Text = "<emy_2_m>";
            find.Format = true;
            find.Replacement.Text = emy_2_m.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);
            find.Text = "<emy_3_m>";
            find.Format = true;
            find.Replacement.Text = emy_3_m.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            progressBar1.Value = 50;
            progressBar1.Update();

            #endregion

            //=========================================================================================================================
            //HEADER ve FOOTER DEĞİŞİKLİKLERİ EKLENECEK !!!
            //=========================================================================================================================

            foreach (Microsoft.Office.Interop.Word.Section section in doc.Sections)
            {
                Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Find.Text = "<rfq_1>";
                headerRange.Find.Replacement.Text = rfq_1;
                headerRange.Find.Execute(Replace: WdReplace.wdReplaceAll);

                headerRange.Find.Text = "<ay_yil>";
                headerRange.Find.Replacement.Text = Convert.ToString(DateTime.Now.Month+"-"+DateTime.Now.Year);
                headerRange.Find.Execute(Replace: WdReplace.wdReplaceAll);

            }

            //=========================================================================================================================
            //EXCEL GRAFİKLERİNİ .JPG KAYDETME
            //=========================================================================================================================

                
                Microsoft.Office.Interop.Excel.Application exc = new Microsoft.Office.Interop.Excel.Application();
                exc.Visible = false;
                exc.Workbooks.Open(textBox1.Text);
                Worksheet sheet = exc.Worksheets[1];
                sheet.Activate();

                foreach (Worksheet ws in exc.Worksheets)
                {
                    ChartObjects chartobjects = ws.ChartObjects();

                    foreach (ChartObject co in chartobjects)
                    {
                        co.Select();
                        Microsoft.Office.Interop.Excel.Chart chart = co.Chart;
                        chart.Export(graphFolder.SelectedPath + @"\" + chart.Name + ".jpg");

                    }
                }
                exc.Workbooks.Close();
            
            //=========================================================================================================================
            //GRAFİKLERİ EKLEME
            //=========================================================================================================================
            double graphwidth, graphheight, j = 1;
            double p;

            while (File.Exists(graphFolder.SelectedPath + @"\Vurus01_Time_History Chart " + j + ".jpg"))
            {
                using (Image image = Image.FromFile(graphFolder.SelectedPath + @"\Vurus01_Time_History Chart "+j+".jpg"))
                {
                    p = (Convert.ToDouble(image.Size.Height) / Convert.ToDouble(image.Size.Width));

                    graphwidth = Convert.ToDouble(image.Width);
                    graphheight = Convert.ToDouble(image.Height);

                    if (p >= 1)
                    {
                        graphheight = 600;
                        graphwidth = graphheight / p;
                    }
                    else
                    {
                        graphwidth = 600;
                        graphheight = graphwidth * p;
                    }

                    new Bitmap(image, Convert.ToInt32(graphwidth), Convert.ToInt32(graphheight)).Save(graphFolder.SelectedPath + @"\graph_" + j + ".jpg");
                }

                object range = doc.Bookmarks["grp_" + j].Range;
                doc.InlineShapes.AddPicture(graphFolder.SelectedPath + @"\graph_" + j + ".jpg", false, true, ref range);
                j++;

               

            }

            //=========================================================================================================================
            //RESİMLERİ EKLEME 
            //=========================================================================================================================
            #region
            info.Text = "...resimler yükleniyor.";

            double width, height, i = 1;
            double c;

            while (File.Exists(jpgFolder.SelectedPath + @"\pic_" + i + ".jpg"))
            {
                using (Image image = Image.FromFile(jpgFolder.SelectedPath + @"\pic_" + i + ".jpg"))
                {
                    c = (Convert.ToDouble(image.Size.Height) / Convert.ToDouble(image.Size.Width));

                    width = Convert.ToDouble(image.Width);
                    height = Convert.ToDouble(image.Height);

                    if (c >= 1)
                    {
                        height = 600;
                        width = height / c;
                    }
                    else
                    {
                        width = 600;
                        height = width * c;
                    }
                                
                 new Bitmap(image, Convert.ToInt32(width), Convert.ToInt32(height)).Save(jpgFolder.SelectedPath + @"\pic_0" + i + ".jpg");
                }

                object rng = doc.Bookmarks["pic_" + i].Range;
                doc.InlineShapes.AddPicture(jpgFolder.SelectedPath + @"\pic_0" + i + ".jpg",false,true, ref rng);
                i++;

                progressBar1.Value = 50 + Convert.ToInt32(i);
                progressBar1.Update();
                
            }

            progressBar1.Value = 100;
            progressBar1.Update();
            info.Text = "...rapor hazırlandı.";

            #endregion

            //=========================================================================================================================
            //RFQ ADIYLA WORD DOSYASINI KAYDET
            //=========================================================================================================================
            wrd.Visible = false;
            doc.SaveAs2(SaveAs.SelectedPath + @"\" + rfq_1);
            button6.Enabled = true;
            





            //GEÇİCİ HATIRLATMA NOTLARI==========================================SİLİNECEK===========================> 11.07.2019 23:32
            //boyutları farklı olması gereken resimler için ayrıca kodlanmalı.
            //"pic_n.jpg" formatında rapor_otomasyon klasörü içerisine kaydedilen resimi yeniden boyutlandırıp
            //rapor_otomasyon\hazirlik içerisine aynı isimle kaydediyor.
            //rapor_otomasyon\hazirlik içine kaydedilen resmi raporda ilgili değişkenin yerine yapıştırıyor.
            //uzun sürebilmesi durumu için progress bar eklenecek.
            //kanal tanımlama görseli ve kalibrasyon sertifikası gibi farklı boyutlarda olabilecek resimler unutulmamalı. 
            //bu döngüden her resimaynı boyutta çıkar.
            //Excel grafiklerinden jpg kaydetme, header ve footer konusu Yunus'ta. ondan alacağım.

            //=======================================================================================================> 12.07.2019 13:31
            //Eğer resmin yüksekliği genişliğinden büyük ise farklı bir aspect ratio ayarlanabilir.
            //=========================================================================================================>===============


        }

        public void Button6_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(SaveAs.SelectedPath);
            

        }
    }
}


// toplam kütle hesaplarını gözden geçir. gerekirse beam seçenekleri ekle... kütle hesabı kolaylaşmalı...