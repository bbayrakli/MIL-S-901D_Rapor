﻿using System;
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
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            string ekipman_adi, musteri_adi, kanal_1_seri, kanal_2_seri, kanal_3_seri, kanal_4_seri, seri_no, hazirlayan_1, musteri_adres, rfq_1, test_tar,
                   ekipman_tar, rapor_tar, grade_1, type_1, class_1, en_01, boy_01, yukseklik_01;

            double ekipman_kutle, fikstur_kutle, toplam_kutle, efikstu_kutle, etopla_kutle,
                    dcy_1, dcy_2, dcy_3, ecy_1, ecy_2, ecy_3, dmy_1_m, dmy_2_m, dmy_3_m, emy_1_m, emy_2_m, emy_3_m;



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
            //etopla_kutle   = Eğik Fikstür Toplam Kütle       = efikstu_kutle + ekipman_kutle
            //epoun_1        = Ingiliz Ağırlık Birimi          = etopla_kutle * (2.2)


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
            toplam_kutle = ekipman_kutle + fikstur_kutle;
            double pound = toplam_kutle * 2.2;
            efikstu_kutle = Convert.ToDouble(efk_1.Text);
            etopla_kutle = efikstu_kutle + ekipman_kutle;
            double epoun_1 = etopla_kutle * 2.2;

            //=========================================================================================================================
            //901D TABLO İŞLEMLERİ
            //=========================================================================================================================

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


            if (epoun_1 < 1000)
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
                if (1000 <= epoun_1 && pound < 2000)
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
                    if (2000 <= epoun_1 && epoun_1 < 3000)
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
                        if (3000 <= epoun_1 && epoun_1 < 3500)
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
                            if (3500 <= epoun_1 && epoun_1 < 4000)
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
                                if (4000 <= epoun_1 && epoun_1 < 4200)
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
                                    if (4200 <= epoun_1 && epoun_1 < 4400)
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
                                        if (4400 <= epoun_1 && epoun_1 < 4600)
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
                                            if (4600 <= epoun_1 && epoun_1 < 4800)
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
                                                if (4800 <= epoun_1 && epoun_1 < 5000)
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
                                                    if (5000 <= epoun_1 && epoun_1 < 5200)
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
                                                        if (5200 <= epoun_1 && epoun_1 < 5400)
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
                                                            if (5400 <= epoun_1 && epoun_1 < 5600)
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
                                                                if (5600 <= epoun_1 && epoun_1 < 6200)

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
                                                                    if (6200 <= epoun_1 && epoun_1 < 6800)
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
                                                                        if (6800 <= epoun_1 && epoun_1 < 7400)
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

            //=========================================================================================================================
            //DEĞİŞKENLEİN DEĞERLERİ ATANDI
            //=========================================================================================================================

            Microsoft.Office.Interop.Word.Application wrd = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = wrd.Documents.Open(@"C:\users\bbayrakli\desktop\rapor_otomasyon\901D.docx");
            //=========================================================================================================================
            //WORD ÜZERİNDEKİ DEĞİŞKENLERİ DEĞİŞTİRME
            //=========================================================================================================================
            //Ekipman Adını Değiştir
            //====================================================================================================================
            var find = doc.Range().Find;
            find.Text = "<ekipman_adi>";
            find.Format = true;
            find.Replacement.Text = ekipman_adi;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Müşteri Adını Değiştir
            //====================================================================================================================
            find.Text = "<musteri_adi>";
            find.Format = true;
            find.Replacement.Text = musteri_adi;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Kanal 1 Seri No Değiştir
            //====================================================================================================================
            find.Text = "<kanal_1_seri>";
            find.Format = true;
            find.Replacement.Text = kanal_1_seri;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Kanal 2 Seri No Değiştir
            //====================================================================================================================
            find.Text = "<kanal_2_seri>";
            find.Format = true;
            find.Replacement.Text = kanal_2_seri;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Kanal 3 Seri No Değiştir
            //====================================================================================================================
            find.Text = "<kanal_3_seri>";
            find.Format = true;
            find.Replacement.Text = kanal_3_seri;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Kanal 4 Seri No Değiştir
            //====================================================================================================================
            find.Text = "<kanal_4_seri>";
            find.Format = true;
            find.Replacement.Text = kanal_4_seri;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Seri No Değiştir
            //====================================================================================================================
            find.Text = "<seri_no>";
            find.Format = true;
            find.Replacement.Text = seri_no;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Hazırlayanı Değiştir
            //====================================================================================================================
            find.Text = "<hazirlayan>";
            find.Format = true;
            find.Replacement.Text = hazirlayan_1;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Müşteri Adresi Değiştir
            //====================================================================================================================
            find.Text = "<musteri_adres>";
            find.Format = true;
            find.Replacement.Text = musteri_adres;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //RFQ No Değiştir
            //====================================================================================================================
            find.Text = "<rfq_1>";
            find.Format = true;
            find.Replacement.Text = rfq_1;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Test Tarihi Değiştir
            //====================================================================================================================
            find.Text = "<test_tar>";
            find.Format = true;
            find.Replacement.Text = test_tar;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Ekipman Tarihi Değiştir
            //====================================================================================================================
            find.Text = "<ekipman_tar>";
            find.Format = true;
            find.Replacement.Text = ekipman_tar;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Rapor Tarihi Değiştir
            //====================================================================================================================
            find.Text = "<rapor_tar>";
            find.Format = true;
            find.Replacement.Text = rapor_tar;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Grade Değiştir
            //====================================================================================================================
            find.Text = "<grade_1>";
            find.Format = true;
            find.Replacement.Text = grade_1;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Type Değiştir
            //====================================================================================================================
            find.Text = "<type_1>";
            find.Format = true;
            find.Replacement.Text = type_1;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Class Değiştir
            //====================================================================================================================
            find.Text = "<class_1>";
            find.Format = true;
            find.Replacement.Text = class_1;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //En Boy Yükseklik Değiştir
            //====================================================================================================================
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

            //====================================================================================================================
            //Ekipman Kütlesi Değiştir
            //====================================================================================================================
            find.Text = "<ekipman_kutle>";
            find.Format = true;
            find.Replacement.Text = ekipman_kutle.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Fikstür Kütlesi Değiştir
            //====================================================================================================================
            find.Text = "<fikstur_kutle>";
            find.Format = true;
            find.Replacement.Text = fikstur_kutle.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Toplam Kütle Değiştir
            //====================================================================================================================
            find.Text = "<toplam_kutle>";
            find.Format = true;
            find.Replacement.Text = toplam_kutle.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Pound Değeri Değiştir
            //====================================================================================================================
            find.Text = "<pound_1>";
            find.Format = true;
            find.Replacement.Text = pound.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Eğik Fikstür Kütlesi Değiştir
            //====================================================================================================================
            find.Text = "<efikstu_kutle>";
            find.Format = true;
            find.Replacement.Text = efikstu_kutle.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Eğik Fikstür Toplam Kütle Değiştir
            //====================================================================================================================
            find.Text = "<etopla_kutle>";
            find.Format = true;
            find.Replacement.Text = etopla_kutle.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Eğik Fikstür Pound Değeri Değiştir
            //====================================================================================================================
            find.Text = "<epoun_1>";
            find.Format = true;
            find.Replacement.Text = epoun_1.ToString();
            find.Execute(Replace: WdReplace.wdReplaceAll);

            //====================================================================================================================
            //Çekiç Yüksekliklerini Değiştir
            //====================================================================================================================
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

            Clipboard.SetImage(Image.FromFile(@"C:\users\bbayrakli\desktop\rapor_otomasyon\pic_01.jpg"));

            
            var sel = wrd.Selection;
            sel.Find.Text = string.Format("pic_01");
            sel.Find.Execute(Replace: WdReplace.wdReplaceNone);
            sel.Range.Select();


            //This code inserts the image
            var xxx = wrd.Selection;
            xxx.Paste();














            wrd.Visible = true;
            doc.SaveAs2(@"C:\users\bbayrakli\desktop\Raporlar\901D\" + rfq_1);










        }
    }
    }
