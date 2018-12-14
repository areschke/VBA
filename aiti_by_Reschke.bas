Attribute VB_Name = "aiti_by_Reschke"
' ************************************************************************************************************************************************************************************
' Editorial:
'
' Autor:                ai-ti: by Reschke | www.ai-ti.eu
' Version:              1.0 / 30.11.2018
' Haftung:              Für nicht durch den Autor vorgenommene oder durch ihn autorisierte Modifikationen
'                       wird keine Haftung übernommen. Manuelle Anpassungen geschehen dann auf eigenes Risiko.


' ************************************************************************************************************************************************************************************
' Struktur:             Dieses Exceldatei (.xlsm) besteht aus 6 Tabellenblätter (Sheets) und 1 Userform:
'
'           Sheets:     START       =   Eingangsmaske, welche direkt beim Aufruf der Datei angezeigt wird
'                       DATA_UPLOAD =   Speicherung der Daten, Basis der CSV-Datei für den Upload
'                       PARAM       =   Systemeinstellungen, wie z. B. der Speicherort der CSV-Datei
'                       DOKU        =   Dokumentation, Hilfe, Ablaufbeschreibung
'                       NOTIZEN     =   Speicherung Bearbeitungshinweis (ohne Übertrag an Mahnfabrik)
'                       KATALOG     =   Quelle für katalogbasierte Auswahlen (z. B. Anrede)


' ************************************************************************************************************************************************************************************
' Feldzuordnungen:      Beschreibung der Zuordnungen der Objekte in der Userform (= Datenerfassung) zu den
'                       den Spalten des Sheets "DATA_UPLOAD":
'
'           Spaltenname                 Feldinhalt                              Spaltennr.  Objekt          Pflicht Bemerkungen
'           --------------------------- --------------------------------------- ----------  --------------- ------- ------------------------------------------------------------------
'           A01_PartnerID               Mandantenummer bei SEPA Collect               1     tBx0            X       (wird aus Sheet 'PARAM' aus Spalte 6, Zeile 17 ermittelt)
'           A02_Produkt1                1=Produkt gebucht                             2     tBx96                   (wird über Modulbuchung eingestellt)
'           A03_Produkt2                1=Produkt gebucht                             3     tBx97                   (wird über Modulbuchung eingestellt)
'           A04_Produkt3                1=Produkt gebucht                             4     tBx98                   (wird über Modulbuchung eingestellt)
'           A05_Produkt4                1=Produkt gebucht                             5     tBx99                   (wird über Modulbuchung eingestellt)
'           A06_Produkt5                1=Produkt gebucht                             6     tBx100                  (wird über Modulbuchung eingestellt)
'           A07_Dummy1                  Reservefeld für Erweiterungen                 7     tBx101
'           A08_Dummy2                  Reservefeld für Erweiterungen                 8     tBx102
'           A09_Dummy3                  Reservefeld für Erweiterungen                 9     tBx103
'           A10_Dummy4                  Reservefeld für Erweiterungen                10     tBx104
'           A11_Dummy5                  Reservefeld für Erweiterungen                11     tBx105
'           F01 Mietvertragsnummer      Folgenummer aus Wowinex                      12     tBx7            X
'           F02_Belegnummer             OP-Nummer aus Wowinex                        13     tBx8            X
'           F03_Geschaeftsjahr          Geschäftsjahr aus Wertstellung des OPs       14     tBx106          X
'           F04_LfdNr                   Lfd. Nr. des Datensatzes zur Akte            15     tBx107          X       (wird von der Schnittstelle erzeugt)
'           F05_KatalogNr               siehe Sheet 'KATALOG'                        16                     X       ###
'           F06_AnsprVomDatum           Monatserster der OP-Wertstellung             17     tBx108          X
'           F07_AnsprBisDatum           Monatsletzter der OP-Wertstellung            18     tBx109         (X)
'           F08_AnsprGrundMM            siehe Sheet 'KATALOG'                        19                     X       ###
'           F09_VertragsDatum           Unterschrift Mietvertrag                     20     tBx110          X
'           F10_Vertragsnutzung         gewerblich / privat                          21     oPb1, oPb2      X       ###
'           F11_Verwaltungseinheit      WE aus Wowinex                               22     tBx3
'           F12_Etage                   Geschoss aus Wowinex                         23     tBx111
'           F13_Lage                    Wohnlage aus Wowinex                         24     tBx112
'           F14_ Wohnungsnummer         Wohnungs-Nr./Zus.-Nr. aus Wowinex            25     tBx5
'           F15_HFVz                    Hauptforderung                               26     tBx9
'           F16_HFVzWMM                 Währungsmerkmal Hauptforderung               27                             (leer = EUR)    ###
'           F17_HFValuta                Fälligstellung Hauptforderung                28                             (bei Dauerschuldverhältnis = Zinssatz Valuta)   ###
'           F18_MahnKostBetrag          Zu buchende Mahnkosten                       29     tBx10
'           F19_MahnKostBetragWMM       Währungsmerkmal Mahnkosten                   30                             (leer = EUR)    ###
'           F20_MahnKostValuta          Fälligstellung Mahnkosten                    31                             ###
'           F21_AuskunftKostBetrag      Zu buchende Auskunftskosten                  32     tBx11
'           F22_AuskunftKostBetragWM    Währungsmerkmal Auskunftskosten              33                             (leer = EUR)    ###
'           F23_AuskunftKostValuta      Fälligstellung Auskunftskosten               34                             ###
'           F24_BRLKostBetrag           Zu buchende Bankrücklastkosten               35     tBx12                   (in Userform 'RLS-Gebühr' für 'Rücklastschrift')
'           F25_BRLKostBetragWMM        Währungsmerkmal Bankrücklastkosten           36                             (leer = EUR)    ###
'           F26_BRLKostValuta           Fälligstellung Bankrücklastkosten            37                             ###
'           F27_SonstNFKostBetrag       Zu buchende sonstige Nebenforderungen        38     tBx13                   (z. B. Versandkosten)
'           F28_SonstNFKostBetragWMM    Währungsmerkmal sonst. Nebenforderungen      39                             (leer = EUR)    ###
'           F29_SonstNFKostValuta       Fälligstellung sonst. Nebenforderungen       40                             ###
'           F30_LeistungsStrasse        Strasse der Leistungsadresse m. Hausnr.      41     tBx114+tBx115  (X)     (Straße+HausNr; nur für die Katalognummern 19, 20 und 90 Pflicht)
'           F31_LeistungsPLZ            Postleitzahl der Leistungsadresse            42     tBx117         (X)
'           F32_LeistungsOrt            Ort der Leistungsadresse                     43     tBx118         (X)
'           F33_LeistungsNation         Nation der Leistungsadresse                  44     tBx116         (X)
'           F34_LeistungsZusatz         Zusatz zur Leistungsadresse                  45     tBx113
'           F35_ForderungsWMM           Forderungs-Währungsmerkmal                   46                             (leer = EUR)    ###
'           F36_ZinssatzValuta          Datum 1. Mahnung                             47                             (Leer bei Dauerschuldverhältnissen) ###
'           F37_Mandant                 Mandant aus Wowinex                          48     tBx1                    (bisher: F37_Dummy1 | Reservefeld für Erweiterungen)
'           F38_Unternehmen             Unternehmen aus Wowinex                      49     tBx2                    (bisher: F38_Dummy2 | Reservefeld für Erweiterungen)
'           F39_HausNr                  Hausnummer aus Wowinex                       50     tBx4                    (bisher: F39_Dummy3 | Reservefeld für Erweiterungen)
'           F40_WohnZusNr               Wohnungszusatznummer aus Wowinex             51     tBx6                    (bisher: F40_Dummy4 | Reservefeld für Erweiterungen)
'           F41_Dummy5                  Reservefeld für Erweiterungen                52     tBx69
'           F42_Dummy6                  Reservefeld für Erweiterungen                53     tBx70
'           F43_Dummy7                  Reservefeld für Erweiterungen                54     tBx71
'           F44_Dummy8                  Reservefeld für Erweiterungen                55     tBx72
'           F45_Dummy9                  Reservefeld für Erweiterungen                56     tBx73
'           F46_Dummy10                 Reservefeld für Erweiterungen                57     tBx74
'           M1_01_ExternNr              Adressnr aus Wowinex                         58     tBx19           X       (bei Genossenschaften: AdressNr = Mitgliedsnummer)
'           M1_02_Anrede                siehe Sheet 'KATALOG'                        59     tBx15           X
'           M1_03_Name1                 Nachname oder 1. Firmenbezeichnung           60     tBx16           X       (nicht abschneiden falls die Feldlänge überschritten wird)
'           M1_04_Name2                 Vorname oder 2.Firmenbezeichnung             61     tBx17           X       (nicht abschneiden falls die Feldlänge überschritten wird)
'           M1_05_GeburtsDatum          Geburtsdatum                                 62     tBx119          X       (nur bei natürlichen Personen)
'           M1_06_TodesDatum            Todesdatum                                   63     tBx120          X       (nur bei natürlichen Personen)
'           M1_07_ZBStrasse             Straße Zustellanschrift Schuldner            64     tBx20           X
'           M1_08_ZBHausNr              Hausnummer Zustellanschrift Schuldner        65     tBx66           X
'           M1_09_ZBPLZ                 PLZ Zustellanschrift Schuldner               66     tBx22           X
'           M1_10_ZBOrt                 Ort Zustellanschrift Schuldner               67     tBx23           X
'           M1_11_ZBOrtsTeil            Ortsteil Zustellanschrift Schuldner          68     tBx24
'           M1_12_ZBNation              Landeskennzeichen Zustellanschrift Sch.      69     tBx21
'           M1_13_ZBZusatz              Zusatz Zustellanschrift Schuldner            70     tBx18                   (z. B. 'c/o')
'           M1_14_UnbekanntVerzogen     Merkmal unbekannt verzogen                   71     oPb3, Opb4              (1=ja, 0 oder leer=nein)
'           M1_15_Email                 Email des Schuldners                         72     tBx25
'           M1_16_Mobil                 Mobilnummer des Schuldners                   73     tBx26
'           M1_17_TelefonNr             Telefonnummer des Schuldners                 74     tBx27
'           M1_18_TelefaxNr             Telefaxnummer des Schuldners                 75     tBx28
'           M1_19_IBAN                  IBAN des Schuldners                          76     tBx29
'           M1_20_BIC                   BIC des Schuldners                           77     tBx30
'           M1_21_VermerkWeitereIBANs   weitere IBANs des Schuldners                 78     tBx31                   (falls bekannt, hier eintragen; Trennzeichen = Komma ",")
'           M1_22_Dummy1                Reservefeld für Erweiterungen                79     tBx75
'           M1_23_Dummy2                Reservefeld für Erweiterungen                80     tBx76
'           M1_24_Dummy3                Reservefeld für Erweiterungen                81     tBx77
'           M1_25_Dummy4                Reservefeld für Erweiterungen                82     tBx78
'           M1_26_Dummy5                Reservefeld für Erweiterungen                83     tBx79
'           M1_27_Dummy6                Reservefeld für Erweiterungen                84     tBx80
'           M1_28_Dummy7                Reservefeld für Erweiterungen                85     tBx81
'           M2_01_ExternNr              Adressnr aus Wowinex                         86     tBx121        (X)       (sofern 2. Mieter vorhanden)
'           M2_02_Anrede                siehe Sheet 'KATALOG'                        87     tBx34         (X)
'           M2_03_Name1                 Nachname oder 1. Firmenbezeichnung           88     tBx35         (X)
'           M2_04_Name2                 Vorname oder 2.Firmenbezeichnung             89     tBx36         (X)
'           M2_05_GeburtsDatum          Geburtsdatum                                 90     tBx122        (X)
'           M2_06_TodesDatum            Todesdatum                                   91     tBx123        (X)
'           M2_07_ZBStrasse             Straße Zustellanschrift Schuldner            92     tBx38         (X)
'           M2_08_ZBHausNr              Hausnummer Zustellanschrift Schuldner        93     tBx67         (X)
'           M2_09_ZBPLZ                 PLZ Zustellanschrift Schuldner               94     tBx40         (X)
'           M2_10_ZBOrt                 Ort Zustellanschrift Schuldner               95     tBx41         (X)
'           M2_11_ZBOrtsTeil            Ortsteil Zustellanschrift Schuldner          96     tBx42
'           M2_12_ZBNation              Landeskennzeichen Zustellanschrift Sch.      97     tBx39
'           M2_13_ZBZusatz              Zusatz Zustellanschrift Schuldner            98     tBx37
'           M2_14_UnbekanntVerzogen     Merkmal unbekannt verzogen                   99     oPb5, oPb6
'           M2_15_Email                 Email des Schuldners                        100     tBx43
'           M2_16_Mobil                 Mobilnummer des Schuldners                  101     tBx44
'           M2_17_TelefonNr             Telefonnummer des Schuldners                102     tBx45
'           M2_18_TelefaxNr             Telefaxnummer des Schuldners                103     tBx46
'           M2_19_IBAN                  IBAN des Schuldners                         104     tBx47
'           M2_20_BIC                   BIC des Schuldners                          105     tBx48
'           M2_21_VermerkWeitereIBANs   weitere IBANs des Schuldners                106     tBx49
'           M2_22_Dummy1                Reservefeld für Erweiterungen               107     tBx82
'           M2_23_Dummy2                Reservefeld für Erweiterungen               108     tBx83
'           M2_24_Dummy3                Reservefeld für Erweiterungen               109     tBx84
'           M2_25_Dummy4                Reservefeld für Erweiterungen               110     tBx85
'           M2_26_Dummy5                Reservefeld für Erweiterungen               111     tBx86
'           M2_27_Dummy6                Reservefeld für Erweiterungen               112     tBx87
'           M2_28_Dummy7                Reservefeld für Erweiterungen               113     tBx88
'           M3_01_ExternNr              Adressnr aus Wowinex                        114     tBx124        (X)       (sofern 3. Mieter vorhanden)
'           M3_02_Anrede                siehe Sheet 'KATALOG'                       115     tBx50         (X)
'           M3_03_Name1                 Nachname oder 1. Firmenbezeichnung          116     tBx52         (X)
'           M3_04_Name2                 Vorname oder 2.Firmenbezeichnung            117     tBx53         (X)
'           M3_05_GeburtsDatum          Geburtsdatum                                118     tBx125        (X)
'           M3_06_TodesDatum            Todesdatum                                  119     tBx126        (X)
'           M3_07_ZBStrasse             Straße Zustellanschrift Schuldner           120     tBx54         (X)
'           M3_08_ZBHausNr              Hausnummer Zustellanschrift Schuldner       121     tBx68         (X)
'           M3_09_ZBPLZ                 PLZ Zustellanschrift Schuldner              122     tBx56         (X)
'           M3_10_ZBOrt                 Ort Zustellanschrift Schuldner              123     tBx57         (X)
'           M3_11_ZBOrtsTeil            Ortsteil Zustellanschrift Schuldner         124     tBx58
'           M3_12_ZBNation              Landeskennzeichen Zustellanschrift Sch.     125     tBx55
'           M3_13_ZBZusatz              Zusatz Zustellanschrift Schuldner           126     tBx51
'           M3_14_UnbekanntVerzogen     Merkmal unbekannt verzogen                  127     oPb7, oPb8
'           M3_15_Email                 Email des Schuldners                        128     tBx59
'           M3_16_Mobil                 Mobilnummer des Schuldners                  129     tBx60
'           M3_17_TelefonNr             Telefonnummer des Schuldners                130     tBx61
'           M3_18_TelefaxNr             Telefaxnummer des Schuldners                131     tBx62
'           M3_19_IBAN                  IBAN des Schuldners                         132     tBx63
'           M3_20_BIC                   BIC des Schuldners                          133     tBx64
'           M3_21_VermerkWeitereIBANs   weitere IBANs des Schuldners                134     tBx65
'           M3_22_Dummy1                Reservefeld für Erweiterungen               135     tBx89
'           M3_23_Dummy2                Reservefeld für Erweiterungen               136     tBx90
'           M3_24_Dummy3                Reservefeld für Erweiterungen               137     tBx91
'           M3_25_Dummy4                Reservefeld für Erweiterungen               138     tBx92
'           M3_26_Dummy5                Reservefeld für Erweiterungen               139     tBx93
'           M3_27_Dummy6                Reservefeld für Erweiterungen               140     tBx94
'           M3_28_Dummy7                Reservefeld für Erweiterungen               141     tBx95


' ************************************************************************************************************************************************************************************
' Eineindeutiger Key:   Bildung durch Zusammensetzen der nachfolgenden Spalten
'                       Da es mehrere Datensätze zu einem Mieter geben kann, ist die Eineindeutigkeit nur über die 'lfd. Nummer' in Spalte 15 darstellbar.
'
'           Spaltenname     F37_Mandant . F38_Unternehmen . F11_Verwaltungseinheit . F39_HausNr . F14_Wohnungsnummer . F40_WohnNrZus . F01_Mietvertragsnummer . F04_lfdNr
'           --------------  -----------   ---------------   ----------------------   ----------   ------------------   -------------   ----------------------   ---------
'           Spaltennummer   48            49                22                       50           25                   51              12                       15
'           Objekt          tBx1          tBx2              tBx3                     tBx4         tBx5                 tBx6            tBx7


' ************************************************************************************************************************************************************************************
' Passwortschutz
'
'           Es existieren zwei Passwörter:
'
'           a) zum Öffnen des VBA-Quellcodes:   10070001
'           b) zum Entfernen des Blattschutzes: mahnfabrik10070001




