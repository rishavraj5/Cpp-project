#include<bits/stdc++.h>
using namespace std;

//Attribute VB_Name = "mod"
/*================================================================
' Description..: Krishnamurti [KP] Astrology Software
' Software.....: KP New Astro
' Date.........: 05/12/2010
' Version......: 1.0.xx beta
' Language.....: Visual Basic 6.0 Ent - SP 6
' Tested.......: Windows XP Professional - SP 3
' Copyright....: [C] 2009-2010 JSW
' E-Mail.......: kpnewastro@gmail.com
' Web..........: http://www.kpnewastro.blogspot.com
'================================================================
' Released under GNU general public license [version 2 or later]
'================================================================
*/
//Option Explicit

//Attribute VB_Name = "mod"
/*================================================================
// Description..: Krishnamurti [KP] Astrology Software
// Software.....: KP New Astro
// Date.........: 05/12/2010
// Version......: 1.0.xx beta
// Language.....: Visual Basic 6.0 Ent - SP 6
// Tested.......: Windows XP Professional - SP 3
// Copyright....: [C] 2009-2010 JSW
// E-Mail.......: kpnewastro@gmail.com
// Web..........: http://www.kpnewastro.blogspot.com
//================================================================
// Released under GNU general public license [version 2 or later]
//================================================================
*/
//Option Explicit

string rashiNameTr[11]; //Traditional Rashi Names
string rashiNameTrSh[11]; //Traditional Rashi Names [Short]
string rashiLordNameTr[11] ; //Traditional Rashilord Names
string planetName1Tr[11]; //Traditional Planet Names
string planetName1TrR[11]; //Traditional Planet Names [R]
string planetName2Tr[8]; //Traditional Planet Names [Short]
string planetName2TrR[8]; //Traditional Planet Names [R] [Short]
string starName[26]; //Star Names
string weekDayName[6] ; //Week Day Names
string weekDayPlanetNameTr[6]; //Week Day to Planet Name
string marakaAdhipathiTr[11]; //Maraka Adhipathi [Traditional]
string bhadhakaAdhipathiTr[11]; //Bhadhaka Adhipathi [Traditional]
int planet2Int[8]; //Planet int to Dasa Type Order
int rashiLord1Int[11]; //Rashi Lord Interger [Import From Setting file]
int rashiLord2Int[11]; //Rashi Lord Interger [Import From Setting file]
int starLord1Int[26]; //Star Loard Integer
int weekDayInt[6]; //WeekDay Integer
int weeekDayPlanet1Int[6]; //WeekDay to Planet Integer
int weeekDayPlanet2Int[6]; //WeekDay to Planet Integer
double dasaYears[8]; //Dasa Years
double aspectPont[17]; //Aspect Point [Exact values]
string aspNature[17]; //Nature of the aspect
string aspectName[17]; //Name of the aspect
string arToRom[11]; //Arabic to Roman numbers

//Sub assign[]

// rashiNameTr[11] As String
rashiNameTr[0] = "Mesha / Aries ";
rashiNameTr[1] = "Rishaba / Taurus ";
rashiNameTr[2] = "Mithuna / Gemini ";
rashiNameTr[3] = "Kataka / Cancer ";
rashiNameTr[4] = "Simha / Leo ";
rashiNameTr[5] = "Kanya / Virgo ";
rashiNameTr[6] = "Thula / Libra ";
rashiNameTr[7] = "Vrishchika / Scorpio";
rashiNameTr[8] = "Dhanu / Sagittarius ";
rashiNameTr[9] = "Makara / Capricorn ";
rashiNameTr[10] = "Kumbha / Aquarius ";
rashiNameTr[11] = "Meena / Pisces ";

// rashiNameTrSh[11] As String
rashiNameTrSh[0] = "Mes/Ari";
rashiNameTrSh[1] = "Ris/Tau";
rashiNameTrSh[2] = "Mit/Gem";
rashiNameTrSh[3] = "Kat/Can";
rashiNameTrSh[4] = "Sim/Leo";
rashiNameTrSh[5] = "Kan/Vir";
rashiNameTrSh[6] = "Thu/Lib";
rashiNameTrSh[7] = "Vri/Sco";
rashiNameTrSh[8] = "Dha/Sag";
rashiNameTrSh[9] = "Mak/Cap";
rashiNameTrSh[10] = "Kum/Aqu";
rashiNameTrSh[11] = "Mee/Pis";

//load planet names from file
string loadedPlanetNames;
string plN[18];
loadedPlanetNames[] = hp_ReadDatFile[App.Path + "\Sett\nmedet.dat", "_"];
for(int i1=0;i1<18;i1++)
plN[i1] = loadedPlanetNames[i1];

// rashiLordNameTr[11] As String
rashiLordNameTr[0] = "Kuj";
rashiLordNameTr[0] = "Kuj";
rashiLordNameTr[1] = "Suk";
rashiLordNameTr[2] = "Bud";
rashiLordNameTr[3] = "Cha";
rashiLordNameTr[4] = "Rav";
rashiLordNameTr[5] = "Bud";
rashiLordNameTr[6] = "Suk";
rashiLordNameTr[7] = "Kuj";
rashiLordNameTr[8] = "Gur";
rashiLordNameTr[9] = "San";
rashiLordNameTr[10] = "San";
rashiLordNameTr[11] = "Gur";

for(int i2=0;i2<11;i2++)
planetName1Tr[i2] = plN[i2];

// planetName1TrR[11] As String
planetName1TrR[0] = plN[0]; //"Rav"
planetName1TrR[1] = plN[1] ;//"Cha"
planetName1TrR[2] = plN[12] ;// "Ku�"
planetName1TrR[3] = plN[13]; //"Bu�"
planetName1TrR[4] = plN[14]; //"Gu�"
planetName1TrR[5] = plN[15]; //"Su�"
planetName1TrR[6] = plN[16]; //"Sa�"
planetName1TrR[7] = plN[7]; //"Rah"
planetName1TrR[8] = plN[8]; //"Ket"
planetName1TrR[9] = plN[17]; // "Ur�"
planetName1TrR[10] = plN[18]; //"Ne�"
planetName1TrR[11] = plN[11]; //"For"

// planetName2Tr[8] As String
planetName2Tr[0] = plN[8]; //"Ket"
planetName2Tr[1] = plN[5]; //"Suk"////
planetName2Tr[2] = plN[0]; //"Rav"
planetName2Tr[3] = plN[1]; //"Cha"
planetName2Tr[4] = plN[2]; //"Kuj"////
planetName2Tr[5] = plN[7]; //"Rah"
planetName2Tr[6] = plN[4]; //"Gur"////
planetName2Tr[7] = plN[6]; //"San"////
planetName2Tr[8] = plN[3]; //"Bud"////

// planetName2TrR[8] As String
planetName2TrR[0] = plN[8]; //"Ket"
planetName2TrR[1] = plN[15]; //"Su�"
planetName2TrR[2] = plN[0]; //"Rav"
planetName2TrR[3] = plN[1]; //"Cha"
planetName2TrR[4] = plN[12]; //"Ku�"
planetName2TrR[5] = plN[7]; //"Rah"
planetName2TrR[6] = plN[14]; //"Gu�"
planetName2TrR[7] = plN[16]; //"Sa�"
planetName2TrR[8] = plN[13]; //"Bu�"

// starName[26] As String
starName[0] = "Asvida / Aswini";
starName[1] = "Berana / Bharani";
starName[2] = "Kethi / Krithika";
starName[3] = "Rehena / Rohini";
starName[4] = "Muvasirasa / Mrigasira";
starName[5] = "Ada / Arudhra";
starName[6] = "Punavasa / Punarvasu";
starName[7] = "Pusha / Pushyam";
starName[8] = "Aslisa / Ashlesha";

starName[9] = "Ma / Magha";
starName[10] = "Puvapal / Purva Phalgu";
starName[11] = "Uthrapal / Uttara Phalgu";
starName[12] = "Hatha / Hasta";
starName[13] = "Sitha / Chitra";
starName[14] = "Sa / Swati";
starName[15] = "Visa / Vishakha";
starName[16] = "Anura / Anuradha";
starName[17] = "Deta / Jyeshtha";

starName[18] = "Mula / Moola";
starName[19] = "Puvasala / Purva Ashad";
starName[20] = "Uthrasala / Uttara Ashad";
starName[21] = "Suvana / Shravan";
starName[22] = "Denata / Dhanistha";
starName[23] = "Siyavasa / Satabishak";
starName[24] = "Puvaputupa / Purva Bhadra";
starName[25] = "Uthraputupa / Uttara Bhadr";
starName[26] = "Revathi / Revati";

// weekDayName[6] As String
weekDayName[0] = "Monday";
weekDayName[1] = "Tuesday";
weekDayName[2] = "Wednesday";
weekDayName[3] = "Thursday";
weekDayName[4] = "Friday";
weekDayName[5] = "Saturday";
weekDayName[6] = "Sunday";

// weekDayPlanetNameTr[6] As String
weekDayPlanetNameTr[0] = plN[1]; //"Chandra"
weekDayPlanetNameTr[1] = plN[2]; //"Kuja"
weekDayPlanetNameTr[2] = plN[3]; //"Budha"
weekDayPlanetNameTr[3] = plN[4]; //"Guru"
weekDayPlanetNameTr[4] = plN[5]; //"Sukra"
weekDayPlanetNameTr[5] = plN[6]; //"Sani"
weekDayPlanetNameTr[6] = plN[0]; //"Ravi"

// marakaAdhipathiTr[11] As String
marakaAdhipathiTr[0] = "Sukra, Sukra";
marakaAdhipathiTr[1] = "Budha, Kuja";
marakaAdhipathiTr[2] = "Chandra, Guru";
marakaAdhipathiTr[3] = "Ravi, Sani";
marakaAdhipathiTr[4] = "Budha, Sani";
marakaAdhipathiTr[5] = "Sukra, Guru";
marakaAdhipathiTr[6] = "Kuja, Kuja";
marakaAdhipathiTr[7] = "Guru, Sani";
marakaAdhipathiTr[8] = "Sani, Budha";
marakaAdhipathiTr[9] = "Sani, Chandra";
marakaAdhipathiTr[10] = "Guru, Ravi";
marakaAdhipathiTr[11] = "Kuja, Budha";

// bhadhakaAdhipathiTr[11] As String
bhadhakaAdhipathiTr[0] = "Sani";
bhadhakaAdhipathiTr[1] = "Sani";
bhadhakaAdhipathiTr[2] = "Guru";
bhadhakaAdhipathiTr[3] = "Sukra";
bhadhakaAdhipathiTr[4] = "Kuja";
bhadhakaAdhipathiTr[5] = "Guru";
bhadhakaAdhipathiTr[6] = "Ravi";
bhadhakaAdhipathiTr[7] = "Chandra";
bhadhakaAdhipathiTr[8] = "Budha";
bhadhakaAdhipathiTr[9] = "Kuja";
bhadhakaAdhipathiTr[10] = "Sukra";
bhadhakaAdhipathiTr[11] = "Budha";

// planet2Int[8] as Integer
planet2Int[0] = 8;
planet2Int[1] = 5;
planet2Int[2] = 0;
planet2Int[3] = 1;
planet2Int[4] = 2;
planet2Int[5] = 7;
planet2Int[6] = 4;
planet2Int[7] = 6;
planet2Int[8] = 3;

// rashiLord1Int[11] As Byte
rashiLord1Int[0] = 2;
rashiLord1Int[1] = 5;
rashiLord1Int[2] = 3;
rashiLord1Int[3] = 1;
rashiLord1Int[4] = 0;
rashiLord1Int[5] = 3;
rashiLord1Int[6] = 5;
rashiLord1Int[7] = 2;
rashiLord1Int[8] = 4;
rashiLord1Int[9] = 6;
rashiLord1Int[10] = 6;
rashiLord1Int[11] = 4;

// rashiLord2Int[11] As Byte
rashiLord2Int[0] = 4;
rashiLord2Int[1] = 1;
rashiLord2Int[2] = 8;
rashiLord2Int[3] = 3;
rashiLord2Int[4] = 2;
rashiLord2Int[5] = 8;
rashiLord2Int[6] = 1;
rashiLord2Int[7] = 4;
rashiLord2Int[8] = 6;
rashiLord2Int[9] = 7;
rashiLord2Int[10] = 7;
rashiLord2Int[11] = 6;

// starLord1Int[8] As Byte
starLord1Int[0] = 8;
starLord1Int[1] = 5;
starLord1Int[2] = 0;
starLord1Int[3] = 1;
starLord1Int[4] = 2;
starLord1Int[5] = 7;
starLord1Int[6] = 4;
starLord1Int[7] = 6;
starLord1Int[8] = 3;

// weekDayInt[6] As Byte
weekDayInt[0] = 1;
weekDayInt[1] = 2;
weekDayInt[2] = 3;
weekDayInt[3] = 4;
weekDayInt[4] = 5;
weekDayInt[5] = 6;
weekDayInt[6] = 0;

// weeekDayPlanet1Int[6] As Byte
weeekDayPlanet1Int[0] = 1;
weeekDayPlanet1Int[1] = 2;
weeekDayPlanet1Int[2] = 3;
weeekDayPlanet1Int[3] = 4;
weeekDayPlanet1Int[4] = 5;
weeekDayPlanet1Int[5] = 6;
weeekDayPlanet1Int[6] = 0;

// weeekDayPlanet2Int[6] As Byte
weeekDayPlanet2Int[0] = 3;
weeekDayPlanet2Int[1] = 4;
weeekDayPlanet2Int[2] = 8;
weeekDayPlanet2Int[3] = 6;
weeekDayPlanet2Int[4] = 1;
weeekDayPlanet2Int[5] = 7;
weeekDayPlanet2Int[6] = 2;
;
// dasaYears[8] As Double
dasaYears[0] = 7#;
dasaYears[1] = 20#;
dasaYears[2] = 6#;
dasaYears[3] = 10#;
dasaYears[4] = 7#;
dasaYears[5] = 18#;
dasaYears[6] = 16#;
dasaYears[7] = 19#;
dasaYears[8] = 17#;

//Private aspectPont[17] As Single
aspectPont[0] = 0#;
aspectPont[1] = 180#;
aspectPont[2] = 120#;
aspectPont[3] = 150#;
aspectPont[4] = 90#;
aspectPont[5] = 60#;
aspectPont[6] = 144#;
aspectPont[7] = 135#;
aspectPont[8] = 72#;

aspectPont[9] = 45#;
aspectPont[10] = 30#;
aspectPont[11] = 18#;
aspectPont[12] = 24#;
aspectPont[13] = 36#;
aspectPont[14] = 108#;
aspectPont[15] = 54#;
aspectPont[16] = 162#;
aspectPont[17] = 126#;

//Private aspNature[17] As Byte
aspNature[0] = "�";
aspNature[1] = "�";
aspNature[2] = " ";
aspNature[3] = "�";
aspNature[4] = "�";
aspNature[5] = " ";
aspNature[6] = " ";
aspNature[7] = "�";

aspNature[8] = " ";
aspNature[9] = "�";
aspNature[10] = " ";
aspNature[11] = " ";
aspNature[12] = " ";
aspNature[13] = " ";
aspNature[14] = " ";
aspNature[15] = " ";

aspNature[16] = " ";
aspNature[17] = " ";

//Private aspectName[17] As String
aspectName[0] = "Conjc";
aspectName[1] = "Oppos";
aspectName[2] = "Trine";
aspectName[3] = "Quinc";
aspectName[4] = "Squar";
aspectName[5] = "Sexti";
aspectName[6] = "Biqui";
aspectName[7] = "Sesqu";

aspectName[8] = "Quint";
aspectName[9] = "S.Sqr";
aspectName[10] = "S.Sex";
aspectName[11] = "Vigin";
aspectName[12] = "Q.Des";
aspectName[13] = "D.S.Q";
aspectName[14] = "Trede";
aspectName[15] = "54Deg";

aspectName[16] = "162De";
aspectName[17] = "126De";

// arToRom[11] As String
arToRom[0] = "I ";
arToRom[1] = "II ";
arToRom[2] = "III ";
arToRom[3] = "IV ";

arToRom[4] = "V ";
arToRom[5] = "VI ";
arToRom[6] = "VII ";
arToRom[7] = "VIII ";

arToRom[8] = "IX ";
arToRom[9] = "X ";
arToRom[10] = "XI ";
arToRom[11] = "XII ";









/*==================================================
====================================================*/


//Attribute VB_Name = "modFunc"
/*================================================================
' Description..: Krishnamurti (KP) Astrology Software
' Software.....: KP New Astro
' Date.........: 05/12/2010
' Version......: 1.0.xx beta
' Language.....: Visual Basic 6.0 Ent - SP 6
' Tested.......: Windows XP Professional - SP 3
' Copyright....: (C) 2009-2010 JSW
' E-Mail.......: kpnewastro@gmail.com
' Web..........: http://www.kpnewastro.blogspot.com
'================================================================
' Released under GNU general public license (version 2 or later)
'================================================================
*/
//Option Explicit

double kp_GeoCorr(double geoCentricLat){
    //Geocentric correction
    double geoGRad;
    geoGRad = geoCentricLat * 1.74532925199433E-02 ;
    return (Atn(Tan(geoGRad) * 0.99330546)) * 57.2957795130823 ;
}

int kp_RashiInt(double posVal){
    //Returns Rashi as a integer value
    //Aries=0, Taurus=1, Gemini=2,..., Pisces=11

    double midVal[12];

    for(int i=0;i<12;i++)
        midVal[i] = 360# * (CDbl(i) / 12#);

    for(int j=0;j<11;j++)
    {
        if(midval[j]<=posVal && posVal<mid[j+1])
        {
            return j;
        }
    }
}

int kp_RashiLordInt(ByVal posVal As Double, bool isDasa){
    //Retuns rashi lord as a integer value
//Ravi=0, Chandra=1, Kuja=2, ... Fortune=11  -> isDasa=False
//Ketu=0, Sukra=1, Ravi=2, ... Budha=8        -> isDasa=True

    double midVal[12];

    for(int i=0;i<12;i++)
    {
        midVal[i] = 360# * (CDbl(i) / 12#);
    }

    if(isDasa == False){
        for(int j=0;j<11;j++)
        {
            if(midVal[j] <= posVal && posVal < midVal[j + 1])
            {
                return rashiLord1Int(j);
            }
        }
    }
    else{
        for(int j=0;j<11;j++){
            if(midVal[j] <= posVal && posVal < midVal[j + 1]){
                return rashiLord2Int(j);
            }
        }
    }    
}

int kp_StarLordInt(double posVal, bool isDasa){
//Returns Star lord as a integer value
//Ravi=0, Chandra=1, Kuja=2, ... Fortune=11  -> isDasa=False
//Ketu=0, Sukra=1, Sun=2, ... Budha=8        -> isDasa=True

    double midVal[27];

    for(int i=0;i<27;i++){
        midVal[i] = 360# * (CDbl(i) / 27#);
    }

    if(isDasa == false){
        for(int j=0;j<26;j++){
            if(midVal[j] <= posVal And posVal < midVal[j + 1]){
                return starLord1Int(hp_Rnd0To8v(j));
            }
        }
    }
    else{
        for(int j=0;j<26;j++){
            If midVal(j) <= posVal And posVal < midVal(j + 1){
                return hp_Rnd0To8v(j);
            }
        }    
    }
}

int  kp_StarInt(double posVal){
    //Returns Star as a integer value
    //Aswini=0, Bharani=1, Krithika=2, ..., Revati=26

    double midVal[27];

    for(int i=0;i<27;i++){
        midVal[i] = 360# * (CDbl(i) / 27#);
    }

    for(int j=0;j<26;j++){
        if(midVal[j] <= posVal And posVal < midVal[j + 1] Then
            return j;
    }
}

int kp_StarPada(double posVal As Double){
//Returns Star Pada as a integer valu....1 to 4

    double midVal[108];

    for(int i=0;i<108;i++)
        midVal[i] = 360# * (CDbl(i) / 108#);

    For j = 0 To 107
    for(int j=0;j<107;j++){
        If midVal(j) <= posVal And posVal < midVal(j + 1) Then
            return hp_Rnd1To4v(j + 1);
    }
}


string kp_PlanetName(int planetInt,double speedVal){
    //Return plnet name as a string
    //Ravi=0, Chandra=1, Kuja=3, ... ,Fortune=11
    //Restrograde planets with character " � "

    kp_PlanetName = vbNullString

    for(int i=0;i<11;i++){
        if(planetInt == i && speedVal >= 0){
            return planetName1Tr(i);
        }
        else if(planetInt == i And speedVal < 0){
            return planetName1TrR(i);
        }
    }
}


int kp_SubLord(double posVal,bool isSubSub){
    //Returns Sub lord as a integer value
    //Ketu=0, Venus=1, Sun=2, ... Mercury=8
    //Up to 2 sub levels

    double nakBorders[27];

    for(int i1=0;i1<27;i1++)
        nakBorders[i1] = 360# * (CDbl(i1) / 27#);

    double selVal1;

    for(int i2=0;i2<26;i2++){
        if(nakBorders[i2] <= posVal And posVal < nakBorders[i2 + 1]){
            selVal1 = nakBorders[i2];
        }   
    }

    int selNakLord;
    selNakLord = kp_StarLordInt(posVal, True);

    double sub1LordDur[8];
    int sub1LordInt[8];

    for(int j1=0;j1<8;j1++){
        sub1LordDur[j1] = 13.3333333333333 * (dasaYears(hp_Rnd0To8v(selNakLord + j1)) / 120#);
    }

    for(int j2=0;j2<8;j2++){
        sub1LordInt[j2] = hp_Rnd0To8v(selNakLord + j2);
    }

    double sub1Bor[9];

    sub1Bor(0) = selVal1
    for(int j3=0;j3<8;j3++){
        sub1Bor[j3 + 1] = sub1Bor[j3] + sub1LordDur[j3];
    }

    double selVal2;
    double sub1Duration;
    int sub1Lord;

    for(int i3=0;i3<8;i3++){
        if(sub1Bor[i3] <= posVal And posVal < sub1Bor[i3 + 1]{
            selVal2 = sub1Bor[i3]
            sub1Duration = sub1Bor[i3 + 1] - sub1Bor[i3]
            sub1Lord = sub1LordInt[i3];
        }
    }

    double sub2LordDur(8);
    int sub2LordInt[8];

    for(int j4=0;j4<8;j4++){
        sub2LordDur[j4] = sub1Duration * (dasaYears(hp_Rnd0To8v(sub1Lord + j4)) / 120#);
    }

    for(int j5=0;j5<8;j5++){
        sub2LordInt(j5) = hp_Rnd0To8v(sub1Lord + j5);
    }

    double sub2Bor[9];

    sub2Bor[0] = selVal2;
    for(int j6=0;j6<8;j6++){
        sub2Bor[j6 + 1] = sub2Bor[j6] + sub2LordDur[j6];
    }

    double selVal3;
    double sub2Duration;
    int sub2Lord;

    for(int i4=0;i4<8;i4++){
        if(sub2Bor[i4] <= posVal And posVal < sub2Bor[i4 + 1]){
            selVal3 = sub2Bor[i4];
            sub2Duration = sub2Bor[i4 + 1] - sub2Bor[i4];
            sub2Lord = sub2LordInt[i4];
        }
            
    }

    if(isSubSub == false)
        return sub1Lord;
    else
        return sub2Lord;
}

string kp_DasaShesha(double moonPos){
    //Returns dasa sheshaya as a string

    double nakBorder[27];

    for(int i1=0;i1<27;i1++){
        nakBorder[i1] = 360# * (CDbl(i1) / 27#);
    }

    double selNak;

    For i2 = 0 To 26
    for(int i2=0;i2<26;i2++){
        if((nakBorder[i2] <= moonPos) && (moonPos < nakBorder[i2 + 1])){
            selNak = nakBorder[i2];
        }
    }

    int nakLordInt;
    nakLordInt = kp_StarLordInt(moonPos, True);

    double sheshaVal;
    //    Dim virtualVal As Double
    //    virtualVal = moonPos - selNak
    sheshaVal = 13.3333333333333 - (moonPos - selNak);

    double seshaYrs;

    seshaYrs = (sheshaVal * dasaYears(nakLordInt)) / 13.3333333333333;

    string s1=planetName2Tr(nakLordInt);
    string s2=" Maha Dasa-";
    string s3=hp_FormalDate(seshaYrs);
    return s1+s2+s3;
}

float kp_WAspect(float pos1Val,float pos2Val,float aspectPoint,float orbValApp,
                    float orbValSep,bool isSelected){
    //No Aspect ---> Return > 360.0

    float x = 400#;

    if(isSelected == True){

        float aspDiff;

        aspDiff = abs(pos1Val - pos2Val)

        if(aspDiff > 180#){
            aspDiff = 360# - aspDiff;
        }
        If ((aspectPoint - orbValApp) <= aspDiff And aspDiff <= (aspectPoint + orbValSep)) Then
            return aspDiff;
    }
    return x;
}

Function kp_WAspFilter(float aspVal,bool isPDF){
    //Format the western Aspect to printable
    string s1 = " - ";
    string tmpStr;
    int lenStr;

    if(aspVal < 360#){

        //kp_WAspFilter = Format(aspVal, "000")
        tmpStr = str(int(aspVal));
        lenStr = tmpStr.length();
        if(isPDF == false){
            if(lenStr == 1){
                s1 = "  " + tmpStr;
            }
            if(lenStr == 2){
                s1 = " " + tmpStr;
            }
            if(lenStr == 3){
                s1 = tmpStr;
            }
        else
            s1 = tmpStr 
        }
    }
    return s1;
}


/*
string kp_RevJday(dblDate As Double){
    //dblDate = Round(dblDate, 0)
    long y1,m1,d1;
    double h1;
    Call swe_revjul(dblDate, 1, y1, m1, d1, h1)
    //kp_RevJday = Format(m1, "00") + "/" + Format(Int(d1), "00") + "/" + Format(y1, "0000")
    string s = Format(Int(d1), "00") + "/" + Format(m1, "00") + "/" + Format(y1, "0000");
    return s;
}
*/



double kp_Sub249Hor(int sInt1to249){
    //Sub Table

    double nakBor[27];

    for(int a0=0;a0<27;a0++)
    {
        nakBor[a0] = 360# * (a0 / 27#);
    }

    double sub243Bor(26, 8);

    for(int a1=0;a1<26;a1++)
    {
        sub243Bor(a1, 0) = nakBor[a1];
        for(int a2=0;a2<7;a2++)
        {
            sub243Bor(a1, a2 + 1) = sub243Bor(a1, a2) + 13.3333333333333 * (dasaYears(hp_Rnd0To8v(a1 + a2)) / 120#);
        }
    }

    double sub243(1 To 243) As Double
    Dim a3 As Byte, a4 As Byte, a5 As Byte
    int a3,a4,a5=0;

    for(a3=0;a3<26;a3++){
        for(a4=0;a4<8;a4++){
            a5 = a5 + 1;
            sub243[a5] = sub243Bor(a3, a4);
        }
    }
    
    Dim sub249(1 To 249) As Double
    
    //1 to 22
    a6 As Byte
    for(int a6=0;a6<22;a6++){
        sub249[a6] = sub243[a6];
    }
    //23     -1
    sub249[23] = 30# ;
    
    //24 to 62
    for(int a7=24;a7<62;a7++){
        sub249[a7] = sub243[a7 - 1];
    }
    
    //63     -2
    sub249[63] = 90# ;
    
    //64 to 105
    for(int a8=64;a8<105;a8++){
        sub249[a8] = sub243[a8 - 2];
    }
    
    //106     -3
    sub249[106] = 150# ;
     
    //107 to 145
    for(int a9=107;a9<145;a7++){
        sub249[a9] = sub243[a9 - 3];
    }

    //146     -4
    sub249[146] = 210# ;
    
    //147 to 188
    for(int b1=147;a9<188;b1++){
        sub249[b1] = sub243[b1 - 4];
    }
    
    //189     -5
    sub249[189] = 270# ;
    
    //190 to 228
    For b2 = 190 To 
    for(int b2=190;b2<228;b2++){
        sub249[b2] = sub243[b2 - 5];
    }
    
    //229     -6
    sub249[229] = 330# ;
    
    //230 to 249
    for(int b3=230;a9<249;b3++){
        sub249[b3] = sub243[b3 - 6];
    }
    
    return sub249[sInt1to249];
    
}

public:
    int kp_PID2N(int plntIntDas){
        // convert dasa planet int to normal planet int
        //kp_PlntIntDasToNor
        if(plntIntDas == 0) 
            return 8;
        if(plntIntDas == 1) 
            return 5;
        if(plntIntDas == 2) 
            return 0;
        if(plntIntDas == 3) 
            return 1;
        if(plntIntDas == 4) 
            return 2;
        if(plntIntDas == 5) 
            return 7;
        if(plntIntDas == 6) 
            return 4;
        if(plntIntDas == 7) 
            return 6;
        if(plntIntDas == 8) 
            return 3;

    //    If plntIntDas = 9 Then kp_PID2N = 0
    //    If plntIntDas = 10 Then kp_PID2N = 0
    //   If plntIntDas = 11 Then kp_PID2N = 0
}

double kp_NKPA(julDate As Double){
    double B,P,A,T,yDate;

    yDate = (julDate / 365.242199) - 4712.10699807445;                  //days to years
    
    T = yDate - 1900# ;
    
    B = 22# + (22 / 60#) + (16 / 3600#);   // = 22-22-30 @ 15 Apr 1900)
    P = 50.2388475;
    A = 0.000111;   //sec per year
    
    return B + (T * P + (T * T * A)) / 3600#;
}
