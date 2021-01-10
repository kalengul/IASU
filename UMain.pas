unit UMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComObj, ExtCtrls, Grids, UConstParametrs, UGroup, UAuditoria, USemPlan, UNagryzka;

const
  xlDown = -4121;
  xlCenter = -4108;
  xlRight = -4152;



type
  TFMain = class(TForm)
    BtIASY: TButton;
    ODIASY: TOpenDialog;
    Label1: TLabel;
    LaNameFile: TLabel;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    MeProtocol: TMemo;
    Panel7: TPanel;
    BtProverka: TButton;
    SGPrepod: TStringGrid;
    BtCreateResult: TButton;
    SgNagryzka: TStringGrid;
    Panel10: TPanel;
    BtSaveSgNagryzkaXLSX: TButton;
    BtRaspredelenieNagryzki: TButton;
    CbAutoProverkaClose: TCheckBox;
    BtExzToExcel: TButton;
    BtRaspisanie: TButton;
    BtMTO: TButton;
    BtGroup: TButton;
    BtVivodPrepVGr: TButton;
    BtPrepodInOOP: TButton;
    Button2: TButton;
    BtMTOSemestrov: TButton;
    BtAllPOAud: TButton;
    Button1: TButton;
    BtCreateYMK: TButton;
    BtTablePredmet: TButton;
    BtExzGr: TButton;
    Panel11: TPanel;
    PNagrO: TPanel;
    PGroup: TPanel;
    POsn: TPanel;
    PPrepod: TPanel;
    PEkz: TPanel;
    PNagrV: TPanel;
    PRaspPrepod: TPanel;
    PPO: TPanel;
    PRaspGroup: TPanel;
    PRPD: TPanel;
    LaKaf: TLabel;
    LaZavKaf: TLabel;
    SgNagryzkaSearth: TStringGrid;
    PMergeDis: TPanel;
    procedure BtExzGrClick(Sender: TObject);
    procedure BtTablePredmetClick(Sender: TObject);
    procedure BtCreateYMKClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure BtAllPOAudClick(Sender: TObject);
    procedure BtMTOSemestrovClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure BtPrepodInOOPClick(Sender: TObject);
    procedure BtVivodPrepVGrClick(Sender: TObject);
    procedure BtGroupClick(Sender: TObject);
    procedure BtMTOClick(Sender: TObject);
    procedure BtRaspisanieClick(Sender: TObject);
    procedure BtExzToExcelClick(Sender: TObject);
    procedure BtRaspredelenieNagryzkiClick(Sender: TObject);
    procedure SgNagryzkaSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure BtSaveSgNagryzkaXLSXClick(Sender: TObject);
    procedure BtCreateResultClick(Sender: TObject);
    procedure BtProverkaClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormActivate(Sender: TObject);
    procedure BtIASYClick(Sender: TObject);
    procedure SgNagryzkaSearthSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: string);
    procedure SgNagryzkaSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: string);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

type


  THourStudentDis = record
                    Dis,Vid,Group:String;
                    HourForOneStudent:Double;
                    NomRowNagryzka:LongWord;
                    Enabled:byte;
                    end;

var
  FMain: TFMain;
  CurrentDir:String;
  Excel,ExcelBase: Variant;
  Decl : Variant;
  HourStudentDis:array of THourStudentDis;
  NameFileNagryzka:array [1..kolsem] of string;

  ArrAuditSRS,ArrAuditoriiKP,ArrAuditoriiKons,ArrAuditoriiKontrol,ArrAuditoriiObslyz:array of Longword;
  RowSelectSgNagryzka:Longword;
  VivodProtocol:boolean;

implementation

uses UChangeNagryzka, UYchPlan, USaveExcel, ULoadExcel, UMTO, UIndividualPlanPrepod;

{$R *.dfm}

procedure TFMain.BtProverkaClick(Sender: TObject);
var
  NomRow,NomNagryzka,NomNagryzkaBase,NomNagryzkaSdvig,NomPrepod,NomSearchPrepod:Longword;
  Sem:Byte;
  StExcel,OldPrepod:string;
  ArrSt:array[1..kolrownagryzka] of string;
  st:string;
  NomArrSt:byte;
  VivodOsh:boolean;
begin
//��������� �������� ��������������
MeProtocol.Lines.Add('����� ��������');
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  Prepod[NomPrepod].P:=0;
  //Prepod[NomPrepod].FlagP:=0;
  NomNagryzka:=0;
  while NomNagryzka<length(Prepod[NomPrepod].Nagryzka) do
    begin
    Prepod[NomPrepod].Nagryzka[NomNagryzka].P:=0;
    inc(NomNagryzka);
    end;
  inc(NomPrepod);
  end;
//�������� ��������� ����� ����� ������������ ��������
//MeProtocol.Lines.Clear;
for Sem := 1 to kolsem do
if NameFileNagryzka[Sem]<>'' then
  begin
  Excel.Workbooks.Open(NameFileNagryzka[Sem]);
  NomNagryzka:=1;
  while (NomNagryzka<Length(Nagryzka)) do
    begin
    //�������� ������ �� ����� � ������������ � ������� �� ��������
    st:='';
    for NomArrSt := 1 to kolrownagryzka do
      begin
      ArrSt[NomArrSt]:=Excel.Cells[Nagryzka[NomNagryzka].NomRow,NomArrSt];
      st:=st+ArrSt[NomArrSt]+' ';
      end;
    //������ � ���������
    if ((Nagryzka[NomNagryzka].Sem=Sem) and
      not((Nagryzka[NomNagryzka].Dis=ArrSt[1]) and
          (Nagryzka[NomNagryzka].Vid=ArrSt[2]) and
          (Nagryzka[NomNagryzka].Group=ArrSt[3]) and
          (Nagryzka[NomNagryzka].Hour=ArrSt[4]) and
          ((Nagryzka[NomNagryzka].FIOPrep=ArrSt[5]) or (ArrSt[5]='') and (Nagryzka[NomNagryzka].FIOPrep='�� ���������')) and
          (Nagryzka[NomNagryzka].Opisanie=ArrSt[6])and
          (Nagryzka[NomNagryzka].NOMPrep=ArrSt[7])))  then
     begin
     //��� ������������ ���������
     MeProtocol.Lines.Add('������� '+Nagryzka[NomNagryzka].Dis+' '+Nagryzka[NomNagryzka].Vid+' '+Nagryzka[NomNagryzka].Group+' '+Nagryzka[NomNagryzka].Hour+' '+Nagryzka[NomNagryzka].FIOPrep+' '+Nagryzka[NomNagryzka].Opisanie);
     MeProtocol.Lines.Add('��      '+st);
     MeProtocol.Lines.Add('');
     //���� ������� ������������� ������������� ������� ������������� ���� ��������� ��������
     //����� ������������� � �������
     if Nagryzka[NomNagryzka].FIOPrep='�� ���������' then
       NomSearchPrepod:=SeartchPrepodFIO('�� ���������')
     else
       NomSearchPrepod:=SeartchPrepodFIO(Nagryzka[NomNagryzka].FIOPrep);
     //��������� ��� ����� :)
     if NomSearchPrepod<>65000 then
       Prepod[NomSearchPrepod].FlagP:=1
     else
       ShowMessage('������ ������ ������������� '+Nagryzka[NomNagryzka].FIOPrep);
     //� ������ � ������ ������������� ���� ��������� ��������
     //����� ������������� � �������
     NomSearchPrepod:=SeartchPrepodFIO(ArrSt[5]);
     //��������� ��� ����� :)
     if NomSearchPrepod<>65000 then
       Prepod[NomSearchPrepod].FlagP:=1
     else
       ShowMessage('������ ������ ������������� '+ArrSt[5]);
     end;
    inc(NomNagryzka);
    end;
  Excel.Workbooks.Close;
  end;
//MeProtocol.Lines.Clear;
//�������� ��������������
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  if Prepod[NomPrepod].NameFilePrepod<>'' then
    begin
    Excel.Workbooks.Open(Prepod[NomPrepod].NameFilePrepod);
    MeProtocol.Lines.Add(Prepod[NomPrepod].FIO);
    VivodOsh:=false;
    //�������� ����������� ��������

    //�������� �������� ��������
    NomRow:=2;
    StExcel:=Excel.Cells[NomRow,1];
    for Sem := 1 to kolsem do
    begin
    while StExcel<>'' do
      begin
      //��������� �������� �� ����� �������������
      st:='';
      for NomArrSt := 1 to kolrownagryzka do
        begin
        ArrSt[NomArrSt]:=Excel.Cells[NomRow,NomArrSt];
        st:=st+ArrSt[NomArrSt]+' ';
        end;

      //����� �������� ��������� � ����� � ����
      NomNagryzka:=0;
      while (NomNagryzka<Length(Prepod[NomPrepod].Nagryzka)) and
            ((Prepod[NomPrepod].Nagryzka[NomNagryzka].sem<>Sem) or
      not  ((Prepod[NomPrepod].Nagryzka[NomNagryzka].Dis=ArrSt[1]) and
            (Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid=ArrSt[2]) and
            (SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagryzka].Group,ArrSt[3])<>65000) and
            (Prepod[NomPrepod].Nagryzka[NomNagryzka].NOMPrep=ArrSt[6]))) do
        inc(NomNagryzka);

      if    (NomNagryzka<Length(Prepod[NomPrepod].Nagryzka)) and
            (Prepod[NomPrepod].Nagryzka[NomNagryzka].sem=Sem) and
            (Prepod[NomPrepod].Nagryzka[NomNagryzka].Dis=ArrSt[1]) and
            (Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid=ArrSt[2]) and
            (SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagryzka].Group,ArrSt[3])<>65000) and
            (Prepod[NomPrepod].Nagryzka[NomNagryzka].NOMPrep=ArrSt[6])  then
         begin
         inc(Prepod[NomPrepod].Nagryzka[NomNagryzka].P);  //����� �������� �� �����
         //�������� ���������� ��������
         if (Prepod[NomPrepod].Nagryzka[NomNagryzka].Opisanie<>ArrSt[5]) then
           begin
           //����� ���� �������� � ���� ��������
           NomNagryzkaBase:=0;
           while (NomNagryzkaBase<Length(Nagryzka)) and
                 ((Nagryzka[NomNagryzkaBase].Sem<>Sem) or
           not  ((Nagryzka[NomNagryzkaBase].FIOPrep=Prepod[NomPrepod].FIO) and
                 (Nagryzka[NomNagryzkaBase].Dis=ArrSt[1]) and
                 (Nagryzka[NomNagryzkaBase].Vid=ArrSt[2]) and
                 (Nagryzka[NomNagryzkaBase].Group=ArrSt[3]) and
                 (Nagryzka[NomNagryzkaBase].Hour=ArrSt[4]) and
                 (Nagryzka[NomNagryzkaBase].Opisanie<>ArrSt[5]) and
                 (Nagryzka[NomNagryzkaBase].NOMPrep=ArrSt[6]))) do
              inc(NomNagryzkaBase);
           if (NomNagryzkaBase<Length(Nagryzka)) and
              ((Nagryzka[NomNagryzkaBase].FIOPrep=Prepod[NomPrepod].FIO) and
              (Nagryzka[NomNagryzkaBase].Sem=Sem) and
              (Nagryzka[NomNagryzkaBase].Dis=ArrSt[1]) and
              (Nagryzka[NomNagryzkaBase].Vid=ArrSt[2]) and
              (Nagryzka[NomNagryzkaBase].Group=ArrSt[3]) and
              (Nagryzka[NomNagryzkaBase].Hour=ArrSt[4]) and
              (Nagryzka[NomNagryzkaBase].Opisanie<>ArrSt[5]) and
              (Nagryzka[NomNagryzkaBase].NOMPrep=ArrSt[6])) then
              begin
              //�������� EXCEL �����
              ExcelBase.Workbooks.Open(NameFileNagryzka[Sem]);
              MeProtocol.Lines.Add('�������� �������� �������� '+Nagryzka[NomNagryzkaBase].Dis+' '+Nagryzka[NomNagryzkaBase].Vid+' '+Nagryzka[NomNagryzkaBase].Group+' '+Nagryzka[NomNagryzkaBase].Hour+' '+Nagryzka[NomNagryzkaBase].Opisanie);
              MeProtocol.Lines.Add('�� '+ArrSt[5]);
              VivodOsh:=true;
              //��������� �������� ��������
              ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow,6]:=ArrSt[5];
              end;
           ExcelBase.Workbooks[1].Save;
           ExcelBase.Workbooks.Close;
           end;
         //�������� ��������� �����
         if StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagryzka].Hour)<StrTOFloat(ArrSt[4]) then
           showmessage('�������� �� ���������� �������� � ��������� �� ������������� (����������)')
         else
         if (Prepod[NomPrepod].Nagryzka[NomNagryzka].Hour<>ArrSt[4]) then
           begin
           //����� ���� �������� � ���� ��������
           Prepod[NomPrepod].FlagP:=1;
           NomNagryzkaBase:=0;
           while (NomNagryzkaBase<Length(Nagryzka)) and
                 ((Nagryzka[NomNagryzkaBase].Sem<>Sem) or
           not  ((Nagryzka[NomNagryzkaBase].FIOPrep=Prepod[NomPrepod].FIO) and
                 (Nagryzka[NomNagryzkaBase].Dis=ArrSt[1]) and
                 (Nagryzka[NomNagryzkaBase].Vid=ArrSt[2]) and
                 (Nagryzka[NomNagryzkaBase].Hour<>ArrSt[4]) and
                 (Nagryzka[NomNagryzkaBase].Group=ArrSt[3]))) do
              inc(NomNagryzkaBase);
           if (NomNagryzkaBase<Length(Nagryzka)) and
              (Nagryzka[NomNagryzkaBase].FIOPrep=Prepod[NomPrepod].FIO) and
              (Nagryzka[NomNagryzkaBase].Sem=Sem) and
              (Nagryzka[NomNagryzkaBase].Dis=ArrSt[1]) and
              (Nagryzka[NomNagryzkaBase].Vid=ArrSt[2]) and
              (Nagryzka[NomNagryzkaBase].Hour<>ArrSt[4]) and
              (Nagryzka[NomNagryzkaBase].Group=ArrSt[3]) then
              begin
              //�������� EXCEL �����
              ExcelBase.Workbooks.Open(NameFileNagryzka[Sem]);
              MeProtocol.Lines.Add('�������� ���� �������� '+Nagryzka[NomNagryzkaBase].Dis+' '+Nagryzka[NomNagryzkaBase].Vid+' '+Nagryzka[NomNagryzkaBase].Group+' '+Nagryzka[NomNagryzkaBase].Hour+' '+Nagryzka[NomNagryzkaBase].Opisanie);
              MeProtocol.Lines.Add('�� '+ArrSt[4]);
              VivodOsh:=true;
              //��������� ����� ��������
              //�������� ������
              //ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow].Insert(xlDown);
              ExcelBase.ActiveSheet.Rows[Nagryzka[NomNagryzkaBase].NomRow+1].Select;
              ExcelBase.Selection.Insert(Shift :=xlDown);

             //� ����� ������ ���������� ��������� �� ������
              st:=ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow,1];
              ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow+1,1]:=st;
              st:=ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow,2];
              ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow+1,2]:=st;
              st:=ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow,3];
              ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow+1,3]:=st;
              ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow+1,4]:=StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagryzka].Hour)-StrToFloat(ArrSt[4]);
              ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow+1,5]:='';
              st:=ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow,6];
              ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow+1,6]:=st;
              st:=ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow,7];
              ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow+1,7]:=st;
             //�������� ��� ������ � ���� NomRow ��� ���� ��������� ��������
              For NomNagryzkaSdvig:=0 to Length(Nagryzka)-1 do
                if (Nagryzka[NomNagryzkaBase].Sem=Nagryzka[NomNagryzkaSdvig].Sem) and (Nagryzka[NomNagryzkaBase].NomRow<Nagryzka[NomNagryzkaSdvig].NomRow) then
                  inc(Nagryzka[NomNagryzkaSdvig].NomRow);              
              //�������� ����� ������ � ������ ��������
              NomNagryzkaSdvig:=Length(Nagryzka);
              SetLength(Nagryzka,NomNagryzkaSdvig+1);
              Nagryzka[NomNagryzkaSdvig].P:=0;
              Nagryzka[NomNagryzkaSdvig].NomRow:=Nagryzka[NomNagryzkaBase].NomRow+1;
              Nagryzka[NomNagryzkaSdvig].Sem:=Nagryzka[NomNagryzkaBase].Sem;
              Nagryzka[NomNagryzkaSdvig].Dis:=Nagryzka[NomNagryzkaBase].Dis;
              Nagryzka[NomNagryzkaSdvig].Vid:=Nagryzka[NomNagryzkaBase].Vid;
              Nagryzka[NomNagryzkaSdvig].Group:=Nagryzka[NomNagryzkaBase].Group;
              Nagryzka[NomNagryzkaSdvig].Hour:=ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow+1,4];
              Nagryzka[NomNagryzkaSdvig].FIOPrep:='�� ���������';
              Nagryzka[NomNagryzkaSdvig].Opisanie:=Nagryzka[NomNagryzkaBase].Opisanie;
              if Nagryzka[NomNagryzkaSdvig].NOMPrep='' then
                begin
                Nagryzka[NomNagryzkaBase].NOMPrep:='1';
                ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow,7]:='1';
                Nagryzka[NomNagryzkaSdvig].NOMPrep:='2';
                ExcelBase.Cells[Nagryzka[NomNagryzkaSdvig].NomRow,7]:='2';
                end
              else
                begin
                Nagryzka[NomNagryzkaSdvig].NOMPrep:=IntToStr(StrToInt(Nagryzka[NomNagryzkaBase].NOMPrep)+1);
                ExcelBase.Cells[Nagryzka[NomNagryzkaSdvig].NomRow,7]:=Nagryzka[NomNagryzkaSdvig].NOMPrep;
                end;
              //�������� ���������� ����� � �������
              ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow,4]:=ArrSt[4];
              //�������� ���������� ����� � ����
              Nagryzka[NomNagryzkaBase].Hour:=ArrSt[4];
              //����� "�� ���������" � ���������� ����
              NomSearchPrepod:=SeartchPrepodFIO('�� ���������');
              //��������� ��� ����� :)
              if NomSearchPrepod<>65000 then
                Prepod[NomSearchPrepod].FlagP:=1
              else
                ShowMessage('������ ������ ������������� �� ���������');
              end;

           ExcelBase.Workbooks[1].Save;
           ExcelBase.Workbooks.Close;
           end;
         end
      else
        begin
        //���� �� ����� �������� �� ����� (������� ����� ��������)
        //����� ���� �������� � ���� ��������
        NomNagryzkaBase:=0;
        while (NomNagryzkaBase<Length(Nagryzka)) and
              ((Nagryzka[NomNagryzkaBase].Sem<>Sem) or
          not ((Nagryzka[NomNagryzkaBase].Dis=ArrSt[1]) and
            (Nagryzka[NomNagryzkaBase].Vid=ArrSt[2]) and
            (Nagryzka[NomNagryzkaBase].Group=ArrSt[3]) and
            (Nagryzka[NomNagryzkaBase].Hour=ArrSt[4]) and
            (Nagryzka[NomNagryzkaBase].NOMPrep=ArrSt[6]))) do
          inc(NomNagryzkaBase);
        if (NomNagryzkaBase<Length(Nagryzka)) and
           ((Nagryzka[NomNagryzkaBase].Sem=Sem) and
            (Nagryzka[NomNagryzkaBase].Dis=ArrSt[1]) and
            (Nagryzka[NomNagryzkaBase].Vid=ArrSt[2]) and
            (Nagryzka[NomNagryzkaBase].Group=ArrSt[3]) and
            (Nagryzka[NomNagryzkaBase].Hour=ArrSt[4])and
            (Nagryzka[NomNagryzkaBase].NOMPrep=ArrSt[6])) then
           begin
           //�������� EXCEL �����
           ExcelBase.Workbooks.Open(NameFileNagryzka[Sem]);
           OldPrepod:=ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow,5];
           MeProtocol.Lines.Add('��������� �������� '+Nagryzka[NomNagryzkaBase].Dis+' '+Nagryzka[NomNagryzkaBase].Vid+' '+Nagryzka[NomNagryzkaBase].Group+' '+Nagryzka[NomNagryzkaBase].Hour+' '+Nagryzka[NomNagryzkaBase].Opisanie);
           VivodOsh:=true;
           //���� ������� ������������� ������������� ��� ���� ��������� �������� (�������� ���������� ��������)
           if OldPrepod<>Prepod[NomPrepod].FIO then
             begin
             MeProtocol.Lines.Add('� ����� �������� ���� ��������� '+OldPrepod);
             //����� ������������� � �������
             NomSearchPrepod:=SeartchPrepodFIO(Nagryzka[NomNagryzkaBase].FIOPrep);
             //��������� ��� ����� :)
             if NomSearchPrepod<>65000 then
               Prepod[NomSearchPrepod].FlagP:=1
             else
               ShowMessage('������ ������ ������������� '+Nagryzka[NomNagryzka].FIOPrep);

             //��������� �������������
             ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow,5]:=Prepod[NomPrepod].FIO;
             end;
           ExcelBase.Workbooks[1].Save;
           ExcelBase.Workbooks.Close;
           end;
        end;

      //��������� � ��������� ������
      inc(NomRow);
      StExcel:=Excel.Cells[NomRow,1];
      end;
      //������� ������ ������ ��� ����� ���������
      inc(NomRow);
      StExcel:=Excel.Cells[NomRow,1];
      end;
    Excel.Workbooks.Close;

    //����� ��������, ������� �� ������� � ����� ������������� (��������� ��������)
    NomNagryzka:=0;
    while (NomNagryzka<Length(Prepod[NomPrepod].Nagryzka)) do
    begin
      while (NomNagryzka<Length(Prepod[NomPrepod].Nagryzka)) and (Prepod[NomPrepod].Nagryzka[NomNagryzka].P=1) do
        inc(NomNagryzka);
      if (NomNagryzka<Length(Prepod[NomPrepod].Nagryzka)) and (Prepod[NomPrepod].Nagryzka[NomNagryzka].P<>1) then
         begin
        //����� ���� �������� � ���� ��������
        Prepod[NomPrepod].FlagP:=1;
        NomNagryzkaBase:=0;
        while (NomNagryzkaBase<Length(Nagryzka)) and
              ((Nagryzka[NomNagryzkaBase].Sem<>Prepod[NomPrepod].Nagryzka[NomNagryzka].sem) or
        not((Nagryzka[NomNagryzkaBase].Dis=Prepod[NomPrepod].Nagryzka[NomNagryzka].Dis) and
            (Nagryzka[NomNagryzkaBase].Vid=Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid) and
            (SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagryzka].Group,Nagryzka[NomNagryzkaBase].Group)<>65000) and
            (Nagryzka[NomNagryzkaBase].Hour=Prepod[NomPrepod].Nagryzka[NomNagryzka].Hour) and
            (Nagryzka[NomNagryzkaBase].NOMPrep=Prepod[NomPrepod].Nagryzka[NomNagryzka].NOMPrep))) do
          inc(NomNagryzkaBase);
        if (NomNagryzkaBase<Length(Nagryzka)) and
           ((Nagryzka[NomNagryzkaBase].Sem=Prepod[NomPrepod].Nagryzka[NomNagryzka].sem) and
            (Nagryzka[NomNagryzkaBase].Dis=Prepod[NomPrepod].Nagryzka[NomNagryzka].Dis) and
            (Nagryzka[NomNagryzkaBase].Vid=Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid) and
            (SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagryzka].Group,Nagryzka[NomNagryzkaBase].Group)<>65000) and
            (Nagryzka[NomNagryzkaBase].Hour=Prepod[NomPrepod].Nagryzka[NomNagryzka].Hour) and
            (Nagryzka[NomNagryzkaBase].NOMPrep=Prepod[NomPrepod].Nagryzka[NomNagryzka].NOMPrep)) then
           begin
           //�������� EXCEL �����
           ExcelBase.Workbooks.Open(NameFileNagryzka[Prepod[NomPrepod].Nagryzka[NomNagryzka].sem]);
           OldPrepod:=ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow,5];
           MeProtocol.Lines.Add('��������� �������� '+Nagryzka[NomNagryzkaBase].Dis+' '+Nagryzka[NomNagryzkaBase].Vid+' '+Nagryzka[NomNagryzkaBase].Group+' '+Nagryzka[NomNagryzkaBase].Hour+' '+Nagryzka[NomNagryzkaBase].Opisanie);
           VivodOsh:=true;
           //���� ������� ������������� ������������� ��� ���� ��������� �������� (�������� ���������� ��������)
           if OldPrepod=Prepod[NomPrepod].FIO then
             begin
             MeProtocol.Lines.Add('� ����� �������� ������ ��� ������������� ������������ ');
             //��������� �������������
             ExcelBase.Cells[Nagryzka[NomNagryzkaBase].NomRow,5]:='';
             //����� ������������� � �������
             NomSearchPrepod:=SeartchPrepodFIO('�� ���������');
             //��������� ��� ����� :)
             if NomSearchPrepod<>65000 then
               Prepod[NomSearchPrepod].FlagP:=1
             else
               ShowMessage('������ ������ �������������: �� ���������');
             end;
           ExcelBase.Workbooks[1].Save;
           ExcelBase.Workbooks.Close;
           end;
        end;
      inc(NomNagryzka)
      end;
    end;
  if not VivodOsh then
    MeProtocol.Lines.Strings[MeProtocol.Lines.Count-1]:=MeProtocol.Lines.Strings[MeProtocol.Lines.Count-1]+' - OK';
  inc(NomPrepod);
  end;

//��� ���� ��������������, ������� �� ������ ��������, �������� �� �����
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  if Prepod[NomPrepod].FlagP<>0 then
    begin
    MeProtocol.Lines.Add('�������� ����� ��� ������������� '+Prepod[NomPrepod].FIO);
    MeProtocol.Lines.Add(Prepod[NomPrepod].NameFilePrepod);
    DeleteFile(Prepod[NomPrepod].NameFilePrepod);
    MeProtocol.Lines.Add('');
    end;
  inc(NomPrepod);
  end;

//���������� �������� �����������
ProverkaStart;

For NomPrepod:=0 to Length(Prepod)-1 do
  Prepod[NomPrepod].FlagP:=0;

ShowMessage('�������� ���������');
end;

procedure TFMain.BtCreateResultClick(Sender: TObject);
begin
CreateIndividualPlanAllPrepod(MeProtocol);
end;

Procedure GoAllExzamenGroupInExcelFile(NomSem:Byte; FileName:string);
var
NomPrepod,NomNagr,NomRow,NomGroupNagryzka:longword;
StGroup:string;
begin
Excel.WorkBooks.Add;
NomRow:=2;
NomPrepod:=0;
    while NomPrepod<length(Prepod) do
      begin
      NomNagr:=0;
      while NomNagr<length(Prepod[NomPrepod].Nagryzka) do
        begin
        if ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='�������') and (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem)) then
          begin
          StGroup:='';
          NomGroupNagryzka:=0;
          while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
            begin
            StGroup:=StGroup+Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka].Nom+', ';
            inc(NomGroupNagryzka);
            end;
          Excel.Cells[NomRow,1]:=StGroup;
          Excel.Cells[NomRow,2]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
          Excel.Cells[NomRow,3]:=Prepod[NomPrepod].Nagryzka[NomNagr].Vid;
          Excel.Cells[NomRow,4]:=Prepod[NomPrepod].FIO;
          Excel.Cells[NomRow,5]:=Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria;
          inc(NomRow);
          end;
        inc(NomNagr);
        end;
      inc(NomPrepod);
      end;
Excel.Workbooks[1].saveas(FileName);
Fmain.MeProtocol.Lines.Add('������ ���� '+FileName);
Excel.Workbooks.Close;
end;

procedure TFMain.BtExzGrClick(Sender: TObject);
begin
if not DirectoryExists(CurrentDir+'\�������� �� �������') then
    ForceDirectories(CurrentDir+'\�������� �� �������');
GoAllExzamenGroupInExcelFile(2,CurrentDir+'\�������� �� �������\������� � �������.xlsx');

end;

procedure TFMain.BtExzToExcelClick(Sender: TObject);
begin
CreateAllGroup;
GoExzamenToExcel(1);
GoExzamenToExcel(2);


GoTablExzamenGroupToExcel(1);

ShowMessage('������� ����� � ����������� ��������� (����� � �����)');
end;

procedure TFMain.BtIASYClick(Sender: TObject);
begin
if ODIASY.Execute then
  begin
  CurrentDir:=ODIASY.FileName;
  end;
end;

//��������� � ����� ������� ��� ������ (��/��/�������� � �.�.) � "��������" ���������������
Procedure Nomer2PrepodNagryzka(FileName:String);
var
NomRow,NomRowSearch:Longword;
Nom2Prepod:Longword;
st,stNew,SearchSt1,SearchSt2,SearchSt3,SravnenieST1,SravnenieST2,SravnenieST3:string;
begin
if FileExists(FileName) then
begin
Excel.Workbooks.Open(FileName);
NomRow:=2;
st:=Excel.Cells[NomRow,1];
While st<>'' do
  begin
  SearchSt1:=Excel.Cells[NomRow,7];
  if SearchSt1='' then
    begin
    Nom2Prepod:=1;
    SearchSt1:=Excel.Cells[NomRow,1];
    SearchSt2:=Excel.Cells[NomRow,2];
    SearchSt3:=Excel.Cells[NomRow,3];
    NomRowSearch:=NomRow+1;
    stNew:=Excel.Cells[NomRowSearch,1];
    While stNew<>'' do
      begin
      SravnenieST1:=Excel.Cells[NomRowSearch,1];
      SravnenieST2:=Excel.Cells[NomRowSearch,2];
      SravnenieST3:=Excel.Cells[NomRowSearch,3];
      if (SravnenieST1=SearchSt1) and (SravnenieST2=SearchSt2) and (SravnenieST3=SearchSt3) then
        begin
        inc(Nom2Prepod);
        Excel.Cells[NomRow,7]:='1';
        Excel.Cells[NomRowSearch,7]:=Nom2Prepod;

        end;
      inc(NomRowSearch);
      stNew:=Excel.Cells[NomRowSearch,1];
      end;
    end;
  inc(NomRow);
  st:=Excel.Cells[NomRow,1];
  end;
Excel.Workbooks[1].Save;
Excel.Workbooks.Close;
end;
end;

procedure TFMain.BtPrepodInOOPClick(Sender: TObject);
var
st,StNew:string;
NomPrepod,NomNagryzka,NomRow:Longword;
begin
if FileExists(CurrentDir+'\������� ���������.xlsx') then
begin
Excel.Workbooks.Open(CurrentDir+'\������� ���������.xlsx');
NomRow:=1;
st:=Excel.Cells[NomRow,1];
while St<>'End' do
  begin
  StNew:='';
  NomPrepod:=0;
  while NomPrepod<Length(Prepod) do
    begin
    NomNagryzka:=0;
    while NomNagryzka<Length(Prepod[NomPrepod].Nagryzka) do
      begin
      if (Prepod[NomPrepod].Nagryzka[NomNagryzka].Dis=st) and (Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid='��') and (Pos(Prepod[NomPrepod].FIO,StNew)=0)then
        StNew:=StNew+Prepod[NomPrepod].FIO+', ';
      inc(NomNagryzka);
      end;
    inc(NomPrepod);
    end;
  Excel.Cells[NomRow,2]:=StNew;
  inc(NomRow);
  st:=Excel.Cells[NomRow,1];
  end;
//Excel.Workbooks.save;
Excel.Workbooks.Close;
end;
end;

procedure TFMain.FormActivate(Sender: TObject);
begin
CurrentDir := GetCurrentDir;
SetLength(ArrAuditorii,1);
ArrAuditorii[0]:=TAuditoria.Create;
ArrAuditorii[0].Auditoria:='';
//������� ������ Excel
Excel := CreateOleObject('Excel.Application');
ExcelBase := CreateOleObject('Excel.Application');
//������ ���� Excel ���������
Excel.Visible := false;
ExcelBase.Visible := false;
if FileExists(CurrentDir+'\Settings.xlsx') then
    begin
    MeProtocol.Lines.Add('���������� �������� � ����:'+CurrentDir+'\Settings.xlsx');
    LoadInitializationParametrsInExcelFile(CurrentDir+'\Settings.xlsx');
    end
  else
    begin
    InitializationParametrs;
    MeProtocol.Lines.Add('�� ������ ���� � �����������:'+CurrentDir+'\Settings.xlsx');
    MeProtocol.Lines.Add('��������� ����������� ���������');
    end;
LaZavKaf.Caption:='���������� �������� '+ZKafSokr;
LaKaf.Caption:='��� �'+NomKaf;
StartLoadExcel(MeProtocol);
ShowMessage('�������� ���������');
end;

procedure TFMain.BtSaveSgNagryzkaXLSXClick(Sender: TObject);
begin
If ODIASY.Execute then
  begin
  VivodSgExcel(SgNagryzka,ODIASY.FileName);
  MeProtocol.Lines.Add('������ ������� ��������� � ����� '+ODIASY.FileName);
  end;
end;

procedure TFMain.BtVivodPrepVGrClick(Sender: TObject);
begin
if not DirectoryExists(CurrentDir+'\���������� ���� �� �������') then
    ForceDirectories(CurrentDir+'\���������� ���� �� �������');
LoadAllRaspisanieAllGroup(CurrentDir+'\���������� ���� �� �������\');
CreateAllGroup;

SaveAllPrepodGroup;
MeProtocol.Lines.Add('���������� �������������� � ������ ���������.');
end;

Procedure SaveDisPO(FileName:string);
var
NomRow,KolArrDis,NomArrDis:longword;
NomSemYp,NomDis:Longword;
st,st1:string;
i:byte;
ArrDis: array of String;
begin
if FileExists(FileName) then
  Excel.Workbooks.Open(FileName)
else
  Excel.Workbooks.add;
NomRow:=2;
SetLength(ArrDis,0);
KolArrDis:=0;
St1:=Excel.Cells[NomRow,2];
while st1<>'' do
  begin
  SetLength(ArrDis,KolArrDis+1);
  ArrDis[KolArrDis]:=Excel.Cells[NomRow,2];
  inc(KolArrDis);
  inc(NomRow);
  St1:=Excel.Cells[NomRow,2];
  end;
NomSemYp:=0;
while NomSemYp<Length(SemYP) do
  begin
  NomDis:=0;
  While NomDis<Length(SemYp[NomSemYp].Disciplin) do
    begin
    if (SemYp[NomSemYp].Disciplin[NomDis].LR<>0) or (SemYp[NomSemYp].Disciplin[NomDis].PZ<>0) then
    begin
    NomArrDis:=0;
    while (NomArrDis<Length(ArrDis)) and (SemYp[NomSemYp].Disciplin[NomDis].Name<>ArrDis[NomArrDis]) do
      inc(NomArrDis);
    if not (NomArrDis<Length(ArrDis)) then
      begin
      Excel.Cells[NomRow,1]:=NomRow-1;
      Excel.Cells[NomRow,2]:=SemYp[NomSemYp].Disciplin[NomDis].Name;
      for I := 0 to 3 do
        begin
        st:=Excel.Cells[NomRow-1,3+i];
        Excel.Cells[NomRow,3+i]:=st;
        end;
      Excel.Cells[NomRow,7]:=SemYp[NomSemYp].Disciplin[NomDis].PO;
      Excel.Cells[NomRow,8]:=SemYp[NomSemYp].Profil;
      Excel.Cells[NomRow,9]:=SemYp[NomSemYp].Nom;
      Excel.Cells[NomRow,10]:=SemYp[NomSemYp].Disciplin[NomDis].PoRPD;
      Excel.Cells[NomRow,11]:=SemYp[NomSemYp].Disciplin[NomDis].OsnashenieRPD;
      inc(NomRow);
      SetLength(ArrDis,KolArrDis+1);
      ArrDis[KolArrDis]:=SemYp[NomSemYp].Disciplin[NomDis].Name;
      inc(KolArrDis);
      end;
    end;
    inc(NomDis);
    end;
  inc(NomSemYp);
  end;
Excel.Workbooks[1].saveas(FileName);
Fmain.MeProtocol.Lines.Add('�������� ���� '+FileName);
Excel.Workbooks.Close;

end;

procedure TFMain.BtTablePredmetClick(Sender: TObject);
var
  NomProfil,NomSemPlan,NomDisSemYP,NomSemPlanToo,NomDisSemYPToo,NomPrepod,NomNagryzka,kolDisAll:Longword;
  NomRow,NomRow1,Nom:Longword;
  st:string;
begin
LoadSemPlan;
Excel.WorkBooks.Add;
NomRow:=1;
    Excel.Columns[1].ColumnWidth := 3.14;
    Excel.Columns[2].ColumnWidth := 32.00;
    Excel.Columns[3].ColumnWidth := 5.86;
    kolDisAll:=0;

NomSemPlan:=0;
  while NomSemPlan<Length(SemYp) do
    begin
    NomDisSemYP:=0;
    while NomDisSemYP<Length(SemYp[NomSemPlan].Disciplin) do
      begin
      SemYp[NomSemPlan].Disciplin[NomDisSemYP].BYCh:=false;
      inc(NomDisSemYP);
      end;
    inc(NomSemPlan);
    end;
NomProfil:=0;
while NomProfil<Length(ArrProfil) do
  begin

    Excel.Cells[NomRow,1]:=ArrProfil[NomProfil].NameNaprav+' ������� - '+ArrProfil[NomProfil].NameProfil;
   // Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,4]].MergeCells:=true;
    NomRow:=NomRow+2;
    Excel.Cells[NomRow,1]:='� �\�';
    Excel.Cells[NomRow,2]:='������������ ����������';
    Excel.Cells[NomRow,3]:='���';
    Excel.Cells[NomRow,4]:='�������';
    Excel.Cells[NomRow,5]:='�����.';
    Excel.Cells[NomRow,6]:='�������';
    Excel.Cells[NomRow,7]:='������';

    inc(NomRow);
    Nom:=1;
  NomSemPlan:=0;
  while NomSemPlan<Length(ArrProfil[NomProfil].SemYp) do
    begin
    NomDisSemYP:=0;
    while NomDisSemYP<Length(ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin) do
      begin
      if not ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].BYCh then
        begin

        NomSemPlanToo:=0;
        while NomSemPlanToo<Length(ArrProfil[NomProfil].SemYp) do
          begin
          NomDisSemYPToo:=0;
          while NomDisSemYPToo<Length(ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin) do
            begin
            if ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].Name=ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].Name then
              begin
              ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].BYCh:=true;
              end;
            inc(NomDisSemYPToo);
            end;
          inc(NomSemPlanToo);
          end;

      inc(kolDisAll);
        Excel.Cells[NomRow,1]:=Nom;
        if ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].NomElektivDis=65000 then
          Excel.Cells[NomRow,2]:=ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].Name
        else
          Excel.Cells[NomRow,2]:='+'+ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].Name;
        if ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].Kaf='302' then
        Excel.Cells[NomRow,3]:='304'
        else
        Excel.Cells[NomRow,3]:=ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].Kaf;

        NomRow:=NomRow+1;
        inc(Nom);
        end;
      inc(NomDisSemYP);
      end;
    inc(NomSemPlan);
    end;
  NomRow:=NomRow+3;
  inc(NomProfil)
  end;
  Excel.Cells[NomRow+1,1]:=kolDisAll;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,4]].WrapText:=true;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,4]].VerticalAlignment:=xlCenter;
  Excel.Workbooks[1].saveas(CurrentDir+'\��� �������.xlsx');
  Fmain.MeProtocol.Lines.Add('������ ���� '+CurrentDir+'\��� �������.xlsx');
  Excel.Workbooks.Close;


end;

procedure TFMain.BtCreateYMKClick(Sender: TObject);
var
  NomProfil,NomSemPlan,NomDisSemYP,NomSemPlanToo,NomDisSemYPToo,NomPrepod,NomNagryzka:Longword;
  NomRow,NomRow1,Nom:Longword;
  st,st1:string;
begin
LoadSemPlan;
Excel.WorkBooks.Add;
NomRow:=1;
    Excel.Columns[1].ColumnWidth := 3.14;
    Excel.Columns[2].ColumnWidth := 32.00;
    Excel.Columns[3].ColumnWidth := 5.86;
    Excel.Columns[4].ColumnWidth := 33;
    Excel.Columns[5].ColumnWidth := 59.86;
NomSemPlan:=0;
  while NomSemPlan<Length(SemYp) do
    begin
    NomDisSemYP:=0;
    while NomDisSemYP<Length(SemYp[NomSemPlan].Disciplin) do
      begin
      SemYp[NomSemPlan].Disciplin[NomDisSemYP].BYCh:=false;
      inc(NomDisSemYP);
      end;
    inc(NomSemPlan);
    end;
NomProfil:=0;
while NomProfil<Length(ArrProfil) do
  begin

    Excel.Cells[NomRow,1]:=ArrProfil[NomProfil].NameNaprav+' ������� - '+ArrProfil[NomProfil].NameProfil;
    Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,5]].MergeCells:=true;
    NomRow:=NomRow+2;
    Excel.Cells[NomRow,1]:='� �\�';
    Excel.Cells[NomRow,2]:='������������ ����������';
    Excel.Cells[NomRow,3]:='���';
    Excel.Cells[NomRow,4]:='�������������';
    Excel.Cells[NomRow,5]:='����� ���������� ���';
    inc(NomRow);
    Nom:=1;
  NomSemPlan:=0;
  while NomSemPlan<Length(ArrProfil[NomProfil].SemYp) do
    begin
    NomDisSemYP:=0;
    while NomDisSemYP<Length(ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin) do
      begin
      if not ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].BYCh then
        begin
        st1:='';
        NomSemPlanToo:=0;
        while NomSemPlanToo<Length(ArrProfil[NomProfil].SemYp) do
          begin
          NomDisSemYPToo:=0;
          while NomDisSemYPToo<Length(ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin) do
            begin
            if ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].Name=ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].Name then
              begin
              ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].BYCh:=true;
             // st1:=st1+' '+ArrProfil[NomProfil].SemYp[NomSemPlan].
              end;
            inc(NomDisSemYPToo);
            end;
          inc(NomSemPlanToo);
          end;
        st:='';
        NomRow1:=NomRow;
        NomPrepod:=0;
              while NomPrepod<Length(Prepod) do
                begin
                NomNagryzka:=0;
                while NomNagryzka<Length(Prepod[NomPrepod].Nagryzka) do
                  begin
                  if (Prepod[NomPrepod].Nagryzka[NomNagryzka].Dis=ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].Name) and (Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid='��') and (Pos(Prepod[NomPrepod].FIO,st)=0)then
                    begin
                    st:=st+' '+Prepod[NomPrepod].FIO;
                    Excel.Cells[NomRow,5]:=Prepod[NomPrepod].FIO;
                    inc(NomRow);
                    end;
                  inc(NomNagryzka);
                  end;
                inc(NomPrepod);
                end;
        if NomRow>NomRow1 then
          dec(NomRow);
        Excel.Cells[NomRow,1]:=Nom;
        if ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].NomElektivDis=65000 then
          Excel.Cells[NomRow,2]:=ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].Name
        else
          Excel.Cells[NomRow,2]:='+'+ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].Name;
        if ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].Kaf='302' then
        Excel.Cells[NomRow,3]:='304'
        else
        Excel.Cells[NomRow,3]:=ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].Kaf;
        Excel.Cells[NomRow,4]:=st1;
        If NomRow>NomRow1 then
        begin
        Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow1,1]].MergeCells:=true;
        Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow1,2]].MergeCells:=true;
        Excel.Range[Excel.Cells[NomRow,3],Excel.Cells[NomRow1,3]].MergeCells:=true;
        end;
        NomRow:=NomRow+1;
        inc(Nom);
        end;
      inc(NomDisSemYP);
      end;
    inc(NomSemPlan);
    end;
  NomRow:=NomRow+3;
  inc(NomProfil)
  end;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,5]].WrapText:=true;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,6]].VerticalAlignment:=xlCenter;
  Excel.Workbooks[1].saveas(CurrentDir+'\���.xlsx');
  Fmain.MeProtocol.Lines.Add('������ ���� '+CurrentDir+'\���.xlsx');
  Excel.Workbooks.Close;


end;

procedure TFMain.Button1Click(Sender: TObject);
begin
LoadSemPlan;
if not DirectoryExists(CurrentDir+'\���\��') then
    ForceDirectories(CurrentDir+'\���\��');
LoadPo(CurrentDir+'\���\��\��.xlsx');
LoadMTOIASYFile(CurrentDir+'\��� � ���.xlsx');

SaveDisPO(CurrentDir+'\���\��\��.xlsx');
SaveDisPO(CurrentDir+'\��_all.xlsx');
end;

procedure TFMain.Button2Click(Sender: TObject);
var
SearchStr:TSearchRec;
NomRow,NomCol:Longword;
KolYP,KolDis,KolGroup:longword;
st:string;
Nom:Byte;
begin
if FileExists(CurrentDir+'\�������� ��.xlsx') then
  ExcelBase.Workbooks.Open(CurrentDir+'\�������� ��.xlsx');
if not DirectoryExists(CurrentDir+'\��') then
  ForceDirectories(CurrentDir+'\��');
KolYP:=0;
if FindFirst(CurrentDir+'\��\'+'*.xls',faDirectory,SearchStr)=0 then
  begin
  repeat
  Excel.Workbooks.Open(CurrentDir+'\��\'+SearchStr.Name);

  SetLength(YP,KolYP+1);
  YP[KolYP]:=TYp.Create;
  YP[KolYP].Name:=Excel.Cells[3,3];

  NomRow:=2;
  st:=ExcelBase.Cells[NomRow,1];
  while st<>'' do
    begin
    if st=YP[KolYP].Name then
      begin
      KolGroup:=Length(YP[KolYP].Group);
      SetLength(YP[KolYP].Group,KolGroup+1);
      YP[KolYP].Group[KolGroup]:=SearchAndCreateGroup(ExcelBase.Cells[NomRow,2]);
      end;
    inc(NomRow);
    st:=ExcelBase.Cells[NomRow,1];
    end;

  NomRow:=8;
  st:=Excel.Cells[NomRow,1];
  KolDis:=0;
  while st<>'' do
    begin
    st:=Excel.Cells[NomRow,5];
    if st<>'' then
      begin
      st:=Excel.Cells[NomRow,6];
      if st<>'' then
        begin
        Setlength(YP[KolYP].Discipline,KolDis+1);
        YP[KolYP].Discipline[KolDis].Name:=Excel.Cells[NomRow,1];
        YP[KolYP].Discipline[KolDis].kaf:=Excel.Cells[NomRow,5];
        YP[KolYP].Discipline[KolDis].ZE:=Excel.Cells[NomRow,6];
        YP[KolYP].Discipline[KolDis].Hour:=Excel.Cells[NomRow,7];
        YP[KolYP].Discipline[KolDis].HourBezEK:=Excel.Cells[NomRow,8];
        NomCol:=9;
        Nom:=Excel.Cells[7,NomCol];
        repeat
          st:=Excel.Cells[NomRow,NomCol];
          if st<>'' then
            YP[KolYP].Discipline[KolDis].ZESEM[Nom]:=Excel.Cells[NomRow,NomCol];
          NomCol:=NomCol+5;
          Nom:=Excel.Cells[7,NomCol];
        until Nom=1;
        repeat
          YP[KolYP].Discipline[KolDis].EkzSem[Nom]:=Excel.Cells[NomRow,NomCol];
          NomCol:=NomCol+4;
          Nom:=Excel.Cells[7,NomCol];
        until Nom=1;
        repeat
          YP[KolYP].Discipline[KolDis].KyrsSem[Nom]:=Excel.Cells[NomRow,NomCol];
          NomCol:=NomCol+4;
          Nom:=Excel.Cells[7,NomCol];
        until Nom=1;
        repeat
          st:=Excel.Cells[NomRow,NomCol];
          if st<>'' then
            YP[KolYP].Discipline[KolDis].NedelinaiaAZ[Nom]:=Excel.Cells[NomRow,NomCol];
          NomCol:=NomCol+5;
          Nom:=Excel.Cells[7,NomCol];
        until Nom=1;
        repeat
          st:=Excel.Cells[NomRow,NomCol];
          if st<>'' then
            begin
            YP[KolYP].Discipline[KolDis].AuditorObSem[Nom]:=StrToInt(Copy(St,1,Pos('/',st)-1));
            Delete(st,1,Pos('/',st));
            YP[KolYP].Discipline[KolDis].LekSem[Nom]:=StrToInt(Copy(St,1,Pos('/',st)-1));
            Delete(st,1,Pos('/',st));
            YP[KolYP].Discipline[KolDis].PraktSem[Nom]:=StrToInt(Copy(St,1,Pos('/',st)-1));
            Delete(st,1,Pos('/',st));
            YP[KolYP].Discipline[KolDis].LabSem[Nom]:=StrToInt(st);
            end;
          NomCol:=NomCol+12;
          Nom:=Excel.Cells[7,NomCol];
        until Nom=1;
        repeat
          st:=Excel.Cells[NomRow,NomCol];
          if st<>'' then
            YP[KolYP].Discipline[KolDis].SRSSem[Nom]:=Excel.Cells[NomRow,NomCol];
          NomCol:=NomCol+5;
          st:=Excel.Cells[7,NomCol];       //���
          if st<>'���' then
            Nom:=StrToInt(st);
        until st='���';
        YP[KolYP].Discipline[KolDis].SumAuditoria:=Excel.Cells[NomRow,NomCol];
        YP[KolYP].Discipline[KolDis].SumSRS:=Excel.Cells[NomRow,NomCol+1];
        end
      else
        begin
        YP[KolYP].Discipline[KolDis-1].ElektivDis:=Excel.Cells[NomRow,1];
        YP[KolYP].Discipline[KolDis-1].ElektivDisKaf:=Excel.Cells[NomRow,5];
        end;
      inc(KolDis);
      end;
    inc(NomRow);
    st:=Excel.Cells[NomRow,1];
    end;

  inc(KolYP);
  Excel.Workbooks.Close;
  FMain.MeProtocol.Lines.Add('�������� �� '+CurrentDir+'\��\'+SearchStr.Name);
  until FindNext(SearchStr)<>0;
  end;
if FileExists(CurrentDir+'\�������� ��.xlsx') then
  ExcelBase.Workbooks.Close;
end;

procedure TFMain.BtMTOSemestrovClick(Sender: TObject);
begin
LoadSemPlan;
SortSemPlan;
if not DirectoryExists(CurrentDir+'\���\��') then
    ForceDirectories(CurrentDir+'\���\��');
LoadPo(CurrentDir+'\���\��\��.xlsx');

if not DirectoryExists(CurrentDir+'\���������� ���� �� �������') then
    ForceDirectories(CurrentDir+'\���������� ���� �� �������');
LoadAllRaspisanieAllGroup(CurrentDir+'\���������� ���� �� �������\');

CreateAllGroup;
LoadMTOIASYFile(CurrentDir+'\��� � ���.xlsx');

{SaveAllMTOSemPlan(3,'��� ���');
SaveAllMTOSemPlan(4,'��� ��� ��� ���');}
AddAllAud('304');
SaveAllMTOSemPlan(1,'��� ���');
{SaveAllMTOSemPlan(2,'��� ��� ��� ���.');   }

MeProtocol.Lines.Add('���������� ��� ���������.');
end;

procedure TFMain.BtAllPOAudClick(Sender: TObject);
begin
if not DirectoryExists(CurrentDir+'\���\��') then
    ForceDirectories(CurrentDir+'\���\��');
LoadPo(CurrentDir+'\���\��\��.xlsx');

if not DirectoryExists(CurrentDir+'\���������� ���� �� �������') then
    ForceDirectories(CurrentDir+'\���������� ���� �� �������');
LoadAllRaspisanieAllGroup(CurrentDir+'\���������� ���� �� �������\');
CreateAllGroup;
CreateAllPOInAud(1);
end;

procedure TFMain.BtRaspisanieClick(Sender: TObject);
begin
if not DirectoryExists(CurrentDir+'\���������� ����') then
    ForceDirectories(CurrentDir+'\���������� ����');
ExcelBase.Workbooks.Add;
ExcelBase.Cells[1,1]:=2;
//LoadAllRaspisanieAllPrepod(CurrentDir+'\���������� �������.xlsx',1);
LoadAllRaspisanieAllPrepod(CurrentDir+'\���������� ����\',0);
SortAllPrepodDateTime;

ExcelBase.Workbooks[1].SaveAs(CurrentDir+'\����������\��� � ��������.xlsx');
MeProtocol.Lines.Add('������ ���� '+CurrentDir+'\����������\��� � ��������.xlsx');
ExcelBase.Workbooks.Close;
//GoRaspisanieToExcel(CurrentSemestr);
GoRaspisanieKafCv(CurrentSemestr,false);
//GoRaspisanieKaf(CurrentSemestr);
MeProtocol.Lines.Add('���������� �������');
end;

procedure TFMain.BtMTOClick(Sender: TObject);
begin
if not DirectoryExists(CurrentDir+'\���\��') then
    ForceDirectories(CurrentDir+'\���\��');
LoadPo(CurrentDir+'\���\��\��.xlsx');

if not DirectoryExists(CurrentDir+'\���������� ���� �� �������') then
    ForceDirectories(CurrentDir+'\���������� ���� �� �������');
LoadAllRaspisanieAllGroup(CurrentDir+'\���������� ���� �� �������\');
CreateAllGroup;
SaveAllMTO;
MeProtocol.Lines.Add('���������� ��� ���������.');
end;

procedure TFMain.BtGroupClick(Sender: TObject);
begin
CreateAllGroup;
//CreateTabledAud(CurrentDir+'\���������.xlsx');
end;

procedure TFMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
If CbAutoProverkaClose.Checked then
  FMain.BtProverkaClick(Sender);
 //��������� ���������� Excel
 Excel.Quit;
 ExcelBase.Quit;
 //������� ���������� ������
 Excel := Unassigned;
 ExcelBase := Unassigned;

end;

procedure TFMain.SgNagryzkaSearthSetEditText(Sender: TObject; ACol,
  ARow: Integer; const Value: string);
var
 NomNagryzka,MaxNomNagryzka:longword;
begin
NomNagryzka:=1;
MaxNomNagryzka:=1;
SgNagryzka.RowCount:=MaxNomNagryzka;
while NomNagryzka<Length(Nagryzka) do
  begin
  if
     ((SgNagryzkaSearth.Cells[0,1]='') or (Pos(SgNagryzkaSearth.Cells[0,1],Nagryzka[NomNagryzka].Dis)<>0)) and
     ((SgNagryzkaSearth.Cells[1,1]='') or (Pos(SgNagryzkaSearth.Cells[1,1],Nagryzka[NomNagryzka].Vid)<>0)) and
     ((SgNagryzkaSearth.Cells[2,1]='') or (Pos(SgNagryzkaSearth.Cells[2,1],Nagryzka[NomNagryzka].Group)<>0)) and
     ((SgNagryzkaSearth.Cells[4,1]='') or (Pos(SgNagryzkaSearth.Cells[4,1],Nagryzka[NomNagryzka].FIOPrep)<>0)) and
     ((SgNagryzkaSearth.Cells[3,1]='') or (Pos(SgNagryzkaSearth.Cells[3,1],Nagryzka[NomNagryzka].Hour)<>0)) then
    begin
    SgNagryzka.RowCount:=MaxNomNagryzka+1;
    SgNagryzka.Cells[0,MaxNomNagryzka]:=Nagryzka[NomNagryzka].Dis;
    SgNagryzka.Cells[1,MaxNomNagryzka]:=Nagryzka[NomNagryzka].Vid;
    SgNagryzka.Cells[2,MaxNomNagryzka]:=Nagryzka[NomNagryzka].Group;
    SgNagryzka.Cells[3,MaxNomNagryzka]:=Nagryzka[NomNagryzka].Hour;
    SgNagryzka.Cells[4,MaxNomNagryzka]:=Nagryzka[NomNagryzka].FIOPrep;
    SgNagryzka.Cells[5,MaxNomNagryzka]:=IntTostr(Nagryzka[NomNagryzka].KolStudent);
    SgNagryzka.Cells[6,MaxNomNagryzka]:=Nagryzka[NomNagryzka].Opisanie;
    SgNagryzka.Cells[7,MaxNomNagryzka]:=IntToStr(NomNagryzka);
    inc(MaxNomNagryzka);
    end;
  inc(NomNagryzka);
  end;
end;

procedure TFMain.SgNagryzkaSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
RowSelectSgNagryzka:=ARow;
end;

procedure TFMain.SgNagryzkaSetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: string);
var
k:integer;
nom:Double;
NomNagryzka:Longword;
st:string;
begin
val(SgNagryzka.Cells[ACol,ARow],nom,k);
if ((ACol=3) and (k=0)) or (ACol=4) then
  begin
  NomNagryzka:=0;  //��� ������ ���������� ����� ������� � �������� ��������.
  while (NomNagryzka<Length(Nagryzka)) and (not ((SgNagryzka.Cells[0,ARow]=Nagryzka[NomNagryzka].Dis) and
                                                   (SgNagryzka.Cells[1,ARow]=Nagryzka[NomNagryzka].Vid) and
                                                   (SgNagryzka.Cells[2,ARow]=Nagryzka[NomNagryzka].Group))) do
    Inc(NomNagryzka);
  if NomNagryzka<Length(Nagryzka) then
    begin
    Excel.Workbooks.Open(NameFileNagryzka[Nagryzka[NomNagryzka].Sem]);
    st:=Excel.Cells[Nagryzka[NomNagryzka].NomRow,ACol+1];
    Excel.Cells[Nagryzka[NomNagryzka].NomRow,ACol+1]:=SgNagryzka.Cells[ACol,ARow];
    FMain.MeProtocol.Lines.Add('�������� �������� '+Nagryzka[NomNagryzka].Dis+' '+Nagryzka[NomNagryzka].Vid+' '+Nagryzka[NomNagryzka].Group+' '+st);
    FMain.MeProtocol.Lines.Add('               �� '+Nagryzka[NomNagryzka].Dis+' '+Nagryzka[NomNagryzka].Vid+' '+Nagryzka[NomNagryzka].Group+' '+SgNagryzka.Cells[ACol,ARow]);
    Excel.Workbooks[1].Save;
    Excel.Workbooks.Close;
    end;
  end;
end;

procedure TFMain.BtRaspredelenieNagryzkiClick(Sender: TObject);
begin
FChangeNagryzka.ShowModal;
end;

procedure GridDeleteRow(RowNumber: Integer; Grid: TstringGrid);
 var
   i: Integer;
 begin
   Grid.Row := RowNumber;
   if (Grid.Row = Grid.RowCount - 1) then
     { On the last row}
     Grid.RowCount := Grid.RowCount - 1
   else
   begin
     { Not the last row}
     for i := RowNumber to Grid.RowCount - 1 do
       Grid.Rows[i] := Grid.Rows[i + 1];
     Grid.RowCount := Grid.RowCount - 1;
   end;
 end;

end.
