unit UChangeNagryzka;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, UMain, ExtCtrls, Grids;

type
  TFChangeNagryzka = class(TForm)
    Label1: TLabel;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    LADis: TLabel;
    Label3: TLabel;
    LaVid: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    LaGroup: TLabel;
    LaHour: TLabel;
    LaFioPrep: TLabel;
    LaComment: TLabel;
    LbPrep: TListBox;
    Panel4: TPanel;
    Panel5: TPanel;
    BtGoPrepodTable: TButton;
    Panel6: TPanel;
    BtNazn: TButton;
    SgNewNagryzkaPrepod: TStringGrid;
    Label7: TLabel;
    LaNomNagryzka: TLabel;
    Label8: TLabel;
    LaStudent: TLabel;
    procedure BtNaznClick(Sender: TObject);
    procedure SgNewNagryzkaPrepodSetEditText(Sender: TObject; ACol,
      ARow: Integer; const Value: string);
    procedure BtGoPrepodTableClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FChangeNagryzka: TFChangeNagryzka;


implementation

uses UNagryzka, UConstParametrs;

{$R *.dfm}
var
  NomPrepodOsnova:LongWord;
  HourAll:Double;
  NomHourOnStud:array of Longword;
  NomMerge:Longword;

procedure TFChangeNagryzka.BtNaznClick(Sender: TObject);
var
  NomRow,NomCol:Longword;
begin
//Изменение файла Excel
for NomCol := 2 to SgNewNagryzkaPrepod.ColCount - 1 do
  for NomRow := 4 to SgNewNagryzkaPrepod.RowCount - 1 do
    SearchAndAddExcelNagryzka(StrToInt(SgNewNagryzkaPrepod.Cells[NomCol,1]),SeartchPrepodFIO(SgNewNagryzkaPrepod.Cells[0,NomRow]), StrToFloat(SgNewNagryzkaPrepod.Cells[NomCol,NomRow]));

end;

procedure TFChangeNagryzka.BtGoPrepodTableClick(Sender: TObject);
Var
  ColRow,i:Longword;
  NomPrepod:Longword;
  NomSelectedLb:Longword;

begin
ColRow:=SgNewNagryzkaPrepod.RowCount;
SgNewNagryzkaPrepod.RowCount:=ColRow+1;
NomSelectedLb:=0;
while (NomSelectedLb<lbPrep.Items.Count) and (not lbPrep.Selected[NomSelectedLb]) do
  inc(NomSelectedLb);
NomPrepod:=0;
while (NomPrepod<Length(Prepod)) and (Prepod[NomPrepod].FIO<>lbPrep.Items.Strings[NomSelectedLb]) do
  inc(NomPrepod);
SgNewNagryzkaPrepod.Cells[0,ColRow]:=Prepod[NomPrepod].FIO;
for i := 1 to SgNewNagryzkaPrepod.ColCount-1 do
  SgNewNagryzkaPrepod.Cells[i,ColRow]:='0';

end;

procedure TFChangeNagryzka.FormActivate(Sender: TObject);
var
  NomPrepod,NomNagryzka,NomMergeDis,NomHourOnOneStudent,NomCol:Longword;
  ExitWhile:boolean;
begin
//Левая текстовая часть загружается из таблицы по выбранной строке
LADis.Caption:=FMain.SgNagryzka.Cells[0,RowSelectSgNagryzka];
LaVid.Caption:=FMain.SgNagryzka.Cells[1,RowSelectSgNagryzka];
LaGroup.Caption:=FMain.SgNagryzka.Cells[2,RowSelectSgNagryzka];
LaHour.Caption:=FMain.SgNagryzka.Cells[3,RowSelectSgNagryzka];
LaStudent.Caption:=FMain.SgNagryzka.Cells[5,RowSelectSgNagryzka];
HourAll:=StrToFloat(LaHour.Caption);
LaFioPrep.Caption:=FMain.SgNagryzka.Cells[4,RowSelectSgNagryzka];
LaComment.Caption:=FMain.SgNagryzka.Cells[6,RowSelectSgNagryzka];
LaNomNagryzka.Caption:=FMain.SgNagryzka.Cells[7,RowSelectSgNagryzka];

//Добавление всех преподавателей в список
LbPrep.Items.Clear;
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  if (Prepod[NomPrepod].FIO<>LaFioPrep.Caption) and (LbPrep.Items.IndexOf(Prepod[NomPrepod].FIO)=-1) then
    LbPrep.Items.Add(Prepod[NomPrepod].FIO)
  else
    NomPrepodOsnova:=NomPrepod;     //Запомнить номер преподавателя из списка преподов
  inc(NomPrepod);
  end;
if LbPrep.Items.IndexOf('не назначено')=-1 then
  LbPrep.Items.Add('не назначено');

//Заполнить таблицу распределения нагрузки по текущим данным
SgNewNagryzkaPrepod.RowCount:=5;
SgNewNagryzkaPrepod.ColWidths[0]:=300;
SgNewNagryzkaPrepod.ColWidths[1]:=24;
SgNewNagryzkaPrepod.Cells[0,1]:='Номер дисциплины';
SgNewNagryzkaPrepod.Cells[0,2]:='Часов на студента';
SgNewNagryzkaPrepod.Cells[0,3]:='Всего часов';
SgNewNagryzkaPrepod.Cells[1,0]:='Ст.';

SgNewNagryzkaPrepod.Cells[0,4]:=LaFioPrep.Caption;
SgNewNagryzkaPrepod.Cells[1,3]:=LaStudent.Caption;
SgNewNagryzkaPrepod.Cells[1,4]:=LaStudent.Caption;
//SgNewNagryzkaPrepod.Cells[2,1]:=LaHour.Caption;
//SgNewNagryzkaPrepod.Cells[1,1]:=FloatToStr(Prepod[NomPrepodOsnova].AllHour);

//Найти все сгруппированные дисциплины с этой
ExitWhile:=false;
NomMerge:=0;
while (NomMerge<Length(ArrMergeDis)) and (not ExitWhile) do  //Проход по группировкам
  begin
  NomMergeDis:=0;
  while (NomMergeDis<Length(ArrMergeDis[NomMerge])) and not((ArrMergeDis[NomMerge][NomMergeDis].Dis=LADis.Caption) and (ArrMergeDis[NomMerge][NomMergeDis].Vid=LaVid.Caption)) do  //Проход внутри группы
    inc(NomMergeDis);
  if (ArrMergeDis[NomMerge][NomMergeDis].Dis=LADis.Caption) and (ArrMergeDis[NomMerge][NomMergeDis].Vid=LaVid.Caption) then
    ExitWhile:=true //Если нашли совпадающую дисциплину, то Запоминаем гурппу дисциплин
  else
    inc(NomMerge);
  end;

NomCol:=1;
If NomMerge<Length(ArrMergeDis) then
  begin
  //Проходим по группе дисциплин
  NomMergeDis:=0;
  while (NomMergeDis<Length(ArrMergeDis[NomMerge])) do
    begin
    inc(NomCol);
    SgNewNagryzkaPrepod.ColCount:=NomCol+1;
    SgNewNagryzkaPrepod.ColWidths[NomCol]:=140;
    SgNewNagryzkaPrepod.Cells[NomCol,0]:=ArrMergeDis[NomMerge][NomMergeDis].Dis;
    //Каждую дисциплину заносим в таблицу и определяем количество часов
    NomHourOnOneStudent:=0;
    while NomHourOnOneStudent<Length(HourOnOneStudent) do
      begin
      if HourOnOneStudent[NomHourOnOneStudent].Vid=ArrMergeDis[NomMerge][NomMergeDis].Vid then
        begin
        SgNewNagryzkaPrepod.Cells[NomCol,2]:=FloatToStr(HourOnOneStudent[NomHourOnOneStudent].Hour);  //Для каждой дисциплины запомнить количество часов на студента
//        SgNewNagryzkaPrepod.Cells[NomCol,4]:=FloatToStr(HourOnOneStudent[NomHourOnOneStudent].Hour*StrToFloat(LaStudent.Caption));
        end;
      inc(NomHourOnOneStudent);
      end;
    NomNagryzka:=0;  //Для каждой дисциплины найти строчку в основной нагрузке.
    while (NomNagryzka<Length(Nagryzka)) and (not ((ArrMergeDis[NomMerge][NomMergeDis].Dis=Nagryzka[NomNagryzka].Dis) and
                                                   (ArrMergeDis[NomMerge][NomMergeDis].Vid=Nagryzka[NomNagryzka].Vid) and
                                                   (LaGroup.Caption=Nagryzka[NomNagryzka].Group) and
                                                   (LaFioPrep.Caption=Nagryzka[NomNagryzka].FIOPrep))) do
      Inc(NomNagryzka);
    if NomNagryzka<Length(Nagryzka) then
      begin
      SgNewNagryzkaPrepod.Cells[NomCol,1]:=IntToStr(NomNagryzka);
      SgNewNagryzkaPrepod.Cells[NomCol,3]:=Nagryzka[NomNagryzka].Hour;
      SgNewNagryzkaPrepod.Cells[NomCol,4]:=Nagryzka[NomNagryzka].Hour;
      end;
    inc(NomMergeDis);
    end;


  end;

end;

procedure TFChangeNagryzka.SgNewNagryzkaPrepodSetEditText(Sender: TObject; ACol,
  ARow: Integer; const Value: string);
var
  NomRow,NomCol:Longword;
  SumHour:Double;
  k:integer;
  nom:Double;
  Proverka:Boolean;
begin
val(Value,nom,k);
if k=0 then
begin

if ACol=1 then  //Изменение количества человек
  for NomCol := 2 to SgNewNagryzkaPrepod.ColCount-1 do
    begin
    val(SgNewNagryzkaPrepod.Cells[NomCol,2],nom,k);
    if k=0 then
      begin
      val(SgNewNagryzkaPrepod.Cells[1,ARow],nom,k);
      if k=0 then
        SgNewNagryzkaPrepod.Cells[NomCol,ARow]:=FloatToStr(StrToFloat(SgNewNagryzkaPrepod.Cells[NomCol,2])*StrToFloat(SgNewNagryzkaPrepod.Cells[1,ARow]));
      end;
    end;

for NomCol := 2 to SgNewNagryzkaPrepod.ColCount-1 do
begin
SumHour:=0;
for NomRow := 5 to SgNewNagryzkaPrepod.RowCount-1 do
    begin
    val(SgNewNagryzkaPrepod.Cells[NomCol,NomRow],nom,k);
    if k=0 then
      SumHour:=SumHour+StrToFloat(SgNewNagryzkaPrepod.Cells[NomCol,NomRow]);
    end;
val(SgNewNagryzkaPrepod.Cells[NomCol,3],nom,k);
if (k=0) and (StrToFloat(SgNewNagryzkaPrepod.Cells[NomCol,3])-SumHour>=0) then
  SgNewNagryzkaPrepod.Cells[NomCol,4]:=FloatToStr(StrToFloat(SgNewNagryzkaPrepod.Cells[NomCol,3])-SumHour);
end;
// Проверка возможности такого распределения нагрузки
Proverka:=true;
for NomCol := 2 to SgNewNagryzkaPrepod.ColCount-1 do
  begin
  SumHour:=0;
  for NomRow := 4 to SgNewNagryzkaPrepod.RowCount-1 do
    begin
    val(SgNewNagryzkaPrepod.Cells[ACol,NomRow],nom,k);
    if k=0 then
      SumHour:=SumHour+StrToFloat(SgNewNagryzkaPrepod.Cells[ACol,NomRow]);
    end;
  val(SgNewNagryzkaPrepod.Cells[ACol,3],nom,k);
  if (k=0) then
    begin
    if (SumHour<>StrToFloat(SgNewNagryzkaPrepod.Cells[ACol,3])) then
      begin
      Proverka:=false
      end
    else
      begin

      end;
    end;
  end;
//Nazn.Enabled:=Proverka;
end;
end;

end.
