unit ULoadExcel;

interface

uses SysUtils, StdCtrls, Grids, UMain, UNagryzka, UGroup, UConstParametrs, UAuditoria, USemPlan;

Procedure LoadInitializationParametrsInExcelFile (FileName:String);
Procedure ProverkaStart;
Procedure ProverkaFilePrep (Dir:String; Me:TMemo);
Procedure VivodPrepodSG (SG:TStringGrid);
Procedure VivodAllNagryzkaSG(Sg:TStringGrid);
Procedure VivodSgExcel (Sg:TStringGrid; FileName:String);
Procedure StartLoadExcel(MeProtocol:TMemo);
procedure LoadGroup(NameFile:string);
Procedure LoadExcelNagr(FileName:String; TypeSem:Byte);
Procedure LoadAuditorii(FileName:string);
Procedure LoadHourStudentDisFromExcelFile(NameFile:string);
Procedure LoadGroupDisFromExcelFile(NameFile:string);
Procedure LoadPrepodFromExcelFile(NameFile:String);
Procedure LoadAllRaspisanieAllGroup(RaspisanieDir:string);
Procedure LoadAllRaspisanieAllPrepod(RaspisanieDir:string; TypePrepod:byte);
Procedure LoadRaspIASYEkzamenGroupExcel(FileName:String);
Procedure LoadRaspIASYEkzamenExcel(FileName:String);

implementation

uses USaveExcel;

Procedure LoadInitializationParametrsInExcelFile (FileName:String);
var
i:longword;
St:string;
begin
Excel.Workbooks.Open(FileName);
SetLength(TimeSetPar,0);
i:=0;
St:=Excel.Cells[10,2+i];
while st<>'' do
  begin
  SetLength(TimeSetPar,i+1);
  TimeSetPar[i]:=Excel.Cells[10,2+i];
  inc(i);
  St:=Excel.Cells[10,2+i];
  end;
SetLength(HourOnOneStudent,0);
i:=0;
St:=Excel.Cells[12,2+i];
while st<>'' do
  begin
  SetLength(HourOnOneStudent,i+1);
  HourOnOneStudent[i].Vid:=Excel.Cells[12,2+i];
  HourOnOneStudent[i].Hour:=Excel.Cells[13,2+i];
  inc(i);
  St:=Excel.Cells[12,2+i];
  end;
CurrentYear:=Excel.Cells[2,2];
CurrentSemestr:=Excel.Cells[3,2];
YearBakalavr:=Excel.Cells[5,2];
YearAspirant:=Excel.Cells[5,4];
YearMagistr:=Excel.Cells[5,3];
HourStavka:=Excel.Cells[6,2];
ZKaf:=Excel.Cells[8,2];
ZKafSokr:=Excel.Cells[8,3];
NomKaf:=Excel.Cells[7,2];
CreateFilePrep:=false;
Excel.Workbooks.Close;
end;

Procedure ProverkaStart;
var
  SR: TSearchRec;   // поисковая переменная
  FindRes: Integer; // переменная для записи результата поиска
  StErr:string;
  NextGo:Boolean;
  NomPrepod:Longword;
begin
With FMain do
begin
DeatroyAllDate;

CurrentDir := GetCurrentDir;
NextGo:=false;
// задание условий поиска и начало поиска
FindRes := FindFirst(CurrentDir+'\Нагрузка*.xls*', faAnyFile, SR);
StErr:='';

if FindRes = 0 then // Если нашли файл
  begin
  LaNameFile.Caption:=CurrentDir+'\';
  if not DirectoryExists(CurrentDir+'\Нагрузка по преподавателям') then
    ForceDirectories(CurrentDir+'\Нагрузка по преподавателям');
   //Поиск файла с осенней нагрузкой
   FindRes := FindFirst(CurrentDir+'\Нагрузка_осень*.xls*', faAnyFile, SR);
   if FindRes = 0 then // Если нашли файл
     begin
     FMain.PNagrO.Color:=ClGreen;
     NameFileNagryzka[1]:=CurrentDir+'\'+SR.Name;
     MeProtocol.Lines.Add('Загрузка данных из файла:'+CurrentDir+'\'+SR.Name);
     LoadExcelNagr(CurrentDir+'\'+SR.Name,1);
     NextGo:=true;
     end
   else
     MeProtocol.Lines.Add('Не найден файл с осенней нагрузкой:'+CurrentDir+'\Нагрузка_осень*.xls*');
   //Поиск файла с осенней нагрузкой
   FindRes := FindFirst(CurrentDir+'\Нагрузка_весна*.xls*', faAnyFile, SR);
   if FindRes = 0 then // Если нашли файл
     begin
     FMain.PNagrV.Color:=ClGreen;
     NameFileNagryzka[2]:=CurrentDir+'\'+SR.Name;
     MeProtocol.Lines.Add('Загрузка данных из файла:'+CurrentDir+'\'+SR.Name);
     LoadExcelNagr(CurrentDir+'\'+SR.Name,2);
     NextGo:=true;
     end
   else
     MeProtocol.Lines.Add('Не найден файл с весенней нагрузкой:'+CurrentDir+'\Нагрузка_весна*.xls*');
   end
 else
   LaNameFile.Caption:='Необходимо определение';

if NextGo then
  begin
  if not CreateFilePrep then
    CreatePrep
  else
    ProverkaFilePrep (CurrentDir+'\Нагрузка по преподавателям',MeProtocol);
  SortPrepodFIO;
  For NomPrepod:=0 to Length(Prepod)-1 do
    Prepod[NomPrepod].FlagP:=0;
  VivodPrepodSG (SGPrepod);
  if FileExists(CurrentDir+'\Нагрузка_таблица.xlsx') then
    begin
    MeProtocol.Lines.Add('Сохранение нагрузки в файл:'+CurrentDir+'\Нагрузка_таблица.xlsx');
    VivodSgExcel(SGPrepod,CurrentDir+'\Нагрузка_таблица.xlsx');
    end
  else
    MeProtocol.Lines.Add('Не найден файл для таблицы нагрузки:'+CurrentDir+'\Нагрузка_таблица.xlsx');
  ProverkaTsel;
  VivodOshibkiBase(MeProtocol);
  end;

end;
end;

Procedure ProverkaFilePrep (Dir:String; Me:TMemo);
var
  NomPrepod,NomNagr,NomNagryzkaS,NomRow,NomHourStudentDis:Longword;
  Sem:Byte;
  SR: TSearchRec;   // поисковая переменная
  FindRes: Integer; // переменная для записи результата поиска
  st,StGroup,st1,NameFileXlSX:string;
  StExcel,StSokr:String;
  HourNagr:Double;
  HourSem:array [1..kolsem] of Double;
  NomGroupNagryzka:Longword;
begin
NomPrepod:=0;
while (NomPrepod<Length(Prepod))  do
  begin
  // задание условий поиска и начало поиска
  st1:=Prepod[NomPrepod].FIO;
  st:=Copy(st1,1,Pos(' ',st1)-1);
  if Pos(' ',st1)<>0 then
    Delete(st1,1,Pos(' ',st1));
  if Length(st1)<>0 then
    begin
    st:=st+'_'+st1[1];
    if Pos(' ',st1)<>0 then
      Delete(st1,1,Pos(' ',st1));
    if Length(st1)<>0 then
      st:=st+st1[1];
    end;

  FindRes := FindFirst(Dir+'\'+st+'*.xls*', faAnyFile, SR);
  if FindRes <> 0 then // Если нашли файл
    begin
    NameFileXlSX:=Dir+'\'+st+'.xlsx';
    Prepod[NomPrepod].NameFilePrepod:=NameFileXlSX;
    Excel.WorkBooks.Add;
    HourNagr:=0;
    NomRow:=2;
    NomNagryzkaS:=0;
    Excel.Cells[1,1]:=Prepod[NomPrepod].FIO;
    for Sem := 1 to kolsem do
      begin
      NomNagr:=0;
      HourSem[Sem]:=0;
      while NomNagr<Length(Nagryzka) do
        begin
        if (Nagryzka[NomNagr].FIOPrep=Prepod[NomPrepod].FIO) and (Nagryzka[NomNagr].Sem=Sem) then
          begin

          Excel.Cells[NomRow,1]:=Nagryzka[NomNagr].Dis;
          Excel.Cells[NomRow,2]:=Nagryzka[NomNagr].Vid;
          Excel.Cells[NomRow,3]:=Nagryzka[NomNagr].Group;
          Excel.Cells[NomRow,4]:=Nagryzka[NomNagr].Hour;
          Excel.Cells[NomRow,5]:=Nagryzka[NomNagr].Opisanie;
          Excel.Cells[NomRow,6]:=Nagryzka[NomNagr].NOMPrep;

          SetLength(Prepod[NomPrepod].Nagryzka,NomNagryzkaS+1);
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS]:=TNagryzkaPrepod.Create;
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Prepod:=Prepod[NomPrepod];
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].P:=0;
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].sem:=Sem;
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].NomRow:=NomRow;
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Dis:=Nagryzka[NomNagr].Dis;
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Vid:=Nagryzka[NomNagr].Vid;
          //

          StGroup:=Nagryzka[NomNagr].Group ;
          NomGroupNagryzka:=0;
          while Pos(',',StGroup)<>0 do
             begin
             SetLength(Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Group,NomGroupNagryzka+1);
             Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Group[NomGroupNagryzka]:=SearchAndCreateGroup(Copy(StGroup,1,Pos(',',StGroup)-1));
             Delete(StGroup,1,Pos(',',StGroup)+1);
             inc(NomGroupNagryzka);
             end;
           SetLength(Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Group,NomGroupNagryzka+1);
           Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Group[NomGroupNagryzka]:=SearchAndCreateGroup(StGroup);

          NomberOfDis(NomPrepod,NomNagryzkaS);

          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Hour:=Nagryzka[NomNagr].Hour;
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].KolStudent:=0;
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Opisanie:=Nagryzka[NomNagr].Opisanie;
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].NOMPrep:=Nagryzka[NomNagr].NOMPrep;

          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].SOKR:=Nagryzka[NomNagr].SOKR;

          HourNagr:=HourNagr+StrToFloat(Nagryzka[NomNagr].Hour);
          HourSem[Sem]:=HourSem[Sem]+StrToFloat(Nagryzka[NomNagr].Hour);

          inc(NomNagryzkaS);
          inc(NomRow);
          end;
        inc(NomNagr);
        end;
      Prepod[NomPrepod].HourSem[Sem]:=HourSem[Sem];
      inc(NomRow);
      end;
    inc(NomRow);
    Excel.Cells[NomRow,1]:='ИТОГО';
    Excel.Cells[NomRow,4]:=HourNagr;
    Prepod[NomPrepod].AllHour:=HourNagr;
    Excel.Workbooks[1].saveas(NameFileXlSX);
    Excel.Workbooks.Close;
    Me.Lines.Add('Создан файл преподавателей '+Dir+'\'+st+'.xlsx');
    end
  else
    begin
    Me.Lines.Add('Открыт файл преподавателей '+Dir+'\'+SR.Name);
    Excel.Workbooks.Open(Dir+'\'+SR.Name);
    Prepod[NomPrepod].NameFilePrepod:=Dir+'\'+SR.Name;
    HourNagr:=0;
    NomRow:=2;
    NomNagryzkaS:=0;
    For Sem:=1 to kolsem do
      begin
      HourSem[Sem]:=0;
      StExcel:=Excel.Cells[NomRow,1];
      while StExcel<>'' do
        begin
        SetLength(Prepod[NomPrepod].Nagryzka,NomNagryzkaS+1);
        Prepod[NomPrepod].Nagryzka[NomNagryzkaS]:=TNagryzkaPrepod.Create;
        Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Prepod:=Prepod[NomPrepod];
        Prepod[NomPrepod].Nagryzka[NomNagryzkaS].P:=0;
        Prepod[NomPrepod].Nagryzka[NomNagryzkaS].sem:=Sem;
        Prepod[NomPrepod].Nagryzka[NomNagryzkaS].NomRow:=NomRow;
        Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Dis:=Excel.Cells[NomRow,1];
        Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Vid:=Excel.Cells[NomRow,2];
        NomberOfDis(NomPrepod,NomNagryzkaS);
        NomGroupNagryzka:=0;
        StGroup:=Excel.Cells[NomRow,3];
        while pos(',',StGroup)<>0 do
          begin
          SetLength(Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Group,NomGroupNagryzka+1);
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Group[NomGroupNagryzka]:=SearchAndCreateGroup(copy(StGroup,1,pos(',',StGroup)-1));
          inc(NomGroupNagryzka);
          delete(StGroup,1,pos(',',StGroup)+1);
          end;
        SetLength(Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Group,NomGroupNagryzka+1);
        Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Group[NomGroupNagryzka]:=SearchAndCreateGroup(StGroup);
        Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Hour:=Excel.Cells[NomRow,4];
        Prepod[NomPrepod].Nagryzka[NomNagryzkaS].KolStudent:=0;
        Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Opisanie:=Excel.Cells[NomRow,5];
        Prepod[NomPrepod].Nagryzka[NomNagryzkaS].NOMPrep:=Excel.Cells[NomRow,6];
        StSokr:=Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Dis;
        Prepod[NomPrepod].Nagryzka[NomNagryzkaS].SOKR:=upCase(StSokr[1]);
        while pos(' ',StSokr)<>0 do
          begin
          Delete(StSokr,1,pos(' ',StSokr));
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].SOKR:=Prepod[NomPrepod].Nagryzka[NomNagryzkaS].SOKR+UpCase(StSokr[1]);
          end;
        HourNagr:=HourNagr+StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Hour);
        HourSem[Sem]:=HourSem[Sem]+StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Hour);

        inc(NomNagryzkaS);
        inc(NomRow);
        StExcel:=Excel.Cells[NomRow,1];
        end;
      inc(NomRow);
      Prepod[NomPrepod].HourSem[Sem]:=HourSem[Sem];
      end;
    Prepod[NomPrepod].AllHour:=HourNagr;
    Excel.Workbooks.Close;
    end;
  inc(NomPrepod);
  end;
end;

Procedure VivodPrepodSG (SG:TStringGrid);
var
NomPrepod:Longword;
begin
Sg.ColCount:=5;
SG.ColWidths[0]:=200;
SG.ColWidths[1]:=40;
SG.ColWidths[2]:=40;
SG.ColWidths[3]:=40;
SG.ColWidths[4]:=60;
SG.Cells[0,0]:='ФИО';
SG.Cells[1,0]:='Всего';
SG.Cells[2,0]:='Осень';
SG.Cells[3,0]:='Весна';
SG.Cells[4,0]:='Ставка';


NomPrepod:=0;
while (NomPrepod<Length(Prepod))  do
  begin
  Sg.RowCount:=NomPrepod+2;
  SG.Cells[0,NomPrepod+1]:=Prepod[NomPrepod].FIO;
  SG.Cells[1,NomPrepod+1]:=FloatToStr(Prepod[NomPrepod].AllHour);
  SG.Cells[2,NomPrepod+1]:=FloatToStr(Prepod[NomPrepod].HourSem[1]);
  SG.Cells[3,NomPrepod+1]:=FloatToStr(Prepod[NomPrepod].HourSem[2]);
  SG.Cells[4,NomPrepod+1]:=FloatToStr(Prepod[NomPrepod].AllHour/HourStavka);
  inc(NomPrepod);
  end;
end;

Procedure VivodShNagryzkaSG(Sg:TStringGrid);
begin
Sg.RowCount:=2;
Sg.ColCount:=8;
Sg.ColWidths[0]:=360;
Sg.ColWidths[1]:=80;
Sg.ColWidths[2]:=200;
Sg.ColWidths[3]:=20;
Sg.ColWidths[4]:=200;
Sg.ColWidths[5]:=20;
Sg.ColWidths[6]:=200;
Sg.ColWidths[7]:=20;
Sg.Cells[0,0]:='Дисциплина';
Sg.Cells[1,0]:='Вид нагрузки';
Sg.Cells[2,0]:='Группы';
Sg.Cells[3,0]:='Ч';
Sg.Cells[4,0]:='ФИО Преподавателя';
Sg.Cells[5,0]:='Ст';
Sg.Cells[6,0]:='Коментарии';
Sg.Cells[7,0]:='№';
end;

//Процедура вывода всей нагрузки в StringGrid
Procedure VivodAllNagryzkaSG(Sg:TStringGrid);
var
  NomNagryzka,MaxNomNagryzka:Longword;
begin
VivodShNagryzkaSG(Sg);
NomNagryzka:=1;
MaxNomNagryzka:=Length(Nagryzka);
Sg.RowCount:=MaxNomNagryzka;
while NomNagryzka<MaxNomNagryzka do
  begin
  Sg.Cells[0,NomNagryzka]:=Nagryzka[NomNagryzka].Dis;
  Sg.Cells[1,NomNagryzka]:=Nagryzka[NomNagryzka].Vid;
  Sg.Cells[2,NomNagryzka]:=Nagryzka[NomNagryzka].Group;
  Sg.Cells[3,NomNagryzka]:=Nagryzka[NomNagryzka].Hour;
  Sg.Cells[4,NomNagryzka]:=Nagryzka[NomNagryzka].FIOPrep;
  Sg.Cells[5,NomNagryzka]:=IntTostr(Nagryzka[NomNagryzka].KolStudent);
  Sg.Cells[6,NomNagryzka]:=Nagryzka[NomNagryzka].Opisanie;
  Sg.Cells[7,NomNagryzka]:=IntToStr(NomNagryzka);
  inc(NomNagryzka);
  end;
end;

//Процедура вывода значений из StringGrid в EXCEL файл
Procedure VivodSgExcel (Sg:TStringGrid; FileName:String);
var
  NomCol,NomRow:Longword;
begin
Excel.Workbooks.Open(FileName);
for NomCol := 0 to Sg.ColCount do
  for NomRow := 0 to Sg.RowCount do
    Excel.Cells[NomRow+1,NomCol+1]:=Sg.Cells[NomCol,NomRow];
Excel.Workbooks[1].Save;
Excel.Workbooks.Close;
end;

Procedure StartLoadExcel(MeProtocol:TMemo);
var
  SR: TSearchRec;   // поисковая переменная
  FindRes: Integer; // переменная для записи результата поиска
  SearchStr:TSearchRec;
  StErr:string;
  NextGo:Boolean;
begin
if FileExists(CurrentDir+'\аудитории оснащение.xlsx') then
  begin
  FMain.POsn.Color:=ClGreen;
  LoadAuditorii(CurrentDir+'\аудитории оснащение.xlsx');
  MeProtocol.Lines.Add('Загружен файл с информацией об аудиториях:'+CurrentDir+'\аудитории оснащение.xlsx');
  end
else
  MeProtocol.Lines.Add('Не найден файл с информацией об аудиториях:'+CurrentDir+'\аудитории оснащение.xlsx');

if FileExists(CurrentDir+'\Группы.xlsx') then
  begin
  FMain.PGroup.Color:=ClGreen;
  LoadGroup(CurrentDir+'\Группы.xlsx');
  FMain.MeProtocol.Lines.Add('Загружена информация о группах из файла:'+CurrentDir+'\Группы.xlsx');
end
else
  MeProtocol.Lines.Add('Не найден файл с информацией о группах:'+CurrentDir+'\Группы.xlsx');
{if FileExists(CurrentDir+'\ЧАСЫ НА ОДНОГО СТУДЕНТА.xlsx') then
  begin
  LoadHourStudentDisFromExcelFile(CurrentDir+'\ЧАСЫ НА ОДНОГО СТУДЕНТА.xlsx');
  MeProtocol.Lines.Add('Загружен файл с распределением студент/часы:'+CurrentDir+'\ЧАСЫ НА ОДНОГО СТУДЕНТА.xlsx');
  end
else
  MeProtocol.Lines.Add('Не найден файл с распределением студент/часы:'+CurrentDir+'\ЧАСЫ НА ОДНОГО СТУДЕНТА.xlsx'); }
if FileExists(CurrentDir+'\ОБЪЕДИНЕНИЕ ДИСЦИПЛИН.xlsx') then
  begin
  LoadGroupDisFromExcelFile(CurrentDir+'\ОБЪЕДИНЕНИЕ ДИСЦИПЛИН.xlsx');
  FMain.PMergeDis.Color:=ClGreen;
  MeProtocol.Lines.Add('Загружен файл с группами дисциплин:'+CurrentDir+'\ОБЪЕДИНЕНИЕ ДИСЦИПЛИН.xlsx');
  end
else
  MeProtocol.Lines.Add('Не найден файл с группами дисциплин:'+CurrentDir+'\ОБЪЕДИНЕНИЕ ДИСЦИПЛИН.xlsx');

ProverkaStart;
GoKolStudentOnCost;
if FileExists(CurrentDir+'\ПРЕПОДАВАТЕЛИ.xlsx') then
  begin
  FMain.PPrepod.Color:=ClGreen;
  AddAllPrepodNagryzkaToExcelFile(CurrentDir+'\ПРЕПОДАВАТЕЛИ.xlsx');
  LoadPrepodFromExcelFile(CurrentDir+'\ПРЕПОДАВАТЕЛИ.xlsx');
  MeProtocol.Lines.Add('Загружен файл с информацией о преподавателях:'+CurrentDir+'\ПРЕПОДАВАТЕЛИ.xlsx');
  end
else
  MeProtocol.Lines.Add('Не найден файл с информацией о преподавателях:'+CurrentDir+'\ПРЕПОДАВАТЕЛИ.xlsx');
{
if FileExists(CurrentDir+'\Экзамены\Экзамены_ИАСУ_осень.xlsx') then
  begin
  LoadRaspIASYEkzamenExcel(CurrentDir+'\Экзамены\Экзамены_ИАСУ_осень.xlsx');
  MeProtocol.Lines.Add('Загружен файл с информацией о экзаменах:'+CurrentDir+'\Экзамены\Экзамены_ИАСУ_осень.xlsx');
  end
else
  MeProtocol.Lines.Add('Не найден файл с информацией о экзаменах:'+CurrentDir+'\Экзамены\Экзамены_ИАСУ_осень.xlsx');
}
if FileExists(CurrentDir+'\Экзамены по группам\Расписание сессии.xls') then
  begin
  FMain.PEkz.Color:=ClGreen;
  if FindFirst(CurrentDir+'\Экзамены по группам\Расписание сессии*.xls',faDirectory,SearchStr)=0 then
  begin
  repeat
    LoadRaspIASYEkzamenGroupExcel(CurrentDir+'\Экзамены по группам\'+SearchStr.Name);
    MeProtocol.Lines.Add('Загружен файл с информацией о экзаменах:'+CurrentDir+'\Экзамены по группам\'+SearchStr.Name);
  until FindNext(SearchStr)<>0;
  end;
  end
else
  MeProtocol.Lines.Add('Не найден файл с информацией о экзаменах:'+CurrentDir+'\Экзамены по группам\Расписание сессии.xls');
if FileExists(CurrentDir+'\ЭКЗАМЕНЫ КАФЕДРА.xlsx') then
  begin
  FMain.PEkz.Color:=ClGreen;
  LoadRaspIASYEkzamenExcel(CurrentDir+'\ЭКЗАМЕНЫ КАФЕДРА.xlsx');
  MeProtocol.Lines.Add('Загружен файл с информацией о преподавателях:'+CurrentDir+'\ЭКЗАМЕНЫ КАФЕДРА.xlsx');
  end
else
  MeProtocol.Lines.Add('Не найден файл с информацией о преподавателях:'+CurrentDir+'\ЭКЗАМЕНЫ КАФЕДРА.xlsx');

VivodAllNagryzkaSG(FMain.SgNagryzka);
VivodShNagryzkaSG(FMain.SgNagryzkaSearth);
end;

procedure LoadGroup(NameFile:string);
var
NomRow:Longword;
Group:TGroup;
Kyrs:Longword;
st,stbyf:string;
begin

  Excel.Workbooks.Open(NameFile);
  NomRow:=3;
  st:=Excel.Cells[NomRow,1];
  while st<>'' do
    begin
    //Добавить недостающую цифру
    stbyf:=Copy(st,Length(st)-1,2);
    Kyrs:=CurrentYear-StrToInt('20'+stbyf);
    if CurrentSemestr=1 then
      Kyrs:=Kyrs+1;
    stbyf:=Copy(st,1,2);
    Delete(st,1,2);
    stbyf:=stbyf+Copy(st,1,Pos('-',st));
    Delete(st,1,pos('-',st));
    if (st[1]<>'Д') and(st[1]<>'З') then
      st:=StByf+IntToStr(Kyrs)+st
    else
      st:=StByf+st;
    Group:=SearchAndCreateGroup(st);
    Group.Kyrs:=Kyrs;
    Group.Plosh:=Excel.Cells[NomRow,2];
    //Убрать слово
    st:=Excel.Cells[NomRow,4];
    Group.kaf:=Copy(st,Pos(' ',st)+1,length(st)-Pos(' ',st));
    Group.Forma:=Excel.Cells[NomRow,5];
    Group.KolStudent:=Excel.Cells[NomRow,6];
    Group.Napravlenie:=Excel.Cells[NomRow,8];
    Group.ShifrNapravlenie:=Excel.Cells[NomRow,7];
    Group.Profil:=Excel.Cells[NomRow,10];
    Group.ShifrProfil:=Excel.Cells[NomRow,9];

    inc(NomRow);
    st:=Excel.Cells[NomRow,1];
    end;
  Excel.Workbooks.Close;

end;

Procedure LoadExcelNagr(FileName:String; TypeSem:Byte);
var
NomRow,NomNagryzka,NomHourStudentDis,NomGroupDis,NomDisInGroup:Longword;
st,StSokr:string;
NomPrepod,KolPrepod:Longword;
begin
if FileExists(FileName) then
begin
Excel.Workbooks.Open(FileName);
NomRow:=2;
NomNagryzka:=Length(Nagryzka);
st:=Excel.Cells[1,NomRow];
KolPrepod:=Length(Prepod);
While st<>'' do
  begin

  SetLength(Nagryzka,NomNagryzka+1);
  Nagryzka[NomNagryzka].P:=0;
  Nagryzka[NomNagryzka].NomRow:=NomRow;
  Nagryzka[NomNagryzka].Sem:=TypeSem;
  Nagryzka[NomNagryzka].Dis:=Excel.Cells[NomRow,1];
  Nagryzka[NomNagryzka].Vid:=Excel.Cells[NomRow,2];
  Nagryzka[NomNagryzka].Group:=Excel.Cells[NomRow,3];
  Nagryzka[NomNagryzka].Hour:=Excel.Cells[NomRow,4];
  Nagryzka[NomNagryzka].KolStudent:=0;
  Nagryzka[NomNagryzka].FIOPrep:=Excel.Cells[NomRow,5];
  if Nagryzka[NomNagryzka].FIOPrep='' then
    Nagryzka[NomNagryzka].FIOPrep:='не назначено';
  Nagryzka[NomNagryzka].Opisanie:=Excel.Cells[NomRow,6];
  Nagryzka[NomNagryzka].NOMPrep:=Excel.Cells[NomRow,7];
  StSokr:=Nagryzka[NomNagryzka].Dis;
  Nagryzka[NomNagryzka].SOKR:=UpCase(StSokr[1]);
  while (Length(StSokr)<>0) and (pos(' ',StSokr)<>0) do
    begin
    Delete(StSokr,1,pos(' ',StSokr));
    if (Length(StSokr)<>0) then
    Nagryzka[NomNagryzka].SOKR:=Nagryzka[NomNagryzka].SOKR+UpCase(StSokr[1]);
    end;
  NomPrepod:=0;
  while (NomPrepod<Length(Prepod)) and (Prepod[NomPrepod].FIO<>Nagryzka[NomNagryzka].FIOPrep) do
    inc(NomPrepod);
  if not (NomPrepod<Length(Prepod)) or (Length(Prepod)=0) then
    begin
    inc(KolPrepod);
    SetLength(Prepod,KolPrepod);
    Prepod[KolPrepod-1]:=TPrepodAll.Create;
    Prepod[KolPrepod-1].FIO:=Nagryzka[NomNagryzka].FIOPrep;
    Prepod[KolPrepod-1].P:=0;
    end;
  inc(NomRow);
  inc(NomNagryzka);
  st:=Excel.Cells[NomRow,1];
  end;
Excel.Workbooks.Close;
end;
end;

Procedure LoadAuditorii(FileName:string);
var
NomRow,NomArr,i,KolAuditorii:longword;
st,st1:string;
begin
if FileExists(FileName) then
begin
Excel.Workbooks.Open(FileName);
NomRow:=2;
St1:=Excel.Cells[NomRow,1];
while st1<>'' do
  begin
  KolAuditorii:=Length(ArrAuditorii);
  SetLength(ArrAuditorii,KolAuditorii+1);
  ArrAuditorii[KolAuditorii]:=TAuditoria.Create;
  ArrAuditorii[KolAuditorii].Auditoria:=st1;
  st:=Excel.Cells[NomRow,2];
  ArrAuditorii[KolAuditorii].Korpus:=st;

  st:=Excel.Cells[NomRow,3];
  if St='1' then begin NomArr:=Length(ArrAuditoriiKP); SetLength(ArrAuditoriiKP,NomArr+1); ArrAuditoriiKP[NomArr]:=KolAuditorii; end;
  st:=Excel.Cells[NomRow,4];
  if St='1' then begin NomArr:=Length(ArrAuditoriiKons); SetLength(ArrAuditoriiKons,NomArr+1); ArrAuditoriiKons[NomArr]:=KolAuditorii; end;
  st:=Excel.Cells[NomRow,5];
  if St='1' then begin NomArr:=Length(ArrAuditSRS); SetLength(ArrAuditSRS,NomArr+1); ArrAuditSRS[NomArr]:=KolAuditorii; end;
  st:=Excel.Cells[NomRow,6];
  if St='1' then begin NomArr:=Length(ArrAuditoriiKontrol); SetLength(ArrAuditoriiKontrol,NomArr+1); ArrAuditoriiKontrol[NomArr]:=KolAuditorii; end;
  st:=Excel.Cells[NomRow,7];
  if St='1' then begin NomArr:=Length(ArrAuditoriiObslyz); SetLength(ArrAuditoriiObslyz,NomArr+1); ArrAuditoriiObslyz[NomArr]:=KolAuditorii; end;


  for i := 1 to 3 do
    begin
    St:=Excel.Cells[NomRow,i+7];
    if st<>'' then
      ArrAuditorii[KolAuditorii].KolStudentAuditoriaMax[i]:=StrToInt(st);
    end;
  St:=Excel.Cells[NomRow,11];
  if st<>'' then
    ArrAuditorii[KolAuditorii].KomputersAuditoria:=StrToInt(st);
  ArrAuditorii[KolAuditorii].ProektorAuditoria:=Excel.Cells[NomRow,12];
  ArrAuditorii[KolAuditorii].OsnashenieOgrnAuditoria:=Excel.Cells[NomRow,13];
  ArrAuditorii[KolAuditorii].NameAuditoria:=Excel.Cells[NomRow,14];
  inc(NomRow);
  St1:=Excel.Cells[NomRow,1];
  end;
Fmain.MeProtocol.Lines.Add('Загружено оснащение аудиторий из файла '+FileName);
Excel.Workbooks.Close;
end;
end;

Procedure LoadHourStudentDisFromExcelFile(NameFile:string);
var
  NomRow:Longword;
  st:string;
begin
SetLength(HourStudentDis,0);
Excel.Workbooks.Open(NameFile);
NomRow:=1;
st:=Excel.Cells[NomRow,1];
while st<>'' do
  begin
  SetLength(HourStudentDis,NomRow);
  HourStudentDis[NomRow-1].Dis:=Excel.Cells[NomRow,1];
  HourStudentDis[NomRow-1].Vid:=Excel.Cells[NomRow,2];
  HourStudentDis[NomRow-1].Group:=Excel.Cells[NomRow,3];
  HourStudentDis[NomRow-1].HourForOneStudent:=Excel.Cells[NomRow,4];
  inc(NomRow);
  st:=Excel.Cells[NomRow,1];
  end;
Excel.Workbooks.Close;
end;

Procedure LoadGroupDisFromExcelFile(NameFile:string);
var
  NomRow,NomCol,NomGroup:Longword;
  st:string;
begin
SetLength(ArrMergeDis,0);
Excel.Workbooks.Open(NameFile);
NomRow:=1;
st:=Excel.Cells[NomRow,1];
while st<>'' do
  begin
  SetLength(ArrMergeDis,NomRow);
  NomCol:=1;
  while st<>'' do
    begin
    SetLength(ArrMergeDis[NomRow-1],NomCol div 2+1);
    ArrMergeDis[NomRow-1][NomCol div 2].Dis:=Excel.Cells[NomRow,NomCol];
    ArrMergeDis[NomRow-1][NomCol div 2].Vid:=Excel.Cells[NomRow,NomCol+1];
    NomCol:=NomCol+2;
    st:=Excel.Cells[NomRow,NomCol];
    end;
  inc(NomRow);
  st:=Excel.Cells[NomRow,1];
  end;
Excel.Workbooks.Close;
end;

Procedure LoadPrepodFromExcelFile(NameFile:String);
var
  NomRow,NomPrepod:Longword;
  st:string;
begin
Excel.Workbooks.Open(NameFile);
NomRow:=2;
st:=Excel.Cells[NomRow,1];
while st<>'' do
  begin
  NomPrepod:=SeartchPrepodFIO(st);
  if NomPrepod<>65000 then
    begin
    Prepod[NomPrepod].Dolzhnost:=Excel.Cells[NomRow,2];
    Prepod[NomPrepod].Stepen:=Excel.Cells[NomRow,3];
    Prepod[NomPrepod].Zvanie:=Excel.Cells[NomRow,4];
    Prepod[NomPrepod].PovKvalProsh:=Excel.Cells[NomRow,5];
    Prepod[NomPrepod].PovKval:=Excel.Cells[NomRow,6];
    Prepod[NomPrepod].stavka:=Excel.Cells[NomRow,7];
    Prepod[NomPrepod].StavkaSovmest:=Excel.Cells[NomRow,8];
    if Prepod[NomPrepod].StavkaSovmest<>'' then
      Prepod[NomPrepod].MesNeOplat:=Excel.Cells[NomRow,9];
    end;
  inc(NomRow);
  st:=Excel.Cells[NomRow,1];
  end;
Excel.Workbooks.Close;
end;

Procedure LoadAllRaspisanieAllGroup(RaspisanieDir:string);
var
  st:string;
  Auditoria,Dis,Vid,FIOPrep,Group,CurrentGroup,StDate,StTime:string;
  ArrStDate:array of TDateTime;
  CurrentDate,EndDate:TDateTime;
  NomPrepod,NomNagr,NomSt,KolDate,NomSem:Longword;
  BSearch:Boolean;
  Col,Row:Longword;
  SearchStr:TSearchRec;


Procedure AddNewAudTime(NomPrepod,NomNagr:Longword);
var
  NomDateProc,LenArr,NomArr:Longword;
  begin
  if Auditoria='ауд.440(3)' then
    if Random>0.5 then
      Auditoria:='ауд.434а(3)'
    else
      Auditoria:='ауд.434б(3)';
  if (Pos('ауд.каф.(-)',Auditoria)<>0) {and  (Pos('304',DisAudTime.Kaf)<>0)} then
    begin
    if Vid='ЛК' then
     Auditoria:=KafAudLK[random(10)]
    else
      Auditoria:=KafAudLR[random(8)]
    end;

  if VivodProtocol then
    FMain.MeProtocol.Lines.Add('add auditoria '+Auditoria+' '+Prepod[NomPrepod].FIO+' '+InttoStr(NomNagr)+' '+Prepod[NomPrepod].Nagryzka[NomNagr].Group[0].Nom);
  Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria:=SearchInMassAuditoriaName(Auditoria);
  LenArr:=Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime);
//  ShortDateFormat := 'dd.mm';
  if (KolDate<>0) and (length(ArrStDate)<>0) then
  for NomDateProc := 0 to KolDate - 1 do
    begin
    NomArr:=0;
    while (NomArr<lenArr) and not(
          (Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[NomArr].StDate=DateTimeToStr(ArrStDate[NomDateProc])) and
          (Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[NomArr].StTime=StTime)) do
      inc(NomArr);
    if not (NomArr<lenArr) then
      begin
      setLength(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime,LenArr+1);
      Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[LenArr].StDate:=DateTimeToStr(ArrStDate[NomDateProc]);
      Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[LenArr].StTime:=StTime;
      inc(LenArr);
      end;
    end;

end;

Procedure AddNewDISPrepod(NomPrepod,NomNagr:Longword);
var
NomGroupNagryzka:Longword;
  begin
  Prepod[NomPrepod].Nagryzka[NomNagr].Dis:=Dis;
  Prepod[NomPrepod].Nagryzka[NomNagr].Vid:=Vid;
  NomGroupNagryzka:=0;
  while pos(',',Group)<>0 do
    begin
    SetLength(Prepod[NomPrepod].Nagryzka[NomNagr].Group,NomGroupNagryzka+1);
    Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka]:=SearchAndCreateGroup(copy(Group,1,pos(',',Group)-1));
    inc(NomGroupNagryzka);
    delete(Group,1,pos(',',Group)+1);
    end;
  SetLength(Prepod[NomPrepod].Nagryzka[NomNagr].Group,NomGroupNagryzka+1);
  Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka]:=SearchAndCreateGroup(Group);
  Prepod[NomPrepod].Nagryzka[NomNagr].sem:=NomSem;
  AddNewAudTime(NomPrepod,NomNagr);
  end;

Procedure SearchAndAdd(Group:string);
var
  KolPrepod,NomSP,NomDisSp,NomGroupSp,KolNagr:Longword;
  NomSearch:boolean;
begin
NomSem:=1;
NomPrepod:=0;
NomSearch:=false;
while (NomPrepod<Length(Prepod)){ and (Pos(Copy(FIOPrep,1,pos(' ',FIOPrep)-1),Prepod[NomPrepod].FIO)=0)} do
 begin
 NomNagr:=0;
 while (NomNagr<Length(Prepod[NomPrepod].Nagryzka)) and not
      ((Pos(Dis,Prepod[NomPrepod].Nagryzka[NomNagr].Dis)<>0) and
      (Prepod[NomPrepod].Nagryzka[NomNagr].Vid=Vid) and
      (SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagr].Group,Group)<>65000)) do
   begin
      inc(NomNagr);
   end;
 if (NomNagr<Length(Prepod[NomPrepod].Nagryzka)) then
     begin
     AddNewAudTime(NomPrepod,NomNagr);
     NomSearch:=true;
     end;
{    else
      begin
      KolNagr:=Length(Prepod[NomPrepod].Nagryzka);
      SetLength(Prepod[NomPrepod].Nagryzka,KolNagr+1);
      AddNewDISPrepod(NomPrepod,KolNagr);
      end;  }
  inc(NomPrepod);
  end;
  if not NomSearch then
    begin
    if VivodProtocol then
    FMain.MeProtocol.Lines.Add('add prepod '+FIOPrep);
    //Добавить препода в базу
    KolPrepod:=Length(Prepod);
    SetLength(Prepod,KolPrepod+1);
    Prepod[KolPrepod]:=TPrepodAll.Create;
    Prepod[KolPrepod].FIO:=FIOPrep;
    SetLength(Prepod[KolPrepod].Nagryzka,1);
    Prepod[KolPrepod].Nagryzka[0]:=TNagryzkaPrepod.Create;
    Prepod[KolPrepod].Nagryzka[0].Prepod:=Prepod[KolPrepod];
    KolNagr:=0;
    AddNewDISPrepod(KolPrepod,KolNagr);
    end;
NomSP:=0;
While NomSp<Length(SemYP) do
  begin
  NomGroupSp:=0;
  While (NomGroupSp<Length(SemYP[NomSp].Group)) and (SemYP[NomSp].Group[NomGroupSp].Nom<>Group) do
    inc(NomGroupSp);
  if (NomGroupSp<Length(SemYP[NomSp].Group)) then
    begin
    NomDisSp:=0;
    while NomDisSp<Length(SemYP[NomSp].Disciplin) do
      begin
      if (SemYP[NomSp].Disciplin[NomDisSp].Name=Dis) then
        begin
        if (Vid='ЛК') and (SemYP[NomSp].Disciplin[NomDisSp].LK<>0) then
          begin
          SemYP[NomSp].Disciplin[NomDisSp].LKAud:=SearchInMassAuditoriaName(Auditoria);
          if SemYP[NomSp].Disciplin[NomDisSp].NomElektivDis<>65000 then
            SemYP[NomSp].Disciplin[SemYP[NomSp].Disciplin[NomDisSp].NomElektivDis].LKAud:=SemYP[NomSp].Disciplin[NomDisSp].LKAud;
          end;
        if (Vid='ЛР') and (SemYP[NomSp].Disciplin[NomDisSp].LR<>0) then
          begin
          SemYP[NomSp].Disciplin[NomDisSp].LRAud:=SearchInMassAuditoriaName(Auditoria);
          if SemYP[NomSp].Disciplin[NomDisSp].NomElektivDis<>65000 then
            SemYP[NomSp].Disciplin[SemYP[NomSp].Disciplin[NomDisSp].NomElektivDis].LRAud:=SemYP[NomSp].Disciplin[NomDisSp].LRAud;
          end;
        if (Vid='ПЗ') and (SemYP[NomSp].Disciplin[NomDisSp].PZ<>0) then
          begin
          SemYP[NomSp].Disciplin[NomDisSp].PZAud:=SearchInMassAuditoriaName(Auditoria);
          if SemYP[NomSp].Disciplin[NomDisSp].NomElektivDis<>65000 then
            SemYP[NomSp].Disciplin[SemYP[NomSp].Disciplin[NomDisSp].NomElektivDis].PZ:=SemYP[NomSp].Disciplin[NomDisSp].PZ;
          end;
        end;
      inc(NomDisSp);
      end;
    end;

  inc(NomSp);
  end;

end;

begin
//Пройти по всем файлам EXCEL

if FindFirst(RaspisanieDir+'*.xls',faDirectory,SearchStr)=0 then
  begin
  FMain.PRaspGroup.Color:=ClGreen;
  repeat
  Excel.Workbooks.Open(RaspisanieDir+SearchStr.Name);
  Group:=Excel.Cells[2,3];
  Delete(Group,1,Pos('ы',Group)+1);
  Row:=5;
  while Row<18 do
    begin
    Col:=2;
    while Col<8 do
      begin
      st:=Excel.Cells[4,Col];
      delete(st,1,2);
      StTime:=Copy(st,1,Pos('-',st)-1);
      st:=Excel.Cells[Row,Col];
      while st<>'' do
        begin
        //Аудитория, предмет, тип нарузки, преподаватель?группа (дата дд.мм - дата дд.мм)
        FIOPrep:=Copy(st,1,Pos(' ',st)+4);
       { if FIOPrep<>'' then
        begin  }
        Delete(St,1,Pos('-',st)+1);
        Delete(St,1,Pos('а',st)-1);
        Auditoria:=Copy(st,1,Pos(')',st));
        Delete(St,1,Pos(')',st)+2);
        Dis:=Copy(st,1,Pos(',',st)-1);    //В названии дисциплины моет быть запятая
        Delete(St,1,Pos(',',st)+1);
        Vid:=Copy(st,1,Pos('(',st)-2);    //Проверить со стандартным видом,
        Delete(St,1,Pos('(',st));
        if VivodProtocol then
        FMain.MeProtocol.Lines.Add(Group+' '+FIOPrep+' '+Dis+' '+Vid+' '+Auditoria);

 {       While not( (Vid='ЛР') or (Vid='ПЗ') or (Vid='ЛК') or (Vid='КП') or
              (Vid='КР') or (Vid='Консультация') or (Vid='Экзамен') or
              (Vid='Зачет с оценкой') or (Vid='Зачет') or (Vid='Практика') or
              (Vid='Руководство магистрами') or (Vid='Преддипломная практика') or
              (Vid='Диплом') or (Vid='Руководство аспирантами') or
              (Vid='Руководство кафедрой')) do
          begin
          Dis:=Dis+','+Vid;
          Vid:=Copy(st,1,Pos(',',st)-1);    //Проверить со стандартным видом,
          Delete(St,1,Pos(',',st)+1);
          end;      }
     {   NomSt:=1;
        while (st[Nomst]>'9') or (st[Nomst]<'0') do
          inc(NomSt);
        if St[NomSt-1]='М' then
          dec(NomSt);
        Delete(st,1,NomSt-1);
        Group:=Copy(st,1,Pos('(',st)-1);       //Разобрать группы через запятую
        Delete(St,1,Pos('(',st)); }
        StDate:=Copy(st,1,Pos(')',st)-1);      //попробовать разобрать дату
        Delete(St,1,Pos(')',st)+1);
        //Проверить объединение ячеек.
        if (Pos('-',StDate)=0) and (Pos(',',StDate)=0) then
          begin
          CurrentDate:=StrToDateTime(StDate);
          EndDate:=CurrentDate;
          end
        else if Pos(',',StDate)<>0 then
          begin
          CurrentDate:=StrToDateTime(Copy(StDate,1,5));
          EndDate:=StrToDateTime(Copy(StDate,Length(stDate)-5,5));
          end
        else
          begin
          CurrentDate:=StrToDateTime(Copy(StDate,1,Pos('-',StDate)-1));
          EndDate:=StrToDateTime(Copy(StDate,Pos('-',StDate)+1,Length(stDate)));
          end;
        KolDate:=0;
        While CurrentDate<=EndDate do
          begin
          inc(KolDate);
          SetLength(ArrStDate,KolDate);
          ArrStDate[KolDate-1]:=CurrentDate;
          CurrentDate:=CurrentDate+14;
          end;
        //Найти эту дисциплину

        BSearch:=false;

        SearchAndAdd(Group);
        delete(st,1,5);
       { end
        else
          st:='';
          }
        end;
      inc(Col);
      end;
    inc(Row);
    end;
  Excel.Workbooks.Close;
  FMain.MeProtocol.Lines.Add('Загружен файл с расписанием группы '+RaspisanieDir+SearchStr.Name);
  until FindNext(SearchStr)<>0;
  end;

end;

Procedure LoadAllRaspisanieAllPrepod(RaspisanieDir:string; TypePrepod:byte);
var
  st:string;
  Auditoria,Dis,Vid,FIOPrep,Group,CurrentGroup,StDate,StDateAll,StTime:string;
  ArrStDate:array of TDateTime;
  CurrentDate,EndDate:TDateTime;
  NomPrepod,NomNagr,NomSt,KolDate:Longword;
  BSearch:Boolean;
  Col,Row:Longword;
  StMerge:string;
  MergeCells:Boolean;
  NomPrepodRasp:Longword;

Procedure SearchAndAdd(Group:string);
var
  NomDateProc,LenArr,NomArr:Longword;
  NomSearch:boolean;
  NomRowBase:LongWord;
begin
NomSearch:=false;
NomPrepod:=0;
while (NomPrepod<length(Prepod)) do
 begin
 NomNagr:=0;
 while (NomNagr<Length(Prepod[NomPrepod].Nagryzka)) do
   begin
   if (Pos(Dis,Prepod[NomPrepod].Nagryzka[NomNagr].Dis)<>0) and
      (Prepod[NomPrepod].Nagryzka[NomNagr].Vid=Vid) and
      (SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagr].Group,Group)<>65000) then
     begin
     Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria:=SearchInMassAuditoriaName(Auditoria);
     NomSearch:=true;


     //SetLength(Prepod[NomPrepod].Nagryzka[NomNagr].StDate,KolDate);
     LenArr:=Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime);
  //   ShortDateFormat := 'dd.mm';
     for NomDateProc := 0 to KolDate - 1 do
       begin
       NomArr:=0;
       while (NomArr<lenArr) and not(
          (Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[NomArr].StDate=DateTimeToStr(ArrStDate[NomDateProc])) and
          (Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[NomArr].StTime=StTime)) do
           inc(NomArr);
       if not (NomArr<lenArr) then
         begin
         setLength(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime,LenArr+1);
         Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[LenArr].StDate:=DateTimeToStr(ArrStDate[NomDateProc]);
         Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[LenArr].StTime:=StTime;
         inc(LenArr);
         end;
       end;

//     Prepod[NomPrepod].Nagryzka[NomNagr].StTime:=StTime;
     end;
   inc(NomNagr);
   end;
 inc(NomPrepod);
 end;
if not NomSearch then
  begin
  NomRowBase:=ExcelBase.Cells[1,1];
  inc(NomRowBase);
  ExcelBase.Cells[1,1]:=NomRowBase;
  ExcelBase.Cells[NomRowBase,1]:=FIOPrep;
  ExcelBase.Cells[NomRowBase,2]:=Vid;
  ExcelBase.Cells[NomRowBase,3]:=Dis;
  ExcelBase.Cells[NomRowBase,4]:=Group;
  ExcelBase.Cells[NomRowBase,5]:=Auditoria;
  ExcelBase.Cells[NomRowBase,6]:=StTime;
  ExcelBase.Cells[NomRowBase,7]:=StDateAll;
  end;
 end;

Procedure LoadFilePrepod (NameFile:String);
begin
  Excel.Workbooks.Open(NameFile);
  Row:=5;
  while Row<18 do
    begin
    Col:=2;
    while Col<8 do
      begin
      st:=Excel.Cells[4,Col];
      delete(st,1,2);
      StTime:=Copy(st,1,Pos('-',st)-1);
      st:=Excel.Cells[Row,Col];
      while st<>'' do
        begin
        //Проверить объединение ячеек.
        StMerge:=Excel.Cells[Row+1,Col];
        if (Row mod 2<>0) then
          MergeCells:=(Excel.Range[Excel.Cells[Row,Col],Excel.Cells[Row+1,Col]].MergeCells)
        else
          MergeCells:=false;
        //Аудитория, предмет, тип нарузки, преподаватель?группа (дата дд.мм - дата дд.мм)
        Auditoria:=Copy(st,1,Pos(',',st)-1);
        Delete(St,1,Pos(',',st)+1);
        Dis:=Copy(st,1,Pos(',',st)-1);    //В названии дисциплины моет быть запятая
        Delete(St,1,Pos(',',st)+1);
        Vid:=Copy(st,1,Pos(',',st)-1);    //Проверить со стандартным видом,
        Delete(St,1,Pos(',',st)+1);

        While not( (Vid='ЛР') or (Vid='ПЗ') or (Vid='ЛК') or (Vid='КП') or
              (Vid='КР') or (Vid='Консультация') or (Vid='Экзамен') or
              (Vid='Зачет с оценкой') or (Vid='Зачет') or (Vid='Практика') or
              (Vid='Руководство магистрами') or (Vid='Преддипломная практика') or
              (Vid='Диплом') or (Vid='Руководство аспирантами') or
              (Vid='Руководство кафедрой')) do
          begin
          Dis:=Dis+', '+Vid;
          Vid:=Copy(st,1,Pos(',',st)-1);    //Проверить со стандартным видом,
          Delete(St,1,Pos(',',st)+1);
          end;
        while Dis[Length(Dis)]=' ' do
          delete(Dis,Length(Dis),1);
//        FIOPrep:=Copy(st,1,Pos(' ',st)+2);
//        Delete(St,1,Pos(' ',st)+2);
        NomSt:=1;
        while (st[Nomst]>'9') or (st[Nomst]<'0') do
          inc(NomSt);
        FIOPrep:=Copy(st,1,NomSt);
        if St[NomSt-1]='М' then
          dec(NomSt);
        Delete(st,1,NomSt-1);
        Group:=Copy(st,1,Pos('(',st)-2);       //Разобрать группы через запятую
        Delete(St,1,Pos('(',st));
        StDate:=Copy(st,1,Pos(')',st)-1);
        Delete(St,1,Pos(')',st)+1);

        //Проверить объединение ячеек.
        if (Pos('-',StDate)=0) and (Pos(',',StDate)=0) then
          begin
          CurrentDate:=StrToDateTime(StDate);
          EndDate:=CurrentDate;
          end
        else if Pos(',',StDate)<>0 then
          begin
          CurrentDate:=StrToDateTime(Copy(StDate,1,5));
          EndDate:=StrToDateTime(Copy(StDate,Length(stDate)-5,5));
          end
        else
          begin
          CurrentDate:=StrToDateTime(Copy(StDate,1,Pos('-',StDate)-1));
          EndDate:=StrToDateTime(Copy(StDate,Pos('-',StDate)+1,Length(stDate)));
          end;
        KolDate:=0;
        While CurrentDate<=EndDate do
          begin
          inc(KolDate);
          SetLength(ArrStDate,KolDate);
          ArrStDate[KolDate-1]:=CurrentDate;
          CurrentDate:=CurrentDate+14;
          end;

      {  StDateAll:=StDate;
        KolDate:=0;
        While Pos(',',StDate)<>0 do
          begin
          CurrentDate:=StrToDateTime(Copy(StDate,1,Pos(',',StDate)-1));
          inc(KolDate);
          SetLength(ArrStDate,KolDate);
          ArrStDate[KolDate-1]:=CurrentDate;
          Delete(StDate,1,Pos(',',StDate)+1);
          end;
        inc(KolDate);
        SetLength(ArrStDate,KolDate);
        ArrStDate[KolDate-1]:=StrToDateTime(StDate);  }
        //Найти эту дисциплину

        BSearch:=false;
        while pos(',',Group)<>0 do                     //Разобрать группы через запятую
          begin
          SearchAndAdd(Copy(Group,1,pos(',',Group)-1));
          Delete(Group,1,pos(',',Group)+1);
          end;
        SearchAndAdd(Group);
        delete(st,1,Pos('а',st)-1);
        end;
      inc(Col);
      end;
    inc(Row);
    end;
  Excel.Workbooks.Close;
  FMain.MeProtocol.Lines.Add('Загружен файл с расписанием:'+NameFile);

end;

begin
case TypePrepod of
  0:  begin
      NomPrepodRasp:=0;
      while NomPrepodRasp<Length(Prepod) do
        begin
        if FileExists(RaspisanieDir+Prepod[NomPrepodRasp].FIO+'.xlsx') then
          begin
          FMain.PRaspPrepod.Color:=ClGreen;
          LoadFilePrepod (RaspisanieDir+Prepod[NomPrepodRasp].FIO+'.xlsx');
          end;
{        if FileExists(RaspisanieDir+Prepod[NomPrepod].FIO+'_2.xlsx') then
          LoadFilePrepod (RaspisanieDir+Prepod[NomPrepod].FIO+'_2.xlsx'); }
        inc(NomPrepodRasp);
        end;
      end;
  1:  begin
      if FileExists(RaspisanieDir) then
          begin
          FMain.PEkz.Color:=ClGreen;
          LoadFilePrepod (RaspisanieDir);
          end;
      end;
end;


end;

Procedure LoadRaspIASYEkzamenGroupExcel(FileName:String);
var
  st:string;
  Auditoria,Dis,Vid,FIOPrep,Group,StDate,StTime:string;
  NomPrepod,NomNagr,NomSt,KolAuditorii:Longword;
  BSearch:Boolean;
  Col,Row:Longword;
begin
Excel.Workbooks.Open(FileName);
Col:=7;
St:=Excel.Cells[8,Col];
while st<>'' do
  begin
  Group:=st;
  Row:=10;
  St:=Excel.Cells[Row,1];
  while st<>'' do
    begin
    st:=Excel.Cells[Row,Col];
    if st<>'' then
      begin
      //Предмет, ?Преподаватель, время, ауд.
      Dis:=Copy(st,1,Pos(Chr(10),St)-1);
      Delete(st,1,Pos(Chr(10),St));
      if (St[1]<>'0')  and (St[1]<>'1') then
        begin
        FIOPrep:=Copy(st,1,Pos(',',St)-1);
        Delete(st,1,Pos(',',St)+1);
        end;
      StTime:=Copy(st,1,5);
      Delete(st,1,Pos(',',St)+1);
      Auditoria:=st;
      StDate:='';
      St:=Excel.Cells[Row,1];
      st:=Copy(st,1,Pos(Chr(10),St)-1);
      While (St[1]<='9') and (St[1]>='0') do
        begin
        StDate:=StDate +st[1];
        Delete(st,1,1);
        end;
      if pos('декабря',st)<>0 then StDate:=StDate+'.12.'+IntToStr(CurrentYear);
      if pos('января',st)<>0 then StDate:=StDate+'.01.'+IntToStr(CurrentYear);
      if pos('февраля',st)<>0 then StDate:=StDate+'.02.'+IntToStr(CurrentYear);
      if pos('мая',st)<>0 then StDate:=StDate+'.05.'+IntToStr(CurrentYear);
      if pos('июня',st)<>0 then StDate:=StDate+'.06.'+IntToStr(CurrentYear);
      if pos('июля',st)<>0 then StDate:=StDate+'.07.'+IntToStr(CurrentYear);


      //Найти эту дисциплину
      NomPrepod:=0;
      BSearch:=false;
      while (NomPrepod<Length(Prepod)) and (not BSearch) do
        begin
        NomNagr:=0;
        while (NomNagr<Length(Prepod[NomPrepod].Nagryzka)) and (not BSearch) do
          begin
          if (Pos(Dis,Prepod[NomPrepod].Nagryzka[NomNagr].Dis)<>0) and (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Экзамен')
             and (SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagr].Group,Group)<>65000) then
            BSearch:=true;
          if not BSearch then
            inc(NomNagr);
          end;
        if not BSearch then
          inc(NomPrepod);
        end;
      if BSearch then
        begin
        if SearchInMassAuditoriaName(Auditoria)=ArrAuditorii[0] then
          begin
          KolAuditorii:=Length(ArrAuditorii);
          SetLength(ArrAuditorii,KolAuditorii+1);
          ArrAuditorii[KolAuditorii]:=TAuditoria.Create;
          ArrAuditorii[KolAuditorii].Auditoria:=Auditoria;
          end;
        Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria:=SearchInMassAuditoriaName(Auditoria);
        //Доделать на множество времени
        if length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)=0 then
          begin
          SetLength(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime,1);
          Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate:=StDate;
          Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StTime:=StTime;
          end;
//  StrToDateTime
        end;
      end;
    inc(Row);
    St:=Excel.Cells[Row,1];
    end;
  Col:=Col+17;
  St:=Excel.Cells[8,Col];
  end;
FMain.MeProtocol.Lines.Add(st);
Excel.Workbooks.Close;
end;

Procedure LoadRaspIASYEkzamenExcel(FileName:String);
var
  st:string;
  Auditoria,Dis,Vid,FIOPrep,Group,GroupAll,StDate,StTime:string;
  NomPrepod,NomNagr,NomSt:Longword;
  BSearch:Boolean;
  Col,Row:Longword;
begin
Excel.Workbooks.Open(FileName);
Row:=5;
while Row<18 do
  begin
  Col:=2;
  while Col<8 do
    begin
    st:=Excel.Cells[4,Col];
    delete(st,1,2);
    StTime:=Copy(st,1,Pos('-',st)-1);
    st:=Excel.Cells[Row,Col];
    while st<>'' do
      begin
      //Аудитория, предмет, тип нарузки, преподаватель?группа (дата дд.мм - дата дд.мм)
      Auditoria:=Copy(st,1,Pos(',',st)-1);
      Delete(St,1,Pos(',',st)+1);
      Dis:=Copy(st,1,Pos(',',st)-1);    //В названии дисциплины моет быть запятая
      Delete(St,1,Pos(',',st)+1);
      Vid:=Copy(st,1,Pos(',',st)-1);    //Проверить со стандартным видом,
      Delete(St,1,Pos(',',st)+1);

      While not( (Vid='ЛР') or (Vid='ПЗ') or (Vid='ЛК') or (Vid='КП') or
              (Vid='КР') or (Vid='Консультация') or (Vid='Экзамен') or
              (Vid='Зачет с оценкой') or (Vid='Зачет') or (Vid='Практика') or
              (Vid='Руководство магистрами') or (Vid='Преддипломная практика') or
              (Vid='Диплом') or (Vid='Руководство аспирантами') or
              (Vid='Руководство кафедрой')) do
          begin
          Dis:=Dis+', '+Vid;
          Vid:=Copy(st,1,Pos(',',st)-1);    //Проверить со стандартным видом,
          Delete(St,1,Pos(',',st)+1);
          end;
      while Dis[Length(dis)]=' ' do
        Delete(Dis,Length(dis),1);
      NomSt:=1;
      while (st[Nomst]>'9') or (st[Nomst]<'0') do
        inc(NomSt);
      if St[NomSt-1]='М' then
        dec(NomSt);
      Delete(st,1,NomSt-1);
      GroupAll:=Copy(st,1,Pos('(',st)-2);
      GroupAll:=GroupAll+',';
      Delete(St,1,Pos('(',st));
      StDate:=Copy(st,1,Pos(')',st)-1);      //попробовать разобрать дату
      Delete(St,1,Pos(')',st)+1);
      //Найти эту дисциплину
      NomPrepod:=0;
      while pos(',',GroupAll)<>0 do                     //Разобрать группы через запятую
       begin
        Group:=Copy(GroupAll,1,pos(',',GroupAll)-1);
      BSearch:=false;
      while (NomPrepod<Length(Prepod)) and (not BSearch) do
        begin
        NomNagr:=0;
        while (NomNagr<Length(Prepod[NomPrepod].Nagryzka)) and (not BSearch) do
          begin
          if (Pos(Dis,Prepod[NomPrepod].Nagryzka[NomNagr].Dis)<>0) and (Prepod[NomPrepod].Nagryzka[NomNagr].Vid=Vid)
             and (SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagr].Group,Group)<>65000) then
            BSearch:=true;
          if not BSearch then
            inc(NomNagr);
          end;
        if not BSearch then
          inc(NomPrepod);
        end;
      if BSearch then
        begin
        Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria:=SearchAndAddInMassAuditoriaName(Auditoria);
        //Доделать на множество времени
        if length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)=0 then
          begin
          SetLength(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime,1);
          Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate:=StDate+'.'+IntToStr(CurrentYear);
          Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StTime:=StTime;
          end;
//  StrToDateTime
        end;
       Delete(GroupAll,1,pos(',',GroupAll)+1);
       end;
      delete(st,1,Pos('а',st)-1);
      end;
    inc(Col);
    end;
  inc(Row);
  end;
FMain.MeProtocol.Lines.Add(st);
Excel.Workbooks.Close;
end;


end.
