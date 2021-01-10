unit USaveExcel;

interface

uses SysUtils, UMain, UNagryzka, UGroup, UConstParametrs;

procedure VivodPlanNagryzka;
Procedure GoRaspisanieKaf(NomSem:Byte);
Procedure GoRaspisanieKafCv(NomSem:Byte; Cv:Boolean);
Procedure GoRaspisanieToExcel(NomSem:Byte);
procedure GoExzamenToExcel(NomSem:Byte);
Procedure SaveAllPrepodGroup;
procedure AddAllPrepodNagryzkaToExcelFile(NameFile:String);
Procedure GoTablExzamenGroupToExcel(NomSem:Byte);

implementation

procedure AddAllPrepodNagryzkaToExcelFile(NameFile:String);
var
  NomRow,NomPrepod:Longword;
  st:string;
begin
Excel.Workbooks.Open(NameFile);
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  NomRow:=2;
  st:=Excel.Cells[NomRow,1];
  while (st<>'') and (st<>Prepod[NomPrepod].FIO) do
    begin
    inc(NomRow);
    st:=Excel.Cells[NomRow,1];
    end;
  if st='' then
    begin
    Excel.Cells[NomRow,1]:=Prepod[NomPrepod].FIO;
    Excel.Cells[NomRow,2]:=Prepod[NomPrepod].Dolzhnost;
    Excel.Cells[NomRow,3]:=Prepod[NomPrepod].Stepen;
    Excel.Cells[NomRow,4]:=Prepod[NomPrepod].Zvanie;
    Excel.Cells[NomRow,5]:=Prepod[NomPrepod].PovKvalProsh;
    Excel.Cells[NomRow,6]:=Prepod[NomPrepod].PovKval;
    Excel.Cells[NomRow,7]:=Prepod[NomPrepod].stavka;
    Excel.Cells[NomRow,8]:=Prepod[NomPrepod].StavkaSovmest;
    if Prepod[NomPrepod].StavkaSovmest<>'' then
      Excel.Cells[NomRow,9]:=Prepod[NomPrepod].MesNeOplat;
    end;
  inc(NomPrepod)
  end;
Excel.Workbooks[1].save;
Excel.Workbooks.Close;
end;

procedure VivodPlanNagryzka;
Var
NomPrepod,NomRow:Longword;
st,st1:string;
HoursSt,PravSovmest,Summa:Double;
const
  AndNomRow = 9;
begin
Excel.WorkBooks.Add;

Excel.Columns[1].ColumnWidth := 2.29;
Excel.Columns[2].ColumnWidth := 8.86;
Excel.Columns[3].ColumnWidth := 19.71;
Excel.Columns[4].ColumnWidth := 11.57;
Excel.Columns[5].ColumnWidth := 10.43;
Excel.Columns[6].ColumnWidth := 11.14;
Excel.Columns[7].ColumnWidth := 13.14;
Excel.Columns[8].ColumnWidth := 10.86;

Excel.Cells[1,1]:='Проректору по учебной работе';
Excel.Cells[2,1]:='Д.А. Козорезу';
Excel.Cells[4,1]:='Объем выполняемой и плановой педагогической нагрузки';
Excel.Cells[5,1]:='преподавателей кафедры 304 за 2018-2019 учебный год';

Excel.Cells[7,1]:='№';
Excel.Cells[7,2]:='Наим. кафедры';
Excel.Cells[7,3]:='Ф.И.О';
Excel.Cells[7,4]:='Долж.';
Excel.Cells[7,5]:='Учебная  нагрузка';
Excel.Cells[7,8]:='Итого часов';
Excel.Cells[8,5]:='Объем ауд. нагрузки';
Excel.Cells[8,6]:='Объем вне-ауд. нагрузки';
Excel.Cells[8,7]:='Учебно-методическая нагрузка';

Excel.Range[Excel.Cells[1,1],Excel.Cells[1,8]].MergeCells:=true;
Excel.Range[Excel.Cells[2,1],Excel.Cells[2,8]].MergeCells:=true;
Excel.Range[Excel.Cells[4,1],Excel.Cells[4,8]].MergeCells:=true;
Excel.Range[Excel.Cells[5,1],Excel.Cells[5,8]].MergeCells:=true;
Excel.Range[Excel.Cells[7,1],Excel.Cells[8,1]].MergeCells:=true;
Excel.Range[Excel.Cells[7,2],Excel.Cells[8,2]].MergeCells:=true;
Excel.Range[Excel.Cells[7,3],Excel.Cells[8,3]].MergeCells:=true;
Excel.Range[Excel.Cells[7,4],Excel.Cells[8,4]].MergeCells:=true;
Excel.Range[Excel.Cells[7,8],Excel.Cells[8,8]].MergeCells:=true;
Excel.Range[Excel.Cells[7,5],Excel.Cells[7,7]].MergeCells:=true;

Excel.Range[Excel.Cells[4,1],Excel.Cells[4,8]].HorizontalAlignment:=xlCenter;
Excel.Range[Excel.Cells[5,1],Excel.Cells[5,8]].HorizontalAlignment:=xlCenter;

NomPrepod:=0;
NomRow:=0;
while NomPrepod<Length(Prepod) do
   begin
   Excel.Cells[NomRow+AndNomRow,1]:=NomRow+1;
   Excel.Cells[NomRow+AndNomRow,2]:='304';
   Excel.Cells[NomRow+AndNomRow,3]:=Prepod[NomPrepod].FIO;
   st:=Prepod[NomPrepod].Dolzhnost+' ('+Prepod[NomPrepod].stavka+'ст.)';
   Excel.Cells[NomRow+AndNomRow,4]:=st;
   HoursSt:=0;
   PravSovmest:=0;
   if Prepod[NomPrepod].stavka<>'' then
     HoursSt:=StrToFloat(Prepod[NomPrepod].stavka)*830;
   st:=FloatToStr(HoursSt);
   Summa:=HoursSt;
   if Prepod[NomPrepod].StavkaSovmest<>'' then
     begin
     Excel.Cells[NomRow+1+AndNomRow,1]:=NomRow+2;
     Excel.Cells[NomRow+1+AndNomRow,2]:='304';
     Excel.Cells[NomRow+1+AndNomRow,3]:=Prepod[NomPrepod].FIO;
     st1:=Prepod[NomPrepod].Dolzhnost+' ('+Prepod[NomPrepod].StavkaSovmest+'ст.)';
     Excel.Cells[NomRow+1+AndNomRow,4]:=st1;
     HoursSt:=HoursSt+StrToFloat(Prepod[NomPrepod].StavkaSovmest)*830;
     Excel.Cells[NomRow+1+AndNomRow,5]:=StrToFloat(Prepod[NomPrepod].StavkaSovmest)*830;
     Excel.Cells[NomRow+1+AndNomRow,6]:='0';
     if Prepod[NomPrepod].AllHour<>0 then
       PravSovmest:=Prepod[NomPrepod].AllHourPrav*(StrToFloat(Prepod[NomPrepod].StavkaSovmest)*830/Prepod[NomPrepod].AllHour);
     Excel.Cells[NomRow+1+AndNomRow,7]:=PravSovmest;
     end;


   if HoursSt<Prepod[NomPrepod].AllHour then
     begin
     st:=st+'+'+FloatToStr(Prepod[NomPrepod].AllHour-HoursSt);
     Summa:=Summa+Prepod[NomPrepod].AllHour-HoursSt;
     Prepod[NomPrepod].Pochasovka:=Prepod[NomPrepod].AllHour-HoursSt;
     end
   else
     Summa:=Summa+HoursSt;
   Excel.Cells[NomRow+AndNomRow,5]:=st;
   Excel.Cells[NomRow+AndNomRow,9]:=Summa;
   Excel.Cells[NomRow+AndNomRow,6]:=Prepod[NomPrepod].PovKval;
   Excel.Cells[NomRow+AndNomRow,7]:=Prepod[NomPrepod].AllHourPrav-PravSovmest;
   inc(NomRow);
   if Prepod[NomPrepod].StavkaSovmest<>'' then
     inc(NomRow);
   inc(NomPrepod);

   end;

  Excel.Range[Excel.Cells[7,1],Excel.Cells[NomRow-1,8]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[7,1],Excel.Cells[NomRow-1,8]].BorderAround(1);
//  Excel.Range[Excel.Cells[7,1],Excel.Cells[NomRow-1,8]].Font.Size:=10;
  Excel.Range[Excel.Cells[7,1],Excel.Cells[NomRow-1,8]].WrapText:=true;
  Excel.Range[Excel.Cells[7,1],Excel.Cells[NomRow-1,8]].VerticalAlignment:=xlCenter;
  Excel.Range[Excel.Cells[7,1],Excel.Cells[NomRow-1,8]].HorizontalAlignment:=xlCenter;

Excel.Workbooks[1].saveas(CurrentDir+'\Объем плановой и выполняемой нагрузки.xlsx');
FMain.MeProtocol.Lines.Add('Создан файл '+CurrentDir+'\Объем плановой и выполняемой нагрузки.xlsx');
Excel.Workbooks.Close;
end;

Procedure GoRaspisanieKafCv(NomSem:Byte; Cv:Boolean);
var
  NomRow,NomRowNat,NomCol,NomDay,NomTime,NomGroupNagryzka,NomRowKonMax:LongWord;
  NomPrepod,NomNagryzka,NomDate:Longword;
  NomRowDate,NextStringDate:byte;
  NomRowKon:array [1..36]of Longword;
  st:string;
begin
Excel.WorkBooks.Add;
Excel.Cells[1,1]:='ФИО';  Excel.Range[Excel.Cells[1,1],Excel.Cells[2,1]].MergeCells:=true;
Excel.Columns[1].ColumnWidth:=12.86;  //ФИО
//20,71/9=2,3
//9 дат*6 пар*6 дней=324 ячеек
for NomCol := 2 to 325 do
  Excel.Columns[NomCol].ColumnWidth := 1.57;
//9*6=54 на один день недели
Excel.Cells[1,2]:='Понедельник'; Excel.Range[Excel.Cells[1,2],Excel.Cells[1,55]].MergeCells:=true;
Excel.Cells[1,56]:='Вторник';    Excel.Range[Excel.Cells[1,56],Excel.Cells[1,109]].MergeCells:=true;
Excel.Cells[1,110]:='Среда';     Excel.Range[Excel.Cells[1,110],Excel.Cells[1,163]].MergeCells:=true;
Excel.Cells[1,164]:='Четверг';   Excel.Range[Excel.Cells[1,164],Excel.Cells[1,217]].MergeCells:=true;
Excel.Cells[1,218]:='Пятница';   Excel.Range[Excel.Cells[1,218],Excel.Cells[1,271]].MergeCells:=true;
Excel.Cells[1,272]:='Cуббота';   Excel.Range[Excel.Cells[1,272],Excel.Cells[1,325]].MergeCells:=true;

//Время
NomCol:=0;
while NomCol<6  do
  begin
  for NomTime := 0 to 5 do
    begin
    Excel.Cells[2,2+NomCol*54+NomTime*9]:=TimeSetPar[NomTime];
    Excel.Range[Excel.Cells[2,2+NomCol*54+NomTime*9],Excel.Cells[2,2+NomCol*54+(NomTime+1)*9-1]].MergeCells:=true;
    end;
  inc(NomCol);
  end;

Excel.Range[Excel.Cells[1,1],Excel.Cells[2,325]].WrapText:=true; Excel.Range[Excel.Cells[1,1],Excel.Cells[2,325]].VerticalAlignment:=xlCenter; Excel.Range[Excel.Cells[1,1],Excel.Cells[2,325]].HorizontalAlignment:=xlCenter;
Excel.Range[Excel.Cells[1,1],Excel.Cells[2,325]].Borders.Weight := 2; Excel.Range[Excel.Cells[1,1],Excel.Cells[2,325]].BorderAround(1);

NomPrepod:=0;
NomRow:=3;
while NomPrepod<Length(Prepod) do
  begin
  Excel.Cells[NomRow,1]:=Prepod[NomPrepod].FIO;
  //15 и 40 и 15  RowHeight
  NomRowNat:=NomRow;
  NomDay:=1;
  while NomDay<=36 do
    begin
    NomRowKon[NomDay]:=NomRow;
    inc(NomDay);
    end;
  NomDay:=1;
  NomCol:=2;
  while NomDay<=6 do
    begin
    NomTime:=0;
    while NomTime<6  do
      begin
      NomRow:=NomRowNat;
      NomNagryzka:=0;
      while NomNagryzka<length(Prepod[NomPrepod].Nagryzka) do
        begin
        if (Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid<>'Экзамен') then
        begin
        NomDate:=0;
        while (NomDate<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime)) and not (
             (DayOfWeek(StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDate].StDate))=NomDay+1) and
             (Pos(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDate].StTime,TimeSetPar[NomTime])<>0))  do
          inc(NomDate);
        if (Prepod[NomPrepod].Nagryzka[NomNagryzka].sem=NomSem) AND
           (NomDate<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime)) then
           begin
           //Установить все даты, которые соответствуют дню недели
           Excel.Rows[NomRow].RowHeight := 16.50; Excel.Rows[NomRow+1].RowHeight := 25.50; Excel.Rows[NomRow+2].RowHeight := 25.50; Excel.Rows[NomRow+3].RowHeight := 16.50;

           NomRowDate:=0; NextStringDate:=0; NomDate:=0;
           while (NomDate<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime)) do
             begin
             if (DayOfWeek(StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDate].StDate))=NomDay+1) and
                (Pos(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDate].StTime,TimeSetPar[NomTime])<>0) then
               begin
               Excel.Cells[NomRow+NextStringDate,NomCol+NomRowDate]:=FormatDateTime('dd mmm',StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDate].StDate));
               Excel.Cells[NomRow+NextStringDate,NomCol+NomRowDate].Font.Size:=6;
               inc(NomRowDate);
               if NomRowDate>8 then   //Учесть переход на строку ниже (по 9 дат)
                 begin
                 NomRowDate:=0;
                 NextStringDate:=3;
                 end;
               end;
             inc(NomDate);
             end;
           //Вывести информацию по дисциплине
           Excel.Cells[NomRow+1,NomCol]:=Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid; Excel.Range[Excel.Cells[NomRow+1,NomCol],Excel.Cells[NomRow+1,NomCol+1]].MergeCells:=true;
           Excel.Cells[NomRow+1,NomCol].Font.Size:=12;
           Excel.Cells[NomRow+1,NomCol+2]:=Prepod[NomPrepod].Nagryzka[NomNagryzka].Dis; Excel.Range[Excel.Cells[NomRow+1,NomCol+2],Excel.Cells[NomRow+1,NomCol+8]].MergeCells:=true;
           Excel.Cells[NomRow+1,NomCol+2].Font.Size:=8;
           Excel.Cells[NomRow+2,NomCol]:=Prepod[NomPrepod].Nagryzka[NomNagryzka].Auditoria.Auditoria; Excel.Range[Excel.Cells[NomRow+2,NomCol],Excel.Cells[NomRow+2,NomCol+3]].MergeCells:=true;
           Excel.Cells[NomRow+2,NomCol].Font.Size:=7;

           NomGroupNagryzka:=0;
           st:='';
           while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].Group) do
             begin
             st:=st+Prepod[NomPrepod].Nagryzka[NomNagryzka].Group[NomGroupNagryzka].Nom;
             inc(NomGroupNagryzka);
             If NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].Group) then st:=st+', ';
             end;
           Excel.Cells[NomRow+2,NomCol+4]:=st; Excel.Range[Excel.Cells[NomRow+2,NomCol+4],Excel.Cells[NomRow+2,NomCol+8]].MergeCells:=true;
           Excel.Cells[NomRow+2,NomCol+4].Font.Size:=8;

           Excel.Range[Excel.Cells[NomRow,NomCol],Excel.Cells[NomRow+3,NomCol+8]].WrapText:=true; Excel.Range[Excel.Cells[NomRow,NomCol],Excel.Cells[NomRow+3,NomCol+8]].VerticalAlignment:=xlCenter; Excel.Range[Excel.Cells[NomRow,NomCol],Excel.Cells[NomRow+3,NomCol+8]].HorizontalAlignment:=xlCenter;
           Excel.Range[Excel.Cells[NomRow,NomCol],Excel.Cells[NomRow+3,NomCol+8]].BorderAround(1);
           NomRowKon[(NomDay-1)*6+NomTime+1]:=NomRowKon[(NomDay-1)*6+NomTime+1]+4;
           NomRow:=NomRowKon[(NomDay-1)*6+NomTime+1];
           end;
        end;
        inc(NomNagryzka);
        end;
      inc(NomTime);
      NomCol:=NomCol+9;
      end;
    inc(NomDay);
    end;
  //Переход к следущему преподавателю.
  NomRowKonMax:=0;
  NomDay:=1;
  while NomDay<=36 do
    begin
    if NomRowKon[NomDay]>NomRowKonMax then
      NomRowKonMax:=NomRowKon[NomDay];
    inc(NomDay);
    end;
 { NomDay:=1;
  while NomDay<=36 do
    begin
    Excel.Range[Excel.Cells[NomRowKon[NomDay]+1,NomDay+1],Excel.Cells[NomRowKonMax-1,NomDay+1]].MergeCells:=true;
    inc(NomDay);
    end;            }
  Excel.Range[Excel.Cells[NomRowNat,1],Excel.Cells[NomRowKonMax-1,1]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRowNat,1],Excel.Cells[NomRowNat,1]].WrapText:=true; Excel.Cells[NomRowNat,1].VerticalAlignment:=xlCenter; Excel.Range[Excel.Cells[NomRowNat,1],Excel.Cells[NomRowNat,1]].HorizontalAlignment:=xlCenter;
  Excel.Range[Excel.Cells[NomRowNat,1],Excel.Cells[NomRowKonMax-1,325]].BorderAround(1);

  NomRow:=NomRowKonMax;
  Inc(NomPrepod);
  end;

 //Вертикальные столбцы
 Excel.Range[Excel.Cells[3,1],Excel.Cells[NomRowKonMax,1]].BorderAround(1);
NomCol:=0;
while NomCol<6  do
  begin
  for NomTime := 0 to 5 do
    begin
    Excel.Range[Excel.Cells[3,2+NomCol*54+NomTime*9],Excel.Cells[NomRowKonMax,2+NomCol*54+(NomTime+1)*9-1]].BorderAround(1);
    end;
  inc(NomCol);
  end;

Excel.Workbooks[1].saveas(CurrentDir+'\Расписание кафедры Новое.xlsx');
FMain.MeProtocol.Lines.Add('Создан файл с расписанием '+CurrentDir+'\Расписание кафедры Новое.xlsx');
Excel.Workbooks.Close;
end;

Procedure GoRaspisanieKaf(NomSem:Byte);
var
  NomPrepod,NomNagryzka, NomDate:Longword;
  NomRow,NomCol,NomDay,NomTime:byte;
  st,StGroup:string;
  NomGroupNagryzka:longword;
  CurrentDate:TDateTime;
begin
GoRaspisanieKafCv(NomSem,false);
Excel.WorkBooks.Add;
for NomCol := 2 to 38 do
  Excel.Columns[NomCol].ColumnWidth := 24.43;
Excel.Cells[1,1]:='ФИО';
Excel.Cells[1,2]:='Понедельник';
Excel.Cells[1,8]:='Вторник';
Excel.Cells[1,14]:='Среда';
Excel.Cells[1,20]:='Четверг';
Excel.Cells[1,26]:='Пятница';
Excel.Cells[1,32]:='Cуббота';
NomCol:=1;
while NomCol<=6  do
  begin
  for NomTime := 0 to 5 do
    Excel.Cells[2,2+(NomCol-1)*6+NomTime]:=TimeSetPar[NomTime];
  inc(NomCol);
  end;

NomPrepod:=0;
NomRow:=3;
while NomPrepod<Length(Prepod) do
  begin
  Excel.Cells[NomRow,1]:=Prepod[NomPrepod].FIO;
  NomDay:=1;
  NomCol:=2;
  while NomDay<=6 do
    begin
    NomTime:=1;
    while NomTime<=6  do
      begin
      st:='';
      NomNagryzka:=0;
      while NomNagryzka<length(Prepod[NomPrepod].Nagryzka) do
        begin
        if (Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid<>'Экзамен') then
        begin
        NomDate:=0;
        while (NomDate<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime)) and not (
             (DayOfWeek(StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDate].StDate))=NomDay+1) and
             (Pos(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDate].StTime,TimeSetPar[NomTime-1])<>0))  do
          inc(NomDate);
        if (Prepod[NomPrepod].Nagryzka[NomNagryzka].sem=NomSem) AND
           (NomDate<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime)) then
           begin
             if st<>'' then
               st:=st+chr(10)+'-----'+chr(10);
             StGroup:='';
             NomGroupNagryzka:=0;
             while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].Group) do
               begin
               StGroup:=StGroup+Prepod[NomPrepod].Nagryzka[NomNagryzka].Group[NomGroupNagryzka].Nom+', ';
               inc(NomGroupNagryzka);
               end;
             if Prepod[NomPrepod].Nagryzka[NomNagryzka].Auditoria<>nil then
               st:=st+Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid+'; '+Prepod[NomPrepod].Nagryzka[NomNagryzka].Dis+'; '+StGroup+'; '+Prepod[NomPrepod].Nagryzka[NomNagryzka].Auditoria.Auditoria+' ('+copy(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[0].StDate,1,5)+'-'+copy(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime)-1].StDate,1,5)+')';
           end;
        end;

        inc(NomNagryzka);
        end;
      Excel.Cells[NomRow,NomCol]:=st;
      inc(NomTime);
      inc(NomCol);
      end;
    inc(NomDay);
    end;
  Inc(NomRow);
  Inc(NomPrepod);
  end;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,38]].Borders.Weight := 2;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,38]].BorderAround(1);
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,38]].Font.Size:=10;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,38]].WrapText:=true;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,38]].VerticalAlignment:=xlCenter;
//Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRownat,7]].HorizontalAlignment:=xlCenter;
Excel.Workbooks[1].saveas(CurrentDir+'\Расписание кафедры 1.xlsx');
FMain.MeProtocol.Lines.Add('Создан файл с расписанием '+CurrentDir+'\Расписание кафедры 1.xlsx');
Excel.Workbooks.Close;

Excel.WorkBooks.Add;
Excel.Columns[1].ColumnWidth :=14.00;
for NomCol := 2 to 7 do
  Excel.Columns[NomCol].ColumnWidth := 61.14;
Excel.Cells[1,1]:='ФИО';
Excel.Cells[1,2]:='Понедельник';
Excel.Cells[1,3]:='Вторник';
Excel.Cells[1,4]:='Среда';
Excel.Cells[1,5]:='Четверг';
Excel.Cells[1,6]:='Пятница';
Excel.Cells[1,7]:='Cуббота';

NomPrepod:=0;
NomRow:=3;
while NomPrepod<Length(Prepod) do
  begin
  Excel.Cells[NomRow,1]:=Prepod[NomPrepod].FIO;
  NomDay:=1;
  NomCol:=2;
  while NomDay<=6 do
    begin
    st:='';
    NomTime:=1;
    while NomTime<=6  do
      begin
      NomNagryzka:=0;
      while NomNagryzka<length(Prepod[NomPrepod].Nagryzka) do
        begin
        if (Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid<>'Экзамен') then
        begin
        NomDate:=0;
        while (NomDate<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime)) and not(
              (DayOfWeek(StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDate].StDate))=NomDay+1) and
              (Pos(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDate].StTime,TimeSetPar[NomTime-1])<>0))  do
          inc(NomDate);
        if (Prepod[NomPrepod].Nagryzka[NomNagryzka].sem=NomSem) AND
           (NomDate<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime)) then
           begin
             if st<>'' then
               st:=st+chr(10)+'-----'+chr(10);
             StGroup:='';
             NomGroupNagryzka:=0;
             while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].Group) do
               begin
               StGroup:=StGroup+Prepod[NomPrepod].Nagryzka[NomNagryzka].Group[NomGroupNagryzka].Nom+', ';
               inc(NomGroupNagryzka);
               end;
             if Prepod[NomPrepod].Nagryzka[NomNagryzka].Auditoria<>nil then
               st:=st+TimeSetPar[NomTime-1]+'; '+Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid+'; '+Prepod[NomPrepod].Nagryzka[NomNagryzka].Dis+'; '+StGroup+'; '+Prepod[NomPrepod].Nagryzka[NomNagryzka].Auditoria.Auditoria+' ('+Copy(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[0].StDate,1,5)+'-'+Copy(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime)-1].StDate,1,5)+')';
           end;
        end;
        inc(NomNagryzka);
        end;
      inc(NomTime);
      end;
    Excel.Cells[NomRow,NomCol]:=st;
    inc(NomDay);
    inc(NomCol);
    end;
  Inc(NomRow);
  Inc(NomPrepod);
  end;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,7]].Borders.Weight := 2;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,7]].BorderAround(1);
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,7]].Font.Size:=10;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,7]].WrapText:=true;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,7]].VerticalAlignment:=xlCenter;
//Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRownat,7]].HorizontalAlignment:=xlCenter;
Excel.Workbooks[1].saveas(CurrentDir+'\Расписание кафедры 2.xlsx');
FMain.MeProtocol.Lines.Add('Создан файл с расписанием '+CurrentDir+'\Расписание кафедры 2.xlsx');
Excel.Workbooks.Close;

Excel.WorkBooks.Add;
Excel.Columns[1].ColumnWidth := 4.57;
for NomCol := 2 to 7 do
  Excel.Columns[NomCol].ColumnWidth := 61.14;
for NomTime := 0 to 5 do
  Excel.Cells[1,2+NomTime]:=TimeSetPar[NomTime];

NomRow:=2;
CurrentDate:=StrToDateTime('09.02.2020');
while CurrentDate<StrToDateTime('15.06.2020') do
begin
Excel.Cells[NomRow,1]:=Copy(DateTimeToStr(CurrentDate),1,5);

NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  NomNagryzka:=0;
  while NomNagryzka<length(Prepod[NomPrepod].Nagryzka) do
    begin
    if (Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid<>'Экзамен') then
    begin
    NomDate:=0;
    while (NomDate<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime)) and
          (StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDate].StDate)<>CurrentDate) do
      inc(NomDate);
    if (Prepod[NomPrepod].Nagryzka[NomNagryzka].sem=NomSem) AND                      //Нагрузка с нужным номером семестра и нужной датой
       (NomDate<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime)) and
       (StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDate].StDate)=CurrentDate) then
       begin
         //Определить столбец по времени
         NomTime:=1;
         while (NomTime<=6) and (Pos(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDate].StTime,TimeSetPar[NomTime-1])=0) do
           inc(NomTime);
         If NomTime<=6 then
         begin
         st:=Excel.Cells[NomRow,NomTime+1];
         if st<>'' then
           st:=st+chr(10);
         StGroup:='';
         NomGroupNagryzka:=0;
         while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].Group) do
           begin
           StGroup:=StGroup+Prepod[NomPrepod].Nagryzka[NomNagryzka].Group[NomGroupNagryzka].Nom+', ';
           inc(NomGroupNagryzka);
           end;
         if Prepod[NomPrepod].Nagryzka[NomNagryzka].Auditoria<>nil then
           st:=st+Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid+'; '+Prepod[NomPrepod].Nagryzka[NomNagryzka].Dis+'; '+Copy(Prepod[NomPrepod].FIO,1,pos(' ',Prepod[NomPrepod].FIO))+'; '+StGroup+'; '+Prepod[NomPrepod].Nagryzka[NomNagryzka].Auditoria.Auditoria;
         Excel.Cells[NomRow,NomTime+1]:=st;
         end;
       end;
    end;
    inc(NomNagryzka);
    end;
  Inc(NomPrepod);
  end;
Inc(NomRow);
CurrentDate:=CurrentDate+1;
end;
Excel.Workbooks[1].saveas(CurrentDir+'\Расписание кафедры 3.xlsx');
FMain.MeProtocol.Lines.Add('Создан файл с расписанием '+CurrentDir+'\Расписание кафедры 3.xlsx');
Excel.Workbooks.Close;
end;

Procedure GoRaspisanieToExcel(NomSem:Byte);
var
NomPrepod,NomNagr,NomRow,NomRowBase,NomCol,NomRowNat,NomDate,NomTime:Longword;
NomDayWeek:byte;
st:string;
StGroup:string;
NomGroupNagryzka:Longword;
begin
if not DirectoryExists(CurrentDir+'\Расписание') then
    ForceDirectories(CurrentDir+'\Расписание');
NomPrepod:=0;
ExcelBase.WorkBooks.Add;
ExcelBase.Cells[1,1]:='Преподаватель';
ExcelBase.Cells[1,2]:='Нагр';
ExcelBase.Cells[1,3]:='Предмет';
ExcelBase.Cells[1,4]:='Группа';
ExcelBase.Cells[1,5]:='Ч.';
NomRowBase:=2;

while NomPrepod<Length(Prepod) do
  begin
  Excel.WorkBooks.Add;

  Excel.Columns[1].ColumnWidth := 2.86;
  Excel.Columns[2].ColumnWidth := 5;
  Excel.Columns[3].ColumnWidth := 41.14;
  Excel.Columns[4].ColumnWidth := 13.29;
  Excel.Columns[5].ColumnWidth := 2.71;
  Excel.Columns[6].ColumnWidth := 10.71;
  Excel.Columns[7].ColumnWidth := 6.43;

  NomRow:=3;
  Excel.Cells[1,1]:='Не назначено';
  Excel.Cells[2,2]:='Нагр';
  Excel.Cells[2,3]:='Предмет';
  Excel.Cells[2,4]:='Группа';
  Excel.Cells[2,5]:='Ч.';
  Excel.Cells[2,6]:='Действие';
  Excel.Range[Excel.Cells[2,6],Excel.Cells[2,7]].MergeCells:=true;
  NomNagr:=0;
  while NomNagr<Length(Prepod[NomPrepod].Nagryzka) do
    begin
    //Поиск не назначенной нагрузки
    if (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem) and
       (Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)=0)and
       ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛК') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛР') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ПЗ') or
       (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КП') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КР') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Практика') or
       {(Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Руководство аспирантами') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Руководство магистрами') or} (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Преддипломная практика'))then
      begin
      Excel.Cells[NomRow,2]:=Prepod[NomPrepod].Nagryzka[NomNagr].Vid;
      Excel.Cells[NomRow,3]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
      StGroup:='';
      NomGroupNagryzka:=0;
      while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
        begin
        StGroup:=StGroup+Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka].Nom+', ';
        inc(NomGroupNagryzka);
        end;
      Excel.Cells[NomRow,4]:=StGroup;
      Excel.Cells[NomRow,5]:=Prepod[NomPrepod].Nagryzka[NomNagr].Hour;
      {if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛК') then
        begin
        Excel.Cells[NomRow,6]:='Найти в расписании группы на сайте или около деканата';
        Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,6]].Font.Bold := True;
        end
      else }if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛК') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛР') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ПЗ') then
        begin
        Excel.Cells[NomRow,6]:='Найти в расписании группы на сайте или около деканата';
        Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,6]].Font.Bold := True;
        ExcelBase.Cells[NomRowBase,1]:=Prepod[NomPrepod].FIO;
        ExcelBase.Cells[NomRowBase,2]:=Prepod[NomPrepod].Nagryzka[NomNagr].Vid;
        ExcelBase.Cells[NomRowBase,3]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
        StGroup:='';
        NomGroupNagryzka:=0;
        while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
          begin
          StGroup:=StGroup+Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka].Nom+', ';
          inc(NomGroupNagryzka);
          end;
        ExcelBase.Cells[NomRowBase,4]:=StGroup;
        ExcelBase.Cells[NomRowBase,5]:=Prepod[NomPrepod].Nagryzka[NomNagr].Hour;
        inc(NomRowBase);
        end
      else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КП') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КР') then
        begin
        Excel.Cells[NomRow,6]:='Выдать задание и назначить день консультации. Список выданных заданий прислать на кафедру';
        end
      else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Практика') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Преддипломная практика') then
        begin
        Excel.Cells[NomRow,6]:='Выдать задание и назначить день консультации. Список выданных заданий прислать на кафедру';
        end
      else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Руководство аспирантами') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Руководство магистрами') then
        begin
//        Excel.Cells[NomRow,6]:='Выдать назначить день консультации.';
        end ;
      Excel.Range[Excel.Cells[NomRow,6],Excel.Cells[NomRow,7]].MergeCells:=true;
      inc(NomRow);
      end;
    inc(NomNagr);
    end;
  NomRownat:=2;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow-1,7]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow-1,7]].BorderAround(1);
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow-1,7]].Font.Size:=10;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow-1,7]].WrapText:=true;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow-1,7]].VerticalAlignment:=xlCenter;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRownat,7]].HorizontalAlignment:=xlCenter;

// Вывод расписания по дням недели
  inc(NomRow);
  inc(NomRow);
  NomRowNat:=NomRow;
  Excel.Cells[NomRowNat,1]:='День недели';
  Excel.Cells[NomRowNat,2]:='Время';
  Excel.Cells[NomRowNat,3]:='Предмет';
  Excel.Cells[NomRowNat,4]:='Группа';
  Excel.Cells[NomRowNat,5]:='Вид';
  Excel.Cells[NomRowNat,6]:='Аудит.';
  Excel.Cells[NomRowNat,7]:='Даты';
  inc(NomRow);
  For NomDayWeek:=1 to 7 do
    begin
    For NomTime:= 0 to length(TimeSetPar)-1 do
    begin
     NomNagr:=0;
    while NomNagr<Length(Prepod[NomPrepod].Nagryzka) do
      begin
       if (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem) and
       (Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)<>0)then
       begin
       NomDate:=0;
       while (NomDate<Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)) and  not (
             (DayOfWeek(StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[NomDate].StDate))=NomDayWeek) and
             (Pos(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[NomDate].StTime,TimeSetPar[NomTime])<>0)) do
         inc(NomDate);
       if (NomDate<Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)) then
         begin
         case NomDayWeek of
           2:Excel.Cells[NomRow,1]:='Пн.';
           3:Excel.Cells[NomRow,1]:='Вт.';
           4:Excel.Cells[NomRow,1]:='Ср.';
           5:Excel.Cells[NomRow,1]:='Чт.';
           6:Excel.Cells[NomRow,1]:='Пт.';
           7:Excel.Cells[NomRow,1]:='Сб.';
           1:Excel.Cells[NomRow,1]:='Вс.';
         end;
         st:='';
         st:=st+TimeSetPar[NomTime]+', ';
         Excel.Cells[NomRow,2]:=st;
//         Excel.Cells[NomRow,2]:=Prepod[NomPrepod].Nagryzka[NomNagr].StTime;
         Excel.Cells[NomRow,3]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
         StGroup:='';
         NomGroupNagryzka:=0;
         while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
           begin
           StGroup:=StGroup+Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka].Nom+', ';
           inc(NomGroupNagryzka);
           end;
         Excel.Cells[NomRow,4]:=StGroup;
         Excel.Cells[NomRow,5]:=Prepod[NomPrepod].Nagryzka[NomNagr].Vid;
         Excel.Cells[NomRow,6]:=Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria;
         st:='';
         NomDate:=0;
         while (NomDate<Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)) do
           begin
           if (DayOfWeek(StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[NomDate].StDate))=NomDayWeek) and
              (Pos(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[NomDate].StTime,TimeSetPar[NomTime])<>0) then
             st:=st+Copy(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[NomDate].StDate,1,5)+', ';
           inc(NomDate);
           end;
         Excel.Cells[NomRow,7]:=st;
         if Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛК' then
           Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,7]].Font.Bold := True;
         inc(NomRow);
         end;

       end;
      inc(NomNagr);
      end;
    end;
    end;

  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,7]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,7]].BorderAround(1);
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,7]].Font.Size:=10;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,7]].WrapText:=true;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,7]].VerticalAlignment:=xlCenter;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRownat,7]].HorizontalAlignment:=xlCenter;


  Excel.Workbooks[1].saveas(CurrentDir+'\Расписание\'+Prepod[NomPrepod].FIO+'.xlsx');
  FMain.MeProtocol.Lines.Add('Создан файл с расписанием '+CurrentDir+'\Расписание\'+Prepod[NomPrepod].FIO+'.xlsx');
  Excel.Workbooks.Close;
  inc(NomPrepod);
  end;

  ExcelBase.Workbooks[1].saveas(CurrentDir+'\Расписание\Ошибки.xlsx');
  FMain.MeProtocol.Lines.Add('Создан файл с расписанием '+CurrentDir+'\Расписание\Ошибки.xlsx');
  ExcelBase.Workbooks.Close;
{
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  Excel.WorkBooks.Add;
  For NomDayWeek:=1 to 31 do
    begin
    NomNagr:=0;
    while NomNagr<Length(Prepod[NomPrepod].Nagryzka) do
      begin
      if (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem) and
           (Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDate)<>0)then
        begin

        end;
      inc(NomNagr);
      end;
    end;
//  Excel.Workbooks[1].saveas(CurrentDir+'/Расписание 2/'+Prepod[NomPrepod].FIO+'.xlsx');
  FMain.MeProtocol.Lines.Add('Создан файл с расписанием '+CurrentDir+'/Расписание 2/'+Prepod[NomPrepod].FIO+'.xlsx');
  Excel.Workbooks.Close;
  inc(NomPrepod);
  end;          }

end;

procedure GoRaspToExcel(NomSem:Byte);
var
  NomPrepod,NomNagr,NomRow,NomCol,NomRownat:Longword;
  DataCons,CurrentDate:TDateTime;
  MinDate,MaxDate:array [0..1] of TDateTime;
  st,StGroup:string;
  EK:byte;
  NomGroupNagryzka:Longword;
  EnterEkz:boolean;
begin
MinDate[0]:=0;
MaxDate[0]:=0;
MinDate[1]:=0;
MaxDate[1]:=0;
for Ek := 0 to  2 do
begin

Excel.WorkBooks.Add;
if EK=0 then
begin
Excel.Columns[1].ColumnWidth := 35.86;
Excel.Columns[2].ColumnWidth := 12.00;
Excel.Columns[3].ColumnWidth := 5.00;
Excel.Columns[4].ColumnWidth := 5.43;
Excel.Columns[5].ColumnWidth := 9.57;
Excel.Columns[6].ColumnWidth := 5.00;
Excel.Columns[7].ColumnWidth := 5.43;
Excel.Columns[8].ColumnWidth := 9.57;
end
else
begin
Excel.Columns[1].ColumnWidth := 40.71;
Excel.Columns[2].ColumnWidth := 12.71;
Excel.Columns[3].ColumnWidth := 8.43;
Excel.Columns[4].ColumnWidth := 9.29;
Excel.Columns[5].ColumnWidth := 7.00;
Excel.Columns[6].ColumnWidth := 11.71;
end;

NomPrepod:=0;
NomRow:=1;
while NomPrepod<Length(Prepod) do
  begin
  NomNagr:=0;
  while (NomNagr<Length(Prepod[NomPrepod].Nagryzka)) and not (
       (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem) )do
     inc(NomNagr);
  if NomNagr<Length(Prepod[NomPrepod].Nagryzka) then
  begin
  Excel.Cells[NomRow,2]:=Prepod[NomPrepod].FIO;
  Excel.Cells[NomRow,1]:=Prepod[NomPrepod].Dolzhnost;
  Excel.Cells[NomRow+1,1]:='Предмет';
  case Ek of
    0: Excel.Cells[NomRow+1,2]:='Экзамены';
    1: Excel.Cells[NomRow+1,2]:='Зачеты';
    2: Excel.Cells[NomRow+1,2]:='Курсовые/Практика';
  end;
  if Ek=0 then
    begin
    Excel.Cells[NomRow+1,6]:='Консультация';
    Excel.Cells[NomRow+2,2]:='Группа';
    Excel.Cells[NomRow+2,3]:='Дата';
    Excel.Cells[NomRow+2,4]:='Время';
    Excel.Cells[NomRow+2,5]:='Аудитория';
    Excel.Cells[NomRow+2,6]:='Дата';
    Excel.Cells[NomRow+2,7]:='Время';
    Excel.Cells[NomRow+2,8]:='Аудитория';
    Excel.Range[Excel.Cells[NomRow+1,2],Excel.Cells[NomRow+1,5]].MergeCells:=true;
    Excel.Range[Excel.Cells[NomRow+1,6],Excel.Cells[NomRow+1,8]].MergeCells:=true;
    Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,8]].MergeCells:=true;
    end
  else
    begin
    Excel.Cells[NomRow+2,2]:='Группа';
    Excel.Cells[NomRow+2,3]:='Тип';
    Excel.Cells[NomRow+2,4]:='Дата';
    Excel.Cells[NomRow+2,5]:='Время';
    Excel.Cells[NomRow+2,6]:='Аудитория';
    Excel.Range[Excel.Cells[NomRow+1,2],Excel.Cells[NomRow+1,6]].MergeCells:=true;
    Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,6]].MergeCells:=true;
    end;
  Excel.Range[Excel.Cells[NomRow+1,1],Excel.Cells[NomRow+2,1]].MergeCells:=true;

  NomRownat:=NomRow+1;
  NomRow:=NomRow+2;


  NomNagr:=0;
  while NomNagr<Length(Prepod[NomPrepod].Nagryzka) do
    begin
    if (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem) and  (
       ((Ek=0) and (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Экзамен')) or
       ((Ek=1) and ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Зачет') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Зачет с оценкой'))) or
       ((Ek=2) and ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КР') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КП') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Практика')))) then
      begin
      inc(NomRow);
      Excel.Cells[NomRow,1]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
      StGroup:='';
      NomGroupNagryzka:=0;
      while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
        begin
        StGroup:=StGroup+Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka].Nom+', ';
        inc(NomGroupNagryzka);
        end;
      Excel.Cells[NomRow,2]:=Copy(StGroup,1,Length(StGroup)-2);
      if Ek<>0 then
        Excel.Cells[NomRow,3]:=Prepod[NomPrepod].Nagryzka[NomNagr].Vid;
      if (Ek=0) and (Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)<>0) then
        begin
        Excel.Cells[NomRow,3]:=Copy(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate,1,5);

        if (MinDate[0]=0) or (StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate)<MinDate[0]) then
          MinDate[0]:=StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate);
        if (MaxDate[0]=0) or (StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate)>MaxDate[0]) then
          MaxDate[0]:=StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate);

        Excel.Cells[NomRow,4]:=Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StTime;
        Excel.Cells[NomRow,5]:=Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria;
//        ShortDateFormat := 'dd.mm';
        DataCons:=StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate)-1;
        while (DayOfWeek(DataCons)=7) or (DayOfWeek(DataCons)=6) do
          DataCons:=DataCons-1;
        SetLength(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime,2);
        Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[1].StDate:=DateTimeToStr(DataCons);

        if (MinDate[1]=0) or (StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[1].StDate)<MinDate[1]) then
          MinDate[1]:=StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[1].StDate);
        if (MaxDate[1]=0) or (StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[1].StDate)>MaxDate[1]) then
          MaxDate[1]:=StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[1].StDate);

        Excel.Cells[NomRow,6]:=Copy(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[1].StDate,1,5);
        Excel.Cells[NomRow,7]:='10:00';
        Excel.Cells[NomRow,8]:='каф.';
        end;
      end;
    inc(NomNagr);
    end;
  if Ek=0 then
  begin
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,8]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,8]].BorderAround(1);
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,8]].Font.Size:=10;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,8]].WrapText:=true;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,8]].VerticalAlignment:=xlCenter;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRownat+1,8]].HorizontalAlignment:=xlCenter;
  end
  else
  begin
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,6]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,6]].BorderAround(1);
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,6]].Font.Size:=10;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,6]].WrapText:=true;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,6]].VerticalAlignment:=xlCenter;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRownat+1,6]].HorizontalAlignment:=xlCenter;
  end;
  NomRow:=NomRow+2;
  end;
  inc(NomPrepod);

  end;
if Ek=0 then
  begin
  if NomSem=1 then
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Экзамены_осень.xlsx')
  else
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Экзамены_весна.xlsx');
  end;
if Ek=1 then
  begin
  if NomSem=1 then
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Зачеты_осень.xlsx')
  else
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Зачеты_весна.xlsx');
  end;
if Ek=2 then
  begin
  if NomSem=1 then
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Курсовые_осень.xlsx')
  else
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Курсовые_весна.xlsx');
  end;
Excel.Workbooks.Close;
end;

for Ek := 0 to  1 do
begin
Excel.WorkBooks.Add;
Excel.Columns[1].ColumnWidth := 12.57;

NomCol:=2;
CurrentDate:=MinDate[Ek];
while CurrentDate<=MaxDate[Ek] do
  begin
//  ShortDateFormat := 'dd.mm';
  Excel.Cells[1,NomCol]:=Copy(DateTimeToStr(CurrentDate),1,5);
  Excel.Columns[NomCol].ColumnWidth := 8.43;
  inc(NomCol);
  CurrentDate:=CurrentDate+1;
  end;
NomRow:=2;
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  NomNagr:=0;
  while (NomNagr<Length(Prepod[NomPrepod].Nagryzka)) and not ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Экзамен') and (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem)) do
     inc(NomNagr);
  if NomNagr<Length(Prepod[NomPrepod].Nagryzka) then
  begin
  Excel.Cells[NomRow,1]:=Prepod[NomPrepod].FIO;
  NomNagr:=0;
  while NomNagr<Length(Prepod[NomPrepod].Nagryzka) do
    begin
    if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Экзамен') and
       (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem) and
       (Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)<>0) then
      begin
      NomCol:=Trunc(StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[Ek].StDate)-MinDate[Ek]+2);
      StGroup:='';
      NomGroupNagryzka:=0;
      while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
        begin
        StGroup:=StGroup+Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka].Nom+', ';
        inc(NomGroupNagryzka);
        end;
      st:=Excel.Cells[NomRow,NomCol];
//      ShortTimeFormat:= 'hh:mm';
      if Ek=0 then
        st:=st+'  '+stGroup+'  '+Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StTime+'  '+Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria
      else
        st:=st+'  '+StGroup+'  '+'10:00'+'  '+'каф.';

      Excel.Cells[NomRow,NomCol]:=st;
      end;
    inc(NomNagr);

    end;
  inc(NomRow);
  end;
  inc(NomPrepod);
  end;
NomCol:=Trunc(MaxDate[Ek]-MinDate[Ek]+2);
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,NomCol]].Borders.Weight := 2;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,NomCol]].BorderAround(1);
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,1]].Font.Size:=10;
Excel.Range[Excel.Cells[1,1],Excel.Cells[1,NomCol]].Font.Size:=10;
Excel.Range[Excel.Cells[2,2],Excel.Cells[NomRow,NomCol]].Font.Size:=8;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,NomCol]].WrapText:=true;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,NomCol]].VerticalAlignment:=xlCenter;
Excel.Range[Excel.Cells[1,2],Excel.Cells[NomRow,NomCol]].HorizontalAlignment:=xlCenter;
if NomSem=1 then
  if Ek=0 then
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Экзамены_Таблица_осень.xlsx')
  else
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Консультации_Таблица_осень.xlsx')
else
  if Ek=0 then
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Экзамены_Таблица_весна.xlsx')
  else
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Консультации_Таблица_весна.xlsx');
Excel.Workbooks.Close;
end;
end;

Procedure GoTablExzamenGroupToExcel(NomSem:Byte);
var
NomTypeGroup,NomGroup:Longword;
NomPrepod,NomNagr,NomPrepodToo,NomNagrToo:Longword;
NomRow,NomRowNat,Nom,NomStartTabled:Longword;
ArrNagr:array[1..3] of byte;
MaxStudentAuditoria:Longword;
i:byte;
f:TextFile;
st:string;

begin
NomTypeGroup:=0;
while NomTypeGroup<length(NameAllGroup) do
  begin
  if (NomTypeGroup<length(NameAllGroup)) and (NameAllGroup[NomTypeGroup].NameGroupKyrs<>'') then
  begin
  NomGroup:=0;
  while NomGroup<Length(NameAllGroup[NomTypeGroup].Group) do


  begin
  NomPrepod:=0;

  Excel.WorkBooks.Add;

  Excel.Columns[1].ColumnWidth := 9.43;
  Excel.Columns[2].ColumnWidth := 12.71;
  Excel.Columns[3].ColumnWidth := 17;
  Excel.Columns[4].ColumnWidth := 14.71;
  Excel.Columns[5].ColumnWidth := 10.57;
  Excel.Columns[6].ColumnWidth := 10.86;
  Excel.Columns[7].ColumnWidth := 17.14;

  Excel.Range[Excel.Cells[1,1],Excel.Cells[1,7]].MergeCells:=true;
  Excel.Cells[1,1]:='Рассписание промежуточной аттестации (зимний семестр 2020-2021 учебный год)';
  Excel.Range[Excel.Cells[2,1],Excel.Cells[2,7]].MergeCells:=true;
  Excel.Cells[2,1]:='Учебная группа '+NameAllGroup[NomTypeGroup].Group[NomGroup];

  Excel.Cells[3,1]:='Дата';
  Excel.Cells[3,2]:='Вид';
  Excel.Cells[3,3]:='Дисциплина';
  Excel.Cells[3,4]:='ФИО преподавателя';
  Excel.Cells[3,5]:='Время начала, продолжи-тельность';
  Excel.Cells[3,6]:='Платформа,ресурс';
  Excel.Cells[3,7]:='Формат проведения';

  NomRow:=4;
  Nom:=1;
    NomPrepod:=0;
    while NomPrepod<length(Prepod) do
      begin
      NomNagr:=0;
      while NomNagr<length(Prepod[NomPrepod].Nagryzka) do
        begin
        if (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem) and ((SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagr].Group, NameAllGroup[NomTypeGroup].Group[NomGroup])<>65000)) and
           ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КР') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КП') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Практика')) then
          begin
          Excel.Cells[NomRow,1]:='';
          Excel.Cells[NomRow,2]:=Prepod[NomPrepod].Nagryzka[NomNagr].Vid;
          Excel.Cells[NomRow,3]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
          Excel.Cells[NomRow,4]:=Prepod[NomPrepod].FIO;
          Excel.Cells[NomRow,5]:='9:00, 3 часа';
          Excel.Cells[NomRow,6]:='LMS';
          Excel.Cells[NomRow,7]:='Видоконференция с фиксацией ответов обучающегося';
          inc(NomRow);
          end;
        inc(NomNagr);
        end;
      inc(NomPrepod);
      end;
    NomPrepod:=0;
    while NomPrepod<length(Prepod) do
      begin
      NomNagr:=0;
      while NomNagr<length(Prepod[NomPrepod].Nagryzka) do
        begin
        if (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem) and ((SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagr].Group, NameAllGroup[NomTypeGroup].Group[NomGroup])<>65000)) and
           ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Зачет') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Зачет с оценкой')) then
          begin
          Excel.Cells[NomRow,1]:='';
          Excel.Cells[NomRow,2]:=Prepod[NomPrepod].Nagryzka[NomNagr].Vid;
          Excel.Cells[NomRow,3]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
          Excel.Cells[NomRow,4]:=Prepod[NomPrepod].FIO;
          Excel.Cells[NomRow,5]:='9:00, 3 часа';
          Excel.Cells[NomRow,6]:='LMS';
          Excel.Cells[NomRow,7]:='Видоконференция с фиксацией ответов обучающегося';
          inc(NomRow);
          end;
        inc(NomNagr);
        end;
      inc(NomPrepod);
      end;
    NomPrepod:=0;
    while NomPrepod<length(Prepod) do
      begin
      NomNagr:=0;
      while NomNagr<length(Prepod[NomPrepod].Nagryzka) do
        begin
          if (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem) and ((SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagr].Group, NameAllGroup[NomTypeGroup].Group[NomGroup])<>65000)) and
           (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Экзамен') then
          begin
          if (Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)<>0) then
            Excel.Cells[NomRow,1]:=StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[1].StDate);
          Excel.Cells[NomRow,2]:='Консультация к экзамену';
          Excel.Cells[NomRow,3]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
          Excel.Cells[NomRow,4]:=Prepod[NomPrepod].FIO;
          Excel.Cells[NomRow,5]:='10:00, 3 часа';
          Excel.Cells[NomRow,6]:='LMS';
          Excel.Cells[NomRow,7]:='Консультация';
          inc(NomRow);

          if (Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)<>0) then
            Excel.Cells[NomRow,1]:=StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate);
          Excel.Cells[NomRow,2]:='Экзамен';
          Excel.Cells[NomRow,3]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
          Excel.Cells[NomRow,4]:=Prepod[NomPrepod].FIO;
          if (Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)<>0)then
            Excel.Cells[NomRow,5]:=Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StTime+', 3 часа';
          Excel.Cells[NomRow,6]:='LMS';
          Excel.Cells[NomRow,7]:='Видоконференция с фиксацией ответов обучающегося';
          inc(NomRow);
          end;
        inc(NomNagr);
        end;
      inc(NomPrepod);
      end;

  Excel.Range[Excel.Cells[3,3],Excel.Cells[NomRow,3]].Font.Size:=10;
  Excel.Range[Excel.Cells[3,7],Excel.Cells[NomRow,7]].Font.Size:=10;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,7]].WrapText:=true;
  Excel.Range[Excel.Cells[3,1],Excel.Cells[NomRow,7]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[3,1],Excel.Cells[NomRow,7]].BorderAround(1);
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,7]].VerticalAlignment:=xlCenter;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,7]].HorizontalAlignment:=xlCenter;
  if not DirectoryExists(CurrentDir+'\Экзамены\По группам') then
    ForceDirectories(CurrentDir+'\Экзамены\По группам');
  if (NomTypeGroup<length(NameAllGroup)) and (NameAllGroup[NomTypeGroup].Group[NomGroup]<>'') then
    Excel.Workbooks[1].saveas(CurrentDir+'\Экзамены\По группам\'+NameAllGroup[NomTypeGroup].Group[NomGroup]+'.xlsx');
  Fmain.MeProtocol.Lines.Add('Создан файл '+CurrentDir+'\Экзамены\По группам\'+NameAllGroup[NomTypeGroup].Group[NomGroup]+'.xlsx');
  Excel.Workbooks.Close;
    inc(NomGroup);
  end;

  end;
  inc(NomTypeGroup);
  end;
end;

procedure GoExzamenToExcel(NomSem:Byte);
var
  NomPrepod,NomNagr,NomRow,NomCol,NomRownat:Longword;
  DataCons,CurrentDate:TDateTime;
  MinDate,MaxDate:array [0..1] of TDateTime;
  st,StGroup:string;
  EK:byte;
  NomGroupNagryzka:Longword;
  EnterEkz:boolean;
begin
MinDate[0]:=0;
MaxDate[0]:=0;
MinDate[1]:=0;
MaxDate[1]:=0;
for Ek := 0 to  2 do
begin

Excel.WorkBooks.Add;
if EK=0 then
begin
Excel.Columns[1].ColumnWidth := 35.86;
Excel.Columns[2].ColumnWidth := 12.00;
Excel.Columns[3].ColumnWidth := 5.00;
Excel.Columns[4].ColumnWidth := 5.43;
Excel.Columns[5].ColumnWidth := 9.57;
Excel.Columns[6].ColumnWidth := 5.00;
Excel.Columns[7].ColumnWidth := 5.43;
Excel.Columns[8].ColumnWidth := 9.57;
end
else
begin
Excel.Columns[1].ColumnWidth := 40.71;
Excel.Columns[2].ColumnWidth := 12.71;
Excel.Columns[3].ColumnWidth := 8.43;
Excel.Columns[4].ColumnWidth := 9.29;
Excel.Columns[5].ColumnWidth := 7.00;
Excel.Columns[6].ColumnWidth := 11.71;
end;

NomPrepod:=0;
NomRow:=1;
while NomPrepod<Length(Prepod) do
  begin
  NomNagr:=0;
  while (NomNagr<Length(Prepod[NomPrepod].Nagryzka)) and not (
       (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem) and (
       ((Ek=0) and (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Экзамен')) or
       ((Ek=1) and ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Зачет') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Зачет с оценкой'))) or
       ((Ek=2) and ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КР') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КП') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Практика'))))) do
     inc(NomNagr);
  if NomNagr<Length(Prepod[NomPrepod].Nagryzka) then
  begin
  Excel.Cells[NomRow,2]:=Prepod[NomPrepod].FIO;
  Excel.Cells[NomRow,1]:=Prepod[NomPrepod].Dolzhnost;
  Excel.Cells[NomRow+1,1]:='Предмет';
  case Ek of
    0: Excel.Cells[NomRow+1,2]:='Экзамены';
    1: Excel.Cells[NomRow+1,2]:='Зачеты';
    2: Excel.Cells[NomRow+1,2]:='Курсовые/Практика';
  end;
  if Ek=0 then
    begin
    Excel.Cells[NomRow+1,6]:='Консультация';
    Excel.Cells[NomRow+2,2]:='Группа';
    Excel.Cells[NomRow+2,3]:='Дата';
    Excel.Cells[NomRow+2,4]:='Время';
    Excel.Cells[NomRow+2,5]:='Аудитория';
    Excel.Cells[NomRow+2,6]:='Дата';
    Excel.Cells[NomRow+2,7]:='Время';
    Excel.Cells[NomRow+2,8]:='Аудитория';
    Excel.Range[Excel.Cells[NomRow+1,2],Excel.Cells[NomRow+1,5]].MergeCells:=true;
    Excel.Range[Excel.Cells[NomRow+1,6],Excel.Cells[NomRow+1,8]].MergeCells:=true;
    Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,8]].MergeCells:=true;
    end
  else
    begin
    Excel.Cells[NomRow+2,2]:='Группа';
    Excel.Cells[NomRow+2,3]:='Тип';
    Excel.Cells[NomRow+2,4]:='Дата';
    Excel.Cells[NomRow+2,5]:='Время';
    Excel.Cells[NomRow+2,6]:='Аудитория';
    Excel.Range[Excel.Cells[NomRow+1,2],Excel.Cells[NomRow+1,6]].MergeCells:=true;
    Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,6]].MergeCells:=true;
    end;
  Excel.Range[Excel.Cells[NomRow+1,1],Excel.Cells[NomRow+2,1]].MergeCells:=true;

  NomRownat:=NomRow+1;
  NomRow:=NomRow+2;


  NomNagr:=0;
  while NomNagr<Length(Prepod[NomPrepod].Nagryzka) do
    begin
    if (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem) and  (
       ((Ek=0) and (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Экзамен')) or
       ((Ek=1) and ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Зачет') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Зачет с оценкой'))) or
       ((Ek=2) and ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КР') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КП') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Практика')))) then
      begin
      inc(NomRow);
      Excel.Cells[NomRow,1]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
      StGroup:='';
      NomGroupNagryzka:=0;
      while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
        begin
        StGroup:=StGroup+Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka].Nom+', ';
        inc(NomGroupNagryzka);
        end;
      Excel.Cells[NomRow,2]:=Copy(StGroup,1,Length(StGroup)-2);
      if Ek<>0 then
        Excel.Cells[NomRow,3]:=Prepod[NomPrepod].Nagryzka[NomNagr].Vid;
      if (Ek=0) and (Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)<>0) then
        begin
        Excel.Cells[NomRow,3]:=Copy(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate,1,5);

        if (MinDate[0]=0) or (StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate)<MinDate[0]) then
          MinDate[0]:=StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate);
        if (MaxDate[0]=0) or (StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate)>MaxDate[0]) then
          MaxDate[0]:=StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate);

        Excel.Cells[NomRow,4]:=Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StTime;
        Excel.Cells[NomRow,5]:=Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria;
//        ShortDateFormat := 'dd.mm';
        DataCons:=StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StDate)-2;
        while (DayOfWeek(DataCons)=7) or (DayOfWeek(DataCons)=6) do
          DataCons:=DataCons-1;
        SetLength(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime,2);
        Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[1].StDate:=DateTimeToStr(DataCons);

        if (MinDate[1]=0) or (StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[1].StDate)<MinDate[1]) then
          MinDate[1]:=StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[1].StDate);
        if (MaxDate[1]=0) or (StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[1].StDate)>MaxDate[1]) then
          MaxDate[1]:=StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[1].StDate);

        Excel.Cells[NomRow,6]:=Copy(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[1].StDate,1,5);
        Excel.Cells[NomRow,7]:='10:00';
        Excel.Cells[NomRow,8]:='каф.';
        end;
      end;
    inc(NomNagr);
    end;
  if Ek=0 then
  begin
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,8]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,8]].BorderAround(1);
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,8]].Font.Size:=10;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,8]].WrapText:=true;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,8]].VerticalAlignment:=xlCenter;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRownat+1,8]].HorizontalAlignment:=xlCenter;
  end
  else
  begin
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,6]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,6]].BorderAround(1);
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,6]].Font.Size:=10;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,6]].WrapText:=true;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRow,6]].VerticalAlignment:=xlCenter;
  Excel.Range[Excel.Cells[NomRownat,1],Excel.Cells[NomRownat+1,6]].HorizontalAlignment:=xlCenter;
  end;
  NomRow:=NomRow+2;
  end;
  inc(NomPrepod);

  end;
if Ek=0 then
  begin
  if NomSem=1 then
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Экзамены_осень.xlsx')
  else
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Экзамены_весна.xlsx');
  end;
if Ek=1 then
  begin
  if NomSem=1 then
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Зачеты_осень.xlsx')
  else
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Зачеты_весна.xlsx');
  end;
if Ek=2 then
  begin
  if NomSem=1 then
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Курсовые_осень.xlsx')
  else
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Курсовые_весна.xlsx');
  end;
Excel.Workbooks.Close;
end;

for Ek := 0 to  1 do
begin
Excel.WorkBooks.Add;
Excel.Columns[1].ColumnWidth := 12.57;

NomCol:=2;
CurrentDate:=MinDate[Ek];
while CurrentDate<=MaxDate[Ek] do
  begin
//  ShortDateFormat := 'dd.mm';
  Excel.Cells[1,NomCol]:=Copy(DateTimeToStr(CurrentDate),1,5);
  Excel.Columns[NomCol].ColumnWidth := 8.43;
  inc(NomCol);
  CurrentDate:=CurrentDate+1;
  end;
NomRow:=2;
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  NomNagr:=0;
  while (NomNagr<Length(Prepod[NomPrepod].Nagryzka)) and not ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Экзамен') and (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem)) do
     inc(NomNagr);
  if NomNagr<Length(Prepod[NomPrepod].Nagryzka) then
  begin
  Excel.Cells[NomRow,1]:=Prepod[NomPrepod].FIO;
  NomNagr:=0;
  while NomNagr<Length(Prepod[NomPrepod].Nagryzka) do
    begin
    if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Экзамен') and
       (Prepod[NomPrepod].Nagryzka[NomNagr].sem=NomSem) and
       (Length(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime)<>0) then
      begin
      NomCol:=Trunc(StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[Ek].StDate)-MinDate[Ek]+2);
      StGroup:='';
      NomGroupNagryzka:=0;
      while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
        begin
        StGroup:=StGroup+Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka].Nom+', ';
        inc(NomGroupNagryzka);
        end;
      st:=Excel.Cells[NomRow,NomCol];
//      ShortTimeFormat:= 'hh:mm';
      if Ek=0 then
        st:=st+'  '+stGroup+'  '+Prepod[NomPrepod].Nagryzka[NomNagr].StDateTime[0].StTime+'  '+Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria
      else
        st:=st+'  '+StGroup+'  '+'10:00'+'  '+'каф.';

      Excel.Cells[NomRow,NomCol]:=st;
      end;
    inc(NomNagr);

    end;
  inc(NomRow);
  end;
  inc(NomPrepod);
  end;
NomCol:=Trunc(MaxDate[Ek]-MinDate[Ek]+2);
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,NomCol]].Borders.Weight := 2;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,NomCol]].BorderAround(1);
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,1]].Font.Size:=10;
Excel.Range[Excel.Cells[1,1],Excel.Cells[1,NomCol]].Font.Size:=10;
Excel.Range[Excel.Cells[2,2],Excel.Cells[NomRow,NomCol]].Font.Size:=8;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,NomCol]].WrapText:=true;
Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,NomCol]].VerticalAlignment:=xlCenter;
Excel.Range[Excel.Cells[1,2],Excel.Cells[NomRow,NomCol]].HorizontalAlignment:=xlCenter;
if NomSem=1 then
  if Ek=0 then
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Экзамены_Таблица_осень.xlsx')
  else
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Консультации_Таблица_осень.xlsx')
else
  if Ek=0 then
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Экзамены_Таблица_весна.xlsx')
  else
    Excel.Workbooks[1].saveas(CurrentDir+'/Экзамены/Консультации_Таблица_весна.xlsx');
Excel.Workbooks.Close;
end;
end;

Procedure SaveAllPrepodGroup;
var
NomTypeGroup,NomGroup:Longword;
NomPrepod,NomNagr,NomPrepodToo,NomNagrToo:Longword;
NomRow,NomRowNat,Nom,NomStartTabled:Longword;
ArrNagr:array[1..3] of byte;
MaxStudentAuditoria:Longword;
i:byte;
f:TextFile;
st:string;
begin
NomTypeGroup:=0;
while NomTypeGroup<length(NameAllGroup) do
  begin
  if (NomTypeGroup<length(NameAllGroup)) and (NameAllGroup[NomTypeGroup].NameGroupKyrs<>'') then
  begin
  NomPrepod:=0;

  Excel.WorkBooks.Add;

  Excel.Columns[1].ColumnWidth := 3.43;
  Excel.Columns[2].ColumnWidth := 44;
  Excel.Columns[3].ColumnWidth := 3.86;
  Excel.Columns[4].ColumnWidth := 33.14;
  Excel.Columns[5].ColumnWidth := 10.43;
  Excel.Columns[6].ColumnWidth := 8.14;
  Excel.Columns[7].ColumnWidth := 7.14;

  Excel.Range[Excel.Cells[1,1],Excel.Cells[1,6]].MergeCells:=true;
  Excel.Cells[1,1]:='Федеральное государственное бюджетное образовательное учреждение высшего образования';
  Excel.Range[Excel.Cells[2,1],Excel.Cells[2,6]].MergeCells:=true;
  Excel.Cells[2,1]:='«Московский авиационный институт (национальный исследовательский университет)»';
  Excel.Range[Excel.Cells[4,1],Excel.Cells[4,6]].MergeCells:=true;
  Excel.Cells[4,1]:='Справка';
  Excel.Range[Excel.Cells[5,1],Excel.Cells[5,6]].MergeCells:=true;
  st:='о преподовательском обеспечении основной образовательной программы высшего образования – программы';
  case NameAllGroup[NomTypeGroup].NameGroupKyrs[3] of
    'Б' :Excel.Cells[5,1]:=st+' бакалавриата';
    'М' :Excel.Cells[5,1]:=st+' магистратуры';
    'А' :Excel.Cells[5,1]:=st+' аспирантуры';
    'С' :Excel.Cells[5,1]:=st+' специалитета';
  else
    Excel.Cells[5,1]:=st;
  end;
  Excel.Range[Excel.Cells[6,1],Excel.Cells[6,6]].MergeCells:=true;
  if NameAllGroup[NomTypeGroup].NameGroupKyrs='07Б' then
    Excel.Cells[6,1]:='09.03.01 "Информатика и вычислительная техника" профиль - Автоматизированные системы обработки информации и управления'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='09Б' then
    Excel.Cells[6,1]:='09.03.01 "Информатика и вычислительная техника" профиль - Вычислительные машины, комплекы, системы и сети'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='10Б' then
    Excel.Cells[6,1]:='09.03.01 "Информатика и вычислительная техника" профиль - Программное обеспечение средств вычислительной техники и автоматизированных систем'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='11Б' then
    Excel.Cells[6,1]:='09.03.04 "Программная инженерия" профиль - Программно-информационные системы'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='05М' then
    Excel.Cells[6,1]:='09.04.01 "Информатика и вычислительная техника" профиль - Автоматизированные системы обработки информации и управления'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='06М' then
    Excel.Cells[6,1]:='09.04.01 "Информатика и вычислительная техника" профиль - Вычислительные машины, комплекы, системы и сети'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='07М' then
    Excel.Cells[6,1]:='09.04.01 "Информатика и вычислительная техника" профиль - Программное обеспечение средств вычислительной техники и автоматизированных систем'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='08М' then
    Excel.Cells[6,1]:='09.04.04 "Программная инженерия" профиль - Программно-информационные системы' ;
  Excel.Cells[8,1]:='№ п\п';
  Excel.Cells[8,2]:='Наименование дисциплины (модуля), практик в соответствии с учебным планом';
  Excel.Cells[8,3]:='Вид';
  Excel.Cells[8,4]:='ФИО преподавателя';
  Excel.Cells[8,5]:='Должность';
  Excel.Cells[8,6]:='Степень';
  Excel.Cells[8,7]:='Звание';

  NomRow:=9;
  Nom:=1;
    NomPrepod:=0;
    while NomPrepod<length(Prepod) do
      begin
      NomNagr:=0;
      while NomNagr<length(Prepod[NomPrepod].Nagryzka) do
        begin
        if ((SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagr].Group, NameAllGroup[NomTypeGroup].NameGroupKyrs)<>65000)) and
           ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛК') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛР') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ПЗ')) then
          begin
          Excel.Cells[NomRow,1]:=Nom;
          Excel.Cells[NomRow,2]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
          Excel.Cells[NomRow,3]:=Prepod[NomPrepod].Nagryzka[NomNagr].Vid;
          Excel.Cells[NomRow,4]:=Prepod[NomPrepod].FIO;
          Excel.Cells[NomRow,5]:=Prepod[NomPrepod].Dolzhnost;
          Excel.Cells[NomRow,6]:=Prepod[NomPrepod].Stepen;
          Excel.Cells[NomRow,7]:=Prepod[NomPrepod].Zvanie;
          inc(NomRow);
          inc(Nom);
          end;
        inc(NomNagr);
        end;
      inc(NomPrepod);
      end;

  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow,7]].WrapText:=true;
  Excel.Range[Excel.Cells[8,1],Excel.Cells[NomRow,7]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[8,1],Excel.Cells[NomRow,7]].BorderAround(1);
  if not DirectoryExists(CurrentDir+'\Преподаватели в группах') then
    ForceDirectories(CurrentDir+'\Преподаватели в группах');
  if (NomTypeGroup<length(NameAllGroup)) and (NameAllGroup[NomTypeGroup].NameGroupKyrs<>'') then
    Excel.Workbooks[1].saveas(CurrentDir+'\Преподаватели в группах\'+NameAllGroup[NomTypeGroup].NameGroupKyrs+'.xlsx');
  Fmain.MeProtocol.Lines.Add('Создан файл '+CurrentDir+'\Преподаватели в группах\'+NameAllGroup[NomTypeGroup].NameGroupKyrs+'.xlsx');
  Excel.Workbooks.Close;
  end;
  inc(NomTypeGroup);
  end;
end;

end.
