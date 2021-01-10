unit UNagryzka;

interface

uses UGroup,UAuditoria,UConstParametrs,StdCtrls;

type

  TPrepodAll = class;
  TNagryzkaOb = record
                Sem:byte;
                NomRow:Longword;
                P:byte;
                Dis,SOKR,Vid,Group,Hour,FIOPrep,NOMPrep,Opisanie:string;
                KolStudent:longword;
                end;
  TStrDateTime = record
   StDate,StTime:string;
  end;

  TNagryzkaPrepod = class
                    Prepod:TPrepodAll;
                    sem:byte;
                    NomRow:Longword;
                    P:byte;
                    Dis,SOKR,Vid,Hour,NOMPrep,Opisanie:string;
                    KolStudent:Longword;
                    Group:TAGroup;
                    TypeVid:Byte;
                    Auditoria:TAuditoria;
                    FlagIndPlan:byte;
                    PO:string;
                    FlagVivoda:boolean;
                    StDateTime:array of TStrDateTime;
                    Constructor Create;
                    Destructor Destroy;
                    end;

  TPrepodAll = class
               NameFilePrepod:string;
               FIO,Dolzhnost,Stepen,Zvanie,PovKvalProsh:String;
               PovKval:Longword;
               P,FlagP:byte;
               Nagryzka:array of TNagryzkaPrepod;
               Pochasovka:Double;
               Stavka,StavkaSovmest:String;
               MesNeOplat:byte;
               AllHour,AllHourPrav,AllHourVne:Double;
               HourSem:array [1..kolsem] of Double;
               Constructor Create;
               Destructor Destroy;
               end;

  TMergeDiscipline = record
    Dis,Vid:string;
  end;

var
  Nagryzka:Array of TNagryzkaOb;
  Prepod:array of TPrepodAll;
  AllNagryzkaPrepod:array of TNagryzkaPrepod;
  ArrMergeDis:array of array of TMergeDiscipline;

Procedure CreateAllGroup;
Procedure CopyNagryzkaPrepod(var El,NewEl:TNagryzkaPrepod);
Function  SeartchPrepodFIO(FIO:string):Longword;
procedure SortPrepodFIO;
Procedure SortNagryzkaPrepodTypeDis(NomPrepod:Longword);
Procedure SortNagryzkaPrepodNameDis(NomPrepod:Longword);
Procedure SortAllPrepodDateTime;
Procedure ProverkaTsel;
Procedure CreatePrep;
Procedure NomberOfDis(NomPrepod,NomNagr:Longword);
Procedure DeatroyAllDate;
Procedure GoAllKolStudentOn;
Procedure GoKolStudentOnCost;


Procedure VivodOshibkiBase(Me:TMemo);
Procedure VivodPrepodMemo (Me:TMemo);

//Excel
Procedure SearchAndAddExcelNagryzka(NomNagryzka,NomPrepod:LongWord; Hour:Double);
//Procedure ProverkaDisGroup;

implementation

Uses UMain, SysUtils;

Function SeartchPrepodFIO(FIO:string):Longword;
var
S,E,M:Longword;
begin
if Length(Prepod)>0 then
 begin
 S:=0;
 E:=high(Prepod);
 while E-S>1 do
   begin
     M:=(S+E) div 2;
     if FIO>=Prepod[M].FIO then S:=M else E:=M;
   end;
 if FIO=Prepod[E].FIO then
   result:=E
 else if FIO=Prepod[S].FIO then
   result:=S
 else
   Result:=65000;
 end
 else
   Result:=65000;
end;

procedure SortPrepodFIO;
var
  min,Size: integer;
  j: integer; { номер элемента, сравниваемого с минимальным }
  buf: TPrepodAll; { буфер, используемый при обмене элементов массива }
  i: integer;
begin
  Size:=Length(Prepod)-1;
  for i := 0 to Size - 1 do
  begin
    { поиск минимального элемента в части массива от а[1] до a[SIZE]}
    min := i;
    for j := i + 1 to Size do
      if Prepod[j].FIO < Prepod[min].FIO then
        min := j;
    { поменяем местами a [min] и a[i] }
    if i<>min then
      begin
      buf := Prepod[i];
      Prepod[i] := Prepod[min];
      Prepod[min] := buf;
      end;
   end;
end;

Procedure CopyNagryzkaPrepod(var El,NewEl:TNagryzkaPrepod);
var
i,n:longword;
begin
NewEl.Prepod:=El.Prepod;
NewEl.sem:=El.sem;
NewEl.NomRow:=El.NomRow;
NewEl.P:=El.P;
NewEl.Dis:=El.Dis;
NewEl.SOKR:=El.SOKR;
NewEl.Vid:=El.Vid;
NewEl.TypeVid:=El.TypeVid;
NewEl.Group:=El.Group;
NewEl.Hour:=El.Hour;
NewEl.KolStudent:=El.KolStudent;
NewEl.NOMPrep:=El.NOMPrep;
NewEl.Opisanie:=El.Opisanie;
NewEl.Auditoria:=El.Auditoria;

NewEl.FlagIndPlan:=El.FlagIndPlan;
NewEl.PO:=El.PO;
NewEl.FlagVivoda:=El.FlagVivoda;
n:=length(El.StDateTime);
SetLength(NewEl.StDateTime,n);
if n<>0 then
for i := 0 to n-1 do
  NewEl.StDateTime[i]:=El.StDateTime[i];
end;



Procedure VivodOshibkiBase(Me:TMemo);
var
NomPrepod,NomNagryzka,NomNagryzkaPrepod:Longword;
VivodOsh:boolean;
begin
//Обнуление проверки преподавателей
Me.Lines.Add('ПРОВЕРКА ПРЕПОДАВАТЕЛЕЙ');
VivodOsh:=false;
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  if Prepod[NomPrepod].P<>1 then
    begin
    Me.Lines.Add('ОШИБКА у ПРЕПОДАВАТЕЛЯ:'+Prepod[NomPrepod].FIO);
    VivodOsh:=true;
    end;
  NomNagryzkaPrepod:=0;
  while NomNagryzkaPrepod<length(Prepod[NomPrepod].Nagryzka) do
    begin
    if Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].P<>1 then
      begin
      Me.Lines.Add('ОШИБКА в ДИСЦИПЛИНЕ:'+Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].Dis+' '+Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].Vid+' '+Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].Group[0].Nom);
      VivodOsh:=true;
      end;
    inc(NomNagryzkaPrepod);
    end;
  inc(NomPrepod);
  end;
if not VivodOsh then
  Me.Lines.Add('OK');
//Обнуление проверки нагрузки
Me.Lines.Add('ПРОВЕРКА ОБЩЕЙ НАГРУЗКИ');
VivodOsh:=false;
NomNagryzka:=1;
while NomNagryzka<Length(Nagryzka) do
  begin
  if Nagryzka[NomNagryzka].P<>1 then
    begin
    Me.Lines.Add('ОШИБКА в ДИСЦИПЛИНЕ:'+Nagryzka[NomNagryzka].Dis+' '+Nagryzka[NomNagryzka].Vid+' '+Nagryzka[NomNagryzka].Group);
    VivodOsh:=true;
    end;
  inc(NomNagryzka);
  end;
if not VivodOsh then
  Me.Lines.Add('OK');
end;

Procedure VivodPrepodMemo (Me:TMemo);
var
NomPrepod:Longword;
begin
NomPrepod:=0;
while (NomPrepod<Length(Prepod))  do
  begin
  Me.Lines.Add(Prepod[NomPrepod].FIO);
  inc(NomPrepod);
  end;
end;

Procedure SortNagryzkaPrepodTypeDis(NomPrepod:Longword);
var
NomNagryzka,NomCurrentNagryzka,NomMinNagryazka:Longword;
El:TNagryzkaPrepod;
i:longword;
begin
El:=TNagryzkaPrepod.Create;
NomNagryzka:=0;
while NomNagryzka<Length(Prepod[NomPrepod].Nagryzka) do
  begin
  NomMinNagryazka:=NomNagryzka;
  NomCurrentNagryzka:=NomNagryzka+1;
  while NomCurrentNagryzka<Length(Prepod[NomPrepod].Nagryzka) do
    begin
    if (Prepod[NomPrepod].Nagryzka[NomCurrentNagryzka].sem<Prepod[NomPrepod].Nagryzka[NomMinNagryazka].sem) or
       //((Prepod[NomPrepod].Nagryzka[NomCurrentNagryzka].sem=Prepod[NomPrepod].Nagryzka[NomMinNagryazka].sem) and (Prepod[NomPrepod].Nagryzka[NomCurrentNagryzka].TypeVid<10) and (Prepod[NomPrepod].Nagryzka[NomMinNagryazka].TypeVid>=10)) or
       ((Prepod[NomPrepod].Nagryzka[NomCurrentNagryzka].sem=Prepod[NomPrepod].Nagryzka[NomMinNagryazka].sem) and (Prepod[NomPrepod].Nagryzka[NomCurrentNagryzka].TypeVid<Prepod[NomPrepod].Nagryzka[NomMinNagryazka].TypeVid)) then
      NomMinNagryazka:=NomCurrentNagryzka;
    inc(NomCurrentNagryzka);
    end;
  if NomMinNagryazka<>NomNagryzka then
    begin
    CopyNagryzkaPrepod(Prepod[NomPrepod].Nagryzka[NomMinNagryazka],El);
    CopyNagryzkaPrepod(Prepod[NomPrepod].Nagryzka[NomNagryzka],Prepod[NomPrepod].Nagryzka[NomMinNagryazka]);
    CopyNagryzkaPrepod(El,Prepod[NomPrepod].Nagryzka[NomNagryzka]);
    end;
  inc(NomNagryzka);
  end;
end;

Procedure SortNagryzkaPrepodNameDis(NomPrepod:Longword);
var
NomNagryzka,NomCurrentNagryzka,NomMinNagryazka:Longword;
El:TNagryzkaPrepod;
i:longword;
begin
El:=TNagryzkaPrepod.Create;
NomNagryzka:=0;
while NomNagryzka<Length(Prepod[NomPrepod].Nagryzka) do
  begin
  NomMinNagryazka:=NomNagryzka;
  NomCurrentNagryzka:=NomNagryzka+1;
  while NomCurrentNagryzka<Length(Prepod[NomPrepod].Nagryzka) do
    begin
    if (Prepod[NomPrepod].Nagryzka[NomCurrentNagryzka].sem<Prepod[NomPrepod].Nagryzka[NomMinNagryazka].sem) or
       ((Prepod[NomPrepod].Nagryzka[NomCurrentNagryzka].sem=Prepod[NomPrepod].Nagryzka[NomMinNagryazka].sem) and (Prepod[NomPrepod].Nagryzka[NomCurrentNagryzka].TypeVid<10) and (Prepod[NomPrepod].Nagryzka[NomMinNagryazka].TypeVid>=10)) or
       ((Prepod[NomPrepod].Nagryzka[NomCurrentNagryzka].sem=Prepod[NomPrepod].Nagryzka[NomMinNagryazka].sem) and (Prepod[NomPrepod].Nagryzka[NomCurrentNagryzka].Dis<Prepod[NomPrepod].Nagryzka[NomMinNagryazka].Dis)) then
      NomMinNagryazka:=NomCurrentNagryzka;
    inc(NomCurrentNagryzka);
    end;
  if NomMinNagryazka<>NomNagryzka then
    begin
    CopyNagryzkaPrepod(Prepod[NomPrepod].Nagryzka[NomMinNagryazka],El);
    CopyNagryzkaPrepod(Prepod[NomPrepod].Nagryzka[NomNagryzka],Prepod[NomPrepod].Nagryzka[NomMinNagryazka]);
    CopyNagryzkaPrepod(El,Prepod[NomPrepod].Nagryzka[NomNagryzka]);
    end;
  inc(NomNagryzka);
  end;
end;

Procedure SortAllPrepodDateTime;
var
  NomPrepod,NomNagryzka,NomDt,NomCurrDt,NomMinDt:Longword;
  BuffDate,BuffTime:string;
begin
NomPrepod:=0;
while NomPrepod<length(Prepod) do
  begin
  NomNagryzka:=0;
  while NomNagryzka<Length(Prepod[NomPrepod].Nagryzka) do
    begin
    if Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime)>1 then
      begin
      NomDt:=0;
      while NomDt<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime)-1 do
        begin
        NomMinDt:=NomDt;
        NomCurrDt:=NomDt+1;
        while NomCurrDt<Length(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime) do
          begin
          If (StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomCurrDt].StDate)<StrToDateTime(Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomMinDt].StDate)) or
             ((Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomCurrDt].StDate=Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomMinDt].StDate) and
             (Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomCurrDt].StTime<Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomMinDt].StTime)) then
             NomMinDt:=NomCurrDt;
          inc(NomCurrDt);
          end;
        if NomMinDt<>NomDt then
          begin
          BuffDate:=Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomMinDt].StDate;
          BuffTime:=Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomMinDt].StTime;
          Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomMinDt].StDate:=Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDt].StDate;
          Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomMinDt].StTime:=Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDt].StTime;
          Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDt].StDate:=BuffDate;
          Prepod[NomPrepod].Nagryzka[NomNagryzka].StDateTime[NomDt].StTime:=BuffTime;
          end;
        inc(NomDt);
        end;
      end;
    inc(NomNagryzka);
    end;
  inc(NomPrepod);
  end;
end;

Procedure CreateAllGroup;
var
NomPrepod,NomNagr,NomGroupInNagryzka,NomTypeGroup,NomGroup:Longword;
CurrentGroup:string;
begin
SetLength(NameAllGroup,0);
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  NomNagr:=0;
  while NomNagr<Length(Prepod[NomPrepod].Nagryzka) do
    begin
    NomGroupInNagryzka:=0;
    while NomGroupInNagryzka<length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
      begin
    CurrentGroup:=Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupInNagryzka].Nom;
    NomTypeGroup:=0;
    // copy(CurrentGroup,pos('-',CurrentGroup)+2,3)
    while (NomTypeGroup<Length(NameAllGroup)) and (NameAllGroup[NomTypeGroup].NameGroupKyrs<>copy(CurrentGroup,pos('-',CurrentGroup)+2,3)) do
      inc(NomTypeGroup);
    if not (NomTypeGroup<Length(NameAllGroup)) then
      begin
      SetLength(NameAllGroup,NomTypeGroup+1);
      NameAllGroup[NomTypeGroup].NameGroupKyrs:=copy(CurrentGroup,pos('-',CurrentGroup)+2,3);
      setLength(NameAllGroup[NomTypeGroup].Group,0);
      end;
    NomGroup:=0;
    while (NomGroup<Length(NameAllGroup[NomTypeGroup].Group)) and (NameAllGroup[NomTypeGroup].Group[NomGroup]<>CurrentGroup) do
      inc(NomGroup);
    if not (NomGroup<Length(NameAllGroup[NomTypeGroup].Group)) then
      begin
      SetLength(NameAllGroup[NomTypeGroup].Group,NomGroup+1);
      NameAllGroup[NomTypeGroup].Group[NomGroup]:=CurrentGroup;
      end;
      inc(NomGroupInNagryzka);
      end;
    inc(NomNagr);
    end;
  inc(NomPrepod);
  end;
end;

Procedure GoKolStudentOnCost;
var
NomPrepod,NomNagryzka,NomConst:Longword;
begin
NomPrepod:=0;
While NomPrepod<Length(Prepod) do
  begin
  NomNagryzka:=0;
  while NomNagryzka<Length(Prepod[NomPrepod].Nagryzka) do
    begin
    NomConst:=0;
    while (NomConst<Length(HourOnOneStudent)) and (HourOnOneStudent[NomConst].Vid<>Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid) do
      inc(NomConst);
    if NomConst<Length(HourOnOneStudent) then
      Prepod[NomPrepod].Nagryzka[NomNagryzka].KolStudent:=Trunc(StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagryzka].Hour)/HourOnOneStudent[NomConst].Hour)
    else
      Prepod[NomPrepod].Nagryzka[NomNagryzka].KolStudent:=0;
    inc(NomNagryzka);
    end;
  inc(NomPrepod);
  end;

NomNagryzka:=0;
while NomNagryzka<Length(Nagryzka) do
  begin
  NomConst:=0;
  while (NomConst<Length(HourOnOneStudent)) and (HourOnOneStudent[NomConst].Vid<>Nagryzka[NomNagryzka].Vid) do
    inc(NomConst);
  if NomConst<Length(HourOnOneStudent) then
    Nagryzka[NomNagryzka].KolStudent:=Trunc(StrToFloat(Nagryzka[NomNagryzka].Hour)/HourOnOneStudent[NomConst].Hour)
  else
    Nagryzka[NomNagryzka].KolStudent:=0;
  inc(NomNagryzka);
  end;
end;

Procedure GoAllKolStudentOn;
var
NomPrepod,NomNagryzka,NomNagryzkaGl,NomGroup:Longword;
SumStudent:Longword;
begin
NomPrepod:=0;
While NomPrepod<Length(Prepod) do
  begin
  NomNagryzka:=0;
  while NomNagryzka<Length(Prepod[NomPrepod].Nagryzka) do
    begin
    SumStudent:=0;
    NomGroup:=0;
    while NomGroup<length(Prepod[NomPrepod].Nagryzka[NomNagryzka].Group) do
      begin
      SumStudent:=SumStudent+Prepod[NomPrepod].Nagryzka[NomNagryzka].Group[NomGroup].KolStudent;
      inc(NomGroup);
      end;
    Prepod[NomPrepod].Nagryzka[NomNagryzka].KolStudent:=SumStudent;
    //Поиск нагрузки в массиве нагрузок
    NomNagryzkaGl:=0;
    while (NomNagryzkaGl<Length(Nagryzka)) and not
          ((Nagryzka[NomNagryzkaGl].Dis=Prepod[NomPrepod].Nagryzka[NomNagryzka].Dis) and
          (Nagryzka[NomNagryzkaGl].Vid=Prepod[NomPrepod].Nagryzka[NomNagryzka].Vid) and
          (Nagryzka[NomNagryzkaGl].Hour=Prepod[NomPrepod].Nagryzka[NomNagryzka].Hour) and
          (SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagryzka].Group,Nagryzka[NomNagryzkaGl].Group)<>65000)) do
      inc(NomNagryzkaGl);
    if NomNagryzkaGl<Length(Nagryzka) then
      Nagryzka[NomNagryzkaGl].KolStudent:=SumStudent;
    inc(NomNagryzka);
    end;
  inc(NomPrepod);
  end;
end;

Procedure CreatePrep;
var
  NomPrepod,NomNagr,NomNagryzkaS,NomGroupNagryzka,NomHourStudentDis:Longword;
  Sem:byte;
  HourNagr:Double;
  StGroup:string;
  HourSem:array [1..kolsem] of Double;
  st,st1,NameFileXlSX:string;
begin
NomPrepod:=0;
while (NomPrepod<Length(Prepod))  do
  begin
  NomNagryzkaS:=0;
  HourNagr:=0;
  for Sem := 1 to kolsem do
      begin
      NomNagr:=0;
      HourSem[Sem]:=0;
      while NomNagr<Length(Nagryzka) do
        begin
        if (Nagryzka[NomNagr].FIOPrep=Prepod[NomPrepod].FIO) and (Nagryzka[NomNagr].Sem=Sem) then
          begin
          SetLength(Prepod[NomPrepod].Nagryzka,NomNagryzkaS+1);
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS]:=TNagryzkaPrepod.Create;
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Prepod:=Prepod[NomPrepod];
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].P:=0;
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].sem:=Sem;
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].NomRow:=0;
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Dis:=Nagryzka[NomNagr].Dis;
          Prepod[NomPrepod].Nagryzka[NomNagryzkaS].Vid:=Nagryzka[NomNagr].Vid;

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
          end;
        inc(NomNagr);
        end;
      Prepod[NomPrepod].HourSem[Sem]:=HourSem[Sem];
      end;
  Prepod[NomPrepod].AllHour:=HourNagr;
  inc(NomPrepod);
  end;
end;

Procedure NomberOfDis(NomPrepod,NomNagr:Longword);
begin
if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛК')then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=1
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ПЗ') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=8
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛР') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=9
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КП') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=7
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КР') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=6
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Консультация') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=2
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Экзамен') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=3
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Зачет с оценкой') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=4
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Зачет') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=5
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Практика') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=10
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Руководство магистрами') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=11
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Руководство аспирантами') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=12
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Преддипломная практика') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=13
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Диплом') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=14
else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='Руководство кафедрой') then
Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=15
else
  Prepod[NomPrepod].Nagryzka[NomNagr].TypeVid:=250;
end;

Procedure ProverkaTsel;
var
NomPrepod,NomNagryzka,NomNagryzkaPrepod:Longword;
begin
//Обнуление проверки преподавателей
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  Prepod[NomPrepod].P:=0;
  NomNagryzkaPrepod:=0;
  while NomNagryzkaPrepod<length(Prepod[NomPrepod].Nagryzka) do
    begin
    Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].P:=0;
    inc(NomNagryzkaPrepod);
    end;
  inc(NomPrepod);
  end;
//Обнуление проверки нагрузки
NomNagryzka:=0;
while NomNagryzka<Length(Nagryzka) do
  begin
  Nagryzka[NomNagryzka].P:=0;
  inc(NomNagryzka);
  end;
//Проверка дисциплин у преподавателей из таблицы общей нагрузки
NomNagryzka:=0;
while NomNagryzka<Length(Nagryzka) do
  begin
  NomPrepod:=0;
  while (NomPrepod<Length(Prepod)) and (Prepod[NomPrepod].FIO<>Nagryzka[NomNagryzka].FIOPrep) do
    inc(NomPrepod);
  if Prepod[NomPrepod].FIO=Nagryzka[NomNagryzka].FIOPrep then
    begin
    NomNagryzkaPrepod:=0;
    while (NomNagryzkaPrepod<length(Prepod[NomPrepod].Nagryzka)) and not(
          (Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].sem=Nagryzka[NomNagryzka].Sem) and
          (Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].Dis=Nagryzka[NomNagryzka].Dis) and
          (Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].Vid=Nagryzka[NomNagryzka].Vid) and
          (SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].Group,Nagryzka[NomNagryzka].Group)<>65000) and
          (Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].Hour=Nagryzka[NomNagryzka].Hour) and
          (Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].Opisanie=Nagryzka[NomNagryzka].Opisanie) and
          (Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].NOMPrep=Nagryzka[NomNagryzka].NOMPrep)
          ) do
      begin
      inc(NomNagryzkaPrepod);
      end;
    if (NomNagryzkaPrepod<length(Prepod[NomPrepod].Nagryzka)) and
    (Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].sem=Nagryzka[NomNagryzka].Sem) and
          (Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].Dis=Nagryzka[NomNagryzka].Dis) and
          (Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].Vid=Nagryzka[NomNagryzka].Vid) and
          (SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].Group,Nagryzka[NomNagryzka].Group)<>65000) and
          (Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].Hour=Nagryzka[NomNagryzka].Hour) and
          (Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].Opisanie=Nagryzka[NomNagryzka].Opisanie) and
          (Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].NOMPrep=Nagryzka[NomNagryzka].NOMPrep) then
      begin
      inc(Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].P);
      inc(Nagryzka[NomNagryzka].P);
      end;
    end;
  inc(NomNagryzka);
  end;
//Проверка все ли учтено у преподавателей
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  NomNagryzkaPrepod:=0;
  Prepod[NomPrepod].P:=1;
  while NomNagryzkaPrepod<length(Prepod[NomPrepod].Nagryzka) do
    begin
    if Prepod[NomPrepod].Nagryzka[NomNagryzkaPrepod].P<>1 then
      Prepod[NomPrepod].P:=0;
    inc(NomNagryzkaPrepod);
    end;
  inc(NomPrepod);
  end;
end;

//Поиск в файле нагрузки в базе и если нужно добавление новой строки с новым преподавателем
Procedure SearchAndAddExcelNagryzka(NomNagryzka,NomPrepod:LongWord; Hour:Double);
var
  NomNagryzkaSdvig:LongWord;
  St:String;
begin
  Excel.Workbooks.Open(NameFileNagryzka[Nagryzka[NomNagryzka].Sem]);
  //Включить преподавателя в проверку на изменение
  Prepod[NomPrepod].FlagP:=1;
  //Проверка наличия нагрузки у данного преподавателя.
  if Nagryzka[NomNagryzka].FIOPrep=Prepod[NomPrepod].FIO then
    begin
    //Если нагрузка есть, то в нужном месте в фале нагрузки нужно изменить часы
    Excel.Cells[Nagryzka[NomNagryzka].NomRow,4]:=Hour;
    //Изменить количество часов в базе
    Nagryzka[NomNagryzka].Hour:=FloatToStr(Hour);
    end
  else
    begin
    //Если нагрузки нет
    //Добавить новую строку
    Excel.ActiveSheet.Rows[Nagryzka[NomNagryzka].NomRow+1].Select;
    Excel.Selection.Insert(Shift :=xlDown);
    //В новую строку переписать параметры со старой
    st:=Excel.Cells[Nagryzka[NomNagryzka].NomRow,1];
    Excel.Cells[Nagryzka[NomNagryzka].NomRow+1,1]:=st;
    st:=Excel.Cells[Nagryzka[NomNagryzka].NomRow,2];
    Excel.Cells[Nagryzka[NomNagryzka].NomRow+1,2]:=st;
    st:=Excel.Cells[Nagryzka[NomNagryzka].NomRow,3];
    Excel.Cells[Nagryzka[NomNagryzka].NomRow+1,3]:=st;
    Excel.Cells[Nagryzka[NomNagryzka].NomRow+1,4]:=Hour;
    Excel.Cells[Nagryzka[NomNagryzka].NomRow+1,5]:=Prepod[NomPrepod].FIO;
    st:=Excel.Cells[Nagryzka[NomNagryzka].NomRow,6];
    Excel.Cells[Nagryzka[NomNagryzka].NomRow+1,6]:=st;
    st:=Excel.Cells[Nagryzka[NomNagryzka].NomRow,7];
    Excel.Cells[Nagryzka[NomNagryzka].NomRow+1,7]:=st;
    //Сдвинуть все строки в поле NomRow для всей остальной нагрузки
    For NomNagryzkaSdvig:=0 to Length(Nagryzka)-1 do
      if (Nagryzka[NomNagryzka].Sem=Nagryzka[NomNagryzkaSdvig].Sem) and (Nagryzka[NomNagryzka].NomRow<Nagryzka[NomNagryzkaSdvig].NomRow) then
        inc(Nagryzka[NomNagryzkaSdvig].NomRow);
    //Добавить новую запись в массив нагрузки
    NomNagryzkaSdvig:=Length(Nagryzka);
    SetLength(Nagryzka,NomNagryzkaSdvig+1);
    Nagryzka[NomNagryzkaSdvig].P:=0;
    Nagryzka[NomNagryzkaSdvig].NomRow:=Nagryzka[NomNagryzka].NomRow+1;
    Nagryzka[NomNagryzkaSdvig].Sem:=Nagryzka[NomNagryzka].Sem;
    Nagryzka[NomNagryzkaSdvig].Dis:=Nagryzka[NomNagryzka].Dis;
    Nagryzka[NomNagryzkaSdvig].Vid:=Nagryzka[NomNagryzka].Vid;
    Nagryzka[NomNagryzkaSdvig].Group:=Nagryzka[NomNagryzka].Group;
    Nagryzka[NomNagryzkaSdvig].Hour:=FloatToStr(Hour);
    Nagryzka[NomNagryzkaSdvig].FIOPrep:=Prepod[NomPrepod].FIO;
    Nagryzka[NomNagryzkaSdvig].Opisanie:=Nagryzka[NomNagryzka].Opisanie;

    end;

  FMain.MeProtocol.Lines.Add('Изменено ЧАСЫ нагрузки '+Nagryzka[NomNagryzka].Dis+' '+Nagryzka[NomNagryzka].Vid+' '+Nagryzka[NomNagryzka].Group+' '+Nagryzka[NomNagryzka].Hour+' '+Prepod[NomPrepod].FIO+' '+Nagryzka[NomNagryzka].Opisanie);
  FMain.MeProtocol.Lines.Add('На '+FloatToStr(Hour));
  Excel.Workbooks[1].Save;
  Excel.Workbooks.Close;
end;

{Procedure ProverkaDisGroup;
var
NomRow,NomPrepod,NomSearchPrepod,NomGroupDis,NomDisInGroup,NomHourStudentDis:Longword;
StExcel,st:string;
sem,NomArrSt,KolStudentGroupNagryzka:byte;
ArrSt:array[1..kolrownagryzka] of string;
begin
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  if Prepod[NomPrepod].NameFilePrepod<>'' then
    begin
    ExcelBase.Workbooks.Open(Prepod[NomPrepod].NameFilePrepod);
    NomRow:=2;
    StExcel:=ExcelBase.Cells[NomRow,1];
    for Sem := 1 to kolsem do
      begin
      //Обнуление дисциплин
      NomGroupDis:=0;
      while NomGroupDis<Length(GroupDis) do
        begin
        NomDisInGroup:=0;
        while (NomDisInGroup<Length(GroupDis[NomGroupDis])) do
          begin
          GroupDis[NomGroupDis][NomDisInGroup].Enabled:=0;
          inc(NomDisInGroup);
          end;
        inc(NomGroupDis);
        end;
      //Прохождение по всему семестру в файле
      while StExcel<>'' do
        begin
        //Получение нагрузки из файла преподавателя
        st:='';
        for NomArrSt := 1 to kolrownagryzka do
          begin
          ArrSt[NomArrSt]:=ExcelBase.Cells[NomRow,NomArrSt];
          st:=st+ArrSt[NomArrSt]+' ';
          end;

        //Проверка вхождения в дублирующие
        NomGroupDis:=0;
        while NomGroupDis<Length(GroupDis) do
          begin
          NomDisInGroup:=0;
          while (NomDisInGroup<Length(GroupDis[NomGroupDis])) and not((GroupDis[NomGroupDis][NomDisInGroup].Dis=ArrSt[1]) and (GroupDis[NomGroupDis][NomDisInGroup].Vid=ArrSt[2]) and (GroupDis[NomGroupDis][NomDisInGroup].Group=ArrSt[3])) do
            inc(NomDisInGroup);
          if NomDisInGroup<Length(GroupDis[NomGroupDis]) then
            begin
            GroupDis[NomGroupDis][NomDisInGroup].Enabled:=1; //Установление флага нахождения нагрузки
            end;
          inc(NomGroupDis);
          end;
        inc(NomRow);
        StExcel:=ExcelBase.Cells[NomRow,1];
        end;
        //В конце семестра проверить наличие всех дисциплин из группы
        NomGroupDis:=0;
        while NomGroupDis<Length(GroupDis) do
          begin
          NomDisInGroup:=0;
          while (NomDisInGroup<Length(GroupDis[NomGroupDis])) and (GroupDis[NomGroupDis][NomDisInGroup].Enabled=0) do
            inc(NomDisInGroup);
          if NomDisInGroup<Length(GroupDis[NomGroupDis]) then
              begin
              NomHourStudentDis:=0;
              if length(HourStudentDis)<>0 then
                while (NomHourStudentDis<Length(HourStudentDis)) and (HourStudentDis[NomHourStudentDis].Dis=Nagryzka[GroupDis[NomGroupDis][NomDisInGroup].NomRowNagryzka].Dis) and (HourStudentDis[NomHourStudentDis].Vid=Nagryzka[GroupDis[NomGroupDis][NomDisInGroup].NomRowNagryzka].Vid) and (HourStudentDis[NomHourStudentDis].Group=Nagryzka[GroupDis[NomGroupDis][NomDisInGroup].NomRowNagryzka].Group) do
                  inc(NomHourStudentDis);
              if NomHourStudentDis<Length(HourStudentDis) then
                KolStudentGroupNagryzka:=Trunc(StrToFloat(Nagryzka[GroupDis[NomGroupDis][NomDisInGroup].NomRowNagryzka].Hour)/HourStudentDis[NomHourStudentDis].HourForOneStudent);
              //Проверить наличие остальных дисциплин
              NomDisInGroup:=0;
              while (NomDisInGroup<Length(GroupDis[NomGroupDis])) do
                begin
                if GroupDis[NomGroupDis][NomDisInGroup].Enabled=0 then
                  begin
                  NomSearchPrepod:=SeartchPrepodFIO('не назначено');
                  NomHourStudentDis:=0;
                  if length(HourStudentDis)<>0 then
                    while (NomHourStudentDis<Length(HourStudentDis)) and (HourStudentDis[NomHourStudentDis].Dis=Nagryzka[GroupDis[NomGroupDis][NomDisInGroup].NomRowNagryzka].Dis) and (HourStudentDis[NomHourStudentDis].Vid=Nagryzka[GroupDis[NomGroupDis][NomDisInGroup].NomRowNagryzka].Vid) and (HourStudentDis[NomHourStudentDis].Group=Nagryzka[GroupDis[NomGroupDis][NomDisInGroup].NomRowNagryzka].Group) do
                      inc(NomHourStudentDis);
                  if NomHourStudentDis<Length(HourStudentDis) then
                  begin
                  //Уменьшить часы у нагрузки в файле
                  SearchAndAddExcelNagryzka(GroupDis[NomGroupDis][NomDisInGroup].NomRowNagryzka,NomSearchPrepod,StrTOFloat(Nagryzka[GroupDis[NomGroupDis][NomDisInGroup].NomRowNagryzka].Hour)-KolStudentGroupNagryzka*HourStudentDis[NomHourStudentDis].HourForOneStudent);
                  //Добавить нагрузку преподавателю
                  SearchAndAddExcelNagryzka(GroupDis[NomGroupDis][NomDisInGroup].NomRowNagryzka,NomPrepod,KolStudentGroupNagryzka*HourStudentDis[NomHourStudentDis].HourForOneStudent);
                  end;
                  end;
                inc(NomDisInGroup);
                end;
              end;
          inc(NomGroupDis);
          end;
      end;
    ExcelBase.Workbooks.Close;
    end;
  inc(NomPrepod);
  end;
end;
 }


Constructor TNagryzkaPrepod.Create;
var
NomElement:Longword;
  begin
  SetLength(StDateTime,0);
  NomElement:=Length(AllNagryzkaPrepod);
  SetLength(AllNagryzkaPrepod,NomElement+1);
  AllNagryzkaPrepod[NomElement]:=self;
  Auditoria:=ArrAuditorii[0];
  end;
Destructor TNagryzkaPrepod.Destroy;
  begin
  SetLength(StDateTime,0);
  end;

Constructor TPrepodAll.Create;
  begin
  SetLength(Nagryzka,0);
  end;
Destructor TPrepodAll.Destroy;
var
i:longword;
  begin
  i:=0;
  while i<Length(Nagryzka) do
    begin
    Nagryzka[i].Destroy;
    inc(i);
    end;
  SetLength(Nagryzka,0);
  end;

Procedure DeatroyAllDate;
var
NomPrepod:Longword;
begin
//Обнуление преподавателей
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  SetLength(Prepod[NomPrepod].Nagryzka,0);
  inc(NomPrepod);
  end;
SetLength(Prepod,0);
//Обнуление  нагрузки
SetLength(Nagryzka,0);
end;

end.
