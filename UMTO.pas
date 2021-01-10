unit UMTO;

interface

uses SysUtils, UMain, UNagryzka, UGroup, UConstParametrs, USemPlan, UAuditoria;

procedure LoadMTOIASYFile(FileName:string);
procedure LoadPo(FileName:string);
Procedure SaveAllMTOSemPlan(TypeVivod:Byte; NameFolder:string);
Procedure SaveAllMTO;

implementation

procedure LoadMTOIASYFile(FileName:string);
var
NomRow:longword;
NomSemYp,NomDisSemYp:longword;
st,StDis,st1:string;
begin
if FileExists(FileName) then
begin
FMain.PRPD.Color:=$0024AA35;
Excel.Workbooks.Open(FileName);
NomRow:=10;
st:=Excel.Cells[NomRow,1];
while st<>'END' do
  begin
  st:=Excel.Cells[NomRow,9];
  if (St<>'') and (pos('РП дисциплины',st)<>0) then
    begin
    StDis:=Excel.Cells[NomRow-1,1];
    NomSemYp:=0;
    while NomSemYp<Length(SemYp) do
      begin
      NomDisSemYp:=0;
      while NomDisSemYp<Length(SemYp[NomSemYp].Disciplin) do
        begin
        if SemYp[NomSemYp].Disciplin[NomDisSemYp].Name=StDis then
          begin
          st1:=Excel.Cells[NomRow,10];
          if pos(st1,SemYp[NomSemYp].Disciplin[NomDisSemYp].OsnashenieRPD)=0 then
            SemYp[NomSemYp].Disciplin[NomDisSemYp].OsnashenieRPD:=SemYp[NomSemYp].Disciplin[NomDisSemYp].OsnashenieRPD+' '+st1;
          st1:=Excel.Cells[NomRow,11];
          if pos(st1,SemYp[NomSemYp].Disciplin[NomDisSemYp].PoRPD)=0 then
            SemYp[NomSemYp].Disciplin[NomDisSemYp].PoRPD:=SemYp[NomSemYp].Disciplin[NomDisSemYp].PoRPD+' '+st1;
          end;
        inc(NomDisSemYp);
        end;
      inc(NomSemYp);
      end;
    end;
  inc(NomRow);
  st:=Excel.Cells[NomRow,1];
  end;

Fmain.MeProtocol.Lines.Add('Загружено МТО из ИАСУ из файла:'+FileName);
Excel.Workbooks.Close;
end;
end;

procedure LoadPo(FileName:string);
var
NomPrepod,NomNagr:longword;
NomSemYp,NomDisSemYp:longword;
NomRow:longword;
st1:string;
begin
if FileExists(FileName) then
begin
FMain.PPO.Color:=$0024AA35;
Excel.Workbooks.Open(FileName);
NomRow:=2;
St1:=Excel.Cells[NomRow,2];
while st1<>'' do
  begin
  NomPrepod:=0;
    while NomPrepod<length(Prepod) do
      begin
      NomNagr:=0;
      while NomNagr<length(Prepod[NomPrepod].Nagryzka) do
        begin
        if (Prepod[NomPrepod].Nagryzka[NomNagr].Dis=st1) and
        ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛР') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ПЗ')) then
          Prepod[NomPrepod].Nagryzka[NomNagr].PO:=Excel.Cells[NomRow,7];
        inc(NomNagr);
        end;
      inc(NomPrepod);
      end;
  NomSemYp:=0;
  while NomSemYp<Length(SemYp) do
    begin
    NomDisSemYp:=0;
    while NomDisSemYp<Length(SemYp[NomSemYp].Disciplin) do
      begin
      if SemYp[NomSemYp].Disciplin[NomDisSemYp].Name=st1 then
        SemYp[NomSemYp].Disciplin[NomDisSemYp].PO:=Excel.Cells[NomRow,7];
      inc(NomDisSemYp);
      end;
    inc(NomSemYp);
    end;
  inc(NomRow);
  St1:=Excel.Cells[NomRow,2];
  end;
Fmain.MeProtocol.Lines.Add('Загружено программное обеспечение из файла '+FileName);
Excel.Workbooks.Close;
end;
end;

Procedure SaveAllMTOSemPlan(TypeVivod:Byte; NameFolder:string);
Const
  KolPosInDis = 5;
  KolRows = 6;
var
NomProfil,NomSemPlan,NomDisSemYP:Longword;
NomSemPlanToo,NomDisSemYPToo:Longword;
Nom,NomRow,NomDecRow,NomStartTabled,NomKrArr:Longword;
i:byte;
f:TextFile;
ArrType,ArrTypeAud:array [1..KolPosInDis] of boolean;
st:string;
NameDisLR:string;
TrebAudProektor,AudProektorVivod,NameDis:string;

Procedure VivodAuditoria (Aud:TAuditoria; typezan:byte; NomRow:Longword);
var
st:string;
begin
if Aud<>nil then
begin
  ArrTypeAud[typezan]:=true;
  case typezan of
    1:st:='Учебная аудитория для проведения занятий лекционного типа:';
    2:if Aud.NameAuditoria<>'' then st:='Лаборатория: "'+Aud.NameAuditoria+'"' else st:='Лаборатория.'{ для проведения занятий по курсу "'+NameDisLr+'"'};
    3:st:='Учебная аудитория для проведения занятий семинарского типа:';
    4:st:='Учебная аудитория для курсового проектирования (выполнения курсовых работ):';
    5:st:='Учебная аудитория для групповых и индивидуальных консультаций, текущего контроля и промежуточной аттестации:';
    6:st:='Помещение для самостоятельной работы обучающихся:';
    7:st:='Учебная аудитория для текущего контроля и промежуточной аттестации:';
    8:st:='Помещение для хранения и профилактического обслуживания учебного оборудования:';
  end;
  if Pos('Орш',Aud.Auditoria)<>0 then
    st:=st+chr(10)+'г.Москва, Ул Оршанская, 3, 121552'
  else
    st:=st+chr(10)+'125993, г.Москва, Волоколамское шоссе, д.4';
  if ((TypeVivod=1) or(TypeVivod=3)) and (Copy(Aud.Auditoria,5,length(Aud.Auditoria)-4)<>'') then
    st:=st+', корпус - '+Aud.Korpus{+' №'+Copy(Copy(Aud.Auditoria,5,length(Aud.Auditoria)-4),1,Pos('(',Copy(Aud.Auditoria,5,length(Aud.Auditoria)-4))-1)};
  Excel.Cells[NomRow,3]:=st;
  case typezan of
    1:st:='Учебная мебель: столы и стулья для обучающихся; стол и стул для преподавателя; доска с мелом (маркером)';
    2:if Aud.ProektorAuditoria<>'' then st:='Специализированное лабораторное оборудование кафедры.'{ else st:='Специализированное лабораторное оборудование кафедры для проведения лабораторных работ по курсу "'+NameDisLr+'"'};
    3:st:='Учебная мебель: столы и стулья для обучающихся; стол и стул для преподавателя; доска с мелом (маркером)';
    4:st:='Учебная мебель: столы и стулья для обучающихся; Компьютеры с доступом в сеть Internet';
    5:If Pos('актика',NameDis)<>0 then st:='Учебная мебель: столы и стулья для обучающихся; Специализированная мебель и технические средства обучения, служащие для представления учебной информации большой аудитории'+' (презентационная техника: проектор, экран, компьютер/ноутбук). Компьютеры для работы студентов.' else st:='Учебная мебель: столы и стулья для обучающихся;';
    6,7:st:='Учебная мебель: столы и стулья для обучающихся; Компьютеры с доступом в сеть Internet и наличием доступа в электронную информационно-образовательную среду.';
    8:st:='Набор для сборки вычислительных устройств';
  end;

  if (typezan<=3) and (Aud.ProektorAuditoria<>'') then
    st:=st+', '+Aud.ProektorAuditoria
  else
  if (typezan<=3) and ((pos('роектор',TrebAudProektor)<>0) or
  (pos('оутбук',TrebAudProektor)<>0) or (pos('исплей',TrebAudProektor)<>0) or (pos('омпьютер',TrebAudProektor)<>0)  or
     (pos('презентац',TrebAudProektor)<>0) or (pos('слайд',TrebAudProektor)<>0) ) then
    st:=st+', '+'Специализированная мебель и технические средства обучения, служащие для представления учебной информации большой аудитории (презентационная техника: проектор, экран, компьютер/ноутбук).';
  AudProektorVivod:=st;
  Excel.Cells[NomRow,4]:=st;
//  if KolRows>5 then
  if Aud.OsnashenieOgrnAuditoria<>'' then
    Excel.Cells[NomRow,6]:=Aud.OsnashenieOgrnAuditoria
  else
    Excel.Cells[NomRow,6]:='Не приспособлено.';
end;
end;

begin
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

    Excel.WorkBooks.Add;

    Excel.Columns[1].ColumnWidth := 3.00;
    Excel.Columns[2].ColumnWidth := 13.29;
    Excel.Columns[3].ColumnWidth := 25.86;
    Excel.Columns[4].ColumnWidth := 47.57;
    Excel.Columns[5].ColumnWidth := 44.71;
    Excel.Columns[6].ColumnWidth := 13.57;

    Excel.Range[Excel.Cells[1,1],Excel.Cells[1,KolRows]].MergeCells:=true;
    Excel.Cells[1,1]:='федеральное государственное бюджетное образовательное учреждение высшего образования';
    Excel.Range[Excel.Cells[2,1],Excel.Cells[2,KolRows]].MergeCells:=true;
    Excel.Cells[2,1]:='«Московский авиационный институт (национальный исследовательский университет)»';
    Excel.Range[Excel.Cells[4,1],Excel.Cells[4,KolRows]].MergeCells:=true;
    Excel.Cells[4,1]:='Справка';
    Excel.Range[Excel.Cells[5,1],Excel.Cells[5,KolRows]].MergeCells:=true;
    st:='о материально-техническом обеспечении основной образовательной программы высшего образования – программы';
    If Pos('05',ArrProfil[NomProfil].Profil)<>0 then Excel.Cells[5,1]:=st+' аспирантуры';
    If Pos('Б',ArrProfil[NomProfil].Profil)<>0 then Excel.Cells[5,1]:=st+' бакалавриата';
    If Pos('М',ArrProfil[NomProfil].Profil)<>0 then Excel.Cells[5,1]:=st+' магистратуры';
    If Pos('С',ArrProfil[NomProfil].Profil)<>0 then Excel.Cells[5,1]:=st+' специалитета';
    Excel.Range[Excel.Cells[6,1],Excel.Cells[6,KolRows]].MergeCells:=true;
    Excel.Cells[6,1]:=ArrProfil[NomProfil].Naprav+' '+ArrProfil[NomProfil].NameNaprav+' - '+ArrProfil[NomProfil].NameProfil;
    Excel.Cells[8,1]:='№ п\п';
    Excel.Cells[8,2]:='Наименование дисциплины (модуля), практик в соответствии с УП';
    Excel.Cells[8,3]:='Наименование специальных* помещений и помещений для самостоятельной работы';
    Excel.Cells[8,4]:='Оснащенность специальных помещений и помещений для самостоятельной работы';
    Excel.Cells[8,5]:='Перечень лицензионного программного обеспечения. Реквизиты подтверждающего документа ';
    Excel.Cells[8,6]:='Приспособленность помещений для использования инвалидами и лицами с ограниченными возможностями здоровья';
    for i := 1 to 6 do
      Excel.Cells[9,i]:=i;
    Excel.Range[Excel.Cells[1,1],Excel.Cells[9,KolRows]].HorizontalAlignment:=xlCenter;

    NomRow:=10;
    Nom:=1;

  NomSemPlan:=0;
  while NomSemPlan<Length(ArrProfil[NomProfil].SemYp) do
    begin
    NomDisSemYP:=0;
    while NomDisSemYP<Length(ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin) do
      begin
      if not ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].BYCh then
        begin
        for i := 1 to KolPosInDis do
          begin
          ArrType[i]:=false;
          ArrTypeAud[i]:=false;
          end;
        NomSemPlanToo:=0;
        while NomSemPlanToo<Length(ArrProfil[NomProfil].SemYp) do
          begin
          NomDisSemYPToo:=0;
          while NomDisSemYPToo<Length(ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin) do
            begin
            if ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].Name=ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].Name then
              begin
              ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].BYCh:=true;
              NameDis:=ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].Name;
              TrebAudProektor:= ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD;
              if (ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LK<>0) and not (((TypeVivod=3) or (TypeVivod=4)) and (ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LKAud=nil)) then
                begin
                if ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LKAud<>nil then
                  VivodAuditoria(ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LKAud,1,NomRow)
                else
                if (not (ArrTypeAud[1])) and (ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].NomElektivDis<>65000) and
                   (ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].NomElektivDis].LKAud<>nil)then
                  begin
                  ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].NomElektivDis].LKAud:=ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LKAud;
                  VivodAuditoria(ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].NomElektivDis].LKAud,1,NomRow);
                  Excel.Cells[NomRow,9]:='э';
                  end;
                if ((ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LKAud<>nil) and
                   ((Pos('оутбук',AudProektorVivod)<>0) or
                   (Pos('омпьютер',AudProektorVivod)<>0))) or
                   (Pos('оутбук',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('исплей',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('ультимедийн',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('омпьютер',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0)  then
                begin
                  st:='Microsoft Windows Контракт №007-1-0834-18 от 31.05.2018г. Microsoft Office (Включая WORD, EXCEL и т.д.) Контракт №007-1-0834-18 от 31.05.2018г.';
                  Excel.Cells[NomRow,5]:=st+' Kaspersky Endpoint Security Контракт №070-1-0478-17 от 18.04.2017г. Kaspersky Endpoint Security Контракт №2075/M21 от 22.01.2015г.'{+ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PO};
                  end
                else
                begin
                st:=Excel.Cells[NomRow,5];
                if st='' then
                  Excel.Cells[NomRow,5]:='Не требуется.';
                end;
              Excel.Cells[NomRow,7]:= ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD;
              Excel.Cells[NomRow,8]:= ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PoRPD;
              ArrType[1]:=true;
              end;
              if (ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LR<>0) and not (((TypeVivod=3) or (TypeVivod=4)) and (ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LRAud=nil)) then
                begin
                NameDisLR:=ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].Name;
                if ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LRAud<>nil then
                  VivodAuditoria(ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LRAud,2,NomRow+1)
                else
                if (not (ArrTypeAud[2])) and (ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].NomElektivDis<>65000) and
                   (ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].NomElektivDis].LRAud<>nil)then
                   begin
                   ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].NomElektivDis].LRAud:=ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LRAud;
                   VivodAuditoria(ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].NomElektivDis].LRAud,2,NomRow+1);
                   end;
                If (ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].PO<>'') or
                   ((ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LRAud<>nil) and
                   ((Pos('оутбук',AudProektorVivod)<>0) or
                   (Pos('омпьютер',AudProektorVivod)<>0))) or
                   (Pos('оутбук',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('исплей',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('ультимедийн',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('омпьютер',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0)then
                  begin
                  st:='Microsoft Windows Контракт №007-1-0834-18 от 31.05.2018г. Microsoft Office (Включая WORD, EXCEL и т.д.) Контракт №007-1-0834-18 от 31.05.2018г.';
                  Excel.Cells[NomRow+1,5]:=st+' Kaspersky Endpoint Security Контракт №070-1-0478-17 от 18.04.2017г. Kaspersky Endpoint Security Контракт №2075/M21 от 22.01.2015г.'+ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PO;
                  end
                else
                begin
                st:=Excel.Cells[NomRow+1,5];
                if st='' then
                  Excel.Cells[NomRow+1,5]:='Не требуется.';
                end;
                  Excel.Cells[NomRow+1,7]:= ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD;
                  ArrType[2]:=true;
                end;
              if (ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PZ<>0) and not (((TypeVivod=3) or (TypeVivod=4)) and (ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PZAud=nil)) then
                begin
                if ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PZAud<>nil then
                  VivodAuditoria(ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PZAud,3,NomRow+2)
                else
                if (not (ArrTypeAud[3])) and (ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].NomElektivDis<>65000) and
                   (ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].NomElektivDis].PZAud<>nil)then
                  begin
                  ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].NomElektivDis].PZAud:=ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PZAud;
                  VivodAuditoria(ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].NomElektivDis].PZAud,3,NomRow+2);
                  end;
                If (ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].PO<>'') or
                   ((ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PZAud<>nil) and
                   ((Pos('оутбук',AudProektorVivod)<>0) or
                   (Pos('омпьютер',AudProektorVivod)<>0))) or
                   (Pos('оутбук',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('исплей',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('ультимедийн',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('омпьютер',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) then
                  begin
                  st:='Microsoft Windows Контракт №007-1-0834-18 от 31.05.2018г. Microsoft Office (Включая WORD, EXCEL и т.д.) Контракт №007-1-0834-18 от 31.05.2018г.';
                  Excel.Cells[NomRow+2,5]:=st+' Kaspersky Endpoint Security Контракт №070-1-0478-17 от 18.04.2017г. Kaspersky Endpoint Security Контракт №2075/M21 от 22.01.2015г.'+ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PO;
                  end
                else
                  begin
                st:=Excel.Cells[NomRow+2,5];
                if st='' then
                  Excel.Cells[NomRow+2,5]:='Не требуется.';
                  end;
                  Excel.Cells[NomRow+2,7]:= ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD;
                ArrType[3]:=true;
                end;
              if  (ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].KR<>0) and (Length(ArrAuditoriiKP)<>0) then
                begin
                NomKrArr:=random(Length(ArrAuditoriiKP));
                VivodAuditoria(ArrAuditorii[ArrAuditoriiKP[NomKrArr]],4,NomRow+3);
                If true then
                  begin
                  st:='Microsoft Windows Контракт №007-1-0834-18 от 31.05.2018г. Microsoft Office (Включая WORD, EXCEL и т.д.) Контракт №007-1-0834-18 от 31.05.2018г.';
                  Excel.Cells[NomRow+3,5]:=st+' Kaspersky Endpoint Security Контракт №070-1-0478-17 от 18.04.2017г. Kaspersky Endpoint Security Контракт №2075/M21 от 22.01.2015г. Google Chrome (беспланое ПО)'+ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PO;
                  end
                else
                begin
                st:=Excel.Cells[NomRow+3,5];
                if st='' then
                  Excel.Cells[NomRow+3,5]:='Не требуется.';
                end;
                ArrType[4]:=true;
                end;
              if true then
                begin
                if ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LKAud<>nil then
                  VivodAuditoria(ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LKAud,5,NomRow+4)
                else
                if ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LRAud<>nil then
                  VivodAuditoria(ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LRAud,5,NomRow+4)
                else
                if ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PZAud<>nil then
                  VivodAuditoria(ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PZAud,5,NomRow+4);
                If pos('рактика',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].Name)<>0 then
                  begin
                  st:='Microsoft Windows Контракт №007-1-0834-18 от 31.05.2018г. Microsoft Office (Включая WORD, EXCEL и т.д.) Контракт №007-1-0834-18 от 31.05.2018г.';
                  Excel.Cells[NomRow+4,5]:=st+' Kaspersky Endpoint Security Контракт №070-1-0478-17 от 18.04.2017г. Kaspersky Endpoint Security Контракт №2075/M21 от 22.01.2015г. GoogleChrome. '+ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PO;
                  end
                else
                begin
                st:=Excel.Cells[NomRow+4,5];
                if st='' then
                  Excel.Cells[NomRow+4,5]:='Не требуется.';
                end;
                ArrType[5]:=true;
                end;
              end;
            inc(NomDisSemYPToo);
            end;
          inc(NomSemPlanToo);
          end;
        for i := 1 to KolPosInDis do
          if (ArrType[i]) and (not (ArrTypeAud[i])) then
            begin
            VivodAuditoria(ArrAuditorii[0],i,NomRow+i-1);
            end;

        NomDecRow:=0;
        for i := 1 to KolPosInDis do
          if not(ArrType[i]) then
            begin
            Excel.ActiveSheet.Rows[NomRow+i-1-NomDecRow].Select;
            Excel.Selection.Delete(Shift :=-4162);
            inc(NomDecRow);
            end;

        Excel.Cells[NomRow,1]:=Nom;
        Excel.Cells[NomRow,2]:=ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDisSemYP].Name;
        If NomDecRow<KolPosInDis-1 then
          begin
          Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow+KolPosInDis-1-NomDecRow,1]].MergeCells:=true;
          Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow+KolPosInDis-1-NomDecRow,2]].MergeCells:=true;
          end;
        NomRow:=NomRow+KolPosInDis-NomDecRow;
        inc(Nom);
        end;
      inc(NomDisSemYP);
      end;
    inc(NomSemPlan);
    end;

 //ArrAuditSRS,ArrAuditoriiKP,ArrAuditoriiKons,ArrAuditoriiKontrol,ArrAuditoriiObslyz
  Excel.Cells[NomRow,1]:=Nom;
  inc(Nom);
  Excel.Cells[NomRow,2]:='Самостоятельная работа обучающихся';
  i:=0;
  while i<Length(ArrAuditSRS) do
    begin
    VivodAuditoria(ArrAuditorii[ArrAuditSRS[i]],6,NomRow+i);
    st:='Microsoft Windows Контракт №007-1-0834-18 от 31.05.2018г. Microsoft Office (Включая WORD, EXCEL и т.д.) Контракт №007-1-0834-18 от 31.05.2018г.';
    Excel.Cells[NomRow+i,5]:=st+' Kaspersky Endpoint Security Контракт №070-1-0478-17 от 18.04.2017г. Kaspersky Endpoint Security Контракт №2075/M21 от 22.01.2015г. Google Chrome';
    inc(i);
    end;
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow+i-1,1]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow+i-1,2]].MergeCells:=true;
  NomRow:=NomRow+i;
{  Excel.Cells[NomRow,1]:=Nom;
  inc(Nom);
  Excel.Cells[NomRow,2]:='Текущей контроль и промежуточная аттестация';
  i:=0;
  while i<Length(ArrAuditoriiKontrol) do
    begin
    VivodAuditoria(ArrAuditorii[ArrAuditoriiKontrol[i]],7,NomRow+i);
    st:='Microsoft Windows Контракт №007-1-0834-18 от 31.05.2018г. Microsoft Office (Включая WORD, EXCEL и т.д.) Контракт №007-1-0834-18 от 31.05.2018г.';
    Excel.Cells[NomRow+i,5]:=st+' Kaspersky Endpoint Security Контракт №070-1-0478-17 от 18.04.2017г. Kaspersky Endpoint Security Контракт №2075/M21 от 22.01.2015г. Google Chrome';
    inc(i);
    end;
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow+i-1,1]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow+i-1,2]].MergeCells:=true;
  NomRow:=NomRow+i;      }
  Excel.Cells[NomRow,1]:=Nom;
  inc(Nom);
  Excel.Cells[NomRow,2]:='Хранение и профилактическое обслуживание учебного оборудования';
  i:=0;
  while i<Length(ArrAuditoriiObslyz) do
    begin
    VivodAuditoria(ArrAuditorii[ArrAuditoriiObslyz[i]],8,NomRow+i);
    Excel.Cells[NomRow+i,5]:='Дистрибутивы необходимого ПО';
    inc(i);
    end;
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow+i-1,1]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow+i-1,2]].MergeCells:=true;
  NomRow:=NomRow+i;

  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,KolRows]].MergeCells:=true;
  st:='*Специальные помещения - учебные аудитории для проведения занятий лекционного типа, занятий семинарского типа, курсового проектирования (выполнения курсовых работ)';
  Excel.Cells[NomRow,1]:=st+', групповых и индивидуальных консультаций, текущего контроля и промежуточной аттестации, а также помещения для самостоятельной работы.';

  Excel.Range[Excel.Cells[8,1],Excel.Cells[NomRow,KolRows]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[8,1],Excel.Cells[NomRow,KolRows]].BorderAround(1);
  NomRow:=NomRow+2;
  {
  NomStartTabled:=NomRow;
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:='Перечень договоров ЭБС (за период, соответствующий сроку получения образования по ООП)';
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:=1;
  Excel.Cells[NomRow,2]:=2;
  Excel.Cells[NomRow,5]:=3;
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:='Учебный год';
  Excel.Cells[NomRow,2]:='Наименование документа с указанием реквизитов';
  Excel.Cells[NomRow,5]:='Срок действия документа';
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow,6]].HorizontalAlignment:=xlCenter;
  inc(NomRow);
//  Excel.Cells[NomRow,1]:='2018 / 2019';
  if FileExists(CurrentDir+'/МТО/договора ЭБС.txt') then
    begin
    AssignFile(f,CurrentDir+'/МТО/договора ЭБС.txt');
    reset(f);
    while not EOF(f) do
      begin

      readln(f,st);
      if {(CurrentSemestr=1) and (Pos('Б',ArrProfil[NomProfil].Profil)<>0) and (StrToInt('20'+Copy(st,1,2))>=CurrentYear-YearBakalavr) or
         (CurrentSemestr=2) and (Pos('Б',ArrProfil[NomProfil].Profil)<>0) and (StrToInt('20'+Copy(st,4,2))>=CurrentYear-YearBakalavr) or
         (CurrentSemestr=1) and (Pos('М',ArrProfil[NomProfil].Profil)<>0) and (StrToInt('20'+Copy(st,1,2))>=CurrentYear-YearMagistr) or
         (CurrentSemestr=2) and (Pos('М',ArrProfil[NomProfil].Profil)<>0) and (StrToInt('20'+Copy(st,4,2))>=CurrentYear-YearMagistr) {or
         (CurrentSemestr=1) and (Pos('А',ArrProfil[NomProfil].Profil)<>0) and (StrToInt('20'+Copy(st,1,2))>=CurrentYear-YearAspirant) or
         (CurrentSemestr=2) and (Pos('А',ArrProfil[NomProfil].Profil)<>0) and (StrToInt('20'+Copy(st,4,2))>=CurrentYear-YearAspirant) }{ true then
{        begin
        Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,4]].MergeCells:=true;
        Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
        Excel.Cells[NomRow,1]:=st;
        readln(f,st);
        Excel.Cells[NomRow,2]:=st;
        readln(f,st);
        Excel.Cells[NomRow,5]:=st;
        inc(NomRow);
        end;
      end;
    CloseFile(f);
    end;

  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow-1,6]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow-1,6]].BorderAround(1);
  NomRow:=NomRow+2;

  NomStartTabled:=NomRow;
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:='Наименование документа';
  Excel.Cells[NomRow,5]:='Наименование документа (№ документа, дата подписания, организация, выдавшая документ, дата выдачи, срок действия)';
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:=1;
  Excel.Cells[NomRow,5]:=2;
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow,6]].HorizontalAlignment:=xlCenter;
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  st:=' о соответствии зданий, строений, сооружений и помещений, используемых для ведения образовательной деятельности, установленным законодательством РФ требованиям.';
  Excel.Cells[NomRow,1]:='Заключения, выданные в установленном порядке органами, осуществляющими государственный пожарный надзор,'+st;
  Excel.Cells[NomRow,5]:='Заключение №8-28-5-26 о соответствии (несоответствии) объекта защиты требованиям пожарной безопасности от 03 марта 2017 года';
  inc(NomRow);
                                                            }
  {Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow-1,6]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow-1,6]].BorderAround(1); }
{  NomRow:=NomRow+3;
  Excel.Cells[NomRow,1]:='Руководитель организации, ';
  Excel.Cells[NomRow+1,1]:='осуществляющей образовательную деятельность                         ________________________ /______________________ /';
  Excel.Cells[NomRow+2,1]:='                                                                                                                                       подпись                          Ф.И.О. полностью';
  Excel.Cells[NomRow+3,1]:='М.П.';
  Excel.Cells[NomRow+4,1]:='дата составления ________________'; }
//  Excel.Cells[NomRow+1,1]:='';

  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,KolRows]].Font.Size:=8;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,KolRows]].WrapText:=true;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,KolRows]].VerticalAlignment:=xlCenter;
//  Excel.Range[Excel.Cells[1,2],Excel.Cells[NomRow,5]].HorizontalAlignment:=xlCenter;
  if not DirectoryExists(CurrentDir+'\'+NameFolder) then
    ForceDirectories(CurrentDir+'\'+NameFolder);
  if (ArrProfil[NomProfil].Profil<>'') then
    Excel.Workbooks[1].saveas(CurrentDir+'\'+NameFolder+'\'+ArrProfil[NomProfil].Profil+'.xlsx');
  Fmain.MeProtocol.Lines.Add('Создан файл '+CurrentDir+'\'+NameFolder+'\'+ArrProfil[NomProfil].Profil+'.xlsx');
  Excel.Workbooks.Close;
  inc(NomProfil)
  end;
end;

Procedure SaveAllMTO;
var
NomTypeGroup,NomGroup:Longword;
NomPrepod,NomNagr,NomPrepodToo,NomNagrToo:Longword;
NomRow,NomRowNat,Nom,NomStartTabled:Longword;
ArrNagr:array[1..3] of byte;
MaxStudentAuditoria:Longword;
i:byte;
f:TextFile;
st:string;

Procedure VivodInfo(NomRow,NomPrepod,NomNagr:Longword);
var
i:byte;
begin
          st:='Учебная аудитория:                          г.Москва, Волоколамское шоссе, 4, А-80, ГСП-3, 125993, №'+Copy(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria,5,length(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria)-4);
        {  if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[1]<>0 then
            st:=st+'                   Посадочных мест для лекционных занятий - '+IntTostr(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[1]);
          if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[2]<>0 then
            st:=st+'                   Посадочных мест для практических занятий - '+IntTostr(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[2]);
          if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[3]<>0 then
            st:=st+'                   Посадочных мест для лабораторных работ - '+IntTostr(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[3]); }
          Excel.Cells[NomRow,3]:=st;
          if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria='' then
            begin
            Excel.Cells[NomRow,6]:='1';
            if VivodProtocol then
              FMain.MeProtocol.Lines.Add('НЕТ АУД '+Prepod[NomPrepod].FIO+' '+Prepod[NomPrepod].Nagryzka[NomNagr].Dis+' '+Prepod[NomPrepod].Nagryzka[NomNagr].Vid+' '+' '+IntToStr(NomNagr));
//            Excel.Cells[NomRow,3].Pattern:= 1;
//            Excel.Cells[NomRow,3].PatternColorIndex:= -4105;
//            Excel.Cells[NomRow,3].Color:= 65535;
            end;
          if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛК') then  st:='Учебная аудитория для проведения занятий лекционного типа Учебная мебель: столы и стулья для обучающихся; стол и стул для преподавателя; доска с мелом (маркером).'
          else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛР') then  st:='Лаборатория. Специализированное лабораторное оборудование кафедры.'
          else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ПЗ') then  st:='Учебная аудитория для проведения занятий семинарского типа Учебная мебель: столы и стулья для обучающихся; стол и стул для преподавателя; доска с мелом (маркером).'
          else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КП') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='КР') then  st:='Учебная аудитория для проведения занятий семинарского типа Учебная мебель: столы и стулья для обучающихся; стол и стул для преподавателя; доска с мелом (маркером).'
          else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ПЗ') then  st:='Учебная аудитория для проведения занятий семинарского типа Учебная мебель: столы и стулья для обучающихся; стол и стул для преподавателя; доска с мелом (маркером).';

         { MaxStudentAuditoria:=0;
          for i := 1 to 3 do
            if MaxStudentAuditoria<Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[i] then
              MaxStudentAuditoria:=Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[i];
          st:=st+'В аудитории №'+Copy(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria,5,length(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria)-4)+': доска аудиторная, парты на '+IntToStr(MaxStudentAuditoria)+' учащихся, стол преподавателя, стулья';
          if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KomputersAuditoria<>0 then
            st:=st+', '+intToStr(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KomputersAuditoria)+' компьютеров для работы'; }
          st:=st+', '+Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.ProektorAuditoria;
          Excel.Cells[NomRow,4]:=st;
          if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KomputersAuditoria<>0 then
            begin
            st:='Microsoft Windows Контракт №007-1-0834-18 от 31.05.2018г. Microsoft Office (Включая WORD, EXCEL и т.д.) Контракт №007-1-0834-18 от 31.05.2018г.';

            Excel.Cells[NomRow,5]:=st+' Kaspersky Endpoint Security Контракт №070-1-0478-17 от 18.04.2017г. Kaspersky Endpoint Security Контракт №2075/M21 от 22.01.2015г.'+Prepod[NomPrepod].Nagryzka[NomNagr].PO;
            end
          else
            Excel.Cells[NomRow,5]:='Не требуется.';

          if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.OsnashenieOgrnAuditoria<>'' then
            begin
            Excel.Cells[NomRow,6]:=Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.OsnashenieOgrnAuditoria;
            end
          else
            Excel.Cells[NomRow,6]:='Не имеется';
end;

begin

NomTypeGroup:=0;
while NomTypeGroup<length(NameAllGroup) do
  begin
  if (NomTypeGroup<length(NameAllGroup)) and (NameAllGroup[NomTypeGroup].NameGroupKyrs<>'') then
  begin
  NomPrepod:=0;
  while NomPrepod<length(Prepod) do
    begin
    NomNagr:=0;
    while NomNagr<length(Prepod[NomPrepod].Nagryzka) do
      begin
      Prepod[NomPrepod].Nagryzka[NomNagr].FlagVivoda:=false;
      inc(NomNagr);
      end;
    inc(NomPrepod);
    end;
  Excel.WorkBooks.Add;

  Excel.Columns[1].ColumnWidth := 3.43;
  Excel.Columns[2].ColumnWidth := 21.57;
  Excel.Columns[3].ColumnWidth := 27.00;
  Excel.Columns[4].ColumnWidth := 20.14;
  Excel.Columns[5].ColumnWidth := 62.43;
  Excel.Columns[6].ColumnWidth := 25.29;

  Excel.Range[Excel.Cells[1,1],Excel.Cells[1,6]].MergeCells:=true;
  Excel.Cells[1,1]:='Федеральное государственное бюджетное образовательное учреждение высшего образования';
  Excel.Range[Excel.Cells[2,1],Excel.Cells[2,6]].MergeCells:=true;
  Excel.Cells[2,1]:='«Московский авиационный институт (национальный исследовательский университет)»';
  Excel.Range[Excel.Cells[4,1],Excel.Cells[4,6]].MergeCells:=true;
  Excel.Cells[4,1]:='Справка';
  Excel.Range[Excel.Cells[5,1],Excel.Cells[5,6]].MergeCells:=true;
  st:='о материально-техническом обеспечении основной образовательной программы высшего образования – программы';
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
    Excel.Cells[6,1]:='09.04.04 "Программная инженерия" профиль - Программно-информационные системы'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='03А' then
    Excel.Cells[6,1]:='09.06.01 (05.13.15 Вычислительные машины, комплексы и компьютерные сети)'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='05А' then
    Excel.Cells[6,1]:='09.06.01 (05.13.11 Математическое и программное обеспечение вычислительных машин, комплексов и компьютерных сетей)' ;
  Excel.Cells[8,1]:='№ п\п';
  Excel.Cells[8,2]:='Наименование дисциплины (модуля), практик в соответствии с учебным планом';
  Excel.Cells[8,3]:='Наименование специальных* помещений и помещений для самостоятельной работы';
  Excel.Cells[8,4]:='Оснащенность специальных помещений и помещений для самостоятельной работы';
  Excel.Cells[8,5]:='Перечень лицензионного программного обеспечения. Реквизиты подтверждающего документа ';
  Excel.Cells[8,6]:='Приспособленность помещений для использования инвалидами и лицами с ограниченными возможностями здоровья';
  for i := 1 to 6 do
    Excel.Cells[9,i]:=i;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[9,6]].HorizontalAlignment:=xlCenter;

  NomRow:=10;
  Nom:=1;
{  NomGroup:=0;
  while NomGroup<length(NameAllGroup[NomTypeGroup].Group) do
    begin }
//    FMain.MeProtocol.Lines.Add(NameAllGroup[NomTypeGroup].Group[NomGroup]);
    NomPrepod:=0;
    while NomPrepod<length(Prepod) do
      begin
      NomNagr:=0;
      while NomNagr<length(Prepod[NomPrepod].Nagryzka) do
        begin
        if ((SearchInMassiveGroup(Prepod[NomPrepod].Nagryzka[NomNagr].Group, NameAllGroup[NomTypeGroup].NameGroupKyrs)<>65000)) and
          (not Prepod[NomPrepod].Nagryzka[NomNagr].FlagVivoda) and
           ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛК') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛР') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ПЗ')) then
          begin
          Excel.Cells[NomRow,1]:=Nom;
          Excel.Cells[NomRow,2]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
          VivodInfo(NomRow,NomPrepod,NomNagr);
          inc(NomRow);
          inc(Nom);

          for i := 1 to 3 do
            ArrNagr[i]:=0;
          if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛК') then  ArrNagr[1]:=1;
          if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ЛР') then  ArrNagr[2]:=1;
          if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='ПЗ') then  ArrNagr[3]:=1;
          Prepod[NomPrepod].Nagryzka[NomNagr].FlagVivoda:=true;

          //Найти все виды деятельности по данному предмету
          NomRowNat:=NomRow-1;
          NomPrepodToo:=0;
          while NomPrepodToo<length(Prepod) do
            begin
            NomNagrToo:=0;
            while NomNagrToo<length(Prepod[NomPrepodToo].Nagryzka) do
              begin
              if ((SearchInMassiveGroup(Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Group, NameAllGroup[NomTypeGroup].NameGroupKyrs)<>65000)) and
                 (Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Dis=Prepod[NomPrepod].Nagryzka[NomNagr].Dis) and
                 (not Prepod[NomPrepodToo].Nagryzka[NomNagrToo].FlagVivoda)  then
                begin
                Prepod[NomPrepodToo].Nagryzka[NomNagrToo].FlagVivoda:=true;
                if
                 (((Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Vid='ЛК') and (ArrNagr[1]=0)) or
                  ((Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Vid='ЛР') and (ArrNagr[2]=0)) or
                  ((Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Vid='ПЗ') and (ArrNagr[3]=0))) then
                begin
                if (Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Vid='ЛК') then  ArrNagr[1]:=1;
                if (Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Vid='ЛР') then  ArrNagr[2]:=1;
                if (Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Vid='ПЗ') then  ArrNagr[3]:=1;

                VivodInfo(NomRow,NomPrepodToo,NomNagrToo);
                inc(NomRow);
                end;
                end;
              inc(NomNagrToo);
              end;
            inc(NomPrepodToo);
            end;
          Excel.Range[Excel.Cells[NomRowNat,1],Excel.Cells[NomRow-1,1]].MergeCells:=true;
          Excel.Range[Excel.Cells[NomRowNat,2],Excel.Cells[NomRow-1,2]].MergeCells:=true;
          end;
        inc(NomNagr);
        end;
      inc(NomPrepod);
      end;
{    inc(NomGroup);
    end;   }
   Excel.Cells[NomRow,1]:=Nom;
   Excel.Cells[NomRow,2]:='Помещения для самостоятельной работы студетов';
   Excel.Cells[NomRow,3]:='Учебная аудитория:                          г.Москва, Ул Оршанская, 3, № Б-707,          Посадочных мест для самостоятельной работы студентов - 26';
   Excel.Cells[NomRow,4]:='Самостоятельная работа.                В аудитории Б-707: доска аудиторная, парты на 20 учащихся, стол преподавателя, стулья, 11 компьютеров для работы';
   st:='Microsoft Windows Контракт №007-1-0834-18 от 31.05.2018г. Microsoft Office (Включая WORD, EXCEL и т.д.) Контракт №007-1-0834-18 от 31.05.2018г. Kaspersky Endpoint Security Контракт №070-1-0478-17 от 18.04.2017г. ';
   Excel.Cells[NomRow,5]:=st+'Kaspersky Endpoint Security Контракт №2075/M21 от 22.01.2015г. COMP-P(бесплатное ПО, FreeWare) Microsoft Visual Studio(бесплатное ПО);MinGW(бесплатное ПО,  GNU General Public License);';
   Excel.Cells[NomRow+1,3]:='Учебная аудитория:                          г.Москва, Ул Оршанская, 3, № Б-711,          Посадочных мест для самостоятельной работы студентов - 26';
   Excel.Cells[NomRow+1,4]:='Самостоятельная работа.                В аудитории Б-711: доска аудиторная, парты на 20 учащихся, стол преподавателя, стулья, 11 компьютеров для работы';
   st:='Microsoft Windows Контракт №007-1-0834-18 от 31.05.2018г. Microsoft Office (Включая WORD, EXCEL и т.д.) Контракт №007-1-0834-18 от 31.05.2018г. Kaspersky Endpoint Security Контракт №070-1-0478-17 от 18.04.2017г. ';
   Excel.Cells[NomRow+1,5]:=st+'Kaspersky Endpoint Security Контракт №2075/M21 от 22.01.2015г. COMP-P(бесплатное ПО, FreeWare) Microsoft Visual Studio(бесплатное ПО);MinGW(бесплатное ПО,  GNU General Public License);';
   Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow+1,1]].MergeCells:=true;
   Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow+1,2]].MergeCells:=true;
   Excel.Cells[NomRow+2,1]:=Nom+1;
   Excel.Cells[NomRow+2,2]:='Помещения для обслуживания лабораторной и орг. техники';
   Excel.Cells[NomRow+2,3]:='Аудитория:                                       г.Москва, Волоколамское шоссе, 4, А-80, ГСП-3, 125993, №434в(3)';
   Excel.Cells[NomRow+2,4]:='В помещении №434в(3) могут работать 5 сотрудников. Помещение предназначено для НИР';
   Excel.Cells[NomRow+2,5]:='Комплект снаряжения и расходных материалов для ремонта вычислительной и орг. техники';
   Excel.Cells[NomRow+3,3]:='В лабораториях для текущего ремонта';
   Excel.Cells[NomRow+3,4]:='';
   Excel.Cells[NomRow+3,5]:='Комплект снаряжения и расходных материалов для ремонта вычислительной и орг. техники';
   Excel.Range[Excel.Cells[NomRow+2,1],Excel.Cells[NomRow+3,1]].MergeCells:=true;
   Excel.Range[Excel.Cells[NomRow+2,2],Excel.Cells[NomRow+3,2]].MergeCells:=true;
   Excel.Cells[NomRow+4,1]:=Nom+2;
   Excel.Cells[NomRow+4,2]:='Помещения для методической работы';
   Excel.Cells[NomRow+4,3]:='Аудитория:                                           г.Москва, Волоколамское шоссе, 4, А-80, ГСП-3, 125993, №217(3)';
   Excel.Cells[NomRow+4,4]:='В помещении №217(3) могут одновременно работать 3 сотрудников. Осназение: 3 компьютера, столы, стулья, принтер, ксерокс';
   st:='Microsoft Windows Контракт №007-1-0834-18 от 31.05.2018г. Microsoft Office (Включая WORD, EXCEL и т.д.) Контракт №007-1-0834-18 от 31.05.2018г. Kaspersky Endpoint Security Контракт №070-1-0478-17 от 18.04.2017г. ';
   Excel.Cells[NomRow+4,5]:=st+'Kaspersky Endpoint Security Контракт №2075/M21 от 22.01.2015г.';

   for I := 0 to 4 do
     Excel.Cells[NomRow+i,6]:='Не имеется';

   NomRow:=NomRow+5;

  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,6]].MergeCells:=true;
  st:='*Специальные помещения - учебные аудитории для проведения занятий лекционного типа, занятий семинарского типа, курсового проектирования (выполнения курсовых работ)';
  Excel.Cells[NomRow,1]:=st+', групповых и индивидуальных консультаций, текущего контроля и промежуточной аттестации, а также помещения для самостоятельной работы.';

  Excel.Range[Excel.Cells[8,1],Excel.Cells[NomRow,6]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[8,1],Excel.Cells[NomRow,6]].BorderAround(1);
  NomRow:=NomRow+2;

  NomStartTabled:=NomRow;
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:='Перечень договоров ЭБС (за период, соответствующий сроку получения образования по ООП)';
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:=1;
  Excel.Cells[NomRow,2]:=2;
  Excel.Cells[NomRow,5]:=3;
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:='Учебный год';
  Excel.Cells[NomRow,2]:='Наименование документа с указанием реквизитов';
  Excel.Cells[NomRow,5]:='Срок действия документа';
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow,6]].HorizontalAlignment:=xlCenter;
  inc(NomRow);
//  Excel.Cells[NomRow,1]:='2018 / 2019';
  if FileExists(CurrentDir+'/МТО/договора ЭБС.txt') then
    begin
    AssignFile(f,CurrentDir+'/МТО/договора ЭБС.txt');
    reset(f);
    while not EOF(f) do
      begin

      readln(f,st);
      if (CurrentSemestr=1) and (NameAllGroup[NomTypeGroup].NameGroupKyrs[3]='Б') and (StrToInt('20'+Copy(st,1,2))>=CurrentYear-YearBakalavr) or
         (CurrentSemestr=2) and (NameAllGroup[NomTypeGroup].NameGroupKyrs[3]='Б') and (StrToInt('20'+Copy(st,4,2))>=CurrentYear-YearBakalavr) or
         (CurrentSemestr=1) and (NameAllGroup[NomTypeGroup].NameGroupKyrs[3]='М') and (StrToInt('20'+Copy(st,1,2))>=CurrentYear-YearMagistr) or
         (CurrentSemestr=2) and (NameAllGroup[NomTypeGroup].NameGroupKyrs[3]='М') and (StrToInt('20'+Copy(st,4,2))>=CurrentYear-YearMagistr) or
         (CurrentSemestr=1) and (NameAllGroup[NomTypeGroup].NameGroupKyrs[3]='А') and (StrToInt('20'+Copy(st,1,2))>=CurrentYear-YearAspirant) or
         (CurrentSemestr=2) and (NameAllGroup[NomTypeGroup].NameGroupKyrs[3]='А') and (StrToInt('20'+Copy(st,4,2))>=CurrentYear-YearAspirant) then
        begin
        Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,4]].MergeCells:=true;
        Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
        Excel.Cells[NomRow,1]:=st;
        readln(f,st);
        Excel.Cells[NomRow,2]:=st;
        readln(f,st);
        Excel.Cells[NomRow,5]:=st;
        inc(NomRow);
        end;
      end;
    CloseFile(f);
    end;

  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow-1,6]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow-1,6]].BorderAround(1);
  NomRow:=NomRow+2;

  NomStartTabled:=NomRow;
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:='Наименование документа';
  Excel.Cells[NomRow,5]:='Наименование документа (№ документа, дата подписания, организация, выдавшая документ, дата выдачи, срок действия)';
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:=1;
  Excel.Cells[NomRow,5]:=2;
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow,6]].HorizontalAlignment:=xlCenter;
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  st:=' о соответствии зданий, строений, сооружений и помещений, используемых для ведения образовательной деятельности, установленным законодательством РФ требованиям.';
  Excel.Cells[NomRow,1]:='Заключения, выданные в установленном порядке органами, осуществляющими государственный пожарный надзор,'+st;
  Excel.Cells[NomRow,5]:='Заключение №8-28-5-26 о соответствии (несоответствии) объекта защиты требованиям пожарной безопасности от 03 марта 2017 года';
  inc(NomRow);

  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow-1,6]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow-1,6]].BorderAround(1);
  NomRow:=NomRow+3;
  Excel.Cells[NomRow,1]:='Руководитель организации, ';
  Excel.Cells[NomRow+1,1]:='осуществляющей образовательную деятельность                         ________________________ /_Брехов Олег Михайлович_ /';
  Excel.Cells[NomRow+2,1]:='                                                                                                                                       подпись                          Ф.И.О. полностью';
  Excel.Cells[NomRow+3,1]:='М.П.';
  Excel.Cells[NomRow+4,1]:='дата составления ________________';
//  Excel.Cells[NomRow+1,1]:='';

  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,6]].Font.Size:=8;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,6]].WrapText:=true;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,6]].VerticalAlignment:=xlCenter;
//  Excel.Range[Excel.Cells[1,2],Excel.Cells[NomRow,5]].HorizontalAlignment:=xlCenter;
  if not DirectoryExists(CurrentDir+'\МТО') then
    ForceDirectories(CurrentDir+'\МТО');
  if (NomTypeGroup<length(NameAllGroup)) and (NameAllGroup[NomTypeGroup].NameGroupKyrs<>'') then
    Excel.Workbooks[1].saveas(CurrentDir+'\МТО\'+NameAllGroup[NomTypeGroup].NameGroupKyrs+'.xlsx');
  Fmain.MeProtocol.Lines.Add('Создан файл '+CurrentDir+'\МТО\'+NameAllGroup[NomTypeGroup].NameGroupKyrs+'.xlsx');
  Excel.Workbooks.Close;
  end;
  inc(NomTypeGroup);
  end;
end;


end.
