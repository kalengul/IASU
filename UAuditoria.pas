unit UAuditoria;

interface

type
   TAuditoria = class
                    Auditoria:String;
                    Korpus:string;
                    KolStudentAuditoriaMax:array [1..3] of Word;
                    KomputersAuditoria:word;
                    ProektorAuditoria:string;
                    OsnashenieOgrnAuditoria,NameAuditoria:string;
                    Constructor Create;
                    Destructor Destroy;
  end;

var
  ArrAuditorii:array of TAuditoria;

  function SearchAndAddInMassAuditoriaName(Name:string):TAuditoria;
function SearchInMassAuditoriaName(Name:string):TAuditoria;
Procedure AddAllAud(NomKaf:string);

//Excel
Procedure CreateAllPOInAud(TypeVivod:byte);


implementation

uses USemPlan,UConstParametrs,UMain, UNagryzka,SysUtils;

function SearchAndAddInMassAuditoriaName(Name:string):TAuditoria;
var
NomAuditoria:Longword;
begin
NomAuditoria:=0;
while (NomAuditoria<Length(ArrAuditorii)) and (ArrAuditorii[NomAuditoria].Auditoria<>Name) do
  inc(NomAuditoria);
if NomAuditoria<Length(ArrAuditorii) then
  Result:=ArrAuditorii[NomAuditoria]
else
  begin
  NomAuditoria:=Length(ArrAuditorii);
  setLength(ArrAuditorii,NomAuditoria+1);
  ArrAuditorii[NomAuditoria]:=TAuditoria.Create;
  ArrAuditorii[NomAuditoria].Auditoria:=Name;
  Result:=ArrAuditorii[NomAuditoria];
  end;
end;

function SearchInMassAuditoriaName(Name:string):TAuditoria;
var
NomAuditoria:Longword;
begin
NomAuditoria:=0;
while (NomAuditoria<Length(ArrAuditorii)) and (ArrAuditorii[NomAuditoria].Auditoria<>Name) do
  inc(NomAuditoria);
if NomAuditoria<Length(ArrAuditorii) then
  Result:=ArrAuditorii[NomAuditoria]
else
  Result:=ArrAuditorii[0];
end;

Procedure AddAllAud(NomKaf:string);
var
 NomSP,NomDisSp:longword;
 Auditoria:string;
  begin
NomSP:=0;
While NomSp<Length(SemYP) do
  begin

    NomDisSp:=0;
    while NomDisSp<Length(SemYP[NomSp].Disciplin) do
      begin
      if Pos(NomKaf,SemYP[NomSp].Disciplin[NomDisSp].Kaf)<>0 then
        begin
        if (SemYP[NomSp].Disciplin[NomDisSp].LKAud=nil) and (SemYP[NomSp].Disciplin[NomDisSp].LK<>0) then
          begin
          Auditoria:=KafAudLK[random(10)];
          SemYP[NomSp].Disciplin[NomDisSp].LKAud:=SearchInMassAuditoriaName(Auditoria);
          if SemYP[NomSp].Disciplin[NomDisSp].NomElektivDis<>65000 then
            SemYP[NomSp].Disciplin[SemYP[NomSp].Disciplin[NomDisSp].NomElektivDis].LKAud:=SemYP[NomSp].Disciplin[NomDisSp].LKAud;
          end;
        if (SemYP[NomSp].Disciplin[NomDisSp].LRAud=nil) and (SemYP[NomSp].Disciplin[NomDisSp].LR<>0) then
          begin
          Auditoria:=KafAudLR[random(8)];
          SemYP[NomSp].Disciplin[NomDisSp].LRAud:=SearchInMassAuditoriaName(Auditoria);
          if SemYP[NomSp].Disciplin[NomDisSp].NomElektivDis<>65000 then
            SemYP[NomSp].Disciplin[SemYP[NomSp].Disciplin[NomDisSp].NomElektivDis].LRAud:=SemYP[NomSp].Disciplin[NomDisSp].LRAud;
          end;
        if (SemYP[NomSp].Disciplin[NomDisSp].PZAud=nil) and (SemYP[NomSp].Disciplin[NomDisSp].PZ<>0) then
          begin
          Auditoria:=KafAudLR[random(8)];
          SemYP[NomSp].Disciplin[NomDisSp].PZAud:=SearchInMassAuditoriaName(Auditoria);
          if SemYP[NomSp].Disciplin[NomDisSp].NomElektivDis<>65000 then
            SemYP[NomSp].Disciplin[SemYP[NomSp].Disciplin[NomDisSp].NomElektivDis].PZ:=SemYP[NomSp].Disciplin[NomDisSp].PZ;
          end;
        end;
      inc(NomDisSp);
      end;


  inc(NomSp);
  end;
  end;

Procedure CreateAllPOInAud(TypeVivod:byte);
var
NomPrepod,NomNagryzka,NomAud,KolAud,NomRow,NomRowAnd:longword;
Aud:array of String;
Po,PoVivod:String;
begin

KolAud:=0;
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  NomNagryzka:=0;
  while NomNagryzka<Length(Prepod[NomPrepod].Nagryzka) do
    begin
    NomAud:=0;
    while (NomAud<KolAud) and (Prepod[NomPrepod].Nagryzka[NomNagryzka].Auditoria.Auditoria<>Aud[NomAud]) do
      inc(NomAud);
    if not (NomAud<KolAud) then
      begin
      SetLength(Aud,KolAud+1);
      Aud[KolAud]:=Prepod[NomPrepod].Nagryzka[NomNagryzka].Auditoria.Auditoria;
      inc(KolAud);
      end;
    inc(NomNagryzka);
    end;
  inc(NomPrepod);
  end;
Excel.WorkBooks.Add;
NomAud:=0;
NomRow:=2;
while (NomAud<KolAud) do
  begin
  Po:='';
  NomPrepod:=0;
  while NomPrepod<Length(Prepod) do
    begin
    NomNagryzka:=0;
    while NomNagryzka<Length(Prepod[NomPrepod].Nagryzka) do
      begin
      if (Prepod[NomPrepod].Nagryzka[NomNagryzka].Auditoria.Auditoria=Aud[NomAud]) and (Prepod[NomPrepod].Nagryzka[NomNagryzka].PO<>'') then
        Po:=Po+Prepod[NomPrepod].Nagryzka[NomNagryzka].PO;
      inc(NomNagryzka);
      end;
    inc(NomPrepod);
    end;
  if Po<>'' then
  begin
  Po:=Po+';';
  PoVivod:='';
  Excel.Cells[NomRow,1]:=Aud[NomAud];
  NomRowAnd:=NomRow;
  while Po<>'' do
    begin
    if (pos(Copy(Po,1,Pos(';',Po)),PoVivod)=0) then
      begin
      Excel.Cells[NomRow,2]:=Copy(Po,1,Pos(';',Po));
      PoVivod:=PoVivod+Copy(Po,1,Pos(';',Po));
      inc(NomRow);
      end;
    Delete(Po,1,Pos(';',Po));
    end;
  Excel.Range[Excel.Cells[NomRowAnd,1],Excel.Cells[NomRow-1,1]].MergeCells:=true;
  end;
  inc(NomAud);
  end;
Excel.Workbooks[1].saveas(CurrentDir+'\ПО в аудиториях.xlsx');
Fmain.MeProtocol.Lines.Add('Создан файл '+CurrentDir+'\ПО в аудиториях.xlsx');
Excel.Workbooks.Close;
end;



Destructor TAuditoria.Destroy;
  begin

  end;
Constructor TAuditoria.Create;
  begin

  end;


end.
