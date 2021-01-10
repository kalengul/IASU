unit USemPlan;

interface

uses UGroup,UAuditoria;

type
  TDisciplineSemestrYP = record
    Name,Kaf:string;
    Sem,Kyrs:Longword;
    NedChas:Double;
    LK,LR,PZ,SRS:Longword;
    KR,RGR,DZ,REF,Kolokv,Kontr,Testir:byte;
    LKAud,LRAud,PZAud:TAuditoria;
    PO:string;
    VidKontrolia:string;
    ChasEkzamen:string;
    EndDiscipline:byte;
    BYCh:boolean;
    NomElektivDis:Longword;
    OsnashenieRPD,PoRPD:string;
  end;

  TSemestroviYP = class
    Nom,Kaf,Napravlenie,Profil:string;
    Disciplin:array of TDisciplineSemestrYP;
    Group:TAGroup;
    Constructor Create;
    Destructor Destroy;
  end;

  TProfil = record
            SemYp:array of  TSemestroviYP;
            NameProfil,NameNaprav,Profil,Naprav:string;
            end;

var
  SemYP:array of TSemestroviYP;
  ArrProfil:array of TProfil;

Procedure CopyDisSemPlan(var DisCopy,Dis:TDisciplineSemestrYP);
Procedure SortSemPlan;
Procedure LoadSemPlan;

implementation

uses UMain, SysUtils;

Procedure LoadSemPlan;
var
  NomRow,KolSemYP,KolDis,KolGroup,NomGroup:Longword;
  SearchStr:TSearchRec;
  st,st1:string;
  Sem,Kyrs:Longword;
  NomProfil,KolSemYpProfil:Longword;
begin
if not DirectoryExists(CurrentDir+'\Семестровый план') then
  ForceDirectories(CurrentDir+'\Семестровый план');
KolSemYP:=0;
if FindFirst(CurrentDir+'\Семестровый план\'+'*.xls*',faDirectory,SearchStr)=0 then
  begin
  repeat
  Excel.Workbooks.Open(CurrentDir+'\Семестровый план\'+SearchStr.Name);
  SetLength(SemYP,KolSemYP+1);
  SemYP[KolSemYP]:=TSemestroviYP.Create;
  NomRow:=1;
  KolDis:=0;
  KolGroup:=0;
  st:=Excel.Cells[NomRow,2];
  st1:=Excel.Cells[NomRow+1,1];
  While  not ((st='Начальник отдела БД и СФО') and (st1='')) do
    begin
    SemYP[KolSemYP].Nom:=Excel.Cells[NomRow+2,11];
    Delete(SemYP[KolSemYP].Nom,1,18);
    SemYP[KolSemYP].Kaf:=Excel.Cells[NomRow+1,56];
    SemYP[KolSemYP].Napravlenie:=Excel.Cells[NomRow+5,16];
    SemYP[KolSemYP].Profil:=Excel.Cells[NomRow+5,20];
    if NomRow=1 then
      begin
      NomProfil:=0;
      while (NomProfil<Length(ArrProfil)) and (ArrProfil[NomProfil].Profil<>SemYP[KolSemYP].Profil) do
        inc(NomProfil);
      if not (NomProfil<Length(ArrProfil)) then
        begin
        SetLength(ArrProfil,NomProfil+1);
        ArrProfil[NomProfil].Profil:=SemYP[KolSemYP].Profil;
        ArrProfil[NomProfil].Naprav:=SemYP[KolSemYP].Napravlenie;
        end;
      KolSemYpProfil:=Length(ArrProfil[NomProfil].SemYp);
      SetLength(ArrProfil[NomProfil].SemYp,KolSemYpProfil+1);
      ArrProfil[NomProfil].SemYp[KolSemYpProfil]:=SemYP[KolSemYP];
      end;
    {SemYP[KolSemYP].Group}st:=Excel.Cells[NomRow+5,29];    //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    NomGroup:=0;
    while NomGroup<Length(AllGroup) do
      begin
      if Pos(AllGroup[NomGroup].Nom,st)<>0 then
        begin
        SetLength(SemYP[KolSemYP].Group,KolGroup+1);
        SemYP[KolSemYP].Group[KolGroup]:=AllGroup[NomGroup];
        inc(KolGroup);
        if ArrProfil[NomProfil].NameProfil='' then
          begin
          ArrProfil[NomProfil].NameProfil:=AllGroup[NomGroup].Profil;
          ArrProfil[NomProfil].NameNaprav:=AllGroup[NomGroup].Napravlenie;
          end;
        end;
      inc(NomGroup);
      end;
    Kyrs:=Excel.Cells[NomRow+5,12];
    Sem:=Excel.Cells[NomRow+5,14];
    NomRow:=NomRow+10;
    st:=Excel.Cells[NomRow,3];
    while st<>'Итого:' do
      begin
      SetLength(SemYP[KolSemYP].Disciplin,KolDis+1);
      SemYP[KolSemYP].Disciplin[KolDis].Name:=Excel.Cells[NomRow,3];
      SemYP[KolSemYP].Disciplin[KolDis].Kaf:=Excel.Cells[NomRow,15];
      SemYP[KolSemYP].Disciplin[KolDis].Sem:=Sem;
      SemYP[KolSemYP].Disciplin[KolDis].Kyrs:=Kyrs;
      SemYP[KolSemYP].Disciplin[KolDis].NedChas:=Excel.Cells[NomRow,22];
      SemYP[KolSemYP].Disciplin[KolDis].LK:=Excel.Cells[NomRow,26];
      SemYP[KolSemYP].Disciplin[KolDis].LR:=Excel.Cells[NomRow,28];
      SemYP[KolSemYP].Disciplin[KolDis].PZ:=Excel.Cells[NomRow,31];
      SemYP[KolSemYP].Disciplin[KolDis].LKAud:=nil;
      SemYP[KolSemYP].Disciplin[KolDis].LRAud:=nil;
      SemYP[KolSemYP].Disciplin[KolDis].PZAud:=nil;
      SemYP[KolSemYP].Disciplin[KolDis].PO:='';
      st:=Excel.Cells[NomRow,45];
      if Pos('(',st)<>0 then
        Delete(st,Pos('(',st)-1,Length(st)-Pos('(',st)+2);
      if st<>'' then
        SemYP[KolSemYP].Disciplin[KolDis].SRS:=StrToInt(st);
      SemYP[KolSemYP].Disciplin[KolDis].KR:=0;
      st:=Excel.Cells[NomRow,35];
      if st<>'' then
        SemYP[KolSemYP].Disciplin[KolDis].KR:=1;
      st:=Excel.Cells[NomRow,36];
      if st<>'' then
        SemYP[KolSemYP].Disciplin[KolDis].KR:=1;
      st:=Excel.Cells[NomRow,39];
      if st<>'' then
        SemYP[KolSemYP].Disciplin[KolDis].RGR:=Excel.Cells[NomRow,39]
      else
        SemYP[KolSemYP].Disciplin[KolDis].RGR:=0;
      st:=Excel.Cells[NomRow,41];
      if st<>'' then
        SemYP[KolSemYP].Disciplin[KolDis].DZ:=Excel.Cells[NomRow,41]
      else
        SemYP[KolSemYP].Disciplin[KolDis].DZ:=0;
      st:=Excel.Cells[NomRow,42];
      if st<>'' then
        SemYP[KolSemYP].Disciplin[KolDis].REF:=Excel.Cells[NomRow,42]
      else
        SemYP[KolSemYP].Disciplin[KolDis].REF:=0;
      st:=Excel.Cells[NomRow,47];
      if st<>'' then
        SemYP[KolSemYP].Disciplin[KolDis].Kolokv:=Excel.Cells[NomRow,47]
      else
        SemYP[KolSemYP].Disciplin[KolDis].Kolokv:=0;
      st:=Excel.Cells[NomRow,48];
      if st<>'' then
        SemYP[KolSemYP].Disciplin[KolDis].Kontr:=Excel.Cells[NomRow,48]
      else
        SemYP[KolSemYP].Disciplin[KolDis].Kontr:=0;
      st:=Excel.Cells[NomRow,50];
      if st<>'' then
        SemYP[KolSemYP].Disciplin[KolDis].Testir:=Excel.Cells[NomRow,50]
      else
        SemYP[KolSemYP].Disciplin[KolDis].Testir:=0;
      SemYP[KolSemYP].Disciplin[KolDis].VidKontrolia:=Excel.Cells[NomRow,55];
      SemYP[KolSemYP].Disciplin[KolDis].ChasEkzamen:=Excel.Cells[NomRow,56];
      SemYP[KolSemYP].Disciplin[KolDis].EndDiscipline:=0;
      st:=Excel.Cells[NomRow,58];
      if st<>'' then
        SemYP[KolSemYP].Disciplin[KolDis].EndDiscipline:=1;
      st:=Excel.Cells[NomRow,2];
      if pos('.',st)=0 then
        SemYP[KolSemYP].Disciplin[KolDis].NomElektivDis:=65000
      else
        begin
        st1:=Excel.Cells[NomRow-1,2];
        if (st1<>'') and (pos('.',st1)<>0) and (st[1]=st1[1]) then
          begin
          SemYP[KolSemYP].Disciplin[KolDis].NomElektivDis:=KolDis-1;
          SemYP[KolSemYP].Disciplin[KolDis-1].NomElektivDis:=KolDis;
          //CopyDisSemPlan(KolDis-1,KolDis);
          end;
        end;

      inc(KolDis);
      inc(NomRow);
      st:=Excel.Cells[NomRow,3];
      end;
    while st<>'Начальник отдела БД и СФО' do
      begin
      inc(NomRow);
      st:=Excel.Cells[NomRow,2];
      end;
    st1:=Excel.Cells[NomRow+1,1];
    inc(NomRow);
    end;
  inc(KolSemYP);
  Excel.Workbooks.Close;
  FMain.MeProtocol.Lines.Add('Загружен Семестровый план: '+CurrentDir+'\Семестровый план\'+SearchStr.Name);
  until FindNext(SearchStr)<>0;
  end;
FMain.MeProtocol.Lines.Add('Загрузка семестровых планов заверщена');
end;



Procedure SortSemPlan;
var
  SearchStr:TSearchRec;
  NomProfil,NomSemPlan,NomDis:longword;
  NomStr,KolStr,NomElektivDis:Longword;
  St:String;
  buf:TDisciplineSemestrYP;
begin
if not DirectoryExists(CurrentDir+'\Семестровый план\Сортировка') then
  ForceDirectories(CurrentDir+'\Семестровый план\Сортировка');
NomProfil:=0;
while (NomProfil<Length(ArrProfil)) do
  begin
  if FindFirst(CurrentDir+'\Семестровый план\Сортировка\'+ArrProfil[NomProfil].Profil+'сорт'+'.xls*',faDirectory,SearchStr)=0 then
    begin
    Excel.Workbooks.Open(CurrentDir+'\Семестровый план\Сортировка\'+SearchStr.Name);
    NomSemPlan:=0;
    while NomSemPlan<Length(ArrProfil[NomProfil].SemYp) do
      begin
      KolStr:=Excel.Cells[1,1];
      NomStr:=0;
      while NomStr<KolStr do
        begin
        st:=Excel.Cells[NomStr+2,1];   //Загрузили строчку с дисциплиной из Экселя
        //Нашли такуюже строку в дисциплинах
        NomDis:=0;
        while (NomDis<Length(ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin)) and (ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDis].Name<>st) do
          inc(NomDis);

        if (NomDis<Length(ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin)) then
          begin
          if ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDis].NomElektivDis<>65000 then
            begin
            ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDis].NomElektivDis].NomElektivDis:=NomStr;
            end;

          CopyDisSemPlan(ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDis],buf);
          CopyDisSemPlan(ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomStr],ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomDis]);
          CopyDisSemPlan(buf,ArrProfil[NomProfil].SemYp[NomSemPlan].Disciplin[NomStr]);
          //Поменяли местами (поставили на позицию NomStr), если нашли


          end;
        inc(NomStr);
        end;
      inc(NomSemPlan);
      end;
    Excel.Workbooks.Close;
    FMain.MeProtocol.Lines.Add('Загружена сортировка Семестрового плана: '+CurrentDir+'\Семестровый план\Сортировка\'+SearchStr.Name);
    end;
  inc(NomProfil);
  end;
FMain.MeProtocol.Lines.Add('Сортировка семестровых планов заверщена');
end;

Procedure CopyDisSemPlan(var DisCopy,Dis:TDisciplineSemestrYP);
begin
Dis.Name:=DisCopy.Name;
Dis.Kaf:=DisCopy.Kaf;
Dis.Sem:=DisCopy.Sem;
Dis.Kyrs:=DisCopy.Kyrs;
Dis.NedChas:=DisCopy.NedChas;
Dis.LK:=DisCopy.LK;
Dis.LR:=DisCopy.LR;
Dis.PZ:=DisCopy.PZ;
Dis.LKAud:=DisCopy.LKAud;
Dis.LRAud:=DisCopy.LRAud;
Dis.PZAud:=DisCopy.PZAud;
Dis.PO:=DisCopy.PO;
Dis.SRS:=DisCopy.SRS;
Dis.KR:=DisCopy.KR;
Dis.RGR:=DisCopy.RGR;
Dis.DZ:=DisCopy.DZ;
Dis.REF:=DisCopy.REF;
Dis.Kolokv:=DisCopy.Kolokv;
Dis.Kontr:=DisCopy.Kontr;
Dis.Testir:=DisCopy.Testir;
Dis.VidKontrolia:=DisCopy.VidKontrolia;
Dis.ChasEkzamen:=DisCopy.ChasEkzamen;
Dis.EndDiscipline:=DisCopy.EndDiscipline;
Dis.OsnashenieRPD:=DisCopy.OsnashenieRPD;
Dis.PoRPD:=DisCopy.PoRPD;
Dis.NomElektivDis:=DisCopy.NomElektivDis;
end;

Constructor TSemestroviYP.Create;
  begin
  SetLength(Disciplin,0);
  SetLength(Group,0);
  end;
Destructor TSemestroviYP.Destroy;
  begin
  SetLength(Disciplin,0);
  SetLength(Group,0);
  end;

end.
