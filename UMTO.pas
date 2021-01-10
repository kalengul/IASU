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
  if (St<>'') and (pos('�� ����������',st)<>0) then
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

Fmain.MeProtocol.Lines.Add('��������� ��� �� ���� �� �����:'+FileName);
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
        ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��')) then
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
Fmain.MeProtocol.Lines.Add('��������� ����������� ����������� �� ����� '+FileName);
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
    1:st:='������� ��������� ��� ���������� ������� ����������� ����:';
    2:if Aud.NameAuditoria<>'' then st:='�����������: "'+Aud.NameAuditoria+'"' else st:='�����������.'{ ��� ���������� ������� �� ����� "'+NameDisLr+'"'};
    3:st:='������� ��������� ��� ���������� ������� ������������ ����:';
    4:st:='������� ��������� ��� ��������� �������������� (���������� �������� �����):';
    5:st:='������� ��������� ��� ��������� � �������������� ������������, �������� �������� � ������������� ����������:';
    6:st:='��������� ��� ��������������� ������ �����������:';
    7:st:='������� ��������� ��� �������� �������� � ������������� ����������:';
    8:st:='��������� ��� �������� � ����������������� ������������ �������� ������������:';
  end;
  if Pos('���',Aud.Auditoria)<>0 then
    st:=st+chr(10)+'�.������, �� ���������, 3, 121552'
  else
    st:=st+chr(10)+'125993, �.������, ������������� �����, �.4';
  if ((TypeVivod=1) or(TypeVivod=3)) and (Copy(Aud.Auditoria,5,length(Aud.Auditoria)-4)<>'') then
    st:=st+', ������ - '+Aud.Korpus{+' �'+Copy(Copy(Aud.Auditoria,5,length(Aud.Auditoria)-4),1,Pos('(',Copy(Aud.Auditoria,5,length(Aud.Auditoria)-4))-1)};
  Excel.Cells[NomRow,3]:=st;
  case typezan of
    1:st:='������� ������: ����� � ������ ��� �����������; ���� � ���� ��� �������������; ����� � ����� (��������)';
    2:if Aud.ProektorAuditoria<>'' then st:='������������������ ������������ ������������ �������.'{ else st:='������������������ ������������ ������������ ������� ��� ���������� ������������ ����� �� ����� "'+NameDisLr+'"'};
    3:st:='������� ������: ����� � ������ ��� �����������; ���� � ���� ��� �������������; ����� � ����� (��������)';
    4:st:='������� ������: ����� � ������ ��� �����������; ���������� � �������� � ���� Internet';
    5:If Pos('������',NameDis)<>0 then st:='������� ������: ����� � ������ ��� �����������; ������������������ ������ � ����������� �������� ��������, �������� ��� ������������� ������� ���������� ������� ���������'+' (��������������� �������: ��������, �����, ���������/�������). ���������� ��� ������ ���������.' else st:='������� ������: ����� � ������ ��� �����������;';
    6,7:st:='������� ������: ����� � ������ ��� �����������; ���������� � �������� � ���� Internet � �������� ������� � ����������� �������������-��������������� �����.';
    8:st:='����� ��� ������ �������������� ���������';
  end;

  if (typezan<=3) and (Aud.ProektorAuditoria<>'') then
    st:=st+', '+Aud.ProektorAuditoria
  else
  if (typezan<=3) and ((pos('�������',TrebAudProektor)<>0) or
  (pos('������',TrebAudProektor)<>0) or (pos('������',TrebAudProektor)<>0) or (pos('��������',TrebAudProektor)<>0)  or
     (pos('���������',TrebAudProektor)<>0) or (pos('�����',TrebAudProektor)<>0) ) then
    st:=st+', '+'������������������ ������ � ����������� �������� ��������, �������� ��� ������������� ������� ���������� ������� ��������� (��������������� �������: ��������, �����, ���������/�������).';
  AudProektorVivod:=st;
  Excel.Cells[NomRow,4]:=st;
//  if KolRows>5 then
  if Aud.OsnashenieOgrnAuditoria<>'' then
    Excel.Cells[NomRow,6]:=Aud.OsnashenieOgrnAuditoria
  else
    Excel.Cells[NomRow,6]:='�� �������������.';
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
    Excel.Cells[1,1]:='����������� ��������������� ��������� ��������������� ���������� ������� �����������';
    Excel.Range[Excel.Cells[2,1],Excel.Cells[2,KolRows]].MergeCells:=true;
    Excel.Cells[2,1]:='����������� ����������� �������� (������������ ����������������� �����������)�';
    Excel.Range[Excel.Cells[4,1],Excel.Cells[4,KolRows]].MergeCells:=true;
    Excel.Cells[4,1]:='�������';
    Excel.Range[Excel.Cells[5,1],Excel.Cells[5,KolRows]].MergeCells:=true;
    st:='� �����������-����������� ����������� �������� ��������������� ��������� ������� ����������� � ���������';
    If Pos('05',ArrProfil[NomProfil].Profil)<>0 then Excel.Cells[5,1]:=st+' �����������';
    If Pos('�',ArrProfil[NomProfil].Profil)<>0 then Excel.Cells[5,1]:=st+' ������������';
    If Pos('�',ArrProfil[NomProfil].Profil)<>0 then Excel.Cells[5,1]:=st+' ������������';
    If Pos('�',ArrProfil[NomProfil].Profil)<>0 then Excel.Cells[5,1]:=st+' ������������';
    Excel.Range[Excel.Cells[6,1],Excel.Cells[6,KolRows]].MergeCells:=true;
    Excel.Cells[6,1]:=ArrProfil[NomProfil].Naprav+' '+ArrProfil[NomProfil].NameNaprav+' - '+ArrProfil[NomProfil].NameProfil;
    Excel.Cells[8,1]:='� �\�';
    Excel.Cells[8,2]:='������������ ���������� (������), ������� � ������������ � ��';
    Excel.Cells[8,3]:='������������ �����������* ��������� � ��������� ��� ��������������� ������';
    Excel.Cells[8,4]:='������������ ����������� ��������� � ��������� ��� ��������������� ������';
    Excel.Cells[8,5]:='�������� ������������� ������������ �����������. ��������� ��������������� ��������� ';
    Excel.Cells[8,6]:='����������������� ��������� ��� ������������� ���������� � ������ � ������������� ������������� ��������';
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
                  Excel.Cells[NomRow,9]:='�';
                  end;
                if ((ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].LKAud<>nil) and
                   ((Pos('������',AudProektorVivod)<>0) or
                   (Pos('��������',AudProektorVivod)<>0))) or
                   (Pos('������',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('������',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('�����������',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('��������',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0)  then
                begin
                  st:='Microsoft Windows �������� �007-1-0834-18 �� 31.05.2018�. Microsoft Office (������� WORD, EXCEL � �.�.) �������� �007-1-0834-18 �� 31.05.2018�.';
                  Excel.Cells[NomRow,5]:=st+' Kaspersky Endpoint Security �������� �070-1-0478-17 �� 18.04.2017�. Kaspersky Endpoint Security �������� �2075/M21 �� 22.01.2015�.'{+ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PO};
                  end
                else
                begin
                st:=Excel.Cells[NomRow,5];
                if st='' then
                  Excel.Cells[NomRow,5]:='�� ���������.';
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
                   ((Pos('������',AudProektorVivod)<>0) or
                   (Pos('��������',AudProektorVivod)<>0))) or
                   (Pos('������',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('������',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('�����������',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('��������',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0)then
                  begin
                  st:='Microsoft Windows �������� �007-1-0834-18 �� 31.05.2018�. Microsoft Office (������� WORD, EXCEL � �.�.) �������� �007-1-0834-18 �� 31.05.2018�.';
                  Excel.Cells[NomRow+1,5]:=st+' Kaspersky Endpoint Security �������� �070-1-0478-17 �� 18.04.2017�. Kaspersky Endpoint Security �������� �2075/M21 �� 22.01.2015�.'+ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PO;
                  end
                else
                begin
                st:=Excel.Cells[NomRow+1,5];
                if st='' then
                  Excel.Cells[NomRow+1,5]:='�� ���������.';
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
                   ((Pos('������',AudProektorVivod)<>0) or
                   (Pos('��������',AudProektorVivod)<>0))) or
                   (Pos('������',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('������',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('�����������',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) or
                   (Pos('��������',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].OsnashenieRPD)<>0) then
                  begin
                  st:='Microsoft Windows �������� �007-1-0834-18 �� 31.05.2018�. Microsoft Office (������� WORD, EXCEL � �.�.) �������� �007-1-0834-18 �� 31.05.2018�.';
                  Excel.Cells[NomRow+2,5]:=st+' Kaspersky Endpoint Security �������� �070-1-0478-17 �� 18.04.2017�. Kaspersky Endpoint Security �������� �2075/M21 �� 22.01.2015�.'+ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PO;
                  end
                else
                  begin
                st:=Excel.Cells[NomRow+2,5];
                if st='' then
                  Excel.Cells[NomRow+2,5]:='�� ���������.';
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
                  st:='Microsoft Windows �������� �007-1-0834-18 �� 31.05.2018�. Microsoft Office (������� WORD, EXCEL � �.�.) �������� �007-1-0834-18 �� 31.05.2018�.';
                  Excel.Cells[NomRow+3,5]:=st+' Kaspersky Endpoint Security �������� �070-1-0478-17 �� 18.04.2017�. Kaspersky Endpoint Security �������� �2075/M21 �� 22.01.2015�. Google Chrome (��������� ��)'+ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PO;
                  end
                else
                begin
                st:=Excel.Cells[NomRow+3,5];
                if st='' then
                  Excel.Cells[NomRow+3,5]:='�� ���������.';
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
                If pos('�������',ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].Name)<>0 then
                  begin
                  st:='Microsoft Windows �������� �007-1-0834-18 �� 31.05.2018�. Microsoft Office (������� WORD, EXCEL � �.�.) �������� �007-1-0834-18 �� 31.05.2018�.';
                  Excel.Cells[NomRow+4,5]:=st+' Kaspersky Endpoint Security �������� �070-1-0478-17 �� 18.04.2017�. Kaspersky Endpoint Security �������� �2075/M21 �� 22.01.2015�. GoogleChrome. '+ArrProfil[NomProfil].SemYp[NomSemPlanToo].Disciplin[NomDisSemYPToo].PO;
                  end
                else
                begin
                st:=Excel.Cells[NomRow+4,5];
                if st='' then
                  Excel.Cells[NomRow+4,5]:='�� ���������.';
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
  Excel.Cells[NomRow,2]:='��������������� ������ �����������';
  i:=0;
  while i<Length(ArrAuditSRS) do
    begin
    VivodAuditoria(ArrAuditorii[ArrAuditSRS[i]],6,NomRow+i);
    st:='Microsoft Windows �������� �007-1-0834-18 �� 31.05.2018�. Microsoft Office (������� WORD, EXCEL � �.�.) �������� �007-1-0834-18 �� 31.05.2018�.';
    Excel.Cells[NomRow+i,5]:=st+' Kaspersky Endpoint Security �������� �070-1-0478-17 �� 18.04.2017�. Kaspersky Endpoint Security �������� �2075/M21 �� 22.01.2015�. Google Chrome';
    inc(i);
    end;
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow+i-1,1]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow+i-1,2]].MergeCells:=true;
  NomRow:=NomRow+i;
{  Excel.Cells[NomRow,1]:=Nom;
  inc(Nom);
  Excel.Cells[NomRow,2]:='������� �������� � ������������� ����������';
  i:=0;
  while i<Length(ArrAuditoriiKontrol) do
    begin
    VivodAuditoria(ArrAuditorii[ArrAuditoriiKontrol[i]],7,NomRow+i);
    st:='Microsoft Windows �������� �007-1-0834-18 �� 31.05.2018�. Microsoft Office (������� WORD, EXCEL � �.�.) �������� �007-1-0834-18 �� 31.05.2018�.';
    Excel.Cells[NomRow+i,5]:=st+' Kaspersky Endpoint Security �������� �070-1-0478-17 �� 18.04.2017�. Kaspersky Endpoint Security �������� �2075/M21 �� 22.01.2015�. Google Chrome';
    inc(i);
    end;
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow+i-1,1]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow+i-1,2]].MergeCells:=true;
  NomRow:=NomRow+i;      }
  Excel.Cells[NomRow,1]:=Nom;
  inc(Nom);
  Excel.Cells[NomRow,2]:='�������� � ���������������� ������������ �������� ������������';
  i:=0;
  while i<Length(ArrAuditoriiObslyz) do
    begin
    VivodAuditoria(ArrAuditorii[ArrAuditoriiObslyz[i]],8,NomRow+i);
    Excel.Cells[NomRow+i,5]:='������������ ������������ ��';
    inc(i);
    end;
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow+i-1,1]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow+i-1,2]].MergeCells:=true;
  NomRow:=NomRow+i;

  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,KolRows]].MergeCells:=true;
  st:='*����������� ��������� - ������� ��������� ��� ���������� ������� ����������� ����, ������� ������������ ����, ��������� �������������� (���������� �������� �����)';
  Excel.Cells[NomRow,1]:=st+', ��������� � �������������� ������������, �������� �������� � ������������� ����������, � ����� ��������� ��� ��������������� ������.';

  Excel.Range[Excel.Cells[8,1],Excel.Cells[NomRow,KolRows]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[8,1],Excel.Cells[NomRow,KolRows]].BorderAround(1);
  NomRow:=NomRow+2;
  {
  NomStartTabled:=NomRow;
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:='�������� ��������� ��� (�� ������, ��������������� ����� ��������� ����������� �� ���)';
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:=1;
  Excel.Cells[NomRow,2]:=2;
  Excel.Cells[NomRow,5]:=3;
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:='������� ���';
  Excel.Cells[NomRow,2]:='������������ ��������� � ��������� ����������';
  Excel.Cells[NomRow,5]:='���� �������� ���������';
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow,6]].HorizontalAlignment:=xlCenter;
  inc(NomRow);
//  Excel.Cells[NomRow,1]:='2018 / 2019';
  if FileExists(CurrentDir+'/���/�������� ���.txt') then
    begin
    AssignFile(f,CurrentDir+'/���/�������� ���.txt');
    reset(f);
    while not EOF(f) do
      begin

      readln(f,st);
      if {(CurrentSemestr=1) and (Pos('�',ArrProfil[NomProfil].Profil)<>0) and (StrToInt('20'+Copy(st,1,2))>=CurrentYear-YearBakalavr) or
         (CurrentSemestr=2) and (Pos('�',ArrProfil[NomProfil].Profil)<>0) and (StrToInt('20'+Copy(st,4,2))>=CurrentYear-YearBakalavr) or
         (CurrentSemestr=1) and (Pos('�',ArrProfil[NomProfil].Profil)<>0) and (StrToInt('20'+Copy(st,1,2))>=CurrentYear-YearMagistr) or
         (CurrentSemestr=2) and (Pos('�',ArrProfil[NomProfil].Profil)<>0) and (StrToInt('20'+Copy(st,4,2))>=CurrentYear-YearMagistr) {or
         (CurrentSemestr=1) and (Pos('�',ArrProfil[NomProfil].Profil)<>0) and (StrToInt('20'+Copy(st,1,2))>=CurrentYear-YearAspirant) or
         (CurrentSemestr=2) and (Pos('�',ArrProfil[NomProfil].Profil)<>0) and (StrToInt('20'+Copy(st,4,2))>=CurrentYear-YearAspirant) }{ true then
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
  Excel.Cells[NomRow,1]:='������������ ���������';
  Excel.Cells[NomRow,5]:='������������ ��������� (� ���������, ���� ����������, �����������, �������� ��������, ���� ������, ���� ��������)';
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:=1;
  Excel.Cells[NomRow,5]:=2;
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow,6]].HorizontalAlignment:=xlCenter;
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  st:=' � ������������ ������, ��������, ���������� � ���������, ������������ ��� ������� ��������������� ������������, ������������� ����������������� �� �����������.';
  Excel.Cells[NomRow,1]:='����������, �������� � ������������� ������� ��������, ��������������� ��������������� �������� ������,'+st;
  Excel.Cells[NomRow,5]:='���������� �8-28-5-26 � ������������ (��������������) ������� ������ ����������� �������� ������������ �� 03 ����� 2017 ����';
  inc(NomRow);
                                                            }
  {Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow-1,6]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow-1,6]].BorderAround(1); }
{  NomRow:=NomRow+3;
  Excel.Cells[NomRow,1]:='������������ �����������, ';
  Excel.Cells[NomRow+1,1]:='�������������� ��������������� ������������                         ________________________ /______________________ /';
  Excel.Cells[NomRow+2,1]:='                                                                                                                                       �������                          �.�.�. ���������';
  Excel.Cells[NomRow+3,1]:='�.�.';
  Excel.Cells[NomRow+4,1]:='���� ����������� ________________'; }
//  Excel.Cells[NomRow+1,1]:='';

  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,KolRows]].Font.Size:=8;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,KolRows]].WrapText:=true;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,KolRows]].VerticalAlignment:=xlCenter;
//  Excel.Range[Excel.Cells[1,2],Excel.Cells[NomRow,5]].HorizontalAlignment:=xlCenter;
  if not DirectoryExists(CurrentDir+'\'+NameFolder) then
    ForceDirectories(CurrentDir+'\'+NameFolder);
  if (ArrProfil[NomProfil].Profil<>'') then
    Excel.Workbooks[1].saveas(CurrentDir+'\'+NameFolder+'\'+ArrProfil[NomProfil].Profil+'.xlsx');
  Fmain.MeProtocol.Lines.Add('������ ���� '+CurrentDir+'\'+NameFolder+'\'+ArrProfil[NomProfil].Profil+'.xlsx');
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
          st:='������� ���������:                          �.������, ������������� �����, 4, �-80, ���-3, 125993, �'+Copy(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria,5,length(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria)-4);
        {  if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[1]<>0 then
            st:=st+'                   ���������� ���� ��� ���������� ������� - '+IntTostr(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[1]);
          if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[2]<>0 then
            st:=st+'                   ���������� ���� ��� ������������ ������� - '+IntTostr(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[2]);
          if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[3]<>0 then
            st:=st+'                   ���������� ���� ��� ������������ ����� - '+IntTostr(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[3]); }
          Excel.Cells[NomRow,3]:=st;
          if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria='' then
            begin
            Excel.Cells[NomRow,6]:='1';
            if VivodProtocol then
              FMain.MeProtocol.Lines.Add('��� ��� '+Prepod[NomPrepod].FIO+' '+Prepod[NomPrepod].Nagryzka[NomNagr].Dis+' '+Prepod[NomPrepod].Nagryzka[NomNagr].Vid+' '+' '+IntToStr(NomNagr));
//            Excel.Cells[NomRow,3].Pattern:= 1;
//            Excel.Cells[NomRow,3].PatternColorIndex:= -4105;
//            Excel.Cells[NomRow,3].Color:= 65535;
            end;
          if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') then  st:='������� ��������� ��� ���������� ������� ����������� ���� ������� ������: ����� � ������ ��� �����������; ���� � ���� ��� �������������; ����� � ����� (��������).'
          else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') then  st:='�����������. ������������������ ������������ ������������ �������.'
          else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') then  st:='������� ��������� ��� ���������� ������� ������������ ���� ������� ������: ����� � ������ ��� �����������; ���� � ���� ��� �������������; ����� � ����� (��������).'
          else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') then  st:='������� ��������� ��� ���������� ������� ������������ ���� ������� ������: ����� � ������ ��� �����������; ���� � ���� ��� �������������; ����� � ����� (��������).'
          else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') then  st:='������� ��������� ��� ���������� ������� ������������ ���� ������� ������: ����� � ������ ��� �����������; ���� � ���� ��� �������������; ����� � ����� (��������).';

         { MaxStudentAuditoria:=0;
          for i := 1 to 3 do
            if MaxStudentAuditoria<Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[i] then
              MaxStudentAuditoria:=Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KolStudentAuditoriaMax[i];
          st:=st+'� ��������� �'+Copy(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria,5,length(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.Auditoria)-4)+': ����� ����������, ����� �� '+IntToStr(MaxStudentAuditoria)+' ��������, ���� �������������, ������';
          if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KomputersAuditoria<>0 then
            st:=st+', '+intToStr(Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KomputersAuditoria)+' ����������� ��� ������'; }
          st:=st+', '+Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.ProektorAuditoria;
          Excel.Cells[NomRow,4]:=st;
          if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.KomputersAuditoria<>0 then
            begin
            st:='Microsoft Windows �������� �007-1-0834-18 �� 31.05.2018�. Microsoft Office (������� WORD, EXCEL � �.�.) �������� �007-1-0834-18 �� 31.05.2018�.';

            Excel.Cells[NomRow,5]:=st+' Kaspersky Endpoint Security �������� �070-1-0478-17 �� 18.04.2017�. Kaspersky Endpoint Security �������� �2075/M21 �� 22.01.2015�.'+Prepod[NomPrepod].Nagryzka[NomNagr].PO;
            end
          else
            Excel.Cells[NomRow,5]:='�� ���������.';

          if Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.OsnashenieOgrnAuditoria<>'' then
            begin
            Excel.Cells[NomRow,6]:=Prepod[NomPrepod].Nagryzka[NomNagr].Auditoria.OsnashenieOgrnAuditoria;
            end
          else
            Excel.Cells[NomRow,6]:='�� �������';
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
  Excel.Cells[1,1]:='����������� ��������������� ��������� ��������������� ���������� ������� �����������';
  Excel.Range[Excel.Cells[2,1],Excel.Cells[2,6]].MergeCells:=true;
  Excel.Cells[2,1]:='����������� ����������� �������� (������������ ����������������� �����������)�';
  Excel.Range[Excel.Cells[4,1],Excel.Cells[4,6]].MergeCells:=true;
  Excel.Cells[4,1]:='�������';
  Excel.Range[Excel.Cells[5,1],Excel.Cells[5,6]].MergeCells:=true;
  st:='� �����������-����������� ����������� �������� ��������������� ��������� ������� ����������� � ���������';
  case NameAllGroup[NomTypeGroup].NameGroupKyrs[3] of
    '�' :Excel.Cells[5,1]:=st+' ������������';
    '�' :Excel.Cells[5,1]:=st+' ������������';
    '�' :Excel.Cells[5,1]:=st+' �����������';
    '�' :Excel.Cells[5,1]:=st+' ������������';
  else
    Excel.Cells[5,1]:=st;
  end;
  Excel.Range[Excel.Cells[6,1],Excel.Cells[6,6]].MergeCells:=true;
  if NameAllGroup[NomTypeGroup].NameGroupKyrs='07�' then
    Excel.Cells[6,1]:='09.03.01 "����������� � �������������� �������" ������� - ������������������ ������� ��������� ���������� � ����������'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='09�' then
    Excel.Cells[6,1]:='09.03.01 "����������� � �������������� �������" ������� - �������������� ������, ��������, ������� � ����'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='10�' then
    Excel.Cells[6,1]:='09.03.01 "����������� � �������������� �������" ������� - ����������� ����������� ������� �������������� ������� � ������������������ ������'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='11�' then
    Excel.Cells[6,1]:='09.03.04 "����������� ���������" ������� - ����������-�������������� �������'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='05�' then
    Excel.Cells[6,1]:='09.04.01 "����������� � �������������� �������" ������� - ������������������ ������� ��������� ���������� � ����������'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='06�' then
    Excel.Cells[6,1]:='09.04.01 "����������� � �������������� �������" ������� - �������������� ������, ��������, ������� � ����'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='07�' then
    Excel.Cells[6,1]:='09.04.01 "����������� � �������������� �������" ������� - ����������� ����������� ������� �������������� ������� � ������������������ ������'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='08�' then
    Excel.Cells[6,1]:='09.04.04 "����������� ���������" ������� - ����������-�������������� �������'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='03�' then
    Excel.Cells[6,1]:='09.06.01 (05.13.15 �������������� ������, ��������� � ������������ ����)'
  else if NameAllGroup[NomTypeGroup].NameGroupKyrs='05�' then
    Excel.Cells[6,1]:='09.06.01 (05.13.11 �������������� � ����������� ����������� �������������� �����, ���������� � ������������ �����)' ;
  Excel.Cells[8,1]:='� �\�';
  Excel.Cells[8,2]:='������������ ���������� (������), ������� � ������������ � ������� ������';
  Excel.Cells[8,3]:='������������ �����������* ��������� � ��������� ��� ��������������� ������';
  Excel.Cells[8,4]:='������������ ����������� ��������� � ��������� ��� ��������������� ������';
  Excel.Cells[8,5]:='�������� ������������� ������������ �����������. ��������� ��������������� ��������� ';
  Excel.Cells[8,6]:='����������������� ��������� ��� ������������� ���������� � ������ � ������������� ������������� ��������';
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
           ((Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') or (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��')) then
          begin
          Excel.Cells[NomRow,1]:=Nom;
          Excel.Cells[NomRow,2]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
          VivodInfo(NomRow,NomPrepod,NomNagr);
          inc(NomRow);
          inc(Nom);

          for i := 1 to 3 do
            ArrNagr[i]:=0;
          if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') then  ArrNagr[1]:=1;
          if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') then  ArrNagr[2]:=1;
          if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') then  ArrNagr[3]:=1;
          Prepod[NomPrepod].Nagryzka[NomNagr].FlagVivoda:=true;

          //����� ��� ���� ������������ �� ������� ��������
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
                 (((Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Vid='��') and (ArrNagr[1]=0)) or
                  ((Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Vid='��') and (ArrNagr[2]=0)) or
                  ((Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Vid='��') and (ArrNagr[3]=0))) then
                begin
                if (Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Vid='��') then  ArrNagr[1]:=1;
                if (Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Vid='��') then  ArrNagr[2]:=1;
                if (Prepod[NomPrepodToo].Nagryzka[NomNagrToo].Vid='��') then  ArrNagr[3]:=1;

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
   Excel.Cells[NomRow,2]:='��������� ��� ��������������� ������ ��������';
   Excel.Cells[NomRow,3]:='������� ���������:                          �.������, �� ���������, 3, � �-707,          ���������� ���� ��� ��������������� ������ ��������� - 26';
   Excel.Cells[NomRow,4]:='��������������� ������.                � ��������� �-707: ����� ����������, ����� �� 20 ��������, ���� �������������, ������, 11 ����������� ��� ������';
   st:='Microsoft Windows �������� �007-1-0834-18 �� 31.05.2018�. Microsoft Office (������� WORD, EXCEL � �.�.) �������� �007-1-0834-18 �� 31.05.2018�. Kaspersky Endpoint Security �������� �070-1-0478-17 �� 18.04.2017�. ';
   Excel.Cells[NomRow,5]:=st+'Kaspersky Endpoint Security �������� �2075/M21 �� 22.01.2015�. COMP-P(���������� ��, FreeWare) Microsoft Visual Studio(���������� ��);MinGW(���������� ��,  GNU General Public License);';
   Excel.Cells[NomRow+1,3]:='������� ���������:                          �.������, �� ���������, 3, � �-711,          ���������� ���� ��� ��������������� ������ ��������� - 26';
   Excel.Cells[NomRow+1,4]:='��������������� ������.                � ��������� �-711: ����� ����������, ����� �� 20 ��������, ���� �������������, ������, 11 ����������� ��� ������';
   st:='Microsoft Windows �������� �007-1-0834-18 �� 31.05.2018�. Microsoft Office (������� WORD, EXCEL � �.�.) �������� �007-1-0834-18 �� 31.05.2018�. Kaspersky Endpoint Security �������� �070-1-0478-17 �� 18.04.2017�. ';
   Excel.Cells[NomRow+1,5]:=st+'Kaspersky Endpoint Security �������� �2075/M21 �� 22.01.2015�. COMP-P(���������� ��, FreeWare) Microsoft Visual Studio(���������� ��);MinGW(���������� ��,  GNU General Public License);';
   Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow+1,1]].MergeCells:=true;
   Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow+1,2]].MergeCells:=true;
   Excel.Cells[NomRow+2,1]:=Nom+1;
   Excel.Cells[NomRow+2,2]:='��������� ��� ������������ ������������ � ���. �������';
   Excel.Cells[NomRow+2,3]:='���������:                                       �.������, ������������� �����, 4, �-80, ���-3, 125993, �434�(3)';
   Excel.Cells[NomRow+2,4]:='� ��������� �434�(3) ����� �������� 5 �����������. ��������� ������������� ��� ���';
   Excel.Cells[NomRow+2,5]:='�������� ���������� � ��������� ���������� ��� ������� �������������� � ���. �������';
   Excel.Cells[NomRow+3,3]:='� ������������ ��� �������� �������';
   Excel.Cells[NomRow+3,4]:='';
   Excel.Cells[NomRow+3,5]:='�������� ���������� � ��������� ���������� ��� ������� �������������� � ���. �������';
   Excel.Range[Excel.Cells[NomRow+2,1],Excel.Cells[NomRow+3,1]].MergeCells:=true;
   Excel.Range[Excel.Cells[NomRow+2,2],Excel.Cells[NomRow+3,2]].MergeCells:=true;
   Excel.Cells[NomRow+4,1]:=Nom+2;
   Excel.Cells[NomRow+4,2]:='��������� ��� ������������ ������';
   Excel.Cells[NomRow+4,3]:='���������:                                           �.������, ������������� �����, 4, �-80, ���-3, 125993, �217(3)';
   Excel.Cells[NomRow+4,4]:='� ��������� �217(3) ����� ������������ �������� 3 �����������. ���������: 3 ����������, �����, ������, �������, �������';
   st:='Microsoft Windows �������� �007-1-0834-18 �� 31.05.2018�. Microsoft Office (������� WORD, EXCEL � �.�.) �������� �007-1-0834-18 �� 31.05.2018�. Kaspersky Endpoint Security �������� �070-1-0478-17 �� 18.04.2017�. ';
   Excel.Cells[NomRow+4,5]:=st+'Kaspersky Endpoint Security �������� �2075/M21 �� 22.01.2015�.';

   for I := 0 to 4 do
     Excel.Cells[NomRow+i,6]:='�� �������';

   NomRow:=NomRow+5;

  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,6]].MergeCells:=true;
  st:='*����������� ��������� - ������� ��������� ��� ���������� ������� ����������� ����, ������� ������������ ����, ��������� �������������� (���������� �������� �����)';
  Excel.Cells[NomRow,1]:=st+', ��������� � �������������� ������������, �������� �������� � ������������� ����������, � ����� ��������� ��� ��������������� ������.';

  Excel.Range[Excel.Cells[8,1],Excel.Cells[NomRow,6]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[8,1],Excel.Cells[NomRow,6]].BorderAround(1);
  NomRow:=NomRow+2;

  NomStartTabled:=NomRow;
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:='�������� ��������� ��� (�� ������, ��������������� ����� ��������� ����������� �� ���)';
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:=1;
  Excel.Cells[NomRow,2]:=2;
  Excel.Cells[NomRow,5]:=3;
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,2],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:='������� ���';
  Excel.Cells[NomRow,2]:='������������ ��������� � ��������� ����������';
  Excel.Cells[NomRow,5]:='���� �������� ���������';
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow,6]].HorizontalAlignment:=xlCenter;
  inc(NomRow);
//  Excel.Cells[NomRow,1]:='2018 / 2019';
  if FileExists(CurrentDir+'/���/�������� ���.txt') then
    begin
    AssignFile(f,CurrentDir+'/���/�������� ���.txt');
    reset(f);
    while not EOF(f) do
      begin

      readln(f,st);
      if (CurrentSemestr=1) and (NameAllGroup[NomTypeGroup].NameGroupKyrs[3]='�') and (StrToInt('20'+Copy(st,1,2))>=CurrentYear-YearBakalavr) or
         (CurrentSemestr=2) and (NameAllGroup[NomTypeGroup].NameGroupKyrs[3]='�') and (StrToInt('20'+Copy(st,4,2))>=CurrentYear-YearBakalavr) or
         (CurrentSemestr=1) and (NameAllGroup[NomTypeGroup].NameGroupKyrs[3]='�') and (StrToInt('20'+Copy(st,1,2))>=CurrentYear-YearMagistr) or
         (CurrentSemestr=2) and (NameAllGroup[NomTypeGroup].NameGroupKyrs[3]='�') and (StrToInt('20'+Copy(st,4,2))>=CurrentYear-YearMagistr) or
         (CurrentSemestr=1) and (NameAllGroup[NomTypeGroup].NameGroupKyrs[3]='�') and (StrToInt('20'+Copy(st,1,2))>=CurrentYear-YearAspirant) or
         (CurrentSemestr=2) and (NameAllGroup[NomTypeGroup].NameGroupKyrs[3]='�') and (StrToInt('20'+Copy(st,4,2))>=CurrentYear-YearAspirant) then
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
  Excel.Cells[NomRow,1]:='������������ ���������';
  Excel.Cells[NomRow,5]:='������������ ��������� (� ���������, ���� ����������, �����������, �������� ��������, ���� ������, ���� ��������)';
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  Excel.Cells[NomRow,1]:=1;
  Excel.Cells[NomRow,5]:=2;
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow,6]].HorizontalAlignment:=xlCenter;
  inc(NomRow);
  Excel.Range[Excel.Cells[NomRow,1],Excel.Cells[NomRow,4]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomRow,5],Excel.Cells[NomRow,6]].MergeCells:=true;
  st:=' � ������������ ������, ��������, ���������� � ���������, ������������ ��� ������� ��������������� ������������, ������������� ����������������� �� �����������.';
  Excel.Cells[NomRow,1]:='����������, �������� � ������������� ������� ��������, ��������������� ��������������� �������� ������,'+st;
  Excel.Cells[NomRow,5]:='���������� �8-28-5-26 � ������������ (��������������) ������� ������ ����������� �������� ������������ �� 03 ����� 2017 ����';
  inc(NomRow);

  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow-1,6]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomStartTabled,1],Excel.Cells[NomRow-1,6]].BorderAround(1);
  NomRow:=NomRow+3;
  Excel.Cells[NomRow,1]:='������������ �����������, ';
  Excel.Cells[NomRow+1,1]:='�������������� ��������������� ������������                         ________________________ /_������ ���� ����������_ /';
  Excel.Cells[NomRow+2,1]:='                                                                                                                                       �������                          �.�.�. ���������';
  Excel.Cells[NomRow+3,1]:='�.�.';
  Excel.Cells[NomRow+4,1]:='���� ����������� ________________';
//  Excel.Cells[NomRow+1,1]:='';

  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,6]].Font.Size:=8;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,6]].WrapText:=true;
  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomRow-1,6]].VerticalAlignment:=xlCenter;
//  Excel.Range[Excel.Cells[1,2],Excel.Cells[NomRow,5]].HorizontalAlignment:=xlCenter;
  if not DirectoryExists(CurrentDir+'\���') then
    ForceDirectories(CurrentDir+'\���');
  if (NomTypeGroup<length(NameAllGroup)) and (NameAllGroup[NomTypeGroup].NameGroupKyrs<>'') then
    Excel.Workbooks[1].saveas(CurrentDir+'\���\'+NameAllGroup[NomTypeGroup].NameGroupKyrs+'.xlsx');
  Fmain.MeProtocol.Lines.Add('������ ���� '+CurrentDir+'\���\'+NameAllGroup[NomTypeGroup].NameGroupKyrs+'.xlsx');
  Excel.Workbooks.Close;
  end;
  inc(NomTypeGroup);
  end;
end;


end.
