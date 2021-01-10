unit UIndividualPlanPrepod;

interface

uses SysUtils, StdCtrls, Windows, Dialogs, UConstParametrs, UsaveExcel;

Procedure CreateIndividualPlanAllPrepod(MeProtocol:TMemo);

implementation

uses UMain,UNagryzka;

Procedure CreateIndividualPlanAllPrepod(MeProtocol:TMemo);
var
NomPrepod,NomNagr,NomRow,NomYchMetodWork,AddYchMetodWork:Longword;
SumNagryzkaPrepod:Double;
NameOtch,NameOtch1,OsnDis:String;
PosPapki:LongWord;
SdvigHour:Double;
st,st1:string;

procedure AddInfoPrepod(StFileName:string; OgrNagryzkiSem1,OgrNagryzkiSem2:Longword; TypeNagryzki:byte);
var
  CurrentHourSem1,CurrentHourSem2:Double;
  st:string;
  NomGroupNagryzka:Longword;
begin
  Excel.Workbooks.Open(StFileName);

  Excel.Cells[12,3]:=Prepod[NomPrepod].FIO;
  case TypeNagryzki of
    1:begin
      If Prepod[NomPrepod].Stavka='' then
        begin
        if Prepod[NomPrepod].StavkaSovmest='' then
          Excel.Cells[13,2]:=Prepod[NomPrepod].Dolzhnost+' - ��������� ������'
        else
          Excel.Cells[13,2]:=Prepod[NomPrepod].Dolzhnost+' ('+Prepod[NomPrepod].StavkaSovmest+' ��. ����.)';
        end
      else if Prepod[NomPrepod].StavkaSovmest='' then
        Excel.Cells[13,2]:=Prepod[NomPrepod].Dolzhnost+' ('+Prepod[NomPrepod].Stavka+' ��.)'
      else
        Excel.Cells[13,2]:=Prepod[NomPrepod].Dolzhnost+' ('+Prepod[NomPrepod].Stavka+' ��.  +'+Prepod[NomPrepod].StavkaSovmest+' ��. ����.)';
      end;
    2:If Prepod[NomPrepod].Stavka<>'' then Excel.Cells[13,2]:=Prepod[NomPrepod].Dolzhnost+' ('+Prepod[NomPrepod].Stavka+' ��.)';
    3:If Prepod[NomPrepod].StavkaSovmest<>'' then Excel.Cells[13,2]:=Prepod[NomPrepod].Dolzhnost+' ('+Prepod[NomPrepod].StavkaSovmest+' ��.) - ��. ����.';
    4:Excel.Cells[13,2]:=Prepod[NomPrepod].Dolzhnost+' - ��������� ������';
  end;

  Excel.Cells[13,5]:=Prepod[NomPrepod].Stepen;
  Excel.Cells[13,8]:=Prepod[NomPrepod].Zvanie;
  //������ �� ���� ����������� �������������
  SumNagryzkaPrepod:=0;
  CurrentHourSem1:=0; CurrentHourSem2:=0;
  NomNagr:=0;
  NomRow:=0;
  while (NomNagr<Length(Prepod[NomPrepod].Nagryzka)) do
    begin
    if (Prepod[NomPrepod].Nagryzka[NomNagr].FlagIndPlan=0) and
       (((Prepod[NomPrepod].Nagryzka[NomNagr].sem=1) and (CurrentHourSem1+StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagr].Hour)<=OgrNagryzkiSem1)) or
       ((Prepod[NomPrepod].Nagryzka[NomNagr].sem=2) and (CurrentHourSem2+StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagr].Hour)<=OgrNagryzkiSem2))) then
    begin
    //������� ���������� � ����
    Prepod[NomPrepod].Nagryzka[NomNagr].FlagIndPlan:=TypeNagryzki;
    st:='';
    NomGroupNagryzka:=0;
    while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
      begin
      st:=st+Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka].Nom+', ';
      inc(NomGroupNagryzka);
      end;
    Excel.Cells[21+NomRow,1]:=st;
    Excel.Cells[21+NomRow,2]:=Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
    if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') or
       (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') or
       (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') or
       (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') or
       (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') then
      Excel.Cells[21+NomRow,3]:=Prepod[NomPrepod].Nagryzka[NomNagr].Vid
    else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='������������') then
      Excel.Cells[21+NomRow,3]:='����'
    else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='�������') then
      Excel.Cells[21+NomRow,3]:='�'
    else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='����� � �������') then
      Excel.Cells[21+NomRow,3]:='��'
    else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='�����') then
      Excel.Cells[21+NomRow,3]:='��'
    else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��������') then
      Excel.Cells[21+NomRow,3]:='����'
    else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='����������� ����������') then
      Excel.Cells[21+NomRow,3]:='����'
    else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='����������� �����������') then
      Excel.Cells[21+NomRow,3]:='���'
    else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='������������� ��������') then
      Excel.Cells[21+NomRow,3]:='����'
    else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='������') then
      Excel.Cells[21+NomRow,3]:='���'
    else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='����������� ��������') then
      Excel.Cells[21+NomRow,3]:='����';

    Excel.Cells[21+NomRow,8]:=Prepod[NomPrepod].Nagryzka[NomNagr].Hour;
    if Prepod[NomPrepod].Nagryzka[NomNagr].sem=1 then
      begin
      Excel.Cells[21+NomRow,4]:=Prepod[NomPrepod].Nagryzka[NomNagr].Hour;
      CurrentHourSem1:=CurrentHourSem1+StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagr].Hour);
      end
    else
      begin
      Excel.Cells[21+NomRow,6]:=Prepod[NomPrepod].Nagryzka[NomNagr].Hour;
      CurrentHourSem2:=CurrentHourSem2+StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagr].Hour);
      end;
    inc(NomRow);
    end;
    inc(NomNagr);
    end;
  Excel.Range[Excel.Cells[21,1],Excel.Cells[21+NomRow-1,9]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[21,1],Excel.Cells[21+NomRow-1,9]].BorderAround(1);
  Excel.Range[Excel.Cells[21,1],Excel.Cells[21+NomRow-1,9]].WrapText:=true;
  Excel.Range[Excel.Cells[21,1],Excel.Cells[21+NomRow-1,9]].Font.Size:=10;
  Excel.Range[Excel.Cells[21,1],Excel.Cells[21+NomRow-1,9]].VerticalAlignment:=xlCenter; //xlCenter

  Excel.Cells[21+NomRow,1].HorizontalAlignment:=xlRight; //
  Excel.Cells[21+NomRow,1]:='�����';
  Excel.Range[Excel.Cells[21+NomRow,1],Excel.Cells[21+NomRow,3]].MergeCells:=true;
  Excel.Cells[21+NomRow,4]:=CurrentHourSem1;
  Excel.Cells[21+NomRow,6]:=CurrentHourSem2;
  Excel.Cells[21+NomRow,8]:=CurrentHourSem1+CurrentHourSem2;
  SumNagryzkaPrepod:=0;
  Excel.Range[Excel.Cells[21+NomRow,1],Excel.Cells[21+NomRow,9]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[21+NomRow,1],Excel.Cells[21+NomRow,9]].BorderAround(1);

  Excel.Cells[21+NomRow+1,1]:='����������.';
  Excel.Range[Excel.Cells[21+NomRow+2,1],Excel.Cells[21+NomRow+2,2]].MergeCells:=true;
  Excel.Cells[21+NomRow+2,1]:='��������� ��� �������� ������� ��������';
  Excel.Range[Excel.Cells[21+NomRow+2,3],Excel.Cells[21+NomRow+2,9]].MergeCells:=true;
  Excel.Range[Excel.Cells[21+NomRow+2,3],Excel.Cells[21+NomRow+2,9]].Borders[9].LineStyle:=1;
//  Excel.Cells[21+NomNagr+3,2].Underline:= -4142 //xlUnderlineStyleNone
  Excel.Range[Excel.Cells[21+NomRow+4,1],Excel.Cells[21+NomRow+4,9]].MergeCells:=true;
  Excel.Cells[21+NomRow+4,1].Font.Bold := True;
  Excel.Cells[21+NomRow+4,1].HorizontalAlignment:=xlCenter; //xlCenter
  Excel.Cells[21+NomRow+4,1].Font.Size:=14;
  Excel.Cells[21+NomRow+4,1]:='II. ������-������������ ������';
  Excel.Cells[21+NomRow+5,1]:='� �/�';
  Excel.Cells[21+NomRow+5,2]:='������������ ������';
  Excel.Cells[21+NomRow+5,3]:='������������';
  Excel.Cells[21+NomRow+5,4]:='�����, �.';
  Excel.Cells[21+NomRow+5,5]:='���� ����������';
  Excel.Range[Excel.Cells[21+NomRow+5,5],Excel.Cells[21+NomRow+5,7]].MergeCells:=true;
  Excel.Cells[21+NomRow+5,8]:='������� � ����������';
  Excel.Range[Excel.Cells[21+NomRow+5,8],Excel.Cells[21+NomRow+5,9]].MergeCells:=true;
  NomYchMetodWork:=21+NomRow+6;
  NomNagr:=0;
  AddYchMetodWork:=0;
  while NomNagr<Length(Prepod[NomPrepod].Nagryzka) do
    begin
    if Prepod[NomPrepod].Nagryzka[NomNagr].FlagIndPlan=TypeNagryzki then
    begin
    Excel.Cells[NomYchMetodWork+AddYchMetodWork,1]:=AddYchMetodWork+1;
    if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') then
      begin
      st:='';
      NomGroupNagryzka:=0;
      while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
        begin
        st:=st+Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka].Nom+', ';
        inc(NomGroupNagryzka);
        end;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,2]:='���������� � ������ ������. ���� - '+Prepod[NomPrepod].Nagryzka[NomNagr].Dis+' ������ '+st;
      if pos(Prepod[NomPrepod].Nagryzka[NomNagr].Dis,OsnDis)=0 then
        OsnDis:=OsnDis+Prepod[NomPrepod].Nagryzka[NomNagr].Dis+', ';
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,2].Font.Size:=8;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,3]:='�� 1';
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,4]:=1*StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagr].Hour);
      SumNagryzkaPrepod:=SumNagryzkaPrepod+1*StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagr].Hour);
      Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork,5],Excel.Cells[NomYchMetodWork+AddYchMetodWork,7]].MergeCells:=true;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,5]:='30 �������, 30 ������. ��������� ���� ������. ��������� ����� ����� ������ ������ � ������� ��������';
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,5].Font.Size:=8;
      Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork,8],Excel.Cells[NomYchMetodWork+AddYchMetodWork,9]].MergeCells:=true;
      inc(AddYchMetodWork);
      end
    else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') then
      begin
      st:='';
      NomGroupNagryzka:=0;
      while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
        begin
        st:=st+Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka].Nom+', ';
        inc(NomGroupNagryzka);
        end;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,2]:='���������� � ������������ ��������. ���� - '+Prepod[NomPrepod].Nagryzka[NomNagr].Dis+' ������ '+st;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,2].Font.Size:=8;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,3]:='�� 1';
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,4]:=1*StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagr].Hour);
      SumNagryzkaPrepod:=SumNagryzkaPrepod+1*StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagr].Hour);
      Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork,5],Excel.Cells[NomYchMetodWork+AddYchMetodWork,7]].MergeCells:=true;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,5]:='30 �������, 30 ������. ��������� ���� ������������ �������.';
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,5].Font.Size:=8;
      Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork,8],Excel.Cells[NomYchMetodWork+AddYchMetodWork,9]].MergeCells:=true;
      inc(AddYchMetodWork);
      end
    else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') then
      begin
      st:='';
      NomGroupNagryzka:=0;
      while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
        begin
        st:=st+Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka].Nom+', ';
        inc(NomGroupNagryzka);
        end;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,2]:='���������� � ������������ �������. ���� - '+Prepod[NomPrepod].Nagryzka[NomNagr].Dis+' ������ '+st;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,2].Font.Size:=8;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,3]:='�� 0.5';
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,4]:=0.5*StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagr].Hour);
      SumNagryzkaPrepod:=SumNagryzkaPrepod+0.5*StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagr].Hour);
      Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork,5],Excel.Cells[NomYchMetodWork+AddYchMetodWork,7]].MergeCells:=true;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,5]:='30 �������, 30 ������. ��������� ���� ������������� �����.';
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,5].Font.Size:=8;
      Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork,8],Excel.Cells[NomYchMetodWork+AddYchMetodWork,9]].MergeCells:=true;
      inc(AddYchMetodWork);
      end
    else if (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') or
            (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��') or
            (Prepod[NomPrepod].Nagryzka[NomNagr].Vid='��������') then
      begin
      st:='';
      NomGroupNagryzka:=0;
      while NomGroupNagryzka<Length(Prepod[NomPrepod].Nagryzka[NomNagr].Group) do
        begin
        st:=st+Prepod[NomPrepod].Nagryzka[NomNagr].Group[NomGroupNagryzka].Nom+', ';
        inc(NomGroupNagryzka);
        end;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,2]:='�������� ����� �� '+Prepod[NomPrepod].Nagryzka[NomNagr].Vid+'. ���� - '+Prepod[NomPrepod].Nagryzka[NomNagr].Dis+' ������ '+st;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,2].Font.Size:=8;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,3]:='�� 3';
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,4]:=3*12;{*StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagr].KolStudent);}
      SumNagryzkaPrepod:=SumNagryzkaPrepod+3*12;//3*Prepod[NomPrepod].Nagryzka[NomNagr].KolStudent;
      Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork,5],Excel.Cells[NomYchMetodWork+AddYchMetodWork,7]].MergeCells:=true;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,5]:='� ������� ��������';
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,5].Font.Size:=8;
      Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork,8],Excel.Cells[NomYchMetodWork+AddYchMetodWork,9]].MergeCells:=true;
      inc(AddYchMetodWork);

      Excel.Cells[NomYchMetodWork+AddYchMetodWork,1]:=AddYchMetodWork+1;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,2]:='������������ � ����� ����� �� '+Prepod[NomPrepod].Nagryzka[NomNagr].Vid+'. ���� - '+Prepod[NomPrepod].Nagryzka[NomNagr].Dis;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,2].Font.Size:=8;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,3]:='�� 0.4';
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,4]:=0.4*12;{*StrToFloat(Prepod[NomPrepod].Nagryzka[NomNagr].KolStudent);}
      SumNagryzkaPrepod:=SumNagryzkaPrepod+0.4*12;//0.4*Prepod[NomPrepod].Nagryzka[NomNagr].KolStudent;
      Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork,5],Excel.Cells[NomYchMetodWork+AddYchMetodWork,7]].MergeCells:=true;
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,5]:='� ������� ��������';
      Excel.Cells[NomYchMetodWork+AddYchMetodWork,5].Font.Size:=8;
      Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork,8],Excel.Cells[NomYchMetodWork+AddYchMetodWork,9]].MergeCells:=true;
      inc(AddYchMetodWork);
      end;
    end;
    inc(NomNagr);
    end;
  Excel.Range[Excel.Cells[NomYchMetodWork-1,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork,9]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomYchMetodWork-1,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork,9]].BorderAround(1);
  Excel.Range[Excel.Cells[NomYchMetodWork-1,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork,9]].WrapText:=true;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork,3]].MergeCells:=true;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork,1]:='�����';
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork,4],Excel.Cells[NomYchMetodWork+AddYchMetodWork,9]].MergeCells:=true;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork,4]:=SumNagryzkaPrepod;
  if TypeNagryzki=1 then
    Prepod[NomPrepod].AllHourPrav:=SumNagryzkaPrepod;

  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+1,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+1,9]].MergeCells:=true;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+1,1].Font.Bold := True;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+1,1].HorizontalAlignment:=-4108; //xlCenter
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+1,1].Font.Size:=14;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+1,1]:='III. ������ ���� �����';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+2,1]:='� �/�';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+2,2]:='������������ ������';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+2,3]:='������������';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+2,4]:='�����, �.';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+2,5]:='���� ����������';
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+2,5],Excel.Cells[NomYchMetodWork+AddYchMetodWork+2,7]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+3,5],Excel.Cells[NomYchMetodWork+AddYchMetodWork+3,7]].MergeCells:=true;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+2,8]:='������� � ����������';
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+2,8],Excel.Cells[NomYchMetodWork+AddYchMetodWork+2,9]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+3,8],Excel.Cells[NomYchMetodWork+AddYchMetodWork+3,9]].MergeCells:=true;
  if Prepod[NomPrepod].PovKval<>0 then
    begin
    Excel.Cells[NomYchMetodWork+AddYchMetodWork+3,1]:='1';
    Excel.Cells[NomYchMetodWork+AddYchMetodWork+3,2]:='��������� ������������ �������������';
    Excel.Cells[NomYchMetodWork+AddYchMetodWork+3,3]:='�� 250�.';
    Excel.Cells[NomYchMetodWork+AddYchMetodWork+3,4]:=Prepod[NomPrepod].PovKval;
    SumNagryzkaPrepod:=SumNagryzkaPrepod+Prepod[NomPrepod].PovKval;
    Excel.Cells[NomYchMetodWork+AddYchMetodWork+3,5]:='� ������� �������� ����';
    end;

  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+2,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+3,9]].Borders.Weight := 2;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+2,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+3,9]].BorderAround(1);
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+2,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+3,9]].WrapText:=true;

  Excel.Cells[14,2]:=OsnDis;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+5,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+5,9]].MergeCells:=true;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+5,1].Font.Bold := True;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+5,1].HorizontalAlignment:=-4108; //xlCenter
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+5,1].Font.Size:=14;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+5,1]:='IV. ����� � ����� ��������� ������������';
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+6,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+6,9]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+6,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+6,9]].Borders[9].LineStyle:=1;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+6,1]:=Prepod[NomPrepod].PovKvalProsh;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+7,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+7,2]].MergeCells:=true;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+7,1].Font.Bold := True;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+7,1]:='����� ����� �������� ������������� ����������:';
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+7,6],Excel.Cells[NomYchMetodWork+AddYchMetodWork+7,8]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+7,6],Excel.Cells[NomYchMetodWork+AddYchMetodWork+7,8]].Borders[9].LineStyle:=1;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+7,6]:=SumNagryzkaPrepod+CurrentHourSem1+CurrentHourSem2;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+7,9]:='�.';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+9,1].Font.Bold := True;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+9,1]:='�������������:';
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+9,2],Excel.Cells[NomYchMetodWork+AddYchMetodWork+9,2]].Borders[9].LineStyle:=1;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+9,6],Excel.Cells[NomYchMetodWork+AddYchMetodWork+9,7]].Borders[9].LineStyle:=1;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+9,5]:='___';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+9,6]:='��������';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+9,8]:='20__';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+9,9]:='����';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+10,2].Font.Size:=8;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+10,2]:='(�������)';
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+11,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+11,2]].MergeCells:=true;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+11,1]:='��������� �� ������ �������������:';
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+11,3],Excel.Cells[NomYchMetodWork+AddYchMetodWork+11,9]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+11,3],Excel.Cells[NomYchMetodWork+AddYchMetodWork+11,9]].Borders[9].LineStyle:=1;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+12,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+12,9]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+12,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+12,9]].Borders[9].LineStyle:=1;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+14,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+14,2]].MergeCells:=true;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+14,1]:='���������� ����������� ��������� �������:';
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+14,3],Excel.Cells[NomYchMetodWork+AddYchMetodWork+14,9]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+14,3],Excel.Cells[NomYchMetodWork+AddYchMetodWork+14,9]].Borders[9].LineStyle:=1;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+15,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+15,9]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+15,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+15,9]].Borders[9].LineStyle:=1;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+17,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+17,2]].MergeCells:=true;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+17,1]:='���������� � ���������� ����� �� ������� ���:';
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+17,3],Excel.Cells[NomYchMetodWork+AddYchMetodWork+17,9]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+17,3],Excel.Cells[NomYchMetodWork+AddYchMetodWork+17,9]].Borders[9].LineStyle:=1;
  st:='���������� �������� �'+NomKaf+': '+ZKafSokr;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+19,2]:=st;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+19,3],Excel.Cells[NomYchMetodWork+AddYchMetodWork+19,7]].Borders[9].LineStyle:=1;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+19,3],Excel.Cells[NomYchMetodWork+AddYchMetodWork+19,7]].MergeCells:=true;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+21,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+21,2]].MergeCells:=true;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+21,1]:='� ����������� � ���������� ����� ����������.';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+23,1].Font.Size:=10;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+23,1]:='�������������:';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+23,2]:=Prepod[NomPrepod].FIO;
  Excel.Range[Excel.Cells[NomYchMetodWork+AddYchMetodWork+23,6],Excel.Cells[NomYchMetodWork+AddYchMetodWork+23,7]].Borders[9].LineStyle:=1;
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+23,5]:='___';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+23,6]:='��������';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+23,8]:='20__';
  Excel.Cells[NomYchMetodWork+AddYchMetodWork+23,9]:='����';

  Excel.Range[Excel.Cells[1,1],Excel.Cells[NomYchMetodWork+AddYchMetodWork+23,10]].Font.Name:='Times New Roman';
  Excel.Workbooks[1].save;
  Excel.Workbooks.Close;
end;
begin
if FileExists(CurrentDir+'\������ ��������.xlsx') then
begin
//������ �� ���� ��������������
NomPrepod:=0;
while NomPrepod<Length(Prepod) do
  begin
  OsnDis:='';
  //��� ������� ������� ����� ������� � ��� ���
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
  NameOtch1:=CurrentDir+'\�������� �� ��������������\������\'+st+'.xlsx';//Prepod[NomPrepod].NameFilePrepod;
{  PosPapki:=Length(NameOtch)-1;
  while NameOtch[PosPapki]<>'\' do
    dec(PosPapki);
  NameOtch1:=Copy(NameOtch,1,PosPapki-1)+'\������\'+Copy(NameOtch,PosPapki+1,Length(NameOtch)-PosPapki); }
  if not DirectoryExists(CurrentDir+'\�������� �� ��������������\������') then
    ForceDirectories(CurrentDir+'\�������� �� ��������������\������');
  CopyFile(PChar(CurrentDir+'\������ ��������.xlsx'),PChar(NameOtch1),true);
  MeProtocol.Lines.Add('������ ���� ������:'+NameOtch1);
  SortNagryzkaPrepodNameDis(NomPrepod);
  NomNagr:=0;
  while (NomNagr<Length(Prepod[NomPrepod].Nagryzka)) do
    begin
    Prepod[NomPrepod].Nagryzka[NomNagr].FlagIndPlan:=0;
    inc(NomNagr);
    end;
  AddInfoPrepod(NameOtch1,100000,1000000,1);

  {����� ���� ��� ��������}
  SortNagryzkaPrepodTypeDis(NomPrepod);
  if Prepod[NomPrepod].Stavka<>'' then
  begin
  NameOtch1:=CurrentDir+'\�������� �� ��������������\������ ��������\'+st+'.xlsx';
//  NameOtch1:=Copy(NameOtch,1,PosPapki-1)+'\������ ��������\'+Copy(NameOtch,PosPapki+1,Length(NameOtch)-PosPapki);
  if not DirectoryExists(CurrentDir+'\�������� �� ��������������\������ ��������') then
    ForceDirectories(CurrentDir+'\�������� �� ��������������\������ ��������');
  CopyFile(PChar(CurrentDir+'\������ ��������.xlsx'),PChar(NameOtch1),true);
  MeProtocol.Lines.Add('������ ���� ������:'+NameOtch1);
  NomNagr:=0;
  while (NomNagr<Length(Prepod[NomPrepod].Nagryzka)) do
    begin
    Prepod[NomPrepod].Nagryzka[NomNagr].FlagIndPlan:=0;
    inc(NomNagr);
    end;
  SdvigHour:=0;
  if Prepod[NomPrepod].StavkaSovmest<>'' then
    SdvigHour:=StrToFloat(Prepod[NomPrepod].StavkaSovmest)*HourStavka*Prepod[NomPrepod].MesNeOplat/5;
  AddInfoPrepod(NameOtch1,Trunc(StrToFloat(Prepod[NomPrepod].Stavka)*HourStavka*(Prepod[NomPrepod].HourSem[1]/Prepod[NomPrepod].AllHour)+SdvigHour),Trunc(StrToFloat(Prepod[NomPrepod].Stavka)*HourStavka*(Prepod[NomPrepod].HourSem[2]/Prepod[NomPrepod].AllHour)-SdvigHour),2);
  end;

  if Prepod[NomPrepod].StavkaSovmest<>'' then
  begin
  NameOtch1:=CurrentDir+'\�������� �� ��������������\������ ��������\���������������� '+st+'.xlsx';
  //NameOtch1:=Copy(NameOtch,1,PosPapki-1)+'\������ ��������\���������������� '+Copy(NameOtch,PosPapki+1,Length(NameOtch)-PosPapki);
  if not DirectoryExists(CurrentDir+'\�������� �� ��������������\������ ��������') then
    ForceDirectories(CurrentDir+'\�������� �� ��������������\������ ��������');
  CopyFile(PChar(CurrentDir+'\������ ��������.xlsx'),PChar(NameOtch1),true);
  MeProtocol.Lines.Add('������ ���� ������:'+NameOtch1);
  AddInfoPrepod(NameOtch1,Trunc(StrToFloat(Prepod[NomPrepod].StavkaSovmest)*HourStavka*(Prepod[NomPrepod].HourSem[1]/Prepod[NomPrepod].AllHour)-SdvigHour),Trunc(StrToFloat(Prepod[NomPrepod].StavkaSovmest)*HourStavka*(Prepod[NomPrepod].HourSem[2]/Prepod[NomPrepod].AllHour)+SdvigHour),3);
  end;

  NameOtch1:=CurrentDir+'\�������� �� ��������������\������ ��������\��������� '+st+'.xlsx';
//  NameOtch1:=Copy(NameOtch,1,PosPapki-1)+'\������ ��������\��������� '+Copy(NameOtch,PosPapki+1,Length(NameOtch)-PosPapki);
  if not DirectoryExists(CurrentDir+'\�������� �� ��������������\������ ��������') then
    ForceDirectories(CurrentDir+'\�������� �� ��������������\������ ��������');
  CopyFile(PChar(CurrentDir+'\������ ��������.xlsx'),PChar(NameOtch1),true);
  MeProtocol.Lines.Add('������ ���� ������:'+NameOtch1);
  AddInfoPrepod(NameOtch1,100000,1000000,4);

  inc(NomPrepod);
  end;
ShowMessage('�������� ������� ���������');
VivodPlanNagryzka;
end
Else
  MeProtocol.Lines.Add('�� ������ ����:'+CurrentDir+'\������ ��������.xlsx')
end;

end.
