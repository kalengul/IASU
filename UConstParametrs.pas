unit UConstParametrs;

interface

const
  KafAudLR: array [0..7] of string = ('���.207(3)','���.240�(3)','���.240�(3)','���.440(3)','���.444(3)','���.446(3)','���.434�(3)','���.434�(3)');
  KafAudLK: array [0..9] of string = ('���.207(3)','���.213(3)','���.240�(3)','���.240�(3)','���.401(3)','���.440(3)','���.444(3)','���.446(3)','���.434�(3)','���.434�(3)');
  kolsem = 2;
  kolrownagryzka = 7;
  ClGreen = $0024AA35;
  ClRed = $001C189A;

type
THourOnOneStudent = record
  Vid:string;
  Hour:Double;
end;

var
CurrentYear:longword;
CurrentSemestr:byte;
YearBakalavr,YearMagistr,YearAspirant:byte;
TimeSetPar:array of string;
HourStavka:longword;
ZKaf,ZKafSokr:string;
NomKaf:String;
CreateFilePrep:Boolean;
HourOnOneStudent:array of THourOnOneStudent;

Procedure InitializationParametrs;

implementation

Procedure InitializationParametrs;
begin
SetLength(TimeSetPar,6);
TimeSetPar[0]:='09:00-10:30';
TimeSetPar[1]:='10:45-12:15';
TimeSetPar[2]:='13:00-14:30';
TimeSetPar[3]:='14:45-16:15';
TimeSetPar[4]:='16:30-18:00';
TimeSetPar[5]:='18:15-19:45';
CurrentYear:=2019;
CurrentSemestr:=2;
YearBakalavr:=4;
YearAspirant:=4;
YearMagistr:=2;
HourStavka:=830;
ZKaf:='������ ���� ����������';
ZKafSokr:='������ �.�.';
NomKaf:='304';
SetLength(HourOnOneStudent,0);
CreateFilePrep:=false;
end;

end.
