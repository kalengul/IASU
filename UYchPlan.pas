unit UYchPlan;

interface

uses UMain,UGroup, UNagryzka;

type
  TDisciplineYP = record
    Name:string;
    kaf:string;
    ZE:byte;
    ZESEM:array [1..12] of byte;
    NedelinaiaAZ:array [1..12] of Double;
    Hour,HourBezEK:longword;
    EkzSem,KyrsSem:array [1..12] of string;
    AuditorObSem,LekSem,PraktSem,LabSem,SRSSem:array [1..12] of Longword;
    SumAuditoria,SumSRS:Longword;
    ElektivDis:String;
    ElektivDisKaf:string;
    NagryzkaPrepod:array of TNagryzkaPrepod;

  end;

  TYP = class
    Name:string;
    Napravlenie:string;
    Profil:string;
    Group:TAGroup;
    Discipline:array of TDisciplineYP;
  end;

  var
    YP:array of TYP;

implementation

end.
