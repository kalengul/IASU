unit UGroup;

interface

type
  TGroup = class
    Nom:string;
    Plosh:string;
    kaf:string;
    Forma:string;
    KolStudent:Byte;
    Kyrs:Byte;
    Napravlenie,ShifrNapravlenie,Profil,ShifrProfil:String;
    Constructor Create;
    Destructor Destroy;
  end;

  TGroupKyrs = record
               NameGroupKyrs:string;
               Group:array of string;
               end;

  TAGroup = array of TGroup;

  var
  AllGroup:TAGroup;
  NameAllGroup:array of TGroupKyrs;

function SearchAndCreateGroup(NameGroup:string):TGroup;
Function SearchInMassiveGroup(Group:TAGroup; NameGroup:string):Longword;


implementation

function SearchAndCreateGroup(NameGroup:string):TGroup;
var
  Nom:Longword;
  NewGroup:TGroup;
begin
Nom:=0;
While (Nom<Length(AllGroup)) and (AllGroup[Nom].Nom<>NameGroup) do
  inc(Nom);
if (Nom<Length(AllGroup)) then
  result:=AllGroup[Nom]
else
  begin
  NewGroup:=TGroup.Create;
  NewGroup.Nom:=NameGroup;
  result:=AllGroup[Nom];
  end;
end;

Function SearchInMassiveGroup(Group:TAGroup; NameGroup:string):Longword;
var
Nom:Longword;
begin
{while pos(',',NameGroup)<>0 do
  delete(NameGroup,1,pos(',',NameGroup)+1); }
Nom:=0;
while (Nom<Length(Group)) and (Pos(Group[Nom].Nom,NameGroup)=0) do
  inc(Nom);
if (Nom<Length(Group)) then
  result:=nom
else
  result:=65000;
end;

Constructor TGroup.Create;
var
NomElement:Longword;
  begin
  NomElement:=Length(AllGroup);
  SetLength(AllGroup,NomElement+1);
  AllGroup[NomElement]:=self;
  end;
Destructor TGroup.Destroy;
  begin

  end;

end.
