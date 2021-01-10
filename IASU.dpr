program IASU;

uses
  Forms,
  UMain in 'UMain.pas' {FMain},
  UChangeNagryzka in 'UChangeNagryzka.pas' {FChangeNagryzka},
  UConstParametrs in 'UConstParametrs.pas',
  UGroup in 'UGroup.pas',
  UAuditoria in 'UAuditoria.pas',
  UYchPlan in 'UYchPlan.pas',
  USemPlan in 'USemPlan.pas',
  UNagryzka in 'UNagryzka.pas',
  USaveExcel in 'USaveExcel.pas',
  UMTO in 'UMTO.pas',
  ULoadExcel in 'ULoadExcel.pas',
  UIndividualPlanPrepod in 'UIndividualPlanPrepod.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFMain, FMain);
  Application.CreateForm(TFChangeNagryzka, FChangeNagryzka);
  Application.Run;
end.
