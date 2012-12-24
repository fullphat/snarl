program basic_features;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  Snarl in '..\..\..\..\..\hdr\Delphi\Snarl.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
