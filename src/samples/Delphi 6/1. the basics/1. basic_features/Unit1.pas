unit Unit1;

interface

uses
  Snarl, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    Memo1: TMemo;
    Label1: TLabel;
    Label2: TLabel;
    CheckBox1: TCheckBox;
    procedure Button1Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  hr: Integer;
  szIcon: String;

begin

  hr := snarl_register('application/x-delphi-basic_features', 'Delphi 6 Test', '');
  if hr <= 0 then
    ShowMessage('Error registering with Snarl:' + IntToStr(hr))

  else
    begin
      if CheckBox1.Checked then
        szIcon := '&icon=!system-info';

      snDoRequest('notify?app-sig=application/x-delphi-basic_features' +
        szIcon +
        '&title=' + Edit1.Text + '&text=' + Memo1.Lines.Text +
        '&reply-to=' + IntToStr(Form1.Handle) + '&reply=' + IntTostr($0401) +
        '&action=Action 1234,@1234&action=Action 9876,@9876');

    end

  (*
                            "&priority=" & CStr(pri)  _
  *)

end;

procedure TForm1.FormDestroy(Sender: TObject);
begin

  Snarl.snarl_unregister('application/x-delphi-basic_features');

end;

end.
