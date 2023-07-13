program rtf2html;

uses
  Classes, SysUtils, FileUtil, uRtf2Html;

procedure WriteHelp;
begin
  WriteLn('Syntax: rtf2html <filename>.rtf');
  WriteLn('  <filename> ... rtf file to be converted (can contain wildcards)');
  WriteLn('                 The output will be saved as <filename>.html');
  WriteLn('                 An already existing output file will be overwritten without notice.');
  Halt;
end;

var
  stream: TFileStream;
  converter: TRtf2HtmlConverter;
  filename: String;
  htmlFilename: String;
  files: TStrings;
begin
  if ParamCount = 0 then
    WriteHelp;

  files := TStringList.Create;
  try
    FindAllFiles(files, ExtractFileDir(ParamStr(1)), ExtractFileName(ParamStr(1)), false);
    for filename in files do
    begin
      WriteLn('Converting file ', filename, '... ');
      if not FileExists(filename) then
      begin
        WriteLn('  File not found.');
        Continue;
      end;

      stream := TFileStream.Create(filename, fmOpenRead or fmShareDenyWrite);
      try
        converter := TRtf2HtmlConverter.Create(stream);
        try
          converter.Indented := false;

          //WriteLn(converter.ConvertToHtmlString('test'));

          htmlFilename := ChangeFileExt(filename, '.html');
          converter.ConvertToHtml(htmlFilename);
          WriteLn('  Saved as ' + htmlFilename);

        finally
          converter.Free;
        end;
      finally
        stream.Free;
      end;
    end;
  finally
    files.Free;
  end;
end.

