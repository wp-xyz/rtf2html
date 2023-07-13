unit uRtf2Html;

{$mode ObjFPC}{$H+}

{.$DEFINE Debug_rtf}

interface

uses
  Classes, SysUtils, LConvEncoding, RtfPars, LazLoggerBase;

const
  rtfDefault = -1;

type
  TRtfFontStyle = (rfsBold, rfsItalic, rfsStrikeThrough, rfsUnderline, rfsShadow, rfsOutline);
  TRtfFontStyles = set of TRtfFontStyle;

  TRtfCharPosition = (rcpNormal, rcpSubscript, rcpSuperscript);

  TRtfCapsStyle = (rcsNormal, rcsCaps, rcsSmallCaps);
  TRtfUnderlineStyle = (rusSingle, rusDouble, rusDotted, rusWord);

  // Character format --> <span> tag
  TRtfCharFormat = record
    DefaultFontNum: Integer;
    FontNum: Integer;
    FontSize: Double;
    FontStyle: TRtfFontStyles;
    Position: TRtfCharPosition;
    CapsStyle: TRtfCapsStyle;
    UnderlineStyle: TRtfUnderlineStyle;

    FgColor: Integer;
    BkColor: Integer;
    UnicodeCount: Integer;
  end;

  TRtfAlignment = (raLeft, raCenter, raRight, raJustify);

  // Paragraph format ---> <p> tag
  TRtfParFormat = record
    Alignment: TRtfAlignment;
    LeftIndent: Integer;    // in twips (1 inch = 1440 twips)
    RightIndent: Integer;
    FirstIndent: Integer;
    SpacingBefore: Integer; // in twips
    SpacingAfter: Integer;
    SpacingBetween: Integer;
  end;

  TRtfCellBorder = (rcbLeft, rcbTop, rcbRight, rcbBottom);
  TRtfCellBorderStyle = (rcbsNone, rcbsSingle);
  TRtfCellBorderStyles = array[TRtfCellBorder] of TRtfCellBorderStyle;

  TRtfTblAttrib = record
    CellText: array of string;
    ColPos: array of Integer;
    CellBorders: array of TRtfCellBorderStyles;
    HorPadding: Integer;
  end;

  TRtfPictFormat = (rpfUnknown, rpfBMP, rpfPNG, rpfJPEG, rpfWMF);
  TRtfPictAttrib = record
    Width: Integer;      // Twips in KMemo (Pixels according to MS documentation)
    Height: Integer;
    GoalWidth: Integer;  // Twips
    GoalHeight: Integer;
    ScaleX: Integer;     // Percentage
    ScaleY: Integer;
    Format: TRtfPictFormat;
  end;

  TRtfDocAttrib = record
  end;

  TRtfTextState = (
    rtsPreparingParFormat,   // active when any paragraph formatting command has been sent
    rtsPreparingCharFormat,  // active when any character formatting command has been sent
    rtsPreparingTable,       // active inside a table structure
    rtsWaitingForNextRow     // active when a row has been written
  );
  TRtfTextStates = set of TRtfTextState;

  TRtfStackElement = record
    Destination: Integer;
    CharFormat: TRtfCharFormat;
    ParFormat: TRtfParFormat;
  end;

  TRtf2HtmlConverter = class
  private
    FParser: TRtfParser;
    FVersionOK: Boolean;
    FCodePage: String;
    FTitle: String;
    FActiveDestination: Integer;
    FTextState: TRtfTextStates;  // Describes which part of the formatted text is currently selected
    FPixelsPerInch: Integer;
    FIndented: Boolean;

  protected
    // Parser handlers
    procedure DoCharAttributes;
    procedure DoCharSet;
    procedure DoCtrl;
    procedure DoDestination;
    procedure DoDocAttributes;
    procedure DoGroup;
    procedure DoParAttributes;
    procedure DoPictAttributes;
    procedure DoPicture;
    procedure DoSectAttributes;
    procedure DoSpecialChar;
    procedure DoTblAttributes;
    procedure DoText;
  protected
    // Attributes and stack
    FStack: array of TRtfStackElement;
    FFontPending: Boolean;
    FDocAttrib: TRtfDocAttrib;
    FActiveCharFormat: TRtfCharFormat;
    FActiveParFormat: TRtfParFormat;
    FActivePictAttrib: TRtfPictAttrib;
    FTblAttrib: TRtfTblAttrib;
    FTblRows: array of string;
    FCurrCellBorder: TRtfCellBorder;
    FCurrUnicodeCounter: Integer;
    FDefaultFontFamily: String;
    FDefaultFontSize: Double;
    procedure ClearCharFormat;
    procedure ClearParFormat;
    procedure ClearPictAttrib;
    procedure ClearTblAttrib;
    procedure GetDefaultFontName;
    procedure PrepareCharFormat(Enable: Boolean);
    procedure PrepareParFormat(Enable: Boolean);
    procedure PrepareTable(Enable: Boolean);
    function PreparingCharFormat: Boolean;
    function PreparingParFormat: Boolean;
    function PreparingTable: Boolean;

    procedure PushGroup;
    procedure PopGroup;

  protected
    // Writing output
    FOutput: TStrings;
    FCurrParText: String;
    FCurrText: String;
    FAnsiText: RawByteString;
    FPictData: String;
    FPictDataCount: Integer;
    FIndentLevel: Integer;
    FFormatSettings: TFormatSettings;
    procedure Add(AString: String);
    procedure AddText(const AText: String);
    function Indentation: String;
    procedure WriteDefaultFont;
    function WriteFormattedText: String;
    procedure WriteHtmlFooter;
    procedure WriteHtmlHeader;
    procedure WriteParagraph;
    procedure WritePicture;
    procedure WriteTable;
    procedure WriteTableRow;

    procedure DisplayStack(Info: String);
    procedure DisplayOutput;

  public
    constructor Create(RTFStream: TStream; APixelsPerInch: Integer = 96);
    destructor Destroy; override;
    procedure ConvertToHtml(AFileName: String);
    procedure ConvertToHtml(AStream: TStream; ATitle: String);
    function ConvertToHtmlString(ATitle: String): String;
    property Indented: Boolean read FIndented write FIndented default false;
  end;

implementation

uses
  Base64, fpHtml;

// It can be assumed that C is a valid rtf color.
function rtfColorToString(C: PRtfColor): String;
begin
  Result := Format('#%.2x%.2x%.2x', [C^.rtfCRed, C^.rtfCGreen, C^.rtfCBlue]);
end;

function rtfDefaultColor(C: PRtfColor): Boolean;
begin
  Result := (C^.rtfCRed = -1) or (C^.rtfCGreen = -1) or (C^.rtfCBlue = -1);
end;

function TwipsToMM(twips: Integer): Double;
begin
  Result := twips / 1440 * 25.4;
end;

function TwipsToPts(twips: Integer): Double;
begin
  Result := twips / 20;
end;

function TwipsToPx(twips, PPI: Integer): Integer;
begin
  Result := round(twips / 1440 * PPI);
end;


{ TRtf2HtmlConverter }

constructor TRtf2HtmlConverter.Create(RTFStream: TStream; APixelsPerInch: Integer = 96);
begin
  inherited Create;

  FIndented := false;
  FPixelsPerInch := APixelsPerInch;

  FFormatSettings := DefaultFormatSettings;
  FFormatSettings.DecimalSeparator := '.';

  FDefaultFontFamily := '';
  FDefaultFontSize := 0.0;

  FIndentLevel := 0;

  FOutput := TStringList.Create;

  FParser := TRtfParser.Create(RTFStream);
  FParser.ClassCallBacks[rtfText] := @DoText;
  FParser.ClassCallBacks[rtfGroup] := @DoGroup;
  FParser.ClassCallBacks[rtfControl] := @DoCtrl;
  FParser.DestinationCallBacks[rtfPict] := @DoPicture;
end;

destructor TRtf2HtmlConverter.Destroy;
begin
  FParser.Free;
  FOutput.Free;
  inherited;
end;

procedure TRtf2HtmlConverter.Add(AString: String);
begin
  FOutput.Add(Indentation + AString);
end;

procedure TRtf2HtmlConverter.AddText(const AText: String);
var
  n: Integer;
begin
  if AText = '' then
    exit;
  n := FOutput.Count;
  if n = 0 then
    FOutput.Add(AText)
  else
    FOutput[n-1] := FOutput[n-1] + AText;
end;

procedure TRtf2HtmlConverter.ClearCharFormat;
begin
  FActiveCharFormat.FontNum := FActiveCharFormat.DefaultFontNum;
  if FDefaultFontSize = 0 then FDefaultFontSize := 12;
  FActiveCharFormat.FontSize := FDefaultFontSize;
  FActiveCharFormat.FontStyle := [];
  FActiveCharFormat.Position := rcpNormal;
  FActiveCharFormat.CapsStyle := rcsNormal;
  FActiveCharFormat.UnderlineStyle := rusSingle;
  FActiveCharFormat.FgColor := -1;
  FActiveCharFormat.BkColor := -1;
  FActiveCharFormat.UnicodeCount := 1;
end;

procedure TRtf2HtmlConverter.ClearParFormat;
begin
  FActiveParFormat.Alignment := raLeft;
  FActiveParFormat.LeftIndent := 0;
  FActiveParFormat.RightIndent := 0;
  FActiveParFormat.FirstIndent := 0;
  FActiveParFormat.SpacingBefore := 0;
  FActiveParFormat.SpacingAfter := 0;
  FActiveParFormat.SpacingBetween := 0;
end;

procedure TRtf2HtmlConverter.ClearPictAttrib;
begin
  FActivePictAttrib.GoalWidth := 0;
  FActivePictAttrib.GoalHeight := 0;
  FPictData := '';
  FPictDataCount := 0;
end;

procedure TRtf2HtmlConverter.ClearTblAttrib;
begin
  SetLength(FTblAttrib.CellText, 0);
  SetLength(FTblAttrib.ColPos, 0);
  SetLength(FTblAttrib.CellBorders, 0);
end;

procedure TRtf2HtmlConverter.ConvertToHtml(AFileName: String);
var
  stream: TFileStream;
begin
  stream := TFileStream.Create(AFileName, fmCreate or fmShareDenyWrite);
  try
    ConvertToHtml(stream, ExtractFileName(AFileName));
  finally
    stream.Free;
  end;
end;

procedure TRtf2HtmlConverter.ConvertToHtml(AStream: TStream; ATitle: String);
begin
  FTitle := ATitle;
  FActiveDestination := rtfDefault;

  FOutput.Clear;

  WriteHtmlHeader;
  FParser.StartReading;
  WriteHtmlFooter;
  WriteDefaultFont;

  FOutput.SaveToStream(AStream);
end;

function TRtf2HtmlConverter.ConvertToHtmlString(ATitle: String): String;
var
  stream: TStringStream;
begin
  stream := TStringStream.Create('');
  try
    ConvertToHtml(stream, ATitle);
    stream.Position := 0;
    Result := stream.ReadString(stream.Size);
  finally
    stream.Free;
  end;
end;

procedure TRtf2HtmlConverter.DisplayStack(Info: String);
var
  i: Integer;
begin
  {
  WriteLn(Info, ' (', Length(FStack) , ' items on stack)');
  if Length(FStack) > 0 then
  begin
    Write('  ');
    for i := 0 to High(FStack) do
      Write('f',FStack[i].FontNum, 'fs', FStack[i].FontSize:0:1, '  ');
    WriteLn;
  end;
  }
end;

procedure TRtf2HtmlConverter.DisplayOutput;
var
  i: Integer;
begin
  {
  for i := 0 to FOutput.Count-1 do
    Writeln(FOutput[i]);
    }
end;

procedure TRtf2HtmlConverter.DoCharAttributes;
begin
  // Write previously collected text in the still-active character format
  // (which will be changed afterwards here).
  if (FCurrText <> '') or (FAnsiText <> '') then
    FCurrParText := FCurrParText + WriteFormattedText();
  PrepareCharFormat(true);

  // Modify current character attributes.
  case FParser.rtfMinor of
    rtfPlain:     // \plain -- reset font to default
      begin
        ClearCharFormat;
        PrepareCharFormat(true);
      end;
    rtfFontNum:   // \fN --- use font #N
      begin
        FActiveCharFormat.FontNum := FParser.rtfParam;
        PrepareCharFormat(true);
      end;
    rtfFontSize:   // \fsN -- use font size N
      begin
        FActiveCharFormat.FontSize := FParser.rtfParam / 2.0;
        if (FActiveCharFormat.FontNum = FActiveCharFormat.DefaultFontNum) and (FDefaultFontSize = 0) then
          FDefaultFontSize := FActiveCharFormat.FontSize;
        PrepareCharFormat(true);
      end;
    rtfBold:       // \b -- bold font
      begin
        if FParser.rtfParam = 0 then
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle - [rfsBold]
        else
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle + [rfsBold];
        PrepareCharFormat(true);
      end;
    rtfItalic:      // \i -- italic font
      begin
        if FParser.rtfParam = 0 then
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle - [rfsItalic]
        else
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle + [rfsItalic];
        PrepareCharFormat(true);
      end;
    rtfUnderline,   // \ul -- single underline, \ul0 -- no underline
    rtfWUnderline:  // \ulw -- word underline
      begin
        if FParser.rtfParam = 0 then
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle - [rfsUnderline]
        else
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle + [rfsUnderline];
        FActiveCharFormat.UnderlineStyle := rusSingle;
        PrepareCharFormat(true);
      end;
    rtfDUnderline:  // \uld -- dotted underline
      begin
        if FParser.rtfParam = 0 then
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle - [rfsUnderline]
        else
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle + [rfsUnderline];
        FActiveCharFormat.UnderlineStyle := rusDotted;
        PrepareCharFormat(true);
      end;
    rtfDbUnderline:   // \uldb -- double underline
      begin
        if FParser.rtfParam = 0 then
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle - [rfsUnderline]
        else
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle + [rfsUnderline];
        FActiveCharFormat.UnderlineStyle := rusDouble;
        PrepareCharFormat(true);
      end;
    rtfNoUnderline:  // \ulnone -- no underline
      begin
        FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle - [rfsUnderline];
        PrepareCharFormat(true);
      end;
    rtfStrikeThru:     // strike-through text
      begin
        if FParser.rtfParam = 0 then
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle - [rfsStrikeThrough]
        else
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle + [rfsStrikeThrough];
        PrepareCharFormat(true);
      end;
    rtfAllCaps:
      begin
        if FParser.rtfParam =0 then
          FActiveCharFormat.CapsStyle := rcsNormal
        else
          FActiveCharformat.CapsStyle := rcsCaps;
      end;
    rtfSmallCaps:
      begin
        if FParser.rtfParam =0 then
          FActiveCharFormat.CapsStyle := rcsNormal
        else
          FActiveCharformat.CapsStyle := rcsSmallCaps;
      end;
    rtfShadow:  // Shadow effect
      begin
        if FParser.rtfParam = 0 then
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle - [rfsShadow]
        else
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontStyle + [rfsShadow];
        PrepareCharFormat(true);
      end;
    rtfOutline:   // Outline effect
      begin
        if FParser.rtfParam = 0 then
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontSTyle - [rfsOutline]
        else
          FActiveCharFormat.FontStyle := FActiveCharFormat.FontSTyle + [rfsOutline];
        PrepareCharFormat(true);
      end;
    rtfSubScript:   // \sub -- subscript character
      begin
        if FParser.rtfParam = 0 then
          FActiveCharFormat.Position := rcpNormal
        else
          FActiveCharFormat.Position := rcpSubscript;
        PrepareCharFormat(true);
      end;
    rtfSuperScript:    // \super -- superscript
      begin
        if FParser.rtfParam = 0 then
          FActiveCharFormat.Position := rcpNormal
        else
          FActiveCharFormat.Position := rcpSuperscript;
        PrepareCharFormat(true);
      end;
    rtfNoSuperSub:    // \nosupersub -- normal text position
      begin
        FActiveCharFormat.Position := rcpNormal;
        PrepareCharFormat(true);
      end;
    rtfForeColor:     // \cfN -- text color (foreground)
      begin
        FActiveCharFormat.FgColor := FParser.rtfParam;
        PrepareCharFormat(true);
      end;
    rtfBackColor:     // \cbN -- text background color
      begin
        FActiveCharFormat.BkColor := FParser.rtfParam;
        PrepareCharFormat(true);
      end;
  end;
end;

procedure TRtf2HtmlConverter.DoCharSet;
begin
  case FParser.rtfMinor of
    rtfMacCharSet:
      FCodePage := encodingCPMac;
    rtfAnsiCharSet:
      FCodePage := encodingAnsi;
    rtfPCCharSet:
      FCodePage := encodingCP437;
    rtfPcaCharSet:
      FCodePage := encodingCP850;
    otherwise
      FCodePage := encodingAnsi;
  end;
  // there is also a \ansicpgN which is handled by DoSpecialChar
end;

procedure TRtf2HtmlConverter.DoCtrl;
begin
  case FParser.rtfMajor of
    rtfVersion:
      FVersionOK := (FParser.rtfParam = 1);
    rtfDefFont:
      FActiveCharFormat.DefaultFontNum := FParser.rtfParam;
    rtfCharAttr:
      DoCharAttributes;
    rtfCharSet:
      DoCharset;
    {
    rtfDestination:
      DoDestination;
      }
    rtfDocAttr:
      DoDocAttributes;
    rtfSpecialChar:
      DoSpecialChar;
    rtfParAttr:
      DoParAttributes;
    rtfPictAttr:
      DoPictAttributes;
    rtfSectAttr:
      DoSectAttributes;
    rtfTblAttr:
      DoTblAttributes;
  end;
end;

procedure TRtf2HtmlConverter.DoDestination;
begin
  //
end;

procedure TRtf2HtmlConverter.DoDocAttributes;
begin
  {
  case FParser.rtfMinor of
  end;
  }
end;

procedure TRtf2HtmlConverter.DoGroup;
var
  prevDestination: Integer;
begin
  if not FVersionOK then
    exit;

  case FParser.rtfMajor of
    rtfBeginGroup:
      begin
        {$IFDEF Debug_rtf}
        DebugLn('[DoGroup] {');
        {$ENDIF}
        inc(FIndentLevel);
        PushGroup;
      end;
    rtfEndGroup:
      begin
        {$IFDEF Debug_rtf}
        DebugLn('[DoGroup] }');
        {$ENDIF}
        prevDestination := FActiveDestination;
        if prevDestination = rtfFontTbl then
          GetDefaultFontName;
        if (prevDestination = rtfDefault) then
        begin
          if PreparingTable then
            WriteTable;
          if (FCurrText <> '') or (FAnsiText <> '') then   // after closing inlined groups
            FCurrParText := FCurrParText + WriteFormattedText();
          if PreparingParFormat and (FCurrParText <> '') then
            WriteParagraph;
        end;

        PopGroup;

        if (prevDestination = rtfPict) and (FActiveDestination <> rtfPict) then
          WritePicture;

        dec(FIndentLevel);
      end;
  end;
end;

procedure TRtf2HtmlConverter.DoParAttributes;
var
  n: Integer;
begin
  case FParser.RtfMinor of
    rtfParDef:    // \pard -- begins a new paragraph and defines its properties
      begin
        {$IFDEF Debug_rtf}
        DebugLn('[DoParAttributes] \pard');
        {$ENDIF}
        // Write any previously collected paragraph text
        if (FCurrText <> '') or (FAnsiText <> '') then
          FCurrParText := FCurrParText + WriteFormattedText();
        if (FCurrParText <> '') then
          WriteParagraph();
        ClearParFormat;
        PrepareParFormat(true);
      end;
    rtfQuadLeft:   // \ql -- left-aligned paragraph
      begin
        FActiveParFormat.Alignment := raLeft;
        PrepareParFormat(true);
      end;
    rtfQuadRight:  // \qr -- right-aligned paragraph
      begin
        FActiveParFormat.Alignment := raRight;
        PrepareParFormat(true);
      end;
    rtfQuadCenter:   // \qc -- centered paragraph
      begin
        FActiveParFormat.Alignment := raCenter;
        PrepareParFormat(true);
      end;
    rtfQuadJust:    // \qj -- justified paragrah
      begin
        FActiveParFormat.Alignment := raJustify;
        PrepareParFormat(true);
      end;
    rtfLeftIndent:  // \li -- left margin of paragraph
      begin
        FActiveParFormat.LeftIndent := FParser.rtfParam;
        PrepareParFormat(true);
      end;
    rtfRightIndent:  // \ri -- right margin of paragraph
      begin
        FActiveParFormat.RightIndent := FParser.rtfParam;
        PrepareParFormat(true);
      end;
    rtfFirstIndent:  // \fi -- indentation of first line of paragraph
      begin
        FActiveParFormat.FirstIndent := FParser.rtfParam;
        PrepareParFormat(true);
      end;
    rtfSpaceAfter:   // \sa -- space after paragraph
      begin
        FActiveParFormat.SpacingAfter := FParser.rtfParam;
        PrepareParFormat(true);
      end;
    rtfSpaceBefore:   // \sb -- space before paragraph
      begin
        FActiveParFormat.SpacingBefore := FParser.rtfParam;
        PrepareParFormat(true);
      end;
    rtfSpaceBetween:   // \sl -- space between paragraphs
      begin
        FActiveParFormat.SpacingBetween := FParser.rtfParam;
        PrepareParFormat(true);
      end;

    // Table-related
    rtfInTable:      // \intbl - signals that paragraph belongs to a table
      begin
        {
        FInTable := true;
        }
      end;
    rtfBorderSingle:  // Single cell border
      begin
        n := Length(FTblAttrib.CellBorders);
        SetLength(FTblAttrib.CellBorders, n+1);
        FTblAttrib.CellBorders[n, FCurrCellBorder] := rcbsSingle;
      end;
  end;
end;

procedure TRtf2HtmlConverter.DoPictAttributes;
begin
  case FParser.rtfMinor of
    rtfPicWid:
      FActivePictAttrib.Width := FParser.rtfParam;
    rtfPicHt:
      FActivePictAttrib.Height := FParser.rtfParam;
    rtfPicGoalWid:
      FActivePictAttrib.GoalWidth := FParser.rtfParam;
    rtfPicScaleX:
      FActivePictAttrib.ScaleX := FParser.rtfParam;
    rtfPicScaleY:
      FActivePictAttrib.ScaleY := FParser.rtfParam;
    rtfPicGoalHt:
      FActivePictAttrib.GoalHeight := FParser.rtfParam;
    rtfPicCropLeft:
      ;                         // to be completed...
    rtfPicCropTop:
      ;
    rtfPicCropRight:
      ;
    rtfPicCropBottom:
      ;
    rtfPngBlip:
      FActivePictAttrib.Format := rpfPNG;
    rtfJpegBlip:
      FActivePictAttrib.Format := rpfJPEG;
    rtfWinMetaFile:
      ;
      //FActivePictAttrib.Format := rpfWMF;  -- not working
  end;
end;

procedure TRtf2HtmlConverter.DoPicture;
begin
  // Before switching to picture mode write the previously accumulated
  // paragraph to html.
  if FActiveDestination = rtfDefault then
  begin
    if (FCurrText <> '') or (FAnsiText <> '') then
      FCurrParText := FCurrParText + WriteFormattedText();
  end;

  // Switch to picture mode
  FActiveDestination := rtfPict;

  // Erase the picture attribute record
  ClearPictAttrib;
end;

procedure TRtf2HtmlConverter.DoSectAttributes;
begin
  {
  case FParser.RtfMinor of
  end;
  }
end;

procedure TRtf2HtmlConverter.DoSpecialChar;
var
  c: WideChar;
  n: Integer;
begin
  case FParser.RtfMinor of
    rtfLine:       // \line --> new line tag
      FAnsiText := FAnsiText + '<br />';
    rtfNoBrkSpace: // \~ --> non-breaking space
      FAnsiText := FAnsiText + '&nbsp;';
    rtfTab:        // tab character, but: is not rendered in html
      FAnsiText := FAnsiText + '&#9;';
    rtfUnicodeID:  // \u --> signals that a unicode character is sent
      begin
        if FAnsiText <> '' then
        begin;
          FCurrText := FCurrText + ConvertEncoding(FAnsiText, FCodePage, encodingUTF8);
          FAnsiText := '';
        end;
        c := WideChar(FParser.rtfParam);
        FCurrText := FCurrText + UTF8Encode(c);
        FCurrUnicodeCounter := FActiveCharFormat.UnicodeCount;
      end;
    rtfUnicodeCount:  // \ucN --> number of bytes to skip after \u command
      FActiveCharFormat.UnicodeCount := FParser.rtfParam;
    rtfPar:        // \par --> new paragraph (with same settings as before)
      begin
        // Write the previously collected paragraph text
        if (FCurrText <> '') or (FAnsiText <> '') then
          FCurrParText := FCurrParText + WriteFormattedText();
        WriteParagraph();
        FTextState := FTextState + [rtsPreparingParFormat];
        FAnsiText := '';
        FCurrText := '';
        FCurrParText := '';
      end;
    rtfOptDest:    // \* --> ignore up to end of group
      begin
        FParser.SkipGroup;
        PopGroup;
      end;
    rtfAnsiCodePage:
      FCodePage := 'cp' + IntToStr(FParser.rtfParam);

    // Table-related
    rtfCell:       // \cell -- Ends a table cell
      begin
        n := Length(FTblAttrib.CellText);
        SetLength(FTblAttrib.CellText, n+1);
        if FAnsiText <> '' then
          FCurrText := FCurrText + ConvertEncoding(FAnsiText, FCodePage, encodingUTF8);
        FTblAttrib.CellText[n] := FCurrText;
        FCurrText := '';
        FCurrParText := '';
        FAnsiText := '';
      end;
    rtfRow:        // \row -- Ends a table row
      WriteTableRow;
  end;
end;

procedure TRtf2HtmlConverter.DoTblAttributes;
var
  n: Integer;
begin
  case FParser.RtfMinor of
    rtfRowDef:      // \trowd -- begins a new row
      begin
        if not PreparingTable then
        begin
          if PreparingCharFormat then
          begin
            if ((FCurrText <> '') or (FAnsiText <> '')) and (FCurrParText <> '') then
              FCurrParText := FCurrParText + WriteFormattedText();
          end;
          if PreparingParFormat then
            WriteParagraph();
          PrepareTable(true);
          PrepareParFormat(true);
          PrepareCharFormat(true);
          SetLength(FTblRows, 0);
        end;
        FTextState := FTextState - [rtsWaitingForNextRow];
        ClearTblAttrib;
      end;
    rtfCellPos:     // \cellx  -- position of right edge of cell
      begin
        n := Length(FTblAttrib.ColPos);
        SetLength(FTblAttrib.ColPos, n+1);
        FTblAttrib.ColPos[n] := FParser.rtfParam;
      end;
    rtfCellBordLeft:     // \clbrdrl
      FCurrCellBorder := rcbLeft;
    rtfCellBordTop:      // \clbrdrt
      FCurrCellBorder := rcbTop;
    rtfCellBordRight:    // \clbrdrr
      FCurrCellBorder := rcbRight;
    rtfCellBordBottom:   // clbrdrb
      FCurrCellBorder := rcbBottom;
    rtfRowGapH:          // trgaph -- horiz distance between cell text and border
      FTblAttrib.HorPadding := FParser.rtfParam;
  end;

  // We have reached the end of the table
  if rtsWaitingForNextRow in FTextState then
    WriteTable;
end;

procedure TRtf2HtmlConverter.DoText;
const
  DATA_BLOCK_SIZE = 1024;
var
  c: Char;
  s: String;
begin
  if FActiveDestination = rtfDefault then
  begin
    if (rtsPreparingParFormat in FTextState) and (FCurrParText <> '') then
      WriteParagraph();
    FTextState := FTextState - [rtsPreparingParFormat, rtsPreparingCharFormat];
  end;

  // A unicode character has just been read. It is followed by an ANSI
  // replacement character which must be ignored here since unicode is handled
  // by our reader.
  // TO DO: We assume \uc1 which may not be true always. Support \ucN where N denotes the number of bytes in the replacement character.
  if FCurrUnicodeCounter > 0 then
  begin
    dec(FCurrUnicodeCounter);
    exit;
  end;

  // Get the character...
  c := chr(FParser.RTFMajor);
  if c = #0 then   // last character
    exit;

  // ... and append it to current text
  if FActiveDestination = rtfPict then
  begin
    if FPictDataCount = Length(FPictData) then
      SetLength(FPictData, Length(FPictData) + DATA_BLOCK_SIZE);
    inc(FPictDataCount);
    FPictData[FPictDataCount] := c;
  end else
  begin
    FAnsiText := FAnsiText + c;
    {
    if c > #127 then
    begin
      s := ConvertEncoding(c, FCodePage, encodingUTF8);
      FCurrText := FCurrText + s;
    end else
      FCurrText := FCurrText + c;
    }
  end;
end;

{ Stores the default font parameters for the body style which will be written
  later }
procedure TRtf2HtmlConverter.GetDefaultFontName;
var
  font: PRtfFont;
begin
  font := FParser.Fonts[FActiveCharFormat.DefaultFontNum];
  if Assigned(font) then
    FDefaultFontFamily := font^.rtfFName
  else
    FDefaultFontFamily := '';
end;

{ Calculates the indentation string for nicely formatted HTML, but only if
  FIndented is true}
function TRtf2HtmlConverter.Indentation: String;
const
  SPACES_PER_LEVEL = 2;
var
  i: Integer;
begin
  if FIndented then
  begin
    SetLength(Result, FIndentLevel * SPACES_PER_LEVEL);
    if FIndentLevel > 0 then
      FillChar(Result[1], Length(Result), ' ');
  end else
    Result := '';
end;

procedure TRtf2HtmlConverter.PrepareCharFormat(Enable: Boolean);
begin
  if Enable then
    FTextState := FTextState + [rtsPreparingCharFormat]
  else
    FTextState := FTextState - [rtsPreparingCharFormat];
end;

procedure TRtf2HtmlConverter.PrepareParFormat(Enable: Boolean);
begin
  if Enable then
    FTextState := FTextState + [rtsPreparingParFormat]
  else
    FTextState := FTextState - [rtsPreparingParFormat];
end;

procedure TRtf2HtmlConverter.PrepareTable(Enable: Boolean);
begin
  if Enable then
    FTextState := FTextState + [rtsPreparingTable]
  else
    FTextState := FTextState - [rtsPreparingTable];
end;

function TRtf2HtmlConverter.PreparingCharFormat: Boolean;
begin
  Result := rtsPreparingCharFormat in FTextState;
end;

function TRtf2HtmlConverter.PreparingParFormat: Boolean;
begin
  Result := rtsPreparingParFormat in FTextState;
end;

function TRtf2HtmlConverter.PreparingTable: Boolean;
begin
  Result := rtsPreparingTable in FTextState;
end;

{ Whenever a "group" is sent (beginning with curly bracket) the state of the
  converter is pushed to a stack. }
procedure TRtf2HtmlConverter.PushGroup;
var
  n: Integer;
begin
  n := Length(FStack);
  SetLength(FStack, n+1);
  FStack[n].Destination := FActiveDestination;
  FStack[n].CharFormat := FActiveCharFormat;
  FStack[n].ParFormat := FActiveParFormat;
end;

procedure TRtf2HtmlConverter.PopGroup;
var
  n: Integer;
begin
  n := Length(FStack);
  if n <= 0 then
    exit;
  FActiveDestination := FStack[n-1].Destination;
  FActiveCharFormat := FStack[n-1].CharFormat;
  FActiveParFormat := FStack[n-1].ParFormat;
  SetLength(FStack, n-1);
end;

procedure TRtf2HtmlConverter.WriteDefaultFont;
var
  i: Integer;
  s: String;
begin
  if FDefaultFontFamily = '' then
    exit;

  for i := 0 to FOutput.Count-1 do
  begin
    s := FOutput[i];
    if pos('<body>', s) > 0 then
    begin
      Delete(s, Length(s), 1);
      s := s + Format(' style="font-family:%s;font-size:%.1fpt;">',
        [FDefaultFontFamily, FDefaultFontSize], FFormatSettings);
      FOutput[i] := s;
      exit;
    end;
  end;
end;

function TRtf2HtmlConverter.WriteFormattedText: String;
var
  font: PRTFFont;
  fontFamily: String;
  fontSize: String;
  fontStyle: String = '';
  fontColor: String = '';
  bkColor: String = '';
  position: String = '';
  caps: String = '';
  rtfColor: PRtfColor;
  fontSizeValue: Double;
  utf8Text: String;
begin
  font := FParser.Fonts[FActiveCharFormat.FontNum];
  if font <> nil then
    fontFamily := Format('font-family:%s;', [font^.rtfFName]);

  fontSizeValue := FActiveCharFormat.FontSize;
  if fontSizeValue = 0 then
    fontSizeValue := 12;
  if FActiveCharFormat.Position <> rcpNormal then
    fontSizeValue := fontSizeValue * 0.67;
  fontSize := Format('font-size:%.1fpt;', [fontSizeValue], FFormatSettings);

  if (FActiveCharFormat.FgColor <> -1) then
  begin
    rtfColor := FParser.Colors[FActiveCharFormat.FgColor];
    if (rtfColor <> nil) and not rtfDefaultColor(rtfColor) then
      fontColor := Format('color:%s;', [rtfColorToString(rtfColor)]);
  end;

  if FActiveCharFormat.BkColor <> -1 then
  begin
    rtfColor := FParser.Colors[FActiveCharFormat.BkColor];
    if (rtfColor <> nil) and not rtfDefaultColor(rtfColor) then
      bkColor := Format('background-color:%s;', [rtfColorToString(rtfColor)]);
  end;

  fontStyle := '';
  if (rfsBold in FActiveCharFormat.FontStyle) then
    fontStyle := fontStyle + 'font-weight:bold;';
  if (rfsItalic in FActiveCharFormat.FontStyle) then
    fontStyle := fontStyle + 'font-style:italic;';
  if (rfsUnderline in FActiveCharFormat.FontStyle) then
    case FActiveCharFormat.UnderlineStyle of
      rusSingle,
      rusWord: fontStyle := fontStyle + 'text-decoration:underline;';
      rusDouble: fontStyle := fontStyle + 'text-decoration-line:underline;text-decoration-style:double;';
    end;
  if (rfsStrikeThrough in FActiveCharFormat.FontStyle) then
    fontStyle := fontStyle + 'text-decoration:line-through;';
  if (rfsShadow in FActiveCharFormat.FontStyle) then
    fontStyle := fontStyle + 'text-shadow:2px 2px #CCCCCC;';
  if (rfsOutline in FActiveCharFormat.FontStyle) then
    fontStyle := fontStyle + 'text-shadow: -1px 1px 2px #000, 1px 1px 2px #000, 1px -1px 0px #000, -1px -1px 0px #000;color:white;';
       //'text-shadow: -1px 1px 2px #000, 1px 2px 6px #000, 1px -1px 0px #000, -1px -1px 0px #000;',  // https://www.educative.io/answers/how-to-create-a-text-outline-using-css

  case FActiveCharFormat.Position of
    rcpNormal: position := 'vertical-align:baseline;';
    rcpSubscript: position := 'vertical-align:sub;';
    rcpSuperscript: position := 'vertical-align:super;';
  end;

  case FActiveCharFormat.CapsStyle of
    rcsNormal: caps := '';
    rcsCaps: caps := 'text-transform:uppercase';
    rcsSmallCaps: caps := 'font-variant:small-caps';
  end;

  if (FAnsiText <> '') then
    FCurrText := FCurrText + ConvertEncoding(FAnsiText, FCodePage, encodingUTF8);

  Result := Format('<span style="%s%s%s%s%s%s%s">', [fontFamily, fontSize, fontColor, fontStyle, caps, bkColor, position]) +
    FCurrText + '</span>';

  // Reset parser text collecting status
  FTextState := FTextState - [rtsPreparingParFormat, rtsPreparingCharFormat];

  // Reset buffer strings
  FCurrText := '';
  FAnsiText := '';
end;

procedure TRtf2HtmlConverter.WriteHtmlFooter;
begin
  if PreparingTable then
    WriteTable;
  FOutput.Add('  </body>');
  FOutput.Add('</html>');
end;

procedure TRtf2HtmlConverter.WriteHtmlHeader;
begin
  FOutput.Add('<html>');
  FOutput.Add('  <head>');
  FOutput.Add('    <meta charset="utf-8">');
  FOutput.Add('    <title>' + FTitle + '</title>');
  FOutput.Add('  </head>');
  FOutput.Add('  <body>');
  FIndentLevel := 1;      // two spaces per indent level
end;

{ Is called when all paragraph and character formats and the entire text of
  a paragraph is available (i.e. when the next paragraph begins).
  Writes the html for the paragraph. }
procedure TRtf2HtmlConverter.WriteParagraph;
const
  ALIGNMENT_NAMES: array[TRtfAlignment] of string = ('left', 'center', 'right', 'justify');
var
  align: String = '';
  leftindent: String = '';
  rightindent: String = '';
  firstindent: String = '';
  spaceBefore: String = '';
  spaceAfter: String = '';
  spaceBetween: String = '';
  par_or_cell: String;
begin
  // Handle empty lines
  if (FCurrParText = '') and (FCurrText = '') and (FAnsiText = '') then
    FCurrParText := '<br />';

  align := Format('text-align:%s;', [ALIGNMENT_NAMES[FActiveParFormat.Alignment]]);
  leftindent := Format('margin-left:%.1fpt;', [TwipsToPts(FActiveParFormat.LeftIndent)], FFormatSettings);
  rightindent := Format('margin-right:%.1fpt;', [TwipsToPts(FActiveParFormat.RightIndent)], FFormatSettings);
  firstindent := Format('text-indent:%.1fpt;', [TwipsToPts(FActiveParFormat.FirstIndent)], FFormatSettings);
  spaceBefore := Format('margin-top:%.1fpt;', [TwipsToPts(FActiveParFormat.SpacingBefore)], FFormatSettings);
  spaceAfter := Format('margin-bottom:%.1fpt;', [TwipsToPts(FActiveParFormat.SpacingAfter)], FFormatSettings);
  if FActiveParFormat.SpacingBetween > 0 then
    spaceBetween := Format('line-height:%.1fpt;', [TwipsToPts(FActiveParFormat.SpacingBetween)], FFormatSettings);

  // Write the <p> or <tc> tag
  if PreparingTable then
    par_or_cell := 'tc'       // table cell
  else
    par_or_cell := 'p';       // normal paragraph

  Add(Format('<%s style="%s%s%s%s%s%s%s">', [
    par_or_cell,
    align, leftindent, rightindent, firstindent,
    spaceBefore, spaceAfter, spaceBetween
  ]));

  // Write the (formatted) text
  inc(FIndentLevel);
  Add(FCurrParText);

  // Write the closing </p> or </tc> tag
  dec(FIndentLevel);
  Add('</' + par_or_cell + '>');

  // Reset parser text collecting status
  FTextState := FTextState - [rtsPreparingParFormat, rtsPreparingCharFormat];

  // Reset buffer strings
  FCurrParText := '';
  FAnsiText := '';
  FCurrText := '';
end;

procedure TRtf2HtmlConverter.WritePicture;
var
  mimeStr: String = '';
  dataStr: String = '';
  i, j: Integer;
  value: byte;
  s: String[3] = '$xy';    // x, y will be replaced by PictData characters
  ppi: Integer;
  wPx, hPx: Integer;
begin
  if FPictData = '' then
    exit;

  j := 1;
  SetLength(dataStr, FPictDataCount div 2);
  for i := 1 to High(dataStr) do
  begin
    s[2] := FPictData[j];
    s[3] := FPictData[j+1];
    value := StrToInt(s);
    dataStr[i] := chr(value);
    inc(j, 2);
  end;

  case FActivePictAttrib.Format of
    rpfPNG:
      mimeStr := 'image/png';
    rpfJPEG:
      mimeStr := 'image/jpeg';
    {         // not working...
    rpfWMF:
      mimeStr := 'image/wmf';
      }
    else
      if copy(dataStr, 1, 3) = 'GIF' then
        mimeStr := 'image/gif'                 // not displayed by Word
      else
      if copy(dataStr, 1, 2) = 'BM' then
        mimeStr := 'image/bmp'                 // untested
      else
        exit;
  end;

  ppi := FPixelsPerInch;

  if FActivePictAttrib.GoalWidth <> 0 then
    wPx := TwipsToPx(FActivePictAttrib.GoalWidth, FPixelsPerInch)
  else
  begin
    wPx := TwipsToPx(FActivePictAttrib.Width, FPixelsPerInch);
    wPx := wPx * FActivePictAttrib.ScaleX div 100;
  end;

  if FActivePictAttrib.GoalHeight <> 0 then
    hPx := TwipsToPx(FActivePictAttrib.GoalHeight, FPixelsPerInch)
  else
  begin
    hPx := TwipsToPx(FActivePictAttrib.Height, FPixelsPerInch);
    hPx := hPx * FActivePictAttrib.ScaleY div 100;
  end;

  FCurrParText += Format('<img src="data:%s;base64,%s" width="%d" height="%d">',
    [mimeStr, EncodeStringBase64(dataStr), wPx, hPx]
  );

  FPictData := '';
end;

procedure TRtf2HtmlConverter.WriteTable;
var
  i: Integer;
begin
  Add('<table style="border-collapse: collapse;">');
  inc(FIndentLevel);
  for i := 0 to High(FTblRows) do
    Add(FTblRows[i]);
  dec(FIndentLevel);
  Add('</table>');
  FTextState := FTextState - [rtsWaitingForNextRow, rtsPreparingTable];
end;

procedure TRtf2HtmlConverter.WriteTableRow;
var
  rowStr: String;
  widthStr: String;
  borderStr: String = '';
  marginStr: String = '';
  c: Integer;
  colWid: array of Integer = nil;
  border: TRtfCellBorderStyle;
  b: TRtfCellBorder;
  n: Integer;
begin
  // Calculate the colum widths
  n := Length(FTblAttrib.ColPos);
  SetLength(colWid, n);
  if n > 0 then
  begin
    colWid[0] := FTblAttrib.ColPos[0];
    for c := 1 to n-1 do
      colWid[c] := FTblAttrib.ColPos[c] - FTblAttrib.ColPos[c-1];
  end;
  if n < Length(FTblAttrib.CellText) then
  begin
    SetLength(colWid, Length(FTblAttrib.CellText));
    for c := n to High(colWid) do
      colWid[c] := 2*1440;
  end;

  // Collapse all borders (all cells, left/top/right/bottom borders) into a single style
  border := rcbsNone;
  for c := 0 to High(FTblAttrib.CellBorders) do
  begin
    for b in TRtfCellBorder do
      if FTblAttrib.CellBorders[c, b] <> rcbsNone then
      begin
        border := FTblAttrib.CellBorders[c, b];
        break;
      end;
    if border <> rcbsNone then break;
  end;

  marginStr := Format('padding-left:%0:.1fmm;padding-right:%0:.1fmm;', [TwipsToMM(FTblAttrib.HorPadding)], FFormatSettings);

  // Create html text
  rowStr := '<tr>'; //Format('%s<tr style="%s">', [Indentation, marginStr]);
  for c := 0 to High(FTblAttrib.CellText) do
  begin
    widthStr := Format('width:%.1fmm;', [TwipsToMM(colWid[c])], FFormatSettings);
    case border of
      rcbsSingle:
        borderStr := 'border:1px solid black;border-collapse:collapse;';
      else
        borderStr := '';
    end;
    rowStr := rowStr + Format('<td style="%s%s%s">%s</td>', [
      widthStr, borderStr, marginStr, FTblAttrib.CellText[c]
    ]);
  end;
  rowstr := rowStr + '</tr>';

  n := Length(FTblRows);
  SetLength(FTblRows, n+1);
  FTblRows[n] := rowStr;

  FCurrText := '';
  FCurrParText := '';
  FAnsiText := '';
  FTextState := FTextState + [rtsWaitingForNextRow];
  // When the next tag is not \tblrowd and this flag is still set we have
  // reached the end of the table.
end;

end.

