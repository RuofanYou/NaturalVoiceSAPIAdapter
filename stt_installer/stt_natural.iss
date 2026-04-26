; STT 自然语音 v7 - 单文件安装器（NaturalVoiceSAPIAdapter + Microsoft Xiaoxiao 嵌入式神经语音）
; 编码: UTF-8 with BOM (Inno Setup 6 要求中文 .iss 必须带 BOM 或用 ANSI)

[Setup]
AppId={{8F3A1C2D-9B5E-4F6A-8D2E-3C7B5A9E6F4B}
AppName=STT 自然语音 (NaturalVoice)
AppVersion=0.7.0
AppPublisher=STT Voice Pack Team
AppPublisherURL=https://github.com/RuofanYou/NaturalVoiceSAPIAdapter

DefaultDirName=C:\STT-NaturalVoice-VoicePack
DisableDirPage=no
UsePreviousAppDir=no
DefaultGroupName=STT 自然语音
DisableProgramGroupPage=yes
UninstallDisplayName=STT 自然语音 v7
UninstallDisplayIcon={app}\NaturalVoiceSAPIAdapter\x64\NaturalVoiceSAPIAdapter.dll

PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

Compression=lzma2/ultra64
SolidCompression=yes

OutputDir=output
OutputBaseFilename=STT-NaturalVoice-Setup-v7

WizardStyle=modern
ShowLanguageDialog=no

[Languages]
Name: "chs"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Files]
; 修改版 NaturalVoiceSAPIAdapter（CI 编译产物 + 运行时 DLL）
Source: "payload\NaturalVoiceSAPIAdapter\x64\*"; DestDir: "{app}\NaturalVoiceSAPIAdapter\x64"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "payload\NaturalVoiceSAPIAdapter\x86\*"; DestDir: "{app}\NaturalVoiceSAPIAdapter\x86"; Flags: ignoreversion recursesubdirs createallsubdirs
; 微软 Xiaoxiao 嵌入式神经语音模型（解压自 MSIX）
Source: "payload\TTS_VOICE\*"; DestDir: "{app}\TTS_VOICE"; Flags: ignoreversion recursesubdirs createallsubdirs

[Run]
Filename: "{sys}\regsvr32.exe"; Parameters: "/s ""{app}\NaturalVoiceSAPIAdapter\x64\NaturalVoiceSAPIAdapter.dll"""; Flags: runascurrentuser waituntilterminated; StatusMsg: "正在注册 64 位 SAPI5 引擎..."
Filename: "{sys}\..\SysWOW64\regsvr32.exe"; Parameters: "/s ""{app}\NaturalVoiceSAPIAdapter\x86\NaturalVoiceSAPIAdapter.dll"""; Flags: runascurrentuser waituntilterminated; StatusMsg: "正在注册 32 位 SAPI5 引擎..."

[UninstallRun]
Filename: "{sys}\..\SysWOW64\regsvr32.exe"; Parameters: "/u /s ""{app}\NaturalVoiceSAPIAdapter\x86\NaturalVoiceSAPIAdapter.dll"""; Flags: runascurrentuser waituntilterminated; RunOnceId: "UnregX86"
Filename: "{sys}\regsvr32.exe"; Parameters: "/u /s ""{app}\NaturalVoiceSAPIAdapter\x64\NaturalVoiceSAPIAdapter.dll"""; Flags: runascurrentuser waituntilterminated; RunOnceId: "UnregX64"

[Registry]
; 写入 NarratorVoicePath 让 NaturalVoiceSAPIAdapter 扫描我们打包的 Xiaoxiao 离线神经语音
Root: HKCU; Subkey: "Software\NaturalVoiceSAPIAdapter\Enumerator"; ValueType: string; ValueName: "NarratorVoicePath"; ValueData: "{app}\TTS_VOICE"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\NaturalVoiceSAPIAdapter\Enumerator"; ValueType: dword; ValueName: "NoNarratorVoices"; ValueData: 0; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\NaturalVoiceSAPIAdapter\Enumerator"; ValueType: dword; ValueName: "NoEdgeVoices"; ValueData: 0; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\NaturalVoiceSAPIAdapter\Enumerator"; ValueType: dword; ValueName: "NoAzureVoices"; ValueData: 1; Flags: uninsdeletevalue

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
const
  XIAO_YA_CLSID = '{{5C7A9D4A-3B6F-4E18-9A2D-7B1E2F8C6A31}';
  STT_LEGACY_TOKEN = 'STT_XIAO_YA_ZH_CN';

// ---- 字符串小工具 ----
function ContainsText(const Hay, Needle: string): Boolean;
begin
  Result := Pos(LowerCase(Needle), LowerCase(Hay)) > 0;
end;

// ---- regsvr32 卸载（吃静默失败）----
procedure SilentUnregister(const DllPath, RegsvrPath: string);
var
  Code: Integer;
begin
  if FileExists(DllPath) then
    Exec(RegsvrPath, '/u /s "' + DllPath + '"', '', SW_HIDE, ewWaitUntilTerminated, Code);
end;

// ---- 调旧 Inno Setup uninstaller（如果有，静默运行）----
procedure CallLegacyUninstaller(const RegRoot: Integer; const SubKey: string);
var
  S, ExePath: string;
  Code: Integer;
begin
  if RegQueryStringValue(RegRoot, SubKey, 'UninstallString', S) then
  begin
    // S 形如 "C:\xxx\unins000.exe" /SILENT 之类，先剥引号
    if (Length(S) > 0) and (S[1] = '"') then
    begin
      ExePath := Copy(S, 2, Pos('"', Copy(S, 2, Length(S))) - 1);
    end
    else
    begin
      ExePath := S;
    end;
    if FileExists(ExePath) then
      Exec(ExePath, '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART', '', SW_HIDE, ewWaitUntilTerminated, Code);
  end;
end;

// ---- 从 PATH 字符串里删掉所有含 needle 的段 ----
function StripPathSegments(const PathStr, Needle: string): string;
var
  Remain, Seg, OutStr: string;
  P: Integer;
begin
  Remain := PathStr;
  OutStr := '';
  while Length(Remain) > 0 do
  begin
    P := Pos(';', Remain);
    if P = 0 then
    begin
      Seg := Remain;
      Remain := '';
    end
    else
    begin
      Seg := Copy(Remain, 1, P - 1);
      Remain := Copy(Remain, P + 1, Length(Remain));
    end;
    if (Length(Seg) > 0) and (not ContainsText(Seg, Needle)) then
    begin
      if Length(OutStr) > 0 then
        OutStr := OutStr + ';';
      OutStr := OutStr + Seg;
    end;
  end;
  Result := OutStr;
end;

// ---- 系统 PATH 清理 ----
procedure CleanSystemPath();
var
  EnvPath, NewPath: string;
begin
  if RegQueryStringValue(HKLM, 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment', 'Path', EnvPath) then
  begin
    NewPath := StripPathSegments(EnvPath, 'STT-xiao_ya');
    NewPath := StripPathSegments(NewPath, 'STT-NaturalVoice');
    if NewPath <> EnvPath then
      RegWriteExpandStringValue(HKLM, 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment', 'Path', NewPath);
  end;
end;

// ---- 多路径候选 regsvr32 /u ----
procedure UnregisterCandidateDlls();
var
  Sys64, Sys32: string;
  Candidates: array of string;
  i: Integer;
begin
  Sys64 := ExpandConstant('{sys}');
  Sys32 := ExpandConstant('{sys}\..\SysWOW64');

  SetArrayLength(Candidates, 10);
  Candidates[0] := 'C:\STT-xiao_ya-VoicePack\stt_xiao_ya_sapi5.dll';
  Candidates[1] := 'C:\STT-xiao_ya-VoicePack\bin\stt_xiao_ya_sapi5.dll';
  Candidates[2] := ExpandConstant('{pf}\STT-xiao_ya-VoicePack\stt_xiao_ya_sapi5.dll');
  Candidates[3] := ExpandConstant('{pf32}\STT-xiao_ya-VoicePack\stt_xiao_ya_sapi5.dll');
  Candidates[4] := 'C:\STT-NaturalVoice-VoicePack\NaturalVoiceSAPIAdapter\x64\NaturalVoiceSAPIAdapter.dll';
  Candidates[5] := 'C:\STT-NaturalVoice-VoicePack\NaturalVoiceSAPIAdapter\x86\NaturalVoiceSAPIAdapter.dll';
  Candidates[6] := 'C:\STT-NaturalVoice-VoicePack\x64\NaturalVoiceSAPIAdapter.dll';
  Candidates[7] := 'C:\STT-NaturalVoice-VoicePack\x86\NaturalVoiceSAPIAdapter.dll';
  Candidates[8] := ExpandConstant('{pf}\NaturalVoiceSAPIAdapter\x64\NaturalVoiceSAPIAdapter.dll');
  Candidates[9] := ExpandConstant('{pf32}\NaturalVoiceSAPIAdapter\x86\NaturalVoiceSAPIAdapter.dll');

  for i := 0 to GetArrayLength(Candidates) - 1 do
  begin
    if ContainsText(Candidates[i], 'x86') or ContainsText(Candidates[i], 'SysWOW64') then
      SilentUnregister(Candidates[i], Sys32 + '\regsvr32.exe')
    else
      SilentUnregister(Candidates[i], Sys64 + '\regsvr32.exe');
    // 再用另一边也试一次（防错位）
    if ContainsText(Candidates[i], 'x86') then
      SilentUnregister(Candidates[i], Sys64 + '\regsvr32.exe')
    else
      SilentUnregister(Candidates[i], Sys32 + '\regsvr32.exe');
  end;
end;

// ---- 总清理过程 ----
procedure CleanupAllHistoricalVersions();
begin
  // 1. regsvr32 /u 多路径候选 DLL
  UnregisterCandidateDlls();

  // 2. 删 SAPI5 voice token：自研 STT_XIAO_YA_ZH_CN
  RegDeleteKeyIncludingSubkeys(HKLM, 'SOFTWARE\Microsoft\Speech\Voices\Tokens\' + STT_LEGACY_TOKEN);
  RegDeleteKeyIncludingSubkeys(HKLM, 'SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens\' + STT_LEGACY_TOKEN);
  RegDeleteKeyIncludingSubkeys(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\Speech\Voices\Tokens\' + STT_LEGACY_TOKEN);
  RegDeleteKeyIncludingSubkeys(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\Speech_OneCore\Voices\Tokens\' + STT_LEGACY_TOKEN);

  // 3. 删 COM CLSID（自研 SAPI5 引擎）
  RegDeleteKeyIncludingSubkeys(HKLM, 'SOFTWARE\Classes\CLSID\' + XIAO_YA_CLSID);
  RegDeleteKeyIncludingSubkeys(HKLM, 'SOFTWARE\Classes\WOW6432Node\CLSID\' + XIAO_YA_CLSID);

  // 4. 调旧 Inno Setup uninstaller（v6 旧 AppId 的 Uninstall 注册项）
  // 旧 v6 用的 AppId 假设含 5C7A9D4A...，列出常见两条，能调到就调
  CallLegacyUninstaller(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{5C7A9D4A-3B6F-4E18-9A2D-7B1E2F8C6A31}_is1');
  CallLegacyUninstaller(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{5C7A9D4A-3B6F-4E18-9A2D-7B1E2F8C6A31}_is1');

  // 5. 清系统 PATH
  CleanSystemPath();

  // 6. 删 NaturalVoiceSAPIAdapter 之前 zip 装过的注册表配置（用户级）
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\NaturalVoiceSAPIAdapter');

  // 7. 删旧目录（失败也吃）
  DelTree('C:\STT-xiao_ya-VoicePack', True, True, True);
  DelTree(ExpandConstant('{pf}\STT-xiao_ya-VoicePack'), True, True, True);
  DelTree(ExpandConstant('{pf32}\STT-xiao_ya-VoicePack'), True, True, True);
  DelTree('C:\STT-NaturalVoice-VoicePack', True, True, True);
end;

// ---- 启动菜单 ----
function InitializeSetup(): Boolean;
var
  Mode: Integer;
begin
  Result := True;
  Mode := MsgBox(
    'STT 自然语音 v7 - 请选择操作：' + #13#10 + #13#10 +
    '  [是]   安装新版（自动卸载所有历史版本，全新清理）' + #13#10 +
    '  [否]   仅卸载所有历史版本，不安装新版' + #13#10 +
    '  [取消] 退出',
    mbConfirmation, MB_YESNOCANCEL);

  if Mode = IDCANCEL then
  begin
    Result := False;
    Exit;
  end;

  // 不论 YES 还是 NO，先清理历史污染
  CleanupAllHistoricalVersions();

  if Mode = IDNO then
  begin
    MsgBox('已清理所有历史版本（自研 sherpa-onnx 残留 + NaturalVoiceSAPIAdapter 旧装）。'#13#10'未安装新版。', mbInformation, MB_OK);
    Result := False;
    Exit;
  end;

  // 走 Inno Setup 标准安装流程
  Result := True;
end;
