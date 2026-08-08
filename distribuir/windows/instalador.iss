; Instalador de um clique para Windows (Inno Setup 6).
;
;     iscc /DVersao=0.1.0 distribuir\windows\instalador.iss
;
; Sai "Instalar Ferramenta NPD.exe": a pessoa baixa, dá dois cliques, avança
; e pronto — atalho na área de trabalho e no menu Iniciar.
;
; Três decisões que valem explicação:
;
; 1. PrivilegesRequired=lowest. A instalação vai para a pasta do próprio
;    usuário (%LOCALAPPDATA%\Programs), não para "Arquivos de Programas". Isso
;    dispensa senha de administrador — que quem usa a ferramenta muitas vezes
;    não tem, e pedir senha é onde a instalação "de um clique" morre.
;
; 2. Nada de associação de arquivo nem de serviço em segundo plano. O programa
;    é uma janela que a pessoa abre quando quer; instalar coisa que fica
;    rodando sozinha é abuso de confiança.
;
; 3. O idioma é português do Brasil, e só. Um instalador que abre em inglês
;    para quem trabalha em português já começa pedindo desculpas.

#ifndef Versao
  #define Versao "0.0.0"
#endif

#define Nome "Ferramenta NPD"

[Setup]
AppId={{8F2B41E7-9C3A-4D66-B0F1-2C7A5E9D4A10}
AppName={#Nome}
AppVersion={#Versao}
AppVerName={#Nome} {#Versao}
AppPublisher=Marchesoni
DefaultDirName={autopf}\{#Nome}
DefaultGroupName={#Nome}
DisableProgramGroupPage=yes
DisableDirPage=auto
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\..\dist-instalador
; Sem espaço no nome: o GitHub troca espaço por ponto ao anexar o arquivo na
; release, e "Instalar.Ferramenta.NPD.exe" chega parecendo arquivo corrompido.
OutputBaseFilename=Instalar-Ferramenta-NPD
SetupIconFile=..\icone\icone.ico
UninstallDisplayIcon={app}\{#Nome}.exe
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
; A janela do programa precisa do WebView2, que já vem no Windows 10 e 11
; atualizados. Sem ele a tela abre no navegador padrão — funciona igual, muda
; onde aparece —, então não é motivo para barrar a instalação.
MinVersion=10.0

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "atalhodesktop"; Description: "Criar um atalho na área de trabalho"; GroupDescription: "Atalhos:"

[Files]
Source: "..\..\dist-app\{#Nome}.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\COMO-USAR.md"; DestDir: "{app}"; DestName: "Como usar.txt"; Flags: ignoreversion isreadme

[Icons]
Name: "{group}\{#Nome}"; Filename: "{app}\{#Nome}.exe"
Name: "{userdesktop}\{#Nome}"; Filename: "{app}\{#Nome}.exe"; Tasks: atalhodesktop

[Run]
Filename: "{app}\{#Nome}.exe"; Description: "Abrir a {#Nome} agora"; Flags: nowait postinstall skipifsilent
