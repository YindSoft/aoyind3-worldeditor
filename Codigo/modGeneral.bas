Attribute VB_Name = "modGeneral"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

''
' modGeneral
'
' @remarks Funciones Generales
' @author unkwown
' @version 0.4.11
' @date 20061015

Option Explicit

Public Type typDevMODE
    dmDeviceName       As String * 32
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * 32
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type
Private Type tNpcs
    NPCIndex As Integer
    Cantidad As Integer
End Type
Private Type tArea
    Mapa As Byte
    X1 As Integer
    Y1 As Integer
    X2 As Integer
    Y2 As Integer
    Npcs As Byte
    NPC() As tNpcs
End Type

Private Type tZona
    Nombre As String
    Mapa As Byte
    X1 As Integer
    Y1 As Integer
    X2 As Integer
    Y2 As Integer
    Segura As Byte
    Acoplar As Byte
    Terreno As Byte
    Musica(1 To 5) As Byte
End Type
Public Zonas() As tZona
Public Areas() As tArea
Public NumZonas As Integer
Public NumAreas As Integer

Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_DISPLAYFREQUENCY = &H400000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

''
' Realiza acciones de desplasamiento segun las teclas que hallamos precionado
'

Public Sub CheckKeys()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
If Not PuedeMover Then Exit Sub
'If HotKeysAllow = False Then Exit Sub
        '[Loopzer]
        If GetKeyState(vbKeyControl) < 0 Then
            If Seleccionando Then
                If GetKeyState(vbKeyC) < 0 Then CopiarSeleccion
                If GetKeyState(vbKeyX) < 0 Then CortarSeleccion
                If GetKeyState(vbKeyB) < 0 Then BlockearSeleccion
                If GetKeyState(vbKeyD) < 0 Then AccionSeleccion
            Else
                If GetKeyState(vbKeyS) < 0 Then DePegar ' GS
                If GetKeyState(vbKeyV) < 0 Then PegarSeleccion
            End If
        End If
        '[/Loopzer]
        
        
If GetKeyState(vbKeyUp) < 0 Then
        If UserPos.Y < 10 Then Exit Sub ' 10
        If LegalPos(UserPos.X, UserPos.Y - 1) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.Y = UserPos.Y - 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.Y = UserPos.Y - 1
        End If
        bRefreshRadar = True ' Radar
        frmMain.SetFocus
        Exit Sub
    End If

    If GetKeyState(vbKeyRight) < 0 Then
        If UserPos.X > 1089 Then Exit Sub ' 89
        If LegalPos(UserPos.X + 1, UserPos.Y) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.X = UserPos.X + 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X + 1
        End If
        bRefreshRadar = True ' Radar
        frmMain.SetFocus
        Exit Sub
    End If

    If GetKeyState(vbKeyDown) < 0 Then
        If UserPos.Y > 1492 Then Exit Sub ' 92
        If LegalPos(UserPos.X, UserPos.Y + 1) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.Y = UserPos.Y + 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.Y = UserPos.Y + 1
        End If
        bRefreshRadar = True ' Radar
        frmMain.SetFocus
        Exit Sub
    End If

    If GetKeyState(vbKeyLeft) < 0 Then
        If UserPos.X < 12 Then Exit Sub ' 12
        If LegalPos(UserPos.X - 1, UserPos.Y) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.X = UserPos.X - 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X - 1
        End If
        bRefreshRadar = True ' Radar
        frmMain.SetFocus
        Exit Sub
    End If
    

End Sub

Public Function ReadField(Pos As Integer, text As String, SepASCII As Integer) As String
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(text)
    CurChar = mid(text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = mid(text, LastPos + 1, (InStr(LastPos + 1, text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i
FieldNum = FieldNum + 1

If FieldNum = Pos Then
    ReadField = mid(text, LastPos + 1)
End If

End Function


''
' Completa y corrije un path
'
' @param Path Especifica el path con el que se trabajara
' @return   Nos devuelve el path completado

Private Function autoCompletaPath(ByVal Path As String) As String
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
Path = Replace(Path, "/", "\")
If Left(Path, 1) = "\" Then
    ' agrego app.path & path
    Path = App.Path & Path
End If
If Right(Path, 1) <> "\" Then
    ' me aseguro que el final sea con "\"
    Path = Path & "\"
End If
autoCompletaPath = Path
End Function

''
' Carga la configuracion del WorldEditor de WorldEditor.ini
'

Private Sub CargarMapIni()
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
On Error GoTo Fallo
Dim tStr As String
Dim Leer As New clsIniReader

IniPath = App.Path & "\"

If FileExist(IniPath & "WorldEditor.ini", vbArchive) = False Then
    frmMain.mnuGuardarUltimaConfig.Checked = True
    DirGraficos = IniPath & "Graficos\"
    DirIndex = IniPath & "INIT\"
    DirMidi = IniPath & "MIDI\"
    frmMusica.fleMusicas.Path = DirMidi
    DirDats = IniPath & "DAT\"
    MaxGrhs = 15000
    UserPos.X = 50
    UserPos.Y = 50
    PantallaX = 19
    PantallaY = 22
    MsgBox "Falta el archivo 'WorldEditor.ini' de configuración.", vbInformation
    Exit Sub
End If

Call Leer.Initialize(IniPath & "WorldEditor.ini")

' Obj de Translado
Cfg_TrOBJ = Val(Leer.GetValue("CONFIGURACION", "ObjTranslado"))
frmMain.mnuAutoCapturarTranslados.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarTrans"))
frmMain.mnuAutoCapturarSuperficie.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarSup"))

' Guardar Ultima Configuracion
frmMain.mnuGuardarUltimaConfig.Checked = Val(Leer.GetValue("CONFIGURACION", "GuardarConfig"))

' Index
MaxGrhs = Val(GetVar(IniPath & "WorldEditor.ini", "INDEX", "MaxGrhs"))
If MaxGrhs < 1 Then MaxGrhs = 15000

'Reciente
'frmMain.Dialog.InitDir = Leer.GetValue("PATH", "UltimoMapa")
DirGraficos = autoCompletaPath(Leer.GetValue("PATH", "DirGraficos"))
If DirGraficos = "\" Then
    DirGraficos = IniPath & "Graficos\"
End If
If FileExist(DirGraficos, vbDirectory) = False Then
    MsgBox "El directorio de Graficos es incorrecto", vbCritical + vbOKOnly
    End
End If
DirMidi = autoCompletaPath(Leer.GetValue("PATH", "DirMidi"))
If DirMidi = "\" Then
    DirMidi = IniPath & "MIDI\"
End If
If FileExist(DirMidi, vbDirectory) = False Then
    MsgBox "El directorio de MIDI es incorrecto", vbCritical + vbOKOnly
    End
End If
frmMusica.fleMusicas.Path = DirMidi
DirIndex = autoCompletaPath(Leer.GetValue("PATH", "DirIndex"))
If DirIndex = "\" Then
    DirIndex = IniPath & "INIT\"
End If
If FileExist(DirIndex, vbDirectory) = False Then
    MsgBox "El directorio de Index es incorrecto", vbCritical + vbOKOnly
    End
End If
DirDats = autoCompletaPath(Leer.GetValue("PATH", "DirDats"))
If DirDats = "\" Then
    DirDats = IniPath & "DAT\"
End If
If FileExist(DirDats, vbDirectory) = False Then
    MsgBox "El directorio de Dats es incorrecto", vbCritical + vbOKOnly
    End
End If

tStr = Leer.GetValue("MOSTRAR", "LastPos") ' x-y
UserPos.X = Val(ReadField(1, tStr, Asc("-")))
UserPos.Y = Val(ReadField(2, tStr, Asc("-")))
If UserPos.X < XMinMapSize Or UserPos.X > XMaxMapSize Then
    UserPos.X = 50
End If
If UserPos.Y < YMinMapSize Or UserPos.Y > YMaxMapSize Then
    UserPos.Y = 50
End If

' Menu Mostrar
frmMain.mnuVerAutomatico.Checked = Val(Leer.GetValue("MOSTRAR", "ControlAutomatico"))
frmMain.mnuVerCapa2.Checked = Val(Leer.GetValue("MOSTRAR", "Capa2"))
frmMain.mnuVerCapa3.Checked = Val(Leer.GetValue("MOSTRAR", "Capa3"))
frmMain.mnuVerCapa4.Checked = Val(Leer.GetValue("MOSTRAR", "Capa4"))
frmMain.mnuVerTranslados.Checked = Val(Leer.GetValue("MOSTRAR", "Translados"))
frmMain.mnuVerObjetos.Checked = Val(Leer.GetValue("MOSTRAR", "Objetos"))
frmMain.mnuVerNPCs.Checked = Val(Leer.GetValue("MOSTRAR", "NPCs"))
frmMain.mnuVerTriggers.Checked = Val(Leer.GetValue("MOSTRAR", "Triggers"))
frmMain.mmTrans.Checked = Val(Leer.GetValue("MOSTRAR", "Trans"))
frmMain.mnuVerGrilla.Checked = Val(Leer.GetValue("MOSTRAR", "Grilla")) ' Grilla
VerGrilla = frmMain.mnuVerGrilla.Checked
frmMain.mnuVerBloqueos.Checked = Val(Leer.GetValue("MOSTRAR", "Bloqueos"))
frmMain.cVerTriggers.value = frmMain.mnuVerTriggers.Checked
frmMain.cVerBloqueos.value = frmMain.mnuVerBloqueos.Checked

' Tamaño de visualizacion
PantallaX = Val(Leer.GetValue("MOSTRAR", "PantallaX"))
PantallaY = Val(Leer.GetValue("MOSTRAR", "PantallaY"))
If PantallaX > 23 Or PantallaX <= 2 Then PantallaX = 23
If PantallaY > 32 Or PantallaY <= 2 Then PantallaY = 32

' [GS] 02/10/06
' Tamaño de visualizacion en el cliente
ClienteHeight = Val(Leer.GetValue("MOSTRAR", "ClienteHeight"))
ClienteWidth = Val(Leer.GetValue("MOSTRAR", "ClienteWidth"))
If ClienteHeight <= 0 Then ClienteHeight = 13
If ClienteWidth <= 0 Then ClienteWidth = 17

Exit Sub
Fallo:
    MsgBox "ERROR " & Err.Number & " en WorldEditor.ini" & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Public Function TomarBPP() As Integer
    Dim ModoDeVideo As typDevMODE
    Call EnumDisplaySettings(0, -1, ModoDeVideo)
    TomarBPP = CInt(ModoDeVideo.dmBitsPerPel)
End Function
Public Sub CambioDeVideo()
'*************************************************
'Author: Loopzer
'*************************************************
Exit Sub
Dim ModoDeVideo As typDevMODE
Dim r As Long
Call EnumDisplaySettings(0, -1, ModoDeVideo)
    If ModoDeVideo.dmPelsWidth < 1024 Or ModoDeVideo.dmPelsHeight < 768 Then
        Select Case MsgBox("La aplicacion necesita una resolucion minima de 1024 X 768 ,¿Acepta el Cambio de resolucion?", vbInformation + vbOKCancel, "World Editor")
            Case vbOK
                ModoDeVideo.dmPelsWidth = 1024
                ModoDeVideo.dmPelsHeight = 768
                ModoDeVideo.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
                r = ChangeDisplaySettings(ModoDeVideo, CDS_TEST)
                If r <> 0 Then
                    MsgBox "Error al cambiar la resolucion, La aplicacion se cerrara."
                    End
                End If
            Case vbCancel
                End
        End Select
    End If
End Sub

Public Sub Main()
'*************************************************
'Author: Unkwown
'Last modified: 25/11/08 - GS
'*************************************************
On Error Resume Next
If App.PrevInstance = True Then End
UserMap = 1
Call CargarMapIni
Call IniciarCabecera(MiCabecera)

If Year(Now) > 2012 Or GetSetting("MAP", "EDITOR", "START", 1) = 5 Then
'    Call SaveSetting("MAP", "EDITOR", "START", "5")
'    End
End If

If FileExist(IniPath & "WorldEditor.jpg", vbArchive) Then frmCargando.picture1.Picture = LoadPicture(IniPath & "WorldEditor.jpg")
frmCargando.verX = "v" & App.Major & "." & App.Minor & "." & App.Revision
frmCargando.Show
frmCargando.SetFocus
DoEvents
frmCargando.X.Caption = "Iniciando DirectSound..."

DoEvents
frmCargando.X.Caption = "Cargando Indice de Superficies..."
modIndices.CargarIndicesSuperficie
DoEvents
frmCargando.X.Caption = "Indexando Cargado de Imagenes..."
LoadGrhData
CargarParticulas
CargarFxs
CargarCuerpos
DoEvents
If FileExist(DirIndex & "AO.dat", vbArchive) Then
    Call LoadClientSetup


End If

'If InitTileEngine(frmMain.hwnd, frmMain.MainViewShp.Top + 47, frmMain.MainViewShp.Left + 4, 32, 32, PantallaX, PantallaY, 9) Then ' 30/05/2006
    'Display form handle, View window offset from 0,0 of display form, Tile Size, Display size in tiles, Screen buffer
    frmCargando.P1.Visible = True
    frmCargando.L(0).Visible = True
    frmCargando.X.Caption = "Cargando Cuerpos..."
   CargarCuerpos
    DoEvents
    frmCargando.P2.Visible = True
    frmCargando.L(1).Visible = True
    frmCargando.X.Caption = "Cargando Cabezas..."
  CargarCabezas
    DoEvents
    frmCargando.P3.Visible = True
    frmCargando.L(2).Visible = True
    frmCargando.X.Caption = "Cargando NPC's..."
    modIndices.CargarIndicesNPC
    DoEvents
    frmCargando.P4.Visible = True
    frmCargando.L(3).Visible = True
    frmCargando.X.Caption = "Cargando Objetos..."
    modIndices.CargarIndicesOBJ
    DoEvents
    frmCargando.P5.Visible = True
    frmCargando.L(4).Visible = True
    frmCargando.X.Caption = "Cargando Triggers..."
    modIndices.CargarIndicesTriggers
    DoEvents
    frmCargando.P6.Visible = True
    frmCargando.L(5).Visible = True
    DoEvents
'End If
frmCargando.SetFocus
frmCargando.X.Caption = "Iniciando Ventana de Edición..."
DoEvents
If LenB(Dir(App.Path & "\manual\index.html", vbArchive)) = 0 Then
    frmMain.mnuManual.Enabled = False
    frmMain.mnuManual.Caption = "&Manual (no implementado)"
End If


frmCargando.Hide
frmMain.Show
frmParticle.Show , frmMain
frmParticle.Visible = False
'modMapIO.NuevoMapa
DoEvents
engine.Engine_Init
prgRun = True

CargarZonas
CargarAreas


engine.Font_Create "Tahoma", 8, False, False

'MouseParticle = General_Particle_Create(59, -1, -1)
    
engine.Start

End Sub
Sub CargarZonas()
On Error Resume Next
    Dim archivoC As String
    
    archivoC = DirDats & "zonas.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar las zonas. Falta el archivo zonas.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Integer
    Dim e As Integer
    NumZonas = GetVar(archivoC, "Config", "Cantidad")
    
    ReDim Zonas(1 To NumZonas)
    For i = 1 To NumZonas
        Zonas(i).Nombre = GetVar(archivoC, "Zona" & CStr(i), "Nombre")
        Zonas(i).Mapa = CByte(GetVar(archivoC, "Zona" & CStr(i), "Mapa"))
        Zonas(i).X1 = CInt(GetVar(archivoC, "Zona" & CStr(i), "X1"))
        Zonas(i).Y1 = CInt(GetVar(archivoC, "Zona" & CStr(i), "Y1"))
        Zonas(i).X2 = CInt(GetVar(archivoC, "Zona" & CStr(i), "X2"))
        Zonas(i).Y2 = CInt(GetVar(archivoC, "Zona" & CStr(i), "Y2"))
        Zonas(i).Terreno = CByte(Val(GetVar(archivoC, "Zona" & CStr(i), "Terreno")))
        Zonas(i).Segura = CByte(GetVar(archivoC, "Zona" & CStr(i), "Segura"))
        Zonas(i).Acoplar = CByte(Val(GetVar(archivoC, "Zona" & CStr(i), "Acoplar")))
        For e = 1 To 5
         Zonas(i).Musica(e) = Val(GetVar(archivoC, "Zona" & CStr(i), "Musica" & e))
        Next e
       

    Next i
    
frmMain.PintarMapa
End Sub

Sub CargarAreas()
On Error Resume Next
    Dim archivoC As String
    
    archivoC = DirDats & "areas.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar las areas. Falta el archivo zonas.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Integer
    Dim e As Integer
    Dim H As Integer
    
For i = 1 To 1100
    For e = 1 To 1500
        MapData(i, e).NpcArea = 0
    Next e
Next i
    
    NumAreas = GetVar(archivoC, "Config", "Cantidad")
    If NumAreas > 0 Then
    frmMain.picRadar.Cls
    frmMain.picRadar.ForeColor = vbGreen
    ReDim Areas(1 To NumAreas)
    For i = 1 To NumAreas
        Areas(i).Mapa = CByte(GetVar(archivoC, "Area" & CStr(i), "Mapa"))
        Areas(i).X1 = CInt(GetVar(archivoC, "Area" & CStr(i), "X1"))
        Areas(i).Y1 = CInt(GetVar(archivoC, "Area" & CStr(i), "Y1"))
        Areas(i).X2 = CInt(GetVar(archivoC, "Area" & CStr(i), "X2"))
        Areas(i).Y2 = CInt(GetVar(archivoC, "Area" & CStr(i), "Y2"))
        Areas(i).Npcs = CByte(GetVar(archivoC, "Area" & CStr(i), "Npcs"))
        If Areas(i).Npcs > 0 Then
            ReDim Areas(i).NPC(1 To Areas(i).Npcs)
            For e = 1 To Areas(i).Npcs
                Areas(i).NPC(e).NPCIndex = Val(GetVar(archivoC, "Area" & CStr(i), "Npc" & e))
                Areas(i).NPC(e).Cantidad = Val(GetVar(archivoC, "Area" & CStr(i), "Cant" & e))
                If Areas(i).Mapa = UserMap Then
                For H = 1 To Areas(i).NPC(e).Cantidad
                    Call CrearNPC(Areas(i).NPC(e).NPCIndex, i)
                Next H
                End If
            Next e
        End If
    Next i
    End If

frmMain.PintarMapa
End Sub
Sub CrearNPC(NPCIndex As Integer, Area As Integer)
Dim Encuentra As Boolean
Dim X As Integer
Dim Y As Integer
Dim Itera As Integer
Do While Not Encuentra And Itera < 500
    Itera = Itera + 1
    X = RandomNumber(Areas(Area).X1, Areas(Area).X2)
    Y = RandomNumber(Areas(Area).Y1, Areas(Area).Y2)
    Encuentra = LegalPosNPC(X, Y, NpcData(NPCIndex).Agua)
    If Encuentra Then
        MapData(X, Y).NpcArea = NPCIndex
    End If
Loop
End Sub
Function LegalPosNPC(ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte) As Boolean
'***************************************************
'Autor: Unkwnown
'Last Modification: 27/04/2009
'Checks if it's a Legal pos for the npc to move to.
'***************************************************
Dim IsDeadChar As Boolean
Dim UserIndex As Integer
    If (X < 1 Or X > 1100 Or Y < 1 Or Y > 1500) Then
        LegalPosNPC = False
        Exit Function
    End If


    If AguaValida = 0 Then
        LegalPosNPC = (MapData(X, Y).Blocked <> 1) And _
        (MapData(X, Y).NpcArea = 0) And _
        (MapData(X, Y).NPCIndex = 0) And _
        (MapData(X, Y).Trigger <> 3) _
        And Not HayAgua(X, Y)
    Else
        LegalPosNPC = (MapData(X, Y).Blocked <> 1) And _
        (MapData(X, Y).NpcArea = 0) And _
        (MapData(X, Y).NPCIndex = 0) And _
        (MapData(X, Y).Trigger <> 3)
    End If
End Function
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean

If X > 0 And X <= XMaxMapSize And Y > 0 And Y <= YMaxMapSize Then
    If ((MapData(X, Y).Graphic(1).grhindex >= 1505 And MapData(X, Y).Graphic(1).grhindex <= 1520) Or _
    (MapData(X, Y).Graphic(1).grhindex >= 5665 And MapData(X, Y).Graphic(1).grhindex <= 5680) Or _
    (MapData(X, Y).Graphic(1).grhindex >= 13547 And MapData(X, Y).Graphic(1).grhindex <= 13562)) And _
       MapData(X, Y).Graphic(2).grhindex = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function
Public Function GetVar(file As String, Main As String, var As String) As String
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim L As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
szReturn = vbNullString
sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
GetPrivateProfileString Main, var, szReturn, sSpaces, Len(sSpaces), file
GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Public Sub WriteVar(file As String, Main As String, var As String, value As String)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
writeprivateprofilestring Main, var, value, file
End Sub

Public Sub ToggleWalkMode()
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************
On Error GoTo fin:
If WalkMode = False Then
    WalkMode = True
Else
    frmMain.mnuModoCaminata.Checked = False
    WalkMode = False
End If

If WalkMode = False Then
    'Erase character
    Call EraseChar(UserCharIndex)
    MapData(UserPos.X, UserPos.Y).CharIndex = 0
Else
    'MakeCharacter
    If LegalPos(UserPos.X, UserPos.Y) Then
        Call MakeChar(NextOpenChar(), 1, 1, SOUTH, UserPos.X, UserPos.Y)
        UserCharIndex = MapData(UserPos.X, UserPos.Y).CharIndex
        frmMain.mnuModoCaminata.Checked = True
    Else
        MsgBox "ERROR: Ubicacion ilegal."
        WalkMode = False
    End If
End If
fin:
End Sub

Public Sub FixCoasts(ByVal grhindex As Integer, ByVal X As Integer, ByVal Y As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If grhindex = 7284 Or grhindex = 7290 Or grhindex = 7291 Or grhindex = 7297 Or _
   grhindex = 7300 Or grhindex = 7301 Or grhindex = 7302 Or grhindex = 7303 Or _
   grhindex = 7304 Or grhindex = 7306 Or grhindex = 7308 Or grhindex = 7310 Or _
   grhindex = 7311 Or grhindex = 7313 Or grhindex = 7314 Or grhindex = 7315 Or _
   grhindex = 7316 Or grhindex = 7317 Or grhindex = 7319 Or grhindex = 7321 Or _
   grhindex = 7325 Or grhindex = 7326 Or grhindex = 7327 Or grhindex = 7328 Or grhindex = 7332 Or _
   grhindex = 7338 Or grhindex = 7339 Or grhindex = 7345 Or grhindex = 7348 Or _
   grhindex = 7349 Or grhindex = 7350 Or grhindex = 7351 Or grhindex = 7352 Or _
   grhindex = 7349 Or grhindex = 7350 Or grhindex = 7351 Or _
   grhindex = 7354 Or grhindex = 7357 Or grhindex = 7358 Or grhindex = 7360 Or _
   grhindex = 7362 Or grhindex = 7363 Or grhindex = 7365 Or grhindex = 7366 Or _
   grhindex = 7367 Or grhindex = 7368 Or grhindex = 7369 Or grhindex = 7371 Or _
   grhindex = 7373 Or grhindex = 7375 Or grhindex = 7376 Then MapData(X, Y).Graphic(2).grhindex = 0

End Sub

Public Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Randomize Timer
RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
End Function


''
' Actualiza todos los Chars en el mapa
'

Public Sub RefreshAllChars()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
On Error Resume Next
Dim loopc As Integer
frmMain.ApuntadorRadar.Move UserPos.X / 2, UserPos.Y / 2
'frmMain.picRadar.Cls
'For loopc = 1 To LastChar
'    If charlist(loopc).active = 1 Then
'        MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).CharIndex = loopc
'        If charlist(loopc).Heading <> 0 Then
'            frmMain.picRadar.ForeColor = vbGreen
'            frmMain.picRadar.Line (0 + charlist(loopc).Pos.X, 0 + charlist(loopc).Pos.Y)-(2 + charlist(loopc).Pos.X, 0 + charlist(loopc).Pos.Y)
'            frmMain.picRadar.Line (0 + charlist(loopc).Pos.X, 1 + charlist(loopc).Pos.Y)-(2 + charlist(loopc).Pos.X, 1 + charlist(loopc).Pos.Y)
'        End If
'    End If
'Next loopc
bRefreshRadar = False
End Sub


''
' Actualiza el Caption del menu principal
'
' @param Trabajando Indica el path del mapa con el que se esta trabajando
' @param Editado Indica si el mapa esta editado

Public Sub CaptionWorldEditor(ByVal Trabajando As String, ByVal Editado As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If Trabajando = vbNullString Then
    Trabajando = "Nuevo Mapa"
End If
frmMain.Caption = "WorldEditor v" & App.Major & "." & App.Minor & " Build " & App.Revision & " - [" & Trabajando & "]"
If Editado = True Then
    frmMain.Caption = frmMain.Caption & " (modificado)"
End If
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 26/05/2006
'26/05/2005 - GS . DirIndex
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    Open DirIndex & "ao.dat" For Binary Access Read Lock Write As fHandle
        Get fHandle, , ClientSetup
    Close fHandle

End Sub
Public Sub CargarParticulas()
'*************************************
'Coded by OneZero (onezero_ss@hotmail.com)
'Last Modified: 6/4/03
'Loads the Particles.ini file to the ComboBox
'Edited by Juan Martín Sotuyo Dodero to add speed and life
'*************************************
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim myBuffer() As Byte
    Dim StreamFile As String
    Dim Leer As New clsIniReader
    
    StreamFile = DirIndex & "Particles.ini"
    
    Leer.Initialize StreamFile
    
    TotalStreams = Val(Leer.GetValue("INIT", "Total"))
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).name = Leer.GetValue(Val(loopc), "Name")
        frmParticle.lstParticle.AddItem loopc & "-" & StreamData(loopc).name
        StreamData(loopc).NumOfParticles = Leer.GetValue(Val(loopc), "NumOfParticles")
        StreamData(loopc).X1 = Leer.GetValue(Val(loopc), "X1")
        StreamData(loopc).Y1 = Leer.GetValue(Val(loopc), "Y1")
        StreamData(loopc).X2 = Leer.GetValue(Val(loopc), "X2")
        StreamData(loopc).Y2 = Leer.GetValue(Val(loopc), "Y2")
        StreamData(loopc).angle = Leer.GetValue(Val(loopc), "Angle")
        StreamData(loopc).vecx1 = Leer.GetValue(Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = Leer.GetValue(Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = Leer.GetValue(Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = Leer.GetValue(Val(loopc), "VecY2")
        StreamData(loopc).life1 = Leer.GetValue(Val(loopc), "Life1")
        StreamData(loopc).life2 = Leer.GetValue(Val(loopc), "Life2")
        StreamData(loopc).friction = Leer.GetValue(Val(loopc), "Friction")
        StreamData(loopc).spin = Leer.GetValue(Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = Leer.GetValue(Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = Leer.GetValue(Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = Leer.GetValue(Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = Leer.GetValue(Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = Leer.GetValue(Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = Leer.GetValue(Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = Leer.GetValue(Val(loopc), "XMove")
        StreamData(loopc).YMove = Leer.GetValue(Val(loopc), "YMove")
        StreamData(loopc).move_x1 = Leer.GetValue(Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = Leer.GetValue(Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = Leer.GetValue(Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = Leer.GetValue(Val(loopc), "move_y2")
        StreamData(loopc).life_counter = Leer.GetValue(Val(loopc), "life_counter")
        StreamData(loopc).Speed = Val(Leer.GetValue(Val(loopc), "Speed"))
        
        StreamData(loopc).NumGrhs = Leer.GetValue(Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = Leer.GetValue(Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = ReadField(Str(i), GrhListing, Asc(","))
        Next i
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = Leer.GetValue(Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = ReadField(1, TempSet, Asc(","))
            StreamData(loopc).colortint(ColorSet - 1).g = ReadField(2, TempSet, Asc(","))
            StreamData(loopc).colortint(ColorSet - 1).b = ReadField(3, TempSet, Asc(","))
        Next ColorSet
    Next loopc
    
End Sub
 
Public Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim L As Long
    Dim Char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
   
    szReturn = ""
   
    sSpaces = Space$(5000)
   
    GetPrivateProfileString Main, var, szReturn, sSpaces, Len(sSpaces), file
   
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function

Public Function General_Field_Read(ByVal field_pos As Long, ByVal text As String, ByVal delimiter As Byte) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long
   
    LastPos = 0
    FieldNum = 0
    For i = 1 To Len(text)
        If delimiter = CByte(Asc(mid$(text, i, 1))) Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                General_Field_Read = mid$(text, LastPos + 1, (InStr(LastPos + 1, text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = field_pos Then
        General_Field_Read = mid$(text, LastPos + 1)
    End If
End Function
