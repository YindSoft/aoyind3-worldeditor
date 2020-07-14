Attribute VB_Name = "modEdicion"
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
' modEdicion
'
' @remarks Funciones de Edicion
' @author gshaxor@gmail.com
' @version 0.1.38
' @date 20061016

Option Explicit

Public PRect As RECT
Public CopyRect As RECT
Public ZonaR As RECT
Public AreaR As RECT
Public OffX As Integer
Public OffY As Integer
Public mCopyX As Integer
Public mCopyY As Integer
Public CopyState As Byte
Public AgregarZona As Byte
Public AgregarArea As Byte
Public SelArea As Integer
Public AddMY As Integer
Public AddMX As Integer
Public LastX As Integer
Public LastY As Integer
Public MapCopy() As MapBlock
'Public MapCopyD() As MapBlock

Private LastEX As Integer
Private LastEY As Integer
Public Type tRetroceder
    Pos As RECT
    Data() As MapBlock
End Type
Public Const MAX_RETROCEDER As Integer = 100
Public Editado As tRetroceder
Public Retroceder(1 To MAX_RETROCEDER) As tRetroceder
Public NumRetroceder As Integer
Public MapCData(1 To 100, 1 To 100) As MapBlock 'Holds map data for current map

Public MapEData(1 To 100, 1 To 100) As MapBlock
Public MX As Integer
Public MY As Integer
Public PonerLuz As Byte
Public ColorMapa As Long
Public Hora As Byte
''
' Manda una advertencia de Edicion Critica
'
' @return   Nos devuelve si acepta o no el cambio
Public Sub Deshacer()
Dim X As Integer
Dim Y As Integer
If NumRetroceder > 0 Then
With Retroceder(NumRetroceder)

For X = .Pos.Left To .Pos.Right
    For Y = .Pos.Top To .Pos.Bottom
        MapData(X, Y) = .Data(X, Y)
    Next Y
Next X

End With
NumRetroceder = NumRetroceder - 1

End If
End Sub
Public Sub AddDeshacer(ByVal X As Integer, ByVal Y As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Dim i As Integer
Dim XX As Integer, YY As Integer
Dim Distinto As Boolean
If NumRetroceder > 0 Then
With Retroceder(NumRetroceder)
If .Pos.Left <> X Or .Pos.Top <> Y Or .Pos.Right <> X2 Or .Pos.Bottom <> Y2 Then
Distinto = True
Else
    For XX = X To X2
        For YY = Y To Y2
            If MapData(XX, YY).Blocked <> Editado.Data(XX, YY).Blocked Or _
               MapData(XX, YY).TileExit.X <> Editado.Data(XX, YY).TileExit.X Or _
               MapData(XX, YY).CharIndex <> Editado.Data(XX, YY).CharIndex Or _
               MapData(XX, YY).OBJInfo.OBJIndex <> Editado.Data(XX, YY).OBJInfo.OBJIndex Or _
               MapData(XX, YY).Graphic(1).grhindex <> Editado.Data(XX, YY).Graphic(1).grhindex Or _
               MapData(XX, YY).Graphic(2).grhindex <> Editado.Data(XX, YY).Graphic(2).grhindex Or _
               MapData(XX, YY).Graphic(3).grhindex <> Editado.Data(XX, YY).Graphic(3).grhindex Or _
               MapData(XX, YY).Graphic(4).grhindex <> Editado.Data(XX, YY).Graphic(4).grhindex Then
                Distinto = True
            End If
        Next YY
    Next XX
End If
End With
Else
Distinto = True
End If
If Distinto Then
If NumRetroceder = MAX_RETROCEDER Then
    For i = 1 To MAX_RETROCEDER - 1
        Retroceder(i) = Retroceder(i + 1)
    Next i
Else
    NumRetroceder = NumRetroceder + 1
End If
With Retroceder(NumRetroceder)
    .Pos.Left = X
    .Pos.Right = X2
    .Pos.Top = Y
    .Pos.Bottom = Y2
    ReDim .Data(X To X2, Y To Y2)
    For XX = X To X2
        For YY = Y To Y2
            .Data(XX, YY) = MapData(XX, YY)
        Next YY
    Next XX
End With
End If
End Sub
Public Sub AddEditado(ByVal X As Integer, ByVal Y As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Dim XX As Integer, YY As Integer
With Editado
    .Pos.Left = X
    .Pos.Right = X2
    .Pos.Top = Y
    .Pos.Bottom = Y2
    ReDim .Data(X To X2, Y To Y2)
    For XX = X To X2
        For YY = Y To Y2
            .Data(XX, YY) = MapData(XX, YY)
        Next YY
    Next XX
End With
End Sub
Private Function EditWarning() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If MsgBox(MSGDang, vbExclamation + vbYesNo) = vbNo Then
    EditWarning = True
Else
    EditWarning = False
End If
End Function


''
' Bloquea los Bordes del Mapa
'

Public Sub Bloquear_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
            MapData(X, Y).Blocked = 1
        End If
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1
End Sub


''
' Coloca la superficie seleccionada al azar en el mapa
'

Public Sub Superficie_Azar()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error Resume Next
Dim Y As Integer
Dim X As Integer
Dim Cuantos As Integer
Dim k As Integer

If Not MapaCargado Then
    Exit Sub
End If

Cuantos = InputBox("Cuantos Grh se deben poner en este mapa?", "Poner Grh Al Azar", 0)
If Cuantos > 0 Then
    For k = 1 To Cuantos
        X = RandomNumber(10, 90)
        Y = RandomNumber(10, 90)
        If frmConfigSup.MOSAICO.value = vbChecked Then
          Dim aux As Integer
          Dim dy As Integer
          Dim dX As Integer
          If frmConfigSup.DespMosaic.value = vbChecked Then
                        dy = Val(frmConfigSup.DMLargo)
                        dX = Val(frmConfigSup.DMAncho.text)
          Else
                    dy = 0
                    dX = 0
          End If
                
          If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                aux = Val(frmMain.cGrh.text) + _
                (((Y + dy) Mod frmConfigSup.mLargo.text) * frmConfigSup.mAncho.text) + ((X + dX) Mod frmConfigSup.mAncho.text)
                If frmMain.cInsertarBloqueo.value = True Then
                    MapData(X, Y).Blocked = 1
                Else
                    MapData(X, Y).Blocked = 0
                End If
                MapData(X, Y).Graphic(Val(frmMain.cCapas.text)).grhindex = aux
                InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.text)), aux
          Else
                Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
                tXX = X
                tYY = Y
                desptile = 0
                For i = 1 To frmConfigSup.mLargo.text
                    For j = 1 To frmConfigSup.mAncho.text
                        aux = Val(frmMain.cGrh.text) + desptile
                         
                        If frmMain.cInsertarBloqueo.value = True Then
                            MapData(tXX, tYY).Blocked = 1
                        Else
                            MapData(tXX, tYY).Blocked = 0
                        End If

                         MapData(tXX, tYY).Graphic(Val(frmMain.cCapas.text)).grhindex = aux
                         
                         InitGrh MapData(tXX, tYY).Graphic(Val(frmMain.cCapas.text)), aux
                         tXX = tXX + 1
                         desptile = desptile + 1
                    Next
                    tXX = X
                    tYY = tYY + 1
                Next
                tYY = Y
          End If
        End If
    Next
End If

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Coloca la superficie seleccionada en todos los bordes
'

Public Sub Superficie_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If


For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then

          If frmConfigSup.MOSAICO.value = vbChecked Then
            Dim aux As Integer
            aux = Val(frmMain.cGrh.text) + _
            ((Y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (X Mod frmConfigSup.mAncho)
            If frmMain.cInsertarBloqueo.value = True Then
                MapData(X, Y).Blocked = 1
            Else
                MapData(X, Y).Blocked = 0
            End If
            MapData(X, Y).Graphic(Val(frmMain.cCapas.text)).grhindex = aux
            'Setup GRH
            InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.text)), aux
          Else
            'Else Place graphic
            If frmMain.cInsertarBloqueo.value = True Then
                MapData(X, Y).Blocked = 1
            Else
                MapData(X, Y).Blocked = 0
            End If
            
            MapData(X, Y).Graphic(Val(frmMain.cCapas.text)).grhindex = Val(frmMain.cGrh.text)
            
            'Setup GRH
    
            InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.text)), Val(frmMain.cGrh.text)
        End If
             'Erase NPCs
            If MapData(X, Y).NPCIndex > 0 Then
                EraseChar MapData(X, Y).CharIndex
                MapData(X, Y).NPCIndex = 0
            End If

            'Erase Objs
            MapData(X, Y).OBJInfo.OBJIndex = 0
            MapData(X, Y).OBJInfo.Amount = 0
            MapData(X, Y).ObjGrh.grhindex = 0

            'Clear exits
            MapData(X, Y).TileExit.map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0

        End If

    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Coloca la misma superficie seleccionada en todo el mapa
'

Public Sub Superficie_Todo()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If


For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        If frmConfigSup.MOSAICO.value = vbChecked Then
            Dim aux As Integer
            aux = Val(frmMain.cGrh.text) + _
            ((Y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (X Mod frmConfigSup.mAncho)
             MapData(X, Y).Graphic(Val(frmMain.cCapas.text)).grhindex = aux
            'Setup GRH
            InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.text)), aux
        Else
            'Else Place graphic
            MapData(X, Y).Graphic(Val(frmMain.cCapas.text)).grhindex = Val(frmMain.cGrh.text)
            'Setup GRH
            InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.text)), Val(frmMain.cGrh.text)
        End If

    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Modifica los bloqueos de todo mapa
'
' @param Valor Especifica el estado de Bloqueo que se asignara


Public Sub Bloqueo_Todo(ByVal Valor As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub


Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If


For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        MapData(X, Y).Blocked = Valor
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Borra todo el Mapa menos los Triggers
'

Public Sub Borrar_Mapa()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub


Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If


For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        MapData(X, Y).Graphic(1).grhindex = 1
        'Change blockes status
        MapData(X, Y).Blocked = 0

        'Erase layer 2 and 3
        MapData(X, Y).Graphic(2).grhindex = 0
        MapData(X, Y).Graphic(3).grhindex = 0
        MapData(X, Y).Graphic(4).grhindex = 0

        'Erase NPCs
        If MapData(X, Y).NPCIndex > 0 Then
            EraseChar MapData(X, Y).CharIndex
            MapData(X, Y).NPCIndex = 0
        End If

        'Erase Objs
        MapData(X, Y).OBJInfo.OBJIndex = 0
        MapData(X, Y).OBJInfo.Amount = 0
        MapData(X, Y).ObjGrh.grhindex = 0

        'Clear exits
        MapData(X, Y).TileExit.map = 0
        MapData(X, Y).TileExit.X = 0
        MapData(X, Y).TileExit.Y = 0
        
        InitGrh MapData(X, Y).Graphic(1), 1

    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Elimita los NPCs del mapa
'
' @param Hostiles Indica si elimita solo hostiles o solo npcs no hostiles

Public Sub Quitar_NPCs(ByVal Hostiles As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub


Dim Y As Integer
Dim X As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(X, Y).NPCIndex > 0 Then
            If (Hostiles = True And MapData(X, Y).NPCIndex >= 500) Or (Hostiles = False And MapData(X, Y).NPCIndex < 500) Then
                Call EraseChar(MapData(X, Y).CharIndex)
                MapData(X, Y).NPCIndex = 0
            End If
        End If
    Next X
Next Y

bRefreshRadar = True ' Radar

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Elimita todos los Objetos del mapa
'

Public Sub Quitar_Objetos()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub


Dim Y As Integer
Dim X As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(X, Y).OBJInfo.OBJIndex > 0 Then
            If MapData(X, Y).Graphic(3).grhindex = MapData(X, Y).ObjGrh.grhindex Then MapData(X, Y).Graphic(3).grhindex = 0
            MapData(X, Y).OBJInfo.OBJIndex = 0
            MapData(X, Y).OBJInfo.Amount = 0
        End If
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Elimina todos los Triggers del mapa
'

Public Sub Quitar_Triggers()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub


Dim Y As Integer
Dim X As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(X, Y).Trigger > 0 Then
            MapData(X, Y).Trigger = 0
        End If
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Elimita todos los translados del mapa
'

Public Sub Quitar_Translados()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************

If EditWarning Then Exit Sub


Dim Y As Integer
Dim X As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(X, Y).TileExit.map > 0 Then
            MapData(X, Y).TileExit.map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0
        End If
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Elimita todo lo que se encuentre en los bordes del mapa
'

Public Sub Quitar_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

'*****************************************************************
'Clears a border in a room with current GRH
'*****************************************************************

Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If


For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        
            MapData(X, Y).Graphic(1).grhindex = 1
            InitGrh MapData(X, Y).Graphic(1), 1
            MapData(X, Y).Blocked = 0
            
             'Erase NPCs
            If MapData(X, Y).NPCIndex > 0 Then
                EraseChar MapData(X, Y).CharIndex
                MapData(X, Y).NPCIndex = 0
            End If

            'Erase Objs
            MapData(X, Y).OBJInfo.OBJIndex = 0
            MapData(X, Y).OBJInfo.Amount = 0
            MapData(X, Y).ObjGrh.grhindex = 0

            'Clear exits
            MapData(X, Y).TileExit.map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0
            
            ' Triggers
            MapData(X, Y).Trigger = 0

        End If

    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Elimita una capa completa del mapa
'
' @param Capa Especifica la capa


Public Sub Quitar_Capa(ByVal Capa As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

'*****************************************************************
'Clears one layer
'*****************************************************************

Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If Capa = 1 Then
            MapData(X, Y).Graphic(Capa).grhindex = 1
        Else
            MapData(X, Y).Graphic(Capa).grhindex = 0
        End If
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Acciona la operacion al hacer doble click en una posicion del mapa
'
' @param tX Especifica la posicion X en el mapa
' @param tY Espeficica la posicion Y en el mapa

Sub DobleClick(tx As Integer, ty As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
' Selecciones
Seleccionando = False ' GS
SeleccionIX = 0
SeleccionIY = 0
SeleccionFX = 0
SeleccionFY = 0
' Translados
Dim tTrans As WorldPos
tTrans = MapData(tx, ty).TileExit
If tTrans.map > 0 Then
    'If LenB(frmMain.Dialog.FileName) <> 0 Then
      '  If FileExist(PATH_Save & NameMap_Save & tTrans.map & ".map", vbArchive) = True Then
      '      Call modMapIO.NuevoMapa
     '       frmMain.Dialog.FileName = PATH_Save & NameMap_Save & tTrans.map & ".map"
     '       modMapIO.AbrirMapa frmMain.Dialog.FileName
     '       UserPos.X = tTrans.X
     '       UserPos.Y = tTrans.Y
     '       If WalkMode = True Then
    '            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
    '            charlist(UserCharIndex).Heading = SOUTH
   '         End If
            'frmMain.mnuReAbrirMapa.Enabled = True
   '     End If
  '  End If
End If
End Sub

''
' Realiza una operacion de edicion aislada sobre el mapa
'
' @param Button Indica el estado del Click del mouse
' @param tX Especifica la posicion X en el mapa
' @param tY Especifica la posicion Y en el mapa

Sub ClickEdit(Button As Integer, tx As Integer, ty As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    Dim loopc As Integer
    Dim NPCIndex As Integer
    Dim OBJIndex As Integer
    Dim Head As Integer
    Dim Body As Integer
    Dim Heading As Byte
    Dim X As Integer
    Dim Y As Integer
    If ty < 1 Or ty > 1500 Then Exit Sub
    If tx < 1 Or tx > 1100 Then Exit Sub
    
    
    If Button = 0 Then
        ' Pasando sobre :P
        SobreY = ty
        SobreX = tx
        
    End If
If CopyState = 2 Then
            CopyRect.Right = tx
            CopyRect.Bottom = ty
End If
    'Right
    
    If Button = vbRightButton Then
        PRect.Left = 0
        PRect.Top = 0
        PRect.Bottom = 0
        PRect.Right = 0
        LastX = 0
        LastY = 0
        frmMain.Check1.ForeColor = vbWhite
    
        ' Posicion
        frmMain.StatTxt.text = frmMain.StatTxt.text & ENDL & ENDL & "Posición " & tx & "," & ty
        
        ' Bloqueos
        If MapData(tx, ty).Blocked = 1 Then frmMain.StatTxt.text = frmMain.StatTxt.text & " (BLOQ)"
        
        ' Translados
        If MapData(tx, ty).TileExit.map > 0 Then
            If frmMain.mnuAutoCapturarTranslados.Checked = True Then
                frmMain.tTMapa.text = MapData(tx, ty).TileExit.map
                frmMain.tTX.text = MapData(tx, ty).TileExit.X
                frmMain.tTY = MapData(tx, ty).TileExit.Y
            End If
            frmMain.StatTxt.text = frmMain.StatTxt.text & " (Trans.: " & MapData(tx, ty).TileExit.map & "," & MapData(tx, ty).TileExit.X & "," & MapData(tx, ty).TileExit.Y & ")"
        End If
        
        ' NPCs
        If MapData(tx, ty).NPCIndex > 0 Then
            If MapData(tx, ty).NPCIndex > 499 Then
                frmMain.StatTxt.text = frmMain.StatTxt.text & " (NPC-Hostil: " & MapData(tx, ty).NPCIndex & " - " & NpcData(MapData(tx, ty).NPCIndex).name & ")"
            Else
                frmMain.StatTxt.text = frmMain.StatTxt.text & " (NPC: " & MapData(tx, ty).NPCIndex & " - " & NpcData(MapData(tx, ty).NPCIndex).name & ")"
            End If
        End If
        
        ' OBJs
        If MapData(tx, ty).OBJInfo.OBJIndex > 0 Then
            frmMain.StatTxt.text = frmMain.StatTxt.text & " (Obj: " & MapData(tx, ty).OBJInfo.OBJIndex & " - " & ObjData(MapData(tx, ty).OBJInfo.OBJIndex).name & " - Cant.:" & MapData(tx, ty).OBJInfo.Amount & ")"
        End If
        
        ' Capas
        frmMain.StatTxt.text = frmMain.StatTxt.text & ENDL & "Capa1: " & MapData(tx, ty).Graphic(1).grhindex & " - Capa2: " & MapData(tx, ty).Graphic(2).grhindex & " - Capa3: " & MapData(tx, ty).Graphic(3).grhindex & " - Capa4: " & MapData(tx, ty).Graphic(4).grhindex
        If frmMain.mnuAutoCapturarSuperficie.Checked = True And frmMain.cSeleccionarSuperficie.value = False Then
            If MapData(tx, ty).Graphic(4).grhindex <> 0 Then
                frmMain.cCapas.text = 4
                frmMain.cGrh.text = MapData(tx, ty).Graphic(4).grhindex
            ElseIf MapData(tx, ty).Graphic(3).grhindex <> 0 Then
                frmMain.cCapas.text = 3
                frmMain.cGrh.text = MapData(tx, ty).Graphic(3).grhindex
            ElseIf MapData(tx, ty).Graphic(2).grhindex <> 0 Then
                frmMain.cCapas.text = 2
                frmMain.cGrh.text = MapData(tx, ty).Graphic(2).grhindex
            ElseIf MapData(tx, ty).Graphic(1).grhindex <> 0 Then
                frmMain.cCapas.text = 1
                frmMain.cGrh.text = MapData(tx, ty).Graphic(1).grhindex
            End If
        End If
        
        ' Limpieza
        If Len(frmMain.StatTxt.text) > 4000 Then
            frmMain.StatTxt.text = Right(frmMain.StatTxt.text, 3000)
        End If
        frmMain.StatTxt.SelStart = Len(frmMain.StatTxt.text)
        
        Exit Sub
    End If
    
    
    'Left click
    If Button = vbLeftButton Then
            If tx < EditLimit.Left Then EditLimit.Left = tx
            If ty < EditLimit.Top Then EditLimit.Top = ty
            If tx > EditLimit.Right Then EditLimit.Right = tx
            If ty > EditLimit.Bottom Then EditLimit.Bottom = ty
        If AgregarZona = 1 Then
            ZonaR.Left = tx
            ZonaR.Top = ty
            ZonaR.Right = tx
            ZonaR.Bottom = ty
            AgregarZona = 2
            Exit Sub
        ElseIf AgregarZona = 2 Then
            ZonaR.Right = tx
            ZonaR.Bottom = ty
            AgregarZona = 3
        End If
        If AgregarArea = 1 Then
            AreaR.Left = tx
            AreaR.Top = ty
            AreaR.Right = tx
            AreaR.Bottom = ty
            AgregarArea = 2
            Exit Sub
        ElseIf AgregarArea = 2 And AreaR.Left <> AreaR.Right And AreaR.Top <> AreaR.Bottom Then
            AreaR.Right = tx
            AreaR.Bottom = ty
            AgregarArea = 3
            frmMain.GuardaArea
        End If
        If CopyState = 1 Then
            CopyRect.Left = tx
            CopyRect.Top = ty
            CopyRect.Right = tx
            CopyRect.Bottom = ty
            CopyState = 2
            Exit Sub
        ElseIf CopyState = 2 And (CopyRect.Left <> tx Or CopyRect.Top <> ty) Then
            CopyRect.Right = tx
            CopyRect.Bottom = ty
            CopyState = 3
            Exit Sub
        ElseIf CopyState = 4 Then
            If mCopyX <> tx Or mCopyY <> ty Then
            mCopyX = tx
            mCopyY = ty
            Call AddDeshacer(tx - AddMX, ty - AddMY, CopyRect.Right - CopyRect.Left + tx - AddMX, CopyRect.Bottom - CopyRect.Top + ty - AddMY)
            For X = CopyRect.Left To CopyRect.Right
                For Y = CopyRect.Top To CopyRect.Bottom
                    'MapCopyD(X - CopyRect.Left, Y - CopyRect.Top) = MapData(tx + X - CopyRect.Left - AddMX, ty + Y - CopyRect.Top - AddMY)

                    MapData(tx + X - CopyRect.Left - AddMX, ty + Y - CopyRect.Top - AddMY) = MapCopy(X - CopyRect.Left, Y - CopyRect.Top)
                Next Y
            Next X
            Call AddEditado(tx - AddMX, ty - AddMY, CopyRect.Right - CopyRect.Left + tx - AddMX, CopyRect.Bottom - CopyRect.Top + ty - AddMY)
            End If
            Exit Sub
        End If
        If frmMain.Check1.value = vbChecked Then
            If PRect.Left = 0 And (tx <> LastX Or ty <> LastY) Then
                PRect.Left = tx - 1
                PRect.Top = ty - 1
                frmMain.Check1.ForeColor = vbRed
                Exit Sub
            ElseIf (PRect.Left <> tx - 1 Or PRect.Top <> ty - 1) Then
                PRect.Right = tx
                PRect.Bottom = ty
                If PRect.Left + 1 > tx Then
                    PRect.Right = PRect.Left + 1
                    PRect.Left = tx - 1
                End If
                If PRect.Top + 1 > ty Then
                    PRect.Bottom = PRect.Top + 1
                    PRect.Top = ty - 1
                End If
                LastX = tx
                LastY = ty
                If PRect.Left + 1 = PRect.Right And PRect.Top + 1 = PRect.Bottom And frmMain.Check1.value = vbChecked Then
                    PRect.Right = 0
                    PRect.Bottom = 0
                    Exit Sub
                End If
                frmMain.Check1.ForeColor = vbWhite
            Else
                Exit Sub
            End If
        End If
            
            
        If PRect.Left = 0 Then
            PRect.Left = tx - 1
            PRect.Right = tx
            PRect.Top = ty - 1
            PRect.Bottom = ty
        End If
        
        Call AddDeshacer(PRect.Left + 1, PRect.Top + 1, Val(PRect.Right), Val(PRect.Bottom))
        
        
        For X = PRect.Left + 1 To PRect.Right
        For Y = PRect.Top + 1 To PRect.Bottom
            tx = X
            ty = Y
            
            'Erase 2-3
            If frmMain.cQuitarEnTodasLasCapas.value = True Then

                MapInfo.Changed = 1 'Set changed flag
                If MapData(tx, ty).LightIndex > 0 Then
                    Call Light_Destroy(MapData(tx, ty).LightIndex)
                End If
                
                For loopc = 2 To 3
                    MapData(tx, ty).Graphic(loopc).grhindex = 0
                Next loopc
                
                'Exit Sub
            End If
    
            'Borrar "esta" Capa
            If frmMain.cQuitarEnEstaCapa.value = True Then
                If Val(frmMain.cCapas.text) = 1 Then
                    If MapData(tx, ty).Graphic(1).grhindex <> 1 Then
                        MapInfo.Changed = 1 'Set changed flag
                        MapData(tx, ty).Graphic(1).grhindex = 1
                        'Exit Sub
                    End If
                ElseIf MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)).grhindex <> 0 Then
                    MapInfo.Changed = 1 'Set changed flag
                    
                    If Val(frmMain.cCapas.text) = 3 Then
                        If MapData(tx, ty).LightIndex > 0 Then
                            Call Light_Destroy(MapData(tx, ty).LightIndex)
                        End If
                    End If
                    
                    
                    MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)).grhindex = 0
                    'Exit Sub
                End If
            End If
        If PonerLuz > 0 Then
            If MapData(tx, ty).LightIndex > 0 Then
                Call Light_Destroy(MapData(tx, ty).LightIndex)
            End If
            MapData(tx, ty).LightIndex = Light_Create(tx, ty, 255, 255, 255, Val(frmMain.tRango.text), PonerLuz - 1)
            MapData(tx, ty).LuzRango = Val(frmMain.tRango.text)
            MapData(tx, ty).Graphic(3).grhindex = -PonerLuz
        End If
        '************** Place grh
        If frmMain.cSeleccionarSuperficie.value = True Then
            
            If frmConfigSup.MOSAICO.value = vbChecked Then
              Dim aux As Integer
              Dim dy As Integer
              Dim dX As Integer
              If frmConfigSup.DespMosaic.value = vbChecked Then
                            dy = Val(frmConfigSup.DMLargo)
                            dX = Val(frmConfigSup.DMAncho.text)
              Else
                        dy = 0
                        dX = 0
              End If
                    
              If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                    MapInfo.Changed = 1 'Set changed flag
                    aux = Val(frmMain.cGrh.text) + _
                    (((ty + dy) Mod frmConfigSup.mLargo.text) * frmConfigSup.mAncho.text) + ((tx + dX) Mod frmConfigSup.mAncho.text)
                     If MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)).grhindex <> aux Or MapData(tx, ty).Blocked <> frmMain.SelectPanel(2).value Then
                        MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)).grhindex = aux
                        InitGrh MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)), aux
                    End If
              Else
                MapInfo.Changed = 1 'Set changed flag
                Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
                tXX = tx
                tYY = ty
                desptile = 0
                For i = 1 To frmConfigSup.mLargo.text
                    For j = 1 To frmConfigSup.mAncho.text
                        aux = Val(frmMain.cGrh.text) + desptile
                        MapData(tXX, tYY).Graphic(Val(frmMain.cCapas.text)).grhindex = aux
                        InitGrh MapData(tXX, tYY).Graphic(Val(frmMain.cCapas.text)), aux
                        tXX = tXX + 1
                        desptile = desptile + 1
                    Next
                    tXX = tx
                    tYY = tYY + 1
                Next
                tYY = MY
                    
                    
              End If
              
            Else
                'Else Place graphic
                If MapData(tx, ty).Blocked <> frmMain.SelectPanel(2).value Or MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)).grhindex <> Val(frmMain.cGrh.text) Then
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)).grhindex = Val(frmMain.cGrh.text)
                    'Setup GRH
                    InitGrh MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)), Val(frmMain.cGrh.text)
                End If
            End If
            
        End If
        '************** Place blocked tile
        If frmMain.cInsertarBloqueo.value = True Then
            If MapData(tx, ty).Blocked <> 1 Then

                MapInfo.Changed = 1 'Set changed flag
                MapData(tx, ty).Blocked = 1
            End If
        ElseIf frmMain.cQuitarBloqueo.value = True Then
            If MapData(tx, ty).Blocked <> 0 Then
                MapInfo.Changed = 1 'Set changed flag
                MapData(tx, ty).Blocked = 0
            End If
        End If
    
        '************** Place exit
        If frmMain.cInsertarTrans.value = True Then
            If Cfg_TrOBJ > 0 And Cfg_TrOBJ <= NumOBJs And frmMain.cInsertarTransOBJ.value = True Then
                If ObjData(Cfg_TrOBJ).OBJType = 19 Then
                    MapInfo.Changed = 1 'Set changed flag
                    InitGrh MapData(tx, ty).ObjGrh, ObjData(Cfg_TrOBJ).grhindex
                    MapData(tx, ty).OBJInfo.OBJIndex = Cfg_TrOBJ
                    MapData(tx, ty).OBJInfo.Amount = 1
                End If
            End If
            If Val(frmMain.tTMapa.text) < 0 Or Val(frmMain.tTMapa.text) > 9000 Then
                MsgBox "Valor de Mapa invalido", vbCritical + vbOKOnly
                Exit Sub
            ElseIf Val(frmMain.tTX.text) < 0 Or Val(frmMain.tTX.text) > 1100 Then
                MsgBox "Valor de X invalido", vbCritical + vbOKOnly
                Exit Sub
            ElseIf Val(frmMain.tTY.text) < 0 Or Val(frmMain.tTY.text) > 1500 Then
                MsgBox "Valor de Y invalido", vbCritical + vbOKOnly
                Exit Sub
            End If
                If frmMain.cUnionManual.value = True Then
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tx, ty).TileExit.map = Val(frmMain.tTMapa.text)
                    If tx >= 90 Then ' 21 ' derecha
                              MapData(tx, ty).TileExit.X = 12
                              MapData(tx, ty).TileExit.Y = ty
                    ElseIf tx <= 11 Then ' 9 ' izquierda
                        MapData(tx, ty).TileExit.X = 91
                        MapData(tx, ty).TileExit.Y = ty
                    End If
                    If ty >= 91 Then ' 94 '''' hacia abajo
                             MapData(tx, ty).TileExit.Y = 11
                             MapData(tx, ty).TileExit.X = tx
                    ElseIf ty <= 10 Then ''' hacia arriba
                        MapData(tx, ty).TileExit.Y = 90
                        MapData(tx, ty).TileExit.X = tx
                    End If
                Else
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tx, ty).TileExit.map = Val(frmMain.tTMapa.text)
                    MapData(tx, ty).TileExit.X = Val(frmMain.tTX.text)
                    MapData(tx, ty).TileExit.Y = Val(frmMain.tTY.text)
                End If
        ElseIf frmMain.cQuitarTrans.value = True Then
                MapInfo.Changed = 1 'Set changed flag
                MapData(tx, ty).TileExit.map = 0
                MapData(tx, ty).TileExit.X = 0
                MapData(tx, ty).TileExit.Y = 0
        End If
    
        '************** Place NPC
        If frmMain.cInsertarFunc(0).value = True Then
            If frmMain.cNumFunc(0).text > 0 Then
                NPCIndex = frmMain.cNumFunc(0).text
                If NPCIndex <> MapData(tx, ty).NPCIndex Then
                    MapInfo.Changed = 1 'Set changed flag
                    Body = NpcData(NPCIndex).Body
                    Head = NpcData(NPCIndex).Head
                    Heading = NpcData(NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, tx, ty)
                    MapData(tx, ty).NPCIndex = NPCIndex
                End If
            End If
        ElseIf frmMain.cInsertarFunc(1).value = True Then
            If frmMain.cNumFunc(1).text > 0 Then
                NPCIndex = frmMain.cNumFunc(1).text
                If NPCIndex <> (MapData(tx, ty).NPCIndex) Then
                    MapInfo.Changed = 1 'Set changed flag
                    Body = NpcData(NPCIndex).Body
                    Head = NpcData(NPCIndex).Head
                    Heading = NpcData(NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, tx, ty)
                    MapData(tx, ty).NPCIndex = NPCIndex
                End If
            End If
        ElseIf frmMain.cQuitarFunc(0).value = True Or frmMain.cQuitarFunc(1).value = True Then
            If MapData(tx, ty).NPCIndex > 0 Then
                MapInfo.Changed = 1 'Set changed flag
                MapData(tx, ty).NPCIndex = 0
                Call EraseChar(MapData(tx, ty).CharIndex)
            End If
        End If
    
        ' ***************** Control de Funcion de Objetos *****************
        If frmMain.cInsertarFunc(2).value = True Then ' Insertar Objeto
            If frmMain.cNumFunc(2).text > 0 Then
                OBJIndex = frmMain.cNumFunc(2).text
                If MapData(tx, ty).OBJInfo.OBJIndex <> OBJIndex Or MapData(tx, ty).OBJInfo.Amount <> Val(frmMain.cCantFunc(2).text) Then
                    MapInfo.Changed = 1 'Set changed flag
                    InitGrh MapData(tx, ty).ObjGrh, ObjData(OBJIndex).grhindex
                    MapData(tx, ty).OBJInfo.OBJIndex = OBJIndex
                    MapData(tx, ty).OBJInfo.Amount = Val(frmMain.cCantFunc(2).text)
                    Select Case ObjData(OBJIndex).OBJType
                        Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                            MapData(tx, ty).Graphic(3) = MapData(tx, ty).ObjGrh
                    End Select
                End If
            End If
        ElseIf frmMain.cQuitarFunc(2).value = True Then ' Quitar Objeto
            If MapData(tx, ty).OBJInfo.OBJIndex <> 0 Or MapData(tx, ty).OBJInfo.Amount <> 0 Then
                MapInfo.Changed = 1 'Set changed flag
                If MapData(tx, ty).Graphic(3).grhindex = MapData(tx, ty).ObjGrh.grhindex Then MapData(tx, ty).Graphic(3).grhindex = 0
                MapData(tx, ty).ObjGrh.grhindex = 0
                MapData(tx, ty).OBJInfo.OBJIndex = 0
                MapData(tx, ty).OBJInfo.Amount = 0
            End If
        End If
        
        ' ***************** Control de Funcion de Triggers *****************
        If frmMain.cInsertarTrigger.value = True Then ' Insertar Trigger
            If MapData(tx, ty).Trigger <> frmMain.lListado(4).ListIndex Then
                MapInfo.Changed = 1 'Set changed flag
                MapData(tx, ty).Trigger = frmMain.lListado(4).ListIndex
            End If
        ElseIf frmMain.cQuitarTrigger.value = True Then ' Quitar Trigger
            If MapData(tx, ty).Trigger <> 0 Then
                MapInfo.Changed = 1 'Set changed flag
                MapData(tx, ty).Trigger = 0
            End If
        End If
        
        ' ***************** Control de Funcion de Particles! *****************
        If frmParticle.cmdAdd.value = True Then ' Insertar Particle
            MapInfo.Changed = 1 'Set changed flag
            General_Particle_Create frmParticle.lstParticle.ListIndex + 1, tx, ty, frmParticle.Life.text
        ElseIf frmParticle.cmdDel.value = True Then ' Quitar Particle
            If MapData(tx, ty).particle_group_index <> 0 Then
                MapInfo.Changed = 1 'Set changed flag
                engine.Particle_Group_Remove MapData(tx, ty).particle_group_index
                MapData(tx, ty).particle_group_index = 0
            End If
        End If
        Next Y
        Next X
        Call AddEditado(PRect.Left + 1, PRect.Top + 1, Val(PRect.Right), Val(PRect.Bottom))
        'If frmMain.Check2.value = vbChecked Then
            PRect.Left = 0
            PRect.Right = 0
            PRect.Bottom = 0
            PRect.Top = 0
        'End If
        
        
    End If

End Sub
