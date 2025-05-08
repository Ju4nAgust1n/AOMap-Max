Attribute VB_Name = "modMapIO"
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
' modMapIO
'
' @remarks Funciones Especificas al trabajo con Archivos de Mapas
' @author gshaxor@gmail.com
' @version 0.1.15
' @date 20060602

Option Explicit

Private MapTitulo As String     ' GS > Almacena el titulo del mapa para el .dat

''
' Obtener el tama�o de un archivo
'
' @param FileName Especifica el path del archivo
' @return   Nos devuelve el tama�o

Public Function FileSize(ByRef FileName As String) As Long
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

    On Error GoTo FalloFile
    Dim nFileNum As Integer
    Dim lFileSize As Long
    
    nFileNum = FreeFile
    Open FileName For Input As nFileNum
    lFileSize = LOF(nFileNum)
    Close nFileNum
    FileSize = lFileSize
    
    Exit Function
FalloFile:
    FileSize = -1
End Function

''
' Nos dice si existe el archivo/directorio
'
' @param file Especifica el path
' @param FileType Especifica el tipo de archivo/directorio
' @return   Nos devuelve verdadero o falso

Public Function FileExist(ByRef File As String, ByVal FileType As VbFileAttribute) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06
'*************************************************

FileExist = (LenB(Dir$(File, FileType)) > 0)
End Function

''
' Abre un Mapa
'
' @param Path Especifica el path del mapa

Public Sub AbrirMapa(ByRef Path As String, ByRef buffer() As MapBlock, Optional ByVal SoloMap As Boolean = False, Optional m As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If FileSize(Path) = 130273 Then
    Call MapaV1_Cargar(Path, buffer, SoloMap)
    frmMain.mnuUtirialNuevoFormato.Checked = False
Else
    Call MapaV2_Cargar(Path, buffer, SoloMap, m)
    frmMain.mnuUtirialNuevoFormato.Checked = True
End If
End Sub

''
' Guarda el Mapa
'
' @param Path Especifica el path del mapa

Public Sub GuardarMapa(Optional ByRef Path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************

frmMain.Dialog.CancelError = True
On Error GoTo ErrHandler

If LenB(Path) = 0 Then
    frmMain.ObtenerNombreArchivo True
    Path = frmMain.Dialog.FileName
    If LenB(Path) = 0 Then Exit Sub
End If

If frmMain.mnuUtirialNuevoFormato.Checked = True Then
    'Call MapaV2_Guardar(Path)
    Call GrabarMapa(Path)
Else
    Call MapaV1_Guardar(Path)
End If

ErrHandler:
End Sub

''
' Nos pregunta donde guardar el mapa en caso de modificarlo
'
' @param Path Especifica si existiera un path donde guardar el mapa

Public Sub DeseaGuardarMapa(Optional ByRef Path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If MapInfo.Changed = 1 Then
    If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
        GuardarMapa Path
    End If
End If
End Sub

''
' Limpia todo el mapa a uno nuevo
'

Public Sub NuevoMapa()
'*************************************************
'Author: ^[GS]^
'Last modified: 21/05/06
'*************************************************

On Error Resume Next

Dim LoopC As Integer
Dim Y As Integer
Dim X As Integer

bAutoGuardarMapaCount = 0

frmMain.mnuUtirialNuevoFormato.Checked = True
frmMain.mnuReAbrirMapa.Enabled = False
frmMain.TimAutoGuardarMapa.Enabled = False
frmMain.lblMapVersion.Caption = 0

MapaCargado = False

For LoopC = 0 To frmMain.MapPest.Count - 1
    frmMain.MapPest(LoopC).Enabled = False
Next

frmMain.MousePointer = 11

ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

For LoopC = 1 To LastChar
    If CharList(LoopC).Active = 1 Then Call EraseChar(LoopC)
Next LoopC

MapInfo.MapVersion = 0
MapInfo.Name = "Nuevo Mapa"
MapInfo.Music = 0
MapInfo.PK = True
MapInfo.MagiaSinEfecto = 0
MapInfo.Terreno = "BOSQUE"
MapInfo.Zona = "CAMPO"
MapInfo.Restringir = "No"
MapInfo.NoEncriptarMP = 0

Call MapInfo_Actualizar

bRefreshRadar = True ' Radar

'Set changed flag
MapInfo.Changed = 0
frmMain.MousePointer = 0

' Vacio deshacer
modEdicion.Deshacer_Clear

MapaCargado = True

frmMain.SetFocus

End Sub

''
' Guardar Mapa con el formato V2
'
' @param SaveAs Especifica donde guardar el mapa

Public Sub MapaV2_Guardar(ByVal saveas As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo ErrorSave
Dim FreeFileMap As Long
Dim FreeFileInf As Long
Dim LoopC As Long
Dim TempInt As Integer
Dim Y As Long
Dim X As Long
Dim ByFlags As Double

If FileExist(saveas, vbNormal) = True Then
    If MsgBox("�Desea sobrescribir " & saveas & "?", vbCritical + vbYesNo) = vbNo Then
        Exit Sub
    Else
        Kill saveas
    End If
End If

frmMain.MousePointer = 11

' y borramos el .inf tambien
If FileExist(Left$(saveas, Len(saveas) - 4) & ".inf", vbNormal) = True Then
    Kill Left$(saveas, Len(saveas) - 4) & ".inf"
End If

'Open .map file
FreeFileMap = FreeFile
Open saveas For Binary As FreeFileMap
Seek FreeFileMap, 1

saveas = Left$(saveas, Len(saveas) - 4)
saveas = saveas & ".inf"

'Open .inf file
FreeFileInf = FreeFile
Open saveas For Binary As FreeFileInf
Seek FreeFileInf, 1

    'map Header
    
    ' Version del Mapa
    If frmMain.lblMapVersion.Caption < 32767 Then
        frmMain.lblMapVersion.Caption = frmMain.lblMapVersion + 1
        frmMapInfo.txtMapVersion = frmMain.lblMapVersion.Caption
    End If
    Put FreeFileMap, , CInt(frmMain.lblMapVersion.Caption)
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
                ByFlags = 0
                
                If MapData(X, Y).Blocked = 1 Then ByFlags = ByFlags Or 1
                If MapData(X, Y).Graphic(2).GrhIndex Then ByFlags = ByFlags Or 2
                If MapData(X, Y).Graphic(3).GrhIndex Then ByFlags = ByFlags Or 4
                If MapData(X, Y).Graphic(4).GrhIndex Then ByFlags = ByFlags Or 8
                If MapData(X, Y).Trigger Then ByFlags = ByFlags Or 16
                If MapData(X, Y).TypeZona Then ByFlags = ByFlags Or 32
                
                Put FreeFileMap, , ByFlags

                Put FreeFileMap, , MapData(X, Y).Graphic(1).GrhIndex
                
                For LoopC = 2 To 4
                    If MapData(X, Y).Graphic(LoopC).GrhIndex Then _
                        Put FreeFileMap, , MapData(X, Y).Graphic(LoopC).GrhIndex
                Next LoopC
                
                If MapData(X, Y).Trigger Then _
                    Put FreeFileMap, , MapData(X, Y).Trigger
                    
                If MapData(X, Y).TypeZona Then _
                    Put FreeFileMap, , MapData(X, Y).TypeZona
                
                '.inf file
                
                ByFlags = 0
                
                If MapData(X, Y).TileExit.Map Then ByFlags = ByFlags Or 1
                If MapData(X, Y).NPCIndex Then ByFlags = ByFlags Or 2
                If MapData(X, Y).OBJInfo.objindex Then ByFlags = ByFlags Or 4
                
                Put FreeFileInf, , ByFlags
                
                If MapData(X, Y).TileExit.Map Then
                    Put FreeFileInf, , MapData(X, Y).TileExit.Map
                    Put FreeFileInf, , MapData(X, Y).TileExit.X
                    Put FreeFileInf, , MapData(X, Y).TileExit.Y
                End If
                
                If MapData(X, Y).NPCIndex Then
                    Put FreeFileInf, , CInt(MapData(X, Y).NPCIndex)
                End If
                
                If MapData(X, Y).OBJInfo.objindex Then
                    Put FreeFileInf, , MapData(X, Y).OBJInfo.objindex
                    Put FreeFileInf, , MapData(X, Y).OBJInfo.Amount
                End If
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    
    'Close .inf file
    Close FreeFileInf


Call Pesta�as(saveas)

'write .dat file
saveas = Left$(saveas, Len(saveas) - 4) & ".dat"
MapInfo_Guardar saveas

'Change mouse icon
frmMain.MousePointer = 0
MapInfo.Changed = 0

Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & err.Number & " - " & err.Description
End Sub

Public Sub MapaV2Simple_Guardar(ByVal saveas As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo ErrorSave
Dim FreeFileMap As Long
Dim FreeFileInf As Long
Dim LoopC As Long
Dim TempInt As Integer
Dim Y As Long
Dim X As Long
Dim ByFlags As Byte

If FileExist(saveas, vbNormal) = True Then
    If MsgBox("�Desea sobrescribir " & saveas & "?. Debe guardar el mapa con un nombre distinto al del server", vbCritical + vbYesNo) = vbNo Then
        Exit Sub
    Else
        Kill saveas
    End If
End If

frmMain.MousePointer = 11

'Open .map file
FreeFileMap = FreeFile
Open saveas For Binary As FreeFileMap
Seek FreeFileMap, 1

    'map Header
    
    ' Version del Mapa
    If frmMain.lblMapVersion.Caption < 32767 Then
        frmMain.lblMapVersion.Caption = frmMain.lblMapVersion + 1
        frmMapInfo.txtMapVersion = frmMain.lblMapVersion.Caption
    End If
    Put FreeFileMap, , CInt(frmMain.lblMapVersion.Caption)
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
                ByFlags = 0
                
                If MapData(X, Y).Blocked = 1 Then ByFlags = ByFlags Or 1
                If MapData(X, Y).Graphic(2).GrhIndex Then ByFlags = ByFlags Or 2
                If MapData(X, Y).Graphic(3).GrhIndex Then ByFlags = ByFlags Or 4
                If MapData(X, Y).Graphic(4).GrhIndex Then ByFlags = ByFlags Or 8
                If MapData(X, Y).Trigger Then ByFlags = ByFlags Or 16
                
                Put FreeFileMap, , ByFlags

                Put FreeFileMap, , MapData(X, Y).Graphic(1).GrhIndex
                
                For LoopC = 2 To 4
                    If MapData(X, Y).Graphic(LoopC).GrhIndex Then _
                        Put FreeFileMap, , MapData(X, Y).Graphic(LoopC).GrhIndex
                Next LoopC
                
                If MapData(X, Y).Trigger Then _
                    Put FreeFileMap, , MapData(X, Y).Trigger
 
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap


Call Pesta�as(saveas)

'Change mouse icon
frmMain.MousePointer = 0
MapInfo.Changed = 0

Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & err.Number & " - " & err.Description
End Sub


''
' Guardar Mapa con el formato V1
'
' @param SaveAs Especifica donde guardar el mapa

Public Sub MapaV1_Guardar(saveas As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo ErrorSave
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim LoopC As Long
    Dim TempInt As Integer
    Dim Y As Long
    Dim X As Long
    
    If FileExist(saveas, vbNormal) = True Then
        If MsgBox("�Desea sobrescribir " & saveas & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill saveas
        End If
    End If
    
    'Change mouse icon
    frmMain.MousePointer = 11
    
    If FileExist(Left$(saveas, Len(saveas) - 4) & ".inf", vbNormal) = True Then
        Kill Left$(saveas, Len(saveas) - 4) & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open saveas For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    
    saveas = Left$(saveas, Len(saveas) - 4)
    saveas = saveas & ".inf"
    'Open .inf file
    FreeFileInf = FreeFile
    Open saveas For Binary As FreeFileInf
    Seek FreeFileInf, 1
    'map Header
    If frmMain.lblMapVersion.Caption < 32767 Then
        frmMain.lblMapVersion.Caption = frmMain.lblMapVersion + 1
        frmMapInfo.txtMapVersion = frmMain.lblMapVersion.Caption
    End If
    Put FreeFileMap, , CInt(frmMain.lblMapVersion.Caption)
    Put FreeFileMap, , MiCabecera
    
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            '.map file
            
            ' Bloqueos
            Put FreeFileMap, , MapData(X, Y).Blocked
            
            ' Capas
            For LoopC = 1 To 4
                If LoopC = 2 Then Call FixCoasts(MapData(X, Y).Graphic(LoopC).GrhIndex, X, Y)
                Put FreeFileMap, , MapData(X, Y).Graphic(LoopC).GrhIndex
            Next LoopC
            
            ' Triggers
            Put FreeFileMap, , MapData(X, Y).Trigger
            Put FreeFileMap, , TempInt
            
            '.inf file
            'Tile exit
            Put FreeFileInf, , MapData(X, Y).TileExit.Map
            Put FreeFileInf, , MapData(X, Y).TileExit.X
            Put FreeFileInf, , MapData(X, Y).TileExit.Y
            
            'NPC
            Put FreeFileInf, , MapData(X, Y).NPCIndex
            
            'Object
            Put FreeFileInf, , MapData(X, Y).OBJInfo.objindex
            Put FreeFileInf, , MapData(X, Y).OBJInfo.Amount
            
            'Empty place holders for future expansion
            Put FreeFileInf, , TempInt
            Put FreeFileInf, , TempInt
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    'Close .inf file
    Close FreeFileInf
    
    Call Pesta�as(saveas)
    
    'write .dat file
    saveas = Left$(saveas, Len(saveas) - 4) & ".dat"
    MapInfo_Guardar saveas
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
Exit Sub
ErrorSave:
    MsgBox "Error " & err.Number & " - " & err.Description
End Sub

''
' Abrir Mapa con el formato V2
'
' @param Map Especifica el Path del mapa

Public Sub MapaV2_Cargar(ByVal Map As String, ByRef buffer() As MapBlock, ByVal SoloMap As Boolean, Optional MapaViejo As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    Dim LoopC As Integer
    Dim TempInt As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As Byte
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    
    'Change mouse icon
    frmMain.MousePointer = 11
       
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    If Not SoloMap Then
        Map = Left$(Map, Len(Map) - 4)
        Map = Map & ".inf"
        
        FreeFileInf = FreeFile
        Open Map For Binary As FreeFileInf
        Seek FreeFileInf, 1
    End If
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
   ' Get FreeFileMap, , TempInt
    
    If Not SoloMap Then
        'Cabecera inf
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
    End If
    
Dim valorMin As Integer
Dim valorMax As Integer

If MapaViejo > 0 Then
      valorMin = 1
      valorMax = 100
Else
     valorMin = 1
     valorMax = 300
End If

    'Load arrays
    For Y = valorMin To valorMax
        For X = valorMin To valorMax
    
            Get FreeFileMap, , ByFlags
            
            buffer(X, Y).Blocked = (ByFlags And 1)
            
            Get FreeFileMap, , buffer(X, Y).Graphic(1).GrhIndex
            InitGrh buffer(X, Y).Graphic(1), buffer(X, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get FreeFileMap, , buffer(X, Y).Graphic(2).GrhIndex
                InitGrh buffer(X, Y).Graphic(2), buffer(X, Y).Graphic(2).GrhIndex
            Else
                buffer(X, Y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get FreeFileMap, , buffer(X, Y).Graphic(3).GrhIndex
                InitGrh buffer(X, Y).Graphic(3), buffer(X, Y).Graphic(3).GrhIndex
            Else
                buffer(X, Y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get FreeFileMap, , buffer(X, Y).Graphic(4).GrhIndex
                InitGrh buffer(X, Y).Graphic(4), buffer(X, Y).Graphic(4).GrhIndex
            Else
                buffer(X, Y).Graphic(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get FreeFileMap, , buffer(X, Y).TypeZona
            Else
                buffer(X, Y).TypeZona = 0
            End If
            
           'If ByFlags And 32 Then
             '   Get FreeFileMap, , buffer(X, Y).TypeZona
           ' Else
                'buffer(X, Y).TypeZona = 0
            'End If
            
            If Not SoloMap Then
                '.inf file
                Get FreeFileInf, , ByFlags
                
                If ByFlags And 1 Then
                    Get FreeFileInf, , buffer(X, Y).TileExit.Map
                    Get FreeFileInf, , buffer(X, Y).TileExit.X
                    Get FreeFileInf, , buffer(X, Y).TileExit.Y
                End If
        
                If ByFlags And 2 Then
                    'Get and make NPC
                    Get FreeFileInf, , buffer(X, Y).NPCIndex
        
                    If buffer(X, Y).NPCIndex < 0 Then
                        buffer(X, Y).NPCIndex = 0
                    Else
                        Body = NpcData(buffer(X, Y).NPCIndex).Body
                        Head = NpcData(buffer(X, Y).NPCIndex).Head
                        Heading = NpcData(buffer(X, Y).NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)
                    End If
                End If
        
                If ByFlags And 4 Then
                    'Get and make Object
                    Get FreeFileInf, , buffer(X, Y).OBJInfo.objindex
                    Get FreeFileInf, , buffer(X, Y).OBJInfo.Amount
                    If buffer(X, Y).OBJInfo.objindex > 0 Then
                        InitGrh buffer(X, Y).ObjGrh, ObjData(buffer(X, Y).OBJInfo.objindex).GrhIndex
                    End If
                End If
                
            End If
            
        Next X
    Next Y
    
    'Close files
    Close FreeFileMap
    
    If Not SoloMap Then
        Close FreeFileInf
        
        Call Pesta�as(Map)
        
        bRefreshRadar = True ' Radar
        
        Map = Left$(Map, Len(Map) - 4) & ".dat"
        
        MapInfo_Cargar Map
        frmMain.lblMapVersion.Caption = MapInfo.MapVersion
        
        'Set changed flag
        MapInfo.Changed = 0
        
        ' Vacia el Deshacer
        modEdicion.Deshacer_Clear
    End If
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True
End Sub

''
' Abrir Mapa con el formato V1
'
' @param Map Especifica el Path del mapa

Public Sub MapaV1_Cargar(ByVal Map As String, ByRef buffer() As MapBlock, ByVal SoloMap As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    On Error Resume Next
    Dim TBlock As Byte
    Dim LoopC As Integer
    Dim TempInt As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As Byte
    Dim Y As Integer
    Dim X As Integer
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    
    
    MsgBox "A"
    End
    'Change mouse icon
    frmMain.MousePointer = 11
    
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    
    
    Seek FreeFileMap, 1
    
    If Not SoloMap Then
        Map = Left$(Map, Len(Map) - 4)
        Map = Map & ".inf"
        FreeFileInf = FreeFile
        Open Map For Binary As #2
        Seek FreeFileInf, 1
    End If
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    
    If Not SoloMap Then
        'Cabecera inf
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
    End If
    
    'Load arrays
    For Y = 1 To 100
        For X = 1 To 100
    
            '.map file
            Get FreeFileMap, , buffer(X, Y).Blocked
            
            For LoopC = 1 To 4
                Get FreeFileMap, , buffer(X, Y).Graphic(LoopC).GrhIndex
                'Set up GRH
                If buffer(X, Y).Graphic(LoopC).GrhIndex > 0 Then
                    InitGrh buffer(X, Y).Graphic(LoopC), buffer(X, Y).Graphic(LoopC).GrhIndex
                End If
            Next LoopC
            'Trigger
            Get FreeFileMap, , buffer(X, Y).Trigger
            
            Get FreeFileMap, , TempInt
            
            If Not SoloMap Then
                '.inf file
                
                'Tile exit
                Get FreeFileInf, , buffer(X, Y).TileExit.Map
                Get FreeFileInf, , buffer(X, Y).TileExit.X
                Get FreeFileInf, , buffer(X, Y).TileExit.Y
                              
                'make NPC
                Get FreeFileInf, , buffer(X, Y).NPCIndex
                If buffer(X, Y).NPCIndex > 0 Then
                    Body = NpcData(buffer(X, Y).NPCIndex).Body
                    Head = NpcData(buffer(X, Y).NPCIndex).Head
                    Heading = NpcData(buffer(X, Y).NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)
                End If
                
                'Make obj
                Get FreeFileInf, , buffer(X, Y).OBJInfo.objindex
                Get FreeFileInf, , buffer(X, Y).OBJInfo.Amount
                If buffer(X, Y).OBJInfo.objindex > 0 Then
                    InitGrh buffer(X, Y).ObjGrh, ObjData(buffer(X, Y).OBJInfo.objindex).GrhIndex
                End If
                
                'Empty place holders for future expansion
                Get FreeFileInf, , TempInt
                Get FreeFileInf, , TempInt
            End If
        Next X
    Next Y
    
    'Close files
    Close FreeFileMap
    
    If Not SoloMap Then
        Close FreeFileInf
        
        Call Pesta�as(Map)
        
        bRefreshRadar = True ' Radar
        
        Map = Left$(Map, Len(Map) - 4) & ".dat"
            
        MapInfo_Cargar Map
        frmMain.lblMapVersion.Caption = MapInfo.MapVersion
        
        'Set changed flag
        MapInfo.Changed = 0
        
        ' Vacia el Deshacer
        modEdicion.Deshacer_Clear
    End If
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True
End Sub



' *****************************************************************************
' MAPINFO *********************************************************************
' *****************************************************************************

''
' Guardar Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Guardar(ByVal Archivo As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

    If LenB(MapTitulo) = 0 Then
        MapTitulo = NameMap_Save
    End If

    Call WriteVar(Archivo, MapTitulo, "Name", MapInfo.Name)
    Call WriteVar(Archivo, MapTitulo, "MusicNum", MapInfo.Music)
    Call WriteVar(Archivo, MapTitulo, "MagiaSinefecto", Val(MapInfo.MagiaSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "NoEncriptarMP", Val(MapInfo.NoEncriptarMP))

    Call WriteVar(Archivo, MapTitulo, "Terreno", MapInfo.Terreno)
    Call WriteVar(Archivo, MapTitulo, "Zona", MapInfo.Zona)
    Call WriteVar(Archivo, MapTitulo, "Restringir", MapInfo.Restringir)
    Call WriteVar(Archivo, MapTitulo, "BackUp", str(MapInfo.BackUp))

    If MapInfo.PK Then
        Call WriteVar(Archivo, MapTitulo, "Pk", "0")
    Else
        Call WriteVar(Archivo, MapTitulo, "Pk", "1")
    End If
    
End Sub

''
' Abrir Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Cargar(ByVal Archivo As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 02/06/06
'*************************************************

On Error Resume Next
    Dim Leer As New clsIniReader
    Dim LoopC As Integer
    Dim Path As String
    
    MapTitulo = Empty
    Leer.Initialize Archivo

    For LoopC = Len(Archivo) To 1 Step -1
        If mid$(Archivo, LoopC, 1) = "\" Then
            Path = Left$(Archivo, LoopC)
            Exit For
        End If
    Next LoopC
    
    Archivo = Right$(Archivo, Len(Archivo) - (Len(Path)))
    MapTitulo = UCase$(Left$(Archivo, Len(Archivo) - 4))

    MapInfo.Name = Leer.GetValue(MapTitulo, "Name")
    MapInfo.Music = Leer.GetValue(MapTitulo, "MusicNum")
    MapInfo.MagiaSinEfecto = Val(Leer.GetValue(MapTitulo, "MagiaSinEfecto"))
    MapInfo.NoEncriptarMP = Val(Leer.GetValue(MapTitulo, "NoEncriptarMP"))
    
    If Val(Leer.GetValue(MapTitulo, "Pk")) = 0 Then
        MapInfo.PK = True
    Else
        MapInfo.PK = False
    End If
    
    MapInfo.Terreno = Leer.GetValue(MapTitulo, "Terreno")
    MapInfo.Zona = Leer.GetValue(MapTitulo, "Zona")
    MapInfo.Restringir = Leer.GetValue(MapTitulo, "Restringir")
    MapInfo.BackUp = Val(Leer.GetValue(MapTitulo, "BACKUP"))
    
    Call MapInfo_Actualizar
    
End Sub

''
' Actualiza el formulario de MapInfo
'

Public Sub MapInfo_Actualizar()
'*************************************************
'Author: ^[GS]^
'Last modified: 02/06/06
'*************************************************

On Error Resume Next
    ' Mostrar en Formularios
    frmMapInfo.txtMapNombre.Text = MapInfo.Name
    frmMapInfo.txtMapMusica.Text = MapInfo.Music
    frmMapInfo.txtMapTerreno.Text = MapInfo.Terreno
    frmMapInfo.txtMapZona.Text = MapInfo.Zona
    frmMapInfo.txtMapRestringir.Text = MapInfo.Restringir
    frmMapInfo.chkMapBackup.value = MapInfo.BackUp
    frmMapInfo.chkMapMagiaSinEfecto.value = MapInfo.MagiaSinEfecto
    frmMapInfo.chkMapNoEncriptarMP.value = MapInfo.NoEncriptarMP
    frmMapInfo.chkMapPK.value = IIf(MapInfo.PK = True, 1, 0)
    frmMapInfo.txtMapVersion = MapInfo.MapVersion
    frmMain.lblMapNombre = MapInfo.Name
    frmMain.lblMapMusica = MapInfo.Music

End Sub

''
' Calcula la orden de Pesta�as
'
' @param Map Especifica path del mapa

Public Sub Pesta�as(ByVal Map As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
On Error Resume Next
Dim LoopC As Integer

PATH_Save = Left$(Map, InStrRev(Map, "\"))
Map = Right$(Map, Len(Map) - Len(PATH_Save))
Map = Left$(Map, Len(Map) - 4) 'Sacamos la extension

For LoopC = Len(Map) To 1 Step -1
    If Not IsNumeric(mid$(Map, LoopC)) Then
        NumMap_Save = Val(mid$(Map, LoopC + 1))
        NameMap_Save = Left$(Map, LoopC)
        Exit For
    End If
Next LoopC

For LoopC = (NumMap_Save - 4) To (NumMap_Save + 8)
    If FileExist(PATH_Save & NameMap_Save & LoopC & ".map", vbArchive) Then
        frmMain.MapPest(LoopC - NumMap_Save + 4).Visible = True
        frmMain.MapPest(LoopC - NumMap_Save + 4).Enabled = True
        frmMain.MapPest(LoopC - NumMap_Save + 4).Caption = NameMap_Save & LoopC
    Else
        frmMain.MapPest(LoopC - NumMap_Save + 4).Visible = False
    End If
Next LoopC
End Sub
