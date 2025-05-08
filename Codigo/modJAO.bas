Attribute VB_Name = "modJAO"

'BORRE EL CONDESHACER

Public Sub GrabarMapa(ByRef MAPFILE As String)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2011
'10/08/2010 - Pato: Implemento el clsByteBuffer para el grabado de mapas
'28/10/2010:ZaMa - Ahora no se hace backup de los pretorianos.
'12/01/2011 - Amraphen: Ahora no se hace backup de NPCs prohibidos (Pretorianos, Mascotas, Invocados y Centinela)
'***************************************************

On Error Resume Next

MAPFILE = Left$(MAPFILE, Len(MAPFILE) - 4)

    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim loopc As Long
    Dim MapWriter As clsByteBuffer
    Dim InfWriter As clsByteBuffer
    Dim IniManager As clsIniManager
    
    Set MapWriter = New clsByteBuffer
    Set InfWriter = New clsByteBuffer
    Set IniManager = New clsIniManager
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Kill MAPFILE & ".map"
    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Kill MAPFILE & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    
    Call MapWriter.initializeWriter(FreeFileMap)
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    
    Call InfWriter.initializeWriter(FreeFileInf)
    
    'map Header
    Call MapWriter.putInteger(frmMapInfo.txtMapVersion = frmMain.lblMapVersion.Caption)
        
    Call MapWriter.putString(MiCabecera.Desc, False)
    Call MapWriter.putLong(MiCabecera.CRC)
    Call MapWriter.putLong(MiCabecera.MagicWord)
    
    Call MapWriter.putDouble(0)
    
    'inf Header
    Call InfWriter.putDouble(0)
    Call InfWriter.putInteger(0)
    
    'Write .map file
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(X, y)
                ByFlags = 0
                
                If .Blocked Then ByFlags = ByFlags Or 1
                If .Graphic(2).GrhIndex > 0 Then ByFlags = ByFlags Or 2
                If .Graphic(3).GrhIndex > 0 Then ByFlags = ByFlags Or 4
                If .Graphic(4).GrhIndex > 0 Then ByFlags = ByFlags Or 8
                If .TypeZona Then ByFlags = ByFlags Or 16
                
                Call MapWriter.putByte(ByFlags)
                
                Call MapWriter.putInteger(.Graphic(1).GrhIndex)
                
                For loopc = 2 To 4
                    If .Graphic(loopc).GrhIndex > 0 Then _
                        Call MapWriter.putInteger(.Graphic(loopc).GrhIndex)
                Next loopc
                
                If .TypeZona Then _
                    Call MapWriter.putInteger(CInt(.TypeZona))
                
                '.inf file
                ByFlags = 0
                
             '   If .OBJInfo.objindex > 0 Then
                  ' If ObjData(.OBJInfo.objindex).ObjType = eOBJType.otFogata Then
                       ' .OBJInfo.objindex = 0
                      '  .OBJInfo.Amount = 0
                    'End If
                'End If
    
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                ' No hacer backup de los NPCs inválidos (Pretorianos, Mascotas, Invocados y Centinela)
                'If .NPCIndex Then
                 '   NpcInvalido = (Npclist(.NPCIndex).NPCtype = eNPCType.Pretoriano) Or (Npclist(.NPCIndex).MaestroUser > 0) Or EsCentinela(.NPCIndex)
                    
                   ' If Not NpcInvalido Then ByFlags = ByFlags Or 2
              '  End If
              
                If .NPCIndex Then ByFlags = ByFlags Or 2
                
                If .OBJInfo.objindex Then ByFlags = ByFlags Or 4
                
                Call InfWriter.putByte(ByFlags)
                
                If .TileExit.Map Then
                    Call InfWriter.putInteger(.TileExit.Map)
                    Call InfWriter.putInteger(.TileExit.X)
                    Call InfWriter.putInteger(.TileExit.y)
                End If
                
                If .NPCIndex Then _
                    Call InfWriter.putInteger(.NPCIndex)
                
                If .OBJInfo.objindex Then
                    Call InfWriter.putInteger(.OBJInfo.objindex)
                    Call InfWriter.putInteger(.OBJInfo.Amount)
                End If
                
            End With
        Next X
    Next y
    
    Call MapWriter.saveBuffer
    Call InfWriter.saveBuffer
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf
    
    Set MapWriter = Nothing
    Set InfWriter = Nothing
    
MapInfo_Guardar MAPFILE & ".dat"
    
    Set IniManager = Nothing
    
Call Pestañas(MAPFILE & ".map")

frmMain.MousePointer = 0
MapInfo.Changed = 0
    
End Sub


