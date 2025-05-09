Attribute VB_Name = "modRender"
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
' modRender
'
' @author Torres Patricio (Pato)
' @version 1.0.0
' @date 20110312

Option Explicit

Public Enum eFormatPic
    bmp
    png
    jpg
End Enum

Public Sub MapCapture(ByRef format As eFormatPic, ByVal Size As Long)
'*************************************************
'Author: Torres Patricio(Pato)
'Last modified:12/03/11
'*************************************************
Dim y           As Long     'Keeps track of where on map we are
Dim X           As Long     'Keeps track of where on map we are
Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
Dim ScreenXOffset   As Integer
Dim ScreenYOffset   As Integer
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer 'For centering grhs
Dim Grh         As Grh      'Temp Grh for show tile and blocked
Dim renderSurface As DirectDrawSurface7
Dim surfaceDesc As DDSURFACEDESC2
Dim srcRect As RECT
Dim destRect As RECT


    With surfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        If ClientSetup.bUseVideo Then
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        Else
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        End If
        .lHeight = 3200 * 3 '32(Tama�o del pixel)*100(Ancho en pixeles)*100(Alto en pixeles)
        .lWidth = 3200 * 3
        
        Set renderSurface = DirectDraw.CreateSurface(surfaceDesc)
    End With

    With srcRect
        .Right = 3200 * 3
        .Bottom = 3200 * 3
    End With
    
    Call renderSurface.BltColorFill(srcRect, 0)
    
    frmRender.pgbProgress.value = 0
    frmRender.pgbProgress.Max = 50000
    
    'Draw floor layer
    For y = 1 To 300
        For X = 1 To 300
            
            'Layer 1 **********************************
            If MapData(X, y).Graphic(1).GrhIndex <> 0 Then
                Call DDrawGrhtoSurface(renderSurface, MapData(X, y).Graphic(1), _
                    (X - 1) * TilePixelWidth, _
                    (y - 1) * TilePixelHeight, _
                    0, 1)
            End If
            '******************************************
            
'            frmRender.pgbProgress.value = frmRender.pgbProgress.value + 1
            frmRender.lblEstado.Caption = "Renderizado de primer capa - " & (y - 1) + (X / 100) & "%"
            DoEvents
        Next X
    Next y
    
    'Draw floor layer 2
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            'Layer 2 **********************************
            If (MapData(X, y).Graphic(2).GrhIndex <> 0) And bVerCapa(2) Then
                Call DDrawTransGrhtoSurface(renderSurface, MapData(X, y).Graphic(2), _
                        (X - 1) * TilePixelWidth, _
                        (y - 1) * TilePixelHeight, _
                        1, 1)
            End If
            '******************************************
            
'            frmRender.pgbProgress.value = frmRender.pgbProgress.value + 1
            frmRender.lblEstado = "Renderizado de segunda capa - " & (y - 1) + (X / 100) & "%"
            DoEvents
        Next X
    Next y
    
    'Draw Transparent Layers
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            PixelOffsetXTemp = (X - 1) * TilePixelWidth
            PixelOffsetYTemp = (y - 1) * TilePixelHeight
            
            With MapData(X, y)
                'Object Layer **********************************
                If (.ObjGrh.GrhIndex <> 0) And bVerObjetos Then
                    Call DDrawTransGrhtoSurface(renderSurface, .ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
                '***********************************************
                
                'Layer 3 *****************************************
                If (.Graphic(3).GrhIndex <> 0) And bVerCapa(3) Then
                    'Draw
                    Call DDrawTransGrhtoSurface(renderSurface, .Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
                '************************************************
                
            '    frmRender.pgbProgress.value = frmRender.pgbProgress.value + 1
                frmRender.lblEstado.Caption = "Renderizado de objetos y tercer capa - " & (y - 1) + (X / 100) & "%"
                DoEvents
            End With
        Next X
    Next y
    
    Grh.FrameCounter = 1
    Grh.Started = 0

    
    destRect.Right = Size * 3
    destRect.Bottom = Size * 3
     
    frmRender.picMap.Width = Size * 3
    frmRender.picMap.Height = Size * 3

    Call renderSurface.BltToDC(frmRender.picMap.hdc, srcRect, destRect)
    frmRender.picMap.Picture = frmRender.picMap.image
    
    If Not FileExist(App.path & "\Screenshots", vbDirectory) Then MkDir (App.path & "\Screenshots")
    
    Select Case format
        Case eFormatPic.bmp
            Call SavePicture(frmRender.picMap.image, App.path & "\Screenshots\" & MapInfo.Name & ".bmp")
            
        Case eFormatPic.png
            Call StartUpGDIPlus(GdiPlusVersion)
            Call SavePictureAsPNG(frmRender.picMap.Picture, App.path & "\Screenshots\" & MapInfo.Name & ".png")
            Call ShutdownGDIPlus
            
        Case eFormatPic.jpg
            Call StartUpGDIPlus(GdiPlusVersion)
            Call SavePictureAsJPG(frmRender.picMap.Picture, App.path & "\Screenshots\" & MapInfo.Name & ".jpg")
            Call ShutdownGDIPlus
    End Select
End Sub


