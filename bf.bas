Attribute VB_Name = "Module1"
Option Explicit
'
'BITMAP FORMAT OVERVIEW
'1) BITMAPFILEHEADER (bmfh)
'2) BITMAPINFOHEADER (bmih)
'3) RGBQUAD          aColors()
'4) BYTE             aBitmapBits() 'this is not avaliable in 24bit bitmaps the aColors replaces this
'
'
'THIS MODULE IS IN BETA STATE
'AND DOESNT SUPPORT THE FOLLOWING:
'1 compressed bitmaps(RLE4,RLE8,JPEG,PNG)
'2 any bitmaps that are not saved in 24-bit DIB
'3 doesnt seem to be working right on some bitmaps
'
'any bugfixes are appreciated,votes also.
'
Public Type BITMAPFILEHEADER
    bfType As Integer       'must be 19778 = "BM"
    bfSize As Long          'size of file in bytes LOF(%bf)
    bfReserved1 As Integer  'Reserved must be set to zero
    bfReserved2 As Integer  'Reserved must be set to zero
    bfOffBits As Long       'the space between this struct and the begining of the actual bmp data
End Type

Public Type BITMAPINFOHEADER '40 bytes
    biSize As Long              'Len(bmih)
    biWidth As Long             'Width of Bitmap Image
    biHeight As Long            'Height of Bitmap Image
    biPlanes As Integer         'Number of Planes for Target Device,must be set to 1
    biBitCount As Integer       'Number of Bits Per Pixel must be either:1(Monochrome),4(16clrs),8(256color),24(RGBQUADS=16777216 colors)
    biCompression As Long       'Compression Modes can be either:BI_bitfields,BI_JPEG,BI_PNG,BI_RLE4,BI_RLE8
    biSizeImage As Long         'Size in bytes of image,can be set to zero if biCompression = BI_RGB
    biXPelsPerMeter As Long     'Horizonal Resolution in Pixels Per Meter
    biYPelsPerMeter As Long     'Vertical Resolution in Pixels Per Meter
    biClrUsed As Long           'the number of colors used by bitmap if its 0 then all colors are used
    biClrImportant As Long      'the number of colors required to display this bitmap if its 0 then their all required
End Type

Public Type RGBTRIBLE
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type

Public Const BI_bitfields = 3&  'UNKNOWN
Public Const BI_JPEG = 4&       'UNKNOWN
Public Const BI_PNG = 5&        'UNKNOWN
Public Const BI_RGB = 0&        '(uncompressed) THIS IS THE ONLY ONE SUPPORTED IN THIS MODULE
Public Const BI_RLE4 = 2&       'RLE RunLength Compression per 4bits(1/2 byte)
Public Const BI_RLE8 = 1&       'RLE RunLength Compression per 8bits(1bytes)

'1) BITMAPFILEHEADER (bmfh)
'2) BITMAPINFOHEADER (bmih)
'3) RGBQUAD          aColors()
'4) BYTE             aBitmapBits()
'
'bmfh,bmih,acolors,abitmapbits
Dim bmfh As BITMAPFILEHEADER
Dim bmih As BITMAPINFOHEADER
Dim aColors() As RGBTRIBLE
'Dim aColors() as RGBQUAD
Dim aBitmapBits() As Byte



Function LoadBitmapImage(strPath$, picOut As PictureBox)

    Dim F%
    F = FreeFile
    Open strPath For Binary Access Read As F
        Get F, , bmfh
        Get F, , bmih
        Seek F, bmfh.bfOffBits + 1
        Select Case bmih.biBitCount
            Case 1 '1bit per pixel, 8pixels per byte
        
                'MONOCHROME, that means 8pixels per
                'byte (1bit per pixel)
                'we CAN load this format
                'but it will be real slow
                'considering that its already slow
            Case 4 '4bits per pixel, 2 pixel per byte
            
                '16 Color bitmaps
                'thats a little faster i guess
                'since u can use the Hex$ function
            Case 8 '8bits(1byte) per pixel,1 pixel per byte
                '256 GrayScale/Colors
                '
            Case 16 '2bytes per pixel,a pixel is stored as an integer
                'about 32 thousand colors(32,768) in exact to be exact
                '
            Case 24 '3bytes per pixel,no color table is used, 1 pixel per RGBQUAD structure
                picOut.AutoRedraw = True
'                Debug.Print Seek(F), LOF(F), "BEFORE BEGIN"
                ReDim Preserve aColors(1 To (bmih.biWidth * bmih.biHeight)) As RGBTRIBLE
                Get F, , aColors
                Dim cy&, cx&
                Dim c&
                For cy = bmih.biHeight To 1 Step -1
                    For cx = 1 To bmih.biWidth
                        c = c + 1
                        picOut.PSet (cx, cy), RGB(aColors(c).rgbRed, aColors(c).rgbGreen, aColors(c).rgbBlue)
                    Next
                Next
                picOut.Refresh
            End Select
'            Debug.Print Seek(F), LOF(F), bmfh.bfOffBits
    Close F
End Function


