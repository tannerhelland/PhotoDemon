VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdFxBilateral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Bilateral Filter
'Copyright 2014-2026 by Tanner Helland
'Created: 19/June/14
'Last updated: 13/December/19
'Last update: final round of minor perf and quality improvements to the recursive BF algorithm
'Additional dependencies: many; see the VBHacks module in particular, as this class uses it liberally
'Additional licenses: includes code adapted from an MIT-licensed project by Ming (https://github.com/ufoym/recursive-bf)
'
'Per Wikipedia (https://en.wikipedia.org/wiki/Bilateral_filter):
' "A bilateral filter is a non-linear, edge-preserving, and noise-reducing smoothing filter for images.
' It replaces the intensity of each pixel with a weighted average of intensity values from nearby pixels.
' This weight can be based on a Gaussian distribution. Crucially, the weights depend not only on
' Euclidean distance of pixels, but also on the radiometric differences (e.g., range differences, such as
' color intensity, depth distance, etc.). This preserves sharp edges."
'
'More details on bilateral filtering can be found at:
' http://www.cs.duke.edu/~tomasi/papers/tomasi/tomasiIccv98.pdf
'
'Because traditional 2D kernel convolution is extremely slow on images of any size, PhotoDemon used
' a separable bilateral filter implementation for many years.  This provided a good approximation of
' a true bilateral, and it transformed the filter from an O(w*h*r^2) process to O(w*h*2r).
'
'For details on a separable bilateral approach, see:
' http://homepage.tudelft.nl/e3q6n/publications/2005/ICME2005_TPLV.pdf
'
'In 2019, I bit the bullet and translated a (lengthy, complicated) recursive bilateral filter
' implementation into VB6.  This is the current state-of-the-art for real-time bilateral filtering.
' It was developed by Qingxiong Yang and first published in an influential 2012 paper:
' https://link.springer.com/content/pdf/10.1007%2F978-3-642-33718-5_29.pdf
'
'This technique reduces the filter to a constant-time filter of just O(w*h).
'
'PD's implementation is based on a 2017 C++ implementation of Yang's work by Ming:
'
'https://github.com/ufoym/recursive-bf
'
'Ming's code is available under an MIT license.  Thank you to him/her/them for sharing their work!
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Event ProgressUpdate(ByVal progressValue As Single, ByRef cancelOperation As Boolean)

'The recursive algorithm allocates a (fairly large) internal buffer as part of its work.
' It is obviously more effective to reuse this buffer (when possible) instead of re-allocating it.
' The recursive function automatically determines when this buffer needs to be regenerated,
' but to free it, you must currently free the entire class.
Private m_Buffer() As Single, m_BufferSize As Long

'Recursive bilateral filter implementation
' Reference: Qingxiong Yang, Recursive Bilateral Filtering, European Conference on Computer Vision (ECCV) 2012, 399-413.
'
'Original C++ implementation by Ming, MIT-licensed: https://github.com/ufoym/recursive-bf
' Translated into VB6 by Tanner Helland in December 2019
'
' INPUTS:
' 1) source DIB; importantly, *this function performs the bilateral in-place* so this DIB will be modified!
' 2) kernelRadius: Any float >= 1
' 3) rangeFactor: any float on the range [0, 100]; it will be translated into a relevant range factor
'                 based on the source image's color depth
' 4) raiseProgressEvents: if TRUE, this class will raise periodic progress events containing a progress
'                          "factor" on the range [0.0, 1.0]
Public Function Bilateral_Recursive(ByRef srcDIB As pdDIB, ByVal kernelRadius As Double, ByVal rangeFactor As Double, Optional ByVal raiseProgressEvents As Boolean = False) As Long
    
    'For perf reasons, this function is locked to 32-bpp inputs
    If (srcDIB.GetDIBColorDepth <> 32) Then Exit Function
    
    'This filter is interruptible
    Dim cancelOp As Boolean
    cancelOp = False
    
    'If the caller wants progress reports, calculate some max/min progress values
    Dim progCheck As Long, progMax As Single
    If raiseProgressEvents Then
        progMax = srcDIB.GetDIBHeight * 3
        progCheck = ProgressBars.FindBestProgBarValue(progMax)
        progMax = (1! / progMax)
    End If
    
    'Enforce a minimum radius; below this, effects won't be noticeable (and accuracy vs
    ' a true bilateral filter declines)
    If (kernelRadius < 1#) Then kernelRadius = 1#
    
    'PD displays spatial and color (range) factors on a [0, 100] scale.
    ' In our standard bilateral filter tool, these are normalized like this:
    rangeFactor = (rangeFactor * 2.55) / 5#
    
    'We repeat that normalization here.  (It limits the strength to 1/5 it's max amount,
    ' where the max amount approximates an actual gaussian [e.g. range is irrelevant].
    ' Note that this is especially relevant for this tool, as it does *not* approximate
    ' a gaussian well as the range factor increases.)
    
    'Convert the supplied radius value to a standard gaussian sigma value, using a similar
    ' formula to ImageJ (per this link: http://stackoverflow.com/questions/21984405/relation-between-sigma-and-radius-on-the-gaussian-blur)
    ' The idea here is to convert the radius to a sigma of sufficient magnitude where the outer edges
    ' of the gaussian no longer represent meaningful values on a [0, 255] scale.  (For 16-bit images,
    ' you would want to calculate e.g. Log(65535) or similar.)
    Dim sigma As Double
    Const LOG_255_BASE_10 As Double = 2.40654018043395
    sigma = (kernelRadius + 1#) / Sqr(2# * LOG_255_BASE_10)
    
    'The original recursive BF implementation auto-normalizes the radius against the
    ' image's dimensions (width and height are scaled separately).
    
    'I have mixed feelings about the usefulness of this, especially when working on
    ' images with e.g. an aspect ratio larger than 2:1... but it *is* useful for
    ' tasks like producing usable previews (because we don't need to manually scale the
    ' preview radius to make it match the radius-at-full-size value).
    
    'Anyway, to make this algorithm behave more closely to PD's standard implementation,
    ' I use a standard normalization approach.  The supplied gaussian sigma is used
    ' 1:1 as the spatial component.
    Dim sigma_spatial As Double, sigma_range As Double
    sigma_spatial = sigma
    
    'Color (range) factor is *not* scaled according to the current radius, obviously;
    ' instead, it is used like a threshold against the maximum pixel difference at the
    ' current color depth (255 for 8-bit images) - see below for details.
    sigma_range = rangeFactor
    
    'Want to test hard-coded spatial and-or range values (e.g. to compare to the reference
    ' implementation)?  You can do so here, but note that you will also need to set the
    ' useReferenceScale flag to TRUE; this modifies two subsequent calculations that
    ' auto-scale the spatial value against the width and height of the image.
    Dim useReferenceScale As Boolean
    useReferenceScale = False
    If useReferenceScale Then
        sigma_spatial = 0.5     'Arbitrary testing value
        sigma_range = 0.5       'Arbitrary testing value
    End If
    
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = srcDIB.GetDIBWidth
    imgHeight = srcDIB.GetDIBHeight
    
    'The recursive function also requires a (large) internal floating-point buffer
    Dim channels As Long
    channels = srcDIB.GetDIBColorDepth \ 8
    
    Dim width_height As Long
    width_height = imgWidth * imgHeight
    Dim width_channel As Long
    width_channel = srcDIB.GetDIBStride
    Dim width_height_channel As Long
    width_height_channel = imgHeight * width_channel
    
    'We allocate the buffer only when its required size changes; this greatly reduces churn
    ' during preview stages
    Dim bufferSize As Long
    bufferSize = (width_height_channel + width_height + width_channel + imgWidth) * 2
    If (m_BufferSize <> bufferSize) Then
        ReDim m_Buffer(0 To bufferSize - 1) As Single
        m_BufferSize = bufferSize
    End If
    
    'We are now going to overlay a bunch of temporary arrays over that (large!) internal buffer.
    ' Note that these *MUST BE SPECIALLY FREED* when we are done with them, as they aren't actual
    ' allocations - they're just pointers into the buffer allocated above.
    Dim sizeRemaining As Long
    sizeRemaining = bufferSize
    
    Dim img_out_f() As Single, img_out_f_SA As SafeArray1D
    VBHacks.WrapArrayAroundPtr_Float img_out_f, img_out_f_SA, VarPtr(m_Buffer(0)), sizeRemaining * 4
    'float * img_out_f = buffer;
    
    Dim img_temp() As Single, img_tmp_SA As SafeArray1D
    sizeRemaining = sizeRemaining - width_height_channel
    VBHacks.WrapArrayAroundPtr_Float img_temp, img_tmp_SA, VarPtr(m_Buffer(bufferSize - sizeRemaining)), sizeRemaining * 4
    'float * img_temp = &img_out_f[width_height_channel];
    
    Dim map_factor_a() As Single, map_factor_a_SA As SafeArray1D
    sizeRemaining = sizeRemaining - width_height_channel
    VBHacks.WrapArrayAroundPtr_Float map_factor_a, map_factor_a_SA, VarPtr(m_Buffer(bufferSize - sizeRemaining)), sizeRemaining * 4
    'float * map_factor_a = &img_temp[width_height_channel];
    
    Dim map_factor_b() As Single, map_factor_b_SA As SafeArray1D
    sizeRemaining = sizeRemaining - width_height
    VBHacks.WrapArrayAroundPtr_Float map_factor_b, map_factor_b_SA, VarPtr(m_Buffer(bufferSize - sizeRemaining)), sizeRemaining * 4
    'float * map_factor_b = &map_factor_a[width_height];
    
    Dim slice_factor_a() As Single, slice_factor_a_SA As SafeArray1D
    sizeRemaining = sizeRemaining - width_height
    VBHacks.WrapArrayAroundPtr_Float slice_factor_a, slice_factor_a_SA, VarPtr(m_Buffer(bufferSize - sizeRemaining)), sizeRemaining * 4
    'float * slice_factor_a = &map_factor_b[width_height];
    
    Dim slice_factor_b() As Single, slice_factor_b_SA As SafeArray1D
    sizeRemaining = sizeRemaining - width_channel
    VBHacks.WrapArrayAroundPtr_Float slice_factor_b, slice_factor_b_SA, VarPtr(m_Buffer(bufferSize - sizeRemaining)), sizeRemaining * 4
    'float * slice_factor_b = &slice_factor_a[width_channel];
    
    Dim line_factor_a() As Single, line_factor_a_SA As SafeArray1D
    sizeRemaining = sizeRemaining - width_channel
    VBHacks.WrapArrayAroundPtr_Float line_factor_a, line_factor_a_SA, VarPtr(m_Buffer(bufferSize - sizeRemaining)), sizeRemaining * 4
    'float * line_factor_a = &slice_factor_b[width_channel];
    
    Dim line_factor_b() As Single, line_factor_b_SA As SafeArray1D
    sizeRemaining = sizeRemaining - imgWidth
    VBHacks.WrapArrayAroundPtr_Float line_factor_b, line_factor_b_SA, VarPtr(m_Buffer(bufferSize - sizeRemaining)), sizeRemaining * 4
    'float * line_factor_b = &line_factor_a[width];
    
    'Wrap an unsigned char array around the source pixel data
    Dim img() As Byte, imgPixelsSA As SafeArray1D
    srcDIB.WrapArrayAroundDIB_1D img, imgPixelsSA
    
    'Compute a lookup table
    Const QX_DEF_CHAR_MAX As Long = 255
    Dim range_table(0 To QX_DEF_CHAR_MAX) As Single
    'float range_table[QX_DEF_CHAR_MAX + 1];
    
    Dim inv_sigma_range As Single
    If useReferenceScale Then
        inv_sigma_range = 1! / (sigma_range * QX_DEF_CHAR_MAX)
        'range_table[i] = static_cast<float>(exp(-i * inv_sigma_range));
    Else
        'We've already scaled the range factor; just divide as-is
        inv_sigma_range = 1! / sigma_range
    End If
    
    Dim i As Long
    For i = 0 To QX_DEF_CHAR_MAX
        range_table(i) = Exp(-i * inv_sigma_range)
    Next i
    
    Dim tmpAlpha As Single
    If useReferenceScale Then
    
        'As mentioned earlier, the original function auto-scales this alpha calculation value
        ' against the width of the image.  This means that e.g. a sigma of 1.0 would be equal
        ' to a radius of imgWidth, as per this line of code:
        tmpAlpha = Exp(-Sqr(2#) / (sigma_spatial * imgWidth))
        'float alpha = static_cast<float>(exp(-sqrt(2.0) / (sigma_spatial * width)));
    
    Else
    
        'Rather than do that, we treat sigma_spatial as a ratio against the user-supplied
        ' radius value.  It has already been normalized earlier in this function.
        tmpAlpha = Exp(-Sqr(2#) / sigma_spatial)
        
    End If
    
    'Note that /a values have been added for alpha channel handling
    Dim ypr As Single, ypg As Single, ypb As Single, ycr As Single, ycg As Single, ycb As Single
    Dim ypa As Single, yca As Single
    Dim fp As Single, fc As Single
    Dim inv_alpha_ As Single
    inv_alpha_ = 1! - tmpAlpha
    
    'Originally declared in inner loops
    Dim tcr As Long, tcg As Long, tcb As Long, tca As Long
    Dim dR As Long, dG As Long, dB As Long
    Dim temp_x As Long, in_x As Long, texture_x As Long
    Dim tpr As Long, tpg As Long, tpb As Long, tpa As Long
    Dim temp_factor_x As Long, range_dist As Long
    Dim weight As Single, alpha_ As Single
    
    Dim x As Long, y As Long
    For y = 0 To imgHeight - 1
        
        temp_x = y * width_channel
        'float * temp_x = &img_temp[y * width_channel];
        
        in_x = y * width_channel
        'unsigned char * in_x = &img[y * width_channel];
        
        texture_x = y * width_channel
        'unsigned char * texture_x = &img[y * width_channel];
        
        ypb = img(in_x)
        img_temp(temp_x) = ypb
        in_x = in_x + 1
        temp_x = temp_x + 1
        '*temp_x++ = ypr = *in_x++;
        
        ypg = img(in_x)
        img_temp(temp_x) = ypg
        in_x = in_x + 1
        temp_x = temp_x + 1
        '*temp_x++ = ypg = *in_x++;
        
        ypr = img(in_x)
        img_temp(temp_x) = ypr
        in_x = in_x + 1
        temp_x = temp_x + 1
        '*temp_x++ = ypb = *in_x++;
        
        ypa = img(in_x)
        img_temp(temp_x) = ypa
        in_x = in_x + 1
        temp_x = temp_x + 1
        
        tpb = img(texture_x)
        texture_x = texture_x + 1
        'unsigned char tpr = *texture_x++;
        
        tpg = img(texture_x)
        texture_x = texture_x + 1
        'unsigned char tpg = *texture_x++;
        
        tpr = img(texture_x)
        texture_x = texture_x + 1
        'unsigned char tpb = *texture_x++;
        
        tpa = img(texture_x)
        texture_x = texture_x + 1
        
        temp_factor_x = y * imgWidth
        'float * temp_factor_x = &map_factor_a[y * width];
        
        fp = 1!
        map_factor_a(temp_factor_x) = fp
        temp_factor_x = temp_factor_x + 1
        '*temp_factor_x++ = fp = 1;
        
        'causal
        For x = 1 To imgWidth - 1
            
            tcb = img(texture_x)
            texture_x = texture_x + 1
            'unsigned char tcr = *texture_x++;
            
            tcg = img(texture_x)
            texture_x = texture_x + 1
            'unsigned char tcg = *texture_x++;
            
            tcr = img(texture_x)
            texture_x = texture_x + 1
            'unsigned char tcb = *texture_x++;
            
            tca = img(texture_x)
            texture_x = texture_x + 1
            
            dR = Abs(tcr - tpr)
            dG = Abs(tcg - tpg)
            dB = Abs(tcb - tpb)
            
            'Reference implementation is modified to emphasize green over red (per luminance)
            range_dist = (dR + dG * 2 + dB) \ 4
            'int range_dist = (((dr << 1) + dg + db) >> 2);
            
            weight = range_table(range_dist)
            alpha_ = weight * tmpAlpha
            
            ycb = inv_alpha_ * img(in_x) + alpha_ * ypb
            img_temp(temp_x) = ycb
            in_x = in_x + 1
            temp_x = temp_x + 1
            '*temp_x++ = ycr = inv_alpha_*(*in_x++) + alpha_*ypr;
            
            ycg = inv_alpha_ * img(in_x) + alpha_ * ypg
            img_temp(temp_x) = ycg
            in_x = in_x + 1
            temp_x = temp_x + 1
            '*temp_x++ = ycg = inv_alpha_*(*in_x++) + alpha_*ypg;
            
            ycr = inv_alpha_ * img(in_x) + alpha_ * ypr
            img_temp(temp_x) = ycr
            in_x = in_x + 1
            temp_x = temp_x + 1
            '*temp_x++ = ycb = inv_alpha_*(*in_x++) + alpha_*ypb;
            
            yca = inv_alpha_ * img(in_x) + alpha_ * ypa
            img_temp(temp_x) = yca
            in_x = in_x + 1
            temp_x = temp_x + 1
            
            tpr = tcr
            tpg = tcg
            tpb = tcb
            tpa = tca
            ypr = ycr
            ypg = ycg
            ypb = ycb
            ypa = yca
            
            fc = inv_alpha_ + alpha_ * fp
            map_factor_a(temp_factor_x) = fc
            temp_factor_x = temp_factor_x + 1
            '*temp_factor_x++ = fc = inv_alpha_ + alpha_*fp;
            
            fp = fc
            
        Next x
        
        temp_x = temp_x - 1
        in_x = in_x - 1
        img_temp(temp_x) = 0.5! * (img_temp(temp_x) + img(in_x))
        '*--temp_x; *temp_x = 0.5f*((*temp_x) + (*--in_x));
        
        temp_x = temp_x - 1
        in_x = in_x - 1
        img_temp(temp_x) = 0.5! * (img_temp(temp_x) + img(in_x))
        '*--temp_x; *temp_x = 0.5f*((*temp_x) + (*--in_x));
        
        temp_x = temp_x - 1
        in_x = in_x - 1
        img_temp(temp_x) = 0.5! * (img_temp(temp_x) + img(in_x))
        '*--temp_x; *temp_x = 0.5f*((*temp_x) + (*--in_x));
        
        temp_x = temp_x - 1
        in_x = in_x - 1
        img_temp(temp_x) = 0.5! * (img_temp(temp_x) + img(in_x))
        
        texture_x = texture_x - 1
        tpb = img(texture_x)
        'tpr = *--texture_x;
        
        texture_x = texture_x - 1
        tpg = img(texture_x)
        'tpg = *--texture_x;
        
        texture_x = texture_x - 1
        tpr = img(texture_x)
        'tpb = *--texture_x;
        
        texture_x = texture_x - 1
        tpa = img(texture_x)
        
        'Fix by Tanner - use correct offsets for right-boundary pixels!
        ypa = img(in_x)
        ypr = img(in_x + 1)
        ypg = img(in_x + 2)
        ypb = img(in_x + 3)
        'ypr = *in_x; ypg = *in_x; ypb = *in_x;
        
        temp_factor_x = temp_factor_x - 1
        map_factor_a(temp_factor_x) = 0.5! * (map_factor_a(temp_factor_x) + 1)
        '*--temp_factor_x; *temp_factor_x = 0.5f*((*temp_factor_x) + 1);
        
        fp = 1!
        
        'anticausal
        For x = imgWidth - 2 To 0 Step -1
            
            texture_x = texture_x - 1
            tcb = img(texture_x)
            'unsigned char tcr = *--texture_x;
            
            texture_x = texture_x - 1
            tcg = img(texture_x)
            'unsigned char tcg = *--texture_x;
            
            texture_x = texture_x - 1
            tcr = img(texture_x)
            'unsigned char tcb = *--texture_x;
            
            texture_x = texture_x - 1
            tca = img(texture_x)
            
            dR = Abs(tcr - tpr)
            dG = Abs(tcg - tpg)
            dB = Abs(tcb - tpb)
            
            'Reference implementation is modified to emphasize green over red (per luminance)
            range_dist = (dR + dG * 2 + dB) \ 4
            'int range_dist = (((dr << 1) + dg + db) >> 2);
            
            weight = range_table(range_dist)
            alpha_ = weight * tmpAlpha
            
            in_x = in_x - 1
            ycb = inv_alpha_ * img(in_x) + alpha_ * ypb
            'ycr = inv_alpha_ * (*--in_x) + alpha_ * ypr;
            
            in_x = in_x - 1
            ycg = inv_alpha_ * img(in_x) + alpha_ * ypg
            'ycg = inv_alpha_ * (*--in_x) + alpha_ * ypg;
            
            in_x = in_x - 1
            ycr = inv_alpha_ * img(in_x) + alpha_ * ypr
            'ycb = inv_alpha_ * (*--in_x) + alpha_ * ypb;
            
            in_x = in_x - 1
            yca = inv_alpha_ * img(in_x) + alpha_ * ypa
            
            temp_x = temp_x - 1
            img_temp(temp_x) = 0.5 * (img_temp(temp_x) + ycb)
            '*--temp_x; *temp_x = 0.5f*((*temp_x) + ycr);
            
            temp_x = temp_x - 1
            img_temp(temp_x) = 0.5 * (img_temp(temp_x) + ycg)
            '*--temp_x; *temp_x = 0.5f*((*temp_x) + ycg);
            
            temp_x = temp_x - 1
            img_temp(temp_x) = 0.5 * (img_temp(temp_x) + ycr)
            '*--temp_x; *temp_x = 0.5f*((*temp_x) + ycb);
            
            temp_x = temp_x - 1
            img_temp(temp_x) = 0.5 * (img_temp(temp_x) + yca)
            
            tpr = tcr
            tpg = tcg
            tpb = tcb
            tpa = tca
            'tpr = tcr; tpg = tcg; tpb = tcb;
            ypr = ycr
            ypg = ycg
            ypb = ycb
            ypa = yca
            'ypr = ycr; ypg = ycg; ypb = ycb;
            
            fc = inv_alpha_ + alpha_ * fp
            temp_factor_x = temp_factor_x - 1
            
            map_factor_a(temp_factor_x) = 0.5! * (map_factor_a(temp_factor_x) + fc)
            '*temp_factor_x = 0.5f*((*temp_factor_x) + fc);
            
            fp = fc
            
        Next x
        
        'UI updates
        If raiseProgressEvents Then
            If ((y And progCheck) = 0) Then
                RaiseEvent ProgressUpdate((CDbl(y) * progMax), cancelOp)
                If cancelOp Then GoTo FilterCleanup
            End If
        End If
        
    Next y
    
    'See earlier comments on scaling the spatial parameter against image dimensions.
    If useReferenceScale Then
    
        'The reference implementation uses this formula:
        tmpAlpha = Exp(-Sqr(2#) / (sigma_spatial * imgHeight))
        'alpha = static_cast<float>(exp(-sqrt(2.0) / (sigma_spatial * height)));
    
    Else
    
        'We use the user-supplied radius value, instead:
        tmpAlpha = Exp(-Sqr(2#) / sigma_spatial)
        
    End If
    
    inv_alpha_ = 1! - tmpAlpha
    
    Dim ycy As Long, ypy As Long, xcy As Long
    Dim tcy As Long, tpy As Long
    
    CopyMemoryStrict VarPtr(img_out_f(0)), VarPtr(img_temp(0)), width_channel * 4
    'memcpy(img_out_f, img_temp, sizeof(float)* width_channel);
    
    '(Manually replaced with VarPtr() instances below)
    'Dim in_factor As Long
    'float * in_factor = map_factor_a;
    
    Dim ycf As Long, ypf As Long, xcf As Long
    
    CopyMemoryStrict VarPtr(map_factor_b(0)), VarPtr(map_factor_a(0)), 4 * imgWidth
    'memcpy(map_factor_b, in_factor, sizeof(float) * width);
    
    For y = 1 To imgHeight - 1
    
        tpy = (y - 1) * width_channel
        'tpy = &img[(y - 1) * width_channel];
        
        tcy = y * width_channel
        'tcy = &img[y * width_channel];
        
        xcy = y * width_channel
        'xcy = &img_temp[y * width_channel];
        
        ypy = (y - 1) * width_channel
        'ypy = &img_out_f[(y - 1) * width_channel];
        
        ycy = y * width_channel
        'ycy = &img_out_f[y * width_channel];
        
        xcf = y * imgWidth
        'xcf = &in_factor[y * width];
        
        ypf = (y - 1) * imgWidth
        'ypf = &map_factor_b[(y - 1) * width];
        
        ycf = y * imgWidth
        'ycf = &map_factor_b[y * width];
        
        For x = 0 To imgWidth - 1
            
            dB = Abs(CLng(img(tcy)) - CLng(img(tpy)))
            tcy = tcy + 1
            tpy = tpy + 1
            'unsigned char dr = abs((*tcy++) - (*tpy++));
            
            dG = Abs(CLng(img(tcy)) - CLng(img(tpy)))
            tcy = tcy + 1
            tpy = tpy + 1
            'unsigned char dg = abs((*tcy++) - (*tpy++));
            
            dR = Abs(CLng(img(tcy)) - CLng(img(tpy)))
            tcy = tcy + 1
            tpy = tpy + 1
            'unsigned char db = abs((*tcy++) - (*tpy++));
            
            'alpha distance isn't required - just advance the pointer
            tcy = tcy + 1
            tpy = tpy + 1
            
            'Reference implementation is modified to emphasize green over red (per luminance)
            range_dist = (dR + dG * 2 + dB) \ 4
            'int range_dist = (((dr << 1) + dg + db) >> 2);
            
            weight = range_table(range_dist)
            alpha_ = weight * tmpAlpha
            
            'Original code uses a per-channel loop here; instead, we unroll it manually
            img_out_f(ycy) = inv_alpha_ * img_temp(xcy) + alpha_ * img_out_f(ypy)
            ycy = ycy + 1: xcy = xcy + 1: ypy = ypy + 1
            img_out_f(ycy) = inv_alpha_ * img_temp(xcy) + alpha_ * img_out_f(ypy)
            ycy = ycy + 1: xcy = xcy + 1: ypy = ypy + 1
            img_out_f(ycy) = inv_alpha_ * img_temp(xcy) + alpha_ * img_out_f(ypy)
            ycy = ycy + 1: xcy = xcy + 1: ypy = ypy + 1
            img_out_f(ycy) = inv_alpha_ * img_temp(xcy) + alpha_ * img_out_f(ypy)
            ycy = ycy + 1: xcy = xcy + 1: ypy = ypy + 1
            '*ycy++ = inv_alpha_*(*xcy++) + alpha_*(*ypy++);
            
            map_factor_b(ycf) = inv_alpha_ * map_factor_a(xcf) + alpha_ * map_factor_b(ypf)
            ycf = ycf + 1
            xcf = xcf + 1
            ypf = ypf + 1
            '*ycf++ = inv_alpha_*(*xcf++) + alpha_*(*ypf++);
        
        Next x
        
        'UI updates
        If raiseProgressEvents Then
            If (((y + imgHeight) And progCheck) = 0) Then
                RaiseEvent ProgressUpdate((CDbl(y + imgHeight) * progMax), cancelOp)
                If cancelOp Then GoTo FilterCleanup
            End If
        End If
        
    Next y
    
    Dim h1 As Long
    h1 = imgHeight - 1
    
    ycf = 0
    ypf = 0
    CopyMemoryStrict VarPtr(line_factor_b(0)), VarPtr(map_factor_a(h1 * imgWidth)), 4 * imgWidth
    'ycf = line_factor_a;
    'ypf = line_factor_b;
    'memcpy(ypf, &in_factor[h1 * width], sizeof(float) * width);
    
    For x = 0 To imgWidth - 1
        
        map_factor_b(h1 * imgWidth + x) = 0.5! * (map_factor_b(h1 * imgWidth + x) + line_factor_b(x))
        'map_factor_b[h1 * width + x] = 0.5f*(map_factor_b[h1 * width + x] + ypf[x]);
        
    Next x
    
    ycy = 0
    ypy = 0
    CopyMemoryStrict VarPtr(slice_factor_b(0)), VarPtr(img_temp(h1 * width_channel)), 4 * width_channel
    'ycy = slice_factor_a;
    'ypy = slice_factor_b;
    'memcpy(ypy, &img_temp[h1 * width_channel], sizeof(float)* width_channel);
    
    Dim k As Long, mapFactorBCache As Single
    k = 0
    
    Dim idx As Long
    For x = 0 To imgWidth - 1
        
        'Original code uses a per-channel loop here; we unroll it manually
        mapFactorBCache = 1! / map_factor_b(h1 * imgWidth + x)
        idx = (h1 * imgWidth + x) * channels
        img_out_f(idx) = 0.5! * (img_out_f(idx) + slice_factor_b(k)) * mapFactorBCache
        k = k + 1
        idx = (h1 * imgWidth + x) * channels + 1
        img_out_f(idx) = 0.5! * (img_out_f(idx) + slice_factor_b(k)) * mapFactorBCache
        k = k + 1
        idx = (h1 * imgWidth + x) * channels + 2
        img_out_f(idx) = 0.5! * (img_out_f(idx) + slice_factor_b(k)) * mapFactorBCache
        k = k + 1
        idx = (h1 * imgWidth + x) * channels + 3
        img_out_f(idx) = 0.5! * (img_out_f(idx) + slice_factor_b(k)) * mapFactorBCache
        k = k + 1
        'img_out_f[idx] = 0.5f*(img_out_f[idx] + ypy[k++]) / map_factor_b[h1 * width + x];
        
    Next x
    
    Dim ycy_ As Long, ypy_ As Long, out_ As Long
    Dim ycf_ As Long, ypf_ As Long, factor_ As Long
    Dim fcc As Single, ycc As Single
        
    For y = h1 - 1 To 0 Step -1
        
        tpy = (y + 1) * width_channel
        'tpy = &img[(y + 1) * width_channel];
        
        tcy = y * width_channel
        'tcy = &img[y * width_channel];
        
        xcy = y * width_channel
        'xcy = &img_temp[y * width_channel];
        
        ycy_ = ycy
        ypy_ = ypy
        
        out_ = y * width_channel
        'float*out_ = &img_out_f[y * width_channel];
        
        xcf = y * imgWidth
        'xcf = &in_factor[y * width];
        
        ycf_ = ycf
        ypf_ = ypf
        
        factor_ = y * imgWidth
        'float*factor_ = &map_factor_b[y * width];
        
        For x = 0 To imgWidth - 1
            
            dB = Abs(CLng(img(tcy)) - CLng(img(tpy)))
            tcy = tcy + 1
            tpy = tpy + 1
            'unsigned char dr = abs((*tcy++) - (*tpy++));
            
            dG = Abs(CLng(img(tcy)) - CLng(img(tpy)))
            tcy = tcy + 1
            tpy = tpy + 1
            'unsigned char dg = abs((*tcy++) - (*tpy++));
            
            dR = Abs(CLng(img(tcy)) - CLng(img(tpy)))
            tcy = tcy + 1
            tpy = tpy + 1
            'unsigned char db = abs((*tcy++) - (*tpy++));
            
            'alpha distance isn't required - just advance the pointer
            tcy = tcy + 1
            tpy = tpy + 1
            
            'Reference implementation is modified to emphasize green over red (per luminance)
            range_dist = (dR + dG * 2 + dB) \ 4
            'int range_dist = (((dr << 1) + dg + db) >> 2);
            
            weight = range_table(range_dist)
            alpha_ = weight * tmpAlpha
            
            fcc = inv_alpha_ * map_factor_a(xcf) + alpha_ * line_factor_b(ypf_)
            xcf = xcf + 1
            ypf_ = ypf_ + 1
            'float fcc = inv_alpha_*(*xcf++) + alpha_*(*ypf_++);
            
            line_factor_a(ycf_) = fcc
            ycf_ = ycf_ + 1
            '*ycf_++ = fcc;
            
            map_factor_b(factor_) = 0.5! * (map_factor_b(factor_) + fcc)
            '*factor_ = 0.5f * (*factor_ + fcc);
            
            'Original code uses a per-channel loop here; we manually unroll it
            mapFactorBCache = 1! / map_factor_b(factor_)
            
            ycc = inv_alpha_ * img_temp(xcy) + alpha_ * slice_factor_b(ypy_)
            xcy = xcy + 1
            ypy_ = ypy_ + 1
            'float ycc = inv_alpha_*(*xcy++) + alpha_*(*ypy_++);
            slice_factor_a(ycy_) = ycc
            ycy_ = ycy_ + 1
            '*ycy_++ = ycc;
            img_out_f(out_) = 0.5! * (img_out_f(out_) + ycc) * mapFactorBCache
            '*out_ = 0.5f * (*out_ + ycc) / (*factor_);
            out_ = out_ + 1
            
            ycc = inv_alpha_ * img_temp(xcy) + alpha_ * slice_factor_b(ypy_)
            xcy = xcy + 1
            ypy_ = ypy_ + 1
            'float ycc = inv_alpha_*(*xcy++) + alpha_*(*ypy_++);
            slice_factor_a(ycy_) = ycc
            ycy_ = ycy_ + 1
            '*ycy_++ = ycc;
            img_out_f(out_) = 0.5! * (img_out_f(out_) + ycc) * mapFactorBCache
            '*out_ = 0.5f * (*out_ + ycc) / (*factor_);
            out_ = out_ + 1
            
            ycc = inv_alpha_ * img_temp(xcy) + alpha_ * slice_factor_b(ypy_)
            xcy = xcy + 1
            ypy_ = ypy_ + 1
            'float ycc = inv_alpha_*(*xcy++) + alpha_*(*ypy_++);
            slice_factor_a(ycy_) = ycc
            ycy_ = ycy_ + 1
            '*ycy_++ = ycc;
            img_out_f(out_) = 0.5! * (img_out_f(out_) + ycc) * mapFactorBCache
            '*out_ = 0.5f * (*out_ + ycc) / (*factor_);
            out_ = out_ + 1
            
            ycc = inv_alpha_ * img_temp(xcy) + alpha_ * slice_factor_b(ypy_)
            xcy = xcy + 1
            ypy_ = ypy_ + 1
            'float ycc = inv_alpha_*(*xcy++) + alpha_*(*ypy_++);
            slice_factor_a(ycy_) = ycc
            ycy_ = ycy_ + 1
            '*ycy_++ = ycc;
            img_out_f(out_) = 0.5! * (img_out_f(out_) + ycc) * mapFactorBCache
            '*out_ = 0.5f * (*out_ + ycc) / (*factor_);
            out_ = out_ + 1
            
            factor_ = factor_ + 1
            
        Next x
        
        'Update slice and line factors
        CopyMemoryStrict VarPtr(slice_factor_b(0)), VarPtr(slice_factor_a(0)), 4 * width_channel
        CopyMemoryStrict VarPtr(line_factor_b(0)), VarPtr(line_factor_a(0)), 4 * imgWidth
        
        'UI updates
        If raiseProgressEvents Then
            If (((imgHeight * 2 + (imgHeight - y)) And progCheck) = 0) Then
                RaiseEvent ProgressUpdate((CDbl(imgHeight * 2 + (imgHeight - y)) * progMax), cancelOp)
                If cancelOp Then GoTo FilterCleanup
            End If
        End If
        
    Next y
    
    'Copy the final result back into the source DIB
    Const FLOAT_TO_LONG_MAX_SAFE As Single = 255!
    
    Dim tmpFloat As Single, tmpLong As Long
    For i = 0 To width_height_channel - 1
        
        'For very small radii, floats may exceed the max value of a long.  (I'm still investigating
        ' what combination of factors can make them this large - probably division of a very small
        ' number somewhere, argh.)  As a failsafe, check for this and correct it before casting.
        tmpFloat = img_out_f(i)
        If (tmpFloat > FLOAT_TO_LONG_MAX_SAFE) Then tmpFloat = FLOAT_TO_LONG_MAX_SAFE
        tmpLong = tmpFloat
        img(i) = tmpLong And &HFF&
        
    Next i
    
FilterCleanup:
    
    'Unwrap the DIB array wrapper
    srcDIB.UnwrapArrayFromDIB img
    
    '"Free" all arrays that unsafely wrap the underlying buffer
    VBHacks.UnwrapArrayFromPtr_Float img_out_f
    VBHacks.UnwrapArrayFromPtr_Float img_temp
    VBHacks.UnwrapArrayFromPtr_Float map_factor_a
    VBHacks.UnwrapArrayFromPtr_Float map_factor_b
    VBHacks.UnwrapArrayFromPtr_Float slice_factor_a
    VBHacks.UnwrapArrayFromPtr_Float slice_factor_b
    VBHacks.UnwrapArrayFromPtr_Float line_factor_a
    VBHacks.UnwrapArrayFromPtr_Float line_factor_b
    
End Function
