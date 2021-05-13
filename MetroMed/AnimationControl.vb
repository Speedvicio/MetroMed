' a_pess&yahoo.com
' omar amin ibrahim
' coding for fun
' OCtober 31, 2008
' dedicated to Bob

Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging

Public Class AnimationControl

    Inherits Control

#Region " Field "

    Private m_timer As New System.Timers.Timer()

    Private m_AnimationSpeed As New TimeSpan(0, 0, 0, 2, 0)
    Private m_AnimationStartTime As DateTime
    Private m_AnimatedBitmap As Bitmap = Nothing
    Private m_AnimatedFadeImage As Bitmap = Nothing
    Private m_AnimationType As AnimationTypes
    Private m_AnimationPercent As Integer = 0

    Private m_backcolor As Color
    Private m_borderColor As Color
    Private m_minSize As Size
    Private m_Divider As Integer = 6

    Private Path As GraphicsPath
    Private Rect As Rectangle
    Private IsAnimating As Boolean = False

    Private m_transparent As Boolean
    Private m_transparentColor As Color
    Private m_opacity As Double

#End Region

#Region " Constructor "

    Public Sub New()

        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        SetStyle(ControlStyles.Opaque, False)
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.UserPaint, True)
        UpdateStyles()

        m_backcolor = Color.Transparent
        m_minSize = New Size(100, 100)
        m_borderColor = Color.LightGray
        m_transparent = False
        m_transparentColor = Color.DodgerBlue
        m_opacity = 1.0R

        AddHandler m_timer.Elapsed, AddressOf TimerTick

    End Sub

#End Region

#Region " Property "

    Protected Overloads Overrides ReadOnly Property DefaultSize() As Size
        Get
            Return New Size(150, 150)
        End Get
    End Property

    Public Overloads Overrides Property MinimumSize() As System.Drawing.Size
        Get
            Return m_minSize
        End Get
        Set(ByVal value As System.Drawing.Size)
            If (value <> (m_minSize)) Then
                m_minSize = value
                Refresh()
            End If
        End Set
    End Property

    <System.ComponentModel.Browsable(False)>
    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
    <System.ComponentModel.DefaultValue(GetType(Color), "Transparent")>
    <System.ComponentModel.Description("Set background color.")>
    <System.ComponentModel.Category("Control Style")>
    Public Overrides Property BackColor() As System.Drawing.Color
        Get
            Return m_backcolor
        End Get
        Set(ByVal value As System.Drawing.Color)
            m_backcolor = value
            Refresh()
        End Set
    End Property

    <System.ComponentModel.Browsable(True)>
    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)>
    <System.ComponentModel.DefaultValue(1.0R)>
    <System.ComponentModel.TypeConverter(GetType(OpacityConverter))>
    <System.ComponentModel.Description("Set the opacity percentage of the control.")>
    <System.ComponentModel.Category("Control Style")>
    Public Overridable Property Opacity() As Double
        Get
            Return m_opacity
        End Get
        Set(ByVal value As Double)
            If value = m_opacity Then
                Return
            End If
            m_opacity = value
            UpdateStyles()
            Refresh()
        End Set
    End Property

    <System.ComponentModel.Browsable(True)>
    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)>
    <System.ComponentModel.DefaultValue(GetType(Boolean), "False")>
    <System.ComponentModel.Description("Enable control trnasparency.")>
    <System.ComponentModel.Category("Control Style")>
    Public Overridable Property Transparent() As Boolean
        Get
            Return m_transparent
        End Get
        Set(ByVal value As Boolean)
            If value = m_transparent Then
                Return
            End If
            m_transparent = value
            Refresh()
        End Set
    End Property

    <System.ComponentModel.Browsable(True)>
    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)>
    <System.ComponentModel.DefaultValue(GetType(Color), "DodgerBlue")>
    <System.ComponentModel.Description("Set the fill color of the control.")>
    <System.ComponentModel.Category("Control Style")>
    Public Overridable Property TransparentColor() As Color
        Get
            Return m_transparentColor
        End Get
        Set(ByVal value As Color)
            m_transparentColor = value
            Refresh()
        End Set
    End Property

    <System.ComponentModel.DefaultValue(GetType(Bitmap), "")>
    <System.ComponentModel.Description("Set animated iamge.")>
    <System.ComponentModel.Category("Control Style")>
    Public Property AnimatedImage() As Bitmap
        Get
            Return m_AnimatedBitmap
        End Get
        Set(ByVal value As Bitmap)
            m_AnimatedBitmap = value
        End Set
    End Property

    <System.ComponentModel.DefaultValue(GetType(Bitmap), "")>
    <System.ComponentModel.Description("Set fade iamge.")>
    <System.ComponentModel.Category("Control Style")>
    Public Property AnimatedFadeImage() As Bitmap
        Get
            Return m_AnimatedFadeImage
        End Get
        Set(ByVal value As Bitmap)
            m_AnimatedFadeImage = value
        End Set
    End Property

    <System.ComponentModel.DefaultValue(GetType(AnimationTypes), "Maximize")>
    <System.ComponentModel.Description("Set animation type.")>
    <System.ComponentModel.Category("Control Style")>
    Public Property AnimationType() As AnimationTypes
        Get
            Return m_AnimationType
        End Get
        Set(ByVal value As AnimationTypes)
            m_AnimationType = value
        End Set
    End Property

    <System.ComponentModel.DefaultValue(GetType(Single), "2")>
    <System.ComponentModel.Description("Set animation speed. the greater value the slowest speed")>
    <System.ComponentModel.Category("Control Style")>
    Public Property AnimationSpeed() As Single
        Get
            Return CSng(m_AnimationSpeed.TotalSeconds)
        End Get
        Set(ByVal value As Single)
            m_AnimationSpeed = New TimeSpan(0, 0, 0, 0, CInt((1000 * value)))
        End Set
    End Property

    <System.ComponentModel.DefaultValue(GetType(Color), "LightGray")>
    <System.ComponentModel.Description("Set border color.")>
    <System.ComponentModel.Category("Control Style")>
    Public Property BorderColor() As Color
        Get
            Return m_borderColor
        End Get
        Set(ByVal value As Color)
            m_borderColor = value
            Invalidate()
        End Set
    End Property

#End Region

#Region " Method "

    Public Sub Animate(ByVal m_Interval As Integer)

        If m_Interval > 100 Then
            m_Interval = 100
        End If

        m_timer.Interval = m_Interval
        m_timer.Enabled = True
        m_AnimationPercent = 0
        m_AnimationStartTime = DateTime.Now
        IsAnimating = True

        Invalidate()

    End Sub

    Private Sub AnimationStop()
        IsAnimating = False
        m_timer.Enabled = False
    End Sub

    Private Sub TimerTick(ByVal source As Object, ByVal e As System.Timers.ElapsedEventArgs)

        Dim ts As TimeSpan = DateTime.Now - m_AnimationStartTime
        m_AnimationPercent = CSng((100.0F / m_AnimationSpeed.TotalSeconds * ts.TotalSeconds))

        If m_AnimationPercent > 100 Then
            m_AnimationPercent = 100
        End If

        Invalidate()

    End Sub

    Private Sub DrawBorder(ByVal g As Graphics, ByVal control As AnimationControl)

        Using p As New Pen(GetDarkColor(Me.BorderColor, 40), -1)
            Rect = New Rectangle(control.ClientRectangle.X, control.ClientRectangle.Y, control.ClientRectangle.Width - 1, control.ClientRectangle.Height - 1)
            g.DrawRectangle(p, Rect)
        End Using

    End Sub

    Private Sub DrawBackground(ByVal g As Graphics, ByVal control As AnimationControl)

        If Transparent Then
            Using sb As New SolidBrush(control.BackColor)
                g.FillRectangle(sb, Rect)

                Using sbt As New SolidBrush(Color.FromArgb(control.Opacity * 255, control.TransparentColor))
                    g.FillRectangle(sbt, Rect)
                End Using
            End Using
        Else

            Using sb As New SolidBrush(control.TransparentColor)
                g.FillRectangle(sb, Rect)
            End Using
        End If

    End Sub

    Protected Sub DrawAnimatedImage(ByVal g As Graphics, ByVal control As AnimationControl)

        If m_AnimatedBitmap IsNot Nothing Then

            Select Case m_AnimationType

                Case AnimationTypes.BottomLeftToTopRight ' Image Slides from bottom left to top right effect

                    Dim mx As New Drawing2D.Matrix(1, 0, 0, 1, (control.Width * m_AnimationPercent / 100) - control.Width, -(control.Height * m_AnimationPercent / 100) + control.Height)
                    g.Transform = mx
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    mx.Dispose()
                    Exit Select

                    ' --------------->
                Case AnimationTypes.BottomRightToTopLeft ' Image Slides from bottom right to top left

                    Dim mx As New Drawing2D.Matrix(1, 0, 0, 1, -(control.Width * m_AnimationPercent / 100) + control.Width, -(control.Height * m_AnimationPercent / 100) + control.Height)
                    g.Transform = mx
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    mx.Dispose()
                    Exit Select

                    ' --------------->
                Case AnimationTypes.ChessBoard ' Image chess board effect

                    Path = New GraphicsPath()
                    Dim cw As Integer = CInt((control.Width * m_AnimationPercent / 100)) / m_Divider
                    Dim ch As Integer = CInt((control.Height * m_AnimationPercent / 100)) / m_Divider
                    Dim row As Integer = 0
                    Dim col As Integer = 0

                    Dim y As Integer = 0
                    While y < control.Height
                        Dim x As Integer = 0
                        While x < control.Width
                            Dim rc As New Rectangle(x, y, cw, ch)

                            If (row And 1) = 1 Then
                                If (col And 1) = 1 Then
                                    rc.Offset(control.Width / (2 * m_Divider), control.Height / (2 * m_Divider))
                                End If

                            End If

                            Path.AddRectangle(rc)

                            If m_AnimationPercent >= 50 AndAlso (row And 1) = 1 AndAlso x = 0 Then
                                If m_AnimationPercent >= 50 AndAlso (col And 1) = 1 AndAlso y = 0 Then

                                    rc.Offset((control.Width / m_Divider), (control.Height / m_Divider))
                                    Path.AddRectangle(rc)

                                End If
                            End If
                            x += control.Width / m_Divider
                        End While
                        row += 1
                        y += control.Height / m_Divider
                    End While
                    col += 1

                    Dim r As New Region(Path)
                    g.SetClip(r, CombineMode.Intersect)
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    r.Dispose()
                    Path.Dispose()
                    Exit Select

                    ' --------------->
                Case AnimationTypes.ChessHorizontal ' Image chess board horizontal effect

                    Path = New GraphicsPath()
                    Dim cw As Integer = CInt((control.Width * m_AnimationPercent / 100)) / m_Divider
                    Dim ch As Integer = control.Height / m_Divider
                    Dim row As Integer = 0
                    Dim y As Integer = 0
                    While y < control.Height
                        Dim x As Integer = 0
                        While x < control.Width
                            Dim rc As New Rectangle(x, y, cw, ch)
                            If (row And 1) = 1 Then
                                rc.Offset(control.Width / (2 * m_Divider), 0)
                            End If
                            Path.AddRectangle(rc)
                            If m_AnimationPercent >= 50 AndAlso (row And 1) = 1 AndAlso x = 0 Then
                                rc.Offset(-(control.Width / m_Divider), 0)
                                Path.AddRectangle(rc)
                            End If
                            x += control.Width / m_Divider
                        End While
                        row += 1
                        y += ch
                    End While
                    Dim r As New Region(Path)
                    g.SetClip(r, CombineMode.Intersect)
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    r.Dispose()
                    Path.Dispose()
                    Exit Select

                    ' --------------->
                Case AnimationTypes.ChessVertical ' Image chess board vertical effect

                    Path = New GraphicsPath()
                    Dim cw As Integer = control.Width / m_Divider
                    Dim ch As Integer = CInt((control.Height * m_AnimationPercent / 100)) / m_Divider
                    Dim col As Integer = 0
                    Dim x As Integer = 0

                    While x < control.Width
                        Dim y As Integer = 0
                        While y < control.Height
                            Dim rc As New Rectangle(x, y, cw, ch)
                            If (col And 1) = 1 Then
                                rc.Offset(0, control.Height / (2 * m_Divider))
                            End If
                            Path.AddRectangle(rc)
                            If m_AnimationPercent >= 50 AndAlso (col And 1) = 1 AndAlso y = 0 Then
                                rc.Offset(0, -(control.Height / m_Divider))
                                Path.AddRectangle(rc)
                            End If
                            y += control.Height / m_Divider
                        End While
                        col += 1
                        x += cw
                    End While
                    Dim r As New Region(Path)
                    g.SetClip(r, CombineMode.Intersect)
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    r.Dispose()
                    Path.Dispose()
                    Exit Select

                    ' --------------->
                Case AnimationTypes.Circular ' Image circular effect

                    Path = New Drawing2D.GraphicsPath()
                    Dim w As Integer = CInt(((control.Width * 1.414F) * m_AnimationPercent / 200))
                    Dim h As Integer = CInt(((control.Height * 1.414F) * m_AnimationPercent / 200))

                    Path.AddEllipse(CInt(control.Width / 2) - w, CInt(control.Height / 2) - h, 2 * w, 2 * h)
                    g.SetClip(Path, Drawing2D.CombineMode.Replace)
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    Path.Dispose()

                    Exit Select

                    ' --------------->
                Case AnimationTypes.Fade ' Image fade effect

                    If True Then

                        Dim attr As New ImageAttributes()
                        Dim mx As New ColorMatrix()
                        mx.Matrix33 = 1.0F / 255 * (255 * m_AnimationPercent / 100)
                        attr.SetColorMatrix(mx)
                        g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel, attr)
                        attr.Dispose()

                    End If

                    Exit Select

                    ' --------------->
                Case AnimationTypes.Fade2Images ' fade two image effect

                    If True Then

                        If m_AnimationPercent < 100 Then
                            If m_AnimatedFadeImage IsNot Nothing Then
                                g.DrawImage(m_AnimatedFadeImage, control.ClientRectangle, 0, 0, m_AnimatedFadeImage.Width, m_AnimatedFadeImage.Height, GraphicsUnit.Pixel)
                            End If
                        End If

                        Dim attr As New ImageAttributes()
                        Dim mx As New ColorMatrix()
                        mx.Matrix33 = 1.0F / 255 * (255 * m_AnimationPercent / 100)
                        attr.SetColorMatrix(mx)
                        g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel, attr)
                        attr.Dispose()

                    End If

                    Exit Select

                    ' --------------->
                Case AnimationTypes.DownToTop ' Image slide from down to top effect

                    Dim mx As New Drawing2D.Matrix(1, 0, 0, 1, 0, -(control.Height * m_AnimationPercent / 100) + control.Height)
                    g.Transform = mx
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    mx.Dispose()
                    Exit Select

                    ' --------------->
                Case AnimationTypes.Elliptical ' Image elliptical effect

                    Path = New Drawing2D.GraphicsPath()
                    Dim w As Integer = CInt(((control.Width * 1.1 * 1.42F) * m_AnimationPercent / 200))
                    Dim h As Integer = CInt(((control.Height * 1.3 * 1.42F) * m_AnimationPercent / 200))

                    Path.AddEllipse(CInt(control.Width / 2) - w, CInt(control.Height / 2) - h, 2 * w, 2 * h)
                    g.SetClip(Path, Drawing2D.CombineMode.Replace)
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    Path.Dispose()

                    Exit Select

                    ' --------------->
                Case AnimationTypes.LeftToRight ' Image slide from left to right effect

                    Dim mx As New Drawing2D.Matrix(1, 0, 0, 1, (control.Width * m_AnimationPercent / 100) - control.Width, 0)
                    g.Transform = mx
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    mx.Dispose()
                    Exit Select

                    ' --------------->
                Case AnimationTypes.Maximize ' Image maximize effect

                    Dim m_scale As Single = m_AnimationPercent / 100
                    Dim cX As Single = control.Width / 2
                    Dim cY As Single = control.Height / 2

                    If m_scale = 0 Then
                        m_scale = 0.0001F
                    End If
                    Dim mx As New Drawing2D.Matrix(m_scale, 0, 0, m_scale, cX, cY)
                    g.Transform = mx
                    g.DrawImage(m_AnimatedBitmap, New Rectangle(-control.Width / 2, -control.Height / 2, control.Width, control.Height), 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)

                    Exit Select

                    ' --------------->
                Case AnimationTypes.Rectangular ' Image rectangular effect

                    Path = New Drawing2D.GraphicsPath()
                    Dim w As Integer = CInt(((control.Width * 1.414F) * m_AnimationPercent / 200))
                    Dim h As Integer = CInt(((control.Height * 1.414F) * m_AnimationPercent / 200))

                    Dim rect As New Rectangle(CInt(control.Width / 2) - w, CInt(control.Height / 2) - h, 2 * w, 2 * h)
                    Path.AddRectangle(rect)

                    g.SetClip(Path, Drawing2D.CombineMode.Replace)
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    Path.Dispose()

                    Exit Select

                    ' --------------->
                Case AnimationTypes.RighTotLeft ' Image slide right to left effect

                    Dim mx As New Drawing2D.Matrix(1, 0, 0, 1, -(control.Width * m_AnimationPercent / 100) + control.Width, 0)
                    g.Transform = mx
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    mx.Dispose()
                    Exit Select

                    ' --------------->
                Case AnimationTypes.Rotate ' Image rotate effect

                    Dim m_rotation As Single = 360 * m_AnimationPercent / 100
                    Dim cX As Single = control.Width / 2
                    Dim cY As Single = control.Height / 2
                    Dim mx As New Drawing2D.Matrix(1, 0, 0, 1, cX, cY)
                    mx.Rotate(m_rotation, Drawing2D.MatrixOrder.Append)
                    g.Transform = mx
                    g.DrawImage(m_AnimatedBitmap, New Rectangle(-control.Width / 2, -control.Height / 2, control.Width, control.Height), 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)

                    Exit Select

                    ' --------------->
                Case AnimationTypes.Spin ' Image spin effect

                    Dim m_rotation As Single = 360 * m_AnimationPercent / 100
                    Dim cX As Single = control.Width / 2
                    Dim cY As Single = control.Height / 2
                    Dim m_scale As Single = m_AnimationPercent / 100
                    If m_scale = 0 Then
                        m_scale = 0.0001F
                    End If

                    Dim mx As New Drawing2D.Matrix(m_scale, 0, 0, m_scale, cX, cY)
                    mx.Rotate(m_rotation, Drawing2D.MatrixOrder.Prepend)
                    g.Transform = mx
                    g.DrawImage(m_AnimatedBitmap, New Rectangle(-control.Width / 2, -control.Height / 2, control.Width, control.Height), 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)

                    Exit Select

                    ' --------------->
                Case AnimationTypes.TopLeftToBottomRight ' Image slide from top left to bottom right effect

                    Dim mx As New Drawing2D.Matrix(1, 0, 0, 1, (control.Width * m_AnimationPercent / 100) - control.Width, (control.Height * m_AnimationPercent / 100) - control.Height)
                    g.Transform = mx
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    mx.Dispose()

                    Exit Select

                    ' --------------->
                Case AnimationTypes.TopRightToBottomLeft ' Image slide from top right to bottom left effect

                    Dim mx As New Drawing2D.Matrix(1, 0, 0, 1, -(control.Width * m_AnimationPercent / 100) + control.Width, (control.Height * m_AnimationPercent / 100) - control.Height)
                    g.Transform = mx
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    mx.Dispose()

                    Exit Select

                    ' --------------->
                Case AnimationTypes.TopToDown ' Image slide top to down effect

                    Dim mx As New Drawing2D.Matrix(1, 0, 0, 1, 0, (control.Height * m_AnimationPercent / 100) - control.Height)
                    g.Transform = mx
                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)
                    mx.Dispose()

                    Exit Select

                    ' --------------->
                Case AnimationTypes.SplitHorizontal ' Image split horizontal effect

                    g.DrawImage(m_AnimatedBitmap, New Rectangle(0,
                                                                0,
                                                                CInt((control.Width * m_AnimationPercent / 200)), control.Height),
                                                                0,
                                                                0,
                                                                CInt(m_AnimatedBitmap.Width / 2),
                                                                m_AnimatedBitmap.Height,
                                                                GraphicsUnit.Pixel)

                    g.DrawImage(m_AnimatedBitmap, New Rectangle(CInt((control.Width - CInt(control.Width * m_AnimationPercent / 200))),
                                                               0,
                                                               CInt((control.ClientRectangle.Width * m_AnimationPercent / 200)),
                                                               control.ClientRectangle.Height),
                                                               CInt(m_AnimatedBitmap.Width / 2),
                                                               0,
                                                               CInt(m_AnimatedBitmap.Width / 2),
                                                               m_AnimatedBitmap.Height,
                                                               GraphicsUnit.Pixel)

                    Exit Select

                    ' --------------->
                Case AnimationTypes.SplitQuarter ' Image split quarter effect

                    g.DrawImage(m_AnimatedBitmap, New Rectangle(0,
                                                                0,
                                                               CInt((control.Width * m_AnimationPercent / 200)),
                                                               CInt((control.Height * m_AnimationPercent / 200))),
                                                               0,
                                                               0,
                                                               CInt(m_AnimatedBitmap.Width / 2),
                                                               CInt(m_AnimatedBitmap.Height / 2),
                                                               GraphicsUnit.Pixel)

                    g.DrawImage(m_AnimatedBitmap, New Rectangle(CInt((control.Width - CInt(control.Width * m_AnimationPercent / 200))),
                                                               0,
                                                               CInt((control.ClientRectangle.Width * m_AnimationPercent / 200)),
                                                               CInt((control.ClientRectangle.Height * m_AnimationPercent / 200))),
                                                               CInt(m_AnimatedBitmap.Width / 2),
                                                               0,
                                                               CInt(m_AnimatedBitmap.Width / 2),
                                                               CInt(m_AnimatedBitmap.Height / 2),
                                                               GraphicsUnit.Pixel)

                    g.DrawImage(m_AnimatedBitmap, New Rectangle(0,
                                                               CInt((control.Height - CInt(control.Height * m_AnimationPercent / 200))),
                                                               CInt((control.ClientRectangle.Width * m_AnimationPercent / 200)),
                                                               CInt((control.ClientRectangle.Height * m_AnimationPercent / 200))),
                                                               0,
                                                               CInt(m_AnimatedBitmap.Height / 2),
                                                               CInt(m_AnimatedBitmap.Width / 2),
                                                               CInt(m_AnimatedBitmap.Height / 2),
                                                               GraphicsUnit.Pixel)

                    g.DrawImage(m_AnimatedBitmap, New Rectangle(CInt((control.Width - CInt(control.Width * m_AnimationPercent / 200))),
                                                               CInt((control.Height - CInt(control.Height * m_AnimationPercent / 200))),
                                                               CInt((control.ClientRectangle.Width * m_AnimationPercent / 200)),
                                                               CInt((control.ClientRectangle.Height * m_AnimationPercent / 200))),
                                                               CInt(m_AnimatedBitmap.Width / 2),
                                                               CInt(m_AnimatedBitmap.Height / 2),
                                                               CInt(m_AnimatedBitmap.Width / 2),
                                                               CInt(m_AnimatedBitmap.Height / 2),
                                                               GraphicsUnit.Pixel)

                    Exit Select

                    ' --------------->
                Case AnimationTypes.SplitBoom  ' Image split shake effect

                    g.DrawImage(m_AnimatedBitmap, New Rectangle(0,
                                                                0,
                                                                CInt((control.Width * m_AnimationPercent / 200)), control.Rect.Height),
                                                                0,
                                                                0,
                                                                CInt(m_AnimatedBitmap.Width / 2),
                                                                m_AnimatedBitmap.Height,
                                                                GraphicsUnit.Pixel)

                    g.DrawImage(m_AnimatedBitmap, New Rectangle(CInt((control.Width - CInt(control.Width * m_AnimationPercent / 200))),
                                                               0,
                                                               CInt((control.ClientRectangle.Width * m_AnimationPercent / 200)),
                                                               control.ClientRectangle.Height),
                                                               CInt(m_AnimatedBitmap.Width / 2),
                                                               0,
                                                               CInt(m_AnimatedBitmap.Width / 2),
                                                               m_AnimatedBitmap.Height,
                                                               GraphicsUnit.Pixel)

                    g.DrawImage(m_AnimatedBitmap, New Rectangle(0,
                                                                0,
                                                                control.Width,
                                                                CInt((control.Height * m_AnimationPercent / 200))),
                                                                0,
                                                                0,
                                                                m_AnimatedBitmap.Width,
                                                                CInt(m_AnimatedBitmap.Height / 2),
                                                                GraphicsUnit.Pixel)

                    g.DrawImage(m_AnimatedBitmap, New Rectangle(0,
                                                                CInt((control.Height - CInt(control.Height * m_AnimationPercent / 200))),
                                                                control.ClientRectangle.Width,
                                                                CInt((control.ClientRectangle.Height * m_AnimationPercent / 200))),
                                                                0,
                                                                CInt(m_AnimatedBitmap.Height / 2),
                                                                m_AnimatedBitmap.Width,
                                                                CInt(m_AnimatedBitmap.Height / 2),
                                                                GraphicsUnit.Pixel)

                    Exit Select

                    ' --------------->
                Case AnimationTypes.SplitVertical ' Image split vertical effect

                    g.DrawImage(m_AnimatedBitmap, New Rectangle(0,
                                                                0,
                                                                control.Width,
                                                                CInt((control.Height * m_AnimationPercent / 200))),
                                                                0,
                                                                0,
                                                                m_AnimatedBitmap.Width,
                                                                CInt(m_AnimatedBitmap.Height / 2),
                                                                GraphicsUnit.Pixel)

                    g.DrawImage(m_AnimatedBitmap, New Rectangle(0,
                                                                CInt((control.Height - CInt(control.Height * m_AnimationPercent / 200))),
                                                                control.ClientRectangle.Width,
                                                                CInt((control.ClientRectangle.Height * m_AnimationPercent / 200))),
                                                                0,
                                                                CInt(m_AnimatedBitmap.Height / 2),
                                                                m_AnimatedBitmap.Width,
                                                                CInt(m_AnimatedBitmap.Height / 2),
                                                                GraphicsUnit.Pixel)
                    Exit Select

                    ' --------------->
                Case AnimationTypes.Panorama ' Image panorama effect

                    For y As Integer = 0 To m_Divider - 1
                        For x As Integer = 0 To m_Divider - 1

                            Dim src As New Rectangle(x * (m_AnimatedBitmap.Width / m_Divider),
                                                     y * (m_AnimatedBitmap.Height / m_Divider),
                                                     m_AnimatedBitmap.Width / m_Divider,
                                                     m_AnimatedBitmap.Height / m_Divider)

                            Dim drc As New Rectangle(x * (control.Width / m_Divider),
                                                     y * (control.Height / m_Divider),
                                                     CInt(((control.Width / m_Divider) * m_AnimationPercent / 100)),
                                                     CInt(((control.Height / m_Divider) * m_AnimationPercent / 100)))

                            drc.Offset((control.Width / (m_Divider * 2)) - drc.Width / 2, (control.Height / (m_Divider * 2)) - drc.Height / 2)

                            g.DrawImage(m_AnimatedBitmap, drc, src, GraphicsUnit.Pixel)

                        Next
                    Next

                    Exit Select

                    ' --------------->
                Case AnimationTypes.PanoramaHorizontal ' Image panorama horizontal effect

                    For y As Integer = 0 To m_Divider - 1

                        Dim src As New Rectangle(0,
                                                 y * (m_AnimatedBitmap.Height / m_Divider),
                                                 m_AnimatedBitmap.Width,
                                                 m_AnimatedBitmap.Height / m_Divider)

                        Dim drc As New Rectangle(0,
                                                 y * (control.Height / m_Divider),
                                                 control.Width,
                                                 CInt(((control.Height / m_Divider) * m_AnimationPercent / 100)))

                        drc.Offset(0, (control.Height / (m_Divider * 2)) - drc.Height / 2)

                        g.DrawImage(m_AnimatedBitmap, drc, src, GraphicsUnit.Pixel)
                    Next

                    Exit Select

                    ' --------------->
                Case AnimationTypes.PanoramaVertical ' Image panorama vetical effect

                    For x As Integer = 0 To m_Divider - 1
                        Dim src As New Rectangle(x * (m_AnimatedBitmap.Width / m_Divider),
                                                 0,
                                                 m_AnimatedBitmap.Width / m_Divider,
                                                 m_AnimatedBitmap.Height)

                        Dim drc As New Rectangle(x * (control.Width / m_Divider),
                                                 0,
                                                 CInt(((control.Width / m_Divider) * m_AnimationPercent / 100)),
                                                 control.Height)

                        drc.Offset((control.Width / (m_Divider * 2)) - drc.Width / 2, 0)
                        g.DrawImage(m_AnimatedBitmap, drc, src, GraphicsUnit.Pixel)
                    Next
                    Exit Select

                    ' --------------->
                Case AnimationTypes.Spiral ' Image spiral effect

                    If m_AnimationPercent < 100 Then

                        Dim percentageAngle As Double = m_Divider * (Math.PI * 2) / 100
                        Dim percentageDistance As Double = Math.Max(control.Width, control.Height) / 100
                        Path = New GraphicsPath(FillMode.Winding)

                        Dim cx As Single = control.Width / 2
                        Dim cy As Single = control.Height / 2

                        Dim pc1 As Double = m_AnimationPercent - 100
                        Dim pc2 As Double = m_AnimationPercent

                        If pc1 < 0 Then
                            pc1 = 0
                        End If

                        Dim a As Double = percentageAngle * pc2
                        Dim last As New PointF(CSng((cx + (pc1 * percentageDistance * Math.Cos(a)))), CSng((cy + (pc1 * percentageDistance * Math.Sin(a)))))
                        a = percentageAngle * pc1

                        While pc1 <= pc2
                            Dim thisPoint As New PointF(CSng((cx + (pc1 * percentageDistance * Math.Cos(a)))), CSng((cy + (pc1 * percentageDistance * Math.Sin(a)))))
                            Path.AddLine(last, thisPoint)
                            last = thisPoint
                            pc1 += 0.1
                            a += percentageAngle / 10
                        End While

                        Path.CloseFigure()
                        g.SetClip(Path, CombineMode.Replace)
                        Path.Dispose()

                    End If

                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)

                    Exit Select

                Case AnimationTypes.SpiralBoom ' Image spiral boom effect

                    If m_AnimationPercent < 100 Then

                        Dim percentageAngle As Double = m_Divider * (Math.PI * 2) / 100
                        Dim percentageDistance As Double = Math.Max(control.Width, control.Height) / 100
                        Path = New GraphicsPath(FillMode.Winding)

                        Dim cx As Single = control.Width / 2
                        Dim cy As Single = control.Height / 2

                        Dim pc1 As Double = m_AnimationPercent - 100
                        Dim pc2 As Double = m_AnimationPercent

                        If pc1 < 0 Then
                            pc1 = 0
                        End If

                        Dim a As Double = percentageAngle * pc2
                        Dim last As New PointF(CSng((cx + (pc1 * percentageDistance * Math.Cos(a)))), CSng((cy + (pc1 * percentageDistance * Math.Sin(a)))))
                        a = percentageAngle * pc1

                        While pc1 <= pc2
                            Dim thisPoint As New PointF(CSng((cx + (pc1 * percentageDistance * Math.Cos(a)))), CSng((cy + (pc1 * percentageDistance * Math.Sin(a)))))
                            Path.AddLine(last, thisPoint)
                            last = thisPoint
                            pc1 += 0.1
                            a += percentageAngle / 10
                        End While

                        Path.CloseFigure()
                        g.SetClip(Path, CombineMode.Exclude)
                        Path.Dispose()

                    End If

                    g.DrawImage(m_AnimatedBitmap, control.ClientRectangle, 0, 0, m_AnimatedBitmap.Width, m_AnimatedBitmap.Height, GraphicsUnit.Pixel)

                    Exit Select

            End Select

        End If

        If m_AnimationPercent = 100 Then
            AnimationStop()
        End If

    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)

        DrawBorder(e.Graphics, Me)
        DrawBackground(e.Graphics, Me)
        DrawAnimatedImage(e.Graphics, Me)
        MyBase.OnPaint(e)

    End Sub

    Protected Overrides Sub OnResize(ByVal e As System.EventArgs)
        MyBase.OnResize(e)
        Invalidate()
    End Sub

#End Region

#Region " Function "

    Private Function GetLightColor(ByVal colorIn As Color, ByVal percent As Single) As Color
        If percent < 0 OrElse percent > 100 Then
            Throw New ArgumentOutOfRangeException("percent must be between 0 and 100")
        End If
        Dim a As Int32 = colorIn.A * Me.Opacity
        Dim r As Int32 = colorIn.R + CInt(((255 - colorIn.R) / 100) * percent)
        Dim g As Int32 = colorIn.G + CInt(((255 - colorIn.G) / 100) * percent)
        Dim b As Int32 = colorIn.B + CInt(((255 - colorIn.B) / 100) * percent)
        Return Color.FromArgb(a, r, g, b)
    End Function

    Private Function GetDarkColor(ByVal colorIn As Color, ByVal percent As Single) As Color
        If percent < 0 OrElse percent > 100 Then
            Throw New ArgumentOutOfRangeException("percent must be between 0 and 100")
        End If
        Dim a As Int32 = colorIn.A * Me.Opacity
        Dim r As Int32 = colorIn.R - CInt((colorIn.R / 100) * percent)
        Dim g As Int32 = colorIn.G - CInt((colorIn.G / 100) * percent)
        Dim b As Int32 = colorIn.B - CInt((colorIn.B / 100) * percent)
        Return Color.FromArgb(a, r, g, b)
    End Function

#End Region

End Class