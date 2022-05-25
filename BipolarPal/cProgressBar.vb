Imports System.ComponentModel
Imports System.Drawing.Drawing2D

<DesignerCategory("Code")>
Public Class ProgressBarCustomColor
    Inherits ProgressBar

    Private m_ParentBackColor As Color = SystemColors.Window
    Private m_ProgressBarStyle As BarStyle = BarStyle.Standard
    Private m_FlatBorderColor As Color = Color.White

    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.UserPaint Or
                 ControlStyles.OptimizedDoubleBuffer, True)
    End Sub

    Public Enum BarStyle
        Standard
        Flat
    End Enum

    Public Property ProgressBarStyle As BarStyle
        Get
            Return m_ProgressBarStyle
        End Get
        Set
            m_ProgressBarStyle = Value
            If DesignMode Then Me.Parent?.Refresh()
        End Set
    End Property

    Public Property FlatBorderColor As Color
        Get
            Return m_FlatBorderColor
        End Get
        Set
            m_FlatBorderColor = Value
            If DesignMode Then Me.Parent?.Refresh()
        End Set
    End Property

    Protected Overrides Sub OnParentChanged(e As EventArgs)
        MyBase.OnParentChanged(e)
        m_ParentBackColor = If(Me.Parent IsNot Nothing, Me.Parent.BackColor, SystemColors.Window)
    End Sub

    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        MyBase.OnPaintBackground(e)
        Dim rect = Me.ClientRectangle

        If m_ProgressBarStyle = BarStyle.Standard Then
            ProgressBarRenderer.DrawHorizontalBar(e.Graphics, rect)

            Dim baseColor = Color.FromArgb(240, Me.BackColor)
            Dim highColor = Color.FromArgb(160, baseColor)
            Using baseBrush = New SolidBrush(baseColor),
                  highBrush = New SolidBrush(highColor)
                e.Graphics.FillRectangle(highBrush, 1, 2, rect.Width, 9)
                e.Graphics.FillRectangle(baseBrush, 1, 7, rect.Width, rect.Height - 1)
            End Using
        Else
            Using pen = New Pen(m_FlatBorderColor, 2),
                  baseBrush = New SolidBrush(m_ParentBackColor)
                e.Graphics.FillRectangle(baseBrush, 0, 0, rect.Width - 1, rect.Height - 1)
                e.Graphics.DrawRectangle(pen, 1, 1, Me.ClientSize.Width - 2, Me.ClientSize.Height - 2)
            End Using
        End If
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim rect = New RectangleF(PointF.Empty, Me.ClientSize)
        rect.Width = CType(rect.Width * (CType(Me.Value, Single) / Me.Maximum), Integer)
        rect.Size = New SizeF(rect.Width - 2, rect.Height)

        Dim hsl As Single = Me.ForeColor.GetBrightness()
        If m_ProgressBarStyle = BarStyle.Standard Then DrawStandardBar(e.Graphics, rect)
        If m_ProgressBarStyle = BarStyle.Flat Then DrawFlatBar(e.Graphics, rect)
    End Sub

    Private Sub DrawStandardBar(g As Graphics, rect As RectangleF)
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.CompositingQuality = CompositingQuality.HighQuality

        Dim baseColor = Color.FromArgb(240, Me.ForeColor)
        Dim highColor = Color.FromArgb(160, baseColor)

        Using baseBrush = New SolidBrush(baseColor),
              highBrush = New SolidBrush(highColor)


            g.FillRectangle(highBrush, 1, 2, rect.Width, 9)
            g.FillRectangle(baseBrush, 1, 7, rect.Width, rect.Height - 1)


        End Using
    End Sub

    Shared Sub DrawFlatBar(g As Graphics, rect As RectangleF)
        Using baseBrush = New SolidBrush(SystemColors.Highlight)
            g.FillRectangle(baseBrush, 2, 2, rect.Width - 2, rect.Height - 4)
        End Using
    End Sub
End Class
