Imports Wolfram.NETLink

Public Class MathWrapper

#Region " Variables "

	Private myKernelReady As Boolean
	'Mathematica kernel link
	Private ml As IKernelLink
	'Mathematica kernel
	Private kernel As MathKernel

#End Region

#Region " Properties "

	Public ReadOnly Property KernelReady As Boolean
		Get
			Return Me.myKernelReady
		End Get
	End Property

#End Region

#Region " New "

	Public Sub New(kernelPath As String)
		Dim mlArgs As String

		Try
			'start Mathematica kernel
			myKernelReady = False

			mlArgs = "-linkmode launch -linkname '" & kernelPath & "'"
			'Me.ml = MathLinkFactory.CreateKernelLink(mlArgs)
			'Me.ml.WaitAndDiscardAnswer()

			'Dim tmpMl As IKernelLink
			'tmpMl = MathLinkFactory.CreateKernelLink(mlArgs)

			Me.kernel = New Wolfram.NETLink.MathKernel()

			Me.kernel.ResultFormat = MathKernel.ResultFormatType.OutputForm

			Me.kernel.GraphicsHeight = 100
			Me.kernel.GraphicsWidth = 100
			Me.kernel.AutoCloseLink = True
			Me.kernel.CaptureGraphics = True
			Me.kernel.CaptureMessages = True
			Me.kernel.CapturePrint = True
			Me.kernel.GraphicsFormat = "Automatic"
			Me.kernel.GraphicsHeight = 0
			Me.kernel.GraphicsResolution = 0
			Me.kernel.GraphicsWidth = 0
			Me.kernel.HandleEvents = True
			Me.kernel.Input = Nothing
			Me.kernel.LinkArguments = Nothing
			Me.kernel.PageWidth = 60
			Me.kernel.UseFrontEnd = True

			Me.kernel.Connect()

			myKernelReady = True
		Catch ex As Exception
			Windows.Forms.MessageBox.Show("Error starting Wolfram Mathematica kernel" & ControlChars.CrLf & "error message " & ex.Message, "CRITICAL ERROR", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
		End Try
	End Sub

#End Region

#Region " Public Methods "

	Public Sub KernelStop()
		'Me.ml.Close()
	End Sub

	''' <summary>
	''' evaluate an expression and returns result as an image
	''' </summary>
	''' <param name="expression">expression to evaluate</param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function WaitAndEvaluateAsImage(ByVal expression As String, ByVal pbo As Windows.Forms.PictureBox) As Drawing.Image
		Return Me.WaitAndEvaluateAsImage(expression, pbo.Width, pbo.Height)
	End Function

	''' <summary>
	''' evaluate an expression and returns result as an image
	''' </summary>
	''' <param name="expression">expression to evaluate</param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function WaitAndEvaluateAsImage(ByVal expression As String, ByVal width As Integer, ByVal height As Integer) As Drawing.Image
		Try
			Me.kernel.CaptureGraphics = True
			Me.kernel.Compute(expression)
			Return Me.kernel.Graphics(0)
		Catch ex As Exception
			Return Nothing
		End Try
	End Function

    ''' <summary>
    ''' evaluate an expression and returns result as an array of double
    ''' </summary>
    ''' <param name="expression">expression to evaluate</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function WaitAndEvaluateAsDoubleArray(ByVal expression As String) As Double()
        ml.Evaluate(expression)
        ml.WaitForAnswer()
        If ml.GetExpressionType = ExpressionType.Symbol Then
            Throw New Exception()
        Else
            Try
                Return ml.GetDoubleArray
            Catch ex As Exception
                Throw
            End Try
        End If
        'aggiungere controllo: se valore di ritorno non è double ritornare errore
    End Function
    ''' <summary>
    ''' evaluate an expression and returns result as a double
    ''' </summary>
    ''' <param name="expression">expression to evaluate</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function WaitAndEvaluateAsDouble(ByVal expression As String) As Double
		'Dim tmpKernel As MathKernel

		'ml.Evaluate(expression)
		'tmpKernel = New Wolfram.NETLink.MathKernel(ml)

		Me.kernel.Compute(expression)

		'ml.WaitForAnswer()
		'If ml.GetExpressionType = ExpressionType.Symbol OrElse ml.GetExpressionType = ExpressionType.Function Then
		'	Throw New Exception()
		'Else
		Try
			Return CType(Me.kernel.Result, Double)
		Catch ex As Exception
			Throw
		End Try
		'End If
		'aggiungere controllo: se valore di ritorno non è double ritornare errore
	End Function

	''' <summary>
	''' evaluate an expression and returns result as a string
	''' </summary>
	''' <param name="expression">expression to evaluate</param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function WaitAndEvaluateAsString(ByVal expression As String) As String
		ml.Evaluate(expression)
		ml.WaitForAnswer()

		Return ml.GetString
	End Function

	''' <summary>
	''' evaluate an expression and returns result as an expression
	''' </summary>
	''' <param name="expression">expression to evaluate</param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function WaitAndDiscardAnswer(ByVal expression As String) As Wolfram.NETLink.Expr
		'Me.ml.Evaluate(expression)
		'Me.ml.WaitAndDiscardAnswer()

		Me.kernel.Compute(expression)

		Return Nothing
	End Function

#End Region

End Class
