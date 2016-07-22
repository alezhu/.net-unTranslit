Imports System.Text
Imports System.IO
Module Module1

    Dim objWord As Word.ApplicationClass
    Dim objDoc As Word.Document
    Const EngTable As String = "abcdefghijklmnopqrstuvwxyz"
    Const RusTable As String = "àáöäåôãõèæêëìíîï ðñòóââ éç"
    Dim EngTableEx() As String = {"tsja", "tsja", "zh", "x", "ch", "sh", "sch", "'", "ji", "yi", "ye", "je", "yu", "ju", "ya", "ja", "ts"}
    Dim RusTableEx() As String = {"òñÿ", "òñÿ", "æ", "êñ", "÷", "ø", "ù", "ü", "û", "û", "ý", "ý", "þ", "þ", "ÿ", "ÿ", "ö"}

    Function GetSpellCheck(ByVal AValue As String) As String
        If objWord Is Nothing Then
			'objWord = New Word.ApplicationClass()
			objWord = New Word.Application
        End If
        Dim objSpellingSuggestions As Word.SpellingSuggestions
        Dim objSpellingSuggestion As Word.SpellingSuggestion
        If objDoc Is Nothing Then
            objDoc = objWord.Documents.Add()
        End If
        Dim t As String = AValue.ToLower
        objSpellingSuggestions = objWord.GetSpellingSuggestions(t)
        Dim s As String, b As Boolean = False
        'Dim Bites() As Byte = Encoding.Unicode.GetBytes(AValue.ToLower)
        'Dim t As String = Encoding.ASCII.GetString(Encoding.Convert(Encoding.Unicode, Encoding.ASCII, Bites))


        If objSpellingSuggestions.Count > 0 Then
            For Each objSpellingSuggestion In objSpellingSuggestions
                If Not b Then
                    s = objSpellingSuggestion.Name
                    If (s.Length = t.Length) Then
                        If String.Compare(s, t) = 0 Then
                        Else
                            GetSpellCheck = s
                            b = True
                        End If
                    End If
                End If
            Next
            If Not b Then
                GetSpellCheck = AValue
            End If
        Else
            GetSpellCheck = AValue
        End If

    End Function

    Function Translate(ByVal AValue As String) As String
        Dim i As Int32, pos As Int32, s As String
        Dim sb As New StringBuilder(AValue)
        Dim res As New StringBuilder(AValue.Length)

        i = 0
        Do While (i < sb.Length)
            Dim j As Int32, Was As Boolean = True
            Do
                Was = False
                j = 0
                For Each s In EngTableEx
                    If String.Compare(s, 0, AValue, i, s.Length, True) = 0 Then
                        res.Append(RusTableEx(j))
                        i = i + s.Length
                        Was = True
                    End If
                    j = j + 1
                Next
            Loop While Was
            If i < sb.Length Then
                Dim UCase As Boolean
                UCase = False
                pos = EngTable.IndexOf(AValue.Chars(i))
                If pos >= 0 Then
                    res.Append(RusTable.Chars(pos))
                Else
                    pos = EngTable.IndexOf(Char.ToLower(AValue.Chars(i)))
                    If pos >= 0 Then
                        res.Append(Char.ToUpper(RusTable.Chars(pos)))
                    Else
                        res.Append(AValue.Chars(i))
                    End If
                End If
                i = i + 1
            End If
        Loop
        Translate = res.ToString
    End Function

    Function DoTranslate(ByVal AValue As String) As String
        Dim s, sp As String
        Dim sWords() As String
        sWords = AValue.Split(" ")
        Dim sb As StringBuilder = New StringBuilder()
        For Each AValue In sWords
            s = Translate(AValue)
            If Not (s = AValue) Then
                sp = GetSpellCheck(s)
            Else
                sp = s
            End If
            sb.AppendFormat(" {0}", sp)
        Next
        sp = sb.ToString.Trim
        Console.WriteLine("{0} => {1} => {2}", AValue, s, sp)
        DoTranslate = sp
    End Function

    Overloads Sub TranslateFile(ByVal AFilename As String)
        Dim fi As FileInfo
        fi = New FileInfo(AFilename)
        TranslateFile(fi)
    End Sub

    Overloads Sub TranslateFile(ByVal AFile As FileInfo)
        Dim s As String
        s = DoTranslate(Path.ChangeExtension(AFile.Name, Nothing))
        File.Move(AFile.FullName, AFile.Directory.FullName & Path.DirectorySeparatorChar & s & AFile.Extension)
    End Sub

    Sub TranslateFolder(ByVal AFolderName As String)
        Dim di As DirectoryInfo = New DirectoryInfo(AFolderName)
        Dim fn As String
        Dim fi As FileInfo
        Dim s As String
        fn = di.FullName & Path.PathSeparator
        Dim aFiles() As FileInfo = di.GetFiles()
        For Each fi In aFiles
            TranslateFile(fi)
        Next

    End Sub

    Sub Main()
        objWord = Nothing
        objDoc = Nothing
        Dim aWords() As String = System.Environment.GetCommandLineArgs()
        Dim sWord, s As String, first As Boolean = True
        Dim sWords() As String
        Dim fi As FileInfo

		Try
			For Each sWord In aWords
				If first Then
					first = False
				Else
					If File.Exists(sWord) Then
						TranslateFile(sWord)
					ElseIf Directory.Exists(sWord) Then
						TranslateFolder(sWord)
					ElseIf Left(sWord, 3) = "/L:" And File.Exists(Mid(sWord, 4)) Then
						sWord = Mid(sWord, 4)
						Dim sr As StreamReader = New StreamReader(sWord, System.Text.Encoding.Default)
						Do
							sWord = sr.ReadLine
							If Not sWord Is Nothing Then
								If File.Exists(sWord) Then
									TranslateFile(sWord)
								ElseIf Directory.Exists(sWord) Then
									TranslateFolder(sWord)
								End If
							End If
						Loop Until sWord Is Nothing
						sr.Close()
					End If
				End If
			Next

		Finally
			If Not objDoc Is Nothing Then
				Call CType(objDoc, Word._Document).Close()
			End If
			If Not objWord Is Nothing Then
				objWord.Quit()
			End If
		End Try
	End Sub

End Module
