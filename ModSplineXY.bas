Attribute VB_Name = "ModSplineXY"
Option Explicit

'SplineXY          ・・・元場所：FukamiAddins3.ModApproximate
'Spline            ・・・元場所：FukamiAddins3.ModApproximate
'SplineKeisu       ・・・元場所：FukamiAddins3.ModApproximate
'F_Minverse        ・・・元場所：FukamiAddins3.ModMatrix     
'正方行列かチェック・・・元場所：FukamiAddins3.ModMatrix     
'F_MDeterm         ・・・元場所：FukamiAddins3.ModMatrix     
'F_Mgyoirekae      ・・・元場所：FukamiAddins3.ModMatrix     
'F_Mgyohakidasi    ・・・元場所：FukamiAddins3.ModMatrix     
'F_Mjyokyo         ・・・元場所：FukamiAddins3.ModMatrix     
'F_MMult           ・・・元場所：FukamiAddins3.ModMatrix     



Public Function SplineXY(ByVal ArrayXY2D, InputX As Double)
    'スプライン補間計算を行う
    '＜出力値の説明＞
    '入力値Xに対する補間値Y
    
    '＜入力値の説明＞
    'ArrayXY2D：補間の対象となるX,Yの値が格納された配列
    'ArrayXY2Dの1列目がX,2列目がYとなるようにする。
    'X:補間位置のXの値
    
    '入力値のチェック及び修正'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
    Dim RangeNaraTrue As Boolean: RangeNaraTrue = False
    If IsObject(ArrayXY2D) Then
        ArrayXY2D = ArrayXY2D.Value
        RangeNaraTrue = True
    End If
    
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(ArrayXY2D, 1) <> 1 Or LBound(ArrayXY2D, 2) <> 1 Then
        ArrayXY2D = Application.Transpose(Application.Transpose(ArrayXY2D))
    End If
    
    Dim ArrayX1D
    Dim ArrayY1D
    Dim I As Integer
    Dim N As Integer
    N = UBound(ArrayXY2D, 1)
    ReDim ArrayX1D(1 To N)
    ReDim ArrayY1D(1 To N)
    
    For I = 1 To N
        ArrayX1D(I) = ArrayXY2D(I, 1)
        ArrayY1D(I) = ArrayXY2D(I, 2)
    Next I
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim OutputY As Double
    OutputY = Spline(ArrayX1D, ArrayY1D, InputX)
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    If RangeNaraTrue Then
        'ワークシート関数の場合
        SplineXY = Application.Transpose(OutputY)
    Else
        'VBA上での処理の場合
        SplineXY = OutputY
    End If
    
End Function

Private Function Spline(ByVal ArrayX1D, ByVal ArrayY1D, InputX As Double)
        
    '20171124修正
    '20180309改良
    
    'スプライン補間計算を行う
    
    '<出力値の説明>
    '入力値Xに対する補間値Y
    
    '<入力値の説明>
    'ArrayX1D：補間の対象とする配列X
    'ArrayY1D：補間の対象とする配列Y
    'InputX  ：補間位置のXの値
    
    '入力値のチェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    
    '配列の次元チェック
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(ArrayY1D, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が2なら次元1にする。例)配列(1 to N,1 to 1)→配列(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I As Integer
    Dim N As Integer
    Dim K As Integer
    Dim A
    Dim B
    Dim C
    Dim D
    Dim OutputY As Double '出力値Y
    Dim SotoNaraTrue As Boolean
    SotoNaraTrue = False
    N = UBound(ArrayX1D, 1)
       
    'スプライン計算用の各係数を計算する。参照渡しでA,B,C,Dに格納
    Dim Dummy
    Dummy = SplineKeisu(ArrayX1D, ArrayY1D)
    A = Dummy(1)
    B = Dummy(2)
    C = Dummy(3)
    D = Dummy(4)
        
    For I = 1 To N - 1
        If ArrayX1D(I) < ArrayX1D(I + 1) Then 'Xが単調増加の場合
            If I = 1 And ArrayX1D(1) > InputX Then '範囲に入らないとき(開始点より前)
                OutputY = ArrayY1D(1)
                SotoNaraTrue = True
                Exit For
            
            ElseIf I = N - 1 And ArrayX1D(I + 1) <= InputX Then '範囲に入らないとき(終了点より後)
                OutputY = ArrayY1D(N)
                SotoNaraTrue = True
                Exit For
                
            ElseIf ArrayX1D(I) <= InputX And ArrayX1D(I + 1) > InputX Then '範囲内
                K = I: Exit For
            
            End If
        Else 'Xが単調減少の場合
        
            If I = 1 And ArrayX1D(1) < InputX Then '範囲に入らないとき(開始点より前)
                OutputY = ArrayY1D(1)
                SotoNaraTrue = True
                Exit For
            
            ElseIf I = N - 1 And ArrayX1D(I + 1) >= InputX Then '範囲に入らないとき(終了点より後)
                OutputY = ArrayY1D(N)
                SotoNaraTrue = True
                Exit For
                
            ElseIf ArrayX1D(I + 1) < InputX And ArrayX1D(I) >= InputX Then '範囲内
                K = I: Exit For
            
            End If
        
        End If
    Next I
        
    If SotoNaraTrue = False Then
        OutputY = A(K) + B(K) * (InputX - ArrayX1D(K)) + C(K) * (InputX - ArrayX1D(K)) ^ 2 + D(K) * (InputX - ArrayX1D(K)) ^ 3
    End If
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Spline = OutputY

End Function

Private Function SplineKeisu(ByVal ArrayX1D, ByVal ArrayY1D)

    '参考：http://www5d.biglobe.ne.jp/stssk/maze/spline.html
    Dim I As Integer
    Dim N As Integer
    Dim A
    Dim B
    Dim C
    Dim D
    N = UBound(ArrayX1D, 1)
    ReDim A(1 To N)
    ReDim B(1 To N)
    ReDim D(1 To N)
    
    Dim h()         As Double
    Dim ArrayL2D()  As Double '左辺の配列 要素数(1 to N,1 to N)
    Dim ArrayR1D()  As Double '右辺の配列 要素数(1 to N,1 to 1)
    Dim ArrayLm2D() As Double '左辺の配列の逆行列 要素数(1 to N,1 to N)
    
    ReDim h(1 To N - 1)
    ReDim ArrayL2D(1 To N, 1 To N)
    ReDim ArrayR1D(1 To N, 1 To 1)
    
    'hi = xi+1 - x
    For I = 1 To N - 1
        h(I) = ArrayX1D(I + 1) - ArrayX1D(I)
    Next I
    
    'di = yi
    For I = 1 To N
        A(I) = ArrayY1D(I)
    Next I
    
    '右辺の配列の計算
    For I = 1 To N
        If I = 1 Or I = N Then
            ArrayR1D(I, 1) = 0
        Else
            ArrayR1D(I, 1) = 3 * (ArrayY1D(I + 1) - ArrayY1D(I)) / h(I) - 3 * (ArrayY1D(I) - ArrayY1D(I - 1)) / h(I - 1)
        End If
    Next I
    
    '左辺の配列の計算
    For I = 1 To N
        If I = 1 Then
            ArrayL2D(I, 1) = 1
        ElseIf I = N Then
            ArrayL2D(N, N) = 1
        Else
            ArrayL2D(I - 1, I) = h(I - 1)
            ArrayL2D(I, I) = 2 * (h(I) + h(I - 1))
            ArrayL2D(I + 1, I) = h(I)
        End If
    Next I
    
    '左辺の配列の逆行列
    ArrayLm2D = F_Minverse(ArrayL2D)
    
    'Cの配列を求める
    C = F_MMult(ArrayLm2D, ArrayR1D)
    C = Application.Transpose(C)
    
    'Bの配列を求める
    For I = 1 To N - 1
        B(I) = (A(I + 1) - A(I)) / h(I) - h(I) * (C(I + 1) + 2 * C(I)) / 3
    Next I
    
    'Dの配列を求める
    For I = 1 To N - 1
        D(I) = (C(I + 1) - C(I)) / (3 * h(I))
    Next I
    
    '出力
    Dim Output(1 To 4)
    Output(1) = A
    Output(2) = B
    Output(3) = C
    Output(4) = D
    
    SplineKeisu = Output

End Function

Private Function F_Minverse(ByVal Matrix)
    '20210603改良
    'F_Minverse(input_M)
    'F_Minverse(配列)
    '余因子行列を用いて逆行列を計算
    
    '入力値チェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(Matrix, 1) <> 1 Or LBound(Matrix, 2) <> 1 Then
        Matrix = Application.Transpose(Application.Transpose(Matrix))
    End If
    
    '入力値のチェック
    Call 正方行列かチェック(Matrix)
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I        As Integer
    Dim J        As Integer
    Dim N        As Integer
    Dim Output() As Double
    N = UBound(Matrix, 1)
    ReDim Output(1 To N, 1 To N)
    
    Dim detM As Double '行列式の値を格納
    detM = F_MDeterm(Matrix) '行列式を求める
    
    Dim Mjyokyo '指定の列・行を除去した配列を格納
    
    For I = 1 To N '各列
        For J = 1 To N '各行
            
            'I列,J行を除去する
            Mjyokyo = F_Mjyokyo(Matrix, J, I)
            
            'I列,J行の余因子を求めて出力する逆行列に格納
            Output(I, J) = F_MDeterm(Mjyokyo) * (-1) ^ (I + J) / detM
    
        Next J
    Next I
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_Minverse = Output
    
End Function

Private Sub 正方行列かチェック(Matrix)
    '20210603追加
    
    If UBound(Matrix, 1) <> UBound(Matrix, 2) Then
        MsgBox ("正方行列を入力してください" & vbLf & _
                "入力された配列の要素数は" & "「" & _
                UBound(Matrix, 1) & "×" & UBound(Matrix, 2) & "」" & "です")
        Stop
        End
    End If

End Sub

Private Function F_MDeterm(Matrix)
    '20210603改良
    'F_MDeterm(Matrix)
    'F_MDeterm(配列)
    '行列式を計算
    
    '入力値チェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(Matrix, 1) <> 1 Or LBound(Matrix, 2) <> 1 Then
        Matrix = Application.Transpose(Application.Transpose(Matrix))
    End If
    
    '入力値のチェック
    Call 正方行列かチェック(Matrix)
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim N As Integer
    N = UBound(Matrix, 1)
    
    Dim Matrix2 '掃き出しを行う行列
    Matrix2 = Matrix
    
    For I = 1 To N '各列
        For J = I To N '掃き出し元の行の探索
            If Matrix2(J, I) <> 0 Then
                K = J '掃き出し元の行
                Exit For
            End If
            
            If J = N And Matrix2(J, I) = 0 Then '掃き出し元の値が全て0なら行列式の値は0
                F_MDeterm = 0
                Exit Function
            End If
            
        Next J
        
        If K <> I Then '(I列,I行)以外で掃き出しとなる場合は行を入れ替え
            Matrix2 = F_Mgyoirekae(Matrix2, I, K)
        End If
        
        '掃き出し
        Matrix2 = F_Mgyohakidasi(Matrix2, I, I)
              
    Next I
    
    
    '行列式の計算
    Dim Output As Double
    Output = 1
    
    For I = 1 To N '各(I列,I行)を掛け合わせていく
        Output = Output * Matrix2(I, I)
    Next I
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_MDeterm = Output
    
End Function

Private Function F_Mgyoirekae(Matrix, Row1 As Integer, Row2 As Integer)
    '20210603改良
    'F_Mgyoirekae(Matrix, Row1, Row2)
    'F_Mgyoirekae(配列,指定行番号①,指定行番号②)
    '行列Matrixの①行と②行を入れ替える
    
    Dim I     As Integer
    Dim J     As Integer
    Dim K     As Integer
    Dim M     As Integer
    Dim N     As Integer
    Dim Output
    
    Output = Matrix
    M = UBound(Matrix, 2) '列数取得
    
    For I = 1 To M
        Output(Row2, I) = Matrix(Row1, I)
        Output(Row1, I) = Matrix(Row2, I)
    Next I
    
    F_Mgyoirekae = Output
End Function

Private Function F_Mgyohakidasi(Matrix, Row As Integer, Col As Integer)
    '20210603改良
    'F_Mgyohakidasi(Matrix, Row, Col)
    'F_Mgyohakidasi(配列,指定行,指定列)
    '行列MatrixのRow行､Col列の値で各行を掃き出す
    
    Dim I     As Integer
    Dim J     As Integer
    Dim N     As Integer
    Dim Output
    
    Output = Matrix
    N = UBound(Output, 1) '行数取得
    
    Dim Hakidasi '掃き出し元の行
    Dim X As Double '掃き出し元の値
    Dim Y As Double
    ReDim Hakidasi(1 To N)
    X = Matrix(Row, Col)
    
    For I = 1 To N '掃き出し元の1行を作成
        Hakidasi(I) = Matrix(Row, I)
    Next I
    
    For I = 1 To N '各行
        If I = Row Then
            '掃き出し元の行の場合はそのまま
            For J = 1 To N
                Output(I, J) = Matrix(I, J)
            Next J
        
        Else
            '掃き出し元の行以外の場合は掃き出し
            Y = Matrix(I, Col) '掃き出し基準の列の値
            For J = 1 To N
                Output(I, J) = Matrix(I, J) - Hakidasi(J) * Y / X
            Next J
        End If
    
    Next I
    
    F_Mgyohakidasi = Output
    
End Function

Private Function F_Mjyokyo(Matrix, Row As Integer, Col As Integer)
    '20210603改良
    'F_Mjyokyo(Matrix, Row, Col)
    'F_Mjyokyo(配列,指定行,指定列)
    '行列MatrixのRow行、Col列を除去した行列を返す
    
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim M As Integer
    Dim N As Integer '数え上げ用(Integer型)
    Dim Output '指定した行・列を除去後の配列
    
    N = UBound(Matrix, 1) '行数取得
    M = UBound(Matrix, 2) '列数取得
    ReDim Output(1 To N - 1, 1 To M - 1)
    
    Dim I2 As Integer
    Dim J2 As Integer
    
    I2 = 0 '行方向数え上げ初期化
    For I = 1 To N
        If I = Row Then
            'なにもしない
        Else
            I2 = I2 + 1 '行方向数え上げ
            
            J2 = 0 '列方向数え上げ初期化
            For J = 1 To M
                If J = Col Then
                    'なにもしない
                Else
                    J2 = J2 + 1 '列方向数え上げ
                    Output(I2, J2) = Matrix(I, J)
                End If
            Next J
            
        End If
    Next I
    
    F_Mjyokyo = Output

End Function

Private Function F_MMult(ByVal Matrix1, ByVal Matrix2)
    'F_MMult(Matrix1, Matrix2)
    'F_MMult(配列①,配列②)
    '行列の積を計算
    '20180213改良
    '20210603改良
    
    '入力値のチェックと修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '配列の次元チェック
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(Matrix1, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(Matrix2, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が1なら次元2にする。例)配列(1 to N)→配列(1 to N,1 to 1)
    If IsEmpty(JigenCheck1) Then
        Matrix1 = Application.Transpose(Matrix1)
    End If
    If IsEmpty(JigenCheck2) Then
        Matrix2 = Application.Transpose(Matrix2)
    End If
    
    '行列の開始要素を1に変更（計算しやすいから）
    If UBound(Matrix1, 1) = 0 Or UBound(Matrix1, 2) = 0 Then
        Matrix1 = Application.Transpose(Application.Transpose(Matrix1))
    End If
    If UBound(Matrix2, 1) = 0 Or UBound(Matrix2, 2) = 0 Then
        Matrix2 = Application.Transpose(Application.Transpose(Matrix2))
    End If
    
    '入力値のチェック
    If UBound(Matrix1, 2) <> UBound(Matrix2, 1) Then
        MsgBox ("配列1の列数と配列2の行数が一致しません。" & vbLf & _
               "(出力) = (配列1)(配列2)")
        Stop
        End
    End If
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I        As Integer
    Dim J        As Integer
    Dim K        As Integer
    Dim M        As Integer
    Dim N        As Integer
    Dim M2       As Integer
    Dim Output() As Double '出力する配列
    N = UBound(Matrix1, 1) '配列1の行数
    M = UBound(Matrix1, 2) '配列1の列数
    M2 = UBound(Matrix2, 2) '配列2の列数
    
    ReDim Output(1 To N, 1 To M2)
    
    For I = 1 To N '各行
        For J = 1 To M2 '各列
            For K = 1 To M '(配列1のI行)と(配列2のJ列)を掛け合わせる
                Output(I, J) = Output(I, J) + Matrix1(I, K) * Matrix2(K, J)
            Next K
        Next J
    Next I
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_MMult = Output
    
End Function


