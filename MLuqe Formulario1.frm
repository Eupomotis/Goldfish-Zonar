VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   13920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   24150
   LinkTopic       =   "Form1"
   ScaleHeight     =   13920
   ScaleWidth      =   24150
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   4215
      Left            =   16320
      ScaleHeight     =   4155
      ScaleWidth      =   7155
      TabIndex        =   3
      Top             =   9120
      Width           =   7215
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   4215
      Left            =   16320
      ScaleHeight     =   4155
      ScaleWidth      =   7155
      TabIndex        =   2
      Top             =   4680
      Width           =   7215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "START"
      Height          =   1335
      Left            =   10080
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4215
      Left            =   16320
      ScaleHeight     =   4155
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      Height          =   195
      Left            =   10080
      TabIndex        =   9
      Top             =   3000
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   195
      Left            =   10080
      TabIndex        =   8
      Top             =   3600
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   195
      Left            =   10080
      TabIndex        =   7
      Top             =   3300
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   10080
      TabIndex        =   6
      Top             =   2340
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   10080
      TabIndex        =   5
      Top             =   1740
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   10080
      TabIndex        =   4
      Top             =   2040
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()


Picture1.Cls
Picture2.Cls
Picture3.Cls


'INICIO ALGORITMO***********************************************************************
'Dimensionando matrices
ReDim Capas(30, 40, 3)  'Estructura: Fila,Columna,Profundidad. 1=Capa1; 2=Capa2; 3=Capa3
ReDim Transecto(30, 40, 3)
ReDim Simulado(30, 40, 3)

'Cargando Capas. Cada archivo CSV (1 por cada capa: Capa1.csv; Capa2.csv; Capa3.csv) define la superficie que ocupada por la capa
'Superficie ocupada=1
'Superficie no ocupada=0
'Indicar la ruta como:
'Ruta1$ = "C:\Datos\.....\Algoritmo DEF" + "\Capa1.csv"
'Ruta2$ = "C:\Datos\.....\Algoritmo DEF" + "\Capa2.csv"
'Ruta3$ = "C:\Datos\.....\Algoritmo DEF" + "\Capa3.csv"

'Almacenando los datos de las capas
Open Ruta1$ For Input As #1
Open Ruta2$ For Input As #2
Open Ruta3$ For Input As #3
    For FILA = 1 To 26
        For COLUMNA = 1 To 35
            Input #1, CP1
            Input #2, CP2
            Input #3, CP3
            
            Capas(FILA, COLUMNA, 1) = CP1
            Capas(FILA, COLUMNA, 2) = CP2
            Capas(FILA, COLUMNA, 3) = CP3
            
            If Capas(FILA, COLUMNA, 1) = 1 Then Picture1.Line (COLUMNA - 1, FILA - 1)-(COLUMNA, FILA), RGB(0, 0, 0), B
            If Capas(FILA, COLUMNA, 2) = 1 Then Picture2.Line (COLUMNA - 1, FILA - 1)-(COLUMNA, FILA), RGB(255, 0, 0), B
            If Capas(FILA, COLUMNA, 3) = 1 Then Picture3.Line (COLUMNA - 1, FILA - 1)-(COLUMNA, FILA), RGB(0, 0, 255), B
            
        Next COLUMNA
    Next FILA
Close #1
Close #2
Close #3


'Cargando transecto
'Llenando la matriz Transecto
For FILA = 1 To 30
    For COLUMNA = 1 To 40
        For PROF = 1 To 3
            Transecto(FILA, COLUMNA, PROF) = -1
        Next PROF
    Next COLUMNA
Next FILA

'Cargando capa de transecto. Es un archivo CSV que define la ruta seguida en el transecto
'Archivo con 5 columnas
'Columna A: posición X de una celda atravesada por la línea de transecto
'Columna B: posición Y de una celda atravesada por la línea de transecto
'Columna C: número de eco-marcas detectadas en cada celda de la capa 1 atravesada por la línea de transecto
'Columna D: número de eco-marcas detectadas en cada celda de la capa 2 atravesada por la línea de transecto
'Columna E: número de eco-marcas detectadas en cada celda de la capa 3 atravesada por la línea de transecto
'Indicar la ruta como:
'Ruta$ = "C:\Datos\.....\Algoritmo DEF" + "\Transecto2.csv"
Open Ruta$ For Input As #1
    Input #1, F$
    Input #1, C$
    Input #1, NP1$
    Input #1, NP2$
    Input #1, NP3$
    For FILA = 1 To 8   'Número de celdas atravesadas por la linea de transecto en la capa1
        Input #1, Fi
        Input #1, Co
        Input #1, Npe1
        Input #1, Npe2
        Input #1, Npe3
        
        Transecto(Fi, Co, 1) = Npe1
        Transecto(Fi, Co, 2) = Npe2
        Transecto(Fi, Co, 3) = Npe3
    Next FILA
Close #1

'Estableciendo ausencia-presencia simulada para el total de la laguna
NRepHamm = 100
Repeticiones = 10000
AbundanciaObjetivo = 70000  'Estable la abundancia máxima para un simulación
Label6.Caption = "Densidad objetivo media: " + Format$(AbundanciaObjetivo / 523, "0.00")
LimiteHamm = 0.8
PruebasTotal = NRepHamm * Repeticiones
SumaPruebas = 0
ControlSEG = 0
MinEuclidea = 1E+17
ReDim Euclideas(Repeticiones, NRepHamm)
ReDim AbundanciaT(Repeticiones, NRepHamm)

TiempoStart = Timer
For RepHamm = 1 To NRepHamm
    Contador = 0
    ControlSEG = ControlSEG + 1
  
    If RepHamm > 1 Then
        TiempoFin = Timer
        Label3.Caption = "Restante: " + Format$(((TiempoFin - TiempoStart) * (NRepHamm - RepHamm - 1)) / 3600, "0.00") + " horas"
        TiempoStart = TiempoFin
    End If
        
    Do
        DoEvents
        ReDim Simulado(30, 40, 3)
        SumaHamming = 0
        Contador = Contador + 1
        Label2.Caption = "Buscando configuración de distribución para Hamming=" + Str$(LimiteHamm)
        
        For PROF = 1 To 3
            For FILA = 1 To 30
                For COLUMNA = 1 To 40
                
                    If Capas(FILA, COLUMNA, PROF) = 1 Then
                        Randomize Timer
                        ValorAleatorio = Rnd()
                    
                        'Probabilidades del TRANSECTO1 en función de la profundidad de la capa
                        'If PROF = 1 And ValorAleatorio >= 0.57 Then Simulado(FILA, COLUMNA, PROF) = 1
                        'If PROF = 2 And ValorAleatorio >= 0.24 Then Simulado(FILA, COLUMNA, PROF) = 1
                        'If PROF = 3 And ValorAleatorio >= 0.81 Then Simulado(FILA, COLUMNA, PROF) = 1
                        
                        'Probabilidades del TRANSECTO2*******
                        If PROF = 1 And ValorAleatorio >= 0.38 Then Simulado(FILA, COLUMNA, PROF) = 1
                        If PROF = 2 And ValorAleatorio >= 0.63 Then Simulado(FILA, COLUMNA, PROF) = 1
                        If PROF = 3 And ValorAleatorio >= 1 Then Simulado(FILA, COLUMNA, PROF) = 1
                    End If
                
                    'Recopilando comparaciones para calcular la similitud de Hamming
                    If Val(Transecto(FILA, COLUMNA, PROF)) > 0 And Simulado(FILA, COLUMNA, PROF) = 1 Then
                        SumaHamming = SumaHamming + 1
                    End If
                
                    If Val(Transecto(FILA, COLUMNA, PROF)) = 0 And Simulado(FILA, COLUMNA, PROF) = 0 Then
                        SumaHamming = SumaHamming + 1
                    End If
                
                Next COLUMNA
            Next FILA
        Next PROF
    
    Loop Until SumaHamming / 24 >= LimiteHamm    'El denominador depende del transecto. 63 para transecto #1 y 24 àra el transecto #1. Proporciona un Hamming relativo <=0.8
    Debug.Print "RepHamm:"; RepHamm, "Suma Hamming:"; SumaHamming, SumaHamming / 24, Contador    '63 para transecto 1

    'Repitiendo para la configuración que cumple Hamming
    For REP = 1 To Repeticiones
        SumaPruebas = SumaPruebas + 1
        Label1.Caption = Format$(SumaPruebas / PruebasTotal * 100, "0.00") + " %"
        Label2.Caption = "Calculando diferentes distribuciones de abundancia"
    
        'Creando una matriz de copia
        ReDim Copia(30, 40, 3)
        SumaAbundancia = 0
    
        For PROF = 1 To 3
            For FILA = 1 To 30
                For COLUMNA = 1 To 40
                    Copia(FILA, COLUMNA, PROF) = Simulado(FILA, COLUMNA, PROF)
                    If Copia(FILA, COLUMNA, PROF) = 1 Then
                        Randomize Timer
                        Valor = Rnd() * (AbundanciaObjetivo / 523)  '523=número de celdas consideradas)
                        Copia(FILA, COLUMNA, PROF) = Int(Valor)
                        SumaAbundancia = SumaAbundancia + Copia(FILA, COLUMNA, PROF)
                    End If
                Next COLUMNA
            Next FILA
        Next PROF
    
    
        'Calculando la distancia euclídea respecto al transecto
        SumaRestas = 0
        For PROF = 1 To 3
            For FILA = 1 To 30
                For COLUMNA = 1 To 40
                    If Transecto(FILA, COLUMNA, PROF) >= 0 Then
                        SumaRestas = SumaRestas + (Copia(FILA, COLUMNA, PROF) - Transecto(FILA, COLUMNA, PROF)) ^ 2
                    End If
                Next COLUMNA
            Next FILA
        Next PROF
        
        Euclideas(REP, RepHamm) = Sqr(SumaRestas)
        AbundanciaT(REP, RepHamm) = SumaAbundancia
        
        'Guardando la distribución de la abundancia para la distancia euclídea mínima
        'Mostrando resultados gráficos en el formulario
        If Euclideas(REP, RepHamm) < MinEuclidea Then
            Picture1.Cls
            Picture2.Cls
            Picture3.Cls
        
            ReDim Mejor(30, 40, 3)
            For PROF = 1 To 3
                For FILA = 1 To 30
                    For COLUMNA = 1 To 40
                        Mejor(FILA, COLUMNA, PROF) = Copia(FILA, COLUMNA, PROF)
                        
                        If PROF = 1 And Mejor(FILA, COLUMNA, PROF) > 0 Then
                            Intensidad = Mejor(FILA, COLUMNA, PROF)
                            If Mejor(FILA, COLUMNA, PROF) > 255 Then Intensidad = 255
                            Picture1.Line (COLUMNA - 1, FILA - 1)-(COLUMNA, FILA), RGB(255 - Intensidad, 255 - Intensidad, 255 - Intensidad), BF
                        End If
                        
                        If PROF = 2 And Mejor(FILA, COLUMNA, PROF) > 0 Then
                            Intensidad = Mejor(FILA, COLUMNA, PROF)
                            If Mejor(FILA, COLUMNA, PROF) > 255 Then Intensidad = 255
                            Picture2.Line (COLUMNA - 1, FILA - 1)-(COLUMNA, FILA), RGB(255, 255 - Intensidad, 255 - Intensidad), BF
                        End If
    
                        If PROF = 3 And Mejor(FILA, COLUMNA, PROF) > 0 Then
                            Intensidad = Mejor(FILA, COLUMNA, PROF)
                            If Mejor(FILA, COLUMNA, PROF) > 255 Then Intensidad = 255
                            Picture3.Line (COLUMNA - 1, FILA - 1)-(COLUMNA, FILA), RGB(255 - Intensidad, 255 - Intensidad, 255), BF
                        End If
                        
                        If Capas(FILA, COLUMNA, 1) = 1 Then Picture1.Line (COLUMNA - 1, FILA - 1)-(COLUMNA, FILA), RGB(0, 0, 0), B
                        If Capas(FILA, COLUMNA, 2) = 1 Then Picture2.Line (COLUMNA - 1, FILA - 1)-(COLUMNA, FILA), RGB(255, 0, 0), B
                        If Capas(FILA, COLUMNA, 3) = 1 Then Picture3.Line (COLUMNA - 1, FILA - 1)-(COLUMNA, FILA), RGB(0, 0, 255), B
                    Next COLUMNA
                Next FILA
            Next PROF
            Mejor(1, 0, 1) = LimiteHamm
            Mejor(2, 0, 1) = AbundanciaT(REP, RepHamm)
            Mejor(3, 0, 1) = RepHamm
            Mejor(4, 0, 1) = REP
            Mejor(5, 0, 1) = Euclideas(REP, RepHamm)
            MinEuclidea = Euclideas(REP, RepHamm)
            
            AbundanciaT(1, 0) = "Hamming"
            AbundanciaT(2, 0) = LimiteHamm
            AbundanciaT(3, 0) = "Configuración->COL"
            AbundanciaT(4, 0) = RepHamm
            AbundanciaT(5, 0) = "Repetición->FIL"
            AbundanciaT(6, 0) = REP
            AbundanciaT(7, 0) = "Abun. Min. Euclidea"
            AbundanciaT(8, 0) = SumaAbundancia
            AbundanciaT(9, 0) = "Densidad Media Objetivo"
            AbundanciaT(10, 0) = AbundanciaObjetivo / 523

            
            Euclideas(1, 0) = "Hamming"
            Euclideas(2, 0) = LimiteHamm
            Euclideas(3, 0) = "Configuración->COL"
            Euclideas(4, 0) = RepHamm
            Euclideas(5, 0) = "Repetición->FIL"
            Euclideas(6, 0) = REP
            Euclideas(7, 0) = "Abun. Min. Euclidea"
            Euclideas(8, 0) = SumaAbundancia
            Euclideas(9, 0) = "Densidad Media Objetivo"
            Euclideas(10, 0) = AbundanciaObjetivo / 523
            
            
            Label4.Caption = "Distancia: " + Format$(MinEuclidea, "0.00")
            Label5.Caption = "Abundancia Total: " + Format$(AbundanciaT(REP, RepHamm))
        End If
    
    Next REP
    
    'Calculando la distancia euclídea media para esta configuración
    SumaEuclidea = 0
    For Meu = 1 To Repeticiones
        SumaEuclidea = SumaEuclidea + Euclideas(Meu, RepHamm)
    Next Meu
    MedEuclidea = SumaEuclidea / Repeticiones
    Debug.Print "ME:", MedEuclidea
    
    'Copias de seguridad
    'Guardando distancias y distribución de abundancias
    If ControlSEG >= NRepHamm / 10 Then
        Fecha$ = Date
        hora$ = Timer
        CopiaSEG$ = Format$(SumaPruebas / PruebasTotal * 100, "0.00")
        RutaRES$ = "h:\Datos 180908\Artículos\2024 Zoñar 1\Algoritmo DEF3\" + "RES-SEG" + CopiaSEG$ + hora$ + ".csv"
        Open RutaRES$ For Output As #1
            For FILA = 1 To Repeticiones
                For COLUMNA = 0 To NRepHamm
                    Print #1, Euclideas(FILA, COLUMNA), ",",
                Next COLUMNA
                Print #1, Chr$(32)
            Next FILA
        Close #1
        ControlSEG = 0
        
        RutaRES$ = "h:\Datos 180908\Artículos\2024 Zoñar 1\Algoritmo DEF3\" + "CONF-SEG" + CopiaSEG$ + hora$ + ".csv"
        Open RutaRES$ For Output As #2
            For PROF = 1 To 3
                
                For FILA = 1 To 30
                    For COLUMNA = 0 To 40
                        Print #2, Mejor(FILA, COLUMNA, PROF), ",",
                    Next COLUMNA
                    Print #2, Chr$(32)
                Next FILA
                Print #2, "**********************************"
            Next PROF
        Close #2
    End If
Next RepHamm

'Guardando archivo final
Fecha$ = Date
hora$ = Timer
RutaRES$ = "h:\Datos 180908\Artículos\2024 Zoñar 1\Algoritmo DEF3\" + "RES" + hora$ + ".csv"
RutaABU$ = "h:\Datos 180908\Artículos\2024 Zoñar 1\Algoritmo DEF3\" + "ABU" + hora$ + ".csv"
Open RutaRES$ For Output As #1
Open RutaABU$ For Output As #2
    For FILA = 1 To Repeticiones
        For COLUMNA = 0 To NRepHamm
            Print #1, Euclideas(FILA, COLUMNA), ",",
            Print #2, AbundanciaT(FILA, COLUMNA), ",",
        Next COLUMNA
        Print #1, Chr$(32)
        Print #2, Chr$(32)
    Next FILA
Close #1
Close #2
Debug.Print "END", Time


End Sub

Private Sub Form_Load()

'Escala de los Gráfica-Lagunas de control
Picture1.Scale (-2, -2)-(37, 28)
Picture2.Scale (-2, -2)-(37, 28)
Picture3.Scale (-2, -2)-(37, 28)

End Sub
