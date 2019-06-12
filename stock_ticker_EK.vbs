{\rtf1\ansi\ansicpg1252\cocoartf1671\cocoasubrtf500
{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub stocks()\
\
  Dim Ticker As String\
\
  Dim Ticker_Total As Double\
  Ticker_Total = 0\
  \
  Dim Summary_Table_Row As Integer\
  Summary_Table_Row = 2\
  lrow = Cells(Rows.Count - 1, 1).End(xlUp).Row\
  \
  For i = 2 To lrow\
\
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then\
\
      Ticker = Cells(i, 1).Value\
\
      Ticker_Total = Ticker_Total + Cells(i, 7).Value\
\
      Range("I" & Summary_Table_Row).Value = Ticker\
\
      Range("J" & Summary_Table_Row).Value = Ticker_Total\
\
      Summary_Table_Row = Summary_Table_Row + 1\
      \
      Ticker_Total = 0\
\
    Else\
\
      Ticker_Total = Ticker_Total + Cells(i, 7).Value\
\
    End If\
\
  Next i\
\
End Sub\
\
}