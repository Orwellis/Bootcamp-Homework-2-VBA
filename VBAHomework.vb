Sub Easy():

'Define where were are collecting data from
Dim tickerCol As Integer
Dim volumeCol As Integer
Dim row As long
tickerCol = 1
volumeCol = 7
row = 2 'row 1 is the header

'Define where we will return the data we collect
dim returnTickerCol As Integer
dim returnVolumeCol As Integer
dim returnRow As long
'We want the Tickers to show up in column "I" (9)
returnTickerCol = 9
'We want the Volume to show up in column "J" (10)
returnVolumeCol = 10
'We want to start printing our data at row 2
returnRow = 1

'Define peramater for while loop
Dim rowHasData As Boolean

'Set up While loop by checking for first line of data
rowHasData = Not IsEmpty(Cells(row, tickerCol))


'xlup!!!

'Set up variables for while loop
Dim ticker As String
Dim volume As Long
Dim totalTickers as Long
totalTickers = 0
Dim uniqueTicker as Boolean
'The first ticker we find will be unique
uniqueTicker = True
dim oldTicker as Long

While rowHasData = True
    'Record first ticker (row starts at 2)
    ticker = Cells(row, tickerCol).Value
    'and its volume
    volume = Cells(row, volumeCol).Value
    'Check to see if this is an old Ticker Symbol
    'if it is add the current volume to the running total for that ticker
    for oldTicker = 2 to returnRow + 3
        if ticker = Cells(oldTicker,returnTickerCol).Value then
            uniqueTicker = False
            Cells(oldTicker,returnVolumeCol).Value = Cells(oldTicker,returnVolumeCol).Value + volume
            Exit For
        end if
    next oldTicker
    'If this is a ticker we haven't seen yet add it to the next row
    if uniqueTicker then
        returnRow = returnRow + 1
        Cells(returnRow,returnTickerCol).Value = ticker
        Cells(returnRow,returnVolumeCol).Value = volume
    End If
    'add some breaks as this code can take a long time to run
    if row mod 10000 = 0 then
        msgbox("We made it to row: " + str(row))
    end if
    'Reset uniqueTicker
    uniqueTicker = True
    'step row
    row = row + 1
    'update data checker for while loop
    rowHasData = Not IsEmpty(Cells(row, tickerCol))
wend

End Sub
