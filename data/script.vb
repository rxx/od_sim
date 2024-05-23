Sub SimHour(ByVal hr As Integer, ByVal funct As String)
'On Error Resume Next

    simhr = hr + 4
    'Starting at row 4 because of extra added row (due to uniform table headers)
    Dim actions As String
    Dim timeline As String
    'simulation hour section
    timeline = "====== Protection Hour: " & hr + 1 & _
    "  ( Local Time: " & FormatDateTime(Worksheets("Imps").Range("BY" & simhr).Value, vbLongTime) _
    & " " & FormatDateTime(Worksheets("Imps").Range("BY" & simhr).Value, vbShortDate) _
    & " )  ( Domtime: " & FormatDateTime(Worksheets("Imps").Range("BZ" & simhr).Value, vbLongTime) _
    & " " & FormatDateTime(Worksheets("Imps").Range("BZ" & simhr).Value, vbShortDate) _
    & " ) ======" & vbNewLine

    'draftrate section
    If Worksheets("Military").Range("Y" & simhr).Value <> "" And Worksheets("Military").Range("Y" & simhr).Value <> Worksheets("Military").Range("Z" & simhr - 1).Value Then
    actions = actions & "Draftrate changed to " & Worksheets("Military").Range("Y" & simhr).Value * 100 _
    & "%." & vbNewLine
    End If

    'releasing section
    'Check seems OK, note unit names in row 2. Table starts in row 3 with uniform headers


    un1 = Worksheets("Military").Range("AX" & simhr).Value
    un2 = Worksheets("Military").Range("AY" & simhr).Value
    un3 = Worksheets("Military").Range("AZ" & simhr).Value
    un4 = Worksheets("Military").Range("BA" & simhr).Value
    un5 = Worksheets("Military").Range("BB" & simhr).Value
    un6 = Worksheets("Military").Range("BC" & simhr).Value
    un7 = Worksheets("Military").Range("BD" & simhr).Value
    un8 = Worksheets("Military").Range("BE" & simhr).Value
    Draftees = Worksheets("Military").Range("AW" & simhr).Value
    released = un1 <> 0 Or un2 <> 0 Or un3 <> 0 Or un4 <> 0 Or un5 <> 0 Or un6 <> 0 Or un7 <> 0 Or un8 <> 0
    units = Array(un1, un2, un3, un4, un5, un6, un7, un8)
    unitname = Array( _
        Worksheets("Military").Range("AX2").Value, _
        Worksheets("Military").Range("AY2").Value, _
        Worksheets("Military").Range("AZ2").Value, _
        Worksheets("Military").Range("BA2").Value, _
        Worksheets("Military").Range("BB2").Value, _
        Worksheets("Military").Range("BC2").Value, _
        Worksheets("Military").Range("BD2").Value, _
        Worksheets("Military").Range("BE2").Value)
    comma = False
    If released = True Then
    actions = actions & "You successfully released "
    For u = 0 To 6
    If units(u) <> 0 Then
    If comma = True Then actions = actions & ", "
    actions = actions & units(u) & " " & unitname(u)
    comma = True
    End If
    Next u
    actions = actions & "." & vbNewLine
    End If
    If Draftees <> 0 Then
    actions = actions & "You successfully released " & Draftees & " draftees into the peasantry." & vbNewLine
    End If

    'Self spells section
    '
    If Worksheets("Explore").Range("S" & simhr).Value <> 0 And Worksheets("Magic").Range("G" & simhr).Value <> 0 Then
    actions = actions & "Your wizards successfully cast Gaia's Watch at a cost of " & _
    Format((Worksheets("Magic").Range("B" & simhr).Value - 20) * 2, "0") & " mana." & vbNewLine
    ElseIf Worksheets("Magic").Range("G" & simhr).Value <> 0 Then
    actions = actions & "Your wizards successfully cast Gaia's Watch at a cost of " & _
    Format(Worksheets("Magic").Range("B" & simhr).Value * 2, "0") & " mana." & vbNewLine
    End If

    If Worksheets("Explore").Range("S" & simhr).Value <> 0 And Worksheets("Magic").Range("H" & simhr).Value <> 0 Then
    actions = actions & "Your wizards successfully cast Mining Strength at a cost of " & _
    Format((Worksheets("Magic").Range("B" & simhr).Value - 20) * 2, "0") & " mana." & vbNewLine
    ElseIf Worksheets("Magic").Range("H" & simhr).Value <> 0 Then
    actions = actions & "Your wizards successfully cast Mining Strength at a cost of " & _
    Format(Worksheets("Magic").Range("B" & simhr).Value * 2, "0") & " mana." & vbNewLine
    End If

    If Worksheets("Explore").Range("S" & simhr).Value <> 0 And Worksheets("Magic").Range("I" & simhr).Value <> 0 Then
    actions = actions & "Your wizards successfully cast Ares Call at a cost of " & _
    Format((Worksheets("Magic").Range("B" & simhr).Value - 20) * 2.5, "0") & " mana." & vbNewLine
    ElseIf Worksheets("Magic").Range("I" & simhr).Value <> 0 Then
    actions = actions & "Your wizards successfully cast Ares Call at a cost of " & _
    Format(Worksheets("Magic").Range("B" & simhr).Value * 2.5, "0") & " mana." & vbNewLine
    End If

    If Worksheets("Explore").Range("S" & simhr).Value <> 0 And Worksheets("Magic").Range("J" & simhr).Value <> 0 Then
    actions = actions & "Your wizards successfully cast Midas Touch at a cost of " & _
    Format((Worksheets("Magic").Range("B" & simhr).Value - 20) * 2.5, "0") & " mana." & vbNewLine
    ElseIf Worksheets("Magic").Range("J" & simhr).Value <> 0 Then
    actions = actions & "Your wizards successfully cast Midas Touch at a cost of " & _
    Format(Worksheets("Magic").Range("B" & simhr).Value * 2.5, "0") & " mana." & vbNewLine
    End If

    If Worksheets("Explore").Range("S" & simhr).Value <> 0 And Worksheets("Magic").Range("K" & simhr).Value <> 0 Then
    actions = actions & "Your wizards successfully cast Harmony at a cost of " & _
    Format((Worksheets("Magic").Range("B" & simhr).Value - 20) * 2.5, "0") & " mana." & vbNewLine
    ElseIf Worksheets("Magic").Range("K" & simhr).Value <> 0 Then
    actions = actions & "Your wizards successfully cast Harmony at a cost of " & _
    Format(Worksheets("Magic").Range("B" & simhr).Value * 2.5, "0") & " mana." & vbNewLine
    End If

    'Racial spells section
    r1 = Worksheets("Magic").Range("L" & simhr).Value
    r2 = Worksheets("Magic").Range("M" & simhr).Value
    r3 = Worksheets("Magic").Range("N" & simhr).Value
    r4 = Worksheets("Magic").Range("O" & simhr).Value
    r5 = Worksheets("Magic").Range("P" & simhr).Value
    r6 = Worksheets("Magic").Range("Q" & simhr).Value
    r7 = Worksheets("Magic").Range("R" & simhr).Value
    r8 = Worksheets("Magic").Range("S" & simhr).Value
    r9 = Worksheets("Magic").Range("T" & simhr).Value
    r10 = Worksheets("Magic").Range("U" & simhr).Value
    racial = r1 <> 0 Or r2 <> 0 Or r3 <> 0 Or r4 <> 0 Or r5 <> 0 Or r6 <> 0 Or r7 <> 0 Or r8 <> 0 Or r9 <> 0 Or r10 <> 0

    If racial = True And Worksheets("Explore").Range("S" & simhr).Value <> 0 Then
    actions = actions & "Your wizards successfully cast Racial Spell at a cost of " & _
    Format((Worksheets("Magic").Range("B" & simhr).Value - 20) * 5, "0") & " mana." & vbNewLine
    ElseIf racial = True Then
    actions = actions & "Your wizards successfully cast Racial Spell at a cost of " & _
    Format(Worksheets("Magic").Range("B" & simhr).Value * 5, "0") & " mana." & vbNewLine
    End If

    'tech section
    If Worksheets("Techs").Range("K" & simhr).Value <> 0 Then
    actions = actions & "You have unlocked " & Worksheets("Techs").Range("CA" & simhr).Value & vbNewLing
    End If

    'daily plat section
    If Worksheets("Production").Range("C" & simhr).Value <> 0 Then
    actions = actions & "You have been awarded with " & Worksheets("Population").Range("C" & simhr).Value * 4 _
    & " platinum." & vbNewLine
    End If

    'national bank section
    I = 0
    plat = Worksheets("Production").Range("BC" & simhr).Value
    lumber = Worksheets("Production").Range("BD" & simhr).Value
    ore = Worksheets("Production").Range("BE" & simhr).Value
    gems = Worksheets("Production").Range("BF" & simhr).Value
    exchange = plat <> 0 Or wood <> 0 Or ore <> 0 Or gems <> 0
    If exchange = True Then
    If plat < 0 Then
    I = 1
    actions = actions & -plat & " platinum "
    End If
    If lumber < 0 Then
    If I > 0 Then
    I = 2
    actions = actions & "and " & -lumber & " lumber "
    Else
    I = 1
    actions = actions & -lumber & " lumber "
    End If
    End If
    If ore < 0 Then
    If I > 0 Then
    I = 2
    actions = actions & "and " & -ore & " ore "
    Else
    I = 1
    actions = actions & -ore & " ore "
    End If
    End If
    If gems < 0 Then
    If I > 0 Then
    I = 2
    actions = actions & "and " & -gems & " gems "
    Else
    I = 1
    actions = actions & -gems & " gems "
    End If
    End If
    actions = actions & "have been traded for "
    I = 0
    If plat > 0 Then
    I = 1
    actions = actions & plat & " platinum"
    End If
    If lumber > 0 Then
    If I > 0 Then
    I = 2
    actions = actions & " and " & lumber & " lumber"
    Else
    I = 1
    actions = actions & lumber & " lumber"
    End If
    End If
    If ore > 0 Then
    If I > 0 Then
    I = 2
    actions = actions & " and " & ore & " ore"
    Else
    I = 1
    actions = actions & ore & " ore"
    End If
    End If
    If gems > 0 Then
    If I > 0 Then
    I = 2
    actions = actions & " and " & gems & " gems"
    Else
    I = 1
    actions = actions & gems & " gems"
    End If
    End If
    actions = actions & "." & vbNewLine
    End If

    'exploring section
    plain = Worksheets("Explore").Range("T" & simhr).Value
    forest = Worksheets("Explore").Range("U" & simhr).Value
    mtn = Worksheets("Explore").Range("V" & simhr).Value
    hill = Worksheets("Explore").Range("W" & simhr).Value
    swamp = Worksheets("Explore").Range("X" & simhr).Value
    cavern = Worksheets("Explore").Range("Y" & simhr).Value
    water = Worksheets("Explore").Range("Z" & simhr).Value
    explore = plain <> 0 Or forest <> 0 Or mtn <> 0 Or hill <> 0 Or swamp <> 0 Or cavern <> 0 Or water <> 0
    Land = Array(plain, forest, mtn, hill, swamp, cavern, water)
    landname = Array("Plains", "Forest", "Mountains", "Hills", "Swamps", "Caverns", "Water")
    comma = False
    If explore = True Then
    actions = actions & "Exploration for "
    For e = 0 To 6
    If Land(e) <> 0 Then
    If comma = True Then actions = actions & ", "
    actions = actions & Land(e) & " " & landname(e)
    comma = True
    End If
    Next e
    actions = actions & " begun at a cost of " & Worksheets("Explore").Range("AH" & simhr).Value _
    & " platinum and " & Worksheets("Explore").Range("AI" & simhr).Value & " draftees." & vbNewLine
    End If

    'daily land section
    If Worksheets("Explore").Range("S" & simhr).Value <> 0 Then
    actions = actions & "You have been awarded with 20 " & Worksheets("Overview").Range("B70").Value _
    & "." & vbNewLine
    End If

    'destruction section
    'Check seems OK
    Hom = Worksheets("Construction").Range("BW" & simhr).Value
    Alc = Worksheets("Construction").Range("BX" & simhr).Value
    Far = Worksheets("Construction").Range("BY" & simhr).Value
    Smi = Worksheets("Construction").Range("BZ" & simhr).Value
    Mas = Worksheets("Construction").Range("CA" & simhr).Value
    Ly = Worksheets("Construction").Range("CB" & simhr).Value
    Hav = Worksheets("Construction").Range("CC" & simhr).Value
    OM = Worksheets("Construction").Range("CD" & simhr).Value
    GN = Worksheets("Construction").Range("CE" & simhr).Value
    Fac = Worksheets("Construction").Range("CF" & simhr).Value
    GT = Worksheets("Construction").Range("CG" & simhr).Value
    Bar = Worksheets("Construction").Range("CH" & simhr).Value
    Shr = Worksheets("Construction").Range("CI" & simhr).Value
    Tow = Worksheets("Construction").Range("CJ" & simhr).Value
    Tem = Worksheets("Construction").Range("CK" & simhr).Value
    WG = Worksheets("Construction").Range("CL" & simhr).Value
    DM = Worksheets("Construction").Range("CM" & simhr).Value
    Sch = Worksheets("Construction").Range("CN" & simhr).Value
    Doc = Worksheets("Construction").Range("CO" & simhr).Value
    destroy = Hom <> 0 Or Alc <> 0 Or Far <> 0 Or Smi <> 0 Or Mas <> 0 Or _
    Ly <> 0 Or Hav <> 0 Or OM <> 0 Or GN <> 0 Or Fac <> 0 Or GT <> 0 Or Bar <> 0 Or _
    Shr <> 0 Or Tow <> 0 Or Tem <> 0 Or WG <> 0 Or DM <> 0 Or Sch <> 0 Or Doc <> 0
    bldg = Array(Alc, Far, Smi, Mas, Ly, Hav, OM, GN, Fac, GT, Bar, Shr, Tow, _
    Tem, WG, DM, Sch, Doc)
    bldgname = Array("Alchemies", "Farms", "Smithies", "Masonries", "Lumber Yards", "Forest Havens", _
    "Ore Mines", "Gryphon Nests", "Factories", "Guard Towers", "Barracks", "Shrines", "Towers", _
    "Temples", "Wizard Guilds", "Diamond Mines", "Schools", "Docks")
    If destroy = True Then
    actions = actions & "Destruction of "
    If Hom <> 0 Then
    actions = actions & Hom & " Homes"
    comma = True
    End If
    For b = 0 To 17
    If bldg(b) <> 0 Then
    If comma = True Then actions = actions & ", "
    actions = actions & bldg(b) & " " & bldgname(b)
    comma = True
    End If
    Next b
    actions = actions & " is complete." & vbNewLine
    End If

    'rezoning section
    plain = Worksheets("Rezone").Range("L" & simhr).Value
    forest = Worksheets("Rezone").Range("M" & simhr).Value
    mtn = Worksheets("Rezone").Range("N" & simhr).Value
    hill = Worksheets("Rezone").Range("O" & simhr).Value
    swamp = Worksheets("Rezone").Range("P" & simhr).Value
    cavern = Worksheets("Rezone").Range("Q" & simhr).Value
    water = Worksheets("Rezone").Range("R" & simhr).Value
    rezone = plain <> 0 Or forest <> 0 Or mtn <> 0 Or hill <> 0 Or swamp <> 0 Or cavern <> 0 Or water <> 0
    Land = Array(plain, forest, mtn, hill, swamp, cavern, water)
    landname = Array("Plains", "Forest", "Mountains", "Hills", "Swamps", "Caverns", "Water")
    comma = False
    If rezone = True Then
    actions = actions & "Rezoning begun at a cost of " & Worksheets("Rezone").Range("Y" & simhr).Value _
    & " platinum. The changes in land are as following: "
    For r = 0 To 6
    If Land(r) <> 0 Then
    If comma = True Then actions = actions & ", "
    actions = actions & Land(r) & " " & landname(r)
    comma = True
    End If
    Next r
    actions = actions & vbNewLine
    End If

    'construction section
    'Check seems OK
    Hom = Worksheets("Construction").Range("O" & simhr).Value
    Alc = Worksheets("Construction").Range("P" & simhr).Value
    Far = Worksheets("Construction").Range("Q" & simhr).Value
    Smi = Worksheets("Construction").Range("R" & simhr).Value
    Mas = Worksheets("Construction").Range("S" & simhr).Value
    Ly = Worksheets("Construction").Range("T" & simhr).Value
    Hav = Worksheets("Construction").Range("U" & simhr).Value
    OM = Worksheets("Construction").Range("V" & simhr).Value
    GN = Worksheets("Construction").Range("W" & simhr).Value
    Fac = Worksheets("Construction").Range("X" & simhr).Value
    GT = Worksheets("Construction").Range("Y" & simhr).Value
    Bar = Worksheets("Construction").Range("Z" & simhr).Value
    Shr = Worksheets("Construction").Range("AA" & simhr).Value
    Tow = Worksheets("Construction").Range("AB" & simhr).Value
    Tem = Worksheets("Construction").Range("AC" & simhr).Value
    WG = Worksheets("Construction").Range("AD" & simhr).Value
    DM = Worksheets("Construction").Range("AE" & simhr).Value
    Sch = Worksheets("Construction").Range("AF" & simhr).Value
    Doc = Worksheets("Construction").Range("AG" & simhr).Value
    construct = Hom <> 0 Or Alc <> 0 Or Far <> 0 Or Smi <> 0 Or Mas <> 0 Or _
    Ly <> 0 Or Hav <> 0 Or OM <> 0 Or GN <> 0 Or Fac <> 0 Or GT <> 0 Or Bar <> 0 Or _
    Shr <> 0 Or Tow <> 0 Or Tem <> 0 Or WG <> 0 Or DM <> 0 Or Sch <> 0 Or Doc <> 0
    bldg = Array(Alc, Far, Smi, Mas, Ly, Hav, OM, GN, Fac, GT, Bar, Shr, Tow, _
    Tem, WG, DM, Sch, Doc)
    bldgname = Array("Alchemies", "Farms", "Smithies", "Masonries", "Lumber Yards", "Forest Havens", _
    "Ore Mines", "Gryphon Nests", "Factories", "Guard Towers", "Barracks", "Shrines", "Towers", _
    "Temples", "Wizard Guilds", "Diamond Mines", "Schools", "Docks")
    comma = False
    If construct = True Then
    actions = actions & "Construction of "
    If Hom <> 0 Then
    actions = actions & Hom & " Homes"
    comma = True
    End If
    For b = 0 To 17
    If bldg(b) <> 0 Then
    If comma = True Then actions = actions & ", "
    actions = actions & bldg(b) & " " & bldgname(b)
    comma = True
    End If
    Next b
    actions = actions & " started at a cost of " & Worksheets("Construction"). _
    Range("AQ" & simhr).Value & " platinum and " & Worksheets("Construction"). _
    Range("AR" & simhr).Value & " lumber." & vbNewLine
    End If

    'training military section
    'Check seems OK, note unit names in row 2. Table starts in row 3 with uniform headers

    un1 = Worksheets("Military").Range("AG" & simhr).Value
    un2 = Worksheets("Military").Range("AH" & simhr).Value
    un3 = Worksheets("Military").Range("AI" & simhr).Value
    un4 = Worksheets("Military").Range("AJ" & simhr).Value
    un5 = Worksheets("Military").Range("AK" & simhr).Value
    un6 = Worksheets("Military").Range("AL" & simhr).Value + 0
    un7 = Worksheets("Military").Range("AM" & simhr).Value
    un8 = Worksheets("Military").Range("AN" & simhr).Value + 0
    Draftees = un1 + un2 + un3 + un4 + un5 + un7
    trained = un1 <> 0 Or un2 <> 0 Or un3 <> 0 Or un4 <> 0 Or un5 <> 0 Or un6 <> 0 Or un7 <> 0 Or un8 <> 0
    units = Array(un1, un2, un3, un4, un5, un6, un7, un8)
    unitname = Array( _
        Worksheets("Military").Range("AG2").Value, _
        Worksheets("Military").Range("AH2").Value, _
        Worksheets("Military").Range("AI2").Value, _
        Worksheets("Military").Range("AJ2").Value, _
        Worksheets("Military").Range("AK2").Value, _
        Worksheets("Military").Range("AL2").Value, _
        Worksheets("Military").Range("AM2").Value, _
        Worksheets("Military").Range("AN2").Value)
    comma = False
    If trained = True Then
    actions = actions & "Training of "
    For u = 0 To 7
    If units(u) <> 0 Then
    If comma = True Then actions = actions & ", "
    actions = actions & units(u) & " " & unitname(u)
    comma = True
    End If
    Next u
    actions = actions & " begun at a cost of " & Worksheets("Military").Range("AR" & simhr).Value _
    & " platinum, " & Worksheets("Military").Range("AS" & simhr).Value & " ore, " & Draftees & " draftees, " & un6 & " spies, and " & _
    un8 & " wizards." & vbNewLine
    End If

    'improvements section
    If Worksheets("Imps").Range("P" & simhr).Value <> 0 Then
    actions = actions & "You invested " & Worksheets("Imps").Range("P" & simhr).Value _
    & " " & Worksheets("Imps").Range("O" & simhr).Value & " into " & _
    Worksheets("Imps").Range("Q" & simhr).Value & "." & vbNewLine
    End If
    If Worksheets("Imps").Range("S" & simhr).Value <> 0 Then
    actions = actions & "You invested " & Worksheets("Imps").Range("S" & simhr).Value _
    & " " & Worksheets("Imps").Range("R" & simhr).Value & " into " & _
    Worksheets("Imps").Range("T" & simhr).Value & "." & vbNewLine
    End If
    If Worksheets("Imps").Range("V" & simhr).Value <> 0 Then
    actions = actions & "You invested " & Worksheets("Imps").Range("V" & simhr).Value _
    & " " & Worksheets("Imps").Range("U" & simhr).Value & " into " & _
    Worksheets("Imps").Range("W" & simhr).Value & "." & vbNewLine
    End If

    'cut out empty hours
    If actions <> "" Then
    actions = timeline & actions & vbNewLine
    End If

    If funct = "log" Then
    Log.Value = Log.Value & actions
    End If

    If funct = "msg" Then
      If actions = "" Then

        actions = timeline & vbNewLine & "Nothing needs to be done this hour"
      End If
      MsgBox actions
    End If


End Sub

Sub loghr()
'On Error Resume Next
    Log.Value = ""
    Worksheets("Log").TextBoxes("bork").Text = ""

    For hr = 0 To 83
        SimHour hr, "log"
    Next hr
    Worksheets("Log").TextBoxes("bork").Text = Log.Value

End Sub

Sub TextboxCopy()
    If Worksheets("Log").TextBoxes("bork").Text = "" Then
    Exit Sub
    End If

    ClipBoard_SetData Worksheets("Log").TextBoxes("bork").Text

End Sub

Sub clear()
    Worksheets("Log").TextBoxes("bork").Text = ""
End Sub

Sub currenthour()
    Application.Calculate
    hournum = Worksheets("Overview").Range("E17").Value - 1
    If hournum < 0 Then hournum = 0
    SimHour hournum, "msg"
End Sub

Sub stats()
On Error Resume Next

    loghour = Range("I28").Value
    If loghour = "" Then
        loghour = 72
    End If
    hr = loghour + 4

    'clearsight
    statstr = "The Dominion of Simulated Dominion: Hour " & loghour & vbNewLine & "Overview" & vbNewLine & "Ruler:  " & "Whatever" & vbNewLine & _
    "Race:  " & Worksheets("Overview").Range("B14").Value & vbNewLine & "Land:  " & Format(Worksheets("Production").Range("E" & hr).Value, "#,###") & vbNewLine & _
    "Peasants:  " & Format(Worksheets("Population").Range("C" & hr).Value, "#,###") & vbNewLine & "Draftees:  " & Format(Worksheets("Population").Range("E" & hr).Value, "#,###") & vbNewLine & _
    "Employment:  " & Format(Worksheets("Population").Range("I" & hr).Value * 100, "0.00") & "%" & vbNewLine & "Networth:  " & Format(Worksheets("Production").Range("G" & hr).Value, "#,###") & vbNewLine & _
    "Resources" & vbNewLine & "Platinum:  " & Format(Worksheets("Production").Range("H" & hr).Value, "#,###") & vbNewLine & _
    "Food:  " & Format(Worksheets("Production").Range("I" & hr).Value, "#,##0") & vbNewLine & "Lumber:  " & Format(Worksheets("Production").Range("J" & hr).Value, "#,##0") & vbNewLine & _
    "Mana:  " & Format(Worksheets("Production").Range("K" & hr).Value, "#,##0") & vbNewLine & "Ore:  " & Format(Worksheets("Production").Range("L" & hr).Value, "#,##0") & vbNewLine & _
    "Gems:  " & Format(Worksheets("Production").Range("M" & hr).Value, "#,##0") & vbNewLine & "Boats:  " & Format(Worksheets("Production").Range("N" & hr).Value, "#,##0") & vbNewLine & _
    "Military" & vbNewLine & "Morale:  100.00%" & vbNewLine & _
    Worksheets("Overview").Range("A36").Value & ":  " & Format(Worksheets("Military").Range("E" & hr).Value, "#,##0") & vbNewLine & _
    Worksheets("Overview").Range("A37").Value & ":  " & Format(Worksheets("Military").Range("F" & hr).Value, "#,##0") & vbNewLine & _
    Worksheets("Overview").Range("A38").Value & ":  " & Format(Worksheets("Military").Range("G" & hr).Value, "#,##0") & vbNewLine & _
    Worksheets("Overview").Range("A39").Value & ":  " & Format(Worksheets("Military").Range("H" & hr).Value, "#,##0") & vbNewLine & _
    "Spies:  " & Format(Worksheets("Military").Range("I" & hr).Value, "#,##0") & vbNewLine & _
    "Archspies: " & Format(Worksheets("Military").Range("J" & hr).Value, "#,##0") & vbNewLine & _
    "Wizards:  " & Format(Worksheets("Military").Range("K" & hr).Value, "#,##0") & vbNewLine & _
    "Archmages:  " & Format(Worksheets("Military").Range("L" & hr).Value, "#,##0") & vbNewLine

    statstr = statstr & vbNewLine & "--------------------------------------------------" & vbNewLine

    'buildings
    'edited to be easier to change fields. See tab Log_Support for table

    'Old formula
    'statstr = statstr & "Homes:  " & Worksheets("Construction").Range("AS" & hr).Value & "  " & Format(((Worksheets("Construction").Range("AS" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Farms:  " & Worksheets("Construction").Range("AU" & hr).Value & "  " & Format(((Worksheets("Construction").Range("AU" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Factories:  " & Worksheets("Construction").Range("BB" & hr).Value & "  " & Format(((Worksheets("Construction").Range("BB" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Shrines:  " & Worksheets("Construction").Range("BE" & hr).Value & "  " & Format(((Worksheets("Construction").Range("BE" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Smithies:  " & Worksheets("Construction").Range("AV" & hr).Value & "  " & Format(((Worksheets("Construction").Range("AV" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Ore Mines:  " & Worksheets("Construction").Range("AZ" & hr).Value & "  " & Format(((Worksheets("Construction").Range("AZ" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Diamond Mines:  " & Worksheets("Construction").Range("BI" & hr).Value & "  " & Format(((Worksheets("Construction").Range("BI" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Schools:  " & Worksheets("Construction").Range("BJ" & hr).Value & "  " & Format(((Worksheets("Construction").Range("BJ" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Alchemies:  " & Worksheets("Construction").Range("AT" & hr).Value & "  " & Format(((Worksheets("Construction").Range("AT" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Lumberyards:  " & Worksheets("Construction").Range("AX" & hr).Value & "  " & Format(((Worksheets("Construction").Range("AX" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Towers:  " & Worksheets("Construction").Range("BF" & hr).Value & "  " & Format(((Worksheets("Construction").Range("BF" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Docks:  " & Worksheets("Construction").Range("BK" & hr).Value & "  " & Format(((Worksheets("Construction").Range("BK" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Masonries:  " & Worksheets("Construction").Range("AW" & hr).Value & "  " & Format(((Worksheets("Construction").Range("AW" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Gryphon Nests:  " & Worksheets("Construction").Range("BA" & hr).Value & "  " & Format(((Worksheets("Construction").Range("BA" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Barracks:  " & Worksheets("Construction").Range("BD" & hr).Value & "  " & Format(((Worksheets("Construction").Range("BD" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Guard Towers:  " & Worksheets("Construction").Range("BC" & hr).Value & "  " & Format(((Worksheets("Construction").Range("BC" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Wizard Guilds:  " & Worksheets("Construction").Range("BH" & hr).Value & "  " & Format(((Worksheets("Construction").Range("BH" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Temples:  " & Worksheets("Construction").Range("BG" & hr).Value & "  " & Format(((Worksheets("Construction").Range("BG" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Forest Havens:  " & Worksheets("Construction").Range("AY" & hr).Value & "  " & Format(((Worksheets("Construction").Range("AY" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Barren Land:  " & Worksheets("Construction").Range("A" & hr).Value & "  " & Format(((Worksheets("Construction").Range("A" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100), "0.00") & "%" & vbNewLine & _
    "Incoming Buildings:  " & Worksheets("Construction").Range("C" & hr).Value & "  " & Format((Worksheets("Construction").Range("C" & hr).Value / Worksheets("Construction").Range("E" & hr).Value) * 100, "0.00") & "%" & vbNewLine

    Dim tbl As ListObject
    Dim lrow As ListRow

    Set tbl = Worksheets("Log_support").ListObjects("TblStats")

    For Each lrow In tbl.ListRows
       statstr = statstr & lrow.Range(1, tbl.ListColumns("Stats").Index).Value & lrow.Range(1, tbl.ListColumns("Number").Index).Value & " " & FormatPercent(lrow.Range(1, tbl.ListColumns("Percent").Index).Value, 2) & vbNewLine
    Next

    Worksheets("Log").TextBoxes("bork2").Text = statstr

End Sub


Sub clearstat()
    Worksheets("Log").TextBoxes("bork2").Text = ""
End Sub

Sub StatCopy()
    If Worksheets("Log").TextBoxes("bork2").Text = "" Then
    Exit Sub
    End If

    ClipBoard_SetData Worksheets("Log").TextBoxes("bork2").Text

End Sub

Sub ExtractSummary()
    Dim s As String
    Dim I As Integer
    For I = 1 To 133
        s = s & HourSummary(I)
    Next I

    'to do:
    'output s in some way

    Dim c As DataObject
    Set c = New DataObject
    c.SetText s
    c.PutInClipboard
End Sub

Function HourSummary(hr As Integer) As String
    simhr = hr + 3
    Dim actions, timestring As String

    Dim simDate As Date

    simDate = Worksheets("Imps").Range("BZ" & simhr).Value

    timestring = Year(simDate) & "-" & Day(simDate) & "-" & Month(simDate) & " " & Hour(simDate)

    actions = ""

    'exploring section
    plain = Worksheets("Explore").Range("T" & simhr).Value
    forest = Worksheets("Explore").Range("U" & simhr).Value
    mtn = Worksheets("Explore").Range("V" & simhr).Value
    hill = Worksheets("Explore").Range("W" & simhr).Value
    swamp = Worksheets("Explore").Range("X" & simhr).Value
    cavern = Worksheets("Explore").Range("Y" & simhr).Value
    water = Worksheets("Explore").Range("Z" & simhr).Value
    Land = Array(plain, forest, mtn, hill, swamp, cavern, water)
    landname = Array("Plain", "Forest", "Mountain", "Hill", "Swamp", "Cavern", "Water")
    For e = 0 To 6
      If Land(e) <> 0 Then
        actions = actions & timestring & ",Explore," & landname(e) & "," & Land(e) & vbNewLine
      End If
    Next e

    'daily bonus section
    If Worksheets("Explore").Range("S" & simhr).Value <> 0 Then
        actions = actions & timestring & ",LandBonus" & vbNewLine
    End If
    If Worksheets("Production").Range("C" & simhr).Value <> 0 Then
        actions = actions & timestring & ",PlatBonus" & vbNewLine
    End If

    'rezoning section
    plain = Worksheets("Rezone").Range("L" & simhr).Value
    forest = Worksheets("Rezone").Range("M" & simhr).Value
    mtn = Worksheets("Rezone").Range("N" & simhr).Value
    hill = Worksheets("Rezone").Range("O" & simhr).Value
    swamp = Worksheets("Rezone").Range("P" & simhr).Value
    cavern = Worksheets("Rezone").Range("Q" & simhr).Value
    water = Worksheets("Rezone").Range("R" & simhr).Value
    Land = Array(plain, forest, mtn, hill, swamp, cavern, water)
    landname = Array("Plain", "Forest", "Mountain", "Hill", "Swamp", "Cavern", "Water")
    For r = 0 To 6
      If Land(r) <> 0 Then
        actions = actions & timestring & ",Rezone," & landname(r) & "," & Land(r) & vbNewLine
      End If
    Next r

    'Self spells section
    If Worksheets("Magic").Range("G" & simhr).Value <> 0 Then
        actions = actions & timestring & ",Cast,Gaia's Watch" & vbNewLine
    End If
    If Worksheets("Magic").Range("H" & simhr).Value <> 0 Then
        actions = actions & timestring & ",Cast,Mining Strength" & vbNewLine
    End If
    If Worksheets("Magic").Range("I" & simhr).Value <> 0 Then
        actions = actions & timestring & ",Cast,Ares' Call" & vbNewLine
    End If
    If Worksheets("Magic").Range("J" & simhr).Value <> 0 Then
        actions = actions & timestring & ",Cast,Midas Touch" & vbNewLine
    End If
    If Worksheets("Magic").Range("K" & simhr).Value <> 0 Then
        actions = actions & timestring & ",Cast,Harmony" & vbNewLine
    End If

    'Racial spells section
    r1 = Worksheets("Magic").Range("L" & simhr).Value
    r2 = Worksheets("Magic").Range("M" & simhr).Value
    r3 = Worksheets("Magic").Range("N" & simhr).Value
    r4 = Worksheets("Magic").Range("O" & simhr).Value
    r5 = Worksheets("Magic").Range("P" & simhr).Value
    r6 = Worksheets("Magic").Range("Q" & simhr).Value
    r7 = Worksheets("Magic").Range("R" & simhr).Value
    r8 = Worksheets("Magic").Range("S" & simhr).Value
    racial = r1 <> 0 Or r2 <> 0 Or r3 <> 0 Or r4 <> 0 Or r5 <> 0 Or r6 <> 0 Or r7 <> 0 Or r8 <> 0
    If racial Then actions = actions & timestring & ",Cast,Race Spell" & vbNewLine

    'destruction section
    'Changed due to column changes
    Hom = Worksheets("Construction").Range("BV" & simhr).Value
    Alc = Worksheets("Construction").Range("BW" & simhr).Value
    Far = Worksheets("Construction").Range("BX" & simhr).Value
    Smi = Worksheets("Construction").Range("BY" & simhr).Value
    Mas = Worksheets("Construction").Range("BZ" & simhr).Value
    Ly = Worksheets("Construction").Range("CA" & simhr).Value
    Hav = Worksheets("Construction").Range("CB" & simhr).Value
    OM = Worksheets("Construction").Range("CC" & simhr).Value
    GN = Worksheets("Construction").Range("CD" & simhr).Value
    Fac = Worksheets("Construction").Range("CE" & simhr).Value
    GT = Worksheets("Construction").Range("CF" & simhr).Value
    Bar = Worksheets("Construction").Range("CG" & simhr).Value
    Shr = Worksheets("Construction").Range("CH" & simhr).Value
    Tow = Worksheets("Construction").Range("CI" & simhr).Value
    Tem = Worksheets("Construction").Range("CJ" & simhr).Value
    WG = Worksheets("Construction").Range("CK" & simhr).Value
    DM = Worksheets("Construction").Range("CL" & simhr).Value
    Sch = Worksheets("Construction").Range("CM" & simhr).Value
    Doc = Worksheets("Construction").Range("CN" & simhr).Value
    bldg = Array(Hom, Alc, Far, Smi, Mas, Ly, Hav, OM, GN, Fac, GT, Bar, Shr, Tow, _
    Tem, WG, DM, Sch, Doc)
    bldgname = Array("Home", "Alchemy", "Farm", "Smithy", "Masonry", "Lumberyard", "Forest Haven", _
    "Ore Mine", "Gryphon Nest", "Factory", "Guard Tower", "Barracks", "Shrine", "Tower", _
    "Temple", "Wizard Guild", "Diamond Mine", "School", "Dock")
    For b = 0 To 18
       If bldg(b) <> 0 Then actions = actions & timestring & ",Destroy," & bldgname(b) & "," & bldg(b) & vbNewLine
    Next b


    'construction section
    'Changed due to column changes
    Hom = Worksheets("Construction").Range("N" & simhr).Value
    Alc = Worksheets("Construction").Range("O" & simhr).Value
    Far = Worksheets("Construction").Range("P" & simhr).Value
    Smi = Worksheets("Construction").Range("Q" & simhr).Value
    Mas = Worksheets("Construction").Range("R" & simhr).Value
    Ly = Worksheets("Construction").Range("S" & simhr).Value
    Hav = Worksheets("Construction").Range("T" & simhr).Value
    OM = Worksheets("Construction").Range("U" & simhr).Value
    GN = Worksheets("Construction").Range("V" & simhr).Value
    Fac = Worksheets("Construction").Range("W" & simhr).Value
    GT = Worksheets("Construction").Range("X" & simhr).Value
    Bar = Worksheets("Construction").Range("Y" & simhr).Value
    Shr = Worksheets("Construction").Range("Z" & simhr).Value
    Tow = Worksheets("Construction").Range("AA" & simhr).Value
    Tem = Worksheets("Construction").Range("AB" & simhr).Value
    WG = Worksheets("Construction").Range("AC" & simhr).Value
    DM = Worksheets("Construction").Range("AD" & simhr).Value
    Sch = Worksheets("Construction").Range("AE" & simhr).Value
    Doc = Worksheets("Construction").Range("AF" & simhr).Value
    bldg = Array(Hom, Alc, Far, Smi, Mas, Ly, Hav, OM, GN, Fac, GT, Bar, Shr, Tow, _
    Tem, WG, DM, Sch, Doc)
    bldgname = Array("Home", "Alchemy", "Farm", "Smithy", "Masonry", "Lumberyard", "Forest Haven", _
    "Ore Mine", "Gryphon Nest", "Factory", "Guard Tower", "Barracks", "Shrine", "Tower", _
    "Temple", "Wizard Guild", "Diamond Mine", "School", "Dock")
    For b = 0 To 18
       If bldg(b) <> 0 Then actions = actions & timestring & ",Construct," & bldgname(b) & "," & bldg(b) & vbNewLine
    Next b

    'draftrate section
    If (Worksheets("Military").Range("Y" & simhr).Value <> Worksheets("Military").Range("Z" & simhr - 1).Value) _
     And Worksheets("Military").Range("Y" & simhr).Value <> 0 Then
     actions = actions & timestring & ",Draft," & Worksheets("Military").Range("V" & simhr).Value * 100 & "%" & vbNewLine
    End If

    'releasing section
    un1 = Worksheets("Military").Range("AX" & simhr).Value
    un2 = Worksheets("Military").Range("AY" & simhr).Value
    un3 = Worksheets("Military").Range("AZ" & simhr).Value
    un4 = Worksheets("Military").Range("BA" & simhr).Value
    un5 = Worksheets("Military").Range("BB" & simhr).Value
    un6 = Worksheets("Military").Range("BC" & simhr).Value
    un7 = Worksheets("Military").Range("BD" & simhr).Value
    un8 = Worksheets("Military").Range("BE" & simhr).Value
    un9 = Worksheets("Military").Range("AW" & simhr).Value
    units = Array(un1, un2, un3, un4, un5, un6, un7, un8, un9)
    unitname = Array("Spec1", "Spec2", "Elite1", "Elite2", "Spies", "Archspies", "Wizards", "Archmages", "Draftees")
    For u = 0 To 8
        If units(u) <> 0 Then
            actions = actions & timestring & ",Release," & unitname(u) & "," & units(u) & vbNewLine
        End If
    Next u

    'training military section
    un1 = Worksheets("Military").Range("AG" & simhr).Value
    un2 = Worksheets("Military").Range("AH" & simhr).Value
    un3 = Worksheets("Military").Range("AI" & simhr).Value
    un4 = Worksheets("Military").Range("AJ" & simhr).Value
    un5 = Worksheets("Military").Range("AK" & simhr).Value
    un6 = Worksheets("Military").Range("AL" & simhr).Value
    un7 = Worksheets("Military").Range("AM" & simhr).Value
    un8 = Worksheets("Military").Range("AN" & simhr).Value
    units = Array(un1, un2, un3, un4, un5, un6, un7, un8)
    unitname = Array("Spec1", "Spec2", "Elite1", "Elite2", "Spies", "Archspies", "Wizards", "Archmages")
    For u = 0 To 7
        If units(u) <> 0 Then
            actions = actions & timestring & ",Train," & unitname(u) & "," & units(u) & vbNewLine
        End If
    Next u

    '****************************
    '*******    to do:    *******
    '****************************
    ' keep it simple, or......
    '
    ' bank?
    ' improvements?
    ' action order, e.g.
    '  -explore before daily land bonus
    '  -construct before daily land bonus, except the 15 land from landbonus
    '  -rezone before daily land bonus, except 15 land from landbonus
    '  -release draftees, then take platbonus (applicable?)
    '  -destroy factories, then construct
    '  -destroy factories after constructing the rest? what about gradual factory killing?
    '
    '****************************

    HourSummary = actions

End Function




