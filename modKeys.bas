Attribute VB_Name = "modKeys"
Option Compare Database
Option Explicit

'==============================================================================
' modKeys - the join key shared by ZPT receipts and the OMM3 plan  (v1)
'==============================================================================
' Reads   : tbl_CustNameMap, tbl_EngFamilyMap  (LOCAL, this .accdb)
' Writes  : nothing at run time. RBDB_SetupKeyMaps() creates the two tables.
' Sends   : nothing.
'
' FIRST RUN : RBDB_SetupKeyMaps()   <- creates the two map tables if missing
' TEST      : RBDB_TestKeys()       <- asserts every rule below, prints PASS/FAIL
' EXPORT    : RBDB_ExportMaps()     <- writes the maps as Power Query M
'
'==============================================================================
' WHY THIS MODULE EXISTS
'==============================================================================
'   The OMM3 plan is keyed on Customer + Engine Model. ZPT receipts are keyed
'   on Cust Name + Eng Model. They are the same commercial relationship
'   written two different ways, and before these three functions existed the
'   plan and the actuals for one customer landed on two different keys and
'   BOTH sides reported as a failure.
'
'   Merging at engine-family level moved measured attainment from 0.43 to
'   1.04. The earlier figure was a key-matching artifact, not business
'   performance. That is the entire justification for this module: without it
'   every downstream number is wrong in a way that looks plausible.
'
'==============================================================================
' TWO IMPLEMENTATIONS, ONE MEANING - READ THIS BEFORE EDITING A MAP
'==============================================================================
'   These three functions also exist in Power Query (M) as fnNormText,
'   fnCustName and fnEngFamily. If one side gains a mapping and the other does
'   not, keys drift apart and NOBODY NOTICES FOR MONTHS - the counts still
'   look reasonable, the totals still add up, and the attainment figure is
'   quietly measuring the wrong thing.
'
'   So the maps live in TABLES, not in code:
'
'       tbl_CustNameMap  (SourceValue, MappedValue, Notes)
'       tbl_EngFamilyMap (SourceValue, FamilyValue, Notes)
'
'   Access is the master. RBDB_ExportMaps() writes the current contents as a
'   paste-ready M literal, so the Power Query side is REGENERATED rather than
'   maintained by hand, and drift shows up as a diff instead of as a silent
'   divergence. A live Power Query link to this .accdb was considered and
'   rejected: it makes the dashboard refresh depend on this file being
'   reachable, which is a worse failure than the one it prevents.
'
'   ADD A MAPPING: insert a row in the table, run RBDB_TestKeys(), run
'   RBDB_ExportMaps(), paste the output into the .pbix. All four steps.
'
'   If the tables are missing the module falls back to the built-in seed
'   below and MapsAreTableDriven() returns False, so a diagnostic can say so
'   out loud. It never silently runs on a half-configured database.
'
'==============================================================================
' THE KEY
'==============================================================================
'   CustEngPlantKey = <normalised customer> | <engine family> | <plant code>
'
'   Built identically on both sides:
'     ZPT   [Cust Name], [Eng Model], [Plant]     ("5610" / "5650")
'     OMM3  [Customer],  [Engine Model], [Plant]  ("5610" / "5650")
'
'   KEY_SEP IS PART OF THE CONTRACT. The separator never appears in a
'   customer name or an engine model, so it cannot create a false match - but
'   if the Power Query side ever joins ITS key to one exported from here, the
'   two separators have to be the same character. Change it in both or in
'   neither.
'
'==============================================================================
' PERFORMANCE NOTE - WHY THE KEY IS NOT BUILT IN SQL
'==============================================================================
'   These functions are Public, so Access's query engine CAN call them from a
'   SELECT - the way PlanNum() is called in RefreshPlan. Do not do that for
'   the key.
'
'   modAnomalyAlert's AggMap() interpolates its key expression straight into
'   a SELECT and a GROUP BY, and it runs dozens of times per email (weeks x
'   buckets x grains x plants). A per-row VBA callback there would be paid
'   over and over, and it re-introduces exactly the failure the SQL
'   EXPRESSION WARNING at the top of modAnomalyAlert describes: a project
'   that will not compile takes the whole run down with it.
'
'   Instead the key is MATERIALISED ONCE, at staging time, over the ~380
'   DISTINCT customer/engine/plant combinations rather than over every
'   receipt row - then joined back. Downstream every GROUP BY is a bare
'   column name, which is what AggMap is documented to accept.
'
'==============================================================================
' NORMALISATION - EXACT PARITY WITH fnNormText
'==============================================================================
'   In order:
'     1. null-safe            null -> ""
'     2. NBSP -> space        U+00A0 is not whitespace to Trim or to Clean,
'                             and it occurs in this data
'     3. strip control chars  Unicode category Cc: U+0000-U+001F, U+007F-U+009F.
'                             REMOVED, not replaced with a space - Text.Clean
'                             deletes them, so "A<tab>B" is "AB" on both sides
'     4. trim
'     5. uppercase
'     6. collapse internal runs of whitespace to a single space
'
'   Steps 2 and 6 are the ones that matter and the ones a naive port drops.
'   Text.Clean and Text.Trim alone handle neither, and both occur here.
'
'   PUNCTUATION IS PRESERVED. Normalisation does not strip dots, commas or
'   hyphens - "CO.,LTD." survives step 6 intact and is dealt with by the
'   customer map instead. (The keep-list punctuation rule belongs to
'   fnPartSig, which is GTF work and is not in this module.)
'==============================================================================

' The separator inside CustEngPlantKey. See THE KEY above before changing it.
Private Const KEY_SEP As String = "|"

Private Const TBL_CUST_MAP As String = "tbl_CustNameMap"
Private Const TBL_ENG_MAP  As String = "tbl_EngFamilyMap"

' Loaded once and held for the life of the run. ClearKeyCaches() resets them;
' RefreshLocal should call it, the same way it calls ClearCaches().
Private m_dCust As Object
Private m_dEng As Object
Private m_bFromTable As Boolean


'==============================================================================
' THE THREE FUNCTIONS
'==============================================================================

'------------------------------------------------------------------------------
' fnNormText. Applied to every key component before anything else.
'------------------------------------------------------------------------------
Public Function NormText(v As Variant) As String

    Dim s As String

    If IsNull(v) Then
        NormText = ""
        Exit Function
    End If

    s = CStr(v)

    ' ChrW, not Chr. Chr(160) resolves through the ANSI codepage and does not
    ' reliably give U+00A0; ChrW always does.
    s = Replace(s, ChrW(160), " ")

    s = StripControl(s)
    s = Trim(s)
    s = UCase(s)
    s = CollapseSpaces(s)

    NormText = s
End Function


'------------------------------------------------------------------------------
' fnCustName. Normalises, then applies the curated customer map.
'
' One live entry: the same airline is spelled two ways across the two systems,
' with roughly 1.74M of plan split between the spellings.
'
' SUSPECTED BUT UNCONFIRMED, and deliberately NOT seeded: "SETNA I0" (letter O
' against a zero?) and "JET MIDWEST. INC". Verify against the source files
' before adding either - a wrong merge is harder to find than a missed one.
'------------------------------------------------------------------------------
Public Function MapCustName(v As Variant) As String

    Dim s As String

    s = NormText(v)
    If Len(s) = 0 Then
        MapCustName = ""
        Exit Function
    End If

    LoadMaps
    If m_dCust.Exists(s) Then MapCustName = m_dCust(s) Else MapCustName = s

End Function


'------------------------------------------------------------------------------
' fnEngFamily. Normalises, then applies the curated engine-family map.
'
' THIS IS A CURATED MAP, NOT A RULE. Do not replace it with "strip everything
' after the first hyphen" - that merges CFM56-5C into CFM56-7B, which are
' genuinely different engines. Anything not in the map passes through
' unchanged, which is how PW1000, PW1100G-JM, PW1500G, PW1900G, PW100,
' PW2000 and GP7000 stay distinct.
'
' WHY IT EXISTS: OMM3 forecasts PW4000-94; ZPT books the receipts as bare
' PW4000, or as PW4000-94/100. Same engine, three spellings, three keys.
'
' THE ONE GENUINELY AMBIGUOUS MERGE IS CFM56. ZPT distinguishes -5B from -7;
' OMM3 does not split them at all. Merging is the only way to compare them,
' but it should be stated explicitly wherever the number is presented.
'------------------------------------------------------------------------------
Public Function MapEngFamily(v As Variant) As String

    Dim s As String

    s = NormText(v)
    If Len(s) = 0 Then
        MapEngFamily = ""
        Exit Function
    End If

    LoadMaps
    If m_dEng.Exists(s) Then MapEngFamily = m_dEng(s) Else MapEngFamily = s

End Function


'------------------------------------------------------------------------------
' The join key itself. Plant is a code ("5610" / "5650") and is normalised
' only for whitespace and case - it is never mapped.
'------------------------------------------------------------------------------
Public Function CustEngPlantKey(vCust As Variant, vEng As Variant, vPlant As Variant) As String
    CustEngPlantKey = MapCustName(vCust) & KEY_SEP & _
                      MapEngFamily(vEng) & KEY_SEP & _
                      NormText(vPlant)
End Function


'------------------------------------------------------------------------------
' True when both maps came from their tables. False means the module fell back
' to the built-in seed - a diagnostic should say so rather than let a
' half-configured database look healthy.
'------------------------------------------------------------------------------
Public Function MapsAreTableDriven() As Boolean
    LoadMaps
    MapsAreTableDriven = m_bFromTable
End Function


Public Sub ClearKeyCaches()
    Set m_dCust = Nothing
    Set m_dEng = Nothing
End Sub


'==============================================================================
' NORMALISATION INTERNALS
'==============================================================================

'------------------------------------------------------------------------------
' Text.Clean: remove Unicode control characters. REMOVED, not replaced with a
' space - see the header. Category Cc is U+0000-U+001F and U+007F-U+009F.
'------------------------------------------------------------------------------
Private Function StripControl(s As String) As String

    Dim i As Long, c As Long
    Dim sOut As String

    For i = 1 To Len(s)
        ' AscW returns a signed Integer, so anything above U+7FFF comes back
        ' negative. None of those are control characters, but the arithmetic
        ' has to be right before the comparison.
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536

        If Not (c < 32 Or (c >= 127 And c <= 159)) Then
            sOut = sOut & Mid$(s, i, 1)
        End If
    Next i

    StripControl = sOut
End Function


'------------------------------------------------------------------------------
' Collapse internal runs of spaces to one. Split/Join dropping the empties,
' which is what Text.Combine(List.Select(Text.Split(...))) does on the M side.
'------------------------------------------------------------------------------
Private Function CollapseSpaces(s As String) As String

    Dim parts() As String, keep() As String
    Dim i As Long, n As Long

    parts = Split(s, " ")
    ReDim keep(0 To UBound(parts))

    n = 0
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then
            keep(n) = parts(i)
            n = n + 1
        End If
    Next i

    If n = 0 Then
        CollapseSpaces = ""
    Else
        ReDim Preserve keep(0 To n - 1)
        CollapseSpaces = Join(keep, " ")
    End If

End Function


'==============================================================================
' MAP LOADING
'==============================================================================
'   Source values in both tables are stored ALREADY NORMALISED, because that
'   is what the lookup is done against. RBDB_SetupKeyMaps() normalises on the
'   way in; a row typed straight into the table by hand must be normalised
'   too, or it will never match. RBDB_TestKeys() checks for that.
'------------------------------------------------------------------------------
Private Sub LoadMaps()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If Not (m_dCust Is Nothing) Then Exit Sub

    Set m_dCust = CreateObject("Scripting.Dictionary")
    m_dCust.CompareMode = 1                      ' text, case-insensitive
    Set m_dEng = CreateObject("Scripting.Dictionary")
    m_dEng.CompareMode = 1

    Set db = CurrentDb()

    If Not KeyTableExists(db, TBL_CUST_MAP) Or Not KeyTableExists(db, TBL_ENG_MAP) Then
        SeedDefaults m_dCust, m_dEng
        m_bFromTable = False
        Exit Sub
    End If

    Set rs = db.OpenRecordset( _
        "SELECT SourceValue, MappedValue FROM " & TBL_CUST_MAP & _
        " WHERE SourceValue Is Not Null AND MappedValue Is Not Null", dbOpenSnapshot)
    Do While Not rs.EOF
        m_dCust(NormText(rs!SourceValue)) = NormText(rs!MappedValue)
        rs.MoveNext
    Loop
    rs.Close

    Set rs = db.OpenRecordset( _
        "SELECT SourceValue, FamilyValue FROM " & TBL_ENG_MAP & _
        " WHERE SourceValue Is Not Null AND FamilyValue Is Not Null", dbOpenSnapshot)
    Do While Not rs.EOF
        m_dEng(NormText(rs!SourceValue)) = NormText(rs!FamilyValue)
        rs.MoveNext
    Loop
    rs.Close

    m_bFromTable = True

End Sub


'------------------------------------------------------------------------------
' The built-in seed. Two jobs: it populates the tables on first setup, and it
' is the fallback if they are missing. Both callers go through here so there
' is only ever one list to edit.
'------------------------------------------------------------------------------
Private Sub SeedDefaults(dCust As Object, dEng As Object)

    dCust("ALL NIPPON AIRWAYS CO.,LTD.") = "ALL NIPPON AIRWAYS CO LTD"

    dEng("PW4000-94") = "PW4000"
    dEng("PW4000-100") = "PW4000"
    dEng("PW4000-112") = "PW4000"
    dEng("PW4000-94/100") = "PW4000"
    dEng("PW4000-100/112") = "PW4000"
    dEng("V2500_A1") = "V2500"
    dEng("V2500_A5/D5") = "V2500"
    dEng("CFM56-5B") = "CFM56"
    dEng("CFM56-7") = "CFM56"
    dEng("F117-PW-100") = "F117"
    dEng("PW150A") = "PW150"

End Sub


' Same shape as TableExists in modAnomalyAlert - kept local because that one
' is Private and a duplicated four-line helper is cheaper than widening its
' scope.
Private Function KeyTableExists(db As DAO.Database, sName As String) As Boolean

    Dim td As DAO.TableDef

    On Error Resume Next
    Set td = db.TableDefs(sName)
    On Error GoTo 0

    KeyTableExists = Not (td Is Nothing)
End Function


'==============================================================================
' SETUP
'==============================================================================
'   Creates the two map tables IF THEY ARE MISSING and seeds them. An existing
'   table is left completely alone - the same convention RBDB_Setup uses for
'   tbl_Config, and for the same reason: a re-run must never discard mappings
'   somebody added by hand.
'
'   RBDB_ReseedKeyMaps() is the deliberate destructive version.
'==============================================================================
Public Function RBDB_SetupKeyMaps()

    Dim db As DAO.Database
    Dim nMade As Long
    Dim s As String

    Set db = CurrentDb()

    If Not KeyTableExists(db, TBL_CUST_MAP) Then
        db.Execute "CREATE TABLE " & TBL_CUST_MAP & " (" & _
                   "SourceValue TEXT(255), MappedValue TEXT(255), Notes TEXT(255))", dbFailOnError
        nMade = nMade + 1
    End If

    If Not KeyTableExists(db, TBL_ENG_MAP) Then
        db.Execute "CREATE TABLE " & TBL_ENG_MAP & " (" & _
                   "SourceValue TEXT(255), FamilyValue TEXT(255), Notes TEXT(255))", dbFailOnError
        nMade = nMade + 1
    End If

    ' Each table is seeded on its own. Testing them together would leave one
    ' empty forever if the other already had rows.
    FillMapTables db

    ClearKeyCaches

    s = "Key maps ready." & vbCrLf & vbCrLf & _
        TBL_CUST_MAP & ": " & Format(DCount("*", TBL_CUST_MAP), "#,##0") & " row(s)" & vbCrLf & _
        TBL_ENG_MAP & ": " & Format(DCount("*", TBL_ENG_MAP), "#,##0") & " row(s)" & vbCrLf & vbCrLf & _
        IIf(nMade = 0, "Both tables already existed - nothing was overwritten.", _
                       nMade & " table(s) created and seeded.") & vbCrLf & vbCrLf & _
        "Next: RBDB_TestKeys(), then RBDB_ExportMaps() and paste the output" & vbCrLf & _
        "into the .pbix so both sides carry the same mappings."

    Debug.Print s
    MsgBox s, vbInformation, "Key maps"

End Function


'------------------------------------------------------------------------------
' Discard both tables and rebuild them from the built-in seed. Destructive:
' anything added by hand is lost. Kept separate from RBDB_SetupKeyMaps for
' exactly that reason.
'------------------------------------------------------------------------------
Public Function RBDB_ReseedKeyMaps()

    Dim db As DAO.Database

    If MsgBox("This DISCARDS every row in " & TBL_CUST_MAP & " and " & TBL_ENG_MAP & _
              ", including mappings added by hand, and rebuilds both from the " & _
              "built-in seed." & vbCrLf & vbCrLf & "Continue?", _
              vbExclamation + vbYesNo + vbDefaultButton2, "Reseed key maps") <> vbYes Then
        Exit Function
    End If

    Set db = CurrentDb()

    If KeyTableExists(db, TBL_CUST_MAP) Then db.Execute "DELETE FROM " & TBL_CUST_MAP, dbFailOnError
    If KeyTableExists(db, TBL_ENG_MAP) Then db.Execute "DELETE FROM " & TBL_ENG_MAP, dbFailOnError

    FillMapTables db
    ClearKeyCaches

    MsgBox "Reseeded from the built-in defaults." & vbCrLf & vbCrLf & _
           "Run RBDB_ExportMaps() and repaste into the .pbix.", vbInformation, "Key maps"

End Function


'------------------------------------------------------------------------------
' Write the built-in seed into the two tables. Source values are normalised on
' the way in so the stored form is the form the lookup compares against.
'------------------------------------------------------------------------------
Private Sub FillMapTables(db As DAO.Database)

    Dim dCust As Object, dEng As Object
    Dim k As Variant

    Set dCust = CreateObject("Scripting.Dictionary")
    dCust.CompareMode = 1
    Set dEng = CreateObject("Scripting.Dictionary")
    dEng.CompareMode = 1

    SeedDefaults dCust, dEng

    If DCount("*", TBL_CUST_MAP) > 0 Then
        dCust.RemoveAll
    End If

    If DCount("*", TBL_ENG_MAP) > 0 Then
        dEng.RemoveAll
    End If

    For Each k In dCust.Keys
        db.Execute "INSERT INTO " & TBL_CUST_MAP & " (SourceValue, MappedValue, Notes) VALUES ('" & _
                   SqlLit(NormText(k)) & "', '" & SqlLit(NormText(dCust(k))) & "', '" & _
                   SqlLit("same airline, two spellings across the two systems") & "')", dbFailOnError
    Next k

    For Each k In dEng.Keys
        db.Execute "INSERT INTO " & TBL_ENG_MAP & " (SourceValue, FamilyValue, Notes) VALUES ('" & _
                   SqlLit(NormText(k)) & "', '" & SqlLit(NormText(dEng(k))) & "', '" & _
                   SqlLit("curated - OMM3 and ZPT spell this engine differently") & "')", dbFailOnError
    Next k

End Sub


Private Function SqlLit(s As String) As String
    SqlLit = Replace(s, "'", "''")
End Function


'==============================================================================
' TEST
'==============================================================================
'   Feeds known inputs and asserts the outputs. Every case here is a rule the
'   header states, and several are rules a plausible-looking port gets wrong.
'   Run it after ANY edit to either map table.
'
'   Safe to run any time. Reads only.
'==============================================================================
Public Function RBDB_TestKeys()

    Dim s As String
    Dim nBad As Long, nRun As Long
    Dim sNbsp As String

    ClearKeyCaches
    LoadMaps

    s = "KEY TESTS - " & Format(Now(), "dd mmm yyyy hh:nn") & vbCrLf & String(58, "-") & vbCrLf
    s = s & "maps loaded from " & IIf(m_bFromTable, "TABLES", "BUILT-IN SEED (tables missing)") & _
            "   cust=" & m_dCust.Count & "  eng=" & m_dEng.Count & vbCrLf & vbCrLf

    '--- normalisation --------------------------------------------------------
    s = s & "fnNormText:" & vbCrLf

    s = s & Chk("null is empty", NormText(Null), "", nBad, nRun)
    s = s & Chk("trim and uppercase", NormText("  boeing co  "), "BOEING CO", nBad, nRun)

    ' The two steps Text.Clean and Text.Trim do NOT handle on their own.
    sNbsp = "ALL" & ChrW(160) & "NIPPON"
    s = s & Chk("NBSP becomes a space", NormText(sNbsp), "ALL NIPPON", nBad, nRun)
    s = s & Chk("double space collapses", NormText("DELTA   AIR  LINES"), "DELTA AIR LINES", nBad, nRun)

    ' Control characters are DELETED, not turned into spaces - so a tab
    ' between two words joins them. This is what Text.Clean does; if the M
    ' side ever disagrees, this is the assertion that catches it.
    s = s & Chk("control chars removed", NormText("A" & vbTab & "B"), "AB", nBad, nRun)
    s = s & Chk("punctuation survives", NormText(" co.,ltd. "), "CO.,LTD.", nBad, nRun)

    '--- customer map ---------------------------------------------------------
    s = s & vbCrLf & "fnCustName:" & vbCrLf

    s = s & Chk("ANA spelling merged", MapCustName("ALL NIPPON AIRWAYS CO.,LTD."), _
                "ALL NIPPON AIRWAYS CO LTD", nBad, nRun)
    s = s & Chk("ANA already mapped is stable", MapCustName("ALL NIPPON AIRWAYS CO LTD"), _
                "ALL NIPPON AIRWAYS CO LTD", nBad, nRun)
    s = s & Chk("unmapped passes through", MapCustName("some airline inc"), _
                "SOME AIRLINE INC", nBad, nRun)

    '--- engine family map ----------------------------------------------------
    s = s & vbCrLf & "fnEngFamily:" & vbCrLf

    s = s & Chk("PW4000-94", MapEngFamily("PW4000-94"), "PW4000", nBad, nRun)
    s = s & Chk("PW4000-94/100", MapEngFamily("PW4000-94/100"), "PW4000", nBad, nRun)
    s = s & Chk("PW4000-100/112", MapEngFamily("PW4000-100/112"), "PW4000", nBad, nRun)
    s = s & Chk("V2500_A5/D5", MapEngFamily("V2500_A5/D5"), "V2500", nBad, nRun)
    s = s & Chk("CFM56-5B", MapEngFamily("CFM56-5B"), "CFM56", nBad, nRun)
    s = s & Chk("F117-PW-100", MapEngFamily("F117-PW-100"), "F117", nBad, nRun)
    s = s & Chk("PW150A", MapEngFamily("PW150A"), "PW150", nBad, nRun)

    ' The map is curated, NOT "strip after the first hyphen". These four must
    ' survive untouched or genuinely different engines have been merged.
    s = s & Chk("PW1100G-JM untouched", MapEngFamily("PW1100G-JM"), "PW1100G-JM", nBad, nRun)
    s = s & Chk("PW1500G untouched", MapEngFamily("PW1500G"), "PW1500G", nBad, nRun)
    s = s & Chk("PW2000 untouched", MapEngFamily("PW2000"), "PW2000", nBad, nRun)
    s = s & Chk("GP7000 untouched", MapEngFamily("GP7000"), "GP7000", nBad, nRun)

    '--- the key --------------------------------------------------------------
    s = s & vbCrLf & "CustEngPlantKey:" & vbCrLf

    s = s & Chk("composed key", _
                CustEngPlantKey("  all nippon airways co.,ltd. ", "PW4000-94", "5610"), _
                "ALL NIPPON AIRWAYS CO LTD" & KEY_SEP & "PW4000" & KEY_SEP & "5610", nBad, nRun)

    ' The whole point of the module: OMM3 spells it one way, ZPT another, and
    ' they have to land on the same key.
    s = s & Chk("OMM3 and ZPT agree", _
                IIf(CustEngPlantKey("ALL NIPPON AIRWAYS CO.,LTD.", "PW4000-94", "5610") = _
                    CustEngPlantKey("ALL NIPPON AIRWAYS CO LTD", "PW4000-94/100", "5610"), _
                    "same", "DIFFERENT"), "same", nBad, nRun)

    s = s & Chk("null engine gives empty family", MapEngFamily(Null), "", nBad, nRun)

    '--- stored rows must already be normalised -------------------------------
    s = s & vbCrLf & "Table hygiene:" & vbCrLf
    s = s & ChkStoredNormalised(nBad)

    s = s & String(58, "-") & vbCrLf
    If nBad = 0 Then
        s = s & "All " & nRun & " checks passed." & vbCrLf & _
                "Power Query must produce these same outputs. Run RBDB_ExportMaps()" & vbCrLf & _
                "if either map table has changed since the .pbix was last updated."
    Else
        s = s & nBad & " of " & nRun & " check(s) FAILED. Do not build on these keys" & vbCrLf & _
                "until they pass - every downstream figure depends on them."
    End If

    Debug.Print s
    MsgBox s, IIf(nBad = 0, vbInformation, vbExclamation), "Key tests"

End Function


'------------------------------------------------------------------------------
' One assertion.
'------------------------------------------------------------------------------
Private Function Chk(sWhat As String, sGot As String, sWant As String, _
                     ByRef nBad As Long, ByRef nRun As Long) As String

    Dim sPad As String

    nRun = nRun + 1

    sPad = sWhat & String(30, " ")
    sPad = Left(sPad, 30)

    If sGot = sWant Then
        Chk = "  [ok] " & sPad & " " & sGot & vbCrLf
    Else
        Chk = "  [X]  " & sPad & " got '" & sGot & "'  want '" & sWant & "'" & vbCrLf
        nBad = nBad + 1
    End If

End Function


'------------------------------------------------------------------------------
' A source value typed into a map table by hand that is not already normalised
' will never match anything, and nothing else would ever say so.
'------------------------------------------------------------------------------
Private Function ChkStoredNormalised(ByRef nBad As Long) As String

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim s As String
    Dim nOff As Long

    Set db = CurrentDb()

    If Not KeyTableExists(db, TBL_CUST_MAP) Or Not KeyTableExists(db, TBL_ENG_MAP) Then
        ChkStoredNormalised = "  [X]  map tables missing - running on the built-in seed." & vbCrLf & _
                              "       Run RBDB_SetupKeyMaps()." & vbCrLf
        nBad = nBad + 1
        Exit Function
    End If

    Set rs = db.OpenRecordset("SELECT SourceValue FROM " & TBL_CUST_MAP & _
                              " WHERE SourceValue Is Not Null", dbOpenSnapshot)
    Do While Not rs.EOF
        If CStr(rs!SourceValue) <> NormText(rs!SourceValue) Then
            s = s & "  [X]  " & TBL_CUST_MAP & " row not normalised: '" & _
                    CStr(rs!SourceValue) & "'" & vbCrLf
            nOff = nOff + 1
        End If
        rs.MoveNext
    Loop
    rs.Close

    Set rs = db.OpenRecordset("SELECT SourceValue FROM " & TBL_ENG_MAP & _
                              " WHERE SourceValue Is Not Null", dbOpenSnapshot)
    Do While Not rs.EOF
        If CStr(rs!SourceValue) <> NormText(rs!SourceValue) Then
            s = s & "  [X]  " & TBL_ENG_MAP & " row not normalised: '" & _
                    CStr(rs!SourceValue) & "'" & vbCrLf
            nOff = nOff + 1
        End If
        rs.MoveNext
    Loop
    rs.Close

    If nOff = 0 Then
        ChkStoredNormalised = "  [ok] every stored SourceValue is already normalised" & vbCrLf
    Else
        ChkStoredNormalised = s & "       A source value must be stored in its normalised form" & vbCrLf & _
                                  "       or the lookup can never match it." & vbCrLf
        nBad = nBad + 1
    End If

End Function


'==============================================================================
' EXPORT TO POWER QUERY
'==============================================================================
'   Writes both maps as an M literal next to the .accdb. Paste the result into
'   the .pbix so fnCustName and fnEngFamily read the same mappings this module
'   does.
'
'   REGENERATE, DO NOT HAND-EDIT the M side. That is the whole mechanism by
'   which the two implementations stay in step: a mapping added here and not
'   exported shows up as a diff the next time somebody runs this, rather than
'   as a silent divergence six months later.
'==============================================================================
Public Function RBDB_ExportMaps()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sPath As String, s As String
    Dim iFile As Integer
    Dim bFirst As Boolean

    LoadMaps

    Set db = CurrentDb()
    sPath = CurrentProject.Path & "\RBDB_KeyMaps.m"

    s = "// Generated by modKeys.RBDB_ExportMaps() - " & Format(Now(), "dd mmm yyyy hh:nn") & vbCrLf
    s = s & "// Source of truth is the Access database. DO NOT HAND-EDIT: add the" & vbCrLf
    s = s & "// mapping to tbl_CustNameMap / tbl_EngFamilyMap and re-export." & vbCrLf
    s = s & "// Source values are already normalised - fnNormText must be applied to" & vbCrLf
    s = s & "// the INPUT before the lookup, not to these." & vbCrLf & vbCrLf

    '--- customer map ---------------------------------------------------------
    s = s & "CustNameMap = #table(" & vbCrLf
    s = s & "    {""SourceValue"", ""MappedValue""}," & vbCrLf
    s = s & "    {" & vbCrLf

    Set rs = db.OpenRecordset("SELECT SourceValue, MappedValue FROM " & TBL_CUST_MAP & _
                              " WHERE SourceValue Is Not Null ORDER BY SourceValue", dbOpenSnapshot)
    bFirst = True
    Do While Not rs.EOF
        If Not bFirst Then s = s & "," & vbCrLf
        s = s & "        {" & MLit(CStr(rs!SourceValue)) & ", " & MLit(CStr(rs!MappedValue)) & "}"
        bFirst = False
        rs.MoveNext
    Loop
    rs.Close
    s = s & vbCrLf & "    }" & vbCrLf & ")," & vbCrLf & vbCrLf

    '--- engine family map ----------------------------------------------------
    s = s & "EngFamilyMap = #table(" & vbCrLf
    s = s & "    {""SourceValue"", ""FamilyValue""}," & vbCrLf
    s = s & "    {" & vbCrLf

    Set rs = db.OpenRecordset("SELECT SourceValue, FamilyValue FROM " & TBL_ENG_MAP & _
                              " WHERE SourceValue Is Not Null ORDER BY SourceValue", dbOpenSnapshot)
    bFirst = True
    Do While Not rs.EOF
        If Not bFirst Then s = s & "," & vbCrLf
        s = s & "        {" & MLit(CStr(rs!SourceValue)) & ", " & MLit(CStr(rs!FamilyValue)) & "}"
        bFirst = False
        rs.MoveNext
    Loop
    rs.Close
    s = s & vbCrLf & "    }" & vbCrLf & ")" & vbCrLf

    iFile = FreeFile
    Open sPath For Output As #iFile
    Print #iFile, s
    Close #iFile

    Debug.Print s

    MsgBox "Maps exported to:" & vbCrLf & vbCrLf & sPath & vbCrLf & vbCrLf & _
           "Paste into the .pbix so fnCustName and fnEngFamily read the same" & vbCrLf & _
           "mappings this module does. Also printed to the Immediate window.", _
           vbInformation, "Key maps exported"

End Function


'------------------------------------------------------------------------------
' One M string literal. M escapes a double quote by doubling it, same as SQL
' does with a single quote.
'------------------------------------------------------------------------------
Private Function MLit(s As String) As String
    MLit = """" & Replace(s, """", """""") & """"
End Function
