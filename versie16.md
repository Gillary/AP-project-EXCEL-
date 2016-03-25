# AP-project-EXCEL-
Private Sub but_berekeningen_beheerder_Click()
    Berekeningenscherm.Hide
    Beheerder.Show
End Sub

Private Sub but_berekeningen_bruto_Click()
    Worksheets("opgeslagen pensioenen").Range("AW2").Value = 1
    but_berekeningen_bruto.Visible = False
    but_berekeningen_netto.Visible = True
    lbl_berekeningen_bruto.Visible = True
    lbl_berekeningen_netto.Visible = False
    Call Toonresultaat
End Sub

Private Sub but_berekeningen_netto_Click()
    Worksheets("opgeslagen pensioenen").Range("AW2").Value = 0
    but_berekeningen_bruto.Visible = True
    but_berekeningen_netto.Visible = False
    lbl_berekeningen_bruto.Visible = False
    lbl_berekeningen_netto.Visible = True
    Call Toonresultaat
End Sub

Private Sub check_berekeningen_hoog_Click()
        If check_berekeningen_hoog.Value = False Then
        lbl_berekeningen_hoog_1.Enabled = False
        lbl_berekeningen_hoog_2.Enabled = False
        option_berekeningen_hoog_verschil.Enabled = False
        option_berekeningen_hoog_verhouding.Enabled = False
        lbl_berekeningen_hoog_jaar.Enabled = False
        lbl_berekeningen_hoog_maanden.Enabled = False
        dropdown_berekeningen_hoog_jaar.Enabled = False
        dropdown_berekeningen_hoog_maanden.Enabled = False
        textbox_berekeningen_hoog_verschil.Enabled = False
        lbl_berekeningen_hoog_3.Enabled = False
        dropdown_berekeningen_hoog_verhouding_laag.Enabled = False
        but_berekeningen_hoog_hulp.Enabled = False
        Verhouding_Hoog_Laag_100.Enabled = False
    Else
        lbl_berekeningen_hoog_1.Enabled = True
        lbl_berekeningen_hoog_2.Enabled = True
        option_berekeningen_hoog_verschil.Enabled = True
        option_berekeningen_hoog_verhouding.Enabled = True
        lbl_berekeningen_hoog_jaar.Enabled = True
        lbl_berekeningen_hoog_maanden.Enabled = True
        dropdown_berekeningen_hoog_jaar.Enabled = True
        dropdown_berekeningen_hoog_maanden.Enabled = True
        textbox_berekeningen_hoog_verschil.Enabled = True
        lbl_berekeningen_hoog_3.Enabled = True
        dropdown_berekeningen_hoog_verhouding_laag.Enabled = True
        but_berekeningen_hoog_hulp.Enabled = True
        Verhouding_Hoog_Laag_100.Enabled = True
    End If
End Sub

Private Sub check_berekeningen_laag_Click()
    If check_berekeningen_laag.Value = False Then
        lbl_berekeningen_laag_1.Enabled = False
        lbl_berekeningen_laag_2.Enabled = False
        option_berekeningen_laag_verschil.Enabled = False
        option_berekeningen_laag_verhouding.Enabled = False
        lbl_berekeningen_laag_jaar.Enabled = False
        lbl_berekeningen_laag_maanden.Enabled = False
        dropdown_berekeningen_laag_jaar.Enabled = False
        dropdown_berekeningen_laag_maanden.Enabled = False
        textbox_berekeningen_laag_verschil.Enabled = False
        lbl_berekeningen_laag_3.Enabled = False
        dropdown_berekeningen_laag_verhouding_hoog.Enabled = False
        Verhouding_Laag_Hoog_100.Enabled = False
    Else
        lbl_berekeningen_laag_1.Enabled = True
        lbl_berekeningen_laag_2.Enabled = True
        option_berekeningen_laag_verschil.Enabled = True
        option_berekeningen_laag_verhouding.Enabled = True
        lbl_berekeningen_laag_jaar.Enabled = True
        lbl_berekeningen_laag_maanden.Enabled = True
        dropdown_berekeningen_laag_jaar.Enabled = True
        dropdown_berekeningen_laag_maanden.Enabled = True
        textbox_berekeningen_laag_verschil.Enabled = True
        lbl_berekeningen_laag_3.Enabled = True
        dropdown_berekeningen_laag_verhouding_hoog.Enabled = True
        Verhouding_Laag_Hoog_100.Enabled = True
    End If
        
End Sub

Private Sub check_berekeningen_oppp_Click()
    If check_berekeningen_oppp.Value = False Then
        lbl_berekeningen_oppp.Enabled = False
        option_berekeningen_oppp_procenten.Enabled = False
        option_berekeningen_oppp_verhouding.Enabled = False
        dropdown_berekeningen_oppp_percentage.Enabled = False
        dropdown_berekeningen_oppp_verhouding_pp.Enabled = False
        lbl_berekeningen_oppp_2.Enabled = False
        Op_PP_100.Enabled = False
    Else
        lbl_berekeningen_oppp.Enabled = True
        option_berekeningen_oppp_procenten.Enabled = True
        option_berekeningen_oppp_verhouding.Enabled = True
        dropdown_berekeningen_oppp_percentage.Enabled = True
        dropdown_berekeningen_oppp_verhouding_pp.Enabled = True
        lbl_berekeningen_oppp_2.Enabled = True
        Op_PP_100.Enabled = True
    End If

End Sub

Private Sub check_berekeningen_ppop_Click()
    If check_berekeningen_ppop.Value = False Then
        lbl_berekeningen_ppop.Enabled = False
        dropdown_berekeningen_ppop_percentage.Enabled = False
        lbl_berekeningen_ppop_2.Enabled = False
        Op_PP_100.Enabled = False
    Else
        lbl_berekeningen_ppop.Enabled = True
        dropdown_berekeningen_ppop_percentage.Enabled = True
        lbl_berekeningen_ppop_2.Enabled = True
        Op_PP_100.Enabled = True
    End If
End Sub

Private Sub check_berekeningen_uitstellen_Click()
    If check_berekeningen_uitstellen.Value = False Then
        lbl_berekeningen_uitstellen.Enabled = False
        dropdown_berekeningen_uitstellen_jaar.Enabled = False
        dropdown_berekeningen_uitstellen_maand.Enabled = False
        lbl_berekeningen_uitstellen_jaar.Enabled = False
        lbl_berekeningen_uitstellen_maanden.Enabled = False
    Else
        lbl_berekeningen_uitstellen.Enabled = True
        dropdown_berekeningen_uitstellen_jaar.Enabled = True
        dropdown_berekeningen_uitstellen_maand.Enabled = True
        lbl_berekeningen_uitstellen_jaar.Enabled = True
        lbl_berekeningen_uitstellen_maanden.Enabled = True
    End If
    
End Sub

Private Sub check_berekeningen_vervroegen_Click()
    If check_berekeningen_vervroegen.Value = False Then
        lbl_berekeningen_vervroegen.Enabled = False
        check_berekeningen_laag_aowopvullen.Enabled = False
        lbl_berekeningen_vervroegen_jaar.Enabled = False
        lbl_berekeningen_vervroegen_maanden.Enabled = False
        dropdown_berekeningen_vervroegen_jaar.Enabled = False
        dropdown_berekeningen_vervroegen_maanden.Enabled = False
    Else
        check_berekeningen_vervroegen.Enabled = True
        lbl_berekeningen_vervroegen.Enabled = True
        check_berekeningen_laag_aowopvullen.Enabled = True
        lbl_berekeningen_vervroegen_jaar.Enabled = True
        lbl_berekeningen_vervroegen_maanden.Enabled = True
        dropdown_berekeningen_vervroegen_jaar.Enabled = True
        dropdown_berekeningen_vervroegen_maanden.Enabled = True
    End If
End Sub


'selecteer leeftijd uitstellen
Private Sub dropdown_berekeningen_uitstellen_maand_Change()
    check_berekeningen_uitstellen.Value = True
End Sub
'selecteer leeftijd uitstellen
Private Sub dropdown_berekeningen_uitstellen_jaar_Change()
    check_berekeningen_uitstellen.Value = True
End Sub
'selecteer leeftijd vervroegen
Private Sub dropdown_berekeningen_vervroegen_maand_change()
    check_berekeningen_vervroegen.Value = True
End Sub
'selecteer leeftijd vervroegen
Private Sub dropdown_berekeningen_vervroegen_jaar_Change()
    check_berekeningen_vervroegen.Value = True
End Sub
'selecteer ppop
Private Sub dropdown_berekeningen_ppop_percentage_Change()
    check_berekeningen_ppop.Value = True
End Sub



'selecteer oppp
Private Sub option_berekeningen_oppp_procenten_Click()
    check_berekeningen_oppp.Value = True
End Sub
'selecteer oppp
Private Sub option_berekeningen_oppp_verhouding_Click()
    check_berekeningen_oppp.Value = True
End Sub
'selecteer oppp
Private Sub dropdown_berekeningen_oppp_percentage_Change()
    check_berekeningen_oppp.Value = True
    option_berekeningen_oppp_procenten.Value = True
End Sub
'selecteer oppp
Private Sub dropdown_berekeningen_oppp_verhouding_pp_Change()
    check_berekeningen_oppp.Value = True
    option_berekeningen_oppp_verhouding.Value = True
End Sub
'selecteer hooglaag
Private Sub dropdown_berekeningen_hoog_jaar_Change()
    check_berekeningen_hoog.Value = True
End Sub
'selecteer hooglaag
Private Sub dropdown_berekeningen_hoog_maanden_Change()
    check_berekeningen_hoog.Value = True
End Sub
'selecteer hooglaag
Private Sub textbox_berekeningen_hoog_verschil_Change()
    check_berekeningen_hoog.Value = True
    option_berekeningen_hoog_verschil.Value = True
End Sub
'selecteer hooglaag
Private Sub dropdown_berekeningen_hoog_verhouding_laag_Change()
    check_berekeningen_hoog.Value = True
    option_berekeningen_hoog_verhouding.Value = True
End Sub
'selecteer hooglaag
Private Sub option_berekeningen_hoog_verschil_Click()
    check_berekeningen_hoog.Value = True
End Sub
'selecteer hooglaag
Private Sub option_berekeningen_hoog_verhouding_Click()
    check_berekeningen_hoog.Value = True
End Sub
'selecteer laaghoog
Private Sub dropdown_berekeningen_laag_jaar_Change()
    check_berekeningen_laag.Value = True
End Sub
'selecteer laaghoog
Private Sub dropdown_berekeningen_laag_maanden_Change()
    check_berekeningen_laag.Value = True
End Sub
'selecteer laaghoog
Private Sub textbox_berekeningen_laag_verschil_Change()
    check_berekeningen_laag.Value = True
    option_berekeningen_laag_verschil.Value = True
End Sub
'selecteer laaghoog
Private Sub dropdown_berekeningen_laag_verhouding_hoog_Change()
    check_berekeningen_laag.Value = True
    option_berekeningen_laag_verhouding.Value = True
End Sub
'selecteer laaghoog
Private Sub option_berekeningen_laag_verschil_Click()
    check_berekeningen_laag.Value = True
End Sub
'selecteer laaghoog
Private Sub option_berekeningen_laag_verhouding_Click()
    check_berekeningen_laag.Value = True
End Sub
'selecteer laaghoog
Private Sub check_berekeningen_laag_aowopvullen_Click()
    check_berekeningen_vervroegen.Value = True
End Sub
Private Sub UserForm_Activate()
    If Worksheets("opgeslagen pensioenen").Cells(2, 47) = 1 Then
        but_berekeningen_beheerder.Visible = True
    Else
        but_berekeningen_beheerder.Visible = False
    End If
        
    'Kijken of er bruto of netto wordt weergegeven
    If Worksheets("opgeslagen pensioenen").Range("AW2").Value = 0 Or Worksheets("opgeslagen pensioenen").Range("AW2").Value = "" Then
        but_berekeningen_bruto.Visible = False
        but_berekeningen_netto.Visible = True
        lbl_berekeningen_bruto.Visible = True
        lbl_berekeningen_netto.Visible = False
    ElseIf Worksheets("opgeslagen pensioenen").Range("AW2").Value = 1 Then
        but_berekeningen_bruto.Visible = True
        but_berekeningen_netto.Visible = False
        lbl_berekeningen_bruto.Visible = False
        lbl_berekeningen_netto.Visible = True
    End If

    'Bij activeren sheet de juiste flexibiliseringsmogelijkheden tonen.
    If Worksheets("Opgeslagen pensioenen").Cells(2, 5).Value = "1" Then
        check_berekeningen_vervroegen.Visible = True
        lbl_berekeningen_vervroegen.Visible = True
        lbl_berekeningen_vervroegen_jaar.Visible = True
        check_berekeningen_laag_aowopvullen.Visible = True
        lbl_berekeningen_vervroegen_maanden.Visible = True
        dropdown_berekeningen_vervroegen_jaar.Visible = True
        dropdown_berekeningen_vervroegen_maanden.Visible = True
        but_berekeningen_vervroegen_hulp.Visible = True
    Else
        check_berekeningen_vervroegen.Visible = False
        lbl_berekeningen_vervroegen.Visible = False
        lbl_berekeningen_vervroegen_jaar.Visible = False
        check_berekeningen_laag_aowopvullen.Visible = False
        lbl_berekeningen_vervroegen_maanden.Visible = False
        dropdown_berekeningen_vervroegen_jaar.Visible = False
        dropdown_berekeningen_vervroegen_maanden.Visible = False
        but_berekeningen_vervroegen_hulp.Visible = False
    End If
    
    If Worksheets("Opgeslagen pensioenen").Cells(2, 6).Value = "1" Then
        check_berekeningen_uitstellen.Visible = True
        lbl_berekeningen_uitstellen.Visible = True
        dropdown_berekeningen_uitstellen_jaar.Visible = True
        dropdown_berekeningen_uitstellen_maand.Visible = True
        lbl_berekeningen_uitstellen_jaar.Visible = True
        lbl_berekeningen_uitstellen_maanden.Visible = True
        but_berekeningen_uitstellen_hulp.Visible = True
    Else
        check_berekeningen_uitstellen.Visible = False
        lbl_berekeningen_uitstellen.Visible = False
        dropdown_berekeningen_uitstellen_jaar.Visible = False
        dropdown_berekeningen_uitstellen_maand.Visible = False
        lbl_berekeningen_uitstellen_jaar.Visible = False
        lbl_berekeningen_uitstellen_maanden.Visible = False
        but_berekeningen_uitstellen_hulp.Visible = False
    End If
    
    If Worksheets("Opgeslagen pensioenen").Cells(2, 7).Value = "1" Then
        check_berekeningen_oppp.Visible = True
        lbl_berekeningen_oppp.Visible = True
        option_berekeningen_oppp_procenten.Visible = True
        option_berekeningen_oppp_verhouding.Visible = True
        dropdown_berekeningen_oppp_percentage.Visible = True
        dropdown_berekeningen_oppp_verhouding_pp.Visible = True
        lbl_berekeningen_oppp_2.Visible = True
        but_berekeningen_oppp_hulp.Visible = True
        Op_PP_100.Visible = True
        Frame_oppp.Visible = True
    Else
        check_berekeningen_oppp.Visible = False
        lbl_berekeningen_oppp.Visible = False
        option_berekeningen_oppp_procenten.Visible = False
        option_berekeningen_oppp_verhouding.Visible = False
        dropdown_berekeningen_oppp_percentage.Visible = False
        dropdown_berekeningen_oppp_verhouding_pp.Visible = False
        lbl_berekeningen_oppp_2.Visible = False
        but_berekeningen_oppp_hulp.Visible = False
        Op_PP_100.Visible = False
        Frame_oppp.Visible = False
    End If
    
    If Worksheets("Opgeslagen pensioenen").Cells(2, 8).Value = "1" Then
        check_berekeningen_ppop.Visible = True
        lbl_berekeningen_ppop.Visible = True
        dropdown_berekeningen_ppop_percentage.Visible = True
        lbl_berekeningen_ppop_2.Visible = True
        but_berekeningen_ppop_hulp.Visible = True
    Else
        check_berekeningen_ppop.Visible = False
        lbl_berekeningen_ppop.Visible = False
        dropdown_berekeningen_ppop_percentage.Visible = False
        lbl_berekeningen_ppop_2.Visible = False
        but_berekeningen_ppop_hulp.Visible = False
    End If
    
    If Worksheets("Opgeslagen pensioenen").Cells(2, 9).Value = "1" Then
        check_berekeningen_hoog.Visible = True
        lbl_berekeningen_hoog_1.Visible = True
        lbl_berekeningen_hoog_2.Visible = True
        option_berekeningen_hoog_verschil.Visible = True
        option_berekeningen_hoog_verhouding.Visible = True
        lbl_berekeningen_hoog_jaar.Visible = True
        lbl_berekeningen_hoog_maanden.Visible = True
        dropdown_berekeningen_hoog_jaar.Visible = True
        dropdown_berekeningen_hoog_maanden.Visible = True
        textbox_berekeningen_hoog_verschil.Visible = True
        lbl_berekeningen_hoog_3.Visible = True
        dropdown_berekeningen_hoog_verhouding_laag.Visible = True
        but_berekeningen_hoog_hulp.Visible = True
        Verhouding_Hoog_Laag_100.Visible = True
        Frame_hoog.Visible = True
    Else
        check_berekeningen_hoog.Visible = False
        lbl_berekeningen_hoog_1.Visible = False
        lbl_berekeningen_hoog_2.Visible = False
        option_berekeningen_hoog_verschil.Visible = False
        option_berekeningen_hoog_verhouding.Visible = False
        lbl_berekeningen_hoog_jaar.Visible = False
        lbl_berekeningen_hoog_maanden.Visible = False
        dropdown_berekeningen_hoog_jaar.Visible = False
        dropdown_berekeningen_hoog_maanden.Visible = False
        textbox_berekeningen_hoog_verschil.Visible = False
        lbl_berekeningen_hoog_3.Visible = False
        dropdown_berekeningen_hoog_verhouding_laag.Visible = False
        but_berekeningen_hoog_hulp.Visible = False
        Verhouding_Hoog_Laag_100.Visible = False
        Frame_hoog.Visible = False
    End If
    
    If Worksheets("Opgeslagen pensioenen").Cells(2, 10).Value = "1" Then
        check_berekeningen_laag.Visible = True
        lbl_berekeningen_laag_1.Visible = True
        lbl_berekeningen_laag_2.Visible = True
        option_berekeningen_laag_verschil.Visible = True
        option_berekeningen_laag_verhouding.Visible = True
        lbl_berekeningen_laag_jaar.Visible = True
        lbl_berekeningen_laag_maanden.Visible = True
        dropdown_berekeningen_laag_jaar.Visible = True
        dropdown_berekeningen_laag_maanden.Visible = True
        textbox_berekeningen_laag_verschil.Visible = True
        lbl_berekeningen_laag_3.Visible = True
        dropdown_berekeningen_laag_verhouding_hoog.Visible = True
        but_berekeningen_laag_hulp.Visible = True
        Verhouding_Laag_Hoog_100.Visible = True
        Frame_laag.Visible = True
    Else
        check_berekeningen_laag.Visible = False
        lbl_berekeningen_laag_1.Visible = False
        lbl_berekeningen_laag_2.Visible = False
        option_berekeningen_laag_verschil.Visible = False
        option_berekeningen_laag_verhouding.Visible = False
        lbl_berekeningen_laag_jaar.Visible = False
        lbl_berekeningen_laag_maanden.Visible = False
        dropdown_berekeningen_laag_jaar.Visible = False
        dropdown_berekeningen_laag_maanden.Visible = False
        textbox_berekeningen_laag_verschil.Visible = False
        lbl_berekeningen_laag_3.Visible = False
        dropdown_berekeningen_laag_verhouding_hoog.Visible = False
        but_berekeningen_laag_hulp.Visible = False
        Verhouding_Laag_Hoog_100.Visible = False
        Frame_laag.Visible = True
    End If
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'Checkboxen uitzetten als niet visible
    
    If check_berekeningen_vervroegen.Visible = False Then
        check_berekeningen_vervroegen.Value = False
        Worksheets("Opgeslagen pensioenen").Range("K2") = 0
    End If
    If check_berekeningen_uitstellen.Visible = False Then
        check_berekeningen_uitstellen.Value = False
        Worksheets("Opgeslagen pensioenen").Range("N2") = 0
    End If
    If check_berekeningen_oppp.Visible = False Then
        check_berekeningen_oppp.Value = False
        Worksheets("Opgeslagen pensioenen").Range("Q2") = 0
    End If
    If check_berekeningen_ppop.Visible = False Then
        check_berekeningen_ppop.Value = False
        Worksheets("Opgeslagen pensioenen").Range("V2") = 0
    End If
    If check_berekeningen_laag.Visible = False Then
        check_berekeningen_laag.Value = False
        Worksheets("Opgeslagen pensioenen").Range("AF2") = 0
    End If
    If check_berekeningen_hoog.Visible = False Then
        check_berekeningen_hoog.Value = False
        Worksheets("Opgeslagen pensioenen").Range("Y2") = 0
    End If
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'Weergeven huidig pensioen
    
    Call Huidig_pensioen_weergeven
        
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    
    If dropdown_berekeningen_vervroegen_jaar.Visible = True Then
        If Worksheets("Opgeslagen pensioenen").Cells(2, 11) = 1 Then
            check_berekeningen_vervroegen.Value = True
        Else
            check_berekeningen_vervroegen.Value = False
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 39) = 1 Then
            check_berekeningen_laag_aowopvullen.Value = True
        Else
            check_berekeningen_laag_aowopvullen.Value = False
        End If
        dropdown_berekeningen_vervroegen_jaar = Worksheets("Opgeslagen pensioenen").Cells(2, 12)
        dropdown_berekeningen_vervroegen_maanden = Worksheets("Opgeslagen pensioenen").Cells(2, 13)
    End If
    
    If dropdown_berekeningen_uitstellen_jaar.Visible = True Then
        If Worksheets("Opgeslagen pensioenen").Cells(2, 14) = 1 Then
            check_berekeningen_uitstellen.Value = True
        Else
            check_berekeningen_uitstellen.Value = False
        End If
        dropdown_berekeningen_uitstellen_jaar = Worksheets("Opgeslagen pensioenen").Cells(2, 15)
        dropdown_berekeningen_uitstellen_maand = Worksheets("Opgeslagen pensioenen").Cells(2, 16)
    End If
    
    If dropdown_berekeningen_oppp_percentage.Visible = True Then
        If Worksheets("Opgeslagen pensioenen").Cells(2, 17) = 1 Then
            check_berekeningen_oppp.Value = True
        Else
            check_berekeningen_oppp.Value = False
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 18) = 1 Then
            option_berekeningen_oppp_procenten.Value = True
        Else
            option_berekeningen_oppp_procenten.Value = False
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 20) = 1 Then
            option_berekeningen_oppp_verhouding.Value = True
        Else
            option_berekeningen_oppp_verhouding.Value = False
        End If
        dropdown_berekeningen_oppp_percentage = Worksheets("Opgeslagen pensioenen").Cells(2, 19)
        dropdown_berekeningen_oppp_verhouding_pp = Worksheets("Opgeslagen pensioenen").Cells(2, 21)
    End If
    
    If dropdown_berekeningen_ppop_percentage.Visible = True Then
        If Worksheets("Opgeslagen pensioenen").Cells(2, 22) = 1 Then
            check_berekeningen_ppop.Value = True
        Else
            check_berekeningen_ppop.Value = False
        End If
        dropdown_berekeningen_ppop_percentage = Worksheets("Opgeslagen pensioenen").Cells(2, 24)
    End If
    
    If dropdown_berekeningen_hoog_jaar.Visible = True Then
        If Worksheets("Opgeslagen pensioenen").Cells(2, 25) = 1 Then
            check_berekeningen_ppop.Value = True
        Else
            check_berekeningen_ppop.Value = False
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 28) = 1 Then
            option_berekeningen_hoog_verschil.Value = True
        Else
            option_berekeningen_hoog_verschil.Value = False
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 30) = 1 Then
            option_berekeningen_hoog_verhouding.Value = True
        Else
            option_berekeningen_hoog_verhouding.Value = False
        End If
        dropdown_berekeningen_hoog_jaar = Worksheets("Opgeslagen pensioenen").Cells(2, 26)
        dropdown_berekeningen_hoog_maanden = Worksheets("Opgeslagen pensioenen").Cells(2, 27)
        textbox_berekeningen_hoog_verschil = Worksheets("Opgeslagen pensioenen").Cells(2, 29)
        dropdown_berekeningen_hoog_verhouding_laag = Worksheets("Opgeslagen pensioenen").Cells(2, 31)
    End If
    
    If dropdown_berekeningen_laag_jaar.Visible = True Then
        If Worksheets("Opgeslagen pensioenen").Cells(2, 32) = 1 Then
            check_berekeningen_laag.Value = True
        Else
            check_berekeningen_laag.Value = False
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 35) = 1 Then
            option_berekeningen_laag_verschil.Value = True
        Else
            option_berekeningen_laag_verschil.Value = False
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 37) = 1 Then
            option_berekeningen_laag_verhouding.Value = True
        Else
            option_berekeningen_laag_verhouding.Value = False
        End If
        dropdown_berekeningen_laag_jaar = Worksheets("Opgeslagen pensioenen").Cells(2, 33)
        dropdown_berekeningen_laag_maanden = Worksheets("Opgeslagen pensioenen").Cells(2, 34)
        textbox_berekeningen_laag_verschil = Worksheets("Opgeslagen pensioenen").Cells(2, 36)
        dropdown_berekeningen_laag_verhouding_hoog = Worksheets("Opgeslagen pensioenen").Cells(2, 38)
    End If
End Sub

Private Sub but_berekeningen_opslaan_Click()

    Dim i As Integer
    Dim j As Integer
    Dim q As Integer
    i = 4
    Worksheets("Opgeslagen pensioenen").Cells(2, 40) = 0
    PopUpOpslaan.Show
    q = Worksheets("Opgeslagen pensioenen").Cells(2, 41)
    If Worksheets("Opgeslagen pensioenen").Cells(2, 40) = 1 Then
        If q <> 0 Then
            For j = 2 To 39
                Worksheets("Opgeslagen pensioenen").Cells(q, j) = Worksheets("Opgeslagen pensioenen").Cells(2, j)
            Next j
        Else
            Do While Worksheets("Opgeslagen pensioenen").Cells(i, 2) <> ""
                i = i + 1
            Loop
        
            For j = 2 To 39
                Worksheets("Opgeslagen pensioenen").Cells(i, j) = Worksheets("Opgeslagen pensioenen").Cells(2, j)
            Next j
        End If
    End If
End Sub

Private Sub but_berekeningen_volgende_Click()
    'checken of de berekeningen gedaan zijn enzo
    Berekeningenscherm.Hide
    Vergelijken.Show
End Sub

Private Sub but_berekeningen_vorige_Click()
    'opslaan gegevens?
    Berekeningenscherm.Hide
    Keuzemenu.Show
End Sub


Private Sub UserForm_Initialize()
    'Het maken van de dropdown menu's
    dropdown_berekeningen_uitstellen_jaar.List = Array(55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75)
    dropdown_berekeningen_uitstellen_maand.List = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
    
    dropdown_berekeningen_vervroegen_jaar.List = Array(55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75)
    dropdown_berekeningen_vervroegen_maanden.List = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
    
    dropdown_berekeningen_oppp_percentage.List = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100)
    dropdown_berekeningen_oppp_verhouding_pp.List = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100)
    
    dropdown_berekeningen_ppop_percentage.List = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100)
    
    dropdown_berekeningen_hoog_jaar.List = Array(55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75)
    dropdown_berekeningen_hoog_maanden.List = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
    
    dropdown_berekeningen_laag_jaar.List = Array(55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75)
    dropdown_berekeningen_laag_maanden.List = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
    
    End Sub

Function Numeriekcheck() As Boolean
Dim fout As Boolean
    'uitstellen
    If check_berekeningen_uitstellen = True Then
        'checkt maand
        If IsNumeric(dropdown_berekeningen_uitstellen_maand.Value) = False Or dropdown_berekeningen_uitstellen_maand.Value = "" Then
            dropdown_berekeningen_uitstellen_maand.BackColor = vbRed
            MsgBox ("De maand voor uitstellen is niet correct ingevoerd")
            Numeriekcheck = True
            Exit Function
        Else
            dropdown_berekeningen_uitstellen_maand.BackColor = vbWhite
        End If
        'checkt jaar
        If IsNumeric(dropdown_berekeningen_uitstellen_jaar.Value) = False Or dropdown_berekeningen_uitstellen_jaar.Value = "" Then
            dropdown_berekeningen_uitstellen_jaar.BackColor = vbRed
            MsgBox ("Het jaar voor uitstellen is niet correct ingevoerd")
            Numeriekcheck = True
            Exit Function
        Else
            dropdown_berekeningen_uitstellen_jaar.BackColor = vbWhite
        End If
        'checkt of de uitgestelde leeftijd > huidige leeftijd
        If CInt(dropdown_berekeningen_uitstellen_jaar.Value) < CInt(Worksheets("berekeningen").Cells(3, 3).Value) Then
            dropdown_berekeningen_uitstellen_jaar.BackColor = vbRed
            dropdown_berekeningen_uitstellen_maand.BackColor = vbRed
            MsgBox ("De nieuwe pensioensleeftijd moet hoger zijn dan de huidige pensioensleeftijd om uit te stellen")
            Numeriekcheck = True
            Exit Function
        Else
            dropdown_berekeningen_uitstellen_jaar.BackColor = vbWhite
            dropdown_berekeningen_uitstellen_maand.BackColor = vbWhite
        End If
    End If
    'vervroegen
    If check_berekeningen_vervroegen = True Then
        'checkt maand
        If IsNumeric(dropdown_berekeningen_vervroegen_maanden.Value) = False Or dropdown_berekeningen_vervroegen_maanden.Value = "" Then
            dropdown_berekeningen_vervroegen_maanden.BackColor = vbRed
            MsgBox ("De maand voor vervroegen is niet correct ingevoerd")
            Numeriekcheck = True
            Exit Function
        Else
            dropdown_berekeningen_vervroegen_maanden.BackColor = vbWhite
        End If
        'checkt jaar
        If IsNumeric(dropdown_berekeningen_vervroegen_jaar.Value) = False Or dropdown_berekeningen_vervroegen_jaar.Value = "" Then
            dropdown_berekeningen_vervroegen_jaar.BackColor = vbRed
            MsgBox ("Het jaar voor vervroegen is niet correct ingevoerd")
            Numeriekcheck = True
            Exit Function
        Else
            dropdown_berekeningen_vervroegen_jaar.BackColor = vbWhite
        End If
        'checkt of de vervroegde leeftijd < huidige leeftijd
        If CInt(dropdown_berekeningen_vervroegen_jaar.Value) > (Worksheets("berekeningen").Cells(3, 3).Value) Then
            dropdown_berekeningen_vervroegen_jaar.BackColor = vbRed
            dropdown_berekeningen_vervroegen_maanden.BackColor = vbRed
            MsgBox ("De nieuwe pensioensleeftijd moet lager zijn dan de huidige pensioensleeftijd om te vervroegen")
            Numeriekcheck = True
            Exit Function
        Else
            dropdown_berekeningen_vervroegen_jaar.BackColor = vbWhite
            dropdown_berekeningen_vervroegen_maanden.BackColor = vbWhite
        End If
    End If
    'OPPP
    If check_berekeningen_oppp = True Then
        If option_berekeningen_oppp_procenten.Value = False And option_berekeningen_oppp_verhouding.Value = False Then
            MsgBox ("Vul of een percentage of een verhouding in bij het uitruilen van OP naar PP")
        End If
        If option_berekeningen_oppp_procenten.Value = True Then
            'checkt percentage
            If IsNumeric(dropdown_berekeningen_oppp_percentage.Value) = False Or dropdown_berekeningen_oppp_percentage.Value = "" Then
                dropdown_berekeningen_oppp_percentage.BackColor = vbRed
                MsgBox ("Het percentage uit te ruilen ouderdomspensioen is niet correct ingevoerd")
                Numeriekcheck = True
                Exit Function
            Else
                dropdown_berekeningen_oppp_percentage.BackColor = vbWhite
            End If
        ElseIf option_berekeningen_oppp_verhouding.Value = True Then
            'checkt verhouding PP
            If IsNumeric(dropdown_berekeningen_oppp_verhouding_pp.Value) = False Or dropdown_berekeningen_oppp_verhouding_pp.Value = "" Then
                dropdown_berekeningen_oppp_verhouding_pp.BackColor = vbRed
                MsgBox ("De verhouding OP:PP is niet correct ingevoerd")
                Numeriekcheck = True
                Exit Function
            Else
                dropdown_berekeningen_oppp_verhouding_pp.BackColor = vbWhite
            End If
            'checkt of verhouding PP >= 70
            If dropdown_berekeningen_oppp_verhouding_pp.Value < 70 Then
                dropdown_berekeningen_oppp_verhouding_pp.BackColor = vbRed
                MsgBox ("Bij de verhouding OP:PP moet het aandeel van het PP minstens 70 zijn")
                Numeriekcheck = True
                Exit Function
            Else
                dropdown_berekeningen_oppp_verhouding_pp.BackColor = vbWhite
            End If
        End If
    End If
    'PPOP
'    If check_berekeningen_ppop.Value = True Then
'        If IsNumeric(dropdown_berekeningen_ppop_percentage.Value) = False Or dropdown_berekeningen_ppop_percentage.Value = "" Then
'            dropdown_berekeningen_ppop_percentage.BackColor = vbRed
'            MsgBox ("Het percentage uit te ruilen partnerpensioen is niet correct ingevoerd")
'            Numeriekcheck = True
'            Exit Function
'        Else
'            dropdown_berekeningen_ppop_percentage.BackColor = vbWhite
'        End If
'    End If
    'laaghoog
    If check_berekeningen_laag.Value = True Then
        If option_berekeningen_laag_verschil.Value = False And option_berekeningen_laag_verhouding.Value = False Then
            MsgBox ("Vul of een verschil of een verhouding in bij de laag:hoog constructie")
        End If
        'check jaar
        If IsNumeric(dropdown_berekeningen_laag_jaar.Value) = False Or dropdown_berekeningen_laag_jaar.Value = "" Then
            dropdown_berekeningen_laag_jaar.BackColor = vbRed
            MsgBox ("Het jaar voor de laag:hoog constructie is niet correct ingevoerd")
            Numeriekcheck = True
            Exit Function
        Else
            dropdown_berekeningen_laag_jaar.BackColor = vbWhite
        End If
        'check maanden
        If IsNumeric(dropdown_berekeningen_laag_maanden.Value) = False Or dropdown_berekeningen_laag_maanden.Value = "" Then
            dropdown_berekeningen_laag_maanden.BackColor = vbRed
            MsgBox ("De maanden voor de laag:hoog constructie is niet correct ingevoerd")
            Numeriekcheck = True
            Exit Function
        Else
            dropdown_berekeningen_laag_maanden.BackColor = vbWhite
        End If
        If option_berekeningen_laag_verschil.Value = True Then
            'check verschil
            If IsNumeric(textbox_berekeningen_laag_verschil.Value) = False Or textbox_berekeningen_laag_verschil.Value = "" Then
                textbox_berekeningen_laag_verschil.BackColor = vbRed
                MsgBox ("Het verschil bij de laag:hoog constructie is niet correct ingevoerd")
                Numeriekcheck = True
                Exit Function
            Else
                textbox_berekeningen_laag_verschil.BackColor = vbWhite
            End If
        ElseIf option_berekeningen_laag_verhouding.Value = True Then
            'check verhouding
            If IsNumeric(dropdown_berekeningen_laag_verhouding_hoog.Value) = False Or dropdown_berekeningen_laag_verhouding_hoog.Value = "" Then
                dropdown_berekeningen_laag_verhouding_hoog.BackColor = vbRed
                MsgBox ("De verhouding bij de laag:hoog constructie is niet correct ingevoerd")
                Numeriekcheck = True
                Exit Function
            Else
                dropdown_berekeningen_laag_verhouding_hoog.BackColor = vbWhite
            End If
            'checkt of verhouding hoog >= 75
            If dropdown_berekeningen_laag_verhouding_hoog.Value < 75 Then
                dropdown_berekeningen_laag_verhouding_hoog.BackColor = vbRed
                MsgBox ("Bij de verhouding laag:hoog moet het aandeel van hoog minstens 75 zijn")
                Numeriekcheck = True
                Exit Function
            Else
                dropdown_berekeningen_laag_verhouding_hoog.BackColor = vbWhite
            End If
        End If
    End If
        'hooglaag
    If check_berekeningen_hoog.Value = True Then
        If option_berekeningen_hoog_verschil.Value = False And option_berekeningen_hoog_verhouding.Value = False Then
            MsgBox ("Vul of een verschil of een verhouding in bij de hoog:laag constructie")
        End If
        'check jaar
        If IsNumeric(dropdown_berekeningen_hoog_jaar.Value) = False Or dropdown_berekeningen_hoog_jaar.Value = "" Then
            dropdown_berekeningen_hoog_jaar.BackColor = vbRed
            MsgBox ("Het jaar voor de hoog:laag constructie is niet correct ingevoerd")
            Numeriekcheck = True
            Exit Function
        Else
            dropdown_berekeningen_hoog_jaar.BackColor = vbWhite
        End If
        'check maanden
        If IsNumeric(dropdown_berekeningen_hoog_maanden.Value) = False Or dropdown_berekeningen_hoog_maanden.Value = "" Then
            dropdown_berekeningen_hoog_maanden.BackColor = vbRed
            MsgBox ("De maanden voor de hoog:laag constructie is niet correct ingevoerd")
            Numeriekcheck = True
            Exit Function
        Else
            dropdown_berekeningen_hoog_maanden.BackColor = vbWhite
        End If
        If option_berekeningen_hoog_verschil.Value = True Then
            'check verschil
            If IsNumeric(textbox_berekeningen_hoog_verschil.Value) = False Or textbox_berekeningen_hoog_verschil.Value = "" Then
                textbox_berekeningen_hoog_verschil.BackColor = vbRed
                MsgBox ("Het verschil bij de hoog:laag constructie is niet correct ingevoerd")
                Numeriekcheck = True
                Exit Function
            Else
                textbox_berekeningen_hoog_verschil.BackColor = vbWhite
            End If
        ElseIf option_berekeningen_hoog_verhouding.Value = True Then
            'check verhouding
            If IsNumeric(dropdown_berekeningen_hoog_verhouding_laag.Value) = False Or dropdown_berekeningen_hoog_verhouding_laag.Value = "" Then
                dropdown_berekeningen_hoog_verhouding_laag.BackColor = vbRed
                MsgBox ("De verhouding bij de hoog:laag constructie is niet correct ingevoerd")
                Numeriekcheck = True
                Exit Function
            Else
                dropdown_berekeningen_hoog_verhouding_laag.BackColor = vbWhite
            End If
            'checkt of verhouding hoog >= 75
            If dropdown_berekeningen_hoog_verhouding_laag.Value < 75 Then
                dropdown_berekeningen_hoog_verhouding_laag.BackColor = vbRed
                MsgBox ("Bij de verhouding hoog:laag moet het aandeel van laag minstens 75 zijn")
                Numeriekcheck = True
                Exit Function
            Else
                dropdown_berekeningen_hoog_verhouding_laag.BackColor = vbWhite
            End If
        End If
    End If
    Numeriekcheck = False
End Function

Private Sub but_berekeningen_verwerken_Click()
    'checken of alles numeriek is
    If Numeriekcheck = True Then
        Exit Sub
    Else
    'Berekenen
    
        Call Invullen
        Call Berekenen
        Call Toonresultaat
        
    End If
        
End Sub

Public Sub Invullen()
    
    'Updaten pensioeninformatie
    'Oude informatie weghalen
    Worksheets("Opgeslagen pensioenen").Range("L2:M2") = ""
    Worksheets("Opgeslagen pensioenen").Range("O2:P2") = ""
    Worksheets("Opgeslagen pensioenen").Range("S2") = ""
    Worksheets("Opgeslagen pensioenen").Range("U2") = ""
    Worksheets("Opgeslagen pensioenen").Range("X2") = ""
    Worksheets("Opgeslagen pensioenen").Range("Z2:AA2") = ""
    Worksheets("Opgeslagen pensioenen").Range("AC2") = ""
    Worksheets("Opgeslagen pensioenen").Range("AE2") = ""
    Worksheets("Opgeslagen pensioenen").Range("AG2:AH2") = ""
    Worksheets("Opgeslagen pensioenen").Range("AJ2") = ""
    Worksheets("Opgeslagen pensioenen").Range("AL2") = ""
    
    'Invullen in het tijdelijke geheugen
    If dropdown_berekeningen_vervroegen_jaar.Visible = True Then
        If check_berekeningen_vervroegen.Value = True Then
            Worksheets("Opgeslagen pensioenen").Cells(2, 11) = 1
        Else
            Worksheets("Opgeslagen pensioenen").Cells(2, 11) = 0
        End If
        If check_berekeningen_laag_aowopvullen.Value = True Then
            Worksheets("Opgeslagen pensioenen").Cells(2, 39) = 1
        Else
            Worksheets("Opgeslagen pensioenen").Cells(2, 39) = 0
        End If
        Worksheets("Opgeslagen pensioenen").Cells(2, 12) = dropdown_berekeningen_vervroegen_jaar
        Worksheets("Opgeslagen pensioenen").Cells(2, 13) = dropdown_berekeningen_vervroegen_maanden
    End If
    
    If dropdown_berekeningen_uitstellen_jaar.Visible = True Then
        If check_berekeningen_uitstellen.Value = True Then
            Worksheets("Opgeslagen pensioenen").Cells(2, 14) = 1
        Else
            Worksheets("Opgeslagen pensioenen").Cells(2, 14) = 0
        End If
        Worksheets("Opgeslagen pensioenen").Cells(2, 15) = dropdown_berekeningen_uitstellen_jaar
        Worksheets("Opgeslagen pensioenen").Cells(2, 16) = dropdown_berekeningen_uitstellen_maand
    End If
    
    If dropdown_berekeningen_oppp_percentage.Visible = True Then
        If check_berekeningen_oppp.Value = True Then
            Worksheets("Opgeslagen pensioenen").Cells(2, 17) = 1
        Else
            Worksheets("Opgeslagen pensioenen").Cells(2, 17) = 0
        End If
        If option_berekeningen_oppp_procenten.Value = True Then
            Worksheets("Opgeslagen pensioenen").Cells(2, 18) = 1
        Else
            Worksheets("Opgeslagen pensioenen").Cells(2, 18) = 0
        End If
        If option_berekeningen_oppp_verhouding.Value = True Then
            Worksheets("Opgeslagen pensioenen").Cells(2, 20) = 1
        Else
            Worksheets("Opgeslagen pensioenen").Cells(2, 20) = 0
        End If
        Worksheets("Opgeslagen pensioenen").Cells(2, 19) = dropdown_berekeningen_oppp_percentage
        Worksheets("Opgeslagen pensioenen").Cells(2, 21) = dropdown_berekeningen_oppp_verhouding_pp
    End If
    
    If dropdown_berekeningen_ppop_percentage.Visible = True Then
        If check_berekeningen_ppop.Value = True Then
            Worksheets("Opgeslagen pensioenen").Cells(2, 22) = 1
        Else
            Worksheets("Opgeslagen pensioenen").Cells(2, 22) = 0
        End If
        Worksheets("Opgeslagen pensioenen").Cells(2, 24) = dropdown_berekeningen_ppop_percentage
    End If
    
    If dropdown_berekeningen_hoog_jaar.Visible = True Then
        If check_berekeningen_hoog = True Then
            Worksheets("Opgeslagen pensioenen").Cells(2, 25) = 1
        Else
            Worksheets("Opgeslagen pensioenen").Cells(2, 25) = 0
        End If
        If option_berekeningen_hoog_verschil.Value = True Then
            Worksheets("Opgeslagen pensioenen").Cells(2, 28) = 1
        Else
            Worksheets("Opgeslagen pensioenen").Cells(2, 28) = 0
        End If
        If option_berekeningen_hoog_verhouding.Value = True Then
            Worksheets("Opgeslagen pensioenen").Cells(2, 30) = 1
        Else
            Worksheets("Opgeslagen pensioenen").Cells(2, 30) = 0
        End If
        Worksheets("Opgeslagen pensioenen").Cells(2, 26) = dropdown_berekeningen_hoog_jaar
        Worksheets("Opgeslagen pensioenen").Cells(2, 27) = dropdown_berekeningen_hoog_maanden
        Worksheets("Opgeslagen pensioenen").Cells(2, 29) = textbox_berekeningen_hoog_verschil
        Worksheets("Opgeslagen pensioenen").Cells(2, 31) = dropdown_berekeningen_hoog_verhouding_laag
    End If
    
    If dropdown_berekeningen_laag_jaar.Visible = True Then
        If check_berekeningen_laag.Value = True Then
            Worksheets("Opgeslagen pensioenen").Cells(2, 32) = 1
        Else
            Worksheets("Opgeslagen pensioenen").Cells(2, 32) = 0
        End If
        If option_berekeningen_laag_verschil.Value = True Then
            Worksheets("Opgeslagen pensioenen").Cells(2, 35) = 1
        Else
            Worksheets("Opgeslagen pensioenen").Cells(2, 35) = 0
        End If
        If option_berekeningen_laag_verhouding.Value = True Then
            Worksheets("Opgeslagen pensioenen").Cells(2, 37) = 1
        Else
            Worksheets("Opgeslagen pensioenen").Cells(2, 37) = 0
        End If
        Worksheets("Opgeslagen pensioenen").Cells(2, 33) = dropdown_berekeningen_laag_jaar
        Worksheets("Opgeslagen pensioenen").Cells(2, 34) = dropdown_berekeningen_laag_maanden
        Worksheets("Opgeslagen pensioenen").Cells(2, 36) = textbox_berekeningen_laag_verschil
        Worksheets("Opgeslagen pensioenen").Cells(2, 38) = dropdown_berekeningen_laag_verhouding_hoog
    End If
    
End Sub

Public Sub Huidig_pensioen_weergeven()
 'Zoeken positie
    Dim i As Integer
    i = 5
    Do While Worksheets("BESTANDSOPGAVE").Cells(i, 2) <> Worksheets("Opgeslagen pensioenen").Cells(2, 2)
        i = i + 1
    Loop
    
    'De gegevens van gebruiker invullen op de berekenings sheet
    Worksheets("Huidig pensioen").Cells(32, 1) = Worksheets("BESTANDSOPGAVE").Cells(i, 71)
    Worksheets("Huidig pensioen").Cells(32, 2) = Worksheets("BESTANDSOPGAVE").Cells(i, 72)
    Worksheets("Huidig pensioen").Cells(32, 3) = Worksheets("BESTANDSOPGAVE").Cells(i, 73)
    Worksheets("Huidig pensioen").Cells(34, 1) = Worksheets("BESTANDSOPGAVE").Cells(i, 95)
    Worksheets("Huidig pensioen").Cells(34, 2) = Worksheets("BESTANDSOPGAVE").Cells(i, 96)
    Worksheets("Huidig pensioen").Cells(34, 3) = Worksheets("BESTANDSOPGAVE").Cells(i, 97)
    Worksheets("Huidig pensioen").Cells(36, 1) = Worksheets("BESTANDSOPGAVE").Cells(i, 179)
    Worksheets("Huidig pensioen").Cells(36, 2) = Worksheets("BESTANDSOPGAVE").Cells(i, 180)
    Worksheets("Huidig pensioen").Cells(36, 3) = Worksheets("BESTANDSOPGAVE").Cells(i, 16)
    
    Label_huidig_pensioen.Caption = ""
    Label_huidig_pensioen.Caption = Label_huidig_pensioen.Caption & "Uw huidige pensioen:" & vbCrLf
    If Worksheets("Huidig pensioen").Cells(37, 2).Value = Worksheets("BESTANDSOPGAVE").Cells(i, 25).Value And Worksheets("Huidig pensioen").Cells(37, 3).Value = 0 Then
        Label_huidig_pensioen.Caption = Label_huidig_pensioen.Caption & "Ouderdomspensioen vanaf " & Worksheets("BESTANDSOPGAVE").Cells(i, 25).Value & " jaar: €" & Worksheets("Huidig pensioen").Cells(32, 1).Value & vbCrLf
    ElseIf Worksheets("Huidig pensioen").Cells(37, 2).Value = Worksheets("BESTANDSOPGAVE").Cells(i, 25).Value And Worksheets("Huidig pensioen").Cells(37, 3).Value > 0 Then
        Label_huidig_pensioen.Caption = Label_huidig_pensioen.Caption & "Ouderdomspensioen vanaf " & Worksheets("BESTANDSOPGAVE").Cells(i, 25).Value & " jaar: €" & Worksheets("Huidig pensioen").Cells(32, 1).Value & vbCrLf
        Label_huidig_pensioen.Caption = Label_huidig_pensioen.Caption & "Ouderdomspensioen met AOW vanaf " & Worksheets("Huidig pensioen").Cells(37, 2).Value & " jaar en " & Worksheets("Huidig pensioen").Cells(37, 3).Value & " maanden: €" & CStr(Worksheets("Huidig pensioen").Cells(32, 1).Value + Worksheets("Huidig pensioen").Cells(38, 2).Value) & vbCrLf
    ElseIf Worksheets("Huidig pensioen").Cells(37, 2).Value > Worksheets("BESTANDSOPGAVE").Cells(i, 25).Value Then
        Label_huidig_pensioen.Caption = Label_huidig_pensioen.Caption & "Ouderdomspensioen vanaf " & Worksheets("BESTANDSOPGAVE").Cells(i, 25).Value & " jaar: €" & Worksheets("Huidig pensioen").Cells(32, 1).Value & vbCrLf
        Label_huidig_pensioen.Caption = Label_huidig_pensioen.Caption & "Ouderdomspensioen met AOW vanaf " & Worksheets("Huidig pensioen").Cells(37, 2).Value & " jaar en " & Worksheets("Huidig pensioen").Cells(37, 3).Value & " maanden: €" & CStr(Worksheets("Huidig pensioen").Cells(32, 1).Value + Worksheets("Huidig pensioen").Cells(38, 2).Value) & vbCrLf
    ElseIf Worksheets("Huidig pensioen").Cells(37, 2) < Worksheets("BESTANDSOPGAVE").Cells(i, 25).Value Then
        Label_huidig_pensioen.Caption = Label_huidig_pensioen.Caption & "AOW vanaf " & Worksheets("Huidig pensioen").Cells(37, 2).Value & " jaar en " & Worksheets("Huidig pensioen").Cells(37, 3).Value & " maanden: €" & Worksheets("Huidig pensioen").Cells(38, 2) & vbCrLf
        Label_huidig_pensioen.Caption = Label_huidig_pensioen.Caption & "Ouderdomspensioen met AOW vanaf " & Worksheets("BESTANDSOPGAVE").Cells(i, 25).Value & " jaar: €" & CStr(Worksheets("Huidig pensioen").Cells(32, 1).Value + Worksheets("Huidig pensioen").Cells(38, 2).Value) & vbCrLf
    End If
    Label_huidig_pensioen.Caption = Label_huidig_pensioen.Caption & "Partnerpensioen vanaf uw overlijden: €" & Worksheets("Huidig pensioen").Cells(32, 2).Value

    'Call Afbeelding_huidig_pensioen
End Sub

Public Sub Afbeelding_huidig_pensioen()
    Dim objChrt As ChartObject
    Dim myChart As Chart
    Dim i As Integer
    i = 5
    Do While Worksheets("BESTANDSOPGAVE").Cells(i, 2) <> Worksheets("Opgeslagen pensioenen").Cells(2, 2)
        i = i + 1
    Loop
    
    If Worksheets("Huidig pensioen").Cells(37, 2).Value = Worksheets("BESTANDSOPGAVE").Cells(i, 25).Value And Worksheets("Huidig pensioen").Cells(37, 3).Value = 0 Then
        Set objChrt = Sheets("Huidig pensioen").ChartObjects(2)
    ElseIf Worksheets("Huidig pensioen").Cells(37, 2).Value = Worksheets("BESTANDSOPGAVE").Cells(i, 25).Value And Worksheets("Huidig pensioen").Cells(37, 3).Value > 0 Then
        Set objChrt = Sheets("Huidig pensioen").ChartObjects(3)
    ElseIf Worksheets("Huidig pensioen").Cells(37, 2).Value > Worksheets("BESTANDSOPGAVE").Cells(i, 25).Value Then
        Set objChrt = Sheets("Huidig pensioen").ChartObjects(3)
    ElseIf Worksheets("Huidig pensioen").Cells(37, 2).Value < Worksheets("BESTANDSOPGAVE").Cells(i, 25).Value Then
        Set objChrt = Sheets("Huidig pensioen").ChartObjects(1)
    End If
    
    Set myChart = objChrt.Chart

    myFileName = "myChart.gif"

    On Error Resume Next
    Kill ThisWorkbook.Path & "\" & myFileName
    On Error GoTo 0

    myChart.Export Filename:=ThisWorkbook.Path & "\" & myFileName, Filtername:="GIF"
    
    Me.Img_huidig.Picture = LoadPicture(ThisWorkbook.Path & "\" & myFileName)
End Sub

Public Sub Berekenen()
    
    'Zoeken positie
    Dim i As Integer
    i = 5
    Do While Worksheets("BESTANDSOPGAVE").Cells(i, 2) <> Worksheets("Opgeslagen pensioenen").Cells(2, 2)
        i = i + 1
    Loop
    
    'De gegevens van gebruiker invullen op de berekenings sheet
    Worksheets("Berekeningen").Cells(2, 3) = Worksheets("Berekeningen").Cells(3, 3)
    Worksheets("Berekeningen").Cells(2, 4) = ""
    
    Worksheets("Berekeningen").Cells(6, 1) = Worksheets("BESTANDSOPGAVE").Cells(i, 71)
    Worksheets("Berekeningen").Cells(6, 2) = Worksheets("BESTANDSOPGAVE").Cells(i, 72)
    Worksheets("Berekeningen").Cells(6, 3) = Worksheets("BESTANDSOPGAVE").Cells(i, 73)
    Worksheets("Berekeningen").Cells(8, 1) = Worksheets("BESTANDSOPGAVE").Cells(i, 95)
    Worksheets("Berekeningen").Cells(8, 2) = Worksheets("BESTANDSOPGAVE").Cells(i, 96)
    Worksheets("Berekeningen").Cells(8, 3) = Worksheets("BESTANDSOPGAVE").Cells(i, 97)
    Worksheets("Berekeningen").Cells(10, 1) = Worksheets("BESTANDSOPGAVE").Cells(i, 179)
    Worksheets("Berekeningen").Cells(10, 2) = Worksheets("BESTANDSOPGAVE").Cells(i, 180)
    Worksheets("Berekeningen").Cells(10, 3) = Worksheets("BESTANDSOPGAVE").Cells(i, 16)
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    
    'Berekeningen maken
    If Worksheets("Opgeslagen pensioenen").Range("K2") = 1 Then
'        Worksheets("Berekeningen").Cells(2, 3) = Worksheets("Opgeslagen pensioenen").Range("L2") * 1
'        Worksheets("Berekeningen").Cells(2, 4) = Worksheets("Opgeslagen pensioenen").Range("M2") * 1
        
        Worksheets("Berekeningen").Cells(3, 10) = Worksheets("Opgeslagen pensioenen").Range("L2") * 1
        Worksheets("Berekeningen").Cells(3, 11) = Worksheets("Opgeslagen pensioenen").Range("M2") * 1
        
        Worksheets("Berekeningen").Cells(12, 2) = Worksheets("Berekeningen").Cells(8, 10)
        Worksheets("Berekeningen").Cells(13, 2) = Worksheets("Berekeningen").Cells(9, 10)
        
    End If
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    
    If Worksheets("Opgeslagen pensioenen").Range("N2") = 1 Then
'        Worksheets("Berekeningen").Cells(2, 3) = Worksheets("Opgeslagen pensioenen").Range("O2") * 1
'        Worksheets("Berekeningen").Cells(2, 4) = Worksheets("Opgeslagen pensioenen").Range("P2") * 1

        Worksheets("Berekeningen").Cells(14, 10) = Worksheets("Opgeslagen pensioenen").Range("O2") * 1
        Worksheets("Berekeningen").Cells(14, 11) = Worksheets("Opgeslagen pensioenen").Range("P2") * 1
        
        Worksheets("Berekeningen").Cells(12, 2) = Worksheets("Berekeningen").Cells(20, 10)
        Worksheets("Berekeningen").Cells(13, 2) = Worksheets("Berekeningen").Cells(21, 10)
    End If
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'PP-OP
    If Worksheets("Opgeslagen pensioenen").Range("V2") = 1 Then
        If Worksheets("Opgeslagen pensioenen").Range("K2") = 1 Then
            Worksheets("Berekeningen").Cells(2, 3) = Worksheets("Opgeslagen pensioenen").Range("L2") * 1
            Worksheets("Berekeningen").Cells(2, 4) = Worksheets("Opgeslagen pensioenen").Range("M2") * 1
            Worksheets("Berekeningen").Cells(6, 1) = Worksheets("Berekeningen").Cells(12, 2)
            Worksheets("Berekeningen").Cells(6, 2) = Worksheets("Berekeningen").Cells(13, 2)
        ElseIf Worksheets("Opgeslagen pensioenen").Range("N2") = 1 Then
            Worksheets("Berekeningen").Cells(2, 3) = Worksheets("Opgeslagen pensioenen").Range("O2") * 1
            Worksheets("Berekeningen").Cells(2, 4) = Worksheets("Opgeslagen pensioenen").Range("P2") * 1
            Worksheets("Berekeningen").Cells(6, 1) = Worksheets("Berekeningen").Cells(12, 2)
            Worksheets("Berekeningen").Cells(6, 2) = Worksheets("Berekeningen").Cells(13, 2)
        End If
        
        Worksheets("Berekeningen").Cells(5, 6) = Worksheets("Opgeslagen pensioenen").Range("X2") / 100
        
        Worksheets("Berekeningen").Cells(12, 2) = Worksheets("Berekeningen").Cells(7, 6)
        Worksheets("Berekeningen").Cells(13, 2) = Worksheets("Berekeningen").Cells(8, 6)
    End If
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'OP-PP
    If Worksheets("Opgeslagen pensioenen").Range("Q2") = 1 Then
        If Worksheets("Opgeslagen pensioenen").Range("K2") = 1 Then
            Worksheets("Berekeningen").Cells(2, 3) = Worksheets("Opgeslagen pensioenen").Range("L2") * 1
            Worksheets("Berekeningen").Cells(2, 4) = Worksheets("Opgeslagen pensioenen").Range("M2") * 1
            Worksheets("Berekeningen").Cells(6, 1) = Worksheets("Berekeningen").Cells(12, 2)
            Worksheets("Berekeningen").Cells(6, 2) = Worksheets("Berekeningen").Cells(13, 2)
        ElseIf Worksheets("Opgeslagen pensioenen").Range("N2") = 1 Then
            Worksheets("Berekeningen").Cells(2, 3) = Worksheets("Opgeslagen pensioenen").Range("O2") * 1
            Worksheets("Berekeningen").Cells(2, 4) = Worksheets("Opgeslagen pensioenen").Range("P2") * 1
            Worksheets("Berekeningen").Cells(6, 1) = Worksheets("Berekeningen").Cells(12, 2)
            Worksheets("Berekeningen").Cells(6, 2) = Worksheets("Berekeningen").Cells(13, 2)
        End If
                
        If Worksheets("Opgeslagen pensioenen").Range("R2") = 1 Then
            Worksheets("Berekeningen").Cells(27, 6) = Worksheets("Berekeningen").Cells(2, 3)
            Worksheets("Berekeningen").Cells(27, 7) = Worksheets("Berekeningen").Cells(2, 4)
            Worksheets("Berekeningen").Cells(29, 6) = Worksheets("Opgeslagen pensioenen").Range("S2") / 100
            Worksheets("Berekeningen").Cells(12, 2) = Worksheets("Berekeningen").Cells(32, 6)
            Worksheets("Berekeningen").Cells(13, 2) = Worksheets("Berekeningen").Cells(33, 6)
        End If
        
        If Worksheets("Opgeslagen pensioenen").Range("T2") = 1 Then
            Worksheets("Berekeningen").Cells(14, 6) = Worksheets("Berekeningen").Cells(2, 3)
            Worksheets("Berekeningen").Cells(14, 7) = Worksheets("Berekeningen").Cells(2, 4)
            Worksheets("Berekeningen").Cells(16, 6) = Worksheets("Opgeslagen pensioenen").Range("U2") / 100
            Worksheets("Berekeningen").Cells(12, 2) = Worksheets("Berekeningen").Cells(18, 6)
            Worksheets("Berekeningen").Cells(13, 2) = Worksheets("Berekeningen").Cells(19, 6)
        End If
        
    End If
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'AOW-gat
    If Worksheets("Opgeslagen pensioenen").Range("AM2") = 1 Then
        Worksheets("Berekeningen").Cells(6, 1) = Worksheets("Berekeningen").Cells(12, 2)
        Worksheets("Berekeningen").Cells(6, 2) = Worksheets("Berekeningen").Cells(13, 2)
             
        Worksheets("Berekeningen").Cells(16, 14) = Worksheets("Opgeslagen pensioenen").Range("L2") * 1
        Worksheets("Berekeningen").Cells(16, 15) = Worksheets("Opgeslagen pensioenen").Range("M2") * 1
        
        Worksheets("Berekeningen").Cells(14, 2) = Worksheets("Berekeningen").Cells(30, 14)
        Worksheets("Berekeningen").Cells(15, 2) = Worksheets("Berekeningen").Cells(31, 14)
    End If
                
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'Hoog-laag
    If Worksheets("Opgeslagen pensioenen").Range("Y2") = 1 Then
        If Worksheets("Opgeslagen pensioenen").Range("V2") = 1 Or Worksheets("Opgeslagen pensioenen").Range("Q2") = 1 Then
            Worksheets("Berekeningen").Cells(6, 1) = Worksheets("Berekeningen").Cells(12, 2)
            Worksheets("Berekeningen").Cells(6, 2) = Worksheets("Berekeningen").Cells(13, 2)
        ElseIf Worksheets("Opgeslagen pensioenen").Range("K2") = 1 Then 'vervroegen
            Worksheets("Berekeningen").Cells(2, 3) = Worksheets("Opgeslagen pensioenen").Range("L2") * 1
            Worksheets("Berekeningen").Cells(2, 4) = Worksheets("Opgeslagen pensioenen").Range("M2") * 1
            Worksheets("Berekeningen").Cells(6, 1) = Worksheets("Berekeningen").Cells(12, 2)
            Worksheets("Berekeningen").Cells(6, 2) = Worksheets("Berekeningen").Cells(13, 2)
        ElseIf Worksheets("Opgeslagen pensioenen").Range("N2") = 1 Then 'uitstellen
            Worksheets("Berekeningen").Cells(2, 3) = Worksheets("Opgeslagen pensioenen").Range("O2") * 1
            Worksheets("Berekeningen").Cells(2, 4) = Worksheets("Opgeslagen pensioenen").Range("P2") * 1
            Worksheets("Berekeningen").Cells(6, 1) = Worksheets("Berekeningen").Cells(12, 2)
            Worksheets("Berekeningen").Cells(6, 2) = Worksheets("Berekeningen").Cells(13, 2)
        End If
        
        If Worksheets("Opgeslagen pensioenen").Range("AB2") = 1 Then
                Worksheets("Berekeningen").Cells(26, 10) = Worksheets("Opgeslagen pensioenen").Range("Z2") * 1
                Worksheets("Berekeningen").Cells(26, 11) = Worksheets("Opgeslagen pensioenen").Range("AA2") * 1
                Worksheets("Berekeningen").Cells(27, 10) = Worksheets("Opgeslagen pensioenen").Range("AC2") * 1
                Worksheets("Berekeningen").Cells(14, 2) = Worksheets("Berekeningen").Cells(32, 10)
                Worksheets("Berekeningen").Cells(15, 2) = Worksheets("Berekeningen").Cells(33, 10)
        End If
        
        If Worksheets("Opgeslagen pensioenen").Range("AD2") = 1 Then
            Worksheets("Berekeningen").Cells(3, 14) = Worksheets("Opgeslagen pensioenen").Range("Z2") * 1
            Worksheets("Berekeningen").Cells(3, 15) = Worksheets("Opgeslagen pensioenen").Range("AA2") * 1
            Worksheets("Berekeningen").Cells(4, 14) = Worksheets("Opgeslagen pensioenen").Range("AE2") * 1
            Worksheets("Berekeningen").Cells(14, 2) = Worksheets("Berekeningen").Cells(9, 14)
            Worksheets("Berekeningen").Cells(15, 2) = Worksheets("Berekeningen").Cells(10, 14)
        End If
        
        Worksheets("Berekeningen").Cells(32, 18) = Worksheets("Opgeslagen pensioenen").Range("Z2") * 1
        Worksheets("Berekeningen").Cells(32, 19) = Worksheets("Opgeslagen pensioenen").Range("AA2") * 1
        
    End If
        
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'Laag-Hoog
    If Worksheets("Opgeslagen pensioenen").Range("AF2") = 1 Then
        If Worksheets("Opgeslagen pensioenen").Range("V2") = 1 Or Worksheets("Opgeslagen pensioenen").Range("Q2") = 1 Then
            Worksheets("Berekeningen").Cells(6, 1) = Worksheets("Berekeningen").Cells(12, 2)
            Worksheets("Berekeningen").Cells(6, 2) = Worksheets("Berekeningen").Cells(13, 2)
        ElseIf Worksheets("Opgeslagen pensioenen").Range("K2") = 1 Then 'vervroegen
            Worksheets("Berekeningen").Cells(2, 3) = Worksheets("Opgeslagen pensioenen").Range("L2") * 1
            Worksheets("Berekeningen").Cells(2, 4) = Worksheets("Opgeslagen pensioenen").Range("M2") * 1
            Worksheets("Berekeningen").Cells(6, 1) = Worksheets("Berekeningen").Cells(12, 2)
            Worksheets("Berekeningen").Cells(6, 2) = Worksheets("Berekeningen").Cells(13, 2)
        ElseIf Worksheets("Opgeslagen pensioenen").Range("N2") = 1 Then 'uitstellen
            Worksheets("Berekeningen").Cells(2, 3) = Worksheets("Opgeslagen pensioenen").Range("O2") * 1
            Worksheets("Berekeningen").Cells(2, 4) = Worksheets("Opgeslagen pensioenen").Range("P2") * 1
            Worksheets("Berekeningen").Cells(6, 1) = Worksheets("Berekeningen").Cells(12, 2)
            Worksheets("Berekeningen").Cells(6, 2) = Worksheets("Berekeningen").Cells(13, 2)
        End If
        
        If Worksheets("Opgeslagen pensioenen").Range("AI2") = 1 Then
'            If Worksheets("Opgeslagen pensioenen").Range("K2") = 1 Then
                Worksheets("Berekeningen").Cells(3, 18) = Worksheets("Opgeslagen pensioenen").Range("AG2") * 1
                Worksheets("Berekeningen").Cells(3, 19) = Worksheets("Opgeslagen pensioenen").Range("AH2") * 1
                Worksheets("Berekeningen").Cells(4, 18) = Worksheets("Opgeslagen pensioenen").Range("AJ2") * 1
                Worksheets("Berekeningen").Cells(14, 2) = Worksheets("Berekeningen").Cells(12, 18)
                Worksheets("Berekeningen").Cells(15, 2) = Worksheets("Berekeningen").Cells(11, 18)
        End If
        
        If Worksheets("Opgeslagen pensioenen").Range("AK2") = 1 Then
            Worksheets("Berekeningen").Cells(16, 18) = Worksheets("Opgeslagen pensioenen").Range("AG2") * 1
            Worksheets("Berekeningen").Cells(16, 19) = Worksheets("Opgeslagen pensioenen").Range("AH2") * 1
            Worksheets("Berekeningen").Cells(17, 18) = Worksheets("Opgeslagen pensioenen").Range("AL2") * 1
            Worksheets("Berekeningen").Cells(14, 2) = Worksheets("Berekeningen").Cells(25, 18)
            Worksheets("Berekeningen").Cells(15, 2) = Worksheets("Berekeningen").Cells(24, 18)
        End If
        
        Worksheets("Berekeningen").Cells(32, 18) = Worksheets("Opgeslagen pensioenen").Range("AG2") * 1
        Worksheets("Berekeningen").Cells(32, 19) = Worksheets("Opgeslagen pensioenen").Range("AH2") * 1
    End If
End Sub

Public Sub Toonresultaat()

    If Worksheets("opgeslagen pensioenen").Range("AW2").Value = 1 Then
        ' bruto bedragen
        Label1.Caption = ""
        Label1.Caption = Label1.Caption & "Aangepast pensioen:" & vbCrLf
        If Worksheets("Opgeslagen pensioenen").Cells(2, 11) = 1 Then
            Label1.Caption = Label1.Caption & "Pensioen vervroegd naar " & Worksheets("Opgeslagen pensioenen").Cells(2, 12) & " jaar en " & Worksheets("Opgeslagen pensioenen").Cells(2, 13) & " maanden" & vbCrLf
            If Worksheets("Opgeslagen pensioenen").Cells(2, 39) = 1 Then
                Label1.Caption = Label1.Caption & "Uw AOW-gat wordt opgevult." & vbCrLf
            End If
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 14) = 1 Then
            Label1.Caption = Label1.Caption & "Pensioen verlaat naar " & Worksheets("Opgeslagen pensioenen").Cells(2, 15) & " jaar en " & Worksheets("Opgeslagen pensioenen").Cells(2, 16) & " maanden" & vbCrLf
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 17) = 1 Then
            If Worksheets("Opgeslagen pensioenen").Cells(2, 18) = 1 Then
            Label1.Caption = Label1.Caption & "Ouderdomspensioen uitruilen naar partnerpensioen met " & Worksheets("Opgeslagen pensioenen").Cells(2, 19) & " procent" & vbCrLf
            End If
            If Worksheets("Opgeslagen pensioenen").Cells(2, 20) = 1 Then
            Label1.Caption = Label1.Caption & "Ouderdomspensioen uitruilen naar partnerpensioen met een verhouding van 100:" & Worksheets("Opgeslagen pensioenen").Cells(2, 21) & vbCrLf
            End If
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 22) = 1 Then
            Label1.Caption = Label1.Caption & "Partnerpensioen uitruilen naar ouderdomspensioen met " & Worksheets("Opgeslagen pensioenen").Cells(2, 24) & " procent" & vbCrLf
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 25) = 1 Then
            If Worksheets("Opgeslagen pensioenen").Cells(2, 28) = 1 Then
                Label1.Caption = Label1.Caption & "Hoog-Laag constructie tot " & Worksheets("Opgeslagen pensioenen").Cells(2, 26) & " jaar en " & Worksheets("Opgeslagen pensioenen").Cells(2, 27) & " maanden, met een verschil van " & Worksheets("Opgeslagen pensioenen").Cells(2, 29) & ",- " & vbCrLf
            End If
            If Worksheets("Opgeslagen pensioenen").Cells(2, 30) = 1 Then
                Label1.Caption = Label1.Caption & "Hoog-Laag constructie tot " & Worksheets("Opgeslagen pensioenen").Cells(2, 26) & " jaar en " & Worksheets("Opgeslagen pensioenen").Cells(2, 27) & " maanden, met een verhouding van 100:" & Worksheets("Opgeslagen pensioenen").Cells(2, 31) & vbCrLf
            End If
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 32) = 1 Then
            If Worksheets("Opgeslagen pensioenen").Cells(2, 35) = 1 Then
                Label1.Caption = Label1.Caption & "Laag-Hoog constructie tot " & Worksheets("Opgeslagen pensioenen").Cells(2, 33) & " jaar en " & Worksheets("Opgeslagen pensioenen").Cells(2, 34) & " maanden, met een verschil van " & Worksheets("Opgeslagen pensioenen").Cells(2, 36) & ",- " & vbCrLf
            End If
            If Worksheets("Opgeslagen pensioenen").Cells(2, 37) = 1 Then
                Label1.Caption = Label1.Caption & "Laag-Hoog constructie tot " & Worksheets("Opgeslagen pensioenen").Cells(2, 33) & " jaar en " & Worksheets("Opgeslagen pensioenen").Cells(2, 34) & " maanden, met een verhouding van 100:" & Worksheets("Opgeslagen pensioenen").Cells(2, 38) & vbCrLf
            End If
        End If
        
        Label1.Caption = Label1.Caption & vbCrLf
        If Worksheets("Opgeslagen pensioenen").Cells(2, 25) = 0 Then
        If Worksheets("Opgeslagen pensioenen").Cells(2, 32) = 0 Then
        If Worksheets("Opgeslagen pensioenen").Cells(2, 39) = 1 Then
            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 18) & ",-" & vbCrLf
            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 18) & ",-" & vbCrLf
            Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
        Else
            If Worksheets("berekeningen").Cells(30, 18) * 1 > Worksheets("berekeningen").Cells(29, 18) * 1 Then
                Label1.Caption = Label1.Caption & "AOW vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(31, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(36, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
            End If
            If Worksheets("berekeningen").Cells(30, 18) * 1 < Worksheets("berekeningen").Cells(29, 18) * 1 Then
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(35, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(36, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
            End If
            If Worksheets("berekeningen").Cells(30, 18) * 1 = Worksheets("berekeningen").Cells(29, 18) * 1 Then
                If Worksheets("berekeningen").Cells(30, 19) * 1 < Worksheets("berekeningen").Cells(29, 19) * 1 Then
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(35, 18) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(35, 18) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                Else
                    Label1.Caption = Label1.Caption & "AOW vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(31, 18) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(36, 18) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
            End If
        End If
        End If
        End If
        
        If Worksheets("Opgeslagen pensioenen").Cells(2, 25) = 1 Then
        
            If Worksheets("berekeningen").Cells(30, 18) * 1 > Worksheets("berekeningen").Cells(29, 18) * 1 Then
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
            End If
            
            If Worksheets("berekeningen").Cells(30, 18) * 1 < Worksheets("berekeningen").Cells(29, 18) * 1 Then
                If Worksheets("berekeningen").Cells(29, 18) * 1 > Worksheets("berekeningen").Cells(32, 18) * 1 Then
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 18) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 18) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 18) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
                If Worksheets("berekeningen").Cells(29, 18) * 1 < Worksheets("berekeningen").Cells(32, 18) * 1 Then
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
                If Worksheets("berekeningen").Cells(29, 18) * 1 = Worksheets("berekeningen").Cells(32, 18) * 1 Then
                    If Worksheets("berekeningen").Cells(29, 19) * 1 < Worksheets("berekeningen").Cells(32, 19) * 1 Then
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    Else
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    End If
                End If
            End If
            
            If Worksheets("berekeningen").Cells(30, 18) * 1 = Worksheets("berekeningen").Cells(29, 18) * 1 Then
                If Worksheets("berekeningen").Cells(30, 19) * 1 < Worksheets("berekeningen").Cells(29, 19) * 1 Then
                    If Worksheets("berekeningen").Cells(29, 18) * 1 > Worksheets("berekeningen").Cells(32, 18) * 1 Then
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    End If
                    If Worksheets("berekeningen").Cells(29, 18) * 1 < Worksheets("berekeningen").Cells(32, 18) * 1 Then
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    End If
                    If Worksheets("berekeningen").Cells(29, 18) * 1 = Worksheets("berekeningen").Cells(32, 18) * 1 Then
                        If Worksheets("berekeningen").Cells(29, 19) * 1 < Worksheets("berekeningen").Cells(32, 19) * 1 Then
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 18) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 18) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 18) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                        Else
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 18) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 18) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 18) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                        End If
                    End If
                Else
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
            End If
        End If
        
        
    '    XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        
        
        If Worksheets("Opgeslagen pensioenen").Cells(2, 32) = 1 Then
        
            If Worksheets("berekeningen").Cells(30, 18) * 1 > Worksheets("berekeningen").Cells(29, 18) * 1 Then
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
            End If
            
            If Worksheets("berekeningen").Cells(30, 18) * 1 < Worksheets("berekeningen").Cells(29, 18) * 1 Then
                If Worksheets("berekeningen").Cells(29, 18) * 1 > Worksheets("berekeningen").Cells(32, 18) * 1 Then
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 18) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 18) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 18) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
                If Worksheets("berekeningen").Cells(29, 18) * 1 < Worksheets("berekeningen").Cells(32, 18) * 1 Then
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
                If Worksheets("berekeningen").Cells(29, 18) * 1 = Worksheets("berekeningen").Cells(32, 18) * 1 Then
                    If Worksheets("berekeningen").Cells(29, 19) * 1 < Worksheets("berekeningen").Cells(32, 19) * 1 Then
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    Else
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    End If
                End If
            End If
            
            If Worksheets("berekeningen").Cells(30, 18) * 1 = Worksheets("berekeningen").Cells(29, 18) * 1 Then
                If Worksheets("berekeningen").Cells(30, 19) * 1 < Worksheets("berekeningen").Cells(29, 19) * 1 Then
                    If Worksheets("berekeningen").Cells(29, 18) * 1 > Worksheets("berekeningen").Cells(32, 18) * 1 Then
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    End If
                    If Worksheets("berekeningen").Cells(29, 18) * 1 < Worksheets("berekeningen").Cells(32, 18) * 1 Then
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 18) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    End If
                    If Worksheets("berekeningen").Cells(29, 18) * 1 = Worksheets("berekeningen").Cells(32, 18) * 1 Then
                        If Worksheets("berekeningen").Cells(29, 19) * 1 < Worksheets("berekeningen").Cells(32, 19) * 1 Then
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 18) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 18) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 18) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                        Else
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 18) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 18) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 18) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                        End If
                    End If
                Else
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 18) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
            End If
        End If
        
        
        
        Else
        ' Netto bedragen
        Label1.Caption = ""
        Label1.Caption = Label1.Caption & "Aangepast pensioen:" & vbCrLf
        If Worksheets("Opgeslagen pensioenen").Cells(2, 11) = 1 Then
            Label1.Caption = Label1.Caption & "Pensioen vervroegd naar " & Worksheets("Opgeslagen pensioenen").Cells(2, 12) & " jaar en " & Worksheets("Opgeslagen pensioenen").Cells(2, 13) & " maanden" & vbCrLf
            If Worksheets("Opgeslagen pensioenen").Cells(2, 39) = 1 Then
                Label1.Caption = Label1.Caption & "Uw AOW-gat wordt opgevult." & vbCrLf
            End If
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 14) = 1 Then
            Label1.Caption = Label1.Caption & "Pensioen verlaat naar " & Worksheets("Opgeslagen pensioenen").Cells(2, 15) & " jaar en " & Worksheets("Opgeslagen pensioenen").Cells(2, 16) & " maanden" & vbCrLf
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 17) = 1 Then
            If Worksheets("Opgeslagen pensioenen").Cells(2, 18) = 1 Then
            Label1.Caption = Label1.Caption & "Ouderdomspensioen uitruilen naar partnerpensioen met " & Worksheets("Opgeslagen pensioenen").Cells(2, 19) & " procent" & vbCrLf
            End If
            If Worksheets("Opgeslagen pensioenen").Cells(2, 20) = 1 Then
            Label1.Caption = Label1.Caption & "Ouderdomspensioen uitruilen naar partnerpensioen met een verhouding van 100:" & Worksheets("Opgeslagen pensioenen").Cells(2, 21) & vbCrLf
            End If
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 22) = 1 Then
            Label1.Caption = Label1.Caption & "Partnerpensioen uitruilen naar ouderdomspensioen met " & Worksheets("Opgeslagen pensioenen").Cells(2, 24) & " procent" & vbCrLf
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 25) = 1 Then
            If Worksheets("Opgeslagen pensioenen").Cells(2, 28) = 1 Then
                Label1.Caption = Label1.Caption & "Hoog-Laag constructie tot " & Worksheets("Opgeslagen pensioenen").Cells(2, 26) & " jaar en " & Worksheets("Opgeslagen pensioenen").Cells(2, 27) & " maanden, met een verschil van " & Worksheets("Opgeslagen pensioenen").Cells(2, 29) & ",- " & vbCrLf
            End If
            If Worksheets("Opgeslagen pensioenen").Cells(2, 30) = 1 Then
                Label1.Caption = Label1.Caption & "Hoog-Laag constructie tot " & Worksheets("Opgeslagen pensioenen").Cells(2, 26) & " jaar en " & Worksheets("Opgeslagen pensioenen").Cells(2, 27) & " maanden, met een verhouding van 100:" & Worksheets("Opgeslagen pensioenen").Cells(2, 31) & vbCrLf
            End If
        End If
        If Worksheets("Opgeslagen pensioenen").Cells(2, 32) = 1 Then
            If Worksheets("Opgeslagen pensioenen").Cells(2, 35) = 1 Then
                Label1.Caption = Label1.Caption & "Laag-Hoog constructie tot " & Worksheets("Opgeslagen pensioenen").Cells(2, 33) & " jaar en " & Worksheets("Opgeslagen pensioenen").Cells(2, 34) & " maanden, met een verschil van " & Worksheets("Opgeslagen pensioenen").Cells(2, 36) & ",- " & vbCrLf
            End If
            If Worksheets("Opgeslagen pensioenen").Cells(2, 37) = 1 Then
                Label1.Caption = Label1.Caption & "Laag-Hoog constructie tot " & Worksheets("Opgeslagen pensioenen").Cells(2, 33) & " jaar en " & Worksheets("Opgeslagen pensioenen").Cells(2, 34) & " maanden, met een verhouding van 100:" & Worksheets("Opgeslagen pensioenen").Cells(2, 38) & vbCrLf
            End If
        End If
        
        Label1.Caption = Label1.Caption & vbCrLf
        If Worksheets("Opgeslagen pensioenen").Cells(2, 25) = 0 Then
        If Worksheets("Opgeslagen pensioenen").Cells(2, 32) = 0 Then
            If Worksheets("berekeningen").Cells(30, 18) * 1 > Worksheets("berekeningen").Cells(29, 18) * 1 Then
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(36, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
            End If
            If Worksheets("berekeningen").Cells(30, 18) * 1 < Worksheets("berekeningen").Cells(29, 18) * 1 Then
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(35, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(36, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
            End If
            If Worksheets("berekeningen").Cells(30, 18) * 1 = Worksheets("berekeningen").Cells(29, 18) * 1 Then
                If Worksheets("berekeningen").Cells(30, 19) * 1 < Worksheets("berekeningen").Cells(29, 19) * 1 Then
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(35, 20) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(35, 20) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                Else
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(36, 20) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
            End If
        End If
        End If
        
        If Worksheets("Opgeslagen pensioenen").Cells(2, 25) = 1 Then
        
            If Worksheets("berekeningen").Cells(30, 18) * 1 > Worksheets("berekeningen").Cells(29, 18) * 1 Then
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
            End If
            
            If Worksheets("berekeningen").Cells(30, 18) * 1 < Worksheets("berekeningen").Cells(29, 18) * 1 Then
                If Worksheets("berekeningen").Cells(29, 18) * 1 > Worksheets("berekeningen").Cells(32, 18) * 1 Then
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 20) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 20) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 20) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
                If Worksheets("berekeningen").Cells(29, 18) * 1 < Worksheets("berekeningen").Cells(32, 18) * 1 Then
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
                If Worksheets("berekeningen").Cells(29, 18) * 1 = Worksheets("berekeningen").Cells(32, 18) * 1 Then
                    If Worksheets("berekeningen").Cells(29, 19) * 1 < Worksheets("berekeningen").Cells(32, 19) * 1 Then
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    Else
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    End If
                End If
            End If
            
            If Worksheets("berekeningen").Cells(30, 18) * 1 = Worksheets("berekeningen").Cells(29, 18) * 1 Then
                If Worksheets("berekeningen").Cells(30, 19) * 1 < Worksheets("berekeningen").Cells(29, 19) * 1 Then
                    If Worksheets("berekeningen").Cells(29, 18) * 1 > Worksheets("berekeningen").Cells(32, 18) * 1 Then
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    End If
                    If Worksheets("berekeningen").Cells(29, 18) * 1 < Worksheets("berekeningen").Cells(32, 18) * 1 Then
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    End If
                    If Worksheets("berekeningen").Cells(29, 18) * 1 = Worksheets("berekeningen").Cells(32, 18) * 1 Then
                        If Worksheets("berekeningen").Cells(29, 19) * 1 < Worksheets("berekeningen").Cells(32, 19) * 1 Then
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 20) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 20) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 20) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                        Else
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 20) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 20) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(41, 20) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                        End If
                    End If
                Else
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(39, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
            End If
        End If
        
        
    '    XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        
        
        If Worksheets("Opgeslagen pensioenen").Cells(2, 32) = 1 Then
        
            If Worksheets("berekeningen").Cells(30, 18) * 1 > Worksheets("berekeningen").Cells(29, 18) * 1 Then
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
            End If
            
            If Worksheets("berekeningen").Cells(30, 18) * 1 < Worksheets("berekeningen").Cells(29, 18) * 1 Then
                If Worksheets("berekeningen").Cells(29, 18) * 1 > Worksheets("berekeningen").Cells(32, 18) * 1 Then
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 20) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 20) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 20) & ",-" & vbCrLf
                    Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
                If Worksheets("berekeningen").Cells(29, 18) * 1 < Worksheets("berekeningen").Cells(32, 18) * 1 Then
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
                If Worksheets("berekeningen").Cells(29, 18) * 1 = Worksheets("berekeningen").Cells(32, 18) * 1 Then
                    If Worksheets("berekeningen").Cells(29, 19) * 1 < Worksheets("berekeningen").Cells(32, 19) * 1 Then
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    Else
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    End If
                End If
            End If
            
            If Worksheets("berekeningen").Cells(30, 18) * 1 = Worksheets("berekeningen").Cells(29, 18) * 1 Then
                If Worksheets("berekeningen").Cells(30, 19) * 1 < Worksheets("berekeningen").Cells(29, 19) * 1 Then
                    If Worksheets("berekeningen").Cells(29, 18) * 1 > Worksheets("berekeningen").Cells(32, 18) * 1 Then
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    End If
                    If Worksheets("berekeningen").Cells(29, 18) * 1 < Worksheets("berekeningen").Cells(32, 18) * 1 Then
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 20) & ",-" & vbCrLf
                        Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                    End If
                    If Worksheets("berekeningen").Cells(29, 18) * 1 = Worksheets("berekeningen").Cells(32, 18) * 1 Then
                        If Worksheets("berekeningen").Cells(29, 19) * 1 < Worksheets("berekeningen").Cells(32, 19) * 1 Then
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 20) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 20) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(40, 20) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                        Else
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 20) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 20) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(29, 18) & " jaar en " & Worksheets("berekeningen").Cells(29, 19) & " maanden: " & Worksheets("berekeningen").Cells(46, 20) & ",-" & vbCrLf
                            Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                        End If
                    End If
                Else
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(30, 18) & " jaar en " & Worksheets("berekeningen").Cells(30, 19) & " maanden: " & Worksheets("berekeningen").Cells(44, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Ouderdomspensioen vanaf " & Worksheets("berekeningen").Cells(32, 18) & " jaar en " & Worksheets("berekeningen").Cells(32, 19) & " maanden: " & Worksheets("berekeningen").Cells(45, 20) & ",-" & vbCrLf
                Label1.Caption = Label1.Caption & "Partnerpensioen vanaf uw overlijden: " & Worksheets("berekeningen").Cells(13, 2) & ",-" & vbCrLf
                End If
            End If
            End If
        End If

End Sub

