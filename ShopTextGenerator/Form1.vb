Public Class Form1

    Private paw_color As String = "Black"
    Private base_name As String
    Private fiber_content As String
    Private base_URL As String
    Private base_weight As String
    Private ply As String
    Private yards_grams As String
    Private base_info As String
    Private sku As String
    Private delim As String = "-"




    Private Sub btn_generate_text_Click(sender As Object, e As EventArgs) Handles btn_generate.Click

        ' Get our Yarn Base Information
        generateYarnBaseInfo()

        ' Get our color information 
        generatePawColor()

        Dim full_blurb As String
        Dim yarn_name As String
        Dim yarn_blurb As String

        yarn_name = tb_yarn_name.Text
        yarn_blurb = tb_yarn_text.Text

        ' Get our SKU
        generateSKU(yarn_name)

        full_blurb = "<p><b>" + yarn_name + "</b> " + yarn_blurb + "</p>" + vbCrLf

        Dim div_class As String

        div_class = "<div class = ""list-group"">" + vbCrLf +
                    "<div class = ""list-group-item""><span class=""fa fa-paw fa-fw " + paw_color + """ aria-hidden = ""true""></span> Handpainted/Hand-dyed</div>" + vbCrLf +
                    "<div class = ""list-group-item""><span class=""fa fa-paw fa-fw " + paw_color + """ aria-hidden = ""true""></span>" + fiber_content + "</div>" + vbCrLf +
                    "<div class = ""list-group-item""><span class=""fa fa-paw fa-fw " + paw_color + """ aria-hidden = ""true""></span> <b> <a href=" + base_URL + vbCrLf +
                    " target=""_blank"" title = ""White Whisker Studios - Yarns and Fibers - " + base_name + " rel = ""noopener noreferrer"">" + base_name + "</a></b> base</div>" + vbCrLf +
                    "<div class = ""list-group-item""><span class=""fa fa-paw fa-fw " + paw_color + """ aria-hidden = ""true""></span>" + base_weight + " </div>" + vbCrLf +
                    "<div class = ""list-group-item""><span class=""fa fa-paw fa-fw " + paw_color + """ aria-hidden = ""true""></span>" + ply + " </div>" + vbCrLf +
                    "<div class = ""list-group-item""><span class=""fa fa-paw fa-fw " + paw_color + """ aria-hidden = ""true""></span> " + yards_grams + " </div>" + vbCrLf +
                    "<p></p>" + vbCrLf +
                    "<p><em>Each skein is a hand-created, unique work of art and no two skeins will be identical, even within the same dye lot. " +
                    "If working with more than one skein in a project, it is best to alternate skeins every other row. Despite every effort to accurately depict our product colors, " +
                    "actual shades may vary due to differences in monitor settings.</em></p>"

        Dim full_description As String

        If base_info Is String.Empty Then
            full_description = full_blurb + vbCrLf + div_class
        Else
            full_description = full_blurb + vbCrLf + base_info + vbCrLf + "<p>" + vbCrLf + div_class
        End If

        tb_display.Clear()
        tb_display.Text = full_description


    End Sub

    Private Sub generateYarnBaseInfo()

        Select Case True
            Case rb_HeirloomLuxe.Checked
                base_name = "Heirloom Luxe"
                fiber_content = "50/50 Superwash Merino Wool/Silk"
                base_URL = "/pages/yarn-and-fiber-guide/#heirloom-luxe"
                base_weight = "Sock/Fingering Weight"
                ply = "4-ply"
                yards_grams = "Approximately 437 yards/100 grams"
                base_info = ""
                sku = "HLU" + delim

            Case rb_GlitteringLuxe.Checked
                base_name = "Glittering Luxe"
                fiber_content = "75/20/5 Superwash Merino Wool/Nylon/Stellina"
                base_URL = "/pages/yarn-and-fiber-guide/#glittering-luxe"
                base_weight = "Sock/Fingering Weight"
                ply = "4-ply"
                yards_grams = "Approximately 437 yards/100 grams"
                base_info = ""
                sku = "GLU" + delim

            Case rb_MajesticSock.Checked
                base_name = "Majestic Sock Blank"
                fiber_content = "75/25 Superwash Merino Wool/Nylon"
                base_URL = "/pages/yarn-and-fiber-guide/#majestic-sock"
                base_weight = "Sock/Fingering Weight"
                ply = "4-ply"
                yards_grams = "Approximately 463 yards/100 grams"
                base_info = "Sock blanks are worked from one end to the other - either directly from the sock blank (unraveling the sock blank as you go) or after being wound into a ball."
                sku = "MAJ" + delim

            Case rb_HeirloomLoft.Checked
                base_name = "Heirloom Loft"
                fiber_content = "55/45 Blue Faced Leicester Wool/Silk"
                base_URL = "/pages/yarn-and-fiber-guide/#heirloom-loft"
                base_weight = "DK Weight"
                ply = "4-ply"
                yards_grams = "Approximately 232 yards/100 grams"
                base_info = ""
                sku = "HLO" + delim

            Case rb_SuperflyComfort.Checked
                base_name = "Superfly Comfort Singles"
                fiber_content = "100% Superwash Merino Wool"
                base_URL = "/pages/yarn-and-fiber-guide/#superfly-comfort"
                base_weight = "Sock/Fingering Weight"
                ply = "Single ply"
                yards_grams = "Approximately 400 yards/100 grams"
                base_info = ""
                sku = "SFC" + delim

        End Select

    End Sub

    Private Sub generatePawColor()

        Select Case True

            Case rb_pink.Checked
                paw_color = "pink"

            Case rb_darkPurple.Checked
                paw_color = "dark-purple"

            Case rb_green.Checked
                paw_color = "green"

            Case rb_berry_purple.Checked
                paw_color = "berry-purple"

            Case rb_light_blue.Checked
                paw_color = "light-blue"

            Case rb_peach.Checked
                paw_color = "peach"

            Case rb_yellow.Checked
                paw_color = "yellow"

            Case rb_blue.Checked
                paw_color = "blue"

            Case rb_yellow_green.Checked
                paw_color = "yellow-green"

            Case rb_orange.Checked
                paw_color = "orange"

            Case rb_neon_pink.Checked
                paw_color = "neon-pink"

            Case rb_teal.Checked
                paw_color = "teal"

            Case rb_black.Checked
                paw_color = "black"

            Case rb_brick_red.Checked
                paw_color = "brick-red"

            Case rb_dark_green.Checked
                paw_color = "dark-green"

            Case rb_light_purple.Checked
                paw_color = "light-purple"

            Case rb_dark_blue.Checked
                paw_color = "dark-blue"

            Case rb_olive.Checked
                paw_color = "olive"

            Case rb_light_brown.Checked
                paw_color = "light-brown"

            Case rb_light_green.Checked
                paw_color = "light-green"

            Case rb_coral.Checked
                paw_color = "coral"

            Case rb_grey.Checked
                paw_color = "grey"

            Case rb_dark_teal.Checked
                paw_color = "dark-teal"
        End Select
    End Sub

    Private Sub generateSKU(ByVal name As String)

        Dim now As DateTime = DateTime.Today
        sku = sku + now.ToString("MM") + now.Year.ToString + delim


        Dim sub_name As String

        If name.ToString.StartsWith("The ") Then
            sub_name = name.Substring(4, 3)

        ElseIf name.ToString.StartsWith("A ") Then
            sub_name = name.Substring(2, 3)

        Else
            If name.Length > 3 Then
                sub_name = name.Substring(0, 3)
            Else
                sub_name = name

            End If
        End If

        sku = sku + sub_name

        tb_sku.Clear()
        tb_sku.Text = sku

    End Sub

    Private Sub btn_copy_html_Click(sender As Object, e As EventArgs) Handles btn_copy_html.Click

        If tb_display.Text.Length > 0 Then

            Try
                Clipboard.Clear()
                Clipboard.SetText(tb_display.Text)

                lbl_html_green.Visible = True
                Timer1.Interval = 4000
                Timer1.Start()

            Catch ex As Exception

                lbl_html_red.Visible = True
                Timer1.Interval = 4000
                Timer1.Start()

            End Try

        End If

    End Sub

    Private Sub btn_copy_sku_Click(sender As Object, e As EventArgs) Handles btn_copy_sku.Click

        If tb_sku.Text.Length > 0 Then

            Try
                Clipboard.Clear()
                Clipboard.SetText(tb_sku.Text)

                lbl_sku_green.Visible = True
                Timer1.Interval = 4000
                Timer1.Start()

            Catch ex As Exception

                lbl_sku_red.Visible = True
                Timer1.Interval = 4000
                Timer1.Start()

            End Try

        End If

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        lbl_html_red.Visible = False
        lbl_html_green.Visible = False
        lbl_sku_red.Visible = False
        lbl_sku_green.Visible = False

        Timer1.Stop()

    End Sub
End Class
