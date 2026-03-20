$Form1 = New-Object -TypeName System.Windows.Forms.Form
[System.Windows.Forms.ComboBox]$ComboBox1 = $null
[System.Windows.Forms.Label]$Label1 = $null
[System.Windows.Forms.Label]$Label2 = $null
[System.Windows.Forms.RadioButton]$RadioButton1 = $null
[System.Windows.Forms.RadioButton]$RadioButton2 = $null
[System.Windows.Forms.RadioButton]$RadioButton3 = $null
[System.Windows.Forms.RadioButton]$RadioButton4 = $null
[System.Windows.Forms.CheckBox]$CheckBox1 = $null
[System.Windows.Forms.Button]$Button1 = $null
[System.Windows.Forms.Label]$Label3 = $null
[System.Windows.Forms.TextBox]$TextBox1 = $null
[System.Windows.Forms.Label]$Label4 = $null
[System.Windows.Forms.ComboBox]$ComboBox2 = $null
[System.Windows.Forms.Label]$Label5 = $null
[System.Windows.Forms.ComboBox]$ComboBox3 = $null
[System.Windows.Forms.Label]$Label6 = $null
function InitializeComponent
{
$ComboBox1 = (New-Object -TypeName System.Windows.Forms.ComboBox)
$Label1 = (New-Object -TypeName System.Windows.Forms.Label)
$Label2 = (New-Object -TypeName System.Windows.Forms.Label)
$RadioButton1 = (New-Object -TypeName System.Windows.Forms.RadioButton)
$RadioButton2 = (New-Object -TypeName System.Windows.Forms.RadioButton)
$RadioButton3 = (New-Object -TypeName System.Windows.Forms.RadioButton)
$RadioButton4 = (New-Object -TypeName System.Windows.Forms.RadioButton)
$CheckBox1 = (New-Object -TypeName System.Windows.Forms.CheckBox)
$Button1 = (New-Object -TypeName System.Windows.Forms.Button)
$Label3 = (New-Object -TypeName System.Windows.Forms.Label)
$TextBox1 = (New-Object -TypeName System.Windows.Forms.TextBox)
$Label4 = (New-Object -TypeName System.Windows.Forms.Label)
$ComboBox2 = (New-Object -TypeName System.Windows.Forms.ComboBox)
$Label5 = (New-Object -TypeName System.Windows.Forms.Label)
$ComboBox3 = (New-Object -TypeName System.Windows.Forms.ComboBox)
$Label6 = (New-Object -TypeName System.Windows.Forms.Label)
$Form1.SuspendLayout()
#
#ComboBox1
#
$ComboBox1.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]128)))
$ComboBox1.FormattingEnabled = $true
$ComboBox1.ImeMode = [System.Windows.Forms.ImeMode]::On
$ComboBox1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]111,[System.Int32]72))
$ComboBox1.Name = [System.String]'ComboBox1'
$ComboBox1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]281,[System.Int32]27))
$ComboBox1.TabIndex = [System.Int32]1
#
#Label1
#
$Label1.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]128)))
$Label1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]75))
$Label1.Name = [System.String]'Label1'
$Label1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]93,[System.Int32]23))
$Label1.TabIndex = [System.Int32]2
$Label1.Text = [System.String]'文書種別：'
#
#Label2
#
$Label2.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]128)))
$Label2.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]138))
$Label2.Name = [System.String]'Label2'
$Label2.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]109,[System.Int32]23))
$Label2.TabIndex = [System.Int32]3
$Label2.Text = [System.String]'文書の回転：'
$Label2.add_Click($Label2_Click)
#
#RadioButton1
#
$RadioButton1.Checked = $true
$RadioButton1.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]128)))
$RadioButton1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]111,[System.Int32]136))
$RadioButton1.Name = [System.String]'RadioButton1'
$RadioButton1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]104,[System.Int32]24))
$RadioButton1.TabIndex = [System.Int32]4
$RadioButton1.TabStop = $true
$RadioButton1.Text = [System.String]'なし'
$RadioButton1.UseVisualStyleBackColor = $true
#
#RadioButton2
#
$RadioButton2.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]128)))
$RadioButton2.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]248,[System.Int32]136))
$RadioButton2.Name = [System.String]'RadioButton2'
$RadioButton2.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]104,[System.Int32]24))
$RadioButton2.TabIndex = [System.Int32]5
$RadioButton2.Text = [System.String]'180度'
$RadioButton2.UseVisualStyleBackColor = $true
$RadioButton2.add_CheckedChanged($RadioButton2_CheckedChanged)
#
#RadioButton3
#
$RadioButton3.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]128)))
$RadioButton3.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]111,[System.Int32]166))
$RadioButton3.Name = [System.String]'RadioButton3'
$RadioButton3.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]104,[System.Int32]24))
$RadioButton3.TabIndex = [System.Int32]6
$RadioButton3.Text = [System.String]'左90度'
$RadioButton3.UseVisualStyleBackColor = $true
#
#RadioButton4
#
$RadioButton4.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]128)))
$RadioButton4.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]248,[System.Int32]166))
$RadioButton4.Name = [System.String]'RadioButton4'
$RadioButton4.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]104,[System.Int32]24))
$RadioButton4.TabIndex = [System.Int32]7
$RadioButton4.Text = [System.String]'右90度'
$RadioButton4.UseVisualStyleBackColor = $true
#
#CheckBox1
#
$CheckBox1.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]128)))
$CheckBox1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]111,[System.Int32]106))
$CheckBox1.Name = [System.String]'CheckBox1'
$CheckBox1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]241,[System.Int32]24))
$CheckBox1.TabIndex = [System.Int32]9
$CheckBox1.Text = [System.String]'カラー情報無し（グレー）'
$CheckBox1.UseVisualStyleBackColor = $true
#
#Button1
#
$Button1.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]128)))
$Button1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]147,[System.Int32]284))
$Button1.Name = [System.String]'Button1'
$Button1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]110,[System.Int32]35))
$Button1.TabIndex = [System.Int32]10
$Button1.Text = [System.String]'取込み開始'
$Button1.UseVisualStyleBackColor = $true
$Button1.add_MouseClick($clicked)
#
#Label3
#
$Label3.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]128)))
$Label3.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]233))
$Label3.Name = [System.String]'Label3'
$Label3.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]93,[System.Int32]23))
$Label3.TabIndex = [System.Int32]11
$Label3.Text = [System.String]'患者ID：'
#
#TextBox1
#
$TextBox1.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]128)))
$TextBox1.ImeMode = [System.Windows.Forms.ImeMode]::Disable
$TextBox1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]111,[System.Int32]230))
$TextBox1.Name = [System.String]'TextBox1'
$TextBox1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]281,[System.Int32]27))
$TextBox1.TabIndex = [System.Int32]12
#
#Label4
#
$Label4.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$Label4.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]9))
$Label4.Name = [System.String]'Label4'
$Label4.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]93,[System.Int32]23))
$Label4.TabIndex = [System.Int32]13
$Label4.Text = [System.String]'診療科：'
#
#ComboBox2
#
$ComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$ComboBox2.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$ComboBox2.FormattingEnabled = $true
$ComboBox2.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]111,[System.Int32]6))
$ComboBox2.Name = [System.String]'ComboBox2'
$ComboBox2.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]121,[System.Int32]27))
$ComboBox2.TabIndex = [System.Int32]14
#
#Label5
#
$Label5.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$Label5.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]42))
$Label5.Name = [System.String]'Label5'
$Label5.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]93,[System.Int32]23))
$Label5.TabIndex = [System.Int32]15
$Label5.Text = [System.String]'入外：'
#
#ComboBox3
#
$ComboBox3.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$ComboBox3.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Meiryo UI',[System.Single]11.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$ComboBox3.FormattingEnabled = $true
$ComboBox3.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]111,[System.Int32]39))
$ComboBox3.Name = [System.String]'ComboBox3'
$ComboBox3.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]121,[System.Int32]27))
$ComboBox3.TabIndex = [System.Int32]16
#
#Label6
#
$Label6.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$Label6.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]202))
$Label6.Name = [System.String]'Label6'
$Label6.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]380,[System.Int32]2))
$Label6.TabIndex = [System.Int32]17
$Label6.add_Click($Label6_Click)
#
#Form1
#
$Form1.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]404,[System.Int32]331))
$Form1.Controls.Add($Label6)
$Form1.Controls.Add($ComboBox3)
$Form1.Controls.Add($Label5)
$Form1.Controls.Add($ComboBox2)
$Form1.Controls.Add($Label4)
$Form1.Controls.Add($TextBox1)
$Form1.Controls.Add($Label3)
$Form1.Controls.Add($Button1)
$Form1.Controls.Add($CheckBox1)
$Form1.Controls.Add($RadioButton4)
$Form1.Controls.Add($RadioButton3)
$Form1.Controls.Add($RadioButton2)
$Form1.Controls.Add($RadioButton1)
$Form1.Controls.Add($Label2)
$Form1.Controls.Add($Label1)
$Form1.Controls.Add($ComboBox1)
$Form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$Form1.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$Form1.Text = [System.String]'連続スキャン設定'
$Form1.ResumeLayout($false)
$Form1.PerformLayout()
Add-Member -InputObject $Form1 -Name ComboBox1 -Value $ComboBox1 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name Label1 -Value $Label1 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name Label2 -Value $Label2 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name RadioButton1 -Value $RadioButton1 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name RadioButton2 -Value $RadioButton2 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name RadioButton3 -Value $RadioButton3 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name RadioButton4 -Value $RadioButton4 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name CheckBox1 -Value $CheckBox1 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name Button1 -Value $Button1 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name Label3 -Value $Label3 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name TextBox1 -Value $TextBox1 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name Label4 -Value $Label4 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name ComboBox2 -Value $ComboBox2 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name Label5 -Value $Label5 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name ComboBox3 -Value $ComboBox3 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name Label6 -Value $Label6 -MemberType NoteProperty
}
. InitializeComponent
