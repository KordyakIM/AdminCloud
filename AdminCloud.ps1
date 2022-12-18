$x = @()
#создание форм
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")  
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Tabl")
#input data
#SCCM data
$global:CollectionID = "NN0001EC" #Collection _All systems
$global:cache_username_path = "$PWD\work_files\cache_username.csv"
$global:cache_computername_path = "$PWD\work_files\cache_computername.csv"
#-----
#$global:cache_username_number = 0
[string]$not_found = "пусто"
$username = $Env:username
$Admin_account = Get-ADUser $username -Properties EmailAddress,DisplayName,telephoneNumber | select EmailAddress,DisplayName,telephoneNumber
$ADsite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name
$date = Get-Date -Format d
$time = Get-Date -Format t
$Group_AdminCloud_count = (Get-ADGroup "SID group" -Properties member).member.count
$memberof_username = (Get-ADUser $username -Properties memberof).memberof
$version_app = "11.4.2" #version AdminCloud
#SQL input data
$SQLServer = "SQL server"
$SQLDBName = "AdminCloud"
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security=True"
$sqlConnection.Open()
if($?){
	$sqlConnection.Close()
}else{
	[System.Windows.Forms.MessageBox]::Show("Не удалось подключиться к SQL-базе AdminCloud.`nОбратитесь к Администратору.","Внимание!")
	$sqlConnection.Close()
	break
}
#проверка безопасности
$Group_AdminCloud = Get-ADGroup "SID group" | Select-Object -ExpandProperty DistinguishedName
if($memberof_username -eq $Group_AdminCloud){
	#фиксирует открытие приложения
	$SqlQuery = "
		INSERT INTO dbo.Access_Input
	           ([cloud_version],[ad_site],[admin_name],[admin_mail],[access]) 
	    VALUES 
	           ('$($version_app)','$($ADsite)','$($username)',N'$($Admin_account.EmailAddress)','Ok')
	"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet)
	#проверка версии приложения
	$SqlQuery = "Select * FROM dbo.Config"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet)
	$version = $DataSet.Tables[0].Parameter[0]
	if ([version]$version -gt [version]$version_app){
		Start-Process -filepath "$pwd\AdminCloud_Update.exe" -Wait
	}
	#блок заполнения сессий
	$SqlQuery = "Select * FROM dbo.Sessions"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet)
	$sessions_log = $DataSet.Tables[0]
	if(!($sessions_log | where {$username -eq $_.admin_name} | select -Unique)){
	#if($sessions_log.admin_name -notmatch $username){
		$SqlQuery = "
			INSERT INTO dbo.Sessions
				([admin_name],[display_name],[admin_mail],[admin_telephone],[cloud_version],[ad_site])
			VALUES 
				('$($username)',N'$($Admin_account.DisplayName)',N'$($Admin_account.EmailAddress)','$($Admin_account.telephoneNumber)','$($version_app)','$($ADsite)')
		"
		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $SqlCmd
		$DataSet = New-Object System.Data.DataSet
		$SqlAdapter.Fill($DataSet)
	}elseif($sessions_log | where {$username -eq $_.admin_name} | select -Unique){
	#}elseif([string]$sessions_log.admin_name -match $username){
		#удаление админа из списка сессий
		$SqlQuery = "
			DELETE FROM dbo.Sessions WHERE admin_name = '$username'
			INSERT INTO dbo.Sessions
				([admin_name],[display_name],[admin_mail],[admin_telephone],[cloud_version],[ad_site])
			VALUES 
				('$($username)',N'$($Admin_account.DisplayName)',N'$($Admin_account.EmailAddress)','$($Admin_account.telephoneNumber)','$($version_app)','$($ADsite)')
		"
		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $SqlCmd
		$DataSet = New-Object System.Data.DataSet
		$SqlAdapter.Fill($DataSet)
	}
}else{
	#фиксирует открытие приложения
	$SqlQuery = "
		INSERT INTO dbo.Access_Input
	           ([cloud_version],[ad_site],[admin_name],[admin_mail],[access]) 
	    VALUES 
	           ('$($version_app)','$($ADsite)','$($username)',N'$($Admin_account.EmailAddress)','NO')
	"
	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $uid; Password = $pwd;"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet)
	[System.Windows.Forms.MessageBox]::Show($Admin_account.DisplayName+" не входит в группу доступа AdminCloud.`nОбратитесь к Администратору.","Внимание!")
	break
}

#копирует необходимые файлы для работы
try{
	New-Item "$PWD\work_files\" -ItemType Directory -ea stop
	New-Item "$PWD\work_files\00000409\" -ItemType Directory -ea stop
	Copy-Item "$PWD\CmRcViewerRes.dll" -Destination "$PWD\work_files\00000409\"
	Copy-Item "$PWD\CmRcViewer.exe" -Destination "$PWD\work_files\"
	Copy-Item "$PWD\RdpCoreSccm.dll" -Destination "$PWD\work_files\"
	Copy-Item "$PWD\AdminCloud.docx" -Destination "$PWD\work_files\"
	Copy-Item "$PWD\Cloud.ico" -Destination "$PWD\work_files\"
	Copy-Item "$PWD\psexec.exe" -Destination "$PWD\work_files\"
	#Copy-Item "$PWD\Klist_Purge.ps1" -Destination "$PWD\work_files\"
	Copy-Item "$PWD\CMTrace_x32.exe" -Destination "$PWD\work_files\"
	Copy-Item "$PWD\CMTrace_x64.exe" -Destination "$PWD\work_files\"
	Copy-Item "$PWD\gear.ico" -Destination "$PWD\work_files\"
	Copy-Item "$PWD\DartRemoteViewer.exe" -Destination "$PWD\work_files\"
}catch{}
#импортирует библиотеки
Import-Module ActiveDirectory
#Import-Module "\\vnr00-sc2012\SourcesCM$\Module_PS_For_SCCM\ConfigurationManager.psd1"
$Icon = New-Object system.drawing.icon("$pwd\work_files\Cloud.ico")
$IconGear = New-Object system.drawing.icon("$pwd\work_files\gear.ico")
#---------------------CheckBox---------------------#
#$Font0 = New-Object System.Drawing.Font("Calibri",10,[System.Drawing.FontStyle]::Bold)
$Font0 = New-Object Drawing.Font("Microsoft Sans Serif",8.25, [Drawing.FontStyle]::Bold)
$Font1 = New-Object Drawing.Font("Microsoft Sans Serif",8.25)
#---------------------Создаём форму---------------------#
#$form = new-object System.Windows.Forms.form
$frmMain = New-Object System.Windows.Forms.Form;
$count = New-Object System.Windows.Forms.Label;
#$frmMain.icon =[system.drawing.icon]::ExtractAssociatedIcon("C:\Windows\System32\mmc.exe")   
$frmMain.Icon = $Icon
$frmMain.ClientSize = New-Object System.Drawing.Size(750, 900);    
#$frmMain.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink;  
#$frmMain.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog #фиксирует окно, не растягивает
#$frmMain.AutoSize = $true
$frmMain.AutoScaleDimensions = New-Object System.Drawing.SizeF(200, 100);    
$frmMain.MaximizeBox = $true;
$frmMain.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
$frmMain.Text = "AdminCloud v$version_app"; 
$frmMain.add_Closed({
	#удаление админа из списка сессий
	$SqlQuery = "
		DELETE FROM dbo.Sessions WHERE admin_name = '$username'
	"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet)
});

#Выбор сайта площадки
$comboBox_AD_site = New-Object System.Windows.Forms.comboBox  
$comboBox_AD_site.Width = 80
$comboBox_AD_site.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboBox_AD_site.add_Click({
	$comboBox_AD_site.DroppedDown = $true
})

$comboBox_AD_site.add_SelectedValueChanged({
	if($comboBox_AD_site.text -eq "Organiztion 1"){
		$Global:Location = $comboBox_AD_site
		$global:Domain_Controller = "DC Server"
		[string]$Global:SiteName="SCCM Site"
		[string]$Global:SCCMServer="SCCM Server"
		[string]$Global:SCCMNameSpace="root\sms\site_$SiteName"
		citrix_pkg_tables
		AppV_pkg_tables
		SCCM_pkg_tables
	}elseif($comboBox_AD_site.text -eq "Organiztion 2"){
		$Global:Location = $comboBox_AD_site
		$global:Domain_Controller = "DC Server"
		[string]$Global:SiteName="SCCM Site"
		[string]$Global:SCCMServer="SCCM Server"
		[string]$Global:SCCMNameSpace="root\sms\site_$SiteName"
		citrix_pkg_tables
		AppV_pkg_tables
		SCCM_pkg_tables
	}elseif($comboBox_AD_site.text -eq "Organiztion 3"){
		$global:Location = $comboBox_AD_site
		$global:Domain_Controller = "DC Server"
		[string]$Global:SiteName="SCCM Site"
		[string]$Global:SCCMServer="SCCM Server"
		[string]$Global:SCCMNameSpace="root\sms\site_$SiteName"
		citrix_pkg_tables
		AppV_pkg_tables
		SCCM_pkg_tables
	}elseif($comboBox_AD_site.text -eq "Organiztion 4"){
		$global:Location = $comboBox_AD_site
		$global:Domain_Controller = "DC Server"
		[string]$Global:SiteName="SCCM Site"
		[string]$Global:SCCMServer="SCCM Server"
		[string]$Global:SCCMNameSpace="root\sms\site_$SiteName"
		citrix_pkg_tables
		AppV_pkg_tables
		SCCM_pkg_tables
	}
})
# Main Menu Bar
[void]$frmMain.Controls.Add($comboBox_AD_site)
#$groupboxTextbox.Controls.Add($comboBox_AD_site);
#------------------------------------------------------------------------
[array]$Array = "Organization 1","Organization 2","Organization 3","Organization 4" | sort
ForEach ($Item in $Array) { $comboBox_AD_site.Items.Add($Item) } #заполнение combobox
$comboBox_AD_site.SelectedIndex = 3 # Select the first item by default

#разделение главных закладок
$tabControl_All = New-Object System.Windows.Forms.TabControl
$tabControl_All.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = 180
$tabControl_All.Location = $System_Drawing_Point
$tabControl_All.Name = "tabControl_All"
$tabControl_All.SelectedIndex = 0
$tabControl_All.ShowToolTips = $True
$tabControl_All.Dock = "fill"
#$tabControl_All.Height = 705
$tabControl_All.TabIndex = 4
$frmMain.Controls.Add($tabControl_All);

$tabControl_User = New-Object System.Windows.Forms.TabPage
$tabControl_User.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabControl_User.Location = $System_Drawing_Point
$tabControl_User.Name = "tabControl"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabControl_User.Padding = $System_Windows_Forms_Padding
$tabControl_User.Dock = "Fill"
$tabControl_User.TabIndex = 0
$tabControl_User.Text = "User"
$tabControl_User.UseVisualStyleBackColor = $True
$tabControl_All.Controls.Add($tabControl_User)

$tabControl_System = New-Object System.Windows.Forms.TabPage
$tabControl_System.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabControl_System.Location = $System_Drawing_Point
$tabControl_System.Name = "tabControl_System"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabControl_System.TabIndex = 0
$tabControl_System.Text = "System"
$tabControl_System.UseVisualStyleBackColor = $True
$tabControl_All.Controls.Add($tabControl_System)
#------------------------------------------------------------------------------------------------------------

#проверка главных закладок
$tabControl_All.add_MouseClick({
	$SqlQuery = "Select * FROM dbo.Config"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet)
	$version = $DataSet.Tables[0].Parameter[0]
	if([version]$version -gt [version]$version_app){
		function_version_message
	}
})

#разделение закладок User
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = 180
$tabControl.Location = $System_Drawing_Point
$tabControl.Name = "tabControl1"
$tabControl.SelectedIndex = 0
$tabControl.ShowToolTips = $True
#$System_Drawing_Size = New-Object System.Drawing.Size
#$System_Drawing_Size.Height = 500
#$System_Drawing_Size.Width = 693
#$tabControl.Size = $System_Drawing_Size
$tabControl.Dock = "Fill"
$tabControl.TabIndex = 4
$tabControl.Alignment = "Top"
$tabControl_User.Controls.Add($TabControl);
#закладка App-V
$tabControl_AppV = New-Object System.Windows.Forms.TabPage
$tabControl_AppV.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabControl_AppV.Location = $System_Drawing_Point
$tabControl_AppV.Name = "tabControl"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabControl_AppV.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 205
$System_Drawing_Size.Width = 485
$tabControl_AppV.Size = $System_Drawing_Size
$tabControl_AppV.TabIndex = 0
$tabControl_AppV.Text = "App-V Pkgs"
$tabControl_AppV.UseVisualStyleBackColor = $True
$TabControl.Controls.Add($tabControl_AppV)
	#закладка Citrix
	$tabControl_Citrix = New-Object System.Windows.Forms.TabPage
	$tabControl_Citrix.DataBindings.DefaultDataSourceUpdateMode = 0
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 22
	$tabControl_Citrix.Location = $System_Drawing_Point
	$tabControl_Citrix.Name = "tabControl"
	$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
	$System_Windows_Forms_Padding.All = 3
	$System_Windows_Forms_Padding.Bottom = 3
	$System_Windows_Forms_Padding.Left = 3
	$System_Windows_Forms_Padding.Right = 3
	$System_Windows_Forms_Padding.Top = 3
	$tabControl_Citrix.Padding = $System_Windows_Forms_Padding
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 205
	$System_Drawing_Size.Width = 485
	$tabControl_Citrix.Size = $System_Drawing_Size
	$tabControl_Citrix.TabIndex = 0
	$tabControl_Citrix.Text = "Citrix Pkgs"
	$tabControl_Citrix.UseVisualStyleBackColor = $True
	$TabControl.Controls.Add($tabControl_Citrix)
#закладка SCCM
$tabControl_SCCM = New-Object System.Windows.Forms.TabPage
$tabControl_SCCM.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabControl_SCCM.Location = $System_Drawing_Point
$tabControl_SCCM.Name = "tabControl"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabControl_SCCM.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 205
$System_Drawing_Size.Width = 485
$tabControl_SCCM.Size = $System_Drawing_Size
$tabControl_SCCM.TabIndex = 0
$tabControl_SCCM.Text = "SCCM Pkgs"
$tabControl_SCCM.UseVisualStyleBackColor = $True
$tabControl.Controls.Add($tabControl_SCCM)
#закладка Администрирование пользователя
$tabControlP3 = New-Object System.Windows.Forms.TabPage
$tabControlP3.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabControlP3.Location = $System_Drawing_Point
$tabControlP3.Name = "tabControl"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabControlP3.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 205
$System_Drawing_Size.Width = 485
$tabControlP3.Size = $System_Drawing_Size
$tabControlP3.TabIndex = 0
$tabControlP3.Text = "Администрирование пользователя"
$tabControlP3.UseVisualStyleBackColor = $True
$tabControl.Controls.Add($tabControlP3)
#закладка AppV History
$tabControl_AppV_History = New-Object System.Windows.Forms.TabPage
$tabControl_AppV_History.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabControl_AppV_History.Location = $System_Drawing_Point
$tabControl_AppV_History.Name = "tabControl"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabControl_AppV_History.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 205
$System_Drawing_Size.Width = 485
$tabControl_AppV_History.Size = $System_Drawing_Size
$tabControl_AppV_History.TabIndex = 0
$tabControl_AppV_History.Text = "App-V History"
$tabControl_AppV_History.UseVisualStyleBackColor = $True
$tabControl.Controls.Add($tabControl_AppV_History)
#закладка History
$tabControl_History = New-Object System.Windows.Forms.TabPage
$tabControl_History.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabControl_History.Location = $System_Drawing_Point
$tabControl_History.Name = "tabControl"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabControl_History.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 205
$System_Drawing_Size.Width = 485
$tabControl_History.Size = $System_Drawing_Size
$tabControl_History.TabIndex = 0
$tabControl_History.Text = "AdminCloud History"
$tabControl_History.UseVisualStyleBackColor = $True
$tabControl.Controls.Add($tabControl_History)
#-----------------------------------------------------
#проверка закладок
$tabControl.add_MouseClick({
	$SqlQuery = "Select * FROM dbo.Config"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
	$SqlCmd.CommandText = $SqlQuery
	$SqlCmd.Connection = $SqlConnection
	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet)
	$version = $DataSet.Tables[0].Parameter[0]
	if([version]$version -gt [version]$version_app){function_version_message}
})

#включение дополнительных зависимостей
. "$PSScriptRoot\01_AppV_Pkgs.ps1"
. "$PSScriptRoot\02_Citrix_Pkgs.ps1"
. "$PSScriptRoot\03_SCCM_Pkgs.ps1"
. "$PSScriptRoot\04_User_Info.ps1"
. "$PSScriptRoot\05_AppV_History.ps1"
. "$PSScriptRoot\06_AdminCloud_History.ps1"
. "$PSScriptRoot\07_System.ps1"
. "$PSScriptRoot\08_LastLogon.ps1"
. "$PSScriptRoot\09_SCCM_Apps.ps1"
. "$PSScriptRoot\10_Function.ps1"
#-------------------------------------

#строка запроса
$SearchOnType = $false
$groupbox = New-Object System.Windows.Forms.GroupBox
$groupbox.Dock = "top"
$groupbox.Size = New-Object System.Drawing.Size(420, 175);
$groupbox.Text = $null
$tabControl_User.Controls.Add($groupbox)

#listview для SamAccountName
$listSA = New-Object System.Windows.Forms.ListView
$listSA.Dock = "fill"
#$listSA.Width = 477
$listSA.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$listSA.View = [System.Windows.Forms.View]::Details
$listSA.AutoSize = $true
#$listSA.Size = New-Object System.Drawing.Size(400,120)
$listSA.LabelEdit = $true
$listSA.MultiSelect = $true
$LISTSA.FullRowSelect = $True
$listSA.AllowColumnReorder = $True
$listSA.GridLines = $true
$listSA.Columns.Add("SamAccountName").Width = 105
$listSA.Columns.Add("DisplayName").Width = 120
$listSA.Columns.Add("Email").Width = 100
$listSA.Columns.Add("Telephone").Width = 70
$listSA.Columns.Add("OU").Width = 120
$listSA.Columns.Add("Enabled").Width = 70
$listSA.Columns.Add("Locked").Width = 70
#for ($i=0; $i -lt $listSA.Columns.Count; $i++)
#{
#   $listSA.Columns[$i].Width = 120
#}
$listSA.BackColor = 'Azure'
$tooltip= New-Object System.Windows.Forms.ToolTip
$listSA.add_KeyDown({
	param($sender, $e)
	if ($_.KeyCode -eq "C" -and $e.Control){
		Set-ClipBoard_List -Get_list_array $listSA
	}
	if ($_.keycode -eq "A" -and $e.Control){
		foreach ($Item in $listSA.Items){
		    $Item.selected = $true
		}
	}
})
#сортировка
$listSA.add_ColumnClick({
	if ($listSA.Items.Count -gt 1){
		SortListTwoview -column $_.Column -Get_list_array $listSA
	}
})
#контекстно меню
$listSA.add_ItemSelectionChanged({
	if($listSA.SelectedItems -and ($listSA.SelectedItems.text -ne "")){
		if($listSA.SelectedItems | where-object { $_.SubItems[2].Text -ne ""}){
			$send_mail.Enabled = $true
		}else{
			$send_mail.Enabled = $false
		}
	}else{
		$send_mail.Enabled = $false
	}
})
$listSA.add_MouseDoubleClick({
	$SqlQuery = "Select * FROM dbo.Config"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet)
	$version = $DataSet.Tables[0].Parameter[0]
	if ([version]$version -gt [version]$version_app){
		function_version_message
	}else{
		$progressBar.Text = $progressBarStep;$progressBar.refresh()
		$global:SamAccountUser_from_AdminCloud = $listSA.Items[$listSA.FocusedItem.Index].Text #берем имя за основу для всех вкладок
		$global:Email_from_AdminCloud = $listSA.Items[$listSA.FocusedItem.Index].subitems[2].Text #берем e-mail за основу для всех вкладок
		$global:OU_from_AdminCloud = $listSA.Items[$listSA.FocusedItem.Index].subitems[4].Text #берем OU за основу для всех вкладок
		#запись в кэш User
		Cache_list_User -Cache_variable_User $listSA.Items[$listSA.FocusedItem.Index].Text `
			-Cache_variable_DN $listSA.Items[$listSA.FocusedItem.Index].subitems[1].Text `
			-Cache_variable_Email $listSA.Items[$listSA.FocusedItem.Index].subitems[2].Text `
			-Cache_variable_Telephone $listSA.Items[$listSA.FocusedItem.Index].subitems[3].Text `
			-Cache_variable_OU $listSA.Items[$listSA.FocusedItem.Index].subitems[4].Text
		#подписывает пользователя под кнопками
		$infotab_SAN.Text = $listSA.Items[$listSA.FocusedItem.Index].Text
		$infotab_DN.Text = $listSA.Items[$listSA.FocusedItem.Index].subitems[1].Text
		$infotab_OU_split = $listSA.Items[$listSA.FocusedItem.Index].subitems[4].Text
		$infotab_OU.Text = $infotab_OU_split.Split("/")[1]
		$global:user_memberof = (Get-ADUser -Server $Domain_Controller $listSA.Items[$listSA.FocusedItem.Index].Text -Properties memberof).memberof
		if ($tabControl_AppV.Visible){
			$progressBar.Text = $progressBar.Text -replace $progressBarStep_next,$progressBarStep;$progressBar.refresh()
			AppV_count_check;
			$progressBar.Text = $progressBarEnd
		}elseif ($tabControl_Citrix.Visible){
			$progressBar.Text = $progressBar.Text -replace $progressBarStep_next,$progressBarStep;$progressBar.refresh()
			Citrix_count_check;
			$progressBar.Text = $progressBarEnd
		}elseif ($tabControl_SCCM.Visible){
			$progressBar.Text = $progressBar.Text -replace $progressBarStep_next,$progressBarStep;$progressBar.refresh()
			SCCM_count_check;
			$progressBar.Text = $progressBarEnd
		}elseif ($tabControlP3.Visible){
			Get-CompByUser;
		}elseif ($tabControl_AppV_History.Visible){
			AppV_get_list;
		}elseif ($tabControl_History.Visible){
			History_get_list;
		}
	}
})
####Подсказка
	$listSA.add_MouseHover({
		if ($tabControl_AppV.Visible){
				$tooltip = New-Object System.Windows.Forms.ToolTip
				$tooltip.ToolTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
				$toolTip.AutoPopDelay = 50000
				$toolTip.InitialDelay = 300
				$toolTip.ReshowDelay = 100
				$tooltip.ShowAlways =$true
				$tooltip.SetToolTip($listSA, "Пожалуйста, нажмите два раза на Пользователя и выберите необходимый пакет.")
		}elseif ($tabControl_Citrix.Visible){
				$tooltip = New-Object System.Windows.Forms.ToolTip
				$tooltip.ToolTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
				$toolTip.AutoPopDelay = 50000
				$toolTip.InitialDelay = 300
				$toolTip.ReshowDelay = 100
				$tooltip.ShowAlways =$true
				$tooltip.SetToolTip($listSA, "Пожалуйста, нажмите два раза на Пользователя и выберите необходимый пакет.")
		}elseif ($tabControl_SCCM.Visible){
				$tooltip = New-Object System.Windows.Forms.ToolTip
				$tooltip.ToolTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
				$toolTip.AutoPopDelay = 50000
			    $toolTip.InitialDelay = 300
				$toolTip.ReshowDelay = 100
				$tooltip.ShowAlways =$true
				$tooltip.SetToolTip($listSA, "Пожалуйста, нажмите два раза на Пользователя и выберите необходимый пакет.")
		}elseif ($tabControlP3.Visible){
				$tooltip = New-Object System.Windows.Forms.ToolTip
				$tooltip.ToolTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
				$toolTip.AutoPopDelay = 50000
				$toolTip.InitialDelay = 300
				$toolTip.ReshowDelay = 100
				$tooltip.ShowAlways = $true
				$tooltip.SetToolTip($listSA, "Пожалуйста, нажмите два раза на Пользователя. Получите информацию по пользователю.")
		}elseif ($tabControl_AppV.Visible){
				$tooltip = New-Object System.Windows.Forms.ToolTip
				$tooltip.ToolTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
				$toolTip.AutoPopDelay = 50000
				$toolTip.InitialDelay = 300
				$toolTip.ReshowDelay = 100
				$tooltip.ShowAlways = $true
				$tooltip.SetToolTip($listSA, "Пожалуйста, нажмите два раза на Пользователя (по SamAccountName выгружается информация использования App-V).")
		}elseif ($tabControl_History.Visible){
				$tooltip = New-Object System.Windows.Forms.ToolTip
				$tooltip.ToolTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
				$toolTip.AutoPopDelay = 50000
				$toolTip.InitialDelay = 300
				$toolTip.ReshowDelay = 100
				$tooltip.ShowAlways = $true
				$tooltip.SetToolTip($listSA, "Пожалуйста, нажмите два раза на Пользователя (по SamAccountName выгружается история публикации).")
		}
	})
$groupbox.Controls.Add($ListSA)

#группа текста запроса
$groupboxTextbox = New-Object System.Windows.Forms.GroupBox
$groupboxTextbox.Dock = "top"
$groupboxTextbox.Height = 40;
$groupboxTextbox.Text = $null
$tabControl_User.Controls.Add($groupboxTextbox)
#------------------------------------------------------------------------
#Окно ввода
$textBox0 = New-Object System.Windows.Forms.TextBox;    
$textBox0.Dock = "fill"
#$textBox0.Location = New-Object System.Drawing.Point(10, 20);    
#$textBox0.Size = New-Object System.Drawing.Size(400, 10);    
$textBox0.Height = 10
$textBox0.Name = "textBox0";      
#$textBox0.Text = "user";
$textBox0name = 'Введите имя Пользователя!'
$textbox0.ForeColor = 'LightGray'
$textBox0.Text = $textBox0name
$textBox0_AddGM = 0;
$textBox0.add_Click({
	if($textBox0.Text -eq $textBox0name){
        #Clear the text
        $textBox0.Text = ""
        $textBox0.ForeColor = 'WindowText'
    }
	if($textBox0.Text -eq $textBox0.Tag){
        #Clear the text
        $textBox0.Text = ""
        $textBox0.ForeColor = 'WindowText'
    }
})
$textBox0.add_KeyPress({
	if($textBox0.Visible -and $textBox0.Tag -eq $null){
        #Initialize the watermark and save it in the Tag property
        $textBox0.Tag = $textBox0.Text;
        $textBox0.ForeColor = 'LightGray'
        #If we have focus then clear out the text
        if($textBox0.Focused){
            $textBox0.Text = ""
            $textBox0.ForeColor = 'WindowText'
        }
    }
})
$textBox0.add_Leave({
	if($textBox0.Text -eq ""){
        #Display the watermark
        $textBox0.Text = $textBox0name
        $textBox0.ForeColor = 'LightGray'
    }
	if($textBox0.Text -eq ""){
        #Display the watermark
        $textBox0.Text = $textBox0.Tag
        $textBox0.ForeColor = 'LightGray'
    }
})
$textBox0.add_KeyDown({
	param($sender, $e)
	if ($_.keycode -eq "A" -and $e.Control){
		$textBox0.SelectAll(); $textBox0.Focus()
	}
})
#$frmMain.Controls.Add($textBox0);
$groupboxTextbox.Controls.Add($textBox0)

#кнопка вперед
$btnLeft = New-Object System.Windows.Forms.Button;
$btnLeft.Dock = "right"
$btnLeft.Width = 30
#$btnLeft.Size = New-Object System.Drawing.Size(180, 25);    
$btnLeft.Name = "$btnLeft";
$btnLeft.Text = "<-";
$btnLeft.Enabled = $false;
$btnLeft.Add_Click({
	$global:cache_username_number_save = Import-CSV -Path $cache_username_path -Encoding default
	if($global:cache_username_number_save.cache_SamAccountName.count -eq 1){$global:cache_username_number = 1}
	if($global:cache_username_number -ge 1){
		$global:cache_username_number -= 1
		$textbox0.Text = $cache_username[$cache_username_number].cache_SamAccountName
		Fill-List_cache
	}
})
$groupboxTextbox.Controls.Add($btnLeft);
#кнопка назад
$btnRight = New-Object System.Windows.Forms.Button;    
#$btn1.Location = New-Object System.Drawing.Point(450, 100);   
$btnRight.Dock = "right"
$btnRight.Width = 30
#$btnRight.Size = New-Object System.Drawing.Size(180, 25);    
$btnRight.Name = "$btnRight";    
$btnRight.Text = "->"; 
$btnRight.Enabled = $false;
$btnRight.Add_Click({
	$global:cache_username_number_save = Import-CSV -Path $cache_username_path -Encoding default
	if($global:cache_username_number -lt $global:cache_username_number_save.cache_SamAccountName.count-1 ){
		$global:cache_username_number += 1
		$textbox0.Text = $cache_username[$cache_username_number].cache_SamAccountName
		Fill-List_cache
	}
})
$groupboxTextbox.Controls.Add($btnRight);

#кнопка пользовательского списка с кем работал
$btn_users_list = New-Object System.Windows.Forms.Button;    
$btn_users_list.Dock = "right"
$btn_users_list.Width = 30 
$btn_users_list.Name = "$btn_users_list";    
$btn_users_list.Text = "#"; 
$btn_users_list.Enabled = $false;    
$btn_users_list.Add_Click({
		#Users_List
		$users_form = New-Object System.Windows.Forms.Form
		$users_form.Text = "Кеш-список пользователей"
		$users_form.Size = New-Object System.Drawing.Size(666,350)
		$users_form.StartPosition = "CenterScreen"
		$users_form.AutoSize = $true
		$users_form.MinimizeBox = $False
		$users_form.MaximizeBox = $False
		$users_form.SizeGripStyle= "Hide"
		$users_form.WindowState = "Normal"
		$users_form.FormBorderStyle="Fixed3D"
		$List_users = New-Object System.Windows.Forms.ListView
		$List_users.dock = "Fill"
		$List_users.Height = 200
		$List_users.View = "Details"
		$List_users.MultiSelect = $True
		$List_users.FullRowSelect = $True
		$List_users.AutoSize = $true
		$List_users.LabelEdit = $True
		$List_users.AllowColumnReorder = $True
		$List_users.GridLines = $true
		$List_users.Columns.Add("SamAccountName").width = 105
		$List_users.Columns.Add("DisplayName").width = 120
		$List_users.Columns.Add("Email").width = 100
		$List_users.Columns.Add("Telephone").width = 90
		$List_users.Columns.Add("OU").width = 100
		$List_users.Columns.Add("TimeStamp").width = 120
		$List_users.add_MouseDoubleClick({
			$global:SamAccountUser_from_AdminCloud = $List_users.Items[$List_users.FocusedItem.Index].Text #берем имя за основу для всех вкладок
			$textbox0.Text = $SamAccountUser_from_AdminCloud
			$textBox0.ForeColor = 'WindowText'
			$users_form.Close()
			Fill-List_cache
			#----------------------------------------------------------------------------------------------------------------
				$SqlQuery = "Select * FROM dbo.Config"
				$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
				$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
				$SqlAdapter.SelectCommand = $SqlCmd
				$DataSet = New-Object System.Data.DataSet
				$SqlAdapter.Fill($DataSet)
				$version = $DataSet.Tables[0].Parameter[0]
				if ([version]$version -gt [version]$version_app){
					function_version_message
				}else{
					$progressBar.Text = $progressBarStep;$progressBar.refresh()
					#подписывает пользователя под кнопками
					$infotab_SAN.Text = $List_users.Items[$List_users.FocusedItem.Index].Text
					$infotab_DN.Text = $List_users.Items[$List_users.FocusedItem.Index].subitems[1].Text
					$infotab_OU_split = $List_users.Items[$List_users.FocusedItem.Index].subitems[4].Text
					$infotab_OU.Text = $infotab_OU_split.Split("/")[1]
					$global:user_memberof = (Get-ADUser -Server $Domain_Controller $List_users.Items[$List_users.FocusedItem.Index].Text -Properties memberof).memberof
					if ($tabControl_AppV.Visible){
						$progressBar.Text = $progressBar.Text -replace $progressBarStep_next,$progressBarStep;$progressBar.refresh()
						AppV_count_check;
						$progressBar.Text = $progressBarEnd
					}elseif ($tabControl_Citrix.Visible){
						$progressBar.Text = $progressBar.Text -replace $progressBarStep_next,$progressBarStep;$progressBar.refresh()
						Citrix_count_check;
						$progressBar.Text = $progressBarEnd
					}elseif ($tabControl_SCCM.Visible){
						$progressBar.Text = $progressBar.Text -replace $progressBarStep_next,$progressBarStep;$progressBar.refresh()
						SCCM_count_check;
						$progressBar.Text = $progressBarEnd
					}elseif ($tabControlP3.Visible){
						Get-CompByUser;
					}elseif ($tabControl_AppV_History.Visible){
						AppV_get_list;
					}
				}
		#----------------------------------------------------------------------------------------------------------------
			$listSA.Select()
			foreach ($Item in $listSA.Items){
			    $Item.selected = $true
			}
		})
		$List_users.add_KeyDown({
			param($sender, $e)
			if ($_.KeyCode -eq "C" -and $e.Control){
				Set-ClipBoard_List -Get_list_array $List_users
			}
			if ($_.keycode -eq "A" -and $e.Control){
				foreach ($Item in $List_users.Items){
				    $Item.selected = $true
				}
			}
			if ($_.keycode -eq "Escape"){
				$users_form.Close()
			}
		})
		#сортировка
		$List_users.add_ColumnClick({
			if ($List_users.Items.Count -gt 1){
				SortListTwoview -column $global:_.Column -Get_list_array $List_users
			}
		})
		$users_form.Controls.add($List_users)
		$cache_usernames = Import-CSV -Path $cache_username_path -Encoding default
		$List_users.Items.Clear()
		foreach ($item in $cache_usernames) {
			$I = $List_users.Items.Add($item.cache_SamAccountName)
			$I.SubItems.Add($item.cache_DisplayName)
			$I.SubItems.Add($item.cache_Email)
			$I.SubItems.Add($item.cache_Telephone)
			$I.SubItems.Add($item.cache_OU)
			$I.SubItems.Add($item.cache_TimeStamp)
		}
		#сортирует listview по последнему LastLogon
		if ($List_users.Items.Count -gt 1){SortListOneview -Column 5 -Get_list_array $List_users}
		$users_form.Topmost = $True
		  $users_form.Opacity = 0.91
		      $users_form.ShowIcon = $False
		$users_form.Add_Shown({$users_form.Activate()})
		[void] $users_form.ShowDialog()
})
$groupboxTextbox.Controls.Add($btn_users_list);

#CheckBox for list_account - включает режим списков--------------------------------------------------------------------
$CheckBox_list_account = New-Object System.Windows.Forms.CheckBox;    
$CheckBox_list_account.Size = New-Object System.Drawing.Size(98, 20);    
$CheckBox_list_account.text = "Режим списка"; 
$CheckBox_list_account.Enabled = $true;
#$CheckBox_list_account.Checked = $true;
$CheckBox_list_account.Dock = "right"
$CheckBox_list_account.Add_CheckStateChanged({
	if($CheckBox_list_account.Checked){
		$textbox0.Enabled = $false;
		#вычисляем кличество опубликованных пользователей
		$SqlQuery = "
			Select * FROM dbo.History
			WHERE datetime_publish >= DATEADD(day,-1, GETDATE())
		"
		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $SqlCmd
		$DataSet = New-Object System.Data.DataSet
		$SqlAdapter.Fill($DataSet)
		[array]$History_number_one_day = $DataSet.Tables[0]
		$SqlQuery = "Select * FROM dbo.config"
		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $SqlCmd
		$DataSet = New-Object System.Data.DataSet
		$SqlAdapter.Fill($DataSet)
		[string]$Limit_history = $DataSet.Tables[0].parameter[1]
		if($History_number_one_day.count -ge $Limit_history){
			$CheckBox_list_account.Checked = $false;
			$btn_list_account.Enabled = $false;
			$listSA.Items.Clear() #очистка области пользователей
			[System.Windows.Forms.MessageBox]::Show("В режиме списка превышен порог публикации, в день $Limit_history пользователей.","Внимание!")
		}else{
			$Hash_checbox = [HashTable]::Synchronized(@{AppV_Checklistbox = $AppV_Checklistbox; `
			AppV_Hash = $AppV_Hash; Citrix_Hash = $Citrix_Hash; SCCM_Hash = $SCCM_Hash; `
			btn_list_account = $btn_list_account; Citrix_Checklistbox = $Citrix_Checklistbox; `
			SCCM_Checklistbox = $SCCM_Checklistbox; AppV_Сhecklistbox_Hash = $AppV_Сhecklistbox_Hash; `
			Citrix_Checklistbox_Hash = $Citrix_Checklistbox_Hash; SCCM_Checklistbox_Hash = $SCCM_Checklistbox_Hash})
			$Hash_checbox_Runspace = [RunSpaceFactory]::CreateRunspace()
			$Hash_checbox_Runspace.ApartmentState = "STA"
			$Hash_checbox_Runspace.ThreadOptions = "ReuseThread"
			$Hash_checbox_Runspace.Open()
			$Hash_checbox_Runspace.SessionStateProxy.setVariable("Hash_checbox", $Hash_checbox)
			$Hash_checbox_PowerShell = [PowerShell]::Create()
			$Hash_checbox_PowerShell.Runspace = $Hash_checbox_Runspace
			$Hash_checbox_PowerShell.AddScript({
				$Hash_checbox.btn_list_account.Enabled = $true;
				#App-V
				foreach($item in $Hash_checbox.AppV_checklistbox.Items){
					$item_AppV = $Hash_checbox.AppV_Сhecklistbox_Hash.AppV_Hash | where {$_.app_name -eq $item.text}
					$group = Get-ADGroup $item_AppV.group_SID -Properties member, DistinguishedName | select member, DistinguishedName
					$Hash_checbox.AppV_checklistbox.Items[$item.index].SubItems[1].text = $group.member.count
					$Hash_checbox.AppV_checklistbox.Items[$item.index].backcolor="white"; $Hash_checbox.AppV_checklistbox.Items[$item.index].Checked =0
				}
				#Citrix
				foreach($item in $Hash_checbox.Citrix_Checklistbox.Items){
					$item_Citrix = $Hash_checbox.Citrix_Checklistbox_Hash.Citrix_Hash | where {$_.app_name -eq $item.text}
					$group = Get-ADGroup $item_Citrix.group_SID -Properties member, DistinguishedName | select member, DistinguishedName
					$Hash_checbox.Citrix_Checklistbox.Items[$item.index].SubItems[1].text = $group.member.count
					$Hash_checbox.Citrix_Checklistbox.Items[$item.index].backcolor="white"; $Hash_checbox.Citrix_Checklistbox.Items[$item.index].Checked =0
				}
				#SCCM
				foreach($item in $Hash_checbox.SCCM_Checklistbox.Items){
					$item_SCCM = $Hash_checbox.SCCM_Checklistbox_Hash.SCCM_Hash | where {$_.app_name -eq $item.text}
					$group = Get-ADGroup $item_SCCM.group_SID -Properties member, DistinguishedName | select member, DistinguishedName
					$Hash_checbox.SCCM_Checklistbox.Items[$item.index].SubItems[1].text = $group.member.count
					$Hash_checbox.SCCM_Checklistbox.Items[$item.index].backcolor="white"; $Hash_checbox.SCCM_Checklistbox.Items[$item.index].Checked =0
				}
			}).BeginInvoke()
		}
	}else{
		$global:Get_Content_list_account = "" #очистка области списка
		$textbox0.Enabled = $true;
		$btn_list_account.Enabled = $false;
	}
})
$groupboxTextbox.Controls.Add($CheckBox_list_account);

#кнопка для вызова процедуры опубликовывания пакетов по списку
$btn_list_account = New-Object System.Windows.Forms.Button;    
$btn_list_account.Dock = "right"
$btn_list_account.Width = 52
$btn_list_account.Name = "$btn_list_account";    
$btn_list_account.Text = "Список"; 
$btn_list_account.Enabled = $false;
$btn_list_account.Add_Click({
	#вычисляем кличество опубликованных пользователей
	$SqlQuery = "
		Select * FROM dbo.History
		WHERE datetime_publish >= DATEADD(day,-1, GETDATE())
	"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet)
	[array]$History_number_one_day = $DataSet.Tables[0]
	$SqlQuery = "Select * FROM dbo.config"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet)
	[string]$Limit_history = $DataSet.Tables[0].parameter[1]
	if($History_number_one_day.count -ge $Limit_history){
		$CheckBox_list_account.Checked = $false;
		$btn_list_account.Enabled = $false;
		$listSA.Items.Clear() #очистка области пользователей
		[System.Windows.Forms.MessageBox]::Show("В режиме списка превышен порог публикации, в день $Limit_history пользователей.","Внимание!")
	}else{
		function_list_account
	}
})
$groupboxTextbox.Controls.Add($btn_list_account);


#создание запроса с кэша
function Fill-List_cache {
	$s = $textBox0.Text
    $str = Get-ADUser -Filter {SamAccountName -like $s } -Searchbase "DC=domain,DC=ru" | select SamAccountName
	$listSA.Items.Clear()
	if ($str -eq $null){
		if($progressBar.Text -ne $progressBar_user_not_found_in_AD){
			$progressBar.Text = $progressBar_user_not_found_in_AD;$progressBar.refresh()
		}elseif($progressBar.Text -ne $progressBar_user_not_found){
			$progressBar.Text = $progressBar_user_not_found;$progressBar.refresh()
		}
		if(!(Test-ComputerSecureChannel -ErrorAction SilentlyContinue) -or !(Get-Command -Module "ActiveDirectory" -ErrorAction SilentlyContinue)){
			$progressBar.Text = $progressBarAD;$progressBar.refresh()
		}
	}else{
		$i=0
		foreach ( $item in $str ) {
	    	$s1 = $item -split "}"
			#Добавляем элемент в список
		    $string = $s1[0].Substring(17)
			$listSA.Items.Add($string)
		    $userinfo_ListSA = Get-ADUser $string -Properties DisplayName,EmailAddress,telephoneNumber,CanonicalName,LockedOut | select DisplayName,EmailAddress,telephoneNumber,CanonicalName,Enabled,LockedOut
			if ($userinfo_ListSA.DisplayName -ne $null){
				 if($userinfo_ListSA.DisplayName){$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add($userinfo_ListSA.DisplayName)}else{$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add(@{})}
			}
			$userinfo_ListSA_Enabled = $userinfo_ListSA.Enabled
			$userinfo_ListSA_LockedOut = $userinfo_ListSA.LockedOut
			if($userinfo_ListSA.EmailAddress){$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add($userinfo_ListSA.EmailAddress)}else{$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add(@{})}
			if($userinfo_ListSA.telephoneNumber){$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add($userinfo_ListSA.telephoneNumber)}else{$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add(@{})}
			if($userinfo_ListSA.CanonicalName){$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add($userinfo_ListSA.CanonicalName.replace("domain.ru","domainName"))}else{$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add(@{})}
			$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add("$userinfo_ListSA_Enabled")
			$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add("$userinfo_ListSA_LockedOut")
#			if($userinfo_ListSA_LockedOut){$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add("$false")}else{$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add("$true")}
			$i++
		}
		$progressBar.Text = $progressBarEnd+ "Найдено: " +$i+ " user[s].";$progressBar.refresh()
	}
}

#создание запроса
function Fill-List ($Mask = "*") {
	$s = $textBox0.Text + "*"
    $str = Get-ADUser -Filter {SamAccountName -like $s -or DisplayName -like $s } -Searchbase "DC=domain,DC=ru" | select SamAccountName
	#if($str -eq $null){ $str = Get-ADUser -Filter { DisplayName -like $s } -Searchbase "DC=domain,DC=ru" | select SamAccountName }
	$listSA.Items.Clear()
	if ($str -eq $null){
		if($progressBar.Text -ne $progressBar_user_not_found_in_AD){
			$progressBar.Text = $progressBar_user_not_found_in_AD;$progressBar.refresh()
		}elseif($progressBar.Text -ne $progressBar_user_not_found){
			$progressBar.Text = $progressBar_user_not_found;$progressBar.refresh()
		}
		if(!(Test-ComputerSecureChannel -ErrorAction SilentlyContinue) -or !(Get-Command -Module "ActiveDirectory" -ErrorAction SilentlyContinue)){
			$progressBar.Text = $progressBarAD;$progressBar.refresh()
		}
	}else{
		$i=0
		foreach ( $item in $str ) {
	    	$s1 = $item -split "}"
			#Добавляем элемент в список
		    $string = $s1[0].Substring(17)
			$listSA.Items.Add($string)
		    $userinfo_ListSA = Get-ADUser $string -Properties DisplayName,EmailAddress,telephoneNumber,CanonicalName,LockedOut | select DisplayName,EmailAddress,telephoneNumber,CanonicalName,Enabled,LockedOut
			if ($userinfo_ListSA.DisplayName -ne $null){
				 if($userinfo_ListSA.DisplayName){$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add($userinfo_ListSA.DisplayName)}else{$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add(@{})}
			}
			$userinfo_ListSA_Enabled = $userinfo_ListSA.Enabled
			$userinfo_ListSA_LockedOut = $userinfo_ListSA.LockedOut
			if($userinfo_ListSA.EmailAddress){$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add($userinfo_ListSA.EmailAddress)}else{$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add(@{})}
			if($userinfo_ListSA.telephoneNumber){$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add($userinfo_ListSA.telephoneNumber)}else{$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add(@{})}
			if($userinfo_ListSA.CanonicalName){$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add($userinfo_ListSA.CanonicalName.replace("domain.ru","domainName"))}else{$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add(@{})}
			$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add("$userinfo_ListSA_Enabled")
			$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add("$userinfo_ListSA_LockedOut")
#			if($userinfo_ListSA_LockedOut){$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add("$false")}else{$listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add("$true")}
			$i++
		}
		$progressBar.Text = $progressBarEnd+ "Найдено: " +$i+ " user[s].";$progressBar.refresh()
	}
}
if ($SearchOnType){
    #Добавляем обработчик на событие TextChanged, который выполняет функцию Fill-List
    $textBox0.add_TextChanged({Fill-List ("*" + $textBox0.Text + "*")})
}else{ #Ищем только при нажатии Enter
    #Скриптблок (кусок исполняемого кода) который будет выполнен при нажатии клавиши в поле поиска
    $SB_KeyPress = {
        #Если была нажата клавиша Enter (13) то...
        if (13 -eq $_.keychar){
		 if ((!$textBox0.Text) -or ( $textBox0.Text -eq " ") ){
				 #Вызываем функцию Fill-List
			[System.Windows.Forms.MessageBox]::Show("Введите пожалуйста имя пользователя!")
				return
		 	 }else{  
				Fill-List ("*" + $textBox0.Text + "*")
        	 }
		}
    }
    #Добавляем обработчик на событие KeyPress, указав в качестве выполняемого кода $SB_KeyPress
    $textBox0.add_KeyPress($SB_KeyPress)
}
#контекстное меню
$send_mail = New-Object System.Windows.Forms.MenuItem
$send_mail.Text = "Отправить письмо"
$send_mail.Enabled = $false
$send_mail.add_Click({
	[string]$e_mail= ""
	$listSA.SelectedItems | % {if($_.SubItems[2].Text -ne ""){$e_mail += $_.SubItems[2].Text+"; "}}
	$outlook = New-Object -comObject Outlook.Application
	$mail = $outlook.CreateItem(0)
	$mail.To = $e_mail
	$mail.Subject = "AdminCloud"
	$mail.Body = ""
	$inspector = $mail.GetInspector
	$inspector.Activate()
})
$ContextMenu = New-Object System.Windows.Forms.ContextMenu
$ContextMenu.MenuItems.AddRange(@($send_mail))
$groupbox.ContextMenu = $ContextMenu
#-----------------------------------------------------------

#group for CheckBox--------------------------------------------------------------------
$GroupInsideAC = New-Object System.Windows.Forms.GroupBox
$GroupInsideAC.dock = "right"
$GroupInsideAC.Width = 200
#$GroupInsideAC.Height = 40
$groupbox.Controls.Add($GroupInsideAC);
#инфорамация по нажатию SamAccountName
$infotab_SAN = New-Object System.Windows.Forms.Label
$infotab_SAN.Dock = "bottom"
$infotab_SAN.Height = 20
#$infotab_SAN.Size = New-Object System.Drawing.Size(250, 25); 
$infotab_SAN.ForeColor = "black"
$infotab_SAN.Text = ""
$infotab_SAN.Width = 400
$GroupInsideAC.Controls.Add($infotab_SAN);
#инфорамация по нажатию DisplayName
$infotab_DN = New-Object System.Windows.Forms.Label
$infotab_DN.Dock = "bottom"
$infotab_DN.Height = 20
#$infotab_DN.Size = New-Object System.Drawing.Size(300, 25); 
$infotab_DN.ForeColor = "black"
$infotab_DN.Text = ""
$infotab_DN.Width = 400
$GroupInsideAC.Controls.Add($infotab_DN);
#инфорамация по нажатию OU
$infotab_OU = New-Object System.Windows.Forms.Label
$infotab_OU.Dock = "bottom"
$infotab_OU.Height = 20
#$infotab_OU.Size = New-Object System.Drawing.Size(300, 25); 
$infotab_OU.ForeColor = "black"
$infotab_OU.Text = ""
$infotab_OU.Width = 400
$GroupInsideAC.Controls.Add($infotab_OU);

#кнопка добавить в группу
$btn1 = New-Object System.Windows.Forms.Button;    
#$btn1.Location = New-Object System.Drawing.Point(450, 100);   
$btn1.Dock = "top"
$btn1.Size = New-Object System.Drawing.Size(180, 25);    
$btn1.Name = "btn0";    
$btn1.Text = "Опубликовать";
$btn1.Enabled = $true;    
$btn1.Add_Click({
	#SQL test connect
	$SQLServer = "SQL Server"
	$SQLDBName = "AdminCloud"
	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security=True"
	$sqlConnection.Open()
	if($?){
		$sqlConnection.Close()
		#$Get_Content_list_account = Get-Content -Path $TypeBox_list_account_path -ErrorAction SilentlyContinue
		if($CheckBox_list_account.Checked){
			#вычисляем кличество опубликованных пользователей
			$SqlQuery = "
				Select * FROM dbo.History
				WHERE datetime_publish >= DATEADD(day,-1, GETDATE())
			"
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCmd
			$DataSet = New-Object System.Data.DataSet
			$SqlAdapter.Fill($DataSet)
			[array]$History_number_one_day = $DataSet.Tables[0]
			$SqlQuery = "Select * FROM dbo.config"
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection)
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCmd
			$DataSet = New-Object System.Data.DataSet
			$SqlAdapter.Fill($DataSet)
			[string]$Limit_history = $DataSet.Tables[0].parameter[1]
			if($History_number_one_day.count -ge $Limit_history){
				$CheckBox_list_account.Checked = $false;
				$btn_list_account.Enabled = $false;
				$listSA.Items.Clear() #очистка области пользователей
				[System.Windows.Forms.MessageBox]::Show("В режиме списка превышен порог публикации, в день $Limit_history пользователей.","Внимание!")
			}else{
				if([string]::IsNullOrEmpty($Get_Content_list_account)){
					[System.Windows.Forms.MessageBox]::Show("Выбран режим списка, но данные не загружены.","Внимание!")
				}else{
					$found_check = $false
					foreach($item in $AppV_checklistbox.Items){if ($AppV_checklistbox.Items[$item.index].Checked){$found_check = $true}}
					foreach($item in $Citrix_checklistbox.Items){if ($Citrix_checklistbox.Items[$item.index].Checked){$found_check = $true}}
					foreach($item in $SCCM_Checklistbox.Items){if ($SCCM_Checklistbox.Items[$item.index].Checked){$found_check = $true}}
					if ($found_check){
						$output = [System.Windows.Forms.MessageBox]::Show("Выбран режим списка, всего пользователей - "+$listSA.Items.count+", опубликовать?","Внимание!",4)
						if ($output -eq "YES"){
							$global:access_to_group = ""
							$global:access_to_table = $false
							#Получение уч.записи для доступа к группам
							$uid = "username";
							$pwd = "password"
							$SqlConnection_sa = New-Object System.Data.SqlClient.SqlConnection
							$SqlConnection_sa.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $uid; Password = $pwd;"
							#SQL запрос искомых данных в dbo.AdminCloud_Account
							$SqlQuery = "Select * FROM dbo.Account_Audit"
							$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection_sa)
							$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
							$SqlAdapter.SelectCommand = $SqlCmd
							$DataSet = New-Object System.Data.DataSet
							$SqlAdapter.Fill($DataSet)
							$AdminCloud_Account_Audit = $DataSet.Tables[0]
							$AdminCloud_Account_Audit_password = ConvertTo-SecureString $AdminCloud_Account_Audit.password -AsPlainText -Force
							$global:AdminCloud_Account_Audit_cred = New-Object System.Management.Automation.PSCredential ($AdminCloud_Account_Audit.UserLogin, $AdminCloud_Account_Audit_password)
							#----------------------------------------------
							if ($tabControl_AppV.Visible){
								$progressBar.Text = $progressBarStep;$progressBar.refresh()
								AppV_add_group
								if($global:access_to_group -eq "No"){ $progressBar.Text = $progressBar_access_to_group;$progressBar.refresh() }
								elseif($global:access_to_table){ $progressBar.Text = $progressBar_access_to_table;$progressBar.refresh() }
								$global:Get_Content_list_account = "" #очистка области списка
								$listSA.Items.Clear() #очистка области пользователей
							} elseif ($tabControl_Citrix.Visible){
								$progressBar.Text = $progressBarStep;$progressBar.refresh()
								Citrix_add_group;
								if($global:access_to_group -eq "No"){ $progressBar.Text = $progressBar_access_to_group;$progressBar.refresh() }
								elseif($global:access_to_table){ $progressBar.Text = $progressBar_access_to_table;$progressBar.refresh() }
								$global:Get_Content_list_account = "" #очистка области списка
								$listSA.Items.Clear() #очистка области пользователей
							} elseif ($tabControl_SCCM.Visible){
								$progressBar.Text = $progressBarStep;$progressBar.refresh()
								SCCM_add_group;
								if($global:access_to_group -eq "No"){ $progressBar.Text = $progressBar_access_to_group;$progressBar.refresh() }
								elseif($global:access_to_table){ $progressBar.Text = $progressBar_access_to_table;$progressBar.refresh() }
								$global:Get_Content_list_account = "" #очистка области списка
								$listSA.Items.Clear() #очистка области пользователей
							}
						}
					}else{
						[System.Windows.Forms.MessageBox]::Show("Выберите пакет.","Внимание!")
					}
				}
			}
		}else{
			if([string]::IsNullOrEmpty($SamAccountUser_from_AdminCloud)){
				[System.Windows.Forms.MessageBox]::Show("Выберите пользователя.","Внимание!")
			}else{
				$found_check = $false
				[array]$Packages = @()
				if ($tabControl_AppV.Visible){
					foreach($item in $AppV_checklistbox.Items){
						if ($AppV_checklistbox.Items[$item.index].Checked){$found_check = $true; $Packages+=$item.text}
					}
					if($Packages.count -gt 1){$Packages = $Packages -replace "$",","}
				} elseif ($tabControl_Citrix.Visible){
					foreach($item in $Citrix_checklistbox.Items){
						if ($Citrix_checklistbox.Items[$item.index].Checked){$found_check = $true; $Packages+=$item.text}
					}
					if($Packages.count -gt 1){$Packages = $Packages -replace "$",","}
				} elseif ($tabControl_SCCM.Visible){
					foreach($item in $SCCM_Checklistbox.Items){
						if ($SCCM_Checklistbox.Items[$item.index].Checked){$found_check = $true; $Packages+=$item.text}
					}
					if($Packages.count -gt 1){$Packages = $Packages -replace "$",","}
				}
				if ($found_check){
			 		$output = [System.Windows.Forms.MessageBox]::Show("Вы хотите опубликовать "+ $Packages +" пользователю "+$SamAccountUser_from_AdminCloud+"?","Внимание!",4)
					if ($output -eq "YES"){
						$global:access_to_group = ""
						$global:access_to_table = $false
						#Получение уч.записи для доступа к группам
						$uid = "username";
						$pwd = "password"
						$SqlConnection_sa = New-Object System.Data.SqlClient.SqlConnection
						$SqlConnection_sa.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $uid; Password = $pwd;"
						#SQL запрос искомых данных в dbo.AdminCloud_Account
						$SqlQuery = "Select * FROM dbo.Account_Audit"
						$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection_sa)
						$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
						$SqlAdapter.SelectCommand = $SqlCmd
						$DataSet = New-Object System.Data.DataSet
						$SqlAdapter.Fill($DataSet)
						$AdminCloud_Account_Audit = $DataSet.Tables[0]
						$AdminCloud_Account_Audit_password = ConvertTo-SecureString $AdminCloud_Account_Audit.password -AsPlainText -Force
						$global:AdminCloud_Account_Audit_cred = New-Object System.Management.Automation.PSCredential ($AdminCloud_Account_Audit.UserLogin, $AdminCloud_Account_Audit_password)
						#----------------------------------------------
						if ($tabControl_AppV.Visible){
							$progressBar.Text = $progressBarStep;$progressBar.refresh()
							AppV_add_group
							$progressBar.Text = $progressBarEnd;$progressBar.refresh()
							if($global:access_to_group -eq "No"){ $progressBar.Text = $progressBar_access_to_group;$progressBar.refresh() }
							elseif($global:access_to_table){ $progressBar.Text = $progressBar_access_to_table;$progressBar.refresh() }
						} elseif ($tabControl_Citrix.Visible){
							$progressBar.Text = $progressBarStep;$progressBar.refresh()
							Citrix_add_group;
							$progressBar.Text = $progressBarEnd;$progressBar.refresh()
							if($global:access_to_group -eq "No"){ $progressBar.Text = $progressBar_access_to_group;$progressBar.refresh() }
							elseif($global:access_to_table){ $progressBar.Text = $progressBar_access_to_table;$progressBar.refresh() }
						} elseif ($tabControl_SCCM.Visible){
							$progressBar.Text = $progressBarStep;$progressBar.refresh()
							SCCM_add_group;
							$progressBar.Text = $progressBarEnd;$progressBar.refresh()
							if($global:access_to_group -eq "No"){ $progressBar.Text = $progressBar_access_to_group;$progressBar.refresh() }
							elseif($global:access_to_table){ $progressBar.Text = $progressBar_access_to_table;$progressBar.refresh() }
						}
					}
				}else{
					[System.Windows.Forms.MessageBox]::Show("Выберите пакет.","Внимание!")
				}
			}
		}
	}else{
		[System.Windows.Forms.MessageBox]::Show("Не удалось подключиться к SQL-базе AdminCloud.`nОбратитесь к Администратору.","Внимание!")
		$sqlConnection.Close()
	}
})
$GroupInsideAC.Controls.Add($btn1);

#кнопка удалить из группы
$btn2 = New-Object System.Windows.Forms.Button;    
#$btn2.Location = New-Object System.Drawing.Point(450, 150);    
$btn2.Dock = "top"
$btn2.Size = New-Object System.Drawing.Size(180, 25);    
$btn2.Name = "btn0";    
$btn2.Text = "Удалить"; 
$btn2.Enabled = $true;    
$btn2.Add_Click({
	#SQL input data
	$SQLServer = "SQL Server"
	$SQLDBName = "AdminCloud"
	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security=True"
	$sqlConnection.Open()
	if($?){
		$sqlConnection.Close()
		$found_check = $false
		foreach($item in $AppV_checklistbox.Items){if ($AppV_checklistbox.Items[$item.index].Checked){$found_check = $true}}
		foreach($item in $Citrix_checklistbox.Items){if ($Citrix_checklistbox.Items[$item.index].Checked){$found_check = $true}}
		foreach($item in $SCCM_Checklistbox.Items){if ($SCCM_Checklistbox.Items[$item.index].Checked){$found_check = $true}}
#		foreach($item in $checklistbox_SCCM_NT.Items){if ($checklistbox_SCCM_NT.Items[$item.index].Checked){$found_check = $true}}
		if($CheckBox_list_account.Checked){
			if([string]::IsNullOrEmpty($Get_Content_list_account)){
				[System.Windows.Forms.MessageBox]::Show("Выбран режим списка, но данные не загружены.","Внимание!")
			}else{
				if ($found_check){
					$output = [System.Windows.Forms.MessageBox]::Show("Выбран режим списка, всего пользователей - "+$listSA.Items.count+", удалить?","Внимание!",4)
					if ($output -eq "YES"){
						$global:access_to_group = ""
						$global:access_to_table = $false
						#Получение уч.записи для доступа к группам
						$uid = "username";
						$pwd = "password"
						$SqlConnection_sa = New-Object System.Data.SqlClient.SqlConnection
						$SqlConnection_sa.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $uid; Password = $pwd;"
						#SQL запрос искомых данных в dbo.AdminCloud_Account
						$SqlQuery = "Select * FROM dbo.Account_Audit"
						$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection_sa)
						$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
						$SqlAdapter.SelectCommand = $SqlCmd
						$DataSet = New-Object System.Data.DataSet
						$SqlAdapter.Fill($DataSet)
						$AdminCloud_Account_Audit = $DataSet.Tables[0]
						$AdminCloud_Account_Audit_password = ConvertTo-SecureString $AdminCloud_Account_Audit.password -AsPlainText -Force
						$global:AdminCloud_Account_Audit_cred = New-Object System.Management.Automation.PSCredential ($AdminCloud_Account_Audit.UserLogin, $AdminCloud_Account_Audit_password)
						#----------------------------------------------
						if ($tabControl_AppV.Visible){
							$progressBar.Text = $progressBarStep;$progressBar.refresh()
							AppV_rem_group
							if($global:access_to_group -eq "No"){ $progressBar.Text = $progressBar_access_to_group;$progressBar.refresh() }
							elseif($global:access_to_table){ $progressBar.Text = $progressBar_access_to_table;$progressBar.refresh() }
							$global:Get_Content_list_account = ""
							$listSA.Items.Clear() #очистка области пользователей
						}elseif ($tabControl_Citrix.Visible){
							$progressBar.Text = $progressBarStep;$progressBar.refresh()
							Citrix_rem_group
							if($global:access_to_group -eq "No"){ $progressBar.Text = $progressBar_access_to_group;$progressBar.refresh() }
							elseif($global:access_to_table){ $progressBar.Text = $progressBar_access_to_table;$progressBar.refresh() }
							$global:Get_Content_list_account = ""
							$listSA.Items.Clear() #очистка области пользователей
						}elseif ($tabControl_SCCM.Visible){
							$progressBar.Text = $progressBarStep;$progressBar.refresh()
							SCCM_rem_group
							if($global:access_to_group -eq "No"){ $progressBar.Text = $progressBar_access_to_group;$progressBar.refresh() }
							elseif($global:access_to_table){ $progressBar.Text = $progressBar_access_to_table;$progressBar.refresh() }
							$global:Get_Content_list_account = ""
							$listSA.Items.Clear() #очистка области пользователей
						}
					}
				}else{
					[System.Windows.Forms.MessageBox]::Show("Выберите пакет.","Внимание!")
				}
			}
		}else{
			if([string]::IsNullOrEmpty($SamAccountUser_from_AdminCloud)){
				[System.Windows.Forms.MessageBox]::Show("Выберите пользователя.","Внимание!")
			}else{
				$found_check = $false
				[array]$Packages = @()
				if ($tabControl_AppV.Visible){
					foreach($item in $AppV_checklistbox.Items){
						if ($AppV_checklistbox.Items[$item.index].Checked){$found_check = $true; $Packages+=$item.text}
					}
					if($Packages.count -gt 1){$Packages = $Packages -replace "$",","}
				} elseif ($tabControl_Citrix.Visible){
					foreach($item in $Citrix_checklistbox.Items){
						if ($Citrix_checklistbox.Items[$item.index].Checked){$found_check = $true; $Packages+=$item.text}
					}
					if($Packages.count -gt 1){$Packages = $Packages -replace "$",","}
				} elseif ($tabControl_SCCM.Visible){
					foreach($item in $SCCM_Checklistbox.Items){
						if ($SCCM_Checklistbox.Items[$item.index].Checked){$found_check = $true; $Packages+=$item.text}
					}
					if($Packages.count -gt 1){$Packages = $Packages -replace "$",","}
				}
				if ($found_check){
		 			$output = [System.Windows.Forms.MessageBox]::Show("Вы хотите удалить "+ $Packages +" пользователю "+$SamAccountUser_from_AdminCloud+"?","Внимание!",4)
					if ($output -eq "YES"){
						$global:access_to_group = ""
						$global:access_to_table = $false
						#Получение уч.записи для доступа к группам
						$uid = "username";
						$pwd = "password"
						$SqlConnection_sa = New-Object System.Data.SqlClient.SqlConnection
						$SqlConnection_sa.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $uid; Password = $pwd;"
						#SQL запрос искомых данных в dbo.AdminCloud_Account
						$SqlQuery = "Select * FROM dbo.Account_Audit"
						$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConnection_sa)
						$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
						$SqlAdapter.SelectCommand = $SqlCmd
						$DataSet = New-Object System.Data.DataSet
						$SqlAdapter.Fill($DataSet)
						$AdminCloud_Account_Audit = $DataSet.Tables[0]
						$AdminCloud_Account_Audit_password = ConvertTo-SecureString $AdminCloud_Account_Audit.password -AsPlainText -Force
						$global:AdminCloud_Account_Audit_cred = New-Object System.Management.Automation.PSCredential ($AdminCloud_Account_Audit.UserLogin, $AdminCloud_Account_Audit_password)
						#----------------------------------------------
						if ($tabControl_AppV.Visible){
							$progressBar.Text = $progressBarStep;$progressBar.refresh()
							AppV_rem_group
							$progressBar.Text = $progressBarEnd;$progressBar.refresh()
							if($global:access_to_group -eq "No"){ $progressBar.Text = $progressBar_access_to_group;$progressBar.refresh() }
							elseif($global:access_to_table){ $progressBar.Text = $progressBar_access_to_table;$progressBar.refresh() }
						}elseif ($tabControl_Citrix.Visible){
							$progressBar.Text = $progressBarStep;$progressBar.refresh()
							Citrix_rem_group
							$progressBar.Text = $progressBarEnd;$progressBar.refresh()
							if($global:access_to_group -eq "No"){ $progressBar.Text = $progressBar_access_to_group;$progressBar.refresh() }
							elseif($global:access_to_table){ $progressBar.Text = $progressBar_access_to_table;$progressBar.refresh() }
						}elseif ($tabControl_SCCM.Visible){
							$progressBar.Text = $progressBarStep;$progressBar.refresh()
							SCCM_rem_group
							$progressBar.Text = $progressBarEnd;$progressBar.refresh()
							if($global:access_to_group -eq "No"){ $progressBar.Text = $progressBar_access_to_group;$progressBar.refresh() }
							elseif($global:access_to_table){ $progressBar.Text = $progressBar_access_to_table;$progressBar.refresh() }
						}
					}
				}else{
					[System.Windows.Forms.MessageBox]::Show("Выберите пакет.","Внимание!")
				}
			}
		}
	}else{
		[System.Windows.Forms.MessageBox]::Show("Не удалось подключиться к SQL-базе AdminCloud.`nОбратитесь к Администратору.","Внимание!")
		$sqlConnection.Close()
	}
})
$GroupInsideAC.Controls.Add($btn2);

#group for button--------------------------------------------------------------------
$Group_menu_down = New-Object System.Windows.Forms.GroupBox
$Group_menu_down.dock = "bottom"
$Group_menu_down.Height = 42
$frmMain.Controls.Add($Group_menu_down);
#------------------------------------------------------------------------------------
#бегунок
$progressBar = New-Object System.Windows.Forms.Label
$progressBar.ForeColor = "red"
$progressBar.Dock = "bottom"
$progressBar.Height = 18
$progressBarStep = "|||-->"
$progressBarStep_next = "-->"
$progressBarEnd = "Готово! "
$progressBar_access_to_group = "Нет доступа к приложению!"
$progressBar_access_to_table = "Нет доступа к таблице!"
$progressBar_AppV = "Нет соединения с App-V!"
$progressBarSCCM = "Нет соединения с SCCM!"
$progressBarAD = "Нет соединения с ActiveDirectory!"
$progressBar_not_found_history_01 = "Не найдена история - "
$progressBar_not_found_history_02 = "Все еще не найдена история - "
$progressBar_user_not_found_in_AppV = "Пользователь не найден в App-V!"
$progressBar_package_not_found_in_AppV = "Информация по выбранному пакету не найдена!"
$progressBar_user_not_found_in_SCCM = "Пользователь не найден в SCCM!"
$progressBar_comp_not_found_in_SCCM = "Компьютер не найден в SCCM!"
$progressBar_user_not_found_in_AD = "Пользователь не найден в ActiveDirectory!"
$progressBar_comp_not_found_in_AD = "Компьютер не найден в ActiveDirectory!"
$progressBar_user_not_found = "Все еще не найден Пользователь!"
$progressBar_comp_not_found = "Все еще не найден Компьютер!"
$progressBar_dont_update_AppV = "Политика App-V, не обновлена! "
$progressBar_dont_update_SCCM = "Политика SCCM, не обновлена! "
$progressBar.add_TextChanged({
    if($progressBar.Text -eq "|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||-->"){
        $progressBar.Text = $progressBarStep
	}
})
if ( ((Get-Date -UFormat %p) -eq "AM") -or ((Get-Date).Hour -lt 12)) {
	If (((Get-Date).Hour) -ge 3) {
		$progressBar.Text = "Доброе утро, " + $Admin_account.DisplayName + "!"
	}elseif(((Get-Date).Hour) -le 3){
		$progressBar.Text = "Доброй ночи, " + $Admin_account.DisplayName + "!"
	}
}elseif( ((Get-Date -UFormat %p) -eq "PM") -or ((Get-Date).Hour -lt 16) -or ((Get-Date).Hour -ge 16)){
    If(((Get-Date).Hour) -lt 16){
		$progressBar.Text = "Добрый день, " + $Admin_account.DisplayName + "!"
	}elseif(((Get-Date).Hour) -ge 16){
		$progressBar.Text = "Добрый вечер, " + $Admin_account.DisplayName + "!"
	}
} 
#$progressBar.Location = New-Object System.Drawing.Point(10, 250);
$Group_menu_down.Controls.Add($progressBar)
#--------------------------меню file--------------------------------------------------------------------------------------------------------
$menu_AC_Main = New-Object System.Windows.Forms.MenuStrip
$menu_AC_File = New-Object System.Windows.Forms.ToolStripMenuItem
$menu_AC_Feedback = New-Object System.Windows.Forms.ToolStripMenuItem
$menu_AC_about = New-Object System.Windows.Forms.ToolStripMenuItem
$menu_AC_sessions = New-Object System.Windows.Forms.ToolStripMenuItem
$menu_AC_CM_remote = New-Object System.Windows.Forms.ToolStripMenuItem
$menu_AC_exit = New-Object System.Windows.Forms.ToolStripMenuItem
$menu_AC_Dart = New-Object System.Windows.Forms.ToolStripMenuItem
$menu_AC_run = New-Object System.Windows.Forms.ToolStripMenuItem
#[System.Drawing.SystemIcons]::Information
# Main ToolStrip
#[void]$Feedback_form.Controls.Add($mainToolStrip)
 
 
# Main Menu Bar
[void]$frmMain.Controls.Add($menu_AC_Main)

 
# Menu Options - File
$menu_AC_File.Text = ""
$menu_AC_file.Image = $IconGear
$menu_AC_file.Alignment = "right"
$menu_AC_file.DropDownDirection = "BelowLeft"
[void]$menu_AC_Main.Items.Add($menu_AC_File)

# Menu Options - File / CM remote
#$menuOpen.Image        = [System.IconExtractor]::Extract(&quot;shell32.dll&quot;, 4, $true)
#$menuOpen.ShortcutKeys = "Control, O"
$menu_AC_run.Text = "Запустить"
$menu_AC_run.Alignment = "left"
$menu_AC_run.DropDownDirection = "left"
[void]$menu_AC_File.DropDownItems.Add($menu_AC_run)

	$menu_AC_CM_remote.Text = "CM Remote"
	$menu_AC_CM_remote.Add_Click({
		cm_remote
	})
	[void]$menu_AC_run.DropDownItems.Add($menu_AC_CM_remote)
	
	$menu_AC_Dart.Text = "Dart"
	$menu_AC_Dart.Add_Click({
		dart
	})
	[void]$menu_AC_run.DropDownItems.Add($menu_AC_Dart)
 
# Menu Options - File / сессии
#$menuSave.Image        = [System.IconExtractor]::Extract(&quot;shell32.dll&quot;, 36, $true)
#$menuSave.ShortcutKeys = "F2"
$menu_AC_sessions.Text = "Сессии"
$menu_AC_sessions.Add_Click({
	sessions
})
[void]$menu_AC_File.DropDownItems.Add($menu_AC_sessions)
 
# Menu Options - File / обратная связь
#$menuSaveAs.Image        = [System.IconExtractor]::Extract(&quot;shell32.dll&quot;, 45, $true)
#$menuSaveAs.ShortcutKeys = "Control, S"
$menu_AC_Feedback.Text = "Обратная связь"
$menu_AC_Feedback.Add_Click({
	Feedback_button
})
[void]$menu_AC_File.DropDownItems.Add($menu_AC_Feedback)
 
# Menu Options - File / о программе
#$menuExit.Image        = [System.IconExtractor]::Extract(&quot;shell32.dll&quot;, 10, $true)
#$menuExit.ShortcutKeys = "Control, X"
$menu_AC_about.Text = "О программе"
$menu_AC_about.Add_Click({
about_programm
})
[void]$menu_AC_File.DropDownItems.Add($menu_AC_about)

# Menu Options - File / Exit
#$menuExit.Image        = [System.IconExtractor]::Extract(&quot;shell32.dll&quot;, 10, $true)
#$menuExit.ShortcutKeys = "Control, X"
$menu_AC_exit.Text = "Выход"
$menu_AC_exit.Add_Click({$frmMain.Close()})
[void]$menu_AC_File.DropDownItems.Add($menu_AC_exit)
#------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------
citrix_pkg_tables
AppV_pkg_tables
SCCM_pkg_tables
$frmMain.ShowDialog();

$x
