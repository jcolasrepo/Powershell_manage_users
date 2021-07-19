#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#Projet: Scripting 
#
#auteurs :
#-Colas Jéremy SIO_1 2019/2020
#-Thomas Ghisleni SIO_1 2019/2020
#
#
#                            _________
#                           /         /.
#    .-------------.       /_________/ |
#   /             / |      |         | |
#  /+============+\ |      | |====|  | |
#  ||C:\>SIO     || |      |         | |
#  ||            || |      | |====|  | |
#  ||            || |      |   ___   | |
#  ||            || |      |  |404|  | |
#  ||            ||/@@@    |   ---   | |
#  \+============+/    @   |_________|./.
#                     @          ..  ....'
#  ..................@     __.'.'  ''
# /oooooooooooooooo//     ///
#/................//     /_/
#------------------
#
#
#but:
#gérer la création , la suppresion , la migration de groupe à goupe des utilisateurs d'un active directory windows
#
#contact :
#jeremy.colasdev@outlook.fr
#
#paypal:
#
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#form
#code de l'interface graphique
Add-Type -AssemblyName PresentationFramework
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="tj_project" Height="450" Width="800">
        <Window.Background>
		    <ImageBrush ></ImageBrush>
		</Window.Background>
    <Grid>
        <ListBox Name="listbox_etu_groupe" HorizontalAlignment="Left" Height="243" Margin="426,166,0,0" VerticalAlignment="Top" Width="356" SelectionMode="Multiple"/>
        <TabControl HorizontalAlignment="Left" Height="409" VerticalAlignment="Top" Width="421"  TabStripPlacement="Top" Opacity="0.75">
            <TabControl.Background>
                <ImageBrush/>
            </TabControl.Background>
            <TabItem Name="creation_etu" Header="ajouter un utilisateur">
                <Grid Background="#FFE5E5E5">
                    <Button Name="creation" Content="Valider" HorizontalAlignment="Left" Margin="10,218,0,0" VerticalAlignment="Top" Width="75"/>
                    <TextBox Name="textbox_prenom" HorizontalAlignment="Left" Height="22" Margin="10,46,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>
                    <TextBox Name="textbox_nom" HorizontalAlignment="Left" Height="22" Margin="10,119,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>
                    <ComboBox Name="Combobox_grp" HorizontalAlignment="Left" Margin="10,173,0,0" VerticalAlignment="Top" Width="120"/>
                    <ListBox Name="listbox_resultat" HorizontalAlignment="Left" Height="149" Margin="245,46,0,0" VerticalAlignment="Top" Width="160"/>
                    <Label Content="Prénom:" HorizontalAlignment="Left" Margin="10,20,0,0" VerticalAlignment="Top" Width="120"/>
                    <Label Content="Nom:" HorizontalAlignment="Left" Margin="10,88,0,0" VerticalAlignment="Top" Width="120"/>
                    <Label Content="Groupe:" HorizontalAlignment="Left" Margin="10,146,0,0" VerticalAlignment="Top" Width="75"/>
                    <Label Content="Resultat:" HorizontalAlignment="Left" Margin="245,20,0,0" VerticalAlignment="Top" Width="160"/>
                </Grid>
            </TabItem>
            <TabItem Name="supp_etu" Header="supp un/des utilisateur">

                <Grid Background="#FFE5E5E5">
                   <ListBox Name="listbox_etu_supp" HorizontalAlignment="Left" Height="222" Margin="10,108,0,0" VerticalAlignment="Top" Width="220" SelectionMode="Multiple" ItemsSource="{Binding SelectedItems , ElementName=listBox}"/>
                    <Button Name="valider_supp" Content="Supprimer" HorizontalAlignment="Left" Margin="282,198,0,0" VerticalAlignment="Top" Width="74"/>
                    <TextBox Name="textbox_nom1" HorizontalAlignment="Left" Height="23" Margin="10,80,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                    <Label Content="Recherche :" HorizontalAlignment="Left" Margin="10,54,0,0" VerticalAlignment="Top" Width="120"/>
                </Grid>
            </TabItem>
            <TabItem Header="migrer un/des utilisateur" >

                <Grid Background="#FFE5E5E5">
                    <ListBox Name="listbox_tout_etu" HorizontalAlignment="Left" Height="162" Margin="10,78,0,0" VerticalAlignment="Top" Width="120" SelectionMode="Multiple"/>
                    <ComboBox Name="combo_mig" HorizontalAlignment="Left" Margin="285,51,0,0" VerticalAlignment="Top" Width="120"/>
                    <ListBox Name="listbox_etu_group" HorizontalAlignment="Left" Height="162" Margin="285,78,0,0" VerticalAlignment="Top" Width="120"/>
                    <Button Name="valider" Content="Transfert" HorizontalAlignment="Left" Margin="170,148,0,0" VerticalAlignment="Top" Width="75"/>
                    <Label Content="Liste des Etudiant:" HorizontalAlignment="Left" Margin="10,20,0,0" VerticalAlignment="Top" RenderTransformOrigin="-1.974,0.5" Width="120"/>
                    <Label Content="Groupe:" HorizontalAlignment="Left" Margin="285,20,0,0" VerticalAlignment="Top"/>
                    <TextBox Name="text_recherche_migr" HorizontalAlignment="Left" Height="22" Margin="10,50,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>
                </Grid>
            </TabItem>
            <TabItem Header="creer un groupe d'utilisateurs">

                <Grid Background="#FFE5E5E5">
                    <Button Name="Parcourir" Content="Parcourir" HorizontalAlignment="Left" Margin="161,64,0,0" VerticalAlignment="Top" Width="75"/>
                    <ComboBox Name="comboBox_groupe" HorizontalAlignment="Left" Margin="10,115,0,0" VerticalAlignment="Top" Width="120"/>
                    <TextBox Name="textbox_chemin" HorizontalAlignment="Left" Height="24" Margin="10,64,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>
                    <Button Name="valider_creation_groupe" Content="Valider" HorizontalAlignment="Left" Margin="10,157,0,0" VerticalAlignment="Top" Width="75"/>
                    <Label Content="liste des utilisateurs(en .csv):" HorizontalAlignment="Left" Margin="10,33,0,0" VerticalAlignment="Top" Width="180"/>
                    <Label Content="Groupe:" HorizontalAlignment="Left" Margin="10,93,0,0" VerticalAlignment="Top" RenderTransformOrigin="-3.079,-0.942" Width="120"/>
                    <ListBox x:Name="listbox_affichage_csv" HorizontalAlignment="Left" Height="236" Margin="161,115,0,0" VerticalAlignment="Top" Width="244"/>
                </Grid>
            </TabItem>
            <TabItem Header="supp tout les users d'un groupe">

                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="49*"/>
                        <ColumnDefinition Width="367*"/>
                    </Grid.ColumnDefinitions>
                    <ListBox Name="listbox_etu" HorizontalAlignment="Left" Height="220" Margin="10,82,0,0" VerticalAlignment="Top" Width="160" Grid.ColumnSpan="2"/>
                    <Button Name="validation" Content="supprimer" HorizontalAlignment="Left" Margin="219.5,164,0,0" VerticalAlignment="Top" Width="74" Grid.Column="1"/>
                    <ComboBox Name="comboBox_grp_supp" HorizontalAlignment="Left" Margin="10,40,0,0" VerticalAlignment="Top" Width="160" Grid.ColumnSpan="2"/>
                     <ListBox Name="listbox_supression" Grid.Column="1" HorizontalAlignment="Left" Height="100" Margin="170,251,0,0" VerticalAlignment="Top" Width="186"/>
                     <Label Content="Groupe:" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Grid.ColumnSpan="2" Width="160"/>
                    <Label Content="Liste des etudiants du groupe:" HorizontalAlignment="Left" Margin="1,62,0,0" VerticalAlignment="Top" Grid.ColumnSpan="2" Width="169"/>
                    <Label Content="Resultat:" Grid.Column="1" HorizontalAlignment="Left" Margin="170,220,0,0" VerticalAlignment="Top" Width="186"/>
                </Grid>
            </TabItem>
            <TabItem Header="migrer tout un groupe d'utilisateurs">

                <Grid Background="#FFE5E5E5">
                    <ListBox Name="listbox_ancien_groupe" HorizontalAlignment="Left" Height="162" Margin="10,78,0,0" VerticalAlignment="Top" Width="120"/>
                    <ComboBox Name="combobox_1_migration" HorizontalAlignment="Left" Margin="10,51,0,0" VerticalAlignment="Top" Width="120"/>
                    <ComboBox Name="combobox_2_migration" HorizontalAlignment="Left" Margin="285,51,0,0" VerticalAlignment="Top" Width="120"/>
                    <ListBox Name="listbox_nouveau_groupe" HorizontalAlignment="Left" Height="162" Margin="285,78,0,0" VerticalAlignment="Top" Width="120"/>
                    <Button Name="valider_transfert" Content="Transfert" HorizontalAlignment="Left" Margin="170,148,0,0" VerticalAlignment="Top" Width="75"/>
                    <Label Content="Groupe à migrer:" HorizontalAlignment="Left" Margin="10,12,0,0" VerticalAlignment="Top" Width="120"/>
                    <Label Content="Groupe recevant:" HorizontalAlignment="Left" Margin="285,12,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.526,-0.288" Width="120"/>
                </Grid>
            </TabItem>
        </TabControl>
        <ListBox Name="listbox_info_serv" HorizontalAlignment="Left" Height="100" Margin="426,61,0,0" VerticalAlignment="Top" Width="356"/>
        <ComboBox Name="uo_combo" HorizontalAlignment="Left" Margin="423,23,0,0" VerticalAlignment="Top" Width="120"/>
        <Label Content="UO:" HorizontalAlignment="Left" Margin="443,-1,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="refresh_users" Content="Refresh" HorizontalAlignment="Left" Margin="583,7,0,0" VerticalAlignment="Top" Width="74"/>
        <TextBox Name="textbox_nom_srv" HorizontalAlignment="Left" Height="23" Margin="662,34,0,0" TextWrapping="Wrap" Text="WIN-2F6TT0VV10R" VerticalAlignment="Top" Width="120"/>
        <Button Name="button_connexion" Content="connexion" HorizontalAlignment="Left" Margin="583,34,0,0" VerticalAlignment="Top" Width="74"/>


    </Grid>
</Window>

'@

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
#on affiche le bon background
$Form=[Windows.Markup.XamlReader]::Load( $reader )
$Form.add_Loaded({
		try{$Form.Background.ImageSource = "$pwd\background.jpg"}catch{""}
        try{$Form.Icon = "$pwd\script.ico"}catch{""}
	})
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#declaration de variable
#on déclare les objets et, on leurs associe des variables pour pouvoir les utiliser.
#combobox
$combox_creation_etu=$Form.FindName("Combobox_grp")
$combox_migration=$Form.FindName("combo_mig")
$combox_create_grp_etu=$Form.FindName("comboBox_groupe")
$combox_supp_grp_etu=$Form.FindName("comboBox_grp_supp")
$combox_migr_grp_etu_left=$Form.FindName("combobox_1_migration")
$combox_migr_grp_etu_right=$Form.FindName("combobox_2_migration")
$uo_combo=$Form.FindName("uo_combo")
#listbox
$listbox_info=$Form.FindName("listbox_info_serv")
$listbox_etu_supp=$Form.FindName("listbox_etu_supp")
$listbox_etu=$Form.FindName("listbox_etu_groupe")
$listbox_groupe_new=$Form.FindName("listbox_etu_group")
$listbox_etu_AD=$Form.FindName("listbox_tout_etu")
$listbox_supp_etu_AD=$Form.FindName("listbox_etu")
$listbox_migr_etu_AD_old=$Form.FindName("listbox_ancien_groupe")
$listbox_migr_etu_AD_new=$Form.FindName("listbox_nouveau_groupe")
$listbox_supp_resultat=$Form.FindName("listbox_supression")
$listbox_affichage_csv=$Form.FindName("listbox_affichage_csv")
#textbox

#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#méthode 

#methode pour Refresh la listBox donnant les Etudiant de l'AD
function Refresh {
    $Form.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Background, [action]{
    [void] $listbox_etu.Items.Clear()
    $path2=$uo_combo.SelectedItem.ToString()
    $script:commande= Invoke-command -ComputerName $textbox_srv  -credential $Cred -ArgumentList $path2 -ScriptBlock {param($path2)Get-ADUser -Filter * -Searchbase $path2 | Select-Object SamAccountName}
    Foreach ($users in $commande){
        $Sam = $users.SamAccountName
        [void] $listbox_etu.Items.Add("$Sam")}
        })
        
}
#Methode pour Remplir les ListBox pour transférer un étudiant et pour supprimer un étudiant
function Recherche{
    $Form.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Background, [action]{
    [void] $listbox_etu_supp.Items.Clear()
    [void] $listbox_etu_AD.Items.Clear()
    $path3=$uo_combo.SelectedItem.ToString()
    $recherche=Invoke-command -ComputerName $textbox_srv  -credential $Cred -ArgumentList $path3 -ScriptBlock {param($path3) Get-ADUser -Filter * -Searchbase $path3 | Select-Object SamAccountName}
    Foreach ($users in $recherche){
    $Sam = $users.SamAccountName
    [void] $listbox_etu_supp.Items.Add("$Sam")
    [void] $listbox_etu_AD.Items.Add("$Sam")}
    })
    
}


#gestion des états pour  mettre à jour les listbox des pannels de migration
function Refresh_listbox{
        $Form.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Background, [action]{
        #on refresh la listbox
         [void] $listbox_groupe_new.Items.Clear()
         try{$nam_new_grp=$combox_migration.SelectedItem.ToString() 
         $recup_group=Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $nam_new_grp -ScriptBlock {param($nam_new_grp) Get-ADGroupMember -identity  $nam_new_grp | select SamAccountName}
         foreach ($users in $recup_group){
             $sam=$users.SamAccountName
             [void] $listbox_groupe_new.Items.Add("$sam")#commentaire
            }
            }catch{""}
        #on refresh la listbox de gauche du pannel "migrer tout un groupe"
        [void] $listbox_migr_etu_AD_old.Items.Clear()
        try{$nam_old_grp=$combox_migr_grp_etu_left.SelectedItem.ToString() 
        $recup_group=Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $nam_old_grp -ScriptBlock {param($nam_old_grp)Get-ADGroupMember -identity  $nam_old_grp | select SamAccountName}
        foreach ($users in $recup_group)
        {
            $sam=$users.SamAccountName
            [void] $listbox_migr_etu_AD_old.Items.Add("$sam")#commentaire
        }
        }catch{""}
       #on refresh la listbox de droite du pannel "migrer tout un groupe"
        [void] $listbox_migr_etu_AD_new.Items.Clear()
        try{$nam_new_grp=$combox_migr_grp_etu_right.SelectedItem.ToString()
        $recup_group1=Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $nam_new_grp -ScriptBlock {param($nam_new_grp)Get-ADGroupMember -identity  $nam_new_grp | select SamAccountName}
        foreach ($users in $recup_group1)
        {
            $sam=$users.SamAccountName
            [void] $listbox_migr_etu_AD_new.Items.Add("$sam")#commentaire
        }
        }catch{""} 
        #on refresh la listbox du pannel "supp tout un groupe"
        [void]$listbox_supp_etu_AD.Items.Clear()
        try{$nam_grp=$combox_supp_grp_etu.SelectedItem.ToString()
        $recup_group2=Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $nam_grp -ScriptBlock {param($nam_grp)Get-ADGroupMember -identity  $nam_grp | select SamAccountName}
        foreach ($users in $recup_group2)
        {
            $sam=$users.SamAccountName
            [void] $listbox_supp_etu_AD.Items.Add("$sam")#commentaire
        }
        }catch{""}
        })

}
 #New-Object -TypeName Windows.Application
# Fixes le problème de freze en metant les tâches dde mise à jour de l'interface graphique en arrière plan
function Update-Gui {
    # Basically WinForms Application.DoEvents()
    $Form.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Background, [action]{
    Recherche
    Refresh
    Refresh_listbox 
    Write-Host "background"
    })
}

#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#Connexion à distance
$textbox_nom_serveur=$Form.FindName("textbox_nom_srv")

 $Form.FindName("button_connexion").Add_Click( {
    $script:textbox_srv=$textbox_nom_serveur.text
    $script:Cred = Get-Credential
    clear
    $test=Invoke-command -ComputerName $textbox_srv  -credential $Cred -ScriptBlock {Hostname}
    if ($test -eq $textbox_srv){
    write-host "Connection à $test établie"
    [void] $uo_combo.Items.Clear()
     #remplie la combo du choix des UO
     $organizationUnitName=Invoke-command -ComputerName $textbox_srv  -credential $Cred -ScriptBlock {(Get-ADOrganizationalUnit -Filter *).DistinguishedName}
     Foreach ($UO in $organizationUnitName)
     {
     $uo_combo.Items.Add($UO)
     clear
     }

    #ajoute les informations de l'ordinateur distant
    #Information sur l'environment
    $domain= Invoke-command -ComputerName $textbox_srv  -credential $Cred -ScriptBlock {Get-ADDomain -Current LocalComputer} #nom de l'ordinateur
    $nom_computer= Invoke-command -ComputerName $textbox_srv  -credential $Cred -ScriptBlock {(Get-ADDomainController -Filter * | Select Name).name} #nom netbios du domaine
    try{$script:path=$uo_combo.SelectedItem.ToString()}catch{""}
    $script:hour = Get-Date -Format "HH:mm"  #on récupe l'heure
    $script:date = Get-Date -Format "dd/MM/yyyy" #on récupère la date sous la forme "Jour/Mois/Année"
    $script:userName= [Environment]::UserName 
    [void] $listbox_info.Items.Add("Nom de domaine : $domain") #non netbios du domaine
    [void] $listbox_info.Items.Add("Nom de l'ordinateur: $nom_computer") #nom de l'ordinateur
    [void] $listbox_info.Items.Add("Nom de l'utilisateur: $userName")
    [void] $listbox_info.Items.Add("`n`nHorodatage: $hour $date") #date

     
    #on affiche les informations du premier UO trouvé sur l'AD(généralement l'UO "domain Controller" présent par defaut)
    $uo_combo.SelectedIndex = 1
    If ($uo_combo.text -ne $null) {
        [void] $combox_creation_etu.Items.Clear()
        [void] $combox_migration.Items.Clear()
        [void] $combox_create_grp_etu.Items.Clear()
        [void] $combox_supp_grp_etu.Items.Clear()
        [void] $combox_migr_grp_etu_left.Items.Clear()
        [void] $combox_migr_grp_etu_right.Items.Clear()
         $uo_selected=$uo_combo.SelectedItem.ToString()
        $script:ListeGr = Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $uo_selected -ScriptBlock {param($uo_selected) (Get-ADGroup -Properties * -Filter * -SearchBase "$uo_selected").name}
         Foreach ($Usergr in $ListeGr)
        { 
        $combox_creation_etu.Items.Add($Usergr)
        $combox_migration.Items.Add($Usergr)
        $combox_create_grp_etu.Items.Add($Usergr)
        $combox_supp_grp_etu.Items.Add($Usergr)
        $combox_migr_grp_etu_left.Items.Add($Usergr)
        $combox_migr_grp_etu_right.Items.Add($Usergr)
        }
        }
        #Update-Gui
    Refresh
    Recherche
    ADD-content -path $fichier -value "----------------------------------------------------------------------------------------------------`r`nscript lancé le $date à $hour par $userName sur $textbox_srv"


    }
    else 
    {
    write-host "La connection à $textbox_srv n'a pas pu être établie"
    }
})
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
 
 #creation du fichier de log s'il ne l'ai pas
 $script:fichier = "$pwd\log.txt"
 If ((Test-Path $fichier) -eq $True) 
 {
 ADD-content -path $fichier -value "----------------------------------------------------------------------------------------------------`r`nscript exécuter"
 }
else{ADD-content -path $fichier -value "----------------------------------------------------------------------------------------------------`r`nfichier créer"}
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#button Refresh recharge des éléments en cas de bug
 $Form.FindName("refresh_users").Add_Click( {
    Refresh
    Recherche
    Refresh_listbox
    $uo_combo.Add_SelectionChanged($uo_comboIndexChanged)#}
#permets de modifier les groupes des combobox après chaque changement d'UO
$uo_comboIndexChanged=
{
   If ($uo_combo.text -ne $null) {
        [void] $combox_creation_etu.Items.Clear()
        [void] $combox_migration.Items.Clear()
        [void] $combox_create_grp_etu.Items.Clear()
        [void] $combox_supp_grp_etu.Items.Clear()
        [void] $combox_migr_grp_etu_left.Items.Clear()
        [void] $combox_migr_grp_etu_right.Items.Clear()
         $uo_selected=$uo_combo.SelectedItem.ToString()
        $script:ListeGr = Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $uo_selected -ScriptBlock {param($uo_selected) (Get-ADGroup -Properties * -Filter * -SearchBase "$uo_selected").name}
         Foreach ($Usergr in $ListeGr)
        { 
        $combox_creation_etu.Items.Add($Usergr)
        $combox_migration.Items.Add($Usergr)
        $combox_create_grp_etu.Items.Add($Usergr)
        $combox_supp_grp_etu.Items.Add($Usergr)
        $combox_migr_grp_etu_left.Items.Add($Usergr)
        $combox_migr_grp_etu_right.Items.Add($Usergr)
    
        }
    clear
    Refresh
    Recherche
    
   }
}


 })
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#Pannel ,creation d'un etudiant
#variable uniquement pour ce label
$textbox_prenom=$Form.FindName("textbox_prenom")
$textbox_nom=$Form.FindName("textbox_nom")
$listbox_retour=$Form.FindName("listbox_resultat")

#code du boutton
 $Form.FindName("creation").Add_Click( {
    $etu_prenom=$textbox_prenom.text #on récupe le contenu du textbox
    $etu_nom=$textbox_nom.text #on récupe le contenu du textbox
    $nom_complet=$etu_nom + " " +$etu_prenom
    $session=$etu_nom.ToUpper()
     try{
        $nam_grp=$combox_creation_etu.SelectedItem.ToString() #on récupe le contenu du combobox
     }
    catch {""}
    if($etu_prenom.Length -lt 3 -or $etu_nom.Length -lt 3-or $nam_grp -eq $null)#si une DATA n'est pas remplis on ne crée pas l'Etudiant
    {
        [void] $listbox_retour.Items.Add("UNE DATA n'a pas été remplis , merci de la remplir,ou le nom est trop court")#commentaire
        
    }
    else{
        if([bool] (Invoke-command -ComputerName $textbox_srv  -credential $Cred  -ScriptBlock {Get-ADUser -Filter {sAMAccountName -eq  $using:session}}) -eq 1){
            [void] $listbox_retour.Items.Add("Le nom d'ouverture de session est déjà utilisé ")#commentaire
        }
        else{
            $path1=$uo_combo.SelectedItem.ToString()
            $password = ConvertTo-SecureString "P@ssword" -AsPlainText -Force #on set le mdp par défault , il devra etre modifié à la première connection 
            Invoke-command -ComputerName $textbox_srv -credential $Cred -ScriptBlock {New-ADUser -Name $using:nom_complet -GivenName $using:etu_prenom -DisplayName $using:nom_complet -Surname $using:nom_complet -SamAccountName $using:session -UserPrincipalName $using:session -Enabled $true -AccountPassword $using:password -ChangePasswordAtLogon $True -Path $using:path1 }
            try{Invoke-command -ComputerName $textbox_srv  -credential $Cred -ScriptBlock {Add-ADGroupMember -Identity  $using:nam_grp -Members $using:session -PassThru }}catch{""}
            [void] $listbox_retour.Items.Add("$nom_complet ajouté à l'AD")#commentaire
            [void] $listbox_retour.Items.Add("$nom_complet ajouté au groupe $groupe")#commentaire
            ADD-content -path $fichier -value "$session ajouté à l'AD le $date à $hour" #ajout dans le fichier de log 
            [void] $listbox_etu.Items.Add("$session groupe : $nam_grp ")
            #Refresh 
            #Recherche 
            #Refresh_listbox 
            #Update-Gui
            }
        }
 })
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#pannel,supprimer un étudiant
#variable
$textbox_nom_etu=$Form.FindName("textbox_nom1")

#button supprimer l'utilisateur
 $Form.FindName("valider_supp").Add_Click( {
    $nom_etudiant = $listbox_etu_supp.SelectedItems
    if ( $nom_etudiant -ne $null){
    $msgBoxInput =  [System.Windows.MessageBox]::Show('Voulez-vous vraiment supprimer cet étudiant ?','Attention !','YesNo','Warning') #lance une popup de confirmation
  switch  ($msgBoxInput) {

  'Yes' {
      
        
        Foreach ($users in $nom_etudiant)
        {
        $recherche_etudiant_supp=Invoke-command -ComputerName $textbox_srv  -credential $Cred -ArgumentList $users -ScriptBlock {param($users)Get-ADUser -Filter {samaccountname -like $users} -Properties samaccountname | Select-Object samaccountname}
        $sam=$recherche_etudiant_supp.samaccountname
        Invoke-command -ComputerName $textbox_srv  -credential $Cred -ArgumentList $sam -ScriptBlock {param($sam)Remove-ADUser $sam -Confirm:$false} #on supp l'étudiant
        ADD-content -path $fichier -value "$nom_etudiant supp de l'AD le $date à $hour" #ajout dans le fichier de log 
        }
         Refresh_listbox
         Refresh
         Recherche
        }

  'No' {}

  }
  }
 })
 [void] $listbox_etu.Items.Add("test1")
 [void] $listbox_etu.Items.Add("glouglou")
 ################ MUST CREATE BEFORE ASSIGN ################
try{$textbox_nom_etu.Add_SelectionChanged($textbox_nom_etu_SelectedIndexChanged)}catch{""}

#gestion d'état qui recherche les utilisateurs par rapport au contenu de la textbox
$textbox_nom_etu_SelectedIndexChanged=
{
$nom_recherche_etudiant = $textbox_nom_etu.text
 [void] $listbox_etu_supp.Items.Clear()
   if($textbox_nom_etu.text -ne $null) {
   [void] $listbox_etu_supp.Items.Clear()
    foreach($item in $listbox_etu.Items){
    if ($item -match $nom_recherche_etudiant){
    [void] $listbox_etu_supp.Items.Add("$item")}
        }
   }
 
}
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#pannel,Migrer un étudiant
#variable
$text_recherche_migr=$Form.FindName("text_recherche_migr")

#code du bouton 
 $Form.FindName("valider").Add_Click( {
    try{$nom_etu_migr=$listbox_etu_AD.SelectedItems}catch{""}
    try{$nam_new_grp=$combox_migration.SelectedItem.ToString()}catch{""}
    if ($nom_etu_migr -ne $null -and $nam_new_grp -ne $null){
     foreach($group in $ListeGr){
        Invoke-command -ComputerName "WIN-2F6TT0VV10R"  -credential $Cred -ScriptBlock {Remove-ADGroupMember $using:group -Members $using:nom_etu_migr -Confirm:$false}
     }
        Invoke-command -ComputerName "WIN-2F6TT0VV10R"  -credential $Cred -ScriptBlock {Add-ADGroupMember -Identity  $using:nam_new_grp -Members $using:nom_etu_migr -PassThru}
        ADD-content -path $fichier -value "$nom_etu_migr ajouté dans le groupe $nam_new_grp le $date à $hour" #ajout dans le fichier de log 
        Refresh
        Recherche
        Refresh_listbox
 }
  })

   ################ MUST CREATE BEFORE ASSIGN ################
try{$text_recherche_migr.Add_SelectionChanged($text_recherche_migr_SelectedIndexChanged)}catch{""}

#gesion d'état qui recherche les utilisateurs par rapport au contenu de la textbox
$text_recherche_migr_SelectedIndexChanged=
{
$nom_recherche_etudiant_migr = $text_recherche_migr.text
 $nom_contenant_etu_mgr="*$nom_recherche_etudiant_migr*"
 [void] $listbox_etu_AD.Items.Clear()
   if(![string]::IsNullOrEmpty($text_recherche_migr.text)) {
   [void] $listbox_etu_AD.Items.Clear()
    $recherche_fini=Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $nom_contenant_etu_mgr -ScriptBlock {param($nom_contenant_etu_mgr)Get-aduser –filter { SamAccountName -like $nom_contenant_etu_mgr } |select SamAccountName}
    Foreach ($users in $recherche_fini){
    $Sam = $users.SamAccountName
    [void] $listbox_etu_AD.Items.Add("$Sam")
        }
   }
   if([string]::IsNullOrEmpty($text_recherche_migr.text))
   {
   [void] $listbox_etu_AD.Items.Clear()
   $recherche= Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $path -ScriptBlock {param($path)Get-ADUser -Filter * -Searchbase $path | Select-Object SamAccountName}
    Foreach ($users in $recherche)
      {
    $Sam = $users.SamAccountName
    [void] $listbox_etu_AD.Items.Add("$Sam")
      }
    }
    }

#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#pannel,ajouter un groupe d'étudiant
#variable
$textbox_chemin=$Form.FindName("textbox_chemin")
#button parcourir
#affiche la page pour parcourir ces dossiers
$Form.FindName("Parcourir").Add_Click( { 
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $objForm = New-Object System.Windows.Forms.OpenFileDialog
    $objForm.InitialDirectory = "c:\"
    $objForm.Title = "Selectionner un fichier :"
    $objForm.FilterIndex = 3
    $Show = $objForm.ShowDialog()
 
    If ($Show -eq "Cancel")
    {
        "Annulé par l'utilisateur"
    }
    Else
    {
        $chemin=$objForm.FileName
        $textbox_chemin.text =$chemin
    }
})
#button Valider
$Form.FindName("valider_creation_groupe").Add_Click( { 
    try{$group=$combox_create_grp_etu.SelectedItem.ToString()}catch{""}
    $recup_chemin = $textbox_chemin.text
    try {$users=import-csv -path $recup_chemin -delimiter ";"}catch{""}
    if($group -ne $Null -and $users -ne $null){ #si un groupe est choisi et un chemin 
         foreach ($user in $users){
    
        $nom_complet = $user.firstName + " " + $user.lastName
        $prenom = $user.firstName
        $nom = $user.lastName
        $password = ConvertTo-SecureString "P@ssword" -AsPlainText -Force
        $session=$nom.ToUpper()
    
        try {
                Invoke-command -ComputerName "WIN-2F6TT0VV10R"  -credential $Cred -ScriptBlock {New-ADUser -Name $using:nom_complet -GivenName $using:prenom -DisplayName $using:nom_complet  -Surname $using:nom -SamAccountName $using:session -UserPrincipalName $using:session -AccountPassword $using:password -ChangePasswordAtLogon $True -Path $using:path -Enabled $true}
                Invoke-command -ComputerName "WIN-2F6TT0VV10R"  -credential $Cred -ScriptBlock {Add-ADGroupMember -Identity  $using:group -Members $using:nom -PassThru}
                 ADD-content -path $fichier -value "$session ajouté à l'AD le $date à $hour" #ajout dans le fichier de log
                 [void] $listbox_affichage_csv.Items.Add("Bienvenue $session !") #écris dans la listbox du pannel que l'utilisateurs est crée
            } 
        catch{""}   
    }
     Refresh
     Recherche
     Refresh_listbox
                }
})
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#pannel ,supprimer des étudiants

#button supprimer
$Form.FindName("validation").Add_Click( { 
    
    try{$myname_supp=$combox_supp_grp_etu.SelectedItem.ToString()}
    catch{""}
    if ($myname_supp -ne $null)
    {
    $msgBoxInput =  [System.Windows.MessageBox]::Show('Voulez-vous vraiment supprimer tous les membres du groupe ?','Attention !','YesNo','Warning')
    switch  ($msgBoxInput) {

        'Yes' {

            try{$users=Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $myname_supp -ScriptBlock {param($myname_supp)Get-ADGroupMember -identity $myname_supp | select SamAccountName}}catch{""}
            foreach ($user in $users){
                $sam=$user.SamAccountName
                try{Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $sam -ScriptBlock {param($sam)Remove-ADUser $sam -Confirm:$false}}catch{""} #on supp l'étudiant
                [void] $listbox_supp_resultat.Items.Add("bon voyage $sam !")#commentaire     
                ADD-content -path $fichier -value "$sam supp de l'AD le $date à $hour" #ajout dans le fichier de log       
            }

         Refresh_listbox
         Refresh
         Recherche

  }

  'No' {

  [void] $listbox_supp_resultat.Items.Add("Pas de changement dans l'AD !")#commentaire

  }

  }
     
    }
    else
    {
    [void] $listbox_supp_resultat.Items.Add("Merci de choisir un groupe")#commentaire
    }  


})
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#pannel migration de tous les membres d'un groupe

#button transfert
$Form.FindName("valider_transfert").Add_Click( { 

    try{$old_groupe=$combox_migr_grp_etu_left.SelectedItem.ToString()}
        catch{""}
    try{$new_groupe=$combox_migr_grp_etu_right.SelectedItem.ToString()}
        catch{""}
    if($old_groupe -eq $Null -or $new_groupe -eq $null)
    {
        
    }
    else{
        $users = Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $old_groupe -ScriptBlock {param($old_groupe)Get-ADGroupMember -identity $old_groupe | select SamAccountName}
        foreach ($user in $users){
        $sam=$user.SamAccountName
                try{Invoke-command -ComputerName $textbox_srv -credential $Cred -ScriptBlock {remove-ADGroupMember -Identity  $using:old_groupe -Member $using:sam -Confirm:$false}}catch{""} #on retire l'étudiant de l'ancien groupe
                Invoke-command -ComputerName $textbox_srv -credential $Cred  -ScriptBlock {Add-ADGroupMember -Identity  $using:new_groupe -Members $using:sam -PassThru}
                ADD-content -path $fichier -value "$nom_etu_migr ajouté dans le groupe $new_groupe le $date à $hour" #ajout dans le fichier de log 
                
               
        }
         Refresh_listbox
         Refresh
         Recherche

     }
})
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#code pour la gestion des états

#gestion des états pour afficher les membres du groupe pour la migration
$combox_migration_SelectedIndexChanged=
{
   If ($combox_migration.text -ne $null) {
        [void] $listbox_groupe_new.Items.Clear()
        $nam_new_grp=$combox_migration.SelectedItem.ToString() 
        $recup_group=Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $nam_new_grp -ScriptBlock {param($nam_new_grp)Get-ADGroupMember -identity  $nam_new_grp | select SamAccountName}
        foreach ($users in $recup_group){
        $sam=$users.SamAccountName
             [void] $listbox_groupe_new.Items.Add("$sam")#commentaire
        }
   }
}
################ MUST CREATE BEFORE ASSIGN ################
$combox_migration.Add_SelectionChanged($combox_migration_SelectedIndexChanged)

#gestion des états pour afficher les membres du groupe avant leur suppression
$combox_supp_grp_etu_SelectedIndexChanged=
{
   If ($combox_migration.text -ne $null) {
        [void] $listbox_supp_etu_AD.Items.Clear()
        $nam_new_grp=$combox_supp_grp_etu.SelectedItem.ToString() 
        $recup_group=Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $nam_new_grp -ScriptBlock {param($nam_new_grp)Get-ADGroupMember -identity  $nam_new_grp | select SamAccountName}
        foreach ($users in $recup_group){
        $sam=$users.SamAccountName
             [void] $listbox_supp_etu_AD.Items.Add("$sam")#commentaire
        }
   }
}
################ MUST CREATE BEFORE ASSIGN ################
$combox_supp_grp_etu.Add_SelectionChanged($combox_supp_grp_etu_SelectedIndexChanged) 


#gestion des états pour afficher les membres du groupe envoyant les migrants
$combox_migr_grp_etu_left_SelectedIndexChanged=
{
   If ($combox_migr_grp_etu_left.text -ne $null) {
        [void] $listbox_migr_etu_AD_old.Items.Clear()
        $nam_new_grp=$combox_migr_grp_etu_left.SelectedItem.ToString() 
        $recup_group=Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $nam_new_grp -ScriptBlock {param($nam_new_grp)Get-ADGroupMember -identity  $nam_new_grp | select SamAccountName}
        foreach ($users in $recup_group){
            $sam=$users.SamAccountName
             [void] $listbox_migr_etu_AD_old.Items.Add("$sam")#commentaire
        }
   }
}
################ MUST CREATE BEFORE ASSIGN ################
$combox_migr_grp_etu_left.Add_SelectionChanged($combox_migr_grp_etu_left_SelectedIndexChanged)

#gestion des états pour afficher les membres du groupe reçevant les migrants
$combox_migr_grp_etu_right_SelectedIndexChanged=
{
   If ($combox_migr_grp_etu_right.text -ne $null) {
        [void] $listbox_migr_etu_AD_new.Items.Clear()
        $nam_new_grp=$combox_migr_grp_etu_right.SelectedItem.ToString() 
        $recup_group=Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $nam_new_grp -ScriptBlock {param($nam_new_grp)Get-ADGroupMember -identity  $nam_new_grp | select SamAccountName}
        foreach ($users in $recup_group){
             $sam=$users.SamAccountName
             [void] $listbox_migr_etu_AD_new.Items.Add("$sam")#commentaire
        }
   }
}
     ################ MUST CREATE BEFORE ASSIGN ################
     if ($uo_combo.selecteditem -ne $null){
$uo_combo.Add_SelectionChanged($uo_comboIndexChanged)}
#gestion des états pour changer les groupes des combobox après chaque changement d'UO
$uo_comboIndexChanged=
{
   If ($uo_combo.text -ne $null) {
        [void] $combox_creation_etu.Items.Clear()
        [void] $combox_migration.Items.Clear()
        [void] $combox_create_grp_etu.Items.Clear()
        [void] $combox_supp_grp_etu.Items.Clear()
        [void] $combox_migr_grp_etu_left.Items.Clear()
        [void] $combox_migr_grp_etu_right.Items.Clear()
         $uo_selected=$uo_combo.SelectedItem.ToString()
        $script:ListeGr = Invoke-command -ComputerName $textbox_srv -credential $Cred -ArgumentList $uo_selected -ScriptBlock {param($uo_selected) (Get-ADGroup -Properties * -Filter * -SearchBase "$uo_selected").name}
         Foreach ($Usergr in $ListeGr)
        { 
        $combox_creation_etu.Items.Add($Usergr)
        $combox_migration.Items.Add($Usergr)
        $combox_create_grp_etu.Items.Add($Usergr)
        $combox_supp_grp_etu.Items.Add($Usergr)
        $combox_migr_grp_etu_left.Items.Add($Usergr)
        $combox_migr_grp_etu_right.Items.Add($Usergr)
    
        }
    clear
    Refresh
    Recherche
    
   }
}
################ MUST CREATE BEFORE ASSIGN ################
$combox_migr_grp_etu_right.Add_SelectionChanged($combox_migr_grp_etu_right_SelectedIndexChanged)


#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#afficher l'interface graphique
$Form.ShowDialog() | out-null