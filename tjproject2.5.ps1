#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#Projet: Scripting 
#
#auteur :
#-Colas Jérémy SIO_1 2019/2020
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
#
#
#
#
#form creation
#
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
                                        #form

Add-Type -AssemblyName PresentationFramework
#première form , page principale
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="tj_project" Height="450" Width="800">
        <Window.Background>
		    <ImageBrush ></ImageBrush>
		</Window.Background>
    <Grid>
        <ListBox Name="listbox_etu_groupe" HorizontalAlignment="Left" Height="243" Margin="426,166,0,0" VerticalAlignment="Top" Width="356"/>
        <TabControl HorizontalAlignment="Left" Height="409" VerticalAlignment="Top" Width="421"  TabStripPlacement="Top" Opacity="0.75">
            <TabControl.Background>
                <ImageBrush/>
            </TabControl.Background>
            <TabItem Name="creation_etu" Header="creation etudiant">
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
            <TabItem Name="supp_etu" Header="supp etudiant">

                <Grid Background="#FFE5E5E5">
                    <ListBox Name="listbox_etu_supp" HorizontalAlignment="Left" Height="222" Margin="10,10,0,0" VerticalAlignment="Top" Width="220"/>
                    <Button Name="valider_supp" Content="Supprimer" HorizontalAlignment="Left" Margin="290,108,0,0" VerticalAlignment="Top" Width="74"/>
                </Grid>
            </TabItem>
            <TabItem Header="migrer un etudiant" >

                <Grid Background="#FFE5E5E5">
                    <ListBox Name="listbox_tout_etu" HorizontalAlignment="Left" Height="162" Margin="10,78,0,0" VerticalAlignment="Top" Width="120"/>
                    <ComboBox Name="combo_mig" HorizontalAlignment="Left" Margin="285,51,0,0" VerticalAlignment="Top" Width="120"/>
                    <ListBox Name="listbox_etu_group" HorizontalAlignment="Left" Height="162" Margin="285,78,0,0" VerticalAlignment="Top" Width="120"/>
                    <Button Name="valider" Content="Transfert" HorizontalAlignment="Left" Margin="170,148,0,0" VerticalAlignment="Top" Width="75"/>
                    <Label Content="Liste des Etudiant:" HorizontalAlignment="Left" Margin="10,52,0,0" VerticalAlignment="Top" RenderTransformOrigin="-1.974,0.5" Width="120"/>
                    <Label Content="Groupe:" HorizontalAlignment="Left" Margin="285,20,0,0" VerticalAlignment="Top"/>
                </Grid>
            </TabItem>
            <TabItem Header="creer un groupe d'étudiant">

                <Grid Background="#FFE5E5E5">
                    <Button Name="Parcourir" Content="Parcourir" HorizontalAlignment="Left" Margin="161,64,0,0" VerticalAlignment="Top" Width="75"/>
                    <ComboBox Name="comboBox_groupe" HorizontalAlignment="Left" Margin="10,115,0,0" VerticalAlignment="Top" Width="120"/>
                    <TextBox Name="textbox_chemin" HorizontalAlignment="Left" Height="24" Margin="10,64,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>
                    <Button Name="valider_creation_groupe" Content="Valider" HorizontalAlignment="Left" Margin="10,157,0,0" VerticalAlignment="Top" Width="75"/>
                    <Label Content="Chemin:" HorizontalAlignment="Left" Margin="10,33,0,0" VerticalAlignment="Top" Width="120"/>
                    <Label Content="Groupe:" HorizontalAlignment="Left" Margin="10,93,0,0" VerticalAlignment="Top" RenderTransformOrigin="-3.079,-0.942" Width="120"/>
                </Grid>
            </TabItem>
            <TabItem Header="supp des étudiant">

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
            <TabItem Header="migrer un groupe">

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
        <Button Name="refresh_users" Content="Refresh" HorizontalAlignment="Left" Margin="568,26,0,0" VerticalAlignment="Top" Width="74"/>

    </Grid>
</Window>

'@

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
#on load le bon background
$Form=[Windows.Markup.XamlReader]::Load( $reader )
$Form.add_Loaded({
		$Form.Background.ImageSource = "$pwd\background.jpg"
        $Form.Icon = "$pwd\script.ico"
	})
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#declaration de variable

#combobox
$combox_creation_etu=$Form.FindName("Combobox_grp")
$combox_migration=$Form.FindName("combo_mig")
$combox_create_grp_etu=$Form.FindName("comboBox_groupe")
$combox_supp_grp_etu=$Form.FindName("comboBox_grp_supp")
$combox_migr_grp_etu_left=$Form.FindName("combobox_1_migration")
$combox_migr_grp_etu_right=$Form.FindName("combobox_2_migration")
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
#textbox

#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#méthode 

#methode pour refresh la listBox donnant les Etudiant de l'AD
function Refresh {
    import-module activedirectory
    [void] $listbox_etu.Items.Clear()
    $commande= Get-ADUser -Filter * -Searchbase $path | Select-Object SamAccountName
    Foreach ($users in $commande){
    import-module activedirectory
        $Sam = $users.SamAccountName
        $groupe_etu_nom=( ( Get-ADUser $Sam -Property MemberOf ).MemberOf | Get-AdGroup ).Name 
        [void] $listbox_etu.Items.Add("$Sam groupe : $groupe_etu_nom ")}


}
#Remplir les ListBox pour transférer un étudiant et pour supprimer un étudiant
function Recherche{
    import-module activedirectory
    [void] $listbox_etu_supp.Items.Clear()
    [void] $listbox_etu_AD.Items.Clear()
    $recherche= Get-ADUser -Filter * -Searchbase $path | Select-Object SamAccountName
    Foreach ($users in $recherche){
    $Sam = $users.SamAccountName
    [void] $listbox_etu_supp.Items.Add("$Sam")
    [void] $listbox_etu_AD.Items.Add("$Sam")}
}


#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#Démarage du script 
#remplissage des ComboBox
$ListeGr = "RH"  ,"DSI" 
 Foreach ($Usergr in $ListeGr)
    { $combox_creation_etu.Items.Add($Usergr)
    $combox_migration.Items.Add($Usergr)
    $combox_create_grp_etu.Items.Add($Usergr)
    $combox_supp_grp_etu.Items.Add($Usergr)
    $combox_migr_grp_etu_left.Items.Add($Usergr)
    $combox_migr_grp_etu_right.Items.Add($Usergr)
    clear
    }

    #Information sur l'environment
    import-module activedirectory
    $domain= Get-ADDomain -Current LocalComputer #nom de l'ordinateur
    $nom_computer= (Get-ADDomainController -Filter * | Select Name).name #nom netbios du domaine
    $path= "ou=Utilisateurs," + "" + $domain #path ou les modifications auront lieu
    $hour = Get-Date -Format "HH:mm"  #on récupe l'heure
    $date = Get-Date -Format "dd/MM/yyyy" #on récupère la date sous la forme "Jour/Mois/Année"
    $domainName=[System.Net.Dns]::GetHostByName($VM).Hostname.split('.')
    $domainName=$domainName[1]+'.'+$domainName[2]
    $userName= [Environment]::UserName
    $organizationUnitName= (Get-ADOrganizationalUnit -Filter ’Name -like "Utilisateurs"’).Name 
    [void] $listbox_info.Items.Add("Nom de domaine : $domainName") #non netbios du domaine
    [void] $listbox_info.Items.Add("Nom de l'ordinateur: $nom_computer") #nom de l'ordinateur
    [void] $listbox_info.Items.Add("Nom de l'utilisateur: $userName")
    [void] $listbox_info.Items.Add("OU selectionné: $organizationUnitName")
    [void] $listbox_info.Items.Add("`n`nHorodatage: $hour $date") #date

 Recherche
 Refresh
 
 #creation du fichier de log s'il ne l'ai pas
 $fichier = "$pwd\log.txt"
 If ((Test-Path $fichier) -eq $True) 
 {
 ADD-content -path $fichier -value "----------------------------------------------------------------------------------------------------`r`nscript lancé le $date à $hour par $userName"
 }
else{ADD-content -path $fichier -value "----------------------------------------------------------------------------------------------------`r`nscript lancé le $date à $hour par $userName"}
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#button Refresh les utilisateurs de l'AD
 $Form.FindName("refresh_users").Add_Click( {
    Refresh
    Recherche
    Refresh_listbox


 })
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#Pannel ,creation d'un etudiant
#variable
$textbox_prenom=$Form.FindName("textbox_prenom")
$textbox_nom=$Form.FindName("textbox_nom")
$listbox_retour=$Form.FindName("listbox_resultat")
#code button
 $Form.FindName("creation").Add_Click( {
    import-module activedirectory
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
        [void] $listbox_retour.Items.Add("UNE DATA n'a pas été remplis , merci de la remplir")#commentaire
        
    }
    else{
        import-module activedirectory
        if([bool] (Get-ADUser -Filter {sAMAccountName -eq  $session}) -eq 1){
            [void] $listbox_retour.Items.Add("Le nom d'ouverture de session est déjà utilisé ")#commentaire
        }
        else{
            import-module activedirectory
            $password = ConvertTo-SecureString "P@ssword" -AsPlainText -Force #on set le mdp par défault , il devra etre modifié à la première connection 
            New-ADUser -Name $nom_complet -GivenName $etu_prenom -DisplayName $nom_complet -Surname $nom_complet -SamAccountName $session -UserPrincipalName $session -Enabled $true -AccountPassword $password -ChangePasswordAtLogon $True -Path $path 
            try{Add-ADGroupMember -Identity  $nam_grp -Members $session -PassThru }catch{""}
            [void] $listbox_retour.Items.Add("$nom_complet ajouté à l'AD")#commentaire
            [void] $listbox_retour.Items.Add("$nom_complet ajouté au groupe $groupe")#commentaire
            ADD-content -path $fichier -value "$session ajouté à l'AD le $date à $hour" #ajout dans le fichier de log 
            Refresh
            Recherche
            Refresh_listbox
            }
        }
 })
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#pannel,supprimer un étudiant
#variable


#button supprimer
 $Form.FindName("valider_supp").Add_Click( {
    $nom_etudiant = $listbox_etu_supp.SelectedItem
    if ( $nom_etudiant -ne $null){
    $msgBoxInput =  [System.Windows.MessageBox]::Show('Voulez-vous vraiment supprimer cet étudiant ?','Attention !','YesNo','Warning') #lance une popup de confirmation
  switch  ($msgBoxInput) {

  'Yes' {
      
        import-module activedirectory
        Remove-ADUser $nom_etudiant -Confirm:$false #on supp l'étudiant
        ADD-content -path $fichier -value "$nom_etudiant supp de l'AD le $date à $hour" #ajout dans le fichier de log 
         Refresh_listbox
         Refresh
         Recherche
        }

  'No' {}

  }
  }
 })
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#pannel,Migrer un étudiant
#variable

#code du bouton Refresh
#code du bouton Transferer
 $Form.FindName("valider").Add_Click( {
    try{$nom_etu_migr=$listbox_etu_AD.SelectedItem}catch{""}
    try{$nam_new_grp=$combox_migration.SelectedItem.ToString()}catch{""}
    if ($nom_etu_migr -ne $null -and $nam_new_grp -ne $null){
     foreach($group in $ListeGr){
        import-module activedirectory
        try{Remove-ADGroupMember $group -Members $nom_etu_migr -Confirm:$false}catch{""}
     }
        import-module activedirectory
        Add-ADGroupMember -Identity  $nam_new_grp -Members $nom_etu_migr -PassThru
        ADD-content -path $fichier -value "$nom_etu_migr ajouté dans le groupe $nam_new_grp le $date à $hour" #ajout dans le fichier de log 
        Refresh
        Recherche
        Refresh_listbox
 }
  })
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#pannel,ajouter un groupe d'étudiant
#variable
$textbox_chemin=$Form.FindName("textbox_chemin")
#button parcourir
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
                import-module activedirectory
                New-ADUser -Name $nom_complet -GivenName $prenom -DisplayName $nom_complet  -Surname $nom -SamAccountName $session -UserPrincipalName $session -AccountPassword $password -ChangePasswordAtLogon $True -Path $path -Enabled $true
                Add-ADGroupMember -Identity  $group -Members $nom -PassThru
                 ADD-content -path $fichier -value "$session ajouté à l'AD le $date à $hour" #ajout dans le fichier de log
                Refresh
                Recherche
                Refresh_listbox
            } 
        catch{""}   
    }}
})
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#pannel ,supp des étudiant

#button supprimer
$Form.FindName("validation").Add_Click( { 
    
    try{$myname_supp=$combox_supp_grp_etu.SelectedItem.ToString()}
    catch{""}
    if ($myname_supp -ne $null)
    {
    $msgBoxInput =  [System.Windows.MessageBox]::Show('Voulez-vous vraiment supprimer tous les membres du groupe ?','Attention !','YesNo','Warning')
    switch  ($msgBoxInput) {

        'Yes' {

            import-module activedirectory
            #$myname_supp=$combox_supp_grp_etu.SelectedItem.ToString()
            try{$users=Get-ADGroupMember -identity $myname_supp | select SamAccountName}catch{""}
            foreach ($user in $users){
                $sam=$user.SamAccountName
                import-module activedirectory
                try{Remove-ADUser $sam -Confirm:$false}catch{""} #on supp l'étudiant
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
        import-module activedirectory
        $users = Get-ADGroupMember -identity $old_groupe | select SamAccountName
        foreach ($user in $users){
        $sam=$user.SamAccountName
                import-module activedirectory
                try{remove-ADGroupMember -Identity  $old_groupe -Member $sam -Confirm:$false}catch{""} #on retire l'étudiant de l'ancien groupe
                Add-ADGroupMember -Identity  $new_groupe -Members $sam -PassThru  #on ajoute l'étudiant dans son nouveau groupe
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
#méthode après le lancer de la Form

#méthode afficher membre du groupe pour la migration
$combox_migration_SelectedIndexChanged=
{
   If ($combox_migration.text -ne $null) {
        [void] $listbox_groupe_new.Items.Clear()
        $nam_new_grp=$combox_migration.SelectedItem.ToString() 
        Import-Module ActiveDirectory
        $recup_group=Get-ADGroupMember -identity  $nam_new_grp | select SamAccountName
        foreach ($users in $recup_group){
        $sam=$users.SamAccountName
             [void] $listbox_groupe_new.Items.Add("$sam")#commentaire
        }
   }
}
################ MUST CREATE BEFORE ASSIGN ################
$combox_migration.Add_SelectionChanged($combox_migration_SelectedIndexChanged)

#méthode afficher membre du groupe avant leur suppression
$combox_supp_grp_etu_SelectedIndexChanged=
{
   If ($combox_migration.text -ne $null) {
        [void] $listbox_supp_etu_AD.Items.Clear()
        $nam_new_grp=$combox_supp_grp_etu.SelectedItem.ToString() 
        Import-Module ActiveDirectory
        $recup_group=Get-ADGroupMember -identity  $nam_new_grp | select SamAccountName
        foreach ($users in $recup_group){
        $sam=$users.SamAccountName
             [void] $listbox_supp_etu_AD.Items.Add("$sam")#commentaire
        }
   }
}
################ MUST CREATE BEFORE ASSIGN ################
$combox_supp_grp_etu.Add_SelectionChanged($combox_supp_grp_etu_SelectedIndexChanged) 


#méthode afficher membre du groupe envoyant les migrants
$combox_migr_grp_etu_left_SelectedIndexChanged=
{
   If ($combox_migr_grp_etu_left.text -ne $null) {
        [void] $listbox_migr_etu_AD_old.Items.Clear()
        $nam_new_grp=$combox_migr_grp_etu_left.SelectedItem.ToString() 
        Import-Module ActiveDirectory
        $recup_group=Get-ADGroupMember -identity  $nam_new_grp | select SamAccountName
        foreach ($users in $recup_group){
            $sam=$users.SamAccountName
             [void] $listbox_migr_etu_AD_old.Items.Add("$sam")#commentaire
        }
   }
}
################ MUST CREATE BEFORE ASSIGN ################
$combox_migr_grp_etu_left.Add_SelectionChanged($combox_migr_grp_etu_left_SelectedIndexChanged)

#méthode afficher membre du groupe reçevant les migrants
$combox_migr_grp_etu_right_SelectedIndexChanged=
{
   If ($combox_migr_grp_etu_right.text -ne $null) {
        [void] $listbox_migr_etu_AD_new.Items.Clear()
        $nam_new_grp=$combox_migr_grp_etu_right.SelectedItem.ToString() 
        Import-Module ActiveDirectory
        $recup_group=Get-ADGroupMember -identity  $nam_new_grp | select SamAccountName
        foreach ($users in $recup_group){
             $sam=$users.SamAccountName
             [void] $listbox_migr_etu_AD_new.Items.Add("$sam")#commentaire
        }
   }
}
################ MUST CREATE BEFORE ASSIGN ################
$combox_migr_grp_etu_right.Add_SelectionChanged($combox_migr_grp_etu_right_SelectedIndexChanged)

#function qui met à jour les listbox des pannels de migration
function Refresh_listbox{
        #on refresh la listbox
        Import-Module ActiveDirectory
         [void] $listbox_groupe_new.Items.Clear()
         try{$nam_new_grp=$combox_migration.SelectedItem.ToString() 
         Import-Module ActiveDirectory
         $recup_group=Get-ADGroupMember -identity  $nam_new_grp | select SamAccountName
         foreach ($users in $recup_group){
             $sam=$users.SamAccountName
             [void] $listbox_groupe_new.Items.Add("$sam")#commentaire
            }
            }catch{""}
        #on refresh la listbox de gauche du pannel "migrer tout un groupe"
        [void] $listbox_migr_etu_AD_old.Items.Clear()
        try{$nam_old_grp=$combox_migr_grp_etu_left.SelectedItem.ToString() 
        Import-Module ActiveDirectory
        $recup_group=Get-ADGroupMember -identity  $nam_old_grp | select SamAccountName
        foreach ($users in $recup_group)
        {
            $sam=$users.SamAccountName
            [void] $listbox_migr_etu_AD_old.Items.Add("$sam")#commentaire
        }
        }catch{""}
       #on refresh la listbox de droite du pannel "migrer tout un groupe"
        [void] $listbox_migr_etu_AD_new.Items.Clear()
        try{$nam_new_grp=$combox_migr_grp_etu_right.SelectedItem.ToString()
        Import-Module ActiveDirectory
        $recup_group1=Get-ADGroupMember -identity  $nam_new_grp | select SamAccountName
        foreach ($users in $recup_group1)
        {
            $sam=$users.SamAccountName
            [void] $listbox_migr_etu_AD_new.Items.Add("$sam")#commentaire
        }
        }catch{""} 
        #on refresh la listbox du pannel "supp tout un groupe"
        [void]$listbox_supp_etu_AD.Items.Clear()
        try{$nam_grp=$combox_supp_grp_etu.SelectedItem.ToString()
        Import-Module ActiveDirectory
        $recup_group2=Get-ADGroupMember -identity  $nam_grp | select SamAccountName
        foreach ($users in $recup_group2)
        {
            $sam=$users.SamAccountName
            [void] $listbox_supp_etu_AD.Items.Add("$sam")#commentaire
        }
        }catch{""}

}
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
$Form.ShowDialog() | out-null