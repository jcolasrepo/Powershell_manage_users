# Powershell_manage_users
application powershell pour créer,migrer,supprimer des utilisateurs au sein d'un AD

compatible windows server 2008 et + 
Libre d'utilisation 

#description

Le script comprend 6 pannel

-Crée un utilisateurs à la fois et les ajouter à un groupe

-crée plusieurs utilisateurs à la fois via un fichier .csv comprenant leurs Firstname;LastName

-Supprimer 1 ou plusieurs utilisateurs

-Supprimer TOUS les utilisateurs au sein d'un groupe

-Migrer 1 ou plusieurs utilisateurs d'un Groupe à un autre

-Migrer TOUS les utilisateurs d'un Groupe à un autre

application pour gérer les utilisateurs au sein d'un Active Directory Version 2.7


#Ajout  Version 2.8

-Optimisation Visuel du code Source

-Ajout de commentaire

-Possibilité de de se connecter à distance au script via la technologie WinRM de windows

WinRM est activé par défault sur les versions windows server 2012 et supérieur pour vérifier si winRM est activer faire:

-dans un terminal faire la commande :

Get-Service WinRM

Si le status est "running" alors winrm est activer sinon faire:

Enable-PSRemoting

pour activer Winrm

-Vérifier que les ports 5985(http)et 5986(https) entrant sur le serveur sont ouverts

-Vérifier que votre client peut se connecter , les clients qui sont dans un groupe de travail(WORKGROUP)devront être ajouté en tant que "trusted host"ou utilisé une connexion en https faire:

winrm get winrm/config/client

pour afficher la configuration et:

Set-Item WSMan:\localhost\Client\TrustedHosts 192.168.0.1 

ou Set-Item WSMan:\localhost\Client\TrustedHosts CHA_DC_02\test.fr

ou NON RECOMMANDER Set-Item WSMan:\localhost\Client\TrustedHosts * pour autoriser toutes les connexions


