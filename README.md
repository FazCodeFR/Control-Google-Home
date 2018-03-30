# Bienvenue, Contrôler son ordinateur avec sa Google Home
Contrôle de son ordinateur Windows avec Google Home. 

1)

Requis : https://aymkdn.github.io/assistant-plugins/ (! Mettre le dossier "assistant-plugins" dans le dossier : C:\GoogleHome et lancer l'installation a cet endroit !)

2)

Une fois l'installation d'assistant-plugins terminé continué ci-dessous
Mettre GoogleHome.vbs dans le dossier : C:\GoogleHome (Si il n'existe pas, le créer)

3)

Pour lancer le programme au démarrage et en arrière-plan :
[WINDOWS] + [R] 
Une fenêtre exécuter ça s'ouvrir et y mettre "CMD", v
alider une fois l'invite de commande ouverte il mettre ligne par lignes les codes suivants :
'cd C:\GoogleHome\assistant-plugins
npm install pm2 -g
npm install pm2-windows-startup -g
pm2-startup install
pm2 start index.js
Start C:\GoogleHome\GoogleHome.vbs
exit'

Redémarrer votre ordinateur.

4)


IFTTT -> Contrôler son pc : https://ifttt.com/applets/jSNrZ4vJ-controle-de-l-ordinateur-avec-google-assitant

Contrôler son décodeur : https://ifttt.com/applets/N8a5s4BE-controle-du-decodeur-orange-avec-google-assitant


Pour contrôler son pc vocalement vous n'avez plus cas dire : Ok google sur le pc XXX ou Ok google sur l'ordinateur XXX

Exemple : ok google sur le pc lance google chrome
ok google sur le pc donne le mot de passe wifi
ok google sur le pc ouvre le lecteur CD

Tuto vidéo : https://www.youtube.com/watch?v=-Z2OFYVMIKk
