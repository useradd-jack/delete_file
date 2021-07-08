Ajouter les options voulus dans le fichier « listeserveurs.csv ».
Le détail des options est dans le fichier « Options.docx ».


Exemple avec un répertoire de test.

<img width="455" alt="image" src="https://user-images.githubusercontent.com/85949171/124822552-52001380-df70-11eb-8ea7-f280e9df2f01.png">

Le script est lancé en tâche planifiée.
Je le lance manuellement pour le test.

<img width="455" alt="image" src="https://user-images.githubusercontent.com/85949171/124822578-5a584e80-df70-11eb-852f-1ec1bff51393.png">



Dans l’exemple, le fichier csv :

![image](https://user-images.githubusercontent.com/85949171/124873045-42131e80-dfc6-11eb-9f5d-81692ca0a39b.png)


Un fichier de log est dans le répertoire « D:\Sources\ScriptPurge\Log\purgetest.txt » du serveur.
Le script supprime les fichiers dans le répertoire « C:\Users\jack\Desktop\testpurge » qui se terminent en .txt et en garde deux.

Le fichier de log est disponible :
Il reste que deux fichiers en .txt et les autres ne sont pas supprimés étant donnée qu’on demande la suppression des fichiers .txt uniquement.


