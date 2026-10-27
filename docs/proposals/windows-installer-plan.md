# Plan de distribution Windows autonome de Pi for Excel

Statut : plan de mise en œuvre. Les fonctionnalités décrites ci-dessous restent à réaliser ou à valider.

## Objectif

Produire un installateur Windows permettant d'utiliser Pi for Excel localement, sans environnement de développement ni serveur distant à administrer. Internet reste nécessaire pour les appels aux modèles d'IA et le chargement actuel d'Office.js.

Le travail s'appuie sur le [lancement local Windows existant](../enterprise-windows-smoke.md), le [serveur HTTPS](../../scripts/serve-dist-https.mjs) et le [générateur de manifeste local](../../scripts/generate-enterprise-local-manifest.mjs).

## 1. Valider l'installation dans Excel

Prendre comme cible initiale Windows 11 avec Excel Microsoft 365 et WebView2.

Vérifier le chargement du complément depuis le serveur local, la confiance du certificat HTTPS et les restrictions éventuelles du poste. Déterminer la méthode d'enregistrement du complément et les opérations automatisables. Prévoir une activation guidée dans Excel si elle reste nécessaire.

**Livrable :** une procédure reproductible précisant les prérequis et les droits nécessaires.

Cette validation constitue le premier jalon, avant la fabrication de l'installateur.

## 2. Préparer un paquet autonome

Embarquer l'application compilée, un moteur Node.js privé et les composants du serveur local. L'utilisateur n'aura pas à installer Git, Node.js ou des outils de développement.

Réutiliser le mode entreprise existant, ses restrictions réseau et son environnement Python isolé. Préconfigurer les adresses des API autorisées, puis permettre à chaque utilisateur de renseigner ses identifiants au premier lancement. Aucun secret ne sera inclus dans le paquet.

**Livrable :** un dossier distribuable fonctionnant sur un PC sans environnement de développement.

## 3. Créer le lanceur Windows

Développer une petite application accessible depuis le menu Démarrer et la zone de notification. Elle permettra de démarrer et arrêter le serveur, d'afficher son état et d'expliquer les erreurs courantes.

Ajouter une option de démarrage à l'ouverture de session, empêcher les lancements multiples et conserver des adresses locales stables. Prévoir un diagnostic des ports, du certificat et de l'accès à l'API, exportable sans clés ni contenu des classeurs.

**Livrable :** un fonctionnement quotidien sans fenêtre de terminal.

## 4. Construire l'installateur Windows

Utiliser Inno Setup pour produire un fichier `.exe`, avec une installation par utilisateur lorsque le poste le permet.

L'installation préparera les fichiers, les raccourcis, le lancement automatique optionnel et l'enregistrement du complément. Elle générera un certificat propre au poste, protégera sa clé privée et prévoira son renouvellement, conformément aux règles de confiance du poste.

Ajouter une désinstallation retirant les éléments créés par Pi et proposant de conserver ou supprimer les données locales.

**Livrable :** `PiForExcel-Setup-<version>.exe`.

## 5. Organiser la fabrication et les mises à jour

Automatiser la fabrication du paquet sur Windows et attribuer une version identifiable à chaque livraison. Prévoir la signature numérique du lanceur et de l'installateur.

Les mises à jour passeront initialement par un nouvel installateur. Elles devront préserver les paramètres et l'historique, avec une procédure de retour à la version précédente tenant compte des éventuelles migrations de données.

**Livrable :** une chaîne de fabrication reproductible et une procédure de mise à jour.

## 6. Vérifier le parcours complet

Tester sur un PC propre :

- l'installation et le redémarrage de Windows ;
- l'ouverture d'Excel et la connexion à l'API ;
- la lecture et la modification des cellules, ainsi que la récupération ;
- Python et les modes Auto/Confirm ;
- les erreurs de certificat, les ports occupés et les interruptions réseau ;
- les restrictions du poste, la mise à jour et la désinstallation.

**Livrables :** un installateur validé, un compte rendu des vérifications et une notice d'utilisation d'une page.

**Critère d'acceptation :** un utilisateur non technique peut installer, configurer et réutiliser Pi dans Excel sans commande ni assistance technique.
