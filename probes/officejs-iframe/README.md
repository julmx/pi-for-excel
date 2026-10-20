# Expérience Office.js dans une iframe inter-origines

**Objectif unique :** vérifier, dans Excel Windows WebView2, si un script Office.js généré à l'extérieur de l'origine privilégiée du complément peut exécuter une opération Excel native tout en restant isolé du stockage principal.

**Statut :** prototype de faisabilité, **pas une fonctionnalité Enterprise** ni une autorisation d'utiliser des classeurs confidentiels. Aucun changement dans les outils, le manifeste Enterprise, le serveur Enterprise ou la CSP de production.

## Exécution sous Windows

1. Ouvrir un **nouveau classeur jetable**, sans contenu réel.
2. Récupérer la branche `experiment/isolated-officejs-iframe-20261019` du fork.
3. Depuis la racine du dépôt, vérifier que `cert.pem` et `key.pem` couvrent **localhost et 127.0.0.1** (comme dans le guide Enterprise Windows). Ne pas commiter ces fichiers.
4. Lancer :

   ```powershell
   node scripts/probe-officejs-iframe.mjs
   ```

5. Charger dans Excel le manifeste affiché par le script (normalement `%TEMP%\pi-officejs-iframe-probe.xml`) selon la méthode d'installation des compléments autorisée par l'organisation. Le complément apparaît séparément comme **Pi Office.js Probe**. Il possède un identifiant distinct.
6. Ouvrir **Pi Office.js Probe**, cliquer sur **1. Créer la feuille synthétique**, puis **2. Tester Office.js dans l'iframe**.
7. Observer le résultat dans le volet et vérifier visuellement, dans la feuille `__PiIframeProbe`, l'AutoFilter sur `A1:C4` et l'échelle de trois couleurs sur `C2:C4`.

Le test ne réutilise ni la clé du fournisseur LLM, ni l'agent, ni `execute_office_js`. Il ne charge aucune donnée du classeur hors de la feuille synthétique. Le script envoyé à l'iframe est **fixe**, bien que son mode d'exécution reproduise un import dynamique de module JavaScript.

## Architecture et limites

- Hôte : `https://127.0.0.1:3151`, charge Office.js, crée uniquement une feuille absente et conserve un marqueur IndexedDB synthétique.
- Exécuteur : `https://localhost:3152`, iframe **sans attribut `sandbox` pour ce test différentiel**. L’isolation DOM/IndexedDB reste fondée sur des origines HTTPS différentes. Son domaine est déclaré dans `AppDomains` du manifeste de **test**, condition requise par Microsoft pour les appels Office.js depuis une iframe distincte.
- L'hôte envoie un code fixe par `postMessage` avec vérification de l'origine et de la source. L'exécuteur fait un `import(blob:)` et un `Excel.run()` dans **son propre** contexte.
- Seule l'origine exécuteur accepte `blob:` dans `script-src`. La politique principale et celle de Python Enterprise ne sont pas changées.
- Le test vérifie l'accès au DOM parent (qui doit être refusé), la visibilité du marqueur IndexedDB parent (qui doit être fausse) et l'application effective des modifications depuis une seconde lecture côté hôte.
- Les serveurs se limitent aux fichiers de cette expérience et aux adresses loopback. Ils n'exposent aucun relais API ni accès réseau arbitraire. Les deux pages chargent le script Office.js hébergé par Microsoft.

**Attention :** un script ayant accès à Office.js peut toujours lire ou modifier d'autres plages du classeur. L'isolation de l'iframe **ne signifie pas** une limitation des permissions Excel. Une CSP peut aussi laisser des vecteurs de navigation ou des sorties indirectes ; ce test ne constitue donc pas une preuve d'absence d'exfiltration.

## Critères de décision

| Vérification | Attendu |
| --- | --- |
| Office.js dans l'iframe déclarée | `officeReady=true` et aucune erreur d'autorisation Office |
| Opérations Excel natives | Filtre actif et une règle conditionnelle ; vérifiés par l'hôte |
| Séparation entre origines | `parentAccessBlocked=true` et `canaryVisible=false` |
| Réseau / CSP | Examiner les requêtes et violations CSP dans les DevTools WebView2 ; seules les requêtes nécessaires au test doivent apparaître |
| Privilèges Excel | Les restrictions de plage ne sont **pas** garanties ; exiger un design distinct si elles sont requises |

**Ne pas fusionner cette branche seulement parce que les tests fonctionnels passent.** Vérifier en outre le comportement réseau sous WebView2, le chargement d'Office.js dans la sous-frame et les limites de permissions Office. Tester l'ajout/absence d'`AppDomains` sur un manifeste de test séparé si le comportement reste ambigu.

Ce prototype ne résout pas non plus la validation des paramètres de la passerelle Python Enterprise, ni la restauration transactionnelle des scripts arbitraires.

## Nettoyage

Supprimer le complément **Pi Office.js Probe** d'Excel et fermer les deux serveurs avec Ctrl+C. Supprimer le classeur jetable. Le marqueur IndexedDB synthétique est supprimé en fin de test ou à la fermeture du volet, sous réserve que le navigateur effectue l'opération ; en cas d'interruption forcée, effacer le stockage de l'origine de test `https://127.0.0.1:3151`.

Références : documentation Microsoft **Office Add-ins manifest / AppDomains**, **manifest element ordering**, **Excel.AutoFilter** et **Excel.ColorScaleConditionalFormat**.

## Diagnostic du 20 octobre 2026

Premier essai réel sous Excel Windows : l'hôte a annoncé `Office.onReady { host: Excel }`, mais l'iframe a annoncé `officeReady=false`. Le code Office.js fixe n'a donc pas encore été validé. Cette révision retire **uniquement l'attribut `sandbox` de l'iframe**, afin de tester si cette restriction affecte l'initialisation Office.js ; les origines parent et enfant restent différentes. Le résultat est désormais diagnostiqué avec un timeout de 10 s. **L'absence d'attribut sandbox ne constitue pas un renforcement de sécurité** ; si le test fonctionne, une analyse de menace supplémentaire sera indispensable avant toute intégration. Cette variation ne doit pas être transférée automatiquement à l'Enterprise.
