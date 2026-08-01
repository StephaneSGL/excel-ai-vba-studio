# Sécurité / Security

## Versions prises en charge

Excel AI & VBA Studio est actuellement en Preview. Seule la dernière version publiée reçoit des correctifs de sécurité.

| Version | Prise en charge |
| --- | --- |
| Dernière version publiée | Oui |
| Versions antérieures | Non |

## Centre de sécurité d’entreprise

Le Centre de sécurité de l’extension fournit un diagnostic local en lecture
seule. Il distingue les signaux du fichier, les préférences Office et les
stratégies gérées détectées, puis indique quelles capacités de l’extension sont
disponibles, bloquées ou indéterminées.

Les signaux Microsoft 365 Cloud Policy, registre de stratégie Windows,
client Intune, inscription MDM et historique GPO sont affichés séparément. La présence d’un agent ou
d’une inscription MDM ne suffit jamais à attribuer une règle Office précise à
Intune : une même clé de stratégie peut avoir été livrée par GPO, MDM,
Configuration Manager ou un script. Le Centre affiche donc la preuve technique
et conserve l’attribution comme indéterminée lorsqu’elle ne peut pas être
démontrée localement.

Les métadonnées de sensibilité sont suivies par les relations OPC officielles
(`LabelInfo.xml`) ou par la partie de propriétés personnalisées historique
correctement reliée et typée. Elles restent des déclarations locales non
authentifiées. Elles ne prouvent ni la politique actuellement publiée dans le
tenant, ni un rang de confidentialité, ni un chiffrement effectif. Les flux
LabelInfo CFB/IRM chiffrés non analysables sont signalés comme indéterminés ;
la structure IRMDS est distinguée d’un simple package protégé par mot de passe.

Le Centre n’ouvre jamais le classeur inspecté. Cette séparation évite qu’un
`Workbook_Open`, une feuille macro Excel 4.0 ou un autre contenu actif soit
déclenché simplement pour consulter le diagnostic.

L’inspection conserve un verrou de lecture pendant tout le diagnostic afin que
le fichier ne puisse pas être remplacé entre deux contrôles. Les flux
`Zone.Identifier`, les registres, les parties XML et les conteneurs CFB sont
bornés. Un `ZoneId` invalide, des sections `ZoneTransfer` ambiguës, une
hiérarchie CFB excessive ou un inventaire macro incomplet donnent un état
indéterminé plutôt qu’une autorisation implicite.

Les emplacements approuvés incluent les chemins intégrés documentés par Excel
et les clés `LocationN` effectivement présentes. Les variables d’environnement
sont développées localement ; un chemin non résolu ou une énumération tronquée
est signalé comme partiellement illisible.

Ce diagnostic et les liens explicites vers les portails ne remplacent pas Microsoft 365 Apps admin center, Intune, les
résultats de stratégie de groupe, Microsoft Defender ou Microsoft Purview. Il
ne modifie jamais le registre, le Centre de gestion de la confidentialité,
AccessVBOM, ActiveX, les emplacements approuvés ou la marque d’origine Internet.
Une règle gérée doit être modifiée par un administrateur autorisé dans l’outil
d’entreprise qui l’a déployée.

Une signature VBA et une signature numérique du package Office sont deux
protections distinctes. Si le package OOXML contient un graphe de signature
(relations OPC origine/signature et Content Types effectifs, quelle que soit
l’URI valide des parties), l’extension le laisse en lecture seule et refuse
toutes ses voies d’écriture, y compris Enregistrer sous, la grille XLSM, le
bootstrap VBA, les UserForms, les boutons et ActiveX. Si cet état ne peut pas
être vérifié de façon sûre, l’écriture est également refusée.
Le Centre confirme uniquement la structure OPC d’une signature de package ; il
ne valide pas cryptographiquement le certificat, sa chaîne de confiance ou sa
révocation.

## Chaîne de livraison native

Le helper Windows est reconstruit en CI avec Python 3.11.9 et des dépendances
verrouillées par version et SHA-256. Le hash du résultat est publié dans le
résumé du job, puis les tests de sécurité, de mutation et de rollback sont
exécutés contre cet exécutable avant qu’il soit transmis au job de packaging.
Un binaire PyInstaller n’est pas présenté comme reproductible octet par octet
entre des hôtes Windows/Python différents.

Le helper distribué ne possède pas encore de signature Authenticode. La
publication automatique sur le Marketplace reste désactivée tant qu’une
identité de charge de travail Microsoft Entra et la relation de confiance de
l’éditeur ne sont pas configurées. Ces deux limites sont des travaux de
durcissement de livraison ; elles ne réduisent pas les refus d’écriture et les
contrôles locaux décrits ci-dessus.

## Réseau et contenu webview

Le client d’API personnalisée conserve la validation TLS native. Il exige
HTTPS, sauf pour une adresse loopback locale, refuse les identifiants et les
paramètres `key` dans l’URL, transmet les clés dans les en-têtes et ne suit pas
les redirections. Les réponses non streamées sont limitées à 16 Mio ; les flux
à 64 Mio, avec un tampon de ligne limité à 1 Mio de caractères. Le client HTTP
générique limite également chaque corps de réponse à 64 Mio. Les lectures non
streamées refusent aussi les réponses fragmentées en plus de 16 384 blocs.

Les URI d’icônes rendues dans les webviews sont limitées aux ressources SVG
HTTPS émises par l’hôte de ressources VS Code. Le dispatch de messages exige un
gestionnaire explicitement enregistré, les namespaces OOXML sont comparés dans
leur intégralité et les chaînes de police générées pour le canvas sont
échappées avant insertion dans une propriété CSS.

Les anciens arbres d’actifs Vditor/PDF et les points d’entrée inactifs des
providers Markdown/Java qui n’étaient ni construits ni inclus dans le VSIX ont
été supprimés. Leur remise en service nécessiterait une nouvelle revue de
sécurité et, pour PDF.js, une mise à niveau complète plutôt qu’un patch isolé
de l’ancienne copie.

## Signaler une vulnérabilité

Ne publiez pas de vulnérabilité présumée dans une issue, une discussion, un journal ou un classeur partagé publiquement.

Utilisez exclusivement le [signalement privé de vulnérabilité GitHub](https://github.com/StephaneSGL/excel-ai-vba-studio/security/advisories/new). Ce canal permet de collaborer en privé avec le mainteneur avant toute divulgation.

Indiquez, si possible :

- la version de l'extension, de VS Code, de Windows et de Microsoft Excel ;
- le scénario d'attaque, l'impact attendu et les prérequis ;
- des étapes minimales de reproduction avec des données synthétiques ;
- toute mesure de réduction du risque déjà identifiée.

N'envoyez aucun classeur réel, secret, identifiant, donnée personnelle, code VBA propriétaire ou information d'entreprise. Le rapport sera examiné au mieux des disponibilités ; merci de laisser au mainteneur le temps de confirmer et de corriger le problème avant toute publication.

---

Please report suspected vulnerabilities only through [GitHub private vulnerability reporting](https://github.com/StephaneSGL/excel-ai-vba-studio/security/advisories/new), never through a public issue or discussion. Include minimal synthetic reproduction steps and remove all confidential, personal, credential, workbook, and proprietary VBA data. Only the latest published version is supported.
