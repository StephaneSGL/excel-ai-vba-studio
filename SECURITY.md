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

Ce diagnostic ne remplace pas Microsoft 365 Apps admin center, Intune, les
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
