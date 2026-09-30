#Requires -Version 7.0
<#
.SYNOPSIS
    PCPulse Dashboard v2.5.4
.DESCRIPTION
    Lit les JSON produits par le Collector sur tous les PC du parc et
    genere un tableau de bord HTML autonome avec KPIs, filtres, tri,
    drill-down par PC, dark/light mode.
.NOTES
    Auteur       : Damien Gouhier
    Repository   : https://github.com/Damien-Gouhier/pcpulse
    Licence      : MIT
    Version      : 2.5.4
    Runtime      : PowerShell 7+ (pwsh.exe)
.CHANGELOG
    v2.5.4 : [HYBRIDE] 2e source de donnees : conteneur Azure "depot" (postes 100%
           distants qui deposent via la Function). Le Dashboard lit AUSSI le cloud
           (SAS read+list, REST, zero dependance) et FUSIONNE avec le share en
           gardant, par machine, le rapport le plus recent (Machine.CollectedAt).
           Gate par config (CloudDepotBaseUrl + CloudDepotSas) ; vides => share seul
           (inchange). Lecture SEULE. SchemaVersion inchangee (2.2).
    v2.5.3 : [UI] Tag CPU d'age (Recent/Vieillissant/Ancien) remplacable par une
           categorie de RENOUVELLEMENT binaire, pilotee par config (nouvelle cle
           RenewalCandidateMaxYear, lue par le Dashboard). Annee CPU <= seuil =>
           "A renouveler", sinon "Parc courant". Repere 2019 = Intel Core <= 9e gen
           / AMD Ryzen <= 3000. Impacte le badge du tableau, le filtre (renomme
           "Renouvellement"), le KPI (ex "CPU anciens") et l'export CSV (colonne
           CandidatRenouvellement). L'onglet Materiel garde TOUJOURS l'annee + l'age
           reel. Cle absente/nulle => fallback sur l'ancien tag d'age (retrocompat).
           Calcule cote Dashboard depuis CPUYear (deja collecte) : AUCUN changement
           collecteur, aucun redeploy. En mode renouvellement, PLUS AUCUN libelle
           d'age (Ancien/Vieillissant) affiche : les raisons de verdict basees sur
           l'age CPU sont supprimees (l'axe renouvellement vit dans badge/filtre/KPI,
           l'age chiffre reste dans Materiel). En fallback (cle absente), le verdict
           et le tag d'age historiques sont conserves.
    v2.5.2 : [UI] Filtre par version de client VPN dans la barre d'outils (menu
           "VPN :", sur le modele des filtres OS/Modele). Peuple les versions
           distinctes de VpnClient.Version (tri decroissant : 7.x en haut, 6.x
           regroupees en bas), masque si moins de 2 versions. But : traquer les
           vieilles versions FortiClient (6.x) sur le parc heterogene ; selectionner
           une version filtre le tableau (= liste des postes concernes) et l'export
           CSV suit. Ajoute a clearAllFilters et anyFilter. Dashboard-only,
           SchemaVersion inchangee (2.2). Lit le champ peuple par le Collector 2.5.1.
           [UI] Barre d'outils compactee pour tenir Periode + tous les filtres +
           recherche + CSV sur UNE ligne (libelles periode courts 24h/7j/15j/30j,
           gap et padding reduits, recherche elastique). Masquer sains / Vue
           detaillee / Anomalies restent en 2e ligne (toolbar-break inchange).
           [UI] Filtre chassis par ICONES (Portable/Fixe/AIO) dans la barre d'outils,
           en toggle (clic = filtre, re-clic = enleve) plutot qu'un menu deroulant -
           reutilise les glyphes du tableau et le champ Machine.ChassisInfo
           (IsLaptop/IsDesktop/IsAIO). Ajoute a clearAllFilters, anyFilter et
           l'export CSV. Aucun changement collecteur (donnee deja collectee >= v5.6).
    v2.5.1 : [INVENTAIRE] Client VPN (FortiClient) : lecture de l'objet VpnClient
           (Collector 2.5.1) - coverage-check + remap embed + ligne "Client VPN"
           dans l'onglet Securite du drill-down + 3 colonnes CSV
           (VpnPresent/VpnProduct/VpnVersion). Present + version seulement, PAS
           cable au score (inventaire). Vide tant que le Collector 2.5.1 n'a pas
           redeploye (affiche "non collecte").
           [UI] "Silencieux / hors ligne" RETIRE de la file d'actions : un poste
           eteint n'est pas une action a traiter (parc a moitie eteint l'ete). Le
           compte reste dans la tuile "En ligne 24h" et le filtre 'offline' existe
           toujours -> on retire l'incitation a agir, pas la visibilite.
           Dashboard-only : Collector non impacte par ces deux points cote UI,
           SchemaVersion inchangee (2.2).
    v2.5.0 : [UI] Colonne "Boot" retiree du tableau principal (la duree/date de
           boot reste dans le drill-down, onglet Boot Performance). Ajout de deux
           filtres uptime dans la barre d'outils : "Uptime > 7j" (a surveiller) et
           "Uptime > 30j" (jamais reboote), memes mecanique/chip que les autres
           filtres KPI. Va de pair avec le fix uptime du Collector 2.5.0.
           [UI/cockpit UI] Habillage "cockpit" : rail de
           navigation a gauche + file d'actions a droite du tableau. Reskin +
           couche de navigation only ; aucune donnee ni fonction modifiee.
    v2.4.8 : [GARDE-FOU] Coverage-check COMPLETE. Il manquait la marche la plus
           haute : aucun controle au niveau TOP-LEVEL du payload, donc un BLOC
           entier ajoute au Collector pouvait disparaitre de l'embed sans un mot
           (le pire cas de la classe #3 : on pense aux champs d'un objet existant,
           pas a brancher un objet neuf). Ajout de $KnownPayloadKeys + 18 checks
           de sous-objets restants (Meta, Events, BSODs, ResourceWarnings, TopRAM,
           GPUInventory, BatteryInfo, ServicesHealth[.Monitored], BootPerformance
           [.LastBoot/.History/.Stats], MemoryInventory, HardwareHealth[.GPU_TDR/
           .Thermal/.CPUThrottling], Stats). La couverture est desormais totale :
           tous les blocs de l'embed sont remappes champ par champ, la classe de
           bug s'appliquait donc partout, pas seulement aux quelques sous-objets
           gardes depuis la 2.3.1.
           Deux champs emis-mais-non-lus mis au jour et declares : BSODs.Taille
           (deja connu) et BootPerformance.LastBoot/History.BootStartTime (trouve
           EN AJOUTANT le garde-fou -- il fait son travail des l'installation).
           [DIAGNOSTIC] Signal de schema deprecie : les postes encore en schema
           2.1 sont nommes en console a chaque generation. C'est le prerequis
           mesure au resserrage de $AcceptedSchemaVersions sur @('2.2') : s'il en
           reste, ce sont des postes dont le Collector n'a pas pris cinq mises a
           jour d'affilee -> probleme d'auto-update a traiter AVANT de resserrer,
           sinon on ne corrige rien, on les rend juste invisibles.
           Dashboard-only : Collector INCHANGE en 2.4.7, version.txt non touche,
           aucun redeploiement parc. SchemaVersion inchangee (2.2).
    v2.4.7 : Modele machine (Machine.Model/Manufacturer, Collector 2.4.7). Affichage
           dans le panel Materiel (carte OS), recherche texte etendue (modele +
           fabricant), filtre deroulant "Modele" (masque si < 2 modeles), panneau
           de repartition des modeles en bas de page (tuiles cliquables -> filtre),
           colonnes Fabricant/Modele au CSV. Anti-XSS : Model/Manufacturer echappes
           a l'embed via ConvertTo-HtmlSafe + KnownMachineKeys. Retro-compatible :
           champ vide / panneau masque pour les JSON < 2.4.7.
           [FIX] Chevauchement decom : le tableau technicien et les barres "par mois"
           sont empiles verticalement (fini le cote-a-cote ou la table debordait sur
           la barre). Table en table-layout:fixed. [UX] Repartition modeles en barres
           classees compactes (au lieu d'une tuile par modele). [MENAGE] Retention des
           dashboards horodates : on ne garde que les 10 plus recents.
    v2.4.5 : [FIX] Double-encodage HTML corrige (un nom type "O'Brien" s'affichait
           "O&#39;Brien") : la sanitisation en place de Test-PCPulseJson est retiree,
           l'echappement se fait une SEULE fois a l'embed (ConvertTo-HtmlSafe).
           Verifie : chaque champ concerne est bien echappe a l'embed, aucune perte
           de protection XSS. [ROBUSTESSE] CollectedAt caste en [datetime] sous
           try/catch : un JSON avec une date pourrie ne fait plus planter TOUTE la
           generation (le poste s'affiche hors ligne). Dashboard-only (Collector
           inchange en 2.4.4, pas de redeploiement parc).
    v2.4.4 : Aligne sur le numero commun (Collector 2.4.4). Aucun changement
           fonctionnel cote Dashboard ce lot (fix log Collector + retention backups
           Updater).
    v2.4.3 : [SECURITE] XSS pass 3 : DiskInfo (Drive) + champs numeriques disque,
           DurationMin, OSBuild echappes/castes avant innerHTML (un poste compromis
           pouvait injecter du JS chez l'admin). DiskInfo/BootDurations ajoutes au
           garde-fou Test-EmbedCoverage. config.psd1 lu en priorite depuis release\
           (lecture seule postes). SchemaVersion inchangee.
    v2.4.2 : N° de serie machine + audit du cycle de vie. SchemaVersion inchangee.
           - N° de serie (Machine.SerialNumber, collecte Collector 2.4.2) : affiche
             dans le drill-down Materiel (carte OS), colonne 'NumeroSerie' dans
             l'export CSV, et recherche etendue au n° de serie. Sert au report
             decommissionnement vers Excel.
           - Panneau "Cycle de vie du parc" (stats/audit decom) : lit le registre
             COMPLET (y compris PC dont le JSON a ete purge), independant des
             filtres. Compteurs fait / a faire / en retard, delai moyen
             marque->fait, repartition par technicien et par mois.
    v2.4.0 : Dashboard + outillage (le Collector reste en 2.3.2, INCHANGE -> aucun
           redeploiement du parc, version.txt non touche). SchemaVersion inchangee.
           - KPI Securite recompose : "EDR arrete" / "EDR absent" / "OS fin de
             support" (retrait de "PC offline", non pertinent en securite).
           - Tuile "En ligne 24h" cliquable -> filtre les postes hors ligne.
           - Colonne CPU triable (ancien -> recent).
           - Cycle de vie (decommission, EN DEVELOPPEMENT) : lecture du registre
             decommissioning.json (chemin DecommissionRegistryPath), badges par
             poste (a decommissionner / en retard / fait mais en ligne / faite a
             purger), famille KPI dediee + filtres. Outil Decommission-PC.ps1
             (v2.0) pour marquer/cloturer cote donnee. Additif, voue a evoluer
             (purge auto a venir).
    v2.3.2 : Backlog - lisibilite crashers, durcissement, garde-fou, additif.
           - Drill-down crashers (#9) : affichage "X plantes / Y figes", origine
             traduite depuis le module fautif (.exe = application, ntdll/kernelbase
             = interne, clr = .NET, autre dll = conflit de composant) et code
             exception en clair (0xc0000005 = acces memoire invalide, etc.). Le brut
             (module + code) reste en tooltip pour l'admin. Jamais "hang" ni code
             brut a l'ecran. Meme traduction dans la carte "piste memoire".
           - XSS passe 2 (#2) : cast STRICT des champs numeriques du re-mapping
             (EventId, Count(s), Severity, temperatures + SMART, CycleCount,
             ChassisType, VideoOutputTech, annees/age moniteurs, CPUYear/CPUAge...)
             en PRESERVANT null. CPUGen reste HTML-safe (peut valoir "Ultra2").
           - Coverage-check (#3) : generalise aux SOUS-OBJETS de l'embed (crashers,
             WHEA fatal/corrected, moniteurs, disques SMART, barrettes RAM), pas
             seulement Machine. Diagnostic console.
           - CollectorRunAs (#13) : recopie dans l'embed + colonne export CSV
             (audit gMSA -> SYSTEM). Ajoute a KnownMachineKeys.
           - Config (#4) : merge RECURSIF des sous-objets hashtable (ScoreWeights) -
             un override partiel n'efface plus les poids absents (score casse en
             silence). Les listes (MonitoredServices, PriorityApps) gardent la
             semantique de remplacement.
           - Menage (#5) : sanitization en place corrigee - IPAddress -> IP (champ
             reel), branches Manufacturer/Model retirees (inexistantes au niveau
             Machine), CollectorRunAs ajoute.
    v2.3.0 : Release unifiee Collector + Dashboard. Lisibilite (cotes Collector
           2.3.0). Retro-compatible avec les anciens JSON (fallback JS),
           SchemaVersion inchangee.
           - UI : cartes RAM et SMART aerees dans le panel Materiel (etaient
             tassees en colonne etroite : gaps/polices/largeurs revus, sec-ram
             et sec-smart elargies).
           - Utilisateur : affiche LastLoggedUser (dernier user connu) en repli
             quand CurrentUser = "(aucune session)", en muted + tag "dernier".
             Recherche et export CSV etendus a ce champ. Embed + KnownMachineKeys
             + sanitization mis a jour.
           - Ecrans : un ecran a EDID non transmis (dock/adaptateur/KVM) est
             libelle "Ecran non identifie" au lieu de "@@@ / 0000", exclu du top
             fabricants et du calcul d'age, compte dans une tuile dediee. Detection
             via le flag Identified (Collector 2.3.0) ou, a defaut, la signature
             EDID nul deduite cote JS (anciens JSON).
           - RAM : le fabricant JEDEC brut ("80AD000080AD") est decode en clair
             ("SK Hynix") ; decodage cote JS aussi pour les JSON pas encore
             reguleres par le Collector 2.3.0.
    v2.2.1 : Inventaire OS (Windows 10 vs 11) - affichage, filtre, coverage-check.
           - Re-mapping PS->embed : 4 champs Machine recopies (OSProduct / OSBuild /
             OSDisplayVersion / OSEdition). Sans cette recopie ils seraient droppes en
             silence (classe de bug recurrente : Type / BootType / crashers).
           - Colonne "OS" dans le tableau d'inventaire (badge Windows 11 vert / Windows 10
             ambre = EOL / Inconnu gris) + feature update en sous-ligne.
           - Filtre "OS :" dedie (facon site/CPU), populate dynamique, masque s'il n'y a
             qu'un seul OS present. Sert a sortir la liste des postes Windows 10.
           - Bloc OS dans le drill-down (panel Materiel) + colonnes OS dans l'export CSV.
           - COVERAGE-CHECK (anti-recidive #3) : a la generation, diffe les cles Machine
             des JSON vs une liste de cles connues et Write-Host jaune tout champ Machine
             non traite dans l'embed. Purement diagnostic console.
           - Menage : sanitization Machine.OSCaption alignee sur le vrai champ Machine.OS
             (l'ancienne branche visait un nom inexistant = no-op silencieux).
           SchemaVersion INCHANGEE (@('2.1','2.2') toujours accepte ; le champ n'est present
           que sur les JSON du Collector 2.2.1, absence toleree pendant le rollout).
    v2.2.0 : Generalisation pour publication GitHub (de-vendorise, pilote par config).
           Le code ne nomme plus aucun produit ; tout le specifique vit dans config.psd1.
           - Whitelist SchemaVersions : @('2.1') -> @('2.1', '2.2'). Le Dashboard lit les
             deux pour absorber le rollout poste-par-poste : un poste pas encore passe au
             Collector 2.2 (JSON 2.1) affiche l'EDR "en attente" (neutre), pas une alerte.
           - Services critiques : lecture de ServicesHealth.Monitored (liste) au lieu de
             l'ancien objet unique code en dur. Le service Role='EDR' pilote le score et le
             badge ; la liste complete apparait dans le drill-down (Services surveilles).
           - ScoreWeights : poids EDR renomme EDRDown (le code ne nomme plus le vendor).
             Le fallback JS protege un config qui ne l'aurait pas encore.
           - PriorityApps ("A investiguer en priorite") externalisee dans config.psd1.
             Defaut generique = socle bureautique/collab. Les applis metier / securite
             propres au site vivent dans config.psd1, plus dans le code.
           - Colonnes CSV EDR renommees (EDRStatus / EDRInstalle / EDRAlerte).
           NB cote prod : config.psd1 doit declarer MonitoredServices (dont l'entree EDR) et
           PriorityApps ; sinon l'EDR n'est pas suivi et la liste priorite retombe au defaut.
    v2.1.14 : Bouton "Tout effacer" dans la barre des filtres actifs.
           Reinitialise TOUS les filtres d'un coup : les chips (KPI, appli, masquer sains) ET les
           controles qui n'ont pas de chip a eux seuls (menus site/CPU, champ recherche). Le bouton
           apparait des qu'au moins un filtre est actif (condition elargie : state + valeur des
           selects + recherche). Retour a l'etat d'ouverture, tout le parc visible. NB : "masquer
           sains" repart sur off, qui est deja son defaut (MaskHealthyByDefault = false).
    v2.1.13 : Piste materielle "memoire" (Phase 2) + recopie des champs additifs 2.1.5.
           (1) $crasherList recopie desormais les 4 champs du Collector 2.1.5 (HangCount,
           ErrorCount, FaultModule, ExceptionCode) - ils etaient droppes au re-mapping PS->embed,
           donc invisibles cote JS. Compteurs castes [int], module/code htmlsafe. Prepare aussi
           l'affichage detaille du drill-down (a venir).
           (2) Le drill-down PC leve une "piste a verifier : memoire" quand un poste cumule
           (a) >=1 crasher de nature memoire (0xc0000005 acces invalide / 0xc0000374 corruption)
           ET (b) >=1 erreur WHEA de composant RAM. Choix de CO-OCCURRENCE et non de coincidence
           temporelle a la minute (les WHEA corrigees sont un bruit de fond continu -> faux
           positifs ; les fatales provoquent un BSOD, pas un crash d'appli). Signal de fond calcule
           sur l'ensemble des donnees (independant du filtre periode). Encart ambre en tete de la
           carte Hardware, formule "correlation, pas diagnostic - memtest a envisager". 100% cote
           Dashboard (ExceptionCode 2.1.5 + WHEA etaient deja dans le JSON). Seuil bas volontaire
           (>=1 + >=1), a ajuster a l'usage.
    v2.1.12 : Panneau Top Crashers - section "A investiguer en priorite" + filtre par appli.
           (1) Nouvelle section EN TETE du panneau, alimentee par une liste d'applis suivies
           ($priorityApps, desormais dans config.psd1) : applis metier propres au site, socle
           bureautique/collab, services securite/reseau, etc. - liste definie par chaque site. Ces
           applis remontent TOUJOURS, quelle que soit leur dispersion - le score de dispersion
           (concentre = signal / disperse = bruit) noyait sinon une appli metier qui plante
           modestement sur plusieurs postes (ex une appli metier 10 crashs / 5 PC, classee "bruit ambient",
           repliee par defaut). classifyCrasher route ces applis vers le niveau 'priority' avant
           tout autre classement. Volontairement PAS les satellites de fond (updaters/helpers
           Adobe/Office, famille Dell) qui restent dans le bruit.
           (2) Chaque ligne du panneau est desormais CLIQUABLE : un clic filtre le tableau sur
           les seuls PC ayant cette appli en crash (state.appFilter + chip "Appli : ...", applique
           aussi a l'export CSV). Reutilise le mecanisme de filtre existant. Le nom d'appli (deja
           ConvertTo-HtmlSafe cote embed) transite par data-appname, lu via dataset (aucune eval
           JS), avec le & double-encode pour un aller-retour exact. 100% cote Dashboard (aucun
           changement Collector).
    v2.1.11 : Refonte UX de la remontee d'anomalies. Le bandeau jaune "Anomalies
           detectees" (deplie en tete du rapport, avant les KPIs) est RETIRE : il mettait
           en avant un signal SECONDAIRE (qualite de donnee) au-dessus de l'axe principal
           (sante du parc). Les anomalies sont desormais signalees PAR PC, sur un axe
           volontairement distinct de la sante : (1) un badge orange discret sur la ligne
           du PC (survol = detail des anomalies, clic = filtre), (2) un bouton-filtre
           "Anomalies (N)" dans la barre de filtres (N = PC impactes), necessaire car le
           tableau est pagine - un badge seul ne montrerait pas les PC des autres pages.
           Le filtre reutilise le mecanisme kpiFilter existant. Le calcul des anomalies
           (Test-PCPulseAnomalies) est inchange ; elles sont desormais injectees par PC
           dans l'embed (champ Anomalies, ConvertTo-HtmlSafe) puis mappees cote JS. Aucun
           tri force et pas de categorie dediee dans le drill-down (evite les doublons :
           un PC douteux est deja classe par son score, un BSOD deja dans Stabilite, un
           crasher deja dans Crashers).
    v2.1.10 : Anomalie TruncatedAtCollector - message rendu FACTUEL. L'ancien texte
           affirmait "PC en boucle d'erreur (reboot loop, RAM exhaustion, spam Event
           41 ?)" pour TOUTE troncature, ce qui etait faux dans la majorite des cas :
           un TopCrashers tronque signifie juste qu'un poste a plus de 10 applications
           distinctes en echec (bruit applicatif normal, ou un agent qui crashe pour de
           vrai comme Nexthink), sans aucun rapport avec un reboot loop - constate sur
           des postes parfaitement sains (0 Event 41, 0 warning RAM, 1 boot/jour). Le
           message se contente desormais de constater la troncature ("rapport partiel
           pour ce PC") et de pointer, PAR ARRAY, ou regarder le vrai signal (liste des
           crashers pour TopCrashers, onglet Stabilite pour Events/BSOD, etc.). Aucune
           affirmation non verifiee. La troncature reste une info de VOLUMETRIE, pas un
           diagnostic de sante. (Limites Collector inchangees : TopCrashers 10, Events
           et ResourceWarnings 200 - un relevement de la limite TopCrashers cote
           Collector reste une option si le bruit persiste.)
    v2.1.9 : RETRAIT de la detection "SimilarHostname" (ANOMALIE #4). Une fois le fix
           2.1.8 en place, elle s'est revelee inadaptee au parc : la nomenclature etant
           SEQUENTIELLE (ex. PC-001/PC-002, PC-041/PC-042...), toute paire de machines voisines
           a une distance Levenshtein de 1 -> ~276 faux positifs sur 96 PC, qui noyaient
           les vraies anomalies (troncature Collector, timestamps futurs, valeurs
           aberrantes...). L'heuristique "distance <= 1 = suspect" ne peut structurellement
           pas coexister avec une numerotation incrementale, et le scenario vise (spoofing
           par nom RESSEMBLANT) est marginal : un JSON malveillant usurpe le nom EXACT.
           Fonction Get-LevenshteinDistance + boucle + generation d'anomalie retirees ;
           le rendu generique des anomalies (autres types) est inchange. La section 2.1.8
           reste dans l'historique si on veut un jour reintroduire un filtre homoglyphe.
    v2.1.8 : FIX - Get-LevenshteinDistance plantait sur chaque paire de hostnames
           comparee ("op_Subtraction sur [System.Object[]]" qui defilait). Cause : un
           piege de parsing PowerShell - "$d[$i-1, $j]" etait interprete comme
           "$d[$i - (1, $j)]" (soustraction entier - tableau) a cause de la virgule
           dans l'index. Les expressions arithmetiques des index 2D sont desormais
           parenthesees : $d[($i-1), $j], $d[$i, ($j-1)], $d[($i-1), ($j-1)].
           Double effet : plus d'erreurs a l'ecran, ET la detection de hostnames
           quasi-identiques (typo de renommage AD / spoofing par confusion) refonctionne
           - elle etait morte depuis toujours (l'erreur non-terminante vidait $dist).
    v2.1.7 : UI - deux debordements d'affichage en colonne etroite (onglet MATERIEL).
           - Carte Memoire : un Manufacturer de barrette sans espaces (ex "00000000...",
             rempli tel quel par certains BIOS) transpercait la colonne. .ram-module-meta
             passe en min-width:0 + overflow-wrap:anywhere (le flex item retrecit et casse
             le mot au besoin au lieu de deborder).
           - Section CPU Throttling : en colonne etroite, "cumul <duree>" et "xN events"
             debordaient / se chevauchaient. .throttle-row passe en flex-wrap, .throttle-type
             en min-width:0, .throttle-count en nowrap : le bloc duree/events passe proprement
             a la ligne si la largeur manque.
           Cosmetique pur, aucune logique ni donnee touchee.
    v2.1.6 : SECURITE (durcissement XSS, defense en profondeur). Le rendu client
           injecte l'embed JSON dans un bloc <script> puis fait du innerHTML ; un
           endpoint compromis (adversaire du modele de menace) pouvait forger des
           champs string pour casser le bloc script ou injecter du HTML.
           - Couche A : le ConvertTo-Json de l'embed passe en -EscapeHandling EscapeHtml.
             Encode < > & en sequences \uXXXX dans le blob -> plus aucun </script>
             litteral possible, quel que soit le champ (present, futur, ou oublie au
             re-mapping). Ferme le vecteur d'execution de code. Aucun impact d'affichage
             (valeurs identiques apres JSON.parse cote client).
           - Couche B : les champs string pass-through qui atteignent le DOM sont
             desormais ConvertTo-HtmlSafe au re-mapping. Texte : BSODs.Nom, Crashes.Type,
             CrashCause, Boots.PrecedentType/Method/BootType, ResourceWarnings.Type,
             CPUAgeCategory, ConnectionType. Dates : tous les Timestamp / DateBoot /
             FirstSeen / LastSeen / Day / CollectedAt / LastBoot. Neutralise l'injection
             via innerHTML. Sans regression : les valeurs legitimes (libelles, dates ISO)
             n'ont aucun caractere encodable, comparaisons et parsing JS preserves.
           - Backlog (durcissement typage embed) : cast strict des champs numeriques non
             types (EventId, Count, temperatures, CycleCount, ChassisType...) pour fermer
             le residuel innerHTML sur les nombres. Risque mineur : l'execution de code
             est deja fermee par la couche A, et l'exfil est bornee par la CSP.
    v2.1.5 : FIX - le champ Type des TopCrashers etait perdu au re-mapping PS->embed
           (le bloc $crasherList ne recopiait que AppName et CrashCount). Consequence :
           c.Type / tc.Type toujours undefined cote JS -> tout traite en crash (fallback
           legacy), donc la section "Applis en echec recurrent" (2.1.4) restait vide en
           permanence, le badge du drill-down n'apparaissait jamais, et les echecs
           applicatifs (ex Bing Wallpaper) repartaient dans le scoring par dispersion.
           Type est restaure, NORMALISE en whitelist ('app_failure' ou 'crash' par
           defaut) : un endpoint compromis ne peut mettre qu'une valeur prevue. La
           feature 2.1.4 est desormais reellement active. Aucune autre logique touchee.
    v2.1.4 : Top Crashers - qualification "echec applicatif recurrent" (exploite le
           champ Type du Collector 2.1.3). Un crasher n'est plus classe sur sa seule
           dispersion : sa NATURE prime.
           - classifyCrasher recoit le Type. Un 'app_failure' (echec d'install/maj qui
             boucle, ex Bing Wallpaper) part en categorie 'appfail' quelle que soit sa
             concentration, au lieu de remonter en "Signal local" a tort. Type absent
             (JSON d'un Collector < 2.1.3) = traite en crash (retrocompatible).
           - Panneau global : nouvelle section "Applis en echec recurrent" entre
             "Problemes repartis" et "Bruit ambient" (registre orange "a reparer").
           - Drill-down PC : badge "echec recurrent" distinct du badge "bruit".
           - Les blacklists Hard/Soft restent (vrais crashes inactionnables) ; le Type
             est complementaire ET generique (plus besoin de blacklister a la main
             chaque appli bavarde qui echoue sous l'ID 1000).
    v2.1.3 : Mode tache planifiee pour la regeneration automatique.
           - Parametre -OutputPath : chemin de sortie fixe (ecrase a chaque run),
             pour publication. Sans -OutputPath, comportement interactif inchange
             (fichier horodate dans TEMP + ouverture navigateur).
           - Parametre -NoLaunch : supprime l'ouverture navigateur (Start-Process)
             en contexte non interactif (tache planifiee).
           - Ecriture atomique de la sortie (.tmp puis rename) : une consultation
             ne tombe jamais sur un HTML partiel pendant la generation.
           - meta-refresh (600s) injecte UNIQUEMENT en mode -OutputPath : la page
             publiee se recharge seule ; en interactif, pas de refresh (ne casse
             pas les filtres/scroll de l'utilisateur). SchemaVersion inchange.
    v2.1.2 : Passe Dashboard - corrections de rendu / UI.
           - Z-index du popover KPI : au survol d'une carte (.kpi-group-btn),
             le transform applique cree un stacking context qui isole le
             popover ; celui-ci passait donc SOUS le tableau et les boutons
             de pagination. Corrige en hissant le bouton survole (z-index:100).
           - Seuil d'age des ecrans secondaires rendu ajustable via un curseur
             (3-10 ans, defaut 7, persiste en localStorage 'pcpulse_screenAge').
             Pilote de facon coherente les 3 usages : carte "Ecrans >= N ans",
             sous-KPI "Ecrans ages" du popover USURE, et marquage dans le detail
             par PC. Recalcul live via render() (JS pur, aucun changement
             Collector : les ages AgeYears sont deja cote client).
           - "Top Crashers parc global" deplace tout en bas de la page
             (apres le panneau ecrans secondaires).
           - Detail par PC : exploitation des nouveaux champs du Collector
             v2.1.1. Lecture PS de Thermal.Zone et CPUThrottling.TotalSeconds
             (etaient droppes au re-mapping). Affichage : duree cumulee de
             bridage par jour (signal principal, le compte d'events etant
             gonfle par coeur), zone ACPI a cote de la temperature sur les
             events thermiques. typeShort corrige pour l'Event 37 (affichait
             "Event 37"), et message de severite "refroidissement a verifier"
             remplace par "bridage CPU par firmware" (le 37 = limitation
             firmware, souvent power policy, pas forcement thermique).
    v2.1.1 : Theme sombre neutralise (noir profond). Aucun impact fonctionnel
           ni sur le contrat JSON (SchemaVersion inchange).
           - Surfaces du theme dark (:root) passees du bleu-nuit a un noir
             neutre : --bg-main #0a0a0a, --bg-panel #161616, --border #2b2b2b.
             Corrige aussi --bg-elevated (row-detail) qui etait plus sombre
             que le fond -> remis au premier plan (#202020).
           - Gris de texte neutralises (suppression de la micro-teinte bleue).
           - Accent (--accent violet) et couleurs semantiques inchanges.
           - Le theme [data-theme="light"] n'est pas touche.
    v2.1 : Detection d'anomalies + hardening lecture JSON (consolidation
           des deux blocs de travaux v2.1 avant publication GitHub).
           --- Detection d'anomalies (warnings, JSON gardes dans le rapport) ---
           - Nouvelle fonction Test-PCPulseAnomalies appelee apres Test-PCPulseJson.
             6 categories detectees :
               1. Timestamp dans le futur (CollectedAt / DerniereActivite)
                  -> drift NTP, RTC HS, ou JSON antidate volontairement
               2. PC zombie (DerniereActivite > 30 jours)
                  -> PC probablement parti du parc, a faire le menage
               3. Volumetrie suspecte (taille fichier > 5 MB, arrays
                  TopCrashers/BSODs/BootDurations/Events/ResourceWarnings
                  anormalement longs)
               4. (RETIRE en v2.1.9) Detection Levenshtein des hostnames similaires
                  -> supprimee : inadaptee a une nomenclature de parc sequentielle
               5. Valeurs aberrantes (BatteryHealthPct hors [0;100],
                  DiskWearPct hors [0;100], UptimeDays > 365)
               6. Arrays tronques au Collector (lecture Meta.TruncatedArrays)
                  -> volumetrie superieure aux limites : rapport partiel pour ce
                  PC. Le signal de sante reel est dans le contenu (crashers,
                  onglet Stabilite), pas dans le flag de troncature.
           - Nouveau panneau HTML "Anomalies detectees" (jaune) en tete du
             rapport, distinct du panneau "JSON rejetes" (rouge) :
                * Rejets   = JSON refuse, donnee perdue (Test-PCPulseJson)
                * Anomalies = JSON accepte mais flagge (Test-PCPulseAnomalies)
           - Seuils configurables via $AnomalyThresholds en tete du script.
           - Comme pour les rejets : indicateur fort pour le pentester que
             ses tentatives sont detectees et journalisees.
           --- Hardening lecture JSON (defense en profondeur securite) ---
           - Whitelist STRICTE des SchemaVersions : passage de @('2.0') a
             @('2.1'). Les Collector v2.0 sont rejetes (release big bang
             coordonnee, parc cleane avant deploiement v2.1).
           - Nouveau NIVEAU 0 dans Test-PCPulseJson : verification de la
             taille du fichier AVANT Get-Content. Garde-fou anti-DoS contre
             un PC compromis qui generait un JSON gigantesque pour saturer
             la memoire du Dashboard.
                * Hard limit ($MaxJsonSizeBytes = 10 MB)  -> rejet Critical
                * Soft limit ($WarnJsonSizeBytes = 2 MB)  -> Write-Host warning
             Distinction importante : le NIVEAU 0 est une mesure de
             SECURITE (anti-DoS), l'anomalie #3 OversizedJson (5 MB) est
             une mesure de QUALITE (signal pour l'admin). Les seuils sont
             complementaires, pas redondants.
           - Lecture du nouveau bloc Meta.TruncatedArrays produit par le
             Collector v2.1. Si flag a $true sur n'importe quel array,
             genere une anomalie "TruncatedAtCollector". Compatible v2.0
             si bloc Meta absent (test PSObject.Properties).
    v2.0 : Sanity-checks stricts + sanitization + CSP (release security hardening)
           - Whitelist stricte des SchemaVersions acceptees ($AcceptedSchemaVersions = @('2.0')).
             Tout JSON sans SchemaVersion = '2.0' est REJETE (plus de tolerance
             retro-compat sur les anciens schemas).
           - Validation structurelle a 4 niveaux (JSON parsing, SchemaVersion,
             Machine.PC, match nom fichier vs Machine.PC).
           - HTML-encoding (Get-SafeString) sur toutes les valeurs strings
             injectees dans le HTML genere : empeche XSS via JSON corrompu.
           - Match strict nom de fichier vs Machine.PC : un PC compromis ne
             peut pas se faire passer pour un autre PC du parc (anti-spoofing).
           - Nouvelle section HTML "JSON suspects" affichee si rejets detectes.
             Visible par l'admin pour audit + indicateur pentest fort.
           - Content-Security-Policy stricte injectee dans le <head> du HTML
             genere : derniere ligne de defense au cas ou la sanitization
             aurait rate quelque chose. Bloque toute communication reseau,
             scripts/styles externes, iframes, formulaires - dans un rapport
             qui n'en a aucun besoin par design.
           - Voir SECURITY.md pour le trust model complet.
    v1.8 : Exposition des donnees hardware v1.8 du Collector
           - Panel Materiel enrichi de 3 nouveaux blocs :
             * RAM : installee / slots / capacite max + detail barrettes
                     (Slot, Type DDR, Vitesse, Fabricant).
             * GPU : nom + version + date driver pour chaque adapteur.
             * CPU Throttling : agregation quotidienne des events 35/55
                    Kernel-Processor-Power, avec badge alerte >=3 jours.
           - Backward-compat totale : si JSON v1.7 ou anterieur, les
             nouveaux blocs affichent discretement "Donnees non
             disponibles (Collector < v1.8)".
    v1.7 : Debruitage des Top Crashers (Signal vs Bruit)
           - Scoring automatique de chaque processus crasheur :
             score = crashs_par_PC / sqrt(PC_impactes).
             Penalise la dispersion (un crash reparti sur 10 PC est
             noye, un crash concentre sur 1 PC ressort).
           - 3 sections dans le panel Top Crashers parc global :
             * Signaux locaux (score >=3) : a investiguer prioritairement
             * Problemes repartis (2<=score<3) : bug applicatif probable
             * Bruit ambient (score <2) : replie par defaut
           - Blacklist Hard : processus totalement ignores (ex: le
             fameux microsoftsearchbing.exe qui pourrissait deja NXT).
           - Blacklist Soft : processus forces en section Bruit meme
             si leur score les aurait remontes (dellosd, shellexpe,
             gamebar, ASUS utilitaires, Dell techhub).
           - Drill-down par PC : meme logique locale, bruiteurs grises
             avec badge "bruit" au lieu d'etre cache.
    v1.6 : Intelligence de diagnostic (Wave 1)
           - Bandeau verdict global par PC (4 niveaux : Sain / A surveiller
             / Incident probable / Critique) calcule selon seuils sur
             crashs, WHEA, batterie, CPU age, SMART, bursts I/O.
           - Section "Signaux croises" (5 patterns temporels a fenetre
             10 min) : burst I/O -> hard crash, WHEA PCIe -> crash, etc.
           - Nom CPU ajoute dans le panel Materiel.
           - Libelles detailles des Hard crashes (Coupure alim / Reprise
             veille / User bouton power) exploitant CrashCause v1.6.
           - Schema accepte elargi a ('1.4','1.5','1.6').
    v1.5 : Rendu des clusters Event 51 (Disk slow / I/O timeout)
           - Propagation des champs Count/IsBurst/FirstSeen/LastSeen
             produits par le Collector v1.5 jusqu'au rendu JS.
           - Affichage enrichi : "xN events en 1s" + badge BURST si
             Count >= 50 (signal materiel fort).
           - Backward-compat JSON v1.4- : affichage inchange.
           - Schema accepte elargi a ('1.4','1.5').
    Voir CHANGELOG.md du repo pour l'historique des versions.
.EXAMPLE
    .\02_Dashboard.ps1
    # Utilise le SharePath par defaut (C:\PCPulse)
.EXAMPLE
    .\02_Dashboard.ps1 -SharePath "D:\Data\pcpulse"
    # Utilise un chemin custom
.EXAMPLE
    .\02_Dashboard.ps1 -FiltrePC "LAPTOP-*"
    # Ne charge que les PC dont le nom commence par LAPTOP-
#>

param(
    [string]$SharePath  = 'C:\PCPulse',
    [string]$FiltrePC   = '*',
    [string]$OutputPath = '',
    [switch]$NoLaunch,
    # v2.5.4 : 2e source cloud. Le SAS (secret de LECTURE) se passe en PARAMETRE
    # (commande de la tache planifiee, admin-only) et JAMAIS dans release\config.psd1
    # (lisible par les postes). CloudDepotBaseUrl (non secret) peut venir du param ou du config.
    [string]$CloudDepotBaseUrl = '',
    [string]$CloudDepotSas     = ''
)

# ============================================================
# CONSTANTES
# ============================================================

# Whitelist STRICTE des SchemaVersions acceptees.
# v2.2 : @('2.1', '2.2'). Le rename EDR (schema 2.2) se deploie poste-par-poste,
# donc pendant le rollout le share contient un melange de JSON 2.1 (postes pas
# encore a jour) et 2.2. On accepte les deux pour n'exclure aucun poste ; un JSON
# 2.1 affiche juste l'EDR "en attente" (le Dashboard ne lit plus l'ancien champ,
# cf. de-vendoring). Les Collector <= 2.0 restent rejetes. Voir SECURITY.md.
$AcceptedSchemaVersions = @('2.1', '2.2')

# v2.1 : Limites de taille JSON (defense en profondeur SECURITE).
# - Hard limit 10 MB : rejet AVANT Get-Content. Garde-fou anti-DoS
#   contre un PC compromis qui generait un JSON enorme pour saturer
#   la memoire du Dashboard. Le rejet est applique au NIVEAU 0 de
#   Test-PCPulseJson (premier check, avant le parsing).
# - Soft limit 2 MB  : Write-Host warning seulement, JSON traite
#   normalement. Aide a reperer une derive precoce sans bruit user.
# A distinguer de l'anomalie #3 OversizedJson (5 MB) qui est une
# mesure de QUALITE (panneau jaune visible utilisateur). Les deux
# seuils sont complementaires :
#   0-2 MB    : silence
#   2-5 MB    : log warning console
#   5-10 MB   : anomalie soft (panneau jaune utilisateur)
#   > 10 MB   : rejet hard securite (panneau rouge utilisateur)
$MaxJsonSizeBytes  = 10MB
$WarnJsonSizeBytes = 2MB

# v2.0 : Pattern strict pour les noms de PC valides.
# - Alphanumerique, tirets, underscores, points
# - Longueur 1 a 63 caracteres (max NetBIOS hostname)
# - Sert a la fois pour Machine.PC et le nom de fichier .json
$ValidPCNamePattern = '^[A-Za-z0-9_.-]{1,63}$'

# v2.2.0 : la liste des applis suivies ("A investiguer en priorite") est desormais
# pilotee par config (cle PriorityApps de config.psd1, defaut generique dans
# $DefaultConfig). Voir $priorityApps, affecte juste apres le chargement de la config.
# De-vendorise : aucune appli specifique au site (metier / securite) codee en dur ici.

# v2.0+ : Seuils pour la detection d'anomalies (warnings, pas rejets)
# Voir SECURITY.md > "Detection d'anomalies"
$AnomalyThresholds = @{
    # Anomalie 1 : timestamp dans le futur
    # Tolerance de 5 minutes pour absorber les drifts NTP normaux entre PC
    FutureMinutesTolerance = 5

    # Anomalie 2 : PC zombie (n'a pas check-in depuis longtemps)
    # 30 jours = au-dela, le PC est probablement parti du parc
    ZombieDaysThreshold = 30

    # Anomalie 3 : volumetrie suspecte
    # Taille max d'un JSON considere normal (au-dela = potentiel flood)
    MaxJsonSizeMB = 5
    # Nombre max d'entrees dans certaines collections "intentionnellement bornees"
    # Volontairement plus larges que les caps Collector v2.1 (10/10/100/200/200)
    # pour ne se declencher QUE sur Collector compromis (= JSON forge a la main
    # qui bypasse les caps officiels). Ceintures et bretelles.
    MaxTopCrashersCount      = 50
    MaxBSODsCount            = 100
    MaxBootDurationsCount    = 200
    MaxEventsCount           = 500   # v2.1 : nouveau cap Collector (200), seuil anomalie 500
    MaxResourceWarningsCount = 500   # v2.1 : idem

    # Anomalie 5 : valeurs aberrantes
    # Battery > 100% ou < 0% est physiquement impossible
    # Disk wear > 100% est aberrant
    # Uptime > 365 jours = PC qui n'a jamais reboot depuis 1 an = tres suspect
    MaxUptimeDays = 365
}

# v2.4.3 : config.psd1 lu en PRIORITE dans release\ (lecture seule pour les postes,
# durcissement #5). Repli sur la racine pour migration douce.
$ConfigFile       = Join-Path $SharePath 'release\config.psd1'
if (-not (Test-Path $ConfigFile)) { $ConfigFile = Join-Path $SharePath 'config.psd1' }
# v2.1.3 : sortie parametrable. -OutputPath => chemin fixe (mode tache, ecrase a chaque run, pour publication).
# Sans -OutputPath => fichier horodate dans TEMP (mode interactif, comportement historique).
$OutputHTML       = if ($OutputPath) { $OutputPath } else { Join-Path $env:TEMP ("PCPulse-Dashboard-" + (Get-Date -Format 'yyyyMMdd-HHmm') + '.html') }

# Valeurs par defaut si config.psd1 est absent / corrompu
$DefaultConfig = @{
    SeuilBootLong        = 2
    SeuilCrashRecent     = 7
    SeuilDiskAlert       = 10
    SeuilDiskWarning     = 25
    SeuilOfflineJours    = 1
    DashboardTitle       = 'PCPulse'
    DashboardSubtitle    = 'Supervision du parc'
    CsvRanges            = 'ip-ranges.csv'
    MaskHealthyByDefault = $false

    # v2.5.3 : seuil de renouvellement (annee CPU <= seuil => "Candidat au
    # renouvellement", sinon "Parc courant"). $null (defaut) => on garde l'ancien
    # tag Recent/Vieillissant/Ancien (retrocompat). Voir config.psd1.example.
    RenewalCandidateMaxYear = $null

    # v2.5.4 : 2e source CLOUD (postes 100% distants). Le Dashboard lit AUSSI le
    # conteneur Azure "depot" (via SAS read+list) et fusionne avec le share en
    # gardant, par machine, le rapport le plus recent. Vides => cloud desactive.
    CloudDepotBaseUrl = ''   # ex: https://<compte>.blob.core.windows.net/depot
    # NB : le SAS de lecture ne vit PAS dans la config (lisible par les postes) ->
    # il se passe en PARAMETRE -CloudDepotSas au Dashboard (tache planifiee, admin-only).

    # v2.4.0 : dossier ou vit le registre de decommission (ecrit par
    # Decommission-PC.ps1 dans un dossier ecrivable par les techs, hors share
    # durci). Le generateur du Dashboard doit pouvoir LIRE ce dossier. Vide =>
    # on cherche le registre a cote des JSON (retro-compat).
    DecommissionRegistryPath = ''

    # v2.2.0 : applis suivies ("A investiguer en priorite") - de-vendorise, pilote par config.
    # Defaut generique = socle bureautique/collab (universel). Les applis SPECIFIQUES au site
    # (metier maison, services securite/reseau, Java metier...) vivent dans config.psd1, jamais
    # dans le code publie. Le merge remplace cette liste par celle du config.psd1 si presente.
    PriorityApps = @(
        'outlook.exe', 'olk.exe', 'winword.exe', 'excel.exe',
        'ms-teams.exe', 'onedrive.exe', 'acrobat.exe', 'indesign.exe',
        'chrome.exe', 'msedge.exe'
    )

    ScoreWeights         = @{
        BSOD          = 5
        WHEA          = 4
        Crash         = 3
        Thermal       = 3
        GPU_TDR       = 2
        DiskAlert     = 2
        BootLong      = 1
        Offline       = 5
        # v5.3 : EDR down = critique, batterie usee = mineur
        EDRDown       = 5
        Battery       = 1
        # v5.4 : BootPerf et SMART
        BootPerfSlow  = 1
        DiskHealth    = 3
    }
}

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

function Import-MonitorConfig {
    param([string]$Path, [hashtable]$Defaults)
    if (-not (Test-Path $Path)) {
        Write-Host "[!] Config absente ($Path), utilisation des defaults" -ForegroundColor Yellow
        return $Defaults.Clone()
    }
    try {
        $loaded = Import-PowerShellDataFile -Path $Path -ErrorAction Stop
        $merged = @{}
        foreach ($key in $Defaults.Keys) {
            if (-not $loaded.ContainsKey($key)) { $merged[$key] = $Defaults[$key]; continue }
            # v2.3.1 (#4) : pour un sous-objet hashtable (ex ScoreWeights), merge
            # RECURSIF cle par cle - un override PARTIEL ne doit plus effacer les
            # cles absentes (sinon des poids manquants = score casse en silence).
            # Les listes (MonitoredServices, PriorityApps) gardent la semantique de
            # REMPLACEMENT (la liste du config remplace entierement le defaut).
            if ($Defaults[$key] -is [hashtable] -and $loaded[$key] -is [hashtable]) {
                $sub = $Defaults[$key].Clone()
                foreach ($sk in $loaded[$key].Keys) { $sub[$sk] = $loaded[$key][$sk] }
                $merged[$key] = $sub
            } else {
                $merged[$key] = $loaded[$key]
            }
        }
        Write-Host "[+] Config chargee : $Path" -ForegroundColor Green
        return $merged
    } catch {
        Write-Host "[!] Erreur lecture config ($_), utilisation des defaults" -ForegroundColor Yellow
        return $Defaults.Clone()
    }
}

function ConvertTo-HtmlSafe {
    param($Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Text)
}

function Build-CIDRRange {
    param([string]$CIDR, [string]$Site)
    try {
        $parts   = $CIDR.Split('/')
        $prefix  = [int]$parts[1]
        $netRaw  = [System.Net.IPAddress]::Parse($parts[0]).GetAddressBytes()

        $maskBytes = [byte[]]::new(4)
        for ($i = 0; $i -lt 4; $i++) {
            $bitsInByte = [math]::Min(8, [math]::Max(0, $prefix - ($i * 8)))
            if     ($bitsInByte -eq 0) { $maskBytes[$i] = 0 }
            elseif ($bitsInByte -eq 8) { $maskBytes[$i] = 255 }
            else                       { $maskBytes[$i] = [byte](256 - [math]::Pow(2, 8 - $bitsInByte)) }
        }

        $netBytes = [byte[]]::new(4)
        for ($i = 0; $i -lt 4; $i++) {
            $netBytes[$i] = $netRaw[$i] -band $maskBytes[$i]
        }

        return [PSCustomObject]@{
            CIDR      = $CIDR
            Site      = $Site
            NetBytes  = $netBytes
            MaskBytes = $maskBytes
        }
    } catch {
        Write-Warning "Range ignoree ($CIDR) : $_"
        return $null
    }
}

function Get-SiteFromIP {
    param([string]$IP, [array]$Ranges)
    if (-not $IP -or $IP -eq 'N/A' -or $IP -like '169.254.*' -or $IP -like '127.*') {
        return 'Inconnu'
    }
    try {
        $ipBytes = [System.Net.IPAddress]::Parse($IP).GetAddressBytes()
    } catch {
        return 'Inconnu'
    }
    foreach ($r in $Ranges) {
        $match = $true
        for ($i = 0; $i -lt 4; $i++) {
            if (($ipBytes[$i] -band $r.MaskBytes[$i]) -ne $r.NetBytes[$i]) {
                $match = $false
                break
            }
        }
        if ($match) { return $r.Site }
    }
    return 'Inconnu'
}

function ConvertTo-Array {
    param($Data)
    if ($null -eq $Data) { return @() }
    if ($Data -is [array]) { return $Data }
    return @($Data)
}

# v2.3.1 (#2, XSS passe 2) : cast STRICT des champs numeriques du re-mapping.
# Defense en profondeur - un endpoint compromis pourrait forger une chaine
# (voire du markup) dans un champ cense etre un nombre. On force le type ;
# une valeur non convertible OU absente devient $null (jamais 0 : on preserve
# la distinction "absent" vs "zero", et le JS gere deja null). ConvertTo-Json
# emet alors un vrai nombre ou null, jamais une chaine htmlsafe-ee.
function ConvertTo-IntOrNull {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    try { return [int]$Value } catch { return $null }
}
function ConvertTo-NumOrNull {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    try { return [double]$Value } catch { return $null }
}

# ============================================================
# CHARGEMENT DE LA CONFIG
# ============================================================
$cfg = Import-MonitorConfig -Path $ConfigFile -Defaults $DefaultConfig

$SeuilBootLong        = [int]$cfg.SeuilBootLong
$SeuilCrashRecent     = [int]$cfg.SeuilCrashRecent
$SeuilDiskAlert       = [int]$cfg.SeuilDiskAlert
$SeuilDiskWarning     = [int]$cfg.SeuilDiskWarning
$SeuilOfflineJours    = [int]$cfg.SeuilOfflineJours
$DashboardTitle       = [string]$cfg.DashboardTitle
$DashboardSubtitle    = [string]$cfg.DashboardSubtitle
$MaskHealthyByDefault = [bool]$cfg.MaskHealthyByDefault
$ScoreWeights         = $cfg.ScoreWeights
$priorityApps         = @($cfg.PriorityApps)
# v2.5.3 : seuil de renouvellement (annee CPU <= seuil => "Candidat au renouvellement").
# $null => mode desactive, on garde l'ancien tag Recent/Vieillissant/Ancien.
$renewalMaxYear       = if ($null -ne $cfg.RenewalCandidateMaxYear) { [int]$cfg.RenewalCandidateMaxYear } else { $null }

# Resolution du chemin CSV
$csvRel = [string]$cfg.CsvRanges
if ([System.IO.Path]::IsPathRooted($csvRel)) {
    $CsvRanges = $csvRel
} else {
    $CsvRanges = Join-Path $SharePath $csvRel
}

# ============================================================
# LECTURE DU CSV DE RANGES IP
# ============================================================
$rangesIP = @()
if (Test-Path $CsvRanges) {
    Write-Host "[*] Chargement des ranges IP depuis $CsvRanges" -ForegroundColor Cyan
    $raw = Import-Csv -Path $CsvRanges -Delimiter ',' |
        Where-Object { $_.Pattern1 -like '*/*' }

    foreach ($r in $raw) {
        $built = Build-CIDRRange -CIDR $r.Pattern1.Trim() -Site $r.Entity.Trim()
        if ($built) { $rangesIP += $built }
    }
    Write-Host "[+] $($rangesIP.Count) range(s) chargee(s) et pre-compilee(s)" -ForegroundColor Green
} else {
    Write-Host "[!] CSV de ranges non trouve : $CsvRanges (colonne Site desactivee)" -ForegroundColor Yellow
}
$showSite = ($rangesIP.Count -gt 0)

# ============================================================
# LECTURE DU REGISTRE DE DECOMMISSION (v2.4.0)
# Registre SEPARE (jamais dans le JSON du poste, que le Collector reecrit a
# chaque cycle). Ecrit par Decommission-PC.ps1. Cle = nom PC. Lecture seule ici.
# ============================================================
# Registre ecrit par Decommission-PC.ps1 dans un dossier ecrivable par les techs
# (hors share durci). Chemin pilote par DecommissionRegistryPath (config) ; a
# defaut on cherche a cote des JSON (retro-compat).
$decomDir = [string]$cfg.DecommissionRegistryPath
if (-not $decomDir) { $decomDir = $SharePath }
$decomPath = Join-Path $decomDir 'decommissioning.json'
$decomByPc = @{}
$decomAll  = @()   # v2.4.2 : registre COMPLET (y compris PC dont le JSON a ete purge) pour les stats/audit
if (Test-Path $decomPath) {
    try {
        $decomRaw = Get-Content $decomPath -Raw -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($decomRaw)) {
            foreach ($d in @($decomRaw | ConvertFrom-Json)) {
                if ($d.PC) {
                    $decomByPc[[string]$d.PC] = $d
                    # Champs utiles aux stats uniquement (pas d'historique verbeux embarque).
                    $decomAll += [PSCustomObject]@{
                        PC         = ConvertTo-HtmlSafe ([string]$d.PC)
                        Statut     = ConvertTo-HtmlSafe ([string]$d.Statut)
                        Reason     = ConvertTo-HtmlSafe ([string]$d.Reason)
                        AssignedTo = ConvertTo-HtmlSafe ([string]$d.AssignedTo)
                        MarkedAt   = ConvertTo-HtmlSafe ([string]$d.MarkedAt)
                        TargetDate = ConvertTo-HtmlSafe ([string]$d.TargetDate)
                        DoneAt     = ConvertTo-HtmlSafe ([string]$d.DoneAt)
                        DoneBy     = ConvertTo-HtmlSafe ([string]$d.DoneBy)
                    }
                }
            }
        }
        Write-Host "[+] Registre decommission : $($decomByPc.Count) entree(s)" -ForegroundColor Green
    } catch {
        Write-Host "[!] Registre decommission illisible ($_)" -ForegroundColor Yellow
    }
}

# ============================================================
# v2.0 : SANITY-CHECKS - fonctions de validation et sanitization
# ============================================================
# Ces fonctions sont le coeur du hardening pentest v2.0. Elles permettent :
#   1. De rejeter les JSON malformes / spoofs / aux schemas inconnus
#   2. De neutraliser les valeurs string suspectes (HTML-encoding) avant
#      injection dans le rapport HTML genere
#
# Voir SECURITY.md pour le trust model complet et les scenarios d'attaque
# que ces verifications ferment.
# ============================================================

function Get-SafeString {
    <#
    .SYNOPSIS
        HTML-encode une chaine pour empecher l'injection HTML/XSS.
    .DESCRIPTION
        Encode les 5 caracteres dangereux : & < > " '
        Pour toute chaine destinee a etre injectee dans le HTML genere
        (que ce soit dans le JS embarque ou directement dans le DOM),
        on doit passer par cette fonction.

        L'ordre de remplacement est IMPORTANT : & doit etre remplace EN PREMIER
        sinon on encode les ; des autres replacements.
    .EXAMPLE
        Get-SafeString '<script>alert(1)</script>'
        # Retourne : &lt;script&gt;alert(1)&lt;/script&gt;
    #>
    param([string]$InputString)
    if ([string]::IsNullOrEmpty($InputString)) { return '' }
    return $InputString `
        -replace '&', '&amp;' `
        -replace '<', '&lt;' `
        -replace '>', '&gt;' `
        -replace '"', '&quot;' `
        -replace "'", '&#39;'
}

function Test-PCPulseJson {
    <#
    .SYNOPSIS
        Valide un JSON PCPulse selon la politique de securite v2.1.
    .DESCRIPTION
        5 niveaux de validation strict (rejet si KO) :
          0. Taille fichier <= $MaxJsonSizeBytes (10 MB) ?  [v2.1]
             Verification AVANT lecture pour eviter le DoS memoire.
             Warning console si taille >= $WarnJsonSizeBytes (2 MB).
          1. Le fichier parse en JSON ?
          2. SchemaVersion present et dans la whitelist ?
          3. Machine.PC present et alphanumerique valide ?
          4. Machine.PC matche le nom de fichier (anti-spoof) ?

        Apres validation, sanitization en place des champs string les plus
        sensibles (CurrentUser, IPAddress, CPUName, OS, Site).

        Voir SECURITY.md > "Sanity-checks Dashboard" pour le rationnel complet.
    .OUTPUTS
        [pscustomobject] avec les champs :
          - Valid  : [bool]
          - Reason : [string] courte description du rejet
          - Level  : 'Critical' | 'Warning'
          - Detail : [string] info technique pour le diag
          - Data   : [pscustomobject] le JSON parse + sanitise (si Valid=true)
          - File   : [string] nom du fichier
    #>
    param(
        [Parameter(Mandatory)] [string] $FilePath
    )

    $filename = [IO.Path]::GetFileNameWithoutExtension($FilePath)

    # ---- NIVEAU 0 : Taille fichier (v2.1, anti-DoS memoire) ----
    # On verifie la taille AVANT Get-Content pour eviter qu'un PC compromis
    # generant un JSON enorme ne sature la memoire du Dashboard.
    try {
        $fileSize = (Get-Item -Path $FilePath -ErrorAction Stop).Length
    } catch {
        return [pscustomobject]@{
            Valid  = $false
            File   = (Split-Path $FilePath -Leaf)
            Level  = 'Critical'
            Reason = 'Fichier inaccessible'
            Detail = $_.Exception.Message
            Data   = $null
        }
    }
    if ($fileSize -gt $MaxJsonSizeBytes) {
        $sizeMB     = [math]::Round($fileSize / 1MB, 2)
        $maxSizeMB  = [math]::Round($MaxJsonSizeBytes / 1MB, 0)
        return [pscustomobject]@{
            Valid  = $false
            File   = (Split-Path $FilePath -Leaf)
            Level  = 'Critical'
            Reason = 'JSON trop volumineux (rejet anti-DoS)'
            Detail = "$sizeMB MB (limite: $maxSizeMB MB) - PC compromis ou bug Collector ?"
            Data   = $null
        }
    }
    if ($fileSize -gt $WarnJsonSizeBytes) {
        # Soft warning : on continue mais on signale en console.
        # L'anomalie #3 (5 MB) prendra le relais cote utilisateur.
        $sizeMB    = [math]::Round($fileSize / 1MB, 2)
        $warnSizeMB = [math]::Round($WarnJsonSizeBytes / 1MB, 0)
        Write-Host "[!] JSON volumineux : $(Split-Path $FilePath -Leaf) ($sizeMB MB > $warnSizeMB MB warning)" -ForegroundColor DarkYellow
    }

    # ---- NIVEAU 1 : Parsing JSON ----
    try {
        $raw  = Get-Content -Path $FilePath -Raw -Encoding UTF8 -ErrorAction Stop
        $data = $raw | ConvertFrom-Json -ErrorAction Stop
    } catch {
        return [pscustomobject]@{
            Valid  = $false
            File   = (Split-Path $FilePath -Leaf)
            Level  = 'Critical'
            Reason = 'JSON malforme ou illisible'
            Detail = $_.Exception.Message
            Data   = $null
        }
    }

    # ---- NIVEAU 2 : SchemaVersion ----
    if (-not $data.SchemaVersion) {
        return [pscustomobject]@{
            Valid  = $false
            File   = (Split-Path $FilePath -Leaf)
            Level  = 'Critical'
            Reason = 'SchemaVersion absent'
            Detail = 'Le JSON ne contient pas le champ SchemaVersion'
            Data   = $null
        }
    }
    if ($data.SchemaVersion -notin $AcceptedSchemaVersions) {
        return [pscustomobject]@{
            Valid  = $false
            File   = (Split-Path $FilePath -Leaf)
            Level  = 'Critical'
            Reason = 'SchemaVersion non supporte'
            Detail = "trouve='$(Get-SafeString $data.SchemaVersion)', attendu='$($AcceptedSchemaVersions -join "/")'"
            Data   = $null
        }
    }

    # ---- NIVEAU 3 : Machine.PC ----
    if (-not $data.Machine -or -not $data.Machine.PC) {
        return [pscustomobject]@{
            Valid  = $false
            File   = (Split-Path $FilePath -Leaf)
            Level  = 'Critical'
            Reason = 'Machine.PC absent'
            Detail = 'Champ Machine.PC manquant ou vide'
            Data   = $null
        }
    }
    if ([string]$data.Machine.PC -notmatch $ValidPCNamePattern) {
        return [pscustomobject]@{
            Valid  = $false
            File   = (Split-Path $FilePath -Leaf)
            Level  = 'Critical'
            Reason = 'Machine.PC contient des caracteres invalides'
            Detail = "valeur='$(Get-SafeString $data.Machine.PC)' (attendu : alphanumerique, tirets, points, underscores, max 63 chars)"
            Data   = $null
        }
    }

    # ---- NIVEAU 4 : Match fichier <-> Machine.PC (anti-spoof) ----
    if ($data.Machine.PC -ne $filename) {
        return [pscustomobject]@{
            Valid  = $false
            File   = (Split-Path $FilePath -Leaf)
            Level  = 'Critical'
            Reason = 'Mismatch nom de fichier / Machine.PC (spoofing suspect)'
            Detail = "fichier='$filename', JSON='$(Get-SafeString $data.Machine.PC)'"
            Data   = $null
        }
    }

    # ---- ECHAPPEMENT HTML : fait UNE SEULE FOIS, a l'embed ----
    # v2.4.5 : on ne sanitise PLUS en place ici. Ces champs (CurrentUser,
    # LastLoggedUser, IP, CollectorRunAs, CPUName, OS, Site) etaient HTML-encodes
    # ici PUIS re-encodes par ConvertTo-HtmlSafe a l'embed -> DOUBLE encodage
    # ("O'Brien" s'affichait "O&#39;Brien"). Verifie : chacun de ces champs est
    # bien ConvertTo-HtmlSafe a l'embed (point d'echappement unique) et n'est
    # utilise nulle part en brut -> aucune perte de protection XSS.
    # Note : Machine.PC est deja valide par regex.

    return [pscustomobject]@{
        Valid  = $true
        File   = (Split-Path $FilePath -Leaf)
        Level  = $null
        Reason = $null
        Detail = $null
        Data   = $data
    }
}

function Test-PCPulseAnomalies {
    <#
    .SYNOPSIS
        Detecte des anomalies "soft" dans un JSON PCPulse deja valide.
    .DESCRIPTION
        Complement de Test-PCPulseJson : ne rejette pas les JSON, mais
        retourne une liste de signaux suspects qui meritent l'attention
        de l'admin sans pour autant invalider la donnee.

        6 categories detectees :
          1. Timestamp dans le futur (CollectedAt / DerniereActivite)
             -> drift NTP, RTC HS, ou JSON antidate volontairement
          2. PC zombie (DerniereActivite > 30 jours)
             -> PC probablement parti du parc, a faire le menage
          3. Volumetrie suspecte (taille fichier, longueur arrays)
             -> potentiel flood ou bug local
          4. (RETIRE en v2.1.9) Hostnames suspects (detection Levenshtein)
             -> supprimee : parc sequentiel -> faux positifs a chaque voisin
          5. Valeurs aberrantes (HealthPercent hors [0;100], uptime > 1 an, etc.)
             -> bug WMI ou JSON forge
          6. Arrays tronques au Collector (v2.1, lecture Meta.TruncatedArrays)
             -> rapport PARTIEL pour ce PC (volumetrie > limites du Collector).
                Constat factuel, pas un diagnostic : le signal de sante reel
                est dans le contenu (count des crashers, onglet Stabilite),
                pas dans le flag de troncature.

        L'anomalie #4 (similarite hostname) a ete RETIREE en v2.1.9
        (inadaptee a une nomenclature de parc sequentielle).

        Voir SECURITY.md > "Detection d'anomalies" pour le rationale.
    .OUTPUTS
        [array] de [pscustomobject] avec :
          - Type      : code court (ex: 'FutureTimestamp', 'ZombiePC', etc.)
          - Reason    : description courte
          - Detail    : info technique
    #>
    param(
        [Parameter(Mandatory)] [pscustomobject] $Data,
        [Parameter(Mandatory)] [string] $FilePath
    )

    $anomalies = [System.Collections.Generic.List[PSCustomObject]]::new()
    $now = Get-Date

    # --------------------------------------------------------
    # ANOMALIE 1 : Timestamp dans le futur
    # --------------------------------------------------------
    # Tolerance NTP : 5 minutes (les PC peuvent avoir un drift mineur)
    $futureLimit = $now.AddMinutes($AnomalyThresholds.FutureMinutesTolerance)

    foreach ($field in @('CollectedAt')) {
        if ($Data.Machine.PSObject.Properties[$field]) {
            $val = $Data.Machine.$field
            if ($val) {
                try {
                    $parsed = [datetime]::Parse($val)
                    if ($parsed -gt $futureLimit) {
                        $delta = ($parsed - $now)
                        $anomalies.Add([pscustomobject]@{
                            Type   = 'FutureTimestamp'
                            Reason = "Machine.$field dans le futur"
                            Detail = "valeur=$($parsed.ToString('yyyy-MM-dd HH:mm')) (~ $([math]::Round($delta.TotalHours,1))h apres maintenant) - drift NTP ou JSON forge ?"
                        })
                    }
                } catch { }   # parsing rate, on ignore (peut arriver sur des formats exotiques)
            }
        }
    }

    if ($Data.Stats -and $Data.Stats.PSObject.Properties['DerniereActivite']) {
        $val = $Data.Stats.DerniereActivite
        if ($val) {
            try {
                $parsed = [datetime]::Parse($val)
                if ($parsed -gt $futureLimit) {
                    $delta = ($parsed - $now)
                    $anomalies.Add([pscustomobject]@{
                        Type   = 'FutureTimestamp'
                        Reason = "Stats.DerniereActivite dans le futur"
                        Detail = "valeur=$($parsed.ToString('yyyy-MM-dd HH:mm')) (~ $([math]::Round($delta.TotalHours,1))h apres maintenant)"
                    })
                }
            } catch { }
        }
    }

    # --------------------------------------------------------
    # ANOMALIE 2 : PC zombie (CollectedAt trop ancien)
    # --------------------------------------------------------
    if ($Data.Machine.PSObject.Properties['CollectedAt']) {
        $val = $Data.Machine.CollectedAt
        if ($val) {
            try {
                $parsed = [datetime]::Parse($val)
                $age = ($now - $parsed)
                if ($age.TotalDays -gt $AnomalyThresholds.ZombieDaysThreshold) {
                    $anomalies.Add([pscustomobject]@{
                        Type   = 'ZombiePC'
                        Reason = "PC inactif depuis longtemps"
                        Detail = "derniere collecte il y a $([math]::Round($age.TotalDays,0)) jours (seuil: $($AnomalyThresholds.ZombieDaysThreshold)j) - probablement sorti du parc"
                    })
                }
            } catch { }
        }
    }

    # --------------------------------------------------------
    # ANOMALIE 3 : Volumetrie suspecte
    # --------------------------------------------------------
    try {
        $sizeMB = [math]::Round((Get-Item $FilePath).Length / 1MB, 2)
        if ($sizeMB -gt $AnomalyThresholds.MaxJsonSizeMB) {
            $anomalies.Add([pscustomobject]@{
                Type   = 'OversizedJson'
                Reason = "Taille JSON anormale"
                Detail = "$sizeMB MB (seuil: $($AnomalyThresholds.MaxJsonSizeMB) MB) - flood ou bug local ?"
            })
        }
    } catch { }

    # Volumetrie des arrays "intentionnellement bornes"
    if ($Data.PSObject.Properties['TopCrashers'] -and $Data.TopCrashers) {
        $cnt = @($Data.TopCrashers).Count
        if ($cnt -gt $AnomalyThresholds.MaxTopCrashersCount) {
            $anomalies.Add([pscustomobject]@{
                Type   = 'OversizedArray'
                Reason = "TopCrashers anormalement long"
                Detail = "$cnt entrees (seuil: $($AnomalyThresholds.MaxTopCrashersCount)) - PC qui spam des crashs ?"
            })
        }
    }
    if ($Data.PSObject.Properties['BSODs'] -and $Data.BSODs) {
        $cnt = @($Data.BSODs).Count
        if ($cnt -gt $AnomalyThresholds.MaxBSODsCount) {
            $anomalies.Add([pscustomobject]@{
                Type   = 'OversizedArray'
                Reason = "BSODs anormalement long"
                Detail = "$cnt entrees (seuil: $($AnomalyThresholds.MaxBSODsCount)) - PC en panne severe ou JSON forge ?"
            })
        }
    }
    if ($Data.PSObject.Properties['BootDurations'] -and $Data.BootDurations) {
        $cnt = @($Data.BootDurations).Count
        if ($cnt -gt $AnomalyThresholds.MaxBootDurationsCount) {
            $anomalies.Add([pscustomobject]@{
                Type   = 'OversizedArray'
                Reason = "BootDurations anormalement long"
                Detail = "$cnt entrees (seuil: $($AnomalyThresholds.MaxBootDurationsCount)) - HistoriqueJours mal configure ?"
            })
        }
    }
    # v2.1 : nouveaux arrays a verifier (Events / ResourceWarnings)
    if ($Data.PSObject.Properties['Events'] -and $Data.Events) {
        $cnt = @($Data.Events).Count
        if ($cnt -gt $AnomalyThresholds.MaxEventsCount) {
            $anomalies.Add([pscustomobject]@{
                Type   = 'OversizedArray'
                Reason = "Events anormalement long"
                Detail = "$cnt entrees (seuil: $($AnomalyThresholds.MaxEventsCount)) - PC en reboot loop ou JSON forge ?"
            })
        }
    }
    if ($Data.PSObject.Properties['ResourceWarnings'] -and $Data.ResourceWarnings) {
        $cnt = @($Data.ResourceWarnings).Count
        if ($cnt -gt $AnomalyThresholds.MaxResourceWarningsCount) {
            $anomalies.Add([pscustomobject]@{
                Type   = 'OversizedArray'
                Reason = "ResourceWarnings anormalement long"
                Detail = "$cnt entrees (seuil: $($AnomalyThresholds.MaxResourceWarningsCount)) - PC en RAM exhaustion permanente ou JSON forge ?"
            })
        }
    }

    # --------------------------------------------------------
    # ANOMALIE 5 : Valeurs aberrantes
    # --------------------------------------------------------
    # Battery health
    if ($Data.PSObject.Properties['BatteryInfo'] -and $Data.BatteryInfo) {
        $bat = $Data.BatteryInfo
        if ($bat.PSObject.Properties['HasBattery'] -and $bat.HasBattery -eq $true) {
            if ($bat.PSObject.Properties['HealthPercent']) {
                $h = $bat.HealthPercent
                if ($h -lt 0 -or $h -gt 100) {
                    $anomalies.Add([pscustomobject]@{
                        Type   = 'OutOfRangeValue'
                        Reason = "BatteryInfo.HealthPercent hors [0;100]"
                        Detail = "valeur=$h - bug WMI ou JSON forge ?"
                    })
                }
            }
        }
    }

    # Disk wear
    if ($Data.PSObject.Properties['DiskHealth'] -and $Data.DiskHealth) {
        foreach ($d in $Data.DiskHealth) {
            if ($d.PSObject.Properties['WearPct']) {
                $w = $d.WearPct
                if ($w -lt 0 -or $w -gt 100) {
                    $anomalies.Add([pscustomobject]@{
                        Type   = 'OutOfRangeValue'
                        Reason = "DiskHealth.WearPct hors [0;100]"
                        Detail = "disque='$(Get-SafeString $d.FriendlyName)' valeur=$w"
                    })
                }
            }
        }
    }

    # Uptime aberrant
    if ($Data.Machine.PSObject.Properties['UptimeDays']) {
        $u = $Data.Machine.UptimeDays
        if ($u -gt $AnomalyThresholds.MaxUptimeDays) {
            $anomalies.Add([pscustomobject]@{
                Type   = 'OutOfRangeValue'
                Reason = "Uptime anormalement long"
                Detail = "$([math]::Round($u,0)) jours (seuil: $($AnomalyThresholds.MaxUptimeDays)j) - PC jamais reboote ou bug ?"
            })
        }
    }

    # --------------------------------------------------------
    # ANOMALIE 6 : Arrays tronques au Collector (v2.1)
    # --------------------------------------------------------
    # Lecture du bloc Meta.TruncatedArrays produit par le Collector. Si au
    # moins un array est flagge tronque, la volumetrie de ce PC a depasse les
    # limites du Collector -> le rapport est PARTIEL pour lui.
    # v2.1.10 : message rendu FACTUEL. Avant, on affirmait "PC en boucle
    # d'erreur (reboot loop, RAM exhaustion, spam Event 41)" pour TOUTE
    # troncature - faux dans la majorite des cas : TopCrashers deborde des
    # qu'un poste a plus de 10 apps distinctes en echec (bruit applicatif, ou
    # un agent qui crashe en boucle), sans aucun rapport avec un reboot loop.
    # Le vrai signal de sante est dans le CONTENU (count des crashers, onglet
    # Stabilite pour les BSOD/reboots), pas dans le flag de troncature. On se
    # contente donc de constater la troncature et de pointer ou regarder.
    #
    # Compatible v2.0 : si le bloc Meta n'existe pas (JSON antedeluvien
    # par exemple), on skip silencieusement. Pas d'anomalie generee.
    if ($Data.PSObject.Properties['Meta'] -and
        $Data.Meta -and
        $Data.Meta.PSObject.Properties['TruncatedArrays'] -and
        $Data.Meta.TruncatedArrays) {
        $truncatedNames = @()
        foreach ($prop in $Data.Meta.TruncatedArrays.PSObject.Properties) {
            if ($prop.Value -eq $true) {
                $truncatedNames += $prop.Name
            }
        }
        if ($truncatedNames.Count -gt 0) {
            # Nuance factuelle par array, sans presumer d'un diagnostic systeme.
            $hints = @()
            foreach ($tn in $truncatedNames) {
                switch ($tn) {
                    'TopCrashers'      { $hints += "beaucoup d'applications distinctes en echec (voir la liste des crashers)" }
                    'Events'           { $hints += "beaucoup d'evenements systeme (voir l'onglet Stabilite)" }
                    'ResourceWarnings' { $hints += "beaucoup de signaux ressource RAM/CPU/disque (voir l'onglet Materiel)" }
                    'BSODs'            { $hints += "beaucoup de BSOD (voir l'onglet Stabilite)" }
                    'BootDurations'    { $hints += "beaucoup de demarrages sur la periode" }
                }
            }
            $hintTxt = if ($hints.Count -gt 0) { " - " + ($hints -join " ; ") } else { "" }
            $anomalies.Add([pscustomobject]@{
                Type   = 'TruncatedAtCollector'
                Reason = "Array(s) tronque(s) a la collecte - rapport partiel pour ce PC"
                Detail = "Tronque(s) : $($truncatedNames -join ', ')$hintTxt. Volumetrie reelle superieure aux limites du Collector ; verifier le contenu correspondant."
            })
        }
    }

    return $anomalies.ToArray()
}

# ============================================================
# LECTURE DES JSON (avec verification SchemaVersion)
# ============================================================
Write-Host "[*] Lecture des donnees depuis $SharePath ..." -ForegroundColor Cyan

# NB : on exclut decommissioning.json (registre de cycle de vie, pas un poste) au
# cas ou il partage le dossier des JSON (demo, ou repli par defaut sur SharePath).
$jsonFiles      = Get-ChildItem -Path $SharePath -Filter "$FiltrePC.json" -ErrorAction Stop |
                  Where-Object { $_.Name -ne 'decommissioning.json' }
$allData        = [System.Collections.Generic.List[PSCustomObject]]::new()
$rejectedJsons  = [System.Collections.Generic.List[PSCustomObject]]::new()  # v2.0 : pour la section HTML "JSON suspects"
$anomalyJsons   = [System.Collections.Generic.List[PSCustomObject]]::new()  # v2.0+ : anomalies (warnings) detectees
$validFiles     = [System.Collections.Generic.List[PSCustomObject]]::new()  # pour anomalie #4 (cross-PC)

foreach ($file in $jsonFiles) {
    # v2.0 : validation stricte via Test-PCPulseJson (sanity-checks niveaux 1-4)
    $check = Test-PCPulseJson -FilePath $file.FullName

    if ($check.Valid) {
        # Ajouter aux donnees du rapport (deja sanitisees en place)
        $allData.Add($check.Data)
        $validFiles.Add([pscustomobject]@{ File = $file; Data = $check.Data; PC = $check.Data.Machine.PC })

        # v2.0+ : detection des anomalies "soft" (warnings, JSON garde dans le rapport)
        $anomalies = Test-PCPulseAnomalies -Data $check.Data -FilePath $file.FullName
        foreach ($a in $anomalies) {
            $anomalyJsons.Add([pscustomobject]@{
                File   = $file.Name
                PC     = $check.Data.Machine.PC
                Type   = $a.Type
                Reason = $a.Reason
                Detail = $a.Detail
            })
        }
    } else {
        # Rejet : on garde la trace pour la section HTML dediee + log console
        $rejectedJsons.Add($check)
    }
}

# ============================================================
# v2.5.4 : 2e SOURCE - conteneur Azure "depot" (postes 100% distants)
# ------------------------------------------------------------
# Lit les <PC>.json deposes dans le cloud par la Function, les valide via le MEME
# Test-PCPulseJson, et FUSIONNE avec le share : par machine, on garde le rapport
# le PLUS RECENT (Machine.CollectedAt). Un poste baladeur (SMB parfois, cloud
# parfois) affiche donc toujours son dernier etat, quelle que soit la source.
# Gate par config (CloudDepotBaseUrl + CloudDepotSas). Vides -> ignore (share seul).
# Lecture SEULE (SAS read+list), zero dependance (REST via Invoke-RestMethod).
# ============================================================
# Base URL : param prioritaire, sinon config (non secret). SAS : PARAMETRE UNIQUEMENT
# (secret -> jamais dans release\config.psd1 qui est lisible par les postes).
$cloudBase = if ($CloudDepotBaseUrl) { $CloudDepotBaseUrl } else { [string]$cfg.CloudDepotBaseUrl }
$cloudSas  = ([string]$CloudDepotSas).TrimStart('?')
if ($cloudBase -and $cloudSas) {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        # NB : la reponse "list blobs" d'Azure commence par un BOM UTF-8 -> Invoke-RestMethod
        # ne la parse PAS en XML (il rend une chaine) et .EnumerationResults serait $null.
        # On lit le brut, on retire le BOM, puis on cast en [xml].
        $listResp = Invoke-WebRequest -Uri ("{0}?restype=container&comp=list&{1}" -f $cloudBase, $cloudSas) -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
        $listTxt  = [string]$listResp.Content
        if ($listTxt.Length -gt 0 -and $listTxt[0] -eq [char]0xFEFF) { $listTxt = $listTxt.Substring(1) }
        $listXml  = [xml]$listTxt
        $cloudBlobs = @($listXml.EnumerationResults.Blobs.Blob | Where-Object { $_.Name -like '*.json' })
        Write-Host ("[Cloud] {0} JSON dans le conteneur depot" -f $cloudBlobs.Count) -ForegroundColor Cyan
        $cloudTmp = Join-Path $env:TEMP ("pcpulse-cloud-" + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $cloudTmp -Force
        $cloudMerged = 0; $cloudAdded = 0
        foreach ($b in $cloudBlobs) {
            $pcName = [IO.Path]::GetFileNameWithoutExtension([string]$b.Name)
            if ($FiltrePC -ne '*' -and $pcName -ne $FiltrePC) { continue }
            $tmpF = Join-Path $cloudTmp ([string]$b.Name)
            try { Invoke-RestMethod -Uri ("{0}/{1}?{2}" -f $cloudBase, $b.Name, $cloudSas) -OutFile $tmpF -TimeoutSec 30 -ErrorAction Stop }
            catch { Write-Host "[Cloud] telechargement KO $($b.Name) : $_" -ForegroundColor Yellow; continue }
            $cc = Test-PCPulseJson -FilePath $tmpF
            if (-not $cc.Valid) { $rejectedJsons.Add($cc); continue }
            $cloudPcName = [string]$cc.Data.Machine.PC
            $cloudTs     = $cc.Data.Machine.CollectedAt -as [datetime]
            # Machine deja presente cote share ?
            $idx = -1
            for ($k = 0; $k -lt $allData.Count; $k++) { if ([string]$allData[$k].Machine.PC -eq $cloudPcName) { $idx = $k; break } }
            if ($idx -lt 0) {
                $allData.Add($cc.Data); $cloudAdded++            # machine vue SEULEMENT dans le cloud
            } else {
                $shareTs = $allData[$idx].Machine.CollectedAt -as [datetime]
                if ($cloudTs -and (-not $shareTs -or $cloudTs -gt $shareTs)) { $allData[$idx] = $cc.Data; $cloudMerged++ }  # cloud plus recent -> il gagne
            }
        }
        Write-Host ("[Cloud] fusion : {0} ajoutes (cloud seul), {1} remplaces (cloud plus recent)" -f $cloudAdded, $cloudMerged) -ForegroundColor Cyan
        try { Remove-Item -LiteralPath $cloudTmp -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    } catch {
        Write-Host "[Cloud] source Azure ignoree (erreur lecture, on continue en share seul) : $_" -ForegroundColor Yellow
    }
}

# v2.1.9 : la detection "SimilarHostname" (ANOMALIE #4, similarite Levenshtein <= 1) a
# ete RETIREE. Inadaptee a une nomenclature de parc SEQUENTIELLE : deux machines qui se
# suivent (ex. PC-001 / PC-002) ont une distance de 1 -> un faux positif a
# chaque paire de voisins (276 signaux sur 96 PC), ce qui noyait les vraies anomalies
# (troncature Collector, timestamps futurs, valeurs aberrantes...). Le scenario vise
# (spoofing par hostname ressemblant) est marginal : un JSON malveillant usurpe le nom
# EXACT d'une machine, pas un nom "proche". Historique git si besoin de la reintroduire
# un jour avec un filtre homoglyphe (ignorer les ecarts purement numeriques).

# Logs console - rejets
if ($rejectedJsons.Count -gt 0) {
    Write-Host "[!] $($rejectedJsons.Count) JSON rejete(s) :" -ForegroundColor Red
    $rejectedJsons | Select-Object -First 5 | ForEach-Object {
        Write-Host "    - $($_.File) : $($_.Reason)" -ForegroundColor Red
        if ($_.Detail) {
            Write-Host "      $($_.Detail)" -ForegroundColor DarkGray
        }
    }
    if ($rejectedJsons.Count -gt 5) {
        Write-Host "    ... et $($rejectedJsons.Count - 5) autre(s) - voir section dediee dans le rapport" -ForegroundColor Red
    }
}

# Logs console - anomalies (warnings)
if ($anomalyJsons.Count -gt 0) {
    Write-Host "[!] $($anomalyJsons.Count) anomalie(s) detectee(s) :" -ForegroundColor Yellow
    $anomalyJsons | Select-Object -First 5 | ForEach-Object {
        Write-Host "    - $($_.PC) : $($_.Reason)" -ForegroundColor Yellow
        if ($_.Detail) {
            Write-Host "      $($_.Detail)" -ForegroundColor DarkGray
        }
    }
    if ($anomalyJsons.Count -gt 5) {
        Write-Host "    ... et $($anomalyJsons.Count - 5) autre(s) - voir section dediee dans le rapport" -ForegroundColor Yellow
    }
}

if ($allData.Count -eq 0) {
    Write-Host "[-] Aucun JSON valide trouve." -ForegroundColor Red
    if ($rejectedJsons.Count -gt 0) {
        Write-Host "    ($($rejectedJsons.Count) JSON ont ete rejetes - voir SECURITY.md)" -ForegroundColor Red
    }
    exit 1
}
Write-Host "[+] $($allData.Count) PC(s) charge(s)" -ForegroundColor Green

# ============================================================
# v2.4.8 : SIGNAL DE SCHEMA DEPRECIE (prerequis au resserrage de la whitelist)
# ============================================================
# $AcceptedSchemaVersions tolere encore le schema 2.1, tolerance ouverte le temps
# du rollout du rename EDR -- rollout termine depuis cinq redeploiements. La
# tolerance ne se retire pas a l'aveugle pour autant : s'il reste des JSON en 2.1,
# ce sont des postes dont le Collector n'a pas pris CINQ mises a jour d'affilee,
# c'est-a-dire un probleme d'auto-update a traiter AVANT de resserrer -- sinon on
# ne corrige rien, on rend juste ces postes invisibles au dashboard, ce qui est
# exactement le contraire du but.
# Ce bloc transforme donc ce prerequis en mesure : il nomme les postes concernes a
# chaque generation. Console uniquement, aucun impact sur le rendu.
# Quand ce warning ne sort plus sur une generation du parc complet :
# passer $AcceptedSchemaVersions a @('2.2') et vider $DeprecatedSchemaVersions.
$DeprecatedSchemaVersions = @('2.1')
$deprecatedPCs = @($allData | Where-Object { [string]$_.SchemaVersion -in $DeprecatedSchemaVersions } |
    ForEach-Object { "$($_.Machine.PC) (schema $($_.SchemaVersion))" } | Sort-Object)
if ($deprecatedPCs.Count -gt 0) {
    Write-Host "[!] SCHEMA DEPRECIE : $($deprecatedPCs.Count) poste(s) encore en schema $($DeprecatedSchemaVersions -join '/') alors que le parc devrait etre en 2.2 :" -ForegroundColor Yellow
    foreach ($p in $deprecatedPCs) { Write-Host "      - $p" -ForegroundColor Yellow }
    Write-Host "      -> Collector pas a jour sur ces postes : verifier leur auto-update (PCPulse-Updater) AVANT de resserrer la whitelist." -ForegroundColor Yellow
}

# ============================================================
# COVERAGE-CHECK du re-mapping Machine -> embed (anti-recidive #3)
#   Le Dashboard reconstruit un objet embed champ par champ (plus bas).
#   Un champ ajoute au Collector mais oublie dans l'embed est droppe EN
#   SILENCE (deja arrive : Type, BootType, les 4 champs crashers). Ce
#   garde-fou diffe les cles Machine reellement presentes dans les JSON
#   vs la liste des cles connues (recopiees OU volontairement ignorees),
#   et gueule en jaune sur tout champ non traite. Diagnostic console
#   uniquement (rien dans le HTML). A tenir a jour quand on ajoute une
#   cle a machineInfo cote Collector.
# ============================================================
$KnownMachineKeys = @(
    # --- recopiees dans l'embed ($embedData plus bas) ---
    'PC','IP','CurrentUser','LastLoggedUser','CollectorRunAs','LastBoot','UptimeDays','CollectedAt',
    'CPUName','CPUVendor','CPUGen','CPUYear','CPUAge','CPUAgeCategory',
    'ConnectionType','ChassisInfo','SerialNumber','Manufacturer','Model',
    'OSProduct','OSBuild','OSDisplayVersion','OSEdition',
    # --- presentes dans Machine mais volontairement NON embarquees ---
    'OS','LastRealColdBoot','FastStartupEnabled'
)
# v2.3.1 (#3) : coverage-check GENERALISE - meme garde-fou etendu aux SOUS-OBJETS
# de l'embed. Un champ additif d'un sous-objet (crasher, WHEA, moniteur, disque
# SMART, barrette RAM) oublie au re-mapping serait droppe en silence, exactement
# comme au niveau Machine. Diagnostic console uniquement (rien dans le HTML).
function Test-EmbedCoverage {
    param([string]$Label, $Items, [string[]]$KnownKeys)
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($it in @($Items)) {
        if ($null -eq $it) { continue }
        foreach ($prop in $it.PSObject.Properties) { [void]$seen.Add($prop.Name) }
    }
    $unmapped = @($seen | Where-Object { $_ -notin $KnownKeys } | Sort-Object)
    if ($unmapped.Count -gt 0) {
        Write-Host "[!] COVERAGE-CHECK ($Label) : cle(s) presente(s) dans les JSON mais NON traitee(s) dans l'embed :" -ForegroundColor Yellow
        foreach ($k in $unmapped) {
            Write-Host "      - $k  (a recopier dans l'embed, ou a declarer dans la liste connue si ignore volontairement)" -ForegroundColor Yellow
        }
    }
}

# Niveau Machine (comportement inchange)
Test-EmbedCoverage -Label 'Machine' -Items (@($allData | ForEach-Object { $_.Machine })) -KnownKeys $KnownMachineKeys

# Sous-objets de l'embed (v2.3.1) - listes de cles connues a tenir a jour avec le re-mapping plus bas
Test-EmbedCoverage -Label 'TopCrashers' -Items (@($allData | ForEach-Object { ConvertTo-Array $_.TopCrashers })) -KnownKeys @(
    'AppName','CrashCount','Type','HangCount','ErrorCount','FaultModule','ExceptionCode')
Test-EmbedCoverage -Label 'Monitors' -Items (@($allData | ForEach-Object { ConvertTo-Array $_.Monitors })) -KnownKeys @(
    'ManufacturerCode','Manufacturer','Model','SerialNumber','ProductCode','YearOfManufacture',
    'WeekOfManufacture','AgeYears','Active','VideoOutputTech','Identified')
Test-EmbedCoverage -Label 'DiskHealth' -Items (@($allData | ForEach-Object { ConvertTo-Array $_.DiskHealth })) -KnownKeys @(
    'FriendlyName','MediaType','BusType','SizeGB','OperationalStatus','HealthStatus','TemperatureC',
    'TemperatureMaxC','WearPct','PowerOnHours','ReadErrorsTotal','ReadErrorsUncorrected',
    'WriteErrorsTotal','WriteErrorsUncorrected','IsAlert','AlertReasons')
Test-EmbedCoverage -Label 'MemoryInventory.Modules' -Items (@($allData | ForEach-Object { if ($_.MemoryInventory) { ConvertTo-Array $_.MemoryInventory.Modules } })) -KnownKeys @(
    'Slot','Bank','CapacityGB','Type','SpeedMHz','Manufacturer','ManufacturerRaw','PartNumber')
Test-EmbedCoverage -Label 'WHEA_Fatal' -Items (@($allData | ForEach-Object { if ($_.HardwareHealth) { ConvertTo-Array $_.HardwareHealth.WHEA_Fatal } })) -KnownKeys @(
    'Timestamp','EventId','Severity','Component','ErrorSource','BDF','Detail')
Test-EmbedCoverage -Label 'WHEA_Corrected' -Items (@($allData | ForEach-Object { if ($_.HardwareHealth) { ConvertTo-Array $_.HardwareHealth.WHEA_Corrected } })) -KnownKeys @(
    'Component','EventId','ErrorSource','BDF','Count','FirstSeen','LastSeen','Detail')
# v2.4.3 : DiskInfo et BootDurations n'etaient PAS couverts -> c'est l'angle mort
# qui avait laisse passer le XSS DiskInfo (champ non echappe droppe en silence).
Test-EmbedCoverage -Label 'DiskInfo' -Items (@($allData | ForEach-Object { ConvertTo-Array $_.DiskInfo })) -KnownKeys @(
    'Drive','Label','TotalGB','UsedGB','FreeGB','PctUsed','PctFree','IsAlert')
Test-EmbedCoverage -Label 'BootDurations' -Items (@($allData | ForEach-Object { ConvertTo-Array $_.BootDurations })) -KnownKeys @(
    'DateBoot','DurationMin','EstBootLong','PrecedentType','Method','BootType')

# ============================================================
# v2.4.8 : COVERAGE-CHECK COMPLETE - NIVEAU PAYLOAD + SOUS-OBJETS RESTANTS
# ============================================================
# Il manquait la marche la plus haute : AUCUN controle au niveau TOP-LEVEL du
# payload. Un bloc ENTIER ajoute au Collector (comme MemoryInventory ou
# BatteryInfo en leur temps) pouvait donc disparaitre sans un mot -- le pire cas
# de la classe de bug #3, puisque c'est aussi le plus facile a oublier : on pense
# aux champs d'un objet existant, pas a brancher un objet neuf.
#
# Les sous-objets restants sont ajoutes dans le meme mouvement. La revue a montre
# que TOUS les blocs de l'embed sont remappes champ par champ (aucun ne passe en
# bloc) : la classe de bug s'applique donc PARTOUT, pas seulement aux quelques
# sous-objets deja gardes depuis la 2.3.1. Couverture desormais complete.
#
# Rappel du contrat de $KnownKeys : la liste = cles RECOPIEES + cles IGNOREES
# VOLONTAIREMENT. Une cle absente de la liste declenche un warning. C'est
# volontairement une liste a tenir a jour a la main : le cout est une ligne par
# champ ajoute, le gain est qu'aucun champ ne peut se perdre en silence.
$KnownPayloadKeys = @(
    'SchemaVersion','Machine','Meta','Events','BootDurations','BSODs','ResourceWarnings',
    'TopRAM','DiskInfo','BatteryInfo','ServicesHealth','BootPerformance','DiskHealth',
    'Monitors','MemoryInventory','GPUInventory','TopCrashers','HardwareHealth','Stats',
    # v2.5.1 : client VPN (present + version). $null sur les JSON de Collector < 2.5.1.
    'VpnClient'
)
Test-EmbedCoverage -Label 'PAYLOAD (top-level)' -Items (@($allData)) -KnownKeys $KnownPayloadKeys

# v2.5.1 : VpnClient (Collector 2.5.1). Absent des JSON plus anciens (guard $_.VpnClient).
Test-EmbedCoverage -Label 'VpnClient' -Items (@($allData | ForEach-Object { $_.VpnClient })) -KnownKeys @(
    'Present','Product','Version')

# Meta : lu par la detection d'anomalies (Meta.TruncatedArrays), pas par l'embed.
Test-EmbedCoverage -Label 'Meta' -Items (@($allData | ForEach-Object { $_.Meta })) -KnownKeys @(
    'TruncatedArrays')

Test-EmbedCoverage -Label 'Events' -Items (@($allData | ForEach-Object { ConvertTo-Array $_.Events })) -KnownKeys @(
    # EventId n'est pas recopie dans l'embed : il sert de FILTRE (Id=41) a la
    # construction de $crashList, l'info est donc implicite cote JS.
    'Timestamp','EventId','Type','Detail','CrashCause','Message')

Test-EmbedCoverage -Label 'BSODs' -Items (@($allData | ForEach-Object { ConvertTo-Array $_.BSODs })) -KnownKeys @(
    'Date','Nom',
    # 'Taille' : EMIS par le Collector (taille du fichier .dmp) et volontairement
    # NON embarque -- le drill-down BSOD n'affiche que date + nom de dump. Declare
    # ici plutot que retire du Collector : retirer un champ du payload coute un
    # redeploiement de parc, pour un champ purement cosmetique. Si un jour on veut
    # l'afficher, l'info est deja dans les JSON.
    'Taille')

Test-EmbedCoverage -Label 'ResourceWarnings' -Items (@($allData | ForEach-Object { ConvertTo-Array $_.ResourceWarnings })) -KnownKeys @(
    # Count/FirstSeen/LastSeen/IsBurst : presents seulement sur les entrees
    # "disque lent" (clustering Event 51), absents des entrees RAM -> normal.
    'Timestamp','Type','Detail','Count','FirstSeen','LastSeen','IsBurst')

Test-EmbedCoverage -Label 'TopRAM' -Items (@($allData | ForEach-Object { ConvertTo-Array $_.TopRAM })) -KnownKeys @(
    'Name','WorkingSetMB','CPUSeconds')

Test-EmbedCoverage -Label 'GPUInventory' -Items (@($allData | ForEach-Object { ConvertTo-Array $_.GPUInventory })) -KnownKeys @(
    'Name','DriverVersion','DriverDate')

Test-EmbedCoverage -Label 'BatteryInfo' -Items (@($allData | ForEach-Object { $_.BatteryInfo })) -KnownKeys @(
    'HasBattery','Manufacturer','Chemistry','DesignCapacity','FullChargeCapacity',
    'HealthPercent','HealthCategory','CycleCount','CurrentChargePct','Status','IsAlert')

Test-EmbedCoverage -Label 'ServicesHealth' -Items (@($allData | ForEach-Object { $_.ServicesHealth })) -KnownKeys @(
    'Monitored')
Test-EmbedCoverage -Label 'ServicesHealth.Monitored' -Items (@($allData | ForEach-Object { if ($_.ServicesHealth) { ConvertTo-Array $_.ServicesHealth.Monitored } })) -KnownKeys @(
    'Id','DisplayName','ServiceName','Role','Installed','Status','StartType','IsAlert')

Test-EmbedCoverage -Label 'BootPerformance' -Items (@($allData | ForEach-Object { $_.BootPerformance })) -KnownKeys @(
    'LastBoot','History','Stats','IsAlert')
# LastBoot et History partagent la MEME forme (LastBoot = History[0]) : les deux
# listes de cles doivent rester identiques, d'ou la variable partagee.
$KnownBootPerfEventKeys = @(
    'Timestamp','Level','BootTimeMs','MainPathBootTimeMs','BootPostBootTimeMs',
    'UserProfileProcessingTimeMs','ExplorerInitTimeMs','NumStartupApps',
    'IsRebootAfterInstall','IsSlow',
    # 'BootStartTime' : EMIS par le Collector (ConvertFrom-BootPerfEvent) et NON
    # embarque -- deuxieme instance vivante de la classe de bug, trouvee en
    # ajoutant ce garde-fou. Ignore volontairement : le champ est redondant avec
    # Timestamp (heure de l'Event 100) pour tout ce que le drill-down affiche.
    'BootStartTime'
)
Test-EmbedCoverage -Label 'BootPerformance.LastBoot' -Items (@($allData | ForEach-Object { if ($_.BootPerformance) { $_.BootPerformance.LastBoot } })) -KnownKeys $KnownBootPerfEventKeys
Test-EmbedCoverage -Label 'BootPerformance.History' -Items (@($allData | ForEach-Object { if ($_.BootPerformance) { ConvertTo-Array $_.BootPerformance.History } })) -KnownKeys $KnownBootPerfEventKeys
Test-EmbedCoverage -Label 'BootPerformance.Stats' -Items (@($allData | ForEach-Object { if ($_.BootPerformance) { $_.BootPerformance.Stats } })) -KnownKeys @(
    'BootsAnalyzed','AvgBootTimeMs','AvgMainPathMs','AvgPostBootMs','MaxBootTimeMs','SlowBootsCount')

Test-EmbedCoverage -Label 'MemoryInventory' -Items (@($allData | ForEach-Object { $_.MemoryInventory })) -KnownKeys @(
    'TotalInstalledGB','MaxCapacityGB','TotalSlots','OccupiedSlots','FreeSlots','CanUpgrade','Modules')

Test-EmbedCoverage -Label 'HardwareHealth' -Items (@($allData | ForEach-Object { $_.HardwareHealth })) -KnownKeys @(
    'WHEA_Fatal','WHEA_Corrected','GPU_TDR','Thermal','CPUThrottling',
    # Schema v5.0 legacy : converti en WHEA_Fatal par le bloc de compatibilite
    # plus bas. Ne devrait plus apparaitre (whitelist SchemaVersion), declare pour
    # ne pas polluer la console si un JSON antique traine encore sur le share.
    'WHEA_CPU','WHEA_RAM','WHEA_PCIe')
Test-EmbedCoverage -Label 'HardwareHealth.GPU_TDR' -Items (@($allData | ForEach-Object { if ($_.HardwareHealth) { ConvertTo-Array $_.HardwareHealth.GPU_TDR } })) -KnownKeys @(
    'Timestamp','Driver','Detail')
Test-EmbedCoverage -Label 'HardwareHealth.Thermal' -Items (@($allData | ForEach-Object { if ($_.HardwareHealth) { ConvertTo-Array $_.HardwareHealth.Thermal } })) -KnownKeys @(
    'Timestamp','AlertType','Temperature','Zone','Detail')
Test-EmbedCoverage -Label 'HardwareHealth.CPUThrottling' -Items (@($allData | ForEach-Object { if ($_.HardwareHealth) { ConvertTo-Array $_.HardwareHealth.CPUThrottling } })) -KnownKeys @(
    'Day','EventId','Type','Count','TotalSeconds','FirstSeen','LastSeen','Detail')

# Stats : cas particulier, a lire avant de "corriger" un warning ici.
# L'embed n'expose PAS d'objet Stats ; il en aplatit une POIGNEE de champs
# (BootsByType, TotalWHEAFatal, TotalWHEACorr, TotalGPU, TotalThermal,
# TotalHardware, TotalHardCrash) et recalcule le reste cote JS a partir des
# tableaux embarques. Tous les autres champs sont donc IGNORES VOLONTAIREMENT --
# mais voir la note de volumetrie ci-dessous, ce choix n'est pas neutre.
Test-EmbedCoverage -Label 'Stats' -Items (@($allData | ForEach-Object { $_.Stats })) -KnownKeys @(
    # --- aplatis dans l'embed ---
    'BootsByType','TotalWHEAFatal','TotalWHEACorrected','TotalGPU','TotalThermal',
    'TotalHardware','TotalHardCrash',
    # --- ignores : recalcules cote JS depuis les tableaux embarques ---
    # ATTENTION (a arbitrer) : le Collector calcule ces compteurs AVANT le cap des
    # tableaux (Invoke-ArrayCap), precisement pour rester fideles a la volumetrie
    # reelle du poste. Les recalculer cote JS depuis des tableaux potentiellement
    # TRONQUES les fait SOUS-ESTIMER sur les postes bavards -- exactement les
    # postes qui interessent. Meta.TruncatedArrays dit quels tableaux sont
    # concernes. Pas corrige ici : cela changerait des KPI affiches, ce qui se
    # decide en regardant le dashboard, pas en lisant le code.
    'TotalBoots','TotalCrashFreeze','TotalBSOD','TotalRealBoots','BootsLongs',
    'ResourceWarnings','DiskAlerts',
    'TopCrasherApp','TopCrasherCount','TopAppFailureApp','TopAppFailureCount',
    'WHEACorrectedUnique','TotalCPUThrottling',
    'DerniereActivite','Event12Trouve','Event27Trouve',
    # --- ignores : redondants avec un bloc deja embarque ---
    'BatteryHealthPct','BatteryAlert',          # -> BatteryInfo
    'LastBootTimeMs','LastPostBootTimeMs','BootPerfAlert',  # -> BootPerformance
    'DiskWorstWear','DiskHealthAlert',          # -> DiskHealth
    'MonitorsCount'                             # -> Monitors.Count
)

# ============================================================
# CONSTRUCTION DU PAYLOAD JS
# ============================================================
$now       = Get-Date
$embedData = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($pc in $allData) {
    $site = Get-SiteFromIP -IP $pc.Machine.IP -Ranges $rangesIP

    # v2.4.5 (#4) : cast protege. Un JSON avec un CollectedAt non parseable faisait
    # planter TOUTE la generation (un seul poste pourri = dashboard mort). En cas
    # d'echec : fallback tres ancien -> le poste s'affiche HORS LIGNE (defaut sur).
    try {
        $collectedAt = [datetime]$pc.Machine.CollectedAt
    } catch {
        Write-Host "[!] CollectedAt illisible pour $($pc.Machine.PC) ('$($pc.Machine.CollectedAt)') -> poste marque hors ligne" -ForegroundColor Yellow
        $collectedAt = $now.AddYears(-1)
    }
    $hoursAgo    = [math]::Round(($now - $collectedAt).TotalHours, 1)
    $isOffline   = ($hoursAgo -gt (24 * $SeuilOfflineJours))

    $crashList = @($pc.Events | Where-Object { $_.EventId -eq 41 } |
        Sort-Object Timestamp -Descending | ForEach-Object {
            # v1.6 : propager CrashCause (null sur JSON v1.4/1.5, rempli sur v1.6).
            # Le JS gere les deux cas avec un fallback sur Type/Detail.
            [PSCustomObject]@{
                Timestamp  = ConvertTo-HtmlSafe $_.Timestamp
                Type       = if ($_.Type) { ConvertTo-HtmlSafe $_.Type } else { 'Freeze' }
                Detail     = ConvertTo-HtmlSafe $_.Detail
                CrashCause = ConvertTo-HtmlSafe $_.CrashCause
                Message    = ConvertTo-HtmlSafe $_.Message
            }
        })

    # v2.5.0 : redemarrages planifies (Event 1074). Sert a distinguer, dans
    # l'onglet Demarrage, un cycle de MAJ Windows (initie par TrustedInstaller /
    # servicing / Windows Update) d'un vrai cold boot utilisateur. On ne garde que
    # l'horodatage + un flag IsUpdate (aucune donnee sensible, Message non embarque).
    $rebootList = @($pc.Events | Where-Object { $_.EventId -eq 1074 } |
        Sort-Object Timestamp -Descending | ForEach-Object {
            [PSCustomObject]@{
                Timestamp = ConvertTo-HtmlSafe $_.Timestamp
                IsUpdate  = [bool]([string]$_.Message -match 'TrustedInstaller|servicing|Windows Update|wuauclt|MoUso|mise . jour')
            }
        })

    # v2.4.6 (#4) : tri TOLERANT. Avant, Sort-Object { [datetime]$_.DateBoot } jetait
    # sur un DateBoot non parseable -> un seul JSON pourri plantait TOUTE la generation
    # (meme classe de bug que le CollectedAt corrige plus haut). TryParse : les dates
    # illisibles retombent en MinValue (triees en premier), la generation continue.
    $bootList = @($pc.BootDurations | Sort-Object {
            $d = [datetime]::MinValue
            [void][datetime]::TryParse([string]$_.DateBoot, [ref]$d)
            $d
        } | ForEach-Object {
        [PSCustomObject]@{
            DateBoot      = ConvertTo-HtmlSafe $_.DateBoot
            DurationMin   = ConvertTo-NumOrNull $_.DurationMin   # v2.4.3 : cast (rendu en texte -> anti-XSS)
            EstBootLong   = [bool]$_.EstBootLong                 # v2.4.3 : flag booleen strict
            PrecedentType = ConvertTo-HtmlSafe $_.PrecedentType
            Method        = if ($_.Method) { ConvertTo-HtmlSafe $_.Method } else { 'unknown' }
            # v6.0 : fix BootType manquant dans l'embed (le Collector l'ecrivait
            # bien dans le JSON mais le Dashboard ne le recopiait pas dans le
            # payload JS, d'ou le panneau "Repartition demarrages" qui
            # affichait tout en "Inconnu").
            BootType      = if ($_.BootType) { ConvertTo-HtmlSafe $_.BootType } else { 'Unknown' }
        }
    })

    $bsodList = @(ConvertTo-Array $pc.BSODs | Where-Object { $_.Date } | ForEach-Object {
        [PSCustomObject]@{ Date = ConvertTo-HtmlSafe $_.Date; Nom = ConvertTo-HtmlSafe $_.Nom }
    })

    $warningList = @(ConvertTo-Array $pc.ResourceWarnings | Where-Object { $_.Timestamp } | ForEach-Object {
        # v1.5 : propager les champs de clustering Event 51 (Count/IsBurst/FirstSeen/LastSeen).
        # Les JSON v1.4 n'ont pas ces champs : $_.Count retournera $null et sera serialise
        # en null dans le JSON, ce que le JS gere avec 'typeof w.Count === "number"'.
        [PSCustomObject]@{
            Timestamp = ConvertTo-HtmlSafe $_.Timestamp
            Type      = ConvertTo-HtmlSafe $_.Type
            Detail    = ConvertTo-HtmlSafe $_.Detail
            Count     = ConvertTo-IntOrNull $_.Count
            IsBurst   = $_.IsBurst
            FirstSeen = ConvertTo-HtmlSafe $_.FirstSeen
            LastSeen  = ConvertTo-HtmlSafe $_.LastSeen
        }
    })

    $topRAMList = @(ConvertTo-Array $pc.TopRAM | Where-Object { $_.Name } | ForEach-Object {
        [PSCustomObject]@{
            Name         = ConvertTo-HtmlSafe $_.Name
            WorkingSetMB = ConvertTo-NumOrNull $_.WorkingSetMB
            CPUSeconds   = $_.CPUSeconds
        }
    })

    $diskList = @(ConvertTo-Array $pc.DiskInfo | Where-Object { $_.Drive } | ForEach-Object {
        # v2.4.3 : XSS pass 3. Drive et les champs numeriques partaient BRUTS puis
        # atterrissaient en innerHTML -> un poste compromis (JSON force) pouvait
        # injecter du JS chez l'admin. Drive echappe, numeriques cast (ou null).
        [PSCustomObject]@{
            Drive   = ConvertTo-HtmlSafe ([string]$_.Drive)
            Label   = if ($_.Label) { ConvertTo-HtmlSafe $_.Label } else { 'Sans nom' }
            TotalGB = ConvertTo-NumOrNull $_.TotalGB
            UsedGB  = ConvertTo-NumOrNull $_.UsedGB
            FreeGB  = ConvertTo-NumOrNull $_.FreeGB
            PctUsed = ConvertTo-NumOrNull $_.PctUsed
            PctFree = ConvertTo-NumOrNull $_.PctFree
            IsAlert = [bool]$_.IsAlert
        }
    })

    $crasherList = @(ConvertTo-Array $pc.TopCrashers | Where-Object { $_.AppName } | ForEach-Object {
        # v2.1.5 : le champ Type etait perdu ici (recopie AppName/CrashCount seulement),
        # ce qui laissait la section "Applis en echec recurrent" vide et desactivait le
        # badge du drill-down (c.Type / tc.Type toujours undefined cote JS -> tout en crash).
        # On le restaure NORMALISE en whitelist : seule la valeur exacte 'app_failure' passe,
        # tout le reste (y compris une valeur forgee par un endpoint compromis) tombe sur
        # 'crash', le comportement legacy sur. Coherent avec le modele de menace.
        $crasherType = if ("$($_.Type)" -eq 'app_failure') { 'app_failure' } else { 'crash' }
        [PSCustomObject]@{
            AppName       = ConvertTo-HtmlSafe $_.AppName
            CrashCount    = ConvertTo-IntOrNull $_.CrashCount
            Type          = $crasherType
            # v2.1.13 : champs additifs du Collector 2.1.5 (fige/plante + origine), recopies pour
            # (a) la piste materielle memoire ci-dessous, (b) le futur affichage detaille du drill-down.
            # Retrocompat : absents d'un JSON < 2.1.5 -> compteurs a 0, module/code a null.
            HangCount     = [int]$_.HangCount
            ErrorCount    = [int]$_.ErrorCount
            FaultModule   = if ($_.FaultModule)   { ConvertTo-HtmlSafe ([string]$_.FaultModule) }   else { $null }
            ExceptionCode = if ($_.ExceptionCode) { ConvertTo-HtmlSafe ([string]$_.ExceptionCode) } else { $null }
        }
    })

    # ================================================================
    # HARDWARE HEALTH : deux schemas supportes
    #   v5.2 : WHEA_Fatal (par occurrence) + WHEA_Corrected (agrege par signature)
    #   v5.0 : WHEA_CPU / WHEA_RAM / WHEA_PCIe (par ID d'event, imprecise)
    # On normalise tout vers le schema v5.2 en memoire.
    # ================================================================
    $hwHealth = [PSCustomObject]@{
        WHEA_Fatal     = @()
        WHEA_Corrected = @()
        GPU_TDR        = @()
        Thermal        = @()
        # v1.8 : CPU throttling detaille (agregation quotidienne Event 35/55).
        # Reste vide si Collector < v1.8 : pas de regression, le JS gere.
        CPUThrottling  = @()
    }

    if ($pc.HardwareHealth) {
        $hh = $pc.HardwareHealth

        # --- Detection du schema ---
        $hasV52 = ($null -ne $hh.PSObject.Properties['WHEA_Fatal']) -or ($null -ne $hh.PSObject.Properties['WHEA_Corrected'])
        $hasV50 = ($null -ne $hh.PSObject.Properties['WHEA_CPU'])   -or ($null -ne $hh.PSObject.Properties['WHEA_RAM']) -or ($null -ne $hh.PSObject.Properties['WHEA_PCIe'])

        if ($hasV52) {
            # Schema v5.2 : lecture directe
            $hwHealth.WHEA_Fatal = @(ConvertTo-Array $hh.WHEA_Fatal | Where-Object { $_.Timestamp } | ForEach-Object {
                [PSCustomObject]@{
                    Timestamp   = ConvertTo-HtmlSafe $_.Timestamp
                    EventId     = ConvertTo-IntOrNull $_.EventId
                    Severity    = ConvertTo-IntOrNull $_.Severity
                    Component   = ConvertTo-HtmlSafe $_.Component
                    ErrorSource = ConvertTo-HtmlSafe $_.ErrorSource
                    BDF         = ConvertTo-HtmlSafe $_.BDF
                    Detail      = ConvertTo-HtmlSafe $_.Detail
                }
            })
            $hwHealth.WHEA_Corrected = @(ConvertTo-Array $hh.WHEA_Corrected | Where-Object { $_.LastSeen } | ForEach-Object {
                [PSCustomObject]@{
                    Component   = ConvertTo-HtmlSafe $_.Component
                    EventId     = ConvertTo-IntOrNull $_.EventId
                    ErrorSource = ConvertTo-HtmlSafe $_.ErrorSource
                    BDF         = ConvertTo-HtmlSafe $_.BDF
                    Count       = ConvertTo-IntOrNull $_.Count
                    FirstSeen   = ConvertTo-HtmlSafe $_.FirstSeen
                    LastSeen    = ConvertTo-HtmlSafe $_.LastSeen
                    Detail      = ConvertTo-HtmlSafe $_.Detail
                }
            })
        }
        elseif ($hasV50) {
            # Schema v5.0 legacy : on convertit les trois listes en WHEA_Fatal
            # (on est conservateur : tout etait mis en "erreur" donc on considere Fatal)
            $legacyFatal = @()
            foreach ($k in @('WHEA_CPU', 'WHEA_RAM', 'WHEA_PCIe')) {
                $comp = switch ($k) { 'WHEA_CPU' {'CPU'}; 'WHEA_RAM' {'RAM'}; 'WHEA_PCIe' {'PCIe'} }
                $legacyFatal += @(ConvertTo-Array $hh.$k | Where-Object { $_.Timestamp } | ForEach-Object {
                    [PSCustomObject]@{
                        Timestamp   = ConvertTo-HtmlSafe $_.Timestamp
                        EventId     = if ($_.EventId) { $_.EventId } else { 0 }
                        Severity    = 2
                        Component   = $comp
                        ErrorSource = "(legacy v5.0)"
                        BDF         = ""
                        Detail      = ConvertTo-HtmlSafe $_.Detail
                    }
                })
            }
            $hwHealth.WHEA_Fatal = $legacyFatal
            # WHEA_Corrected reste vide : l'info n'existe pas en v5.0
        }

        $hwHealth.GPU_TDR = @(ConvertTo-Array $hh.GPU_TDR | Where-Object { $_.Timestamp } | ForEach-Object {
            [PSCustomObject]@{
                Timestamp = ConvertTo-HtmlSafe $_.Timestamp
                Driver    = ConvertTo-HtmlSafe $_.Driver
                Detail    = ConvertTo-HtmlSafe $_.Detail
            }
        })
        $hwHealth.Thermal = @(ConvertTo-Array $hh.Thermal | Where-Object { $_.Timestamp } | ForEach-Object {
            [PSCustomObject]@{
                Timestamp   = ConvertTo-HtmlSafe $_.Timestamp
                AlertType   = ConvertTo-HtmlSafe $_.AlertType
                Temperature = ConvertTo-HtmlSafe $_.Temperature
                Zone        = ConvertTo-HtmlSafe $_.Zone
                Detail      = ConvertTo-HtmlSafe $_.Detail
            }
        })
        # v1.8 : CPU throttling (present uniquement sur JSON >= 1.8)
        if ($hh.PSObject.Properties['CPUThrottling']) {
            $hwHealth.CPUThrottling = @(ConvertTo-Array $hh.CPUThrottling | Where-Object { $_.Day } | ForEach-Object {
                [PSCustomObject]@{
                    Day          = ConvertTo-HtmlSafe $_.Day
                    EventId      = ConvertTo-IntOrNull $_.EventId
                    Type         = ConvertTo-HtmlSafe $_.Type
                    Count        = ConvertTo-IntOrNull $_.Count
                    TotalSeconds = $_.TotalSeconds
                    FirstSeen    = ConvertTo-HtmlSafe $_.FirstSeen
                    LastSeen     = ConvertTo-HtmlSafe $_.LastSeen
                    Detail       = ConvertTo-HtmlSafe $_.Detail
                }
            })
        }
    }

    # Stats v5.2 si presentes, fallback vers stats v5.0
    $statsObj          = $pc.Stats
    $totalWHEAFatal    = if ($statsObj.TotalWHEAFatal -ne $null)    { $statsObj.TotalWHEAFatal }    else { @($hwHealth.WHEA_Fatal).Count }
    $totalWHEACorr     = if ($statsObj.TotalWHEACorrected -ne $null){ $statsObj.TotalWHEACorrected }else { 0 }
    $totalGPU          = if ($statsObj.TotalGPU)                    { $statsObj.TotalGPU }          else { @($hwHealth.GPU_TDR).Count }
    $totalThermal      = if ($statsObj.TotalThermal)                { $statsObj.TotalThermal }      else { @($hwHealth.Thermal).Count }
    $totalHWAlerting   = $totalWHEAFatal + $totalGPU + $totalThermal   # les "vraies" alertes (hors corrigees)

    # BootsByType : disponible en v5.2 uniquement
    $bootsByType = $null
    if ($statsObj.BootsByType) {
        $bootsByType = [PSCustomObject]@{
            ColdBoot    = if ($statsObj.BootsByType.ColdBoot)    { $statsObj.BootsByType.ColdBoot }    else { 0 }
            FastStartup = if ($statsObj.BootsByType.FastStartup) { $statsObj.BootsByType.FastStartup } else { 0 }
            Resume      = if ($statsObj.BootsByType.Resume)      { $statsObj.BootsByType.Resume }      else { 0 }
            Unknown     = if ($statsObj.BootsByType.Unknown)     { $statsObj.BootsByType.Unknown }     else { 0 }
        }
    }

    $cpuName = if ($pc.Machine.CPUName) { ([string]$pc.Machine.CPUName).Trim() } else { '' }

    # ================================================================
    # BATTERIE (v5.3+) : null si absent (JSON v5.2 ou desktop)
    # ================================================================
    $batteryEmbed = $null
    if ($pc.BatteryInfo) {
        $batteryEmbed = [PSCustomObject]@{
            HasBattery         = [bool]$pc.BatteryInfo.HasBattery
            Manufacturer       = ConvertTo-HtmlSafe ([string]$pc.BatteryInfo.Manufacturer)
            Chemistry          = ConvertTo-HtmlSafe ([string]$pc.BatteryInfo.Chemistry)
            DesignCapacity     = if ($null -ne $pc.BatteryInfo.DesignCapacity)     { [int64]$pc.BatteryInfo.DesignCapacity }     else { 0 }
            FullChargeCapacity = if ($null -ne $pc.BatteryInfo.FullChargeCapacity) { [int64]$pc.BatteryInfo.FullChargeCapacity } else { 0 }
            HealthPercent      = if ($null -ne $pc.BatteryInfo.HealthPercent)      { [double]$pc.BatteryInfo.HealthPercent }     else { 0 }
            HealthCategory     = ConvertTo-HtmlSafe ([string]$pc.BatteryInfo.HealthCategory)
            CycleCount         = ConvertTo-IntOrNull $pc.BatteryInfo.CycleCount   # peut etre $null
            CurrentChargePct   = if ($null -ne $pc.BatteryInfo.CurrentChargePct)   { [int]$pc.BatteryInfo.CurrentChargePct }     else { 0 }
            Status             = ConvertTo-HtmlSafe ([string]$pc.BatteryInfo.Status)
            IsAlert            = [bool]$pc.BatteryInfo.IsAlert
        }
    }

    # ================================================================
    # SERVICES HEALTH (v2.2 : liste Monitored pilotee par config)
    # On recopie chaque service surveille du JSON. HtmlSafe sur le texte, cast
    # strict sur les bool. Le JS repere Role='EDR' pour le score/badge ; les
    # autres services sont affiches dans le drill-down.
    # ================================================================
    $servicesEmbed = $null
    if ($pc.ServicesHealth -and $pc.ServicesHealth.Monitored) {
        $monitoredEmbed = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($svc in @($pc.ServicesHealth.Monitored)) {
            if (-not $svc) { continue }
            $monitoredEmbed.Add([PSCustomObject]@{
                Id          = ConvertTo-HtmlSafe ([string]$svc.Id)
                DisplayName = ConvertTo-HtmlSafe ([string]$svc.DisplayName)
                ServiceName = ConvertTo-HtmlSafe ([string]$svc.ServiceName)
                Role        = ConvertTo-HtmlSafe ([string]$svc.Role)
                Installed   = [bool]$svc.Installed
                Status      = ConvertTo-HtmlSafe ([string]$svc.Status)
                StartType   = ConvertTo-HtmlSafe ([string]$svc.StartType)
                IsAlert     = [bool]$svc.IsAlert
            })
        }
        $servicesEmbed = [PSCustomObject]@{
            Monitored = @($monitoredEmbed)
        }
    }

    # ================================================================
    # BOOT PERFORMANCE (v5.4+) : null si absent
    # ================================================================
    $bootPerfEmbed = $null
    if ($pc.BootPerformance) {
        $bp = $pc.BootPerformance
        $lastBootEmbed = $null
        if ($bp.LastBoot) {
            $lastBootEmbed = [PSCustomObject]@{
                Timestamp                   = ConvertTo-HtmlSafe $bp.LastBoot.Timestamp
                Level                       = ConvertTo-HtmlSafe ([string]$bp.LastBoot.Level)
                BootTimeMs                  = [int64]$bp.LastBoot.BootTimeMs
                MainPathBootTimeMs          = [int64]$bp.LastBoot.MainPathBootTimeMs
                BootPostBootTimeMs          = [int64]$bp.LastBoot.BootPostBootTimeMs
                UserProfileProcessingTimeMs = [int64]$bp.LastBoot.UserProfileProcessingTimeMs
                ExplorerInitTimeMs          = [int64]$bp.LastBoot.ExplorerInitTimeMs
                NumStartupApps              = [int]$bp.LastBoot.NumStartupApps
                IsRebootAfterInstall        = [bool]$bp.LastBoot.IsRebootAfterInstall
                IsSlow                      = [bool]$bp.LastBoot.IsSlow
            }
        }
        $historyEmbed = @()
        if ($bp.History) {
            $historyEmbed = @(ConvertTo-Array $bp.History | ForEach-Object {
                [PSCustomObject]@{
                    Timestamp                   = ConvertTo-HtmlSafe $_.Timestamp
                    Level                       = ConvertTo-HtmlSafe ([string]$_.Level)
                    BootTimeMs                  = [int64]$_.BootTimeMs
                    MainPathBootTimeMs          = [int64]$_.MainPathBootTimeMs
                    BootPostBootTimeMs          = [int64]$_.BootPostBootTimeMs
                    UserProfileProcessingTimeMs = [int64]$_.UserProfileProcessingTimeMs
                    ExplorerInitTimeMs          = [int64]$_.ExplorerInitTimeMs
                    NumStartupApps              = [int]$_.NumStartupApps
                    IsRebootAfterInstall        = [bool]$_.IsRebootAfterInstall
                    IsSlow                      = [bool]$_.IsSlow
                }
            })
        }
        $bootPerfEmbed = [PSCustomObject]@{
            LastBoot = $lastBootEmbed
            History  = $historyEmbed
            Stats    = if ($bp.Stats) {
                [PSCustomObject]@{
                    BootsAnalyzed  = [int]$bp.Stats.BootsAnalyzed
                    AvgBootTimeMs  = [int]$bp.Stats.AvgBootTimeMs
                    AvgMainPathMs  = [int]$bp.Stats.AvgMainPathMs
                    AvgPostBootMs  = [int]$bp.Stats.AvgPostBootMs
                    MaxBootTimeMs  = [int]$bp.Stats.MaxBootTimeMs
                    SlowBootsCount = [int]$bp.Stats.SlowBootsCount
                }
            } else { $null }
            IsAlert = [bool]$bp.IsAlert
        }
    }

    # ================================================================
    # DISK HEALTH SMART (v5.4+) : array, vide si absent
    # ================================================================
    $diskHealthEmbed = @()
    if ($pc.DiskHealth) {
        $diskHealthEmbed = @(ConvertTo-Array $pc.DiskHealth | ForEach-Object {
            [PSCustomObject]@{
                FriendlyName           = ConvertTo-HtmlSafe ([string]$_.FriendlyName)
                MediaType              = ConvertTo-HtmlSafe ([string]$_.MediaType)
                BusType                = ConvertTo-HtmlSafe ([string]$_.BusType)
                SizeGB                 = ConvertTo-NumOrNull $_.SizeGB
                OperationalStatus      = ConvertTo-HtmlSafe ([string]$_.OperationalStatus)
                HealthStatus           = ConvertTo-HtmlSafe ([string]$_.HealthStatus)
                TemperatureC           = ConvertTo-IntOrNull $_.TemperatureC
                TemperatureMaxC        = ConvertTo-IntOrNull $_.TemperatureMaxC
                WearPct                = ConvertTo-IntOrNull $_.WearPct
                PowerOnHours           = ConvertTo-IntOrNull $_.PowerOnHours
                ReadErrorsTotal        = ConvertTo-IntOrNull $_.ReadErrorsTotal
                ReadErrorsUncorrected  = ConvertTo-IntOrNull $_.ReadErrorsUncorrected
                WriteErrorsTotal       = ConvertTo-IntOrNull $_.WriteErrorsTotal
                WriteErrorsUncorrected = ConvertTo-IntOrNull $_.WriteErrorsUncorrected
                IsAlert                = [bool]$_.IsAlert
                AlertReasons           = @(ConvertTo-Array $_.AlertReasons | ForEach-Object { ConvertTo-HtmlSafe ([string]$_) })
            }
        })
    }

    # ================================================================
    # MONITORS (v5.5+) : array, vide si absent ou si que des ecrans internes
    # ================================================================
    $monitorsEmbed = @()
    if ($pc.Monitors) {
        $monitorsEmbed = @(ConvertTo-Array $pc.Monitors | ForEach-Object {
            [PSCustomObject]@{
                ManufacturerCode  = ConvertTo-HtmlSafe ([string]$_.ManufacturerCode)
                Manufacturer      = ConvertTo-HtmlSafe ([string]$_.Manufacturer)
                Model             = ConvertTo-HtmlSafe ([string]$_.Model)
                SerialNumber      = ConvertTo-HtmlSafe ([string]$_.SerialNumber)
                ProductCode       = ConvertTo-HtmlSafe ([string]$_.ProductCode)
                YearOfManufacture = ConvertTo-IntOrNull $_.YearOfManufacture
                WeekOfManufacture = ConvertTo-IntOrNull $_.WeekOfManufacture
                AgeYears          = ConvertTo-IntOrNull $_.AgeYears
                Active            = [bool]$_.Active
                VideoOutputTech   = ConvertTo-IntOrNull $_.VideoOutputTech
                # v5.8 : $true/$false si Collector >= 2.3.0, $null sinon (le JS deduit alors la signature EDID nul)
                Identified        = if ($_.PSObject.Properties['Identified']) { [bool]$_.Identified } else { $null }
            }
        })
    }

    # ================================================================
    # MEMORY INVENTORY (v1.8) : null si Collector < v1.8
    # ================================================================
    $memoryEmbed = $null
    if ($pc.MemoryInventory) {
        $memInv = $pc.MemoryInventory
        $memoryEmbed = [PSCustomObject]@{
            TotalInstalledGB = if ($null -ne $memInv.TotalInstalledGB) { [int]$memInv.TotalInstalledGB } else { 0 }
            MaxCapacityGB    = if ($null -ne $memInv.MaxCapacityGB)    { [int]$memInv.MaxCapacityGB }    else { 0 }
            TotalSlots       = if ($null -ne $memInv.TotalSlots)       { [int]$memInv.TotalSlots }       else { 0 }
            OccupiedSlots    = if ($null -ne $memInv.OccupiedSlots)    { [int]$memInv.OccupiedSlots }    else { 0 }
            FreeSlots        = if ($null -ne $memInv.FreeSlots)        { [int]$memInv.FreeSlots }        else { 0 }
            CanUpgrade       = [bool]$memInv.CanUpgrade
            Modules          = @(ConvertTo-Array $memInv.Modules | ForEach-Object {
                [PSCustomObject]@{
                    Slot         = ConvertTo-HtmlSafe ([string]$_.Slot)
                    Bank         = ConvertTo-HtmlSafe ([string]$_.Bank)
                    CapacityGB   = if ($null -ne $_.CapacityGB) { [int]$_.CapacityGB } else { 0 }
                    Type         = ConvertTo-HtmlSafe ([string]$_.Type)
                    SpeedMHz     = if ($null -ne $_.SpeedMHz) { [int]$_.SpeedMHz } else { 0 }
                    Manufacturer = ConvertTo-HtmlSafe ([string]$_.Manufacturer)
                    # v1.9 : ID JEDEC brut conserve pour l'audit (le JS decode aussi cote client pour les anciens JSON)
                    ManufacturerRaw = if ($_.PSObject.Properties['ManufacturerRaw']) { ConvertTo-HtmlSafe ([string]$_.ManufacturerRaw) } else { '' }
                    PartNumber   = ConvertTo-HtmlSafe ([string]$_.PartNumber)
                }
            })
        }
    }

    # ================================================================
    # GPU INVENTORY (v1.8) : tableau vide si Collector < v1.8
    # ================================================================
    $gpuEmbed = @()
    if ($pc.GPUInventory) {
        $gpuEmbed = @(ConvertTo-Array $pc.GPUInventory | ForEach-Object {
            [PSCustomObject]@{
                Name          = ConvertTo-HtmlSafe ([string]$_.Name)
                DriverVersion = ConvertTo-HtmlSafe ([string]$_.DriverVersion)
                DriverDate    = ConvertTo-HtmlSafe ([string]$_.DriverDate)
            }
        })
    }

    $embedData.Add([PSCustomObject]@{
        PC               = ConvertTo-HtmlSafe ([string]$pc.Machine.PC).Trim()
        IP               = ConvertTo-HtmlSafe $pc.Machine.IP
        Site             = ConvertTo-HtmlSafe $site
        CurrentUser      = ConvertTo-HtmlSafe $pc.Machine.CurrentUser
        # v2.3.0 : dernier user connu (repli d'affichage quand CurrentUser = "(aucune session)").
        LastLoggedUser   = if ($pc.Machine.PSObject.Properties['LastLoggedUser']) { ConvertTo-HtmlSafe ([string]$pc.Machine.LastLoggedUser) } else { '' }
        # v2.3.1 (#13) : compte d'execution reel du Collector (audit gMSA->SYSTEM). Exporte au CSV.
        CollectorRunAs   = if ($pc.Machine.PSObject.Properties['CollectorRunAs']) { ConvertTo-HtmlSafe ([string]$pc.Machine.CollectorRunAs) } else { '' }
        LastBoot         = ConvertTo-HtmlSafe $pc.Machine.LastBoot
        UptimeDays       = $pc.Machine.UptimeDays
        CollectedAt      = ConvertTo-HtmlSafe $pc.Machine.CollectedAt
        IsOffline        = $isOffline
        CPUName          = ConvertTo-HtmlSafe $cpuName
        CPUVendor        = ConvertTo-HtmlSafe $pc.Machine.CPUVendor
        CPUGen           = if ($null -ne $pc.Machine.CPUGen) { ConvertTo-HtmlSafe ([string]$pc.Machine.CPUGen) } else { $null }
        CPUYear          = ConvertTo-IntOrNull $pc.Machine.CPUYear
        CPUAge           = ConvertTo-IntOrNull $pc.Machine.CPUAge
        CPUAgeCategory   = if ($pc.Machine.CPUAgeCategory) { ConvertTo-HtmlSafe $pc.Machine.CPUAgeCategory } else { 'Inconnu' }
        ConnectionType   = if ($pc.Machine.ConnectionType) { ConvertTo-HtmlSafe $pc.Machine.ConnectionType } else { 'Inconnu' }
        # v2.2.1 : OS (Windows 10/11) - ADDITIF. Recopie EXPLICITE des 4 champs Machine :
        #          sans ces lignes ils seraient droppes en silence (classe de bug #3).
        OSProduct        = if ($pc.Machine.OSProduct) { ConvertTo-HtmlSafe ([string]$pc.Machine.OSProduct) } else { 'Inconnu' }
        OSBuild          = ConvertTo-IntOrNull $pc.Machine.OSBuild   # v2.4.3 : cast (affiche + attribut title -> anti-XSS)
        OSDisplayVersion = if ($pc.Machine.OSDisplayVersion) { ConvertTo-HtmlSafe ([string]$pc.Machine.OSDisplayVersion) } else { '' }
        OSEdition        = if ($pc.Machine.OSEdition) { ConvertTo-HtmlSafe ([string]$pc.Machine.OSEdition) } else { '' }
        # v2.4.2 : n° de serie machine (service tag) - inventaire / decommissionnement.
        #          Peut etre '' si Collector < 2.4.2 ou BIOS sans serial exploitable.
        SerialNumber     = if ($pc.Machine.PSObject.Properties['SerialNumber']) { ConvertTo-HtmlSafe ([string]$pc.Machine.SerialNumber) } else { '' }
        # v2.4.7 : modele + fabricant machine (recherche par modele). Anti-XSS via
        #          ConvertTo-HtmlSafe (donnee collectee = non fiable). '' si Collector < 2.4.7.
        Manufacturer     = if ($pc.Machine.PSObject.Properties['Manufacturer']) { ConvertTo-HtmlSafe ([string]$pc.Machine.Manufacturer) } else { '' }
        Model            = if ($pc.Machine.PSObject.Properties['Model']) { ConvertTo-HtmlSafe ([string]$pc.Machine.Model) } else { '' }
        # v5.6 : chassis info (peut etre null si Collector < v5.6)
        ChassisInfo      = if ($pc.Machine.ChassisInfo) {
            [PSCustomObject]@{
                ChassisType  = ConvertTo-IntOrNull $pc.Machine.ChassisInfo.ChassisType
                ChassisLabel = ConvertTo-HtmlSafe ([string]$pc.Machine.ChassisInfo.ChassisLabel)
                IsLaptop     = [bool]$pc.Machine.ChassisInfo.IsLaptop
                IsDesktop    = [bool]$pc.Machine.ChassisInfo.IsDesktop
                IsAIO        = [bool]$pc.Machine.ChassisInfo.IsAIO
            }
        } else { $null }
        Crashes          = $crashList
        Boots            = $bootList
        Reboots          = $rebootList
        BSODs            = $bsodList
        ResourceWarnings = $warningList
        TopRAM           = $topRAMList
        DiskInfo         = $diskList
        TopCrashers      = $crasherList
        HardwareHealth   = $hwHealth
        BootsByType      = $bootsByType
        TotalWHEAFatal   = $totalWHEAFatal
        TotalWHEACorr    = $totalWHEACorr
        TotalGPU         = $totalGPU
        TotalThermal     = $totalThermal
        TotalHardware    = $totalHWAlerting
        # v1.6 : total des hard crashs filtres (BSODSilent+SleepResumeFailed+PowerLoss).
        # null sur JSON v1.4/1.5, int sur v1.6. Le JS gere ce cas dans computeVerdict.
        TotalHardCrash   = if ($null -ne $pc.Stats.TotalHardCrash) { [int]$pc.Stats.TotalHardCrash } else { $null }
        # v5.3 / v5.4 additions (peuvent etre $null selon le schema du JSON source)
        BatteryInfo      = $batteryEmbed
        ServicesHealth   = $servicesEmbed
        # v2.5.1 : client VPN (present + version). Recopie EXPLICITE (classe de bug #3 :
        # sans ces lignes le champ serait droppe en silence). $null si Collector < 2.5.1.
        VpnClient        = if ($pc.VpnClient) {
            [PSCustomObject]@{
                Present = [bool]$pc.VpnClient.Present
                Product = ConvertTo-HtmlSafe ([string]$pc.VpnClient.Product)
                Version = ConvertTo-HtmlSafe ([string]$pc.VpnClient.Version)
            }
        } else { $null }
        BootPerformance  = $bootPerfEmbed
        DiskHealth       = $diskHealthEmbed
        # v5.5 additions
        Monitors         = $monitorsEmbed
        # v1.8 additions
        MemoryInventory  = $memoryEmbed
        GPUInventory     = $gpuEmbed
        # v2.1.11 : anomalies de ce PC (pour le badge + le filtre cote tableau).
        # $anomalyJsons est deja complet a ce stade (rempli a la lecture des fichiers).
        # ConvertTo-HtmlSafe car ces valeurs atteignent le DOM (badge/tooltip).
        Anomalies        = @($anomalyJsons | Where-Object { $_.PC -eq $pc.Machine.PC } | ForEach-Object {
            [PSCustomObject]@{
                Type   = ConvertTo-HtmlSafe ([string]$_.Type)
                Reason = ConvertTo-HtmlSafe ([string]$_.Reason)
                Detail = ConvertTo-HtmlSafe ([string]$_.Detail)
            }
        })
        # v2.4.0 : entree de decommission de ce PC (registre separe), ou $null.
        Decom            = $(
            $d = $decomByPc[[string]$pc.Machine.PC]
            if ($d) {
                [PSCustomObject]@{
                    Statut     = ConvertTo-HtmlSafe ([string]$d.Statut)
                    Reason     = ConvertTo-HtmlSafe ([string]$d.Reason)
                    AssignedTo = ConvertTo-HtmlSafe ([string]$d.AssignedTo)
                    MarkedAt   = ConvertTo-HtmlSafe ([string]$d.MarkedAt)
                    TargetDate = ConvertTo-HtmlSafe ([string]$d.TargetDate)
                    DoneAt     = ConvertTo-HtmlSafe ([string]$d.DoneAt)
                    DoneBy     = ConvertTo-HtmlSafe ([string]$d.DoneBy)
                }
            } else { $null }
        )
    })
}

# Force l'array meme avec 1 seul element (bug ConvertTo-Json connu)
$jsonEmbed    = ConvertTo-Json -InputObject @($embedData) -Depth 6 -EscapeHandling EscapeHtml
$weightsEmbed = ConvertTo-Json -InputObject $ScoreWeights -Compress

$titleHtml    = ConvertTo-HtmlSafe $DashboardTitle
$subtitleHtml = ConvertTo-HtmlSafe $DashboardSubtitle
# v2.5.0 : plus de meta-refresh navigateur. La page ne se recharge plus jamais
# d'elle-meme (l'ancien refresh 600s reinitialisait filtres / recherche / scroll /
# fiche ouverte en plein travail). La republication du HTML par le collecteur
# (~15 min) suffit : l'utilisateur recharge quand il le decide.
$metaRefresh  = ''
$maskHealthyJs = if ($MaskHealthyByDefault) { 'true' } else { 'false' }

# v2.5.3 : mode "Renouvellement". Si RenewalCandidateMaxYear est configure, le tag
# CPU Recent/Vieillissant/Ancien est remplace (tableau + filtre + KPI) par une
# categorie binaire Candidat/Parc courant. Sinon on garde l'historique (fallback).
$renewalMaxYearJs  = if ($null -ne $renewalMaxYear) { [string]$renewalMaxYear } else { 'null' }
$cpuFilterLabelTxt = if ($null -ne $renewalMaxYear) { 'Renouvellement :' } else { 'CPU :' }
$cpuFilterOptions  = if ($null -ne $renewalMaxYear) {
@'
        <option value="">Tous</option>
        <option value="renew">A renouveler</option>
        <option value="keep">Parc courant</option>
'@
} else {
@'
        <option value="">Tous</option>
        <option value="Recent">Recent</option>
        <option value="Vieillissant">Vieillissant</option>
        <option value="Ancien">Ancien</option>
        <option value="Inconnu">Inconnu</option>
'@
}

# v2.1.12 : liste des applis suivies serialisee en minuscules pour le rapprochement JS.
$priorityAppsJson = @($priorityApps | ForEach-Object { ([string]$_).ToLowerInvariant() }) | ConvertTo-Json -Compress
if (-not $priorityAppsJson) { $priorityAppsJson = '[]' }

# v2.4.2 : registre decommission complet (stats/audit). Force l'array (bug ConvertTo-Json 1 element).
$decomRegistryJson = ConvertTo-Json -InputObject @($decomAll) -Depth 4 -EscapeHandling EscapeHtml
if (-not $decomRegistryJson) { $decomRegistryJson = '[]' }
if ($decomRegistryJson.TrimStart()[0] -ne '[') { $decomRegistryJson = '[' + $decomRegistryJson + ']' }
$showSiteJs    = if ($showSite) { 'true' } else { 'false' }

# Favicon SVG "PCPulse" : ligne de pouls (ECG) sur fond violet, encode en data URI.
$faviconSvg = @'
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><rect width="64" height="64" rx="14" fill="#6c63ff"/><path d="M6 34 H20 L25 20 L33 44 L39 28 L43 34 H58" fill="none" stroke="#ffffff" stroke-width="4.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
'@
$faviconBytes = [System.Text.Encoding]::UTF8.GetBytes($faviconSvg)
$faviconB64   = [Convert]::ToBase64String($faviconBytes)

# ============================================================
# v2.0 : SECTION HTML "JSON suspects" (rejets sanity-checks)
# ============================================================
# Si des JSON ont ete rejetes par Test-PCPulseJson, on construit un panneau
# d'alerte qui sera affiche en tete du rapport. C'est un indicateur pentest
# fort : le pentester voit que ses tentatives sont detectees et logees.

$rejectedHtml = ''
if ($rejectedJsons.Count -gt 0) {
    $rejectedRows = $rejectedJsons | ForEach-Object {
        $iconColor = if ($_.Level -eq 'Critical') { '#ff6b6b' } else { '#ffa502' }
        $fileSafe   = Get-SafeString $_.File
        $reasonSafe = Get-SafeString $_.Reason
        $detailSafe = if ($_.Detail) { Get-SafeString $_.Detail } else { '' }
        @"
        <div class="rejected-row">
            <span class="rejected-dot" style="background:$iconColor"></span>
            <div class="rejected-content">
                <div class="rejected-file">$fileSafe</div>
                <div class="rejected-reason"><strong>$reasonSafe</strong></div>
                <div class="rejected-detail">$detailSafe</div>
            </div>
        </div>
"@
    }

    $rejectedHtml = @"
<div class="rejected-panel" id="rejectedPanel">
    <div class="rejected-header">
        <span class="rejected-title">&#9940; Sanity-checks v2.0 : $($rejectedJsons.Count) JSON rejet&eacute;(s)</span>
        <button class="rejected-toggle" onclick="document.getElementById('rejectedPanel').classList.toggle('collapsed')">D&eacute;velopper / R&eacute;duire</button>
    </div>
    <div class="rejected-info">
        Ces fichiers ont &eacute;t&eacute; <strong>refus&eacute;s</strong> par les sanity-checks et n'apparaissent <strong>pas</strong> dans le rapport.
        Causes possibles : JSON corrompu, version de schema obsol&egrave;te, ou tentative de spoofing (cf. <code>SECURITY.md</code>).
    </div>
    <div class="rejected-list">
        $($rejectedRows -join "`n")
    </div>
</div>
"@
}

# v2.1.11 : le panneau "Anomalies detectees" (bandeau jaune deplie en tete) a ete
# RETIRE. Les anomalies sont desormais signalees PAR PC : un badge sur la ligne du
# tableau (+ tooltip au survol) et un filtre "Anomalies" en un clic dans la barre de
# filtres. Le calcul reste (Test-PCPulseAnomalies -> $anomalyJsons) ; ces anomalies
# sont injectees par PC dans l'embed (champ Anomalies) plus bas.
$anomalyHtml = ''

# v2.1.11 : bouton-filtre "Anomalies (N)" dans la barre de filtres (remplace le
# point d'entree qu'etait le bandeau). N = nombre de PC impactes. Absent si aucune.
$anomalyFilterBtn = if ($anomalyJsons.Count -gt 0) {
    $nbAnomPC = @($anomalyJsons | Select-Object -ExpandProperty PC -Unique).Count
    "<button class=""filter-btn filter-btn-anomaly"" id=""anomalyFilterBtn"" onclick=""toggleKpiFilter('anomaly')"" title=""N'afficher que les PC avec une anomalie (donnee a verifier)"">&#9888;&#65039; Anomalies ($nbAnomPC)</button>"
} else { '' }

# ============================================================
# HTML COMPLET
# ============================================================
$html = @"
<!DOCTYPE html>
<html lang="fr" data-theme="light">
<head>
<meta charset="UTF-8">

<!-- v2.0 : Content-Security-Policy stricte (defense en profondeur des sanity-checks).
     Ce rapport HTML est genere localement et est entierement autonome :
       - JS inline (pas de fichier .js externe)
       - CSS inline (pas de fichier .css externe)
       - Favicon en data URI (base64 SVG)
       - Aucun CDN, aucun fetch/AJAX, aucune iframe
     La CSP ferme donc toutes les voies non utilisees. Si malgre les sanity-checks
     un caractere bizarre passait dans le DOM, le navigateur :
       - refuserait de charger un script externe (pas dans script-src)
       - refuserait toute communication reseau (connect-src 'none')
       - empecherait l'exfiltration de donnees (form-action 'none')
     Voir SECURITY.md > "Sanity-checks Dashboard" pour le rationale complet.
     Note : 'unsafe-inline' est requis pour script-src/style-src car le HTML
     genere contient du code inline (par design : zero dependance externe). -->
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data: 'self'; font-src 'self'; connect-src 'none'; frame-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none';">

<meta name="viewport" content="width=device-width, initial-scale=1.0">
$metaRefresh
<link rel="icon" type="image/svg+xml" href="data:image/svg+xml;base64,$faviconB64">
<title>$titleHtml Dashboard</title>
<style>
    /* ===== VARIABLES THEME =====
       v5.6.1 : palette dark retravaillee pour ameliorer le contraste.
       Ancienne palette : text-dim #aaa text-muted #888 text-faint #666
       -> ratio de contraste insuffisant (en dessous du seuil WCAG AA
       pour les textes secondaires). Nouvelle palette : gris plus clairs
       pour matcher la lisibilite du theme clair.
    */
    :root {
        --bg-main:      #0a0a0a;    /* v2.1.1 : dark neutre - noir profond (plus de teinte bleue) */
        --bg-panel:     #161616;    /* cartes : contraste net avec le fond */
        --bg-elevated:  #202020;    /* row-detail : passe au 1er plan (etait + sombre que tout) */
        --border:       #2b2b2b;    /* bordure visible pour delimiter les cartes (pas d'ombre en dark) */
        --text-main:    #f5f5f5;    /* blanc neutre (etait legerement bleute) */
        --text-dim:     #c9c9ce;
        --text-muted:   #a1a1aa;    /* labels - contraste eleve sur fond noir */
        --text-faint:   #76767e;    /* WCAG AA */
        --text-ghost:   #585860;    /* placeholders */
        --accent:       #8d85ff;    /* inchange : seul le FOND bleu posait probleme */
        --green:        #5dd4a8;
        --orange:       #ffb84d;
        --red:          #ff7a7a;
        --pink:         #ffa8f0;
        --cyan:         #38bcd2;
        --purple:       #b570f7;
        --yellow:       #ffc555;
        --bg-danger:    #241a1a;    /* cartes en alerte : teinte sombre calee sur le niveau panel */
        --bg-hw:        #1f1726;
        --bg-warning:   #232017;
        --bg-success:   #15211b;
    }
    [data-theme="light"] {
        --bg-main:      #f4f5fa;
        --bg-panel:     #ffffff;
        --bg-elevated:  #eef0f7;
        --border:       #d8dbe8;
        --text-main:    #1a1a2e;
        --text-dim:     #495266;
        --text-muted:   #6b7380;
        --text-faint:   #8890a0;
        --text-ghost:   #a8b0c0;
        --accent:       #5b52e0;
        --green:        #16a085;
        --orange:       #e67e22;
        --red:          #d63446;
        --pink:         #c941a0;
        --cyan:         #138496;
        --purple:       #8e44ad;
        --yellow:       #d17200;
        --bg-danger:    #fce8eb;
        --bg-hw:        #f3e9fb;
        --bg-warning:   #fdf3dc;
        --bg-success:   #e0f4ec;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background: var(--bg-main);
        color: var(--text-main);
        padding: 24px;
        min-height: 100vh;
        transition: background 0.2s, color 0.2s;
    }
    .header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 24px;
        border-bottom: 1px solid var(--border);
        padding-bottom: 16px;
    }
    .header h1 { font-size: 24px; font-weight: 400; color: var(--accent); }
    .header h1 span { font-weight: 700; color: var(--text-main); }
    .header-right { display: flex; align-items: center; gap: 16px; }
    .timestamp { color: var(--text-muted); font-size: 12px; }

    .theme-toggle {
        background: var(--bg-panel);
        border: 1px solid var(--border);
        color: var(--text-dim);
        padding: 6px 10px;
        border-radius: 6px;
        cursor: pointer;
        font-size: 14px;
        transition: all 0.2s;
    }
    .theme-toggle:hover { border-color: var(--accent); color: var(--accent); }

    /* ===== BARRE DE FILTRES ===== */
    .toolbar {
        display: flex;
        align-items: center;
        gap: 6px;
        padding: 10px 20px;
        background: var(--bg-panel);
        border-radius: 10px 10px 0 0;
        flex-wrap: wrap;
    }
    .toolbar .label { color: var(--text-muted); font-size: 12px; margin-right: 6px; }
    .range-btn, .filter-btn {
        padding: 5px 10px;
        border-radius: 20px;
        border: 1px solid var(--border);
        background: transparent;
        color: var(--text-dim);
        cursor: pointer;
        font-size: 12px;
        transition: all 0.15s;
    }
    .range-btn:hover, .filter-btn:hover { border-color: var(--accent); color: var(--accent); }
    .range-btn.active { background: var(--accent); color: white; border-color: var(--accent); }
    .filter-btn.active { background: var(--accent); color: white; border-color: var(--accent); }
    /* v2.5.2 : filtre chassis par icones (Laptop/Desktop/AIO) - toggle, pas un menu.
       Reutilise les glyphes du tableau. Inactif = attenue, actif = anneau accent. */
    .chassis-filter-btn {
        padding: 4px 9px;
        border-radius: 20px;
        border: 1px solid var(--border);
        background: transparent;
        cursor: pointer;
        font-size: 14px;
        line-height: 1;
        opacity: 0.6;
        transition: all 0.15s ease;
    }
    .chassis-filter-btn:hover { border-color: var(--accent); opacity: 1; }
    .chassis-filter-btn.active { border-color: var(--accent); background: var(--accent); opacity: 1; }

    /* v2.1.11 : badge anomalie sur la ligne PC (donnee a verifier) + bouton-filtre associe.
       Registre orange (--orange) pour rester distinct de l'axe sante (score/pastilles). */
    .anomaly-badge {
        cursor: pointer;
        color: var(--orange);
        font-size: 13px;
        margin-left: 5px;
        vertical-align: middle;
        opacity: 0.85;
    }
    .anomaly-badge:hover { opacity: 1; }
    /* v2.4.0 : badge decommission (cycle de vie) */
    .decom-badge { cursor: pointer; color: var(--purple); font-size: 13px; margin-left: 5px; vertical-align: middle; opacity: 0.9; }
    .decom-badge:hover { opacity: 1; }
    .decom-badge.late, .decom-badge.done-online { color: var(--red); font-weight: 700; }
    .decom-badge.done { color: var(--text-muted); cursor: default; }
    /* v2.4.2 : panneau stats/audit cycle de vie */
    .decom-stat-row { display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 14px; }
    .decom-stat-tile { flex: 1 1 140px; background: var(--bg-elevated, rgba(255,255,255,0.03)); border: 1px solid var(--border, rgba(128,128,128,0.2)); border-radius: 8px; padding: 12px 14px; }
    .decom-stat-val { font-size: 24px; font-weight: 700; line-height: 1.1; }
    .decom-stat-lbl { font-size: 11px; color: var(--text-muted); margin-top: 4px; }
    .decom-stat-tile.ok     .decom-stat-val { color: var(--green); }
    .decom-stat-tile.warn   .decom-stat-val { color: var(--orange); }
    .decom-stat-tile.danger .decom-stat-val { color: var(--red); }
    /* v2.4.7 : les deux sous-blocs (tableau technicien + barres par mois) sont
       EMPILES VERTICALEMENT. En cote-a-cote, la table debordait sa colonne et
       recouvrait la barre "par mois" (chevauchement). L'empilement rend tout
       chevauchement horizontal structurellement impossible. */
    .decom-sub-row { display: flex; flex-direction: column; gap: 20px; }
    .decom-sub { min-width: 0; }
    /* v2.4.7 : repartition modeles en barres classees compactes (vision rapide,
       moins invasif qu'une tuile par modele). Liste scrollable, lignes cliquables. */
    .model-dist-head { font-size: 12px; color: var(--text-muted); margin-bottom: 10px; }
    .model-dist-list { display: flex; flex-direction: column; gap: 3px; max-height: 360px; overflow-y: auto; padding-right: 4px; }
    .model-bar-row { display: flex; align-items: center; gap: 10px; font-size: 12px; padding: 3px 6px; border-radius: 5px; }
    .model-bar-name { width: 240px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .model-bar-track { flex: 1; height: 10px; background: rgba(128,128,128,0.12); border-radius: 5px; overflow: hidden; }
    .model-bar-fill { height: 100%; background: var(--purple); border-radius: 5px; }
    .model-bar-val { width: 78px; text-align: right; color: var(--text-muted); font-variant-numeric: tabular-nums; }
    .decom-sub h5 { margin: 0 0 8px; font-size: 12px; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; }
    /* v2.4.7 : table-layout FIXED + largeurs explicites. Sans ca, une table
       width:100% en layout AUTO se dimensionne sur son contenu et DEBORDE la
       piste de grille (min-width:0 sur l'item ne contraint pas la table) : la
       colonne "A faire" atterrissait hors de sa colonne, sur la barre "par mois".
       En fixed, la table est plafonnee a 100% de sa colonne. */
    .decom-table { width: 100%; table-layout: fixed; border-collapse: collapse; font-size: 13px; }
    .decom-table th { text-align: left; font-size: 11px; color: var(--text-muted); border-bottom: 1px solid var(--border, rgba(128,128,128,0.2)); padding: 4px 6px; }
    .decom-table td { padding: 4px 6px; border-bottom: 1px solid var(--border, rgba(128,128,128,0.08)); }
    /* 1re colonne (Technicien) = reste, tronquee proprement ; les 2 colonnes
       chiffrees fixees etroites -> la table ne peut plus s'etaler vers la droite. */
    .decom-table th:first-child, .decom-table td:first-child { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .decom-table th:nth-child(2), .decom-table td:nth-child(2),
    .decom-table th:nth-child(3), .decom-table td:nth-child(3) { width: 88px; }
    .decom-month { display: flex; align-items: center; gap: 8px; margin-bottom: 5px; font-size: 12px; }
    .decom-month-lbl { width: 62px; color: var(--text-muted); font-variant-numeric: tabular-nums; }
    .decom-month-bar { flex: 1; height: 10px; background: rgba(128,128,128,0.12); border-radius: 5px; overflow: hidden; }
    .decom-month-fill { height: 100%; background: var(--purple); border-radius: 5px; }
    .decom-month-val { width: 24px; text-align: right; font-variant-numeric: tabular-nums; }
    .filter-btn-anomaly { border-color: var(--orange); color: var(--orange); }
    .filter-btn-anomaly:hover { border-color: var(--orange); color: var(--orange); background: rgba(255, 165, 2, 0.12); }
    .filter-btn-anomaly.active { background: var(--orange); color: #1a1a1a; border-color: var(--orange); }

    .divider {
        width: 1px;
        height: 20px;
        background: var(--border);
        margin: 0 4px;
    }

    .toolbar select, .toolbar .search-input {
        padding: 5px 10px;
        border-radius: 6px;
        border: 1px solid var(--border);
        background: var(--bg-main);
        color: var(--text-main);
        font-size: 12px;
        cursor: pointer;
    }
    .toolbar select:focus, .toolbar .search-input:focus {
        outline: none;
        border-color: var(--accent);
    }

    /* v2.5.2 : recherche elastique (flex-shrink) pour tenir sur la ligne 1 avec
       tous les filtres ; margin-left:auto la garde collee a droite avec le CSV. */
    .search-wrap { margin-left: auto; position: relative; flex: 0 1 200px; min-width: 150px; }
    .search-input {
        padding: 6px 30px 6px 32px !important;
        border-radius: 20px !important;
        width: 100%;
    }
    .search-icon {
        position: absolute;
        left: 10px;
        top: 50%;
        transform: translateY(-50%);
        color: var(--text-faint);
        font-size: 14px;
        pointer-events: none;
    }

    .export-btn {
        background: var(--green);
        color: #0a1a12;
        border: none;
        padding: 6px 14px;
        border-radius: 6px;
        cursor: pointer;
        font-size: 12px;
        font-weight: 600;
        transition: all 0.15s;
    }
    .export-btn:hover { transform: translateY(-1px); opacity: 0.9; }

    .toolbar-break { flex-basis: 100%; height: 0; margin: 0; }
    .search-clear {
        position: absolute; right: 10px; top: 50%; transform: translateY(-50%);
        color: var(--text-faint); font-size: 15px; cursor: pointer; line-height: 1; user-select: none;
    }
    .search-clear:hover { color: var(--text-main); }
    .search-input:placeholder-shown ~ .search-clear { display: none; }
    .filter-btn:disabled, .filter-btn:disabled:hover {
        opacity: 0.4; cursor: default; border-color: var(--border); color: var(--text-faint);
    }
    #backToTop {
        position: fixed; right: 24px; bottom: 24px; width: 42px; height: 42px;
        border-radius: 50%; border: 1px solid var(--border); background: var(--bg-panel);
        color: var(--text-main); font-size: 20px; cursor: pointer; display: none;
        align-items: center; justify-content: center; box-shadow: 0 2px 10px rgba(0,0,0,0.25); z-index: 50;
    }
    #backToTop:hover { border-color: var(--accent); color: var(--accent); }


    /* ===== KPI CARDS ===== */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
        gap: 14px;
        margin: 16px 0 20px;
    }
    .kpi-card {
        background: linear-gradient(135deg, var(--bg-panel), var(--bg-main));
        padding: 16px 18px;
        border-radius: 10px;
        border: 1px solid var(--border);
        cursor: pointer;
        transition: all 0.2s;
        position: relative;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        border-color: var(--accent);
    }
    .kpi-card.active {
        border-color: var(--accent);
        box-shadow: 0 0 0 2px var(--accent);
    }
    .kpi-card.card-warning   { border-left: 3px solid var(--orange); }
    .kpi-card.card-danger    { border-left: 3px solid var(--red); }
    .kpi-card.card-info      { border-left: 3px solid var(--cyan); }
    .kpi-card.card-hardware  { border-left: 3px solid var(--purple); }
    .kpi-card.card-success   { border-left: 3px solid var(--green); }
    .kpi-value { font-size: 26px; font-weight: 700; margin-bottom: 4px; }
    /* v5.6 : label KPI plus contraste pour la lisibilite
       (avant : text-muted qui se perdait sur fond sombre) */
    .kpi-label {
        font-size: 11px;
        color: var(--text-dim);
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-weight: 600;
    }

    /* ===== SUMMARY BAR (v5.5) : chiffres essentiels en haut =====
       Un seul coup d'oeil pour savoir "combien de PC au total, combien
       en ligne, combien sains". Les autres KPIs sont regroupes plus bas. */
    .summary-bar {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 12px;
        margin: 16px 0 22px;
    }
    .summary-tile {
        background: var(--bg-panel);
        border: 1px solid var(--border);
        border-radius: 10px;
        padding: 9px 16px;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    /* v2.3.3 : tuile cliquable (ex : "En ligne" -> filtre les hors ligne) */
    .summary-tile.clickable { cursor: pointer; transition: border-color .15s, box-shadow .15s; }
    .summary-tile.clickable:hover { border-color: var(--accent); }
    .summary-tile.clickable.active { border-color: var(--accent); box-shadow: inset 0 0 0 1px var(--accent); }
    .summary-icon {
        width: 34px; height: 34px;
        border-radius: 9px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 16px;
        flex-shrink: 0;
    }
    .summary-icon.blue   { background: rgba(108, 99, 255, 0.15); color: var(--accent); }
    .summary-icon.green  { background: rgba(78, 204, 163, 0.15); color: var(--green); }
    .summary-icon.red    { background: rgba(255, 107, 107, 0.15); color: var(--red); }
    .summary-content { flex: 1; min-width: 0; }
    .summary-value { font-size: 22px; font-weight: 700; color: var(--text); line-height: 1.1; }
    /* v5.6 : label summary plus contraste aussi */
    .summary-label { font-size: 11px; color: var(--text-dim); text-transform: uppercase; letter-spacing: 0.5px; margin-top: 2px; font-weight: 600; }
    .summary-sub   { font-size: 10px; color: var(--text-muted); margin-top: 2px; }

    /* ===== KPI GROUPS (v5.5) : 4 familles thematiques =====
       Ameliore la lisibilite quand il y a beaucoup d'indicateurs :
       l'oeil trouve direct la famille qui l'interesse au lieu de
       scanner 15 cartes eparpillees. */
    .kpi-group {
        margin: 0 0 16px 0;
    }
    /* v5.6 : header de groupe plus visible (text-dim au lieu de text-muted) */
    .kpi-group-header {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 8px;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        color: var(--text-dim);
        font-weight: 700;
    }
    .kpi-group-icon { font-size: 14px; opacity: 0.9; }
    .kpi-group-line {
        flex: 1;
        height: 1px;
        background: var(--border);
    }
    .kpi-group .kpi-grid {
        margin: 0;
        /* Les cartes dans un groupe sont un peu plus compactes */
        grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
        gap: 10px;
    }
    /* Dans un groupe, les cartes sont legerement moins hautes */
    .kpi-group .kpi-card { padding: 12px 14px; }
    .kpi-group .kpi-value { font-size: 22px; margin-bottom: 2px; }
    .kpi-group .kpi-label { font-size: 10.5px; }

    /* v5.6 : cartes "vides" attenuees mais LABEL reste lisible
       (seule la valeur s'eclaircit, le label garde son poids).
       v5.6.1 : valeur attenuee passe de text-faint a text-muted pour
       etre lisible meme en dark mode (le "0" etait trop pale avant). */
    .kpi-card.kpi-quiet .kpi-value { color: var(--text-muted) !important; font-weight: 500; }
    .kpi-card.kpi-quiet { opacity: 0.92; }
    .kpi-card.kpi-quiet:hover { opacity: 1; }
    /* v5.6 : quand la carte est calme, la barre coloree a gauche est
       neutralisee pour ne pas creer de faux signal visuel */
    .kpi-card.kpi-quiet.card-warning,
    .kpi-card.kpi-quiet.card-danger,
    .kpi-card.kpi-quiet.card-info,
    .kpi-card.kpi-quiet.card-hardware {
        border-left-color: var(--border);
    }

    /* Cartes "critiques" : value > 0 et theme danger -> plus d'emphase */
    .kpi-card.kpi-loud {
        background: linear-gradient(135deg, rgba(255, 107, 107, 0.08), var(--bg-panel));
    }
    .kpi-card.kpi-loud .kpi-value { font-weight: 800; }

    .color-green  { color: var(--green); }
    .color-red    { color: var(--red); }
    .color-orange { color: var(--orange); }
    .color-blue   { color: var(--accent); }
    .color-cyan   { color: var(--cyan); }
    .color-purple { color: var(--purple); }
    .color-pink   { color: var(--pink); }
    .color-yellow { color: var(--yellow); }

    /* ===== KPI GROUP BUTTONS (v5.6) =====
       Remplace les 4 lignes pleine largeur par une barre horizontale
       de 4 boutons compacts. Gain d'espace vertical important, plus
       facile a scanner en 1 coup d'oeil.
       - Bouton "calme" (aucune alerte)  -> discret (gris)
       - Bouton "chaud"  (alertes > 0)    -> colore + badge
       - Clic = filtre tableau sur cette famille
       - Survol = popover avec les sous-KPIs detailles (chacun cliquable) */
    .kpi-group-bar {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 10px;
        margin: 16px 0 22px;
    }
    .kpi-group-btn {
        position: relative;
        background: var(--bg-panel);
        border: 1px solid var(--border);
        border-left: 3px solid var(--border);
        border-radius: 8px;
        padding: 12px 14px;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 10px;
        transition: border-color 0.15s, transform 0.15s, background 0.15s;
        text-align: left;
        font-family: inherit;
        color: inherit;
    }
    .kpi-group-btn:hover {
        border-color: var(--accent);
        transform: translateY(-1px);
        z-index: 100;   /* v2.1.2 : le transform ci-dessus cree un stacking context ; sans ce z-index le popover passe sous le tableau et la pagination */
    }
    .kpi-group-btn.active {
        border-color: var(--accent);
        box-shadow: 0 0 0 2px var(--accent);
    }
    .kpi-group-btn-icon {
        width: 34px; height: 34px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 16px;
        flex-shrink: 0;
        background: var(--bg-main);
    }
    .kpi-group-btn-content { flex: 1; min-width: 0; }
    .kpi-group-btn-title {
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: var(--text-dim);
        line-height: 1.2;
    }
    .kpi-group-btn-sub {
        font-size: 10.5px;
        color: var(--text-muted);
        margin-top: 2px;
    }
    .kpi-group-btn-count {
        font-size: 22px;
        font-weight: 800;
        color: var(--text-muted);   /* v5.6.1 : etait text-faint, trop pale en dark */
        margin-left: 8px;
    }

    /* Etats colores selon le nombre d'alertes */
    .kpi-group-btn.warn {
        border-left-color: var(--orange);
        background: linear-gradient(135deg, rgba(255, 165, 2, 0.06), var(--bg-panel));
    }
    .kpi-group-btn.warn .kpi-group-btn-icon { background: rgba(255, 165, 2, 0.15); color: var(--orange); }
    .kpi-group-btn.warn .kpi-group-btn-count { color: var(--orange); }
    .kpi-group-btn.danger {
        border-left-color: var(--red);
        background: linear-gradient(135deg, rgba(255, 107, 107, 0.08), var(--bg-panel));
    }
    .kpi-group-btn.danger .kpi-group-btn-icon { background: rgba(255, 107, 107, 0.15); color: var(--red); }
    .kpi-group-btn.danger .kpi-group-btn-count { color: var(--red); }

    /* Teinte discrete sur l'icone selon la famille (visible quand le
       bouton est calme, sinon overridee par warn/danger) */
    .kpi-group-btn.calm.family-security .kpi-group-btn-icon { color: var(--purple); }
    .kpi-group-btn.calm.family-stability .kpi-group-btn-icon { color: var(--pink); }
    .kpi-group-btn.calm.family-performance .kpi-group-btn-icon { color: var(--yellow); }
    .kpi-group-btn.calm.family-material .kpi-group-btn-icon { color: var(--cyan); }
    .kpi-group-btn.calm.family-lifecycle .kpi-group-btn-icon { color: var(--green); }

    /* Popover au survol : detail des sous-KPIs de la famille */
    .kpi-group-popover {
        position: absolute;
        top: calc(100% + 8px);
        left: -1px;
        right: -1px;
        z-index: 50;
        background: var(--bg-panel);
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 10px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.35);
        display: none;
        min-width: 240px;
    }
    [data-theme="light"] .kpi-group-popover {
        box-shadow: 0 8px 24px rgba(0,0,0,0.12);
    }
    .kpi-group-btn:hover .kpi-group-popover {
        display: block;
    }
    /* Pont invisible pour que la souris puisse transiter du bouton
       au popover sans fermer celui-ci */
    .kpi-group-btn:hover::after {
        content: '';
        position: absolute;
        top: 100%;
        left: 0; right: 0;
        height: 10px;
    }
    .kpi-popover-items {
        display: flex;
        flex-direction: column;
        gap: 3px;
    }
    .kpi-popover-item {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 7px 10px;
        border-radius: 5px;
        cursor: pointer;
        transition: background 0.1s;
        font-size: 12px;
    }
    .kpi-popover-item:hover { background: var(--bg-main); }
    .kpi-popover-item.active {
        background: var(--bg-main);
        outline: 1px solid var(--accent);
    }
    .kpi-popover-item .pop-label { color: var(--text-dim); font-weight: 500; }
    .kpi-popover-item .pop-value {
        font-weight: 700;
        font-size: 14px;
        min-width: 28px;
        text-align: right;
        color: var(--text-muted);   /* v5.6.1 : etait text-faint */
    }
    .kpi-popover-item .pop-value.warn   { color: var(--orange); }
    .kpi-popover-item .pop-value.danger { color: var(--red); }

    /* v5.6 : indicateurs tableau en mode "outline" quand tout va bien
       -> seul ce qui est critique (rouge) ressort visuellement */
    .indicator-dot.ok {
        background: transparent;
        color: var(--green);
        border: 1.5px solid var(--green);
        line-height: 11px;
    }

    /* ===== TABLE ===== */
    .table-container {
        background: var(--bg-panel);
        border-radius: 0 0 10px 10px;
        overflow: hidden;
    }
    .table-header {
        padding: 14px 20px;
        border-bottom: 1px solid var(--border);
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .table-header h2 {
        font-size: 14px;
        font-weight: 600;
        color: var(--text-dim);
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .count { color: var(--text-muted); font-size: 12px; }
    .active-filters {
        display: flex;
        gap: 6px;
        flex-wrap: wrap;
        margin-top: 6px;
    }
    .active-filter-chip {
        background: var(--accent);
        color: white;
        padding: 2px 8px 2px 10px;
        border-radius: 12px;
        font-size: 10px;
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }
    .active-filter-chip .close {
        cursor: pointer;
        font-weight: 700;
        opacity: 0.7;
    }
    .active-filter-chip .close:hover { opacity: 1; }

    /* ===== v1.5 : Pagination ===== */
    .pagination-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 16px;
        padding: 12px 16px;
        background: var(--bg-panel);
        border-top: 1px solid var(--border);
        flex-wrap: wrap;
        font-size: 12px;
    }
    .pagination-info {
        color: var(--text-muted);
        font-size: 12px;
        white-space: nowrap;
    }
    .pagination-controls {
        display: flex;
        align-items: center;
        gap: 4px;
        flex-wrap: wrap;
    }
    .pagination-label {
        color: var(--text-muted);
        font-size: 12px;
        margin-right: 4px;
    }
    .pagination-per, .pagination-page, .pagination-nav, .pagination-showmore, .pagination-showall {
        background: var(--bg-main);
        border: 1px solid var(--border);
        color: var(--text-main);
        padding: 5px 10px;
        border-radius: 4px;
        cursor: pointer;
        font-size: 12px;
        font-weight: 500;
        transition: all 0.15s;
    }
    .pagination-per:hover:not(.active),
    .pagination-page:hover:not(.active),
    .pagination-nav:hover:not([disabled]),
    .pagination-showmore:hover,
    .pagination-showall:hover {
        background: var(--bg-hover);
        border-color: var(--accent);
    }
    .pagination-per.active, .pagination-page.active {
        background: var(--accent);
        color: white;
        border-color: var(--accent);
    }
    .pagination-nav[disabled] {
        opacity: 0.4;
        cursor: not-allowed;
    }
    .pagination-sep {
        display: inline-block;
        width: 1px;
        height: 20px;
        background: var(--border);
        margin: 0 6px;
    }
    .pagination-ellipsis {
        color: var(--text-muted);
        padding: 0 4px;
    }
    .pagination-showmore {
        margin-left: 12px;
        background: var(--bg-main);
        color: var(--accent);
        border-color: var(--accent);
    }
    .pagination-showall {
        margin-left: 4px;
        background: transparent;
        color: var(--text-muted);
        font-style: italic;
    }
    /* Responsive : empiler info + controls sur mobile */
    @media (max-width: 900px) {
        .pagination-bar {
            flex-direction: column;
            align-items: flex-start;
        }
        .pagination-controls {
            width: 100%;
        }
    }

    .table-scroll { overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; min-width: 1200px; }
    th, td { padding: 10px 16px; text-align: left; font-size: 13px; }
    th {
        background: var(--bg-main);
        color: var(--text-muted);
        font-weight: 600;
        text-transform: uppercase;
        font-size: 11px;
        letter-spacing: 0.5px;
        border-bottom: 1px solid var(--border);
        user-select: none;
        white-space: nowrap;
    }
    th.sortable { cursor: pointer; transition: color 0.15s; }
    th.sortable:hover { color: var(--accent); }
    th.sort-active { color: var(--accent); }
    th .sort-arrow { margin-left: 4px; font-size: 9px; color: var(--text-ghost); }
    th.sort-active .sort-arrow { color: var(--accent); }

    tbody tr.row-main    { border-bottom: 1px solid var(--border); transition: background 0.1s; }
    tbody tr.row-main:hover { background: var(--bg-main); }
    /* v2.5.0-color : teinte de fond de ligne retiree. La severite ne vit plus que
       dans le badge de score (1 seul signal couleur par ligne). Les lignes restent
       neutres pour tous les etats ; seul le survol change le fond. */

    .badge {
        padding: 1px 8px;
        border-radius: 12px;
        font-size: 10px;
        font-weight: 600;
        display: inline-block;
        white-space: nowrap;
    }
    .badge-online  { background: var(--bg-elevated); color: var(--text-dim); }
    .badge-offline { background: var(--bg-elevated); color: var(--text-dim); }

    .site-badge {
        padding: 1px 7px;
        border-radius: 4px;
        font-size: 10px;
        background: var(--border);
        color: var(--text-dim);
        white-space: nowrap;
    }
    .site-badge.inconnu { opacity: 0.5; font-style: italic; }

    /* ===== SCORE BADGE ===== */
    .score-badge {
        display: inline-block;
        padding: 2px 7px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 700;
        min-width: 28px;
        text-align: center;
    }
    .score-ok      { background: #1a3d1a; color: var(--green); }
    .score-warn    { background: #3d3d1a; color: var(--orange); }
    .score-danger  { background: #3d1a1a; color: var(--red); }
    .score-critic  { background: var(--red); color: white; }
    [data-theme="light"] .score-ok      { background: #d5f2e9; color: var(--green); }
    [data-theme="light"] .score-warn    { background: #fdf3dc; color: var(--orange); }
    [data-theme="light"] .score-danger  { background: #fcdfe3; color: var(--red); }
    [data-theme="light"] .score-critic  { background: var(--red); color: white; }
    /* v2.5.0-color : 4e etat de severite = hors ligne (gris neutre). Base sur les
       variables existantes -> valable en dark et en clair. */
    .score-offline { background: var(--bg-elevated); color: var(--text-muted); }

    .kpi-crash    { color: var(--red); font-weight: 700; font-size: 15px; }
    .kpi-bsod     { color: var(--pink); font-weight: 700; font-size: 15px; }
    .kpi-hardware { color: var(--purple); font-weight: 700; font-size: 15px; }

    .kpi-warning               { color: var(--yellow); font-weight: 700; }




    .uptime-badge, .cpu-badge, .conn-badge, .chassis-badge, .os-badge {
        display: inline-block;
        padding: 2px 6px;
        border-radius: 4px;
        font-size: 10px;
        font-weight: 600;
    }
    .uptime-ok      { background: #1a3d1a; color: var(--green); }
    .uptime-warning { background: #3d3d1a; color: var(--orange); }
    .uptime-danger  { background: #3d1a1a; color: var(--red); }
    /* v2.5.0-color : uptime < 7 j = etat sain -> neutre (pas de vert positif). */
    .uptime-neutral { background: var(--bg-elevated); color: var(--text-muted); }
    .cpu-recent       { background: #1a3d1a; color: var(--green); }
    .cpu-vieillissant { background: #3d3d1a; color: var(--orange); }
    .cpu-ancien       { background: #3d1a1a; color: var(--red); }
    .cpu-inconnu      { background: #2a2a3a; color: var(--text-muted); }
    .conn-ethernet   { background: #1a3d2a; color: var(--green); }
    .conn-wifi       { background: #1a2a3d; color: #6699ff; }
    .conn-autre      { background: #2a2a3a; color: var(--text-dim); }
    .conn-deconnecte { background: #3d1a1a; color: var(--red); }
    /* v2.5.0-ui : icone chassis inline (avant le CPU), sans texte -> mono-ligne */
    .chassis-icon {
        display: inline-block; vertical-align: middle; margin-right: 5px;
        padding: 1px 4px; border-radius: 3px; font-size: 11px; line-height: 1;
        cursor: help;
    }
    /* v2.5.0-ui : build OS affiche en ligne apres le badge produit */
    .os-build { margin-left: 5px; font-size: 10px; color: var(--text-faint); }
    /* v5.8 : chassis badges */
    .chassis-laptop  { background: #1a2a3d; color: #6699ff; }
    .chassis-desktop { background: #2a1a3d; color: var(--purple); }
    .chassis-aio     { background: #3d2a1a; color: var(--orange); }
    .chassis-autre   { background: #2a2a3a; color: var(--text-muted); }
    /* v2.2.1 : OS badges (Windows 10/11) */
    .os-win11   { background: #1a3d2a; color: var(--green); }
    .os-win10   { background: #3d3d1a; color: var(--orange); }
    .os-inconnu { background: #2a2a3a; color: var(--text-muted); }

    [data-theme="light"] .uptime-ok      { background: #d5f2e9; color: var(--green); }
    [data-theme="light"] .uptime-warning { background: #fdf3dc; color: var(--orange); }
    [data-theme="light"] .uptime-danger  { background: #fcdfe3; color: var(--red); }
    [data-theme="light"] .cpu-recent       { background: #d5f2e9; color: var(--green); }
    [data-theme="light"] .cpu-vieillissant { background: #fdf3dc; color: var(--orange); }
    [data-theme="light"] .cpu-ancien       { background: #fcdfe3; color: var(--red); }
    [data-theme="light"] .cpu-inconnu      { background: #e6e9f2; color: var(--text-muted); }
    [data-theme="light"] .conn-ethernet   { background: #d5f2e9; color: var(--green); }
    [data-theme="light"] .conn-wifi       { background: #dde6f7; color: #3a6cbf; }
    [data-theme="light"] .conn-autre      { background: #e6e9f2; color: var(--text-dim); }
    [data-theme="light"] .conn-deconnecte { background: #fcdfe3; color: var(--red); }
    [data-theme="light"] .chassis-laptop  { background: #dde6f7; color: #3a6cbf; }
    [data-theme="light"] .chassis-desktop { background: #f3e9fb; color: var(--purple); }
    [data-theme="light"] .chassis-aio     { background: #fdf0dc; color: var(--orange); }
    [data-theme="light"] .chassis-autre   { background: #e6e9f2; color: var(--text-muted); }
    [data-theme="light"] .os-win11   { background: #d5f2e9; color: var(--green); }
    [data-theme="light"] .os-win10   { background: #fdf3dc; color: var(--orange); }
    [data-theme="light"] .os-inconnu { background: #e6e9f2; color: var(--text-muted); }

    /* v2.5.0-ui3 : anciennes barres de remplissage disque retirees (remplacees
       par des .kv-row "Remplissage C: -> 45 % . 215 Go libres"). CSS supprime. */

    /* ===== FRESHNESS (derniere activite) ===== */
    /* v2.5.0-color : "vu recemment" n'est pas une alerte -> neutre (etait vert). */
    .freshness-ok       { color: var(--text-muted);  font-size: 11px; }
    .freshness-warning  { color: var(--orange); font-size: 11px; }
    .freshness-danger   { color: var(--red);    font-size: 11px; }

    /* ===== BOOT ===== */
    .boot-badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 11px; font-weight: 700; }
    .boot-ok      { background: #1a3d1a; color: var(--green); }
    .boot-warning { background: #3d3d1a; color: var(--orange); }
    .boot-danger  { background: #3d1a1a; color: var(--red); }
    .boot-date    { font-size: 11px; color: var(--text-faint); margin-top: 2px; }
    [data-theme="light"] .boot-ok      { background: #d5f2e9; color: var(--green); }
    [data-theme="light"] .boot-warning { background: #fdf3dc; color: var(--orange); }
    [data-theme="light"] .boot-danger  { background: #fcdfe3; color: var(--red); }

    /* v2.3.1 (#9) : ligne lisible sous le crasher (nature/code/origine en clair) */
    .crasher-detail { font-size: 10px; color: var(--text-muted); margin-top: 2px; line-height: 1.35; }

    /* ===== DRILL-DOWN ===== */
    .row-main td:first-child { cursor: pointer; user-select: none; }
    .row-main td:first-child:hover { color: var(--accent); }
    .toggle-icon { display: inline-block; margin-right: 6px; font-size: 10px; color: var(--text-muted); transition: transform 0.2s; }
    .row-main.open .toggle-icon { transform: rotate(90deg); }
    /* v2.5.0-ui3 : le panneau deplie doit lire comme une "fiche technique" sur
       surface neutre claire (var --bg-panel, blanc en clair / surface sombre en
       dark), jamais la teinte d'alerte de la ligne. On force la surface + les
       couleurs de texte normales sur la ligne de detail ET sa cellule. */
    .row-detail { display: none; background: var(--bg-panel) !important; }
    .row-detail.visible { display: table-row; }
    .row-detail td { padding: 0 !important; border-bottom: 1px solid var(--border) !important; white-space: normal !important; background: var(--bg-panel) !important; color: var(--text-dim); }

    /* ============================================================
       COCKPIT UI : file d'actions a droite du tableau.
       Le rail de navigation a ete retire (redondant avec les tuiles de
       famille + la toolbar) ; la page reste pleine largeur.
       Aucune donnee ni fonction metier modifiee.
       ============================================================ */

    /* File d'actions a droite du tableau */
    .ck-body { display: grid; grid-template-columns: 1fr 320px; gap: 16px; align-items: start; }
    @media (max-width: 1200px) { .ck-body { grid-template-columns: 1fr; } }
    .ck-body > .table-container { min-width: 0; }
    .ck-actions { background: var(--bg-panel); border: 1px solid var(--border); border-radius: 10px; overflow: hidden; position: sticky; top: 12px; }
    .ck-actions > h3 { margin: 0; padding: 13px 16px; font-size: 12px; text-transform: uppercase; letter-spacing: .8px; color: var(--text-dim); border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: baseline; }
    .ck-actions > h3 small { text-transform: none; letter-spacing: 0; color: var(--text-faint); font-weight: 400; font-size: 10px; }
    .ck-act { padding: 11px 14px; border-bottom: 1px solid var(--border); cursor: pointer; display: flex; gap: 11px; align-items: flex-start; transition: background .12s; }
    .ck-act:last-child { border-bottom: 0; }
    .ck-act:hover { background: var(--bg-elevated); }
    .ck-act.on { background: var(--bg-elevated); box-shadow: inset 3px 0 0 var(--ck-c, var(--accent)); }
    .ck-act-num { flex: 0 0 44px; text-align: center; font-size: 18px; font-weight: 700; color: var(--ck-c, var(--accent)); font-variant-numeric: tabular-nums; line-height: 1.15; }
    .ck-act-num small { display: block; font-size: 9px; font-weight: 500; letter-spacing: .4px; color: var(--text-faint); text-transform: uppercase; }
    .ck-act-txt { min-width: 0; }
    .ck-act-txt b { display: block; font-size: 12.5px; font-weight: 600; color: var(--text-main); }
    .ck-act-txt span { font-size: 11px; color: var(--text-faint); line-height: 1.4; display: block; margin-top: 2px; }
    .ck-act-empty { padding: 22px 16px; text-align: center; color: var(--text-faint); font-size: 12px; }

    /* Sous-section "Mon equipe" (bas de la file d'actions) : remplace l'ancien rail. */
    .ck-team { border-top: 3px solid var(--border); }
    .ck-team-head {
        padding: 11px 14px 7px; font-size: 10.5px; text-transform: uppercase; letter-spacing: .7px;
        color: var(--text-dim); font-weight: 700; display: flex; justify-content: space-between; align-items: baseline;
    }
    .ck-team-head small { text-transform: none; letter-spacing: 0; color: var(--text-faint); font-weight: 400; font-size: 10px; }
    .ck-team-row {
        display: flex; align-items: center; justify-content: space-between; gap: 8px;
        padding: 7px 14px; border-top: 1px solid var(--border); cursor: pointer;
        font-size: 12.5px; color: var(--text-dim); transition: background .12s;
    }
    .ck-team-row:hover { background: var(--bg-elevated); color: var(--text-main); }
    .ck-team-row.on { background: var(--bg-elevated); box-shadow: inset 3px 0 0 var(--accent); color: var(--text-main); font-weight: 600; }
    .ck-team-name { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .ck-team-n { flex: 0 0 auto; font-size: 11px; background: var(--bg-elevated); border: 1px solid var(--border);
        border-radius: 20px; padding: 1px 8px; font-variant-numeric: tabular-nums; color: var(--text-muted); }
    .ck-team-row.on .ck-team-n { background: var(--accent); border-color: var(--accent); color: #fff; }

    /* Anneaux (donut) SVG dans les tuiles de synthese */
    .ck-ring { flex-shrink: 0; }

    /* ===== Lignes du tableau principal resserrees (densite cockpit) =====
       Scope strict : uniquement .row-main. Les lignes de detail (.row-detail)
       gardent leur propre padding (voir plus bas). */
    thead th { padding-top: 8px; padding-bottom: 8px; }
    tbody tr.row-main td {
        padding: 6px 10px;
        font-size: 12px;
        line-height: 1.35;
        vertical-align: middle;
        white-space: nowrap;   /* v2.5.0-ui : chaque cellule sur une seule ligne (densite cockpit) */
    }



    /* Pastilles indicateurs compacts (colonne tableau)
       v5.7 : gap 5px (etait 3px) pour mieux separer les pastilles.
       Avant elles se collaient et donnaient un bloc illisible. */
    .indicator-row {
        display: inline-flex;
        gap: 5px;
        align-items: center;
        padding: 2px 4px;
    }
    .indicator-dot {
        display: inline-block; width: 15px; height: 15px; border-radius: 50%;
        font-size: 9px; font-weight: 700; text-align: center; line-height: 15px;
        color: #fff; cursor: help; user-select: none;
        flex-shrink: 0;
    }
    .indicator-dot.ok   { background: var(--green); }
    .indicator-dot.warn { background: var(--orange); }
    .indicator-dot.ko   { background: var(--red); }
    .indicator-dot.na   { background: var(--border); color: var(--text-muted); }




    /* v2.3.0 : dernier utilisateur connu (pas de session live) */
    .user-last { color: var(--text-muted); font-style: italic; white-space: nowrap; }
    /* v2.5.0-ui : petite puce "DERNIER" inline qui ne passe jamais a la ligne */
    .user-last-tag { display: inline-block; font-style: normal; font-size: 8.5px; text-transform: uppercase; letter-spacing: 0.4px;
                     color: var(--text-ghost); border: 1px solid var(--border); border-radius: 3px;
                     padding: 0 4px; margin-left: 5px; vertical-align: middle; white-space: nowrap; }

    /* ===== INVENTAIRE MONITORS GLOBAL (panneau bas de page) ===== */
    .monitor-inventory {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 10px;
    }
    /* v2.1.2 : bandeau + curseur de seuil d'age des ecrans secondaires */
    .monitor-panel-head {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        flex-wrap: wrap;
        margin-bottom: 14px;
    }
    #monitorPanel h3 { margin: 0; }
    .screen-age-control {
        display: flex;
        align-items: center;
        gap: 10px;
        font-size: 12px;
        color: var(--text-dim);
    }
    .screen-age-control label { white-space: nowrap; text-transform: uppercase; letter-spacing: 0.4px; font-size: 11px; }
    .screen-age-control input[type="range"] { width: 130px; accent-color: var(--accent); cursor: pointer; }
    .screen-age-control .sac-val { color: var(--accent); font-weight: 600; white-space: nowrap; }
    .monitor-inv-tile {
        background: var(--bg-main);
        border: 1px solid var(--border);
        border-radius: 6px;
        padding: 10px 12px;
    }
    .monitor-inv-value {
        font-size: 22px;
        font-weight: 700;
        color: var(--text);
    }
    .monitor-inv-label {
        font-size: 10.5px;
        color: var(--text-muted);
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-top: 2px;
    }
    .monitor-inv-sub {
        font-size: 10.5px;
        color: var(--text-dim);
        margin-top: 4px;
        line-height: 1.5;
    }
    .monitor-inv-sub .chip {
        display: inline-block;
        background: var(--bg-panel);
        padding: 2px 6px;
        border-radius: 3px;
        margin: 2px 3px 2px 0;
        font-size: 10px;
    }

    /* v2.5.0-ui : etat "rien a signaler" = note discrete sur une ligne, pas une carte vide */
    .detail-empty { color: var(--text-muted); font-style: normal; font-size: 11.5px; padding: 2px 0; line-height: 1.4; }

    /* v2.5.0 (Essai A) : DRILL-DOWN = liste de definitions mono-ligne unifiee.
       "Fiche technique" plate. Une donnee = une ligne : gouttiere de libelle a
       gauche (largeur fixe) + valeur condensee a droite, filet pointille, en-tetes
       de groupe legers. Meme format partout -> coherence maximale. Remplace les
       ex-layouts cartes / kv-row / detail-section / detail-item. */
    .dd-list { padding: 2px 0; }
    .dd-group { font-size:11px; letter-spacing:.06em; text-transform:uppercase; color:var(--text-dim); font-weight:600; margin:12px 0 3px; padding-bottom:3px; border-bottom:0.5px solid var(--border); }
    .dd-strong { font-weight:600; color:var(--text-dim); }
    .dd-group:first-child { margin-top:0; }
    .dd-row { display:grid; grid-template-columns:88px 1fr; gap:12px; align-items:baseline; padding:4px 0; border-bottom:0.5px dotted var(--border); }
    .dd-row:last-child { border-bottom:0; }
    .dd-l { font-size:10.5px; text-transform:uppercase; letter-spacing:.03em; color:var(--text-muted); overflow-wrap:anywhere; }
    .dd-v { color:var(--text-dim); overflow-wrap:anywhere; }
    .dd-v .mut { color:var(--text-muted); }
    .dd-v .mono { font-family: ui-monospace, 'SF Mono', Consolas, monospace; }
    .dd-v .ok { color:var(--green); }
    .dd-v .warn { color:var(--orange); }
    .dd-v .ko { color:var(--red); }
    .dd-tag { display:inline-block; font-size:10px; padding:1px 7px; border-radius:999px; margin-left:3px; white-space:nowrap; background:var(--bg-elevated); color:var(--text-muted); }
    .dd-tag.ok { background:var(--bg-success); color:var(--green); }
    .dd-tag.warn { background:var(--bg-warning); color:var(--orange); }
    .dd-tag.ko { background:var(--bg-danger); color:var(--red); }
    .dd-badge { display:inline-block; font-size:10px; padding:1px 7px; border-radius:6px; background:rgba(141,133,255,0.15); color:var(--accent); }

    /* ===== ONGLETS DRILL-DOWN (v5.5) =====
       Les 10 sections de detail sont regroupees en 5 onglets thematiques :
       Vue d'ensemble / Stabilite / Demarrage / Materiel / Securite.
       Evite le "mur de cartes" ou tout avait le meme poids visuel. */
    .detail-tabs {
        display: flex;
        gap: 4px;
        border-bottom: 0.5px solid var(--border);
        padding: 0 20px 0 36px;
        margin: 8px 0 0 0;
        overflow-x: auto;
        scrollbar-width: thin;
    }
    .detail-tab {
        padding: 8px 14px;
        cursor: pointer;
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: var(--text-muted);
        border-bottom: 2px solid transparent;
        margin-bottom: -1px;
        transition: color 0.15s, border-color 0.15s;
        white-space: nowrap;
        display: flex;
        align-items: center;
        gap: 6px;
        background: none;
        border-left: none;
        border-right: none;
        border-top: none;
    }
    .detail-tab:hover { color: var(--text-dim); }
    .detail-tab.active {
        color: var(--accent);
        border-bottom-color: var(--accent);
    }
    /* Badge de compteur sur les onglets (nb d'alertes dans l'onglet) */
    .detail-tab .tab-badge {
        font-size: 10px;
        background: var(--red);
        color: #fff;
        padding: 1px 5px;
        border-radius: 8px;
        font-weight: 700;
        min-width: 14px;
        text-align: center;
    }
    .detail-tab .tab-badge.quiet {
        background: var(--border);
        color: var(--text-muted);
    }
    .detail-tab-panel {
        display: none;
        padding: 14px 20px 16px 36px;
    }
    .detail-tab-panel.active { display: block; }


    /* ===== MODE COMPACT DU TABLEAU (v5.5) =====
       Par defaut on cache les colonnes "avancees" (Crash/BSOD/HW/Disque/Perf)
       pour alleger le tableau. Un bouton toolbar permet de tout afficher. */
    body:not(.advanced-cols) .col-advanced { display: none !important; }

    /* ===== PANNEAU TOP CRASHERS PARC ===== */
    .global-panel {
        margin-top: 24px;
        background: var(--bg-panel);
        border-radius: 10px;
        padding: 20px;
        border: 1px solid var(--border);
    }
    .global-panel h3 {
        font-size: 14px;
        font-weight: 600;
        color: var(--text-dim);
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 14px;
    }
    /* v2.5 : liste verticale simple (les groupes/entetes rythment la lecture) */
    .gc-list { display: block; }
    .global-crasher-row {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 6px 10px;
        background: transparent;
        border-radius: 6px;
        border-left: 2px solid transparent;
    }
    .global-crasher-name { color: var(--text-dim); font-size: 12px; font-weight: 500; min-width: 0; flex: 0 1 auto; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-family: ui-monospace, 'SF Mono', Consolas, monospace; }
    .global-crasher-stats { font-size: 11px; color: var(--text-muted); white-space: nowrap; }

    /* ===== TOP CRASHERS PARC (v2.5) : 3 groupes priorises =====
       Importantes (applis suivies) / Plus fort impact / Bruit de fond.
       Entetes facon .dd-group, tri par volume, top 5 + expander par groupe. */
    .gc-group {
        font-size: 11px;
        letter-spacing: .06em;
        text-transform: uppercase;
        color: var(--text-dim);
        font-weight: 600;
        margin: 16px 0 4px;
        padding-bottom: 3px;
        border-bottom: 0.5px solid var(--border);
        display: flex;
        align-items: baseline;
        gap: 8px;
    }
    .gc-group:first-child { margin-top: 0; }
    .gc-group-count { color: var(--text-muted); font-size: 11px; font-weight: 600; font-variant-numeric: tabular-nums; }
    .gc-empty { color: var(--text-muted); font-size: 11.5px; padding: 3px 2px; }
    .gc-hidden { color: var(--text-muted); font-size: 11px; margin-top: 8px; }

    /* Ligne "voir les N autres" / bruit de fond replie */
    .gc-more {
        cursor: pointer;
        user-select: none;
        color: var(--text-muted);
        font-size: 11px;
        padding: 5px 4px 2px;
        display: inline-flex;
        align-items: center;
        gap: 5px;
    }
    .gc-more:hover { color: var(--text-dim); }
    .gc-more-tri { font-size: 9px; }
    .gc-more-body { display: block; }

    /* Chip tag (uppercase, pilule ~9.5px) */
    .gc-tag {
        display: inline-block;
        font-size: 9.5px;
        line-height: 1.5;
        text-transform: uppercase;
        letter-spacing: .03em;
        padding: 0 7px;
        border-radius: 999px;
        white-space: nowrap;
        font-weight: 600;
        background: var(--bg-elevated);
        color: var(--text-muted);
    }
    .gc-tag-suivi { background: rgba(141,133,255,0.15); color: var(--accent); }
    .gc-tag-fail  { background: var(--bg-warning);      color: var(--orange); }

    /* Metriques a droite : nombres en gras. Rouge dans le groupe "Importantes". */
    .gc-spacer { flex: 1 1 auto; }
    .global-crasher-stats b { color: var(--text-main); font-weight: 700; font-variant-numeric: tabular-nums; }
    .gc-important .global-crasher-stats b { color: var(--red); }

    /* Lignes cliquables (tout le panneau) -> filtre les PC concernes */
    .global-crasher-row.clickable { cursor: pointer; transition: background .1s ease, box-shadow .1s ease; }
    .global-crasher-row.clickable:hover { background: var(--bg-elevated); }
    .global-crasher-row.active { background: rgba(141,133,255,0.14); box-shadow: inset 0 0 0 1px var(--accent); }


    /* v2.1.14 : bouton "tout effacer" dans la barre des filtres actifs */
    .clear-all-filters { background: transparent; border: 1px solid var(--text-faint); color: var(--text-muted); font-size: 11px; padding: 3px 10px; border-radius: 12px; cursor: pointer; margin-left: 4px; font-family: inherit; }
    .clear-all-filters:hover { border-color: var(--red); color: var(--red); }

    /* ===== BOOT BREAKDOWN (v5.2) ===== */
    .boot-breakdown { display: flex; flex-direction: column; gap: 10px; }
    .boot-breakdown-bar {
        display: flex;
        height: 28px;
        background: var(--bg-main);
        border-radius: 6px;
        overflow: hidden;
        border: 1px solid var(--border);
    }
    .boot-breakdown-segment {
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-weight: 700;
        font-size: 11px;
        min-width: 0;
        overflow: hidden;
        white-space: nowrap;
        transition: all 0.3s;
    }
    .boot-breakdown-segment.cold    { background: #3a6cbf; }
    .boot-breakdown-segment.fast    { background: var(--yellow); color: #2a1a00; }
    .boot-breakdown-segment.resume  { background: var(--purple); }
    .boot-breakdown-segment.unknown { background: var(--text-faint); }
    .boot-breakdown-legend {
        display: flex;
        gap: 16px;
        font-size: 11px;
        color: var(--text-dim);
        flex-wrap: wrap;
    }
    .boot-breakdown-legend-item { display: flex; align-items: center; gap: 6px; }
    .boot-breakdown-legend-dot { width: 10px; height: 10px; border-radius: 2px; }
    .boot-breakdown-empty { color: var(--text-ghost); font-style: italic; padding: 10px; text-align: center; }
    .boot-breakdown-note {
        font-size: 11px;
        color: var(--text-muted);
        background: var(--bg-main);
        border-left: 3px solid var(--accent);
        padding: 8px 12px;
        border-radius: 4px;
    }

    /* ===== LEGEND ===== */
    .legend { display: flex; gap: 20px; padding: 12px 20px; font-size: 11px; color: var(--text-faint); flex-wrap: wrap; border-top: 1px solid var(--border); }
    .legend-item { display: flex; align-items: center; gap: 6px; }
    .legend-dot { width: 10px; height: 10px; border-radius: 3px; }
    .dot-danger   { background: #ff6b6b33; border: 1px solid var(--red); }
    .dot-warning  { background: #ffa50233; border: 1px solid var(--orange); }
    .dot-hardware { background: #a855f733; border: 1px solid var(--purple); }
    .dot-ok       { background: #4ecca333; border: 1px solid var(--green); }

    /* v2.0 : panneau "JSON suspects" (rejets sanity-checks) */
    .rejected-panel {
        background: linear-gradient(135deg, rgba(255, 107, 107, 0.12), rgba(255, 165, 2, 0.08));
        border: 1px solid var(--red);
        border-left: 4px solid var(--red);
        border-radius: 8px;
        padding: 16px 20px;
        margin: 16px 24px;
    }
    .rejected-panel.collapsed .rejected-list,
    .rejected-panel.collapsed .rejected-info {
        display: none;
    }
    .rejected-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 8px;
    }
    .rejected-title {
        font-weight: 600;
        font-size: 14px;
        color: var(--red);
    }
    .rejected-toggle {
        background: transparent;
        border: 1px solid var(--text-muted);
        color: var(--text-muted);
        padding: 4px 12px;
        border-radius: 4px;
        cursor: pointer;
        font-size: 11px;
    }
    .rejected-toggle:hover {
        border-color: var(--text);
        color: var(--text);
    }
    .rejected-info {
        font-size: 12px;
        color: var(--text-muted);
        margin-bottom: 12px;
        line-height: 1.5;
    }
    .rejected-info code {
        background: rgba(255,255,255,0.08);
        padding: 1px 5px;
        border-radius: 3px;
        font-size: 11px;
    }
    .rejected-list {
        display: grid;
        gap: 8px;
    }
    .rejected-row {
        display: flex;
        align-items: flex-start;
        gap: 10px;
        background: rgba(0,0,0,0.18);
        padding: 10px 12px;
        border-radius: 6px;
    }
    .rejected-dot {
        flex-shrink: 0;
        width: 10px;
        height: 10px;
        border-radius: 50%;
        margin-top: 4px;
    }
    .rejected-content {
        flex: 1;
        min-width: 0;
    }
    .rejected-file {
        font-family: 'SF Mono', Consolas, monospace;
        font-size: 12px;
        color: var(--text);
        margin-bottom: 2px;
    }
    .rejected-reason {
        font-size: 13px;
        color: var(--text);
        margin-bottom: 2px;
    }
    .rejected-detail {
        font-size: 11px;
        color: var(--text-muted);
        font-family: 'SF Mono', Consolas, monospace;
        word-break: break-word;
    }

</style>
</head>
<body>

<!-- ============================================================
     COCKPIT UI : page pleine largeur (le rail de gauche
     a ete retire). Tableau + file d'actions dans .ck-body plus bas.
     ============================================================ -->

<div class="header">
    <h1><span>$titleHtml</span> &mdash; $subtitleHtml</h1>
    <div class="header-right">
        <button class="theme-toggle" id="themeToggle" onclick="toggleTheme()" title="Basculer clair/sombre">&#9789;</button>
        <div class="timestamp">G&eacute;n&eacute;r&eacute; le : $($now.ToString('dd/MM/yyyy HH:mm'))</div>
    </div>
</div>

$rejectedHtml

$anomalyHtml

<div class="toolbar">
    <span class="label">P&eacute;riode :</span>
    <button class="range-btn active" onclick="setDays(1, this)">24h</button>
    <button class="range-btn" onclick="setDays(7, this)">7j</button>
    <button class="range-btn" onclick="setDays(15, this)">15j</button>
    <button class="range-btn" onclick="setDays(30, this)">30j</button>

    <div class="divider"></div>

    <label style="color: var(--text-muted); font-size: 12px;">Site :</label>
    <select id="siteFilter" onchange="filterChanged()"><option value="">Tous</option></select>
    <label style="color: var(--text-muted); font-size: 12px;">$cpuFilterLabelTxt</label>
    <select id="cpuFilter" onchange="filterChanged()">
$cpuFilterOptions
    </select>
    <label style="color: var(--text-muted); font-size: 12px;">OS :</label>
    <select id="osFilter" onchange="filterChanged()"><option value="">Tous</option></select>
    <label style="color: var(--text-muted); font-size: 12px;">Mod&egrave;le :</label>
    <select id="modelFilter" onchange="filterChanged()"><option value="">Tous</option></select>
    <label style="color: var(--text-muted); font-size: 12px;">VPN :</label>
    <select id="vpnFilter" onchange="filterChanged()"><option value="">Toutes</option></select>

    <div class="divider"></div>
    <label style="color: var(--text-muted); font-size: 12px;">Type :</label>
    <button class="chassis-filter-btn" id="chassisLaptop"  onclick="toggleChassisFilter('laptop')"  title="Portables">&#128187;</button>
    <button class="chassis-filter-btn" id="chassisDesktop" onclick="toggleChassisFilter('desktop')" title="Postes fixes">&#128421;</button>
    <button class="chassis-filter-btn" id="chassisAio"     onclick="toggleChassisFilter('aio')"     title="Tout-en-un (AIO)">&#128444;</button>

    <div class="search-wrap">
        <span class="search-icon">&#128269;</span>
        <input class="search-input" type="text" id="searchInput" placeholder="PC, utilisateur, n&deg; s&eacute;rie..." oninput="filterChanged()">
        <span class="search-clear" id="searchClear" onclick="clearSearch()" title="Effacer la recherche">&times;</span>
    </div>
    <button class="export-btn" onclick="exportCSV()" title="Exporter la vue courante en CSV">&#8681; CSV</button>

    <div class="toolbar-break"></div>

    <button class="filter-btn" id="maskHealthyBtn" onclick="toggleMaskHealthy()" title="Masquer les PC sans probleme">Masquer sains</button>
    <button class="filter-btn" id="advancedColsBtn" onclick="toggleAdvancedCols()" title="Afficher les colonnes techniques (Crash, BSOD, HW, Disque, Perf)">Vue d&eacute;taill&eacute;e</button>
    $anomalyFilterBtn

    <button class="filter-btn" id="resetBtn" onclick="clearAllFilters()" title="R&eacute;initialiser tous les filtres" style="margin-left:auto">R&eacute;initialiser</button>
</div>

<!-- v5.5 : Summary bar avec les 3 chiffres essentiels -->
<div class="summary-bar" id="summaryBar"></div>

<!-- v5.6 : barre horizontale de 4 boutons de groupes (compact)
     remplace les 4 lignes pleine largeur de v5.5 -->
<div class="kpi-group-bar" id="kpiGroupBar"></div>

<!-- Garde pour retrocompat : kpiGrid devient cache, les groupes ci-dessus le remplacent -->
<div class="kpi-grid" id="kpiGrid" style="display:none"></div>
<div id="kpiGroups" style="display:none"></div>

<!-- cockpit : tableau (gauche) + file d'actions (droite) -->
<div class="ck-body">
<div class="table-container">
    <div class="table-header">
        <div>
            <h2>D&eacute;tails par appareil</h2>
            <div class="active-filters" id="activeFilters"></div>
        </div>
        <span class="count" id="pcCount"></span>
    </div>
    <!-- v1.5 : pagination haut -->
    <div id="paginationTop"></div>
    <div class="table-scroll" id="deviceTable">
        <table>
            <thead id="tableHead"></thead>
            <tbody id="tableBody"></tbody>
        </table>
    </div>
    <!-- v1.5 : pagination bas -->
    <div id="paginationBottom"></div>
    <div class="legend">
        <div class="legend-item"><div class="legend-dot" style="background:var(--green)"></div> Sain</div>
        <div class="legend-item"><div class="legend-dot" style="background:var(--orange)"></div> &Agrave; surveiller</div>
        <div class="legend-item"><div class="legend-dot" style="background:var(--red)"></div> Critique</div>
        <div class="legend-item"><div class="legend-dot" style="background:var(--text-muted)"></div> Hors ligne</div>
    </div>
</div>

<section class="ck-actions">
    <h3>File d'actions <small>par gravit&eacute;</small></h3>
    <div id="ckActs"></div>
</section>
</div>
<!-- /ck-body -->

<div class="global-panel">
    <h3>R&eacute;partition des d&eacute;marrages (parc)</h3>
    <div id="bootBreakdown" class="boot-breakdown"></div>
</div>

<!-- v2.4.2 : stats / audit du cycle de vie (decommissionnement) -->
<div class="global-panel" id="decomPanel" style="display:none">
    <h3>Cycle de vie du parc &mdash; d&eacute;commissionnement</h3>
    <div id="decomStats"></div>
</div>

<!-- v5.7 : inventaire des moniteurs externes du parc -->
<div class="global-panel" id="monitorPanel" style="display:none">
    <div class="monitor-panel-head">
        <h3>&Eacute;crans secondaires branch&eacute;s</h3>
        <div class="screen-age-control">
            <label for="screenAgeSlider">Seuil de renouvellement</label>
            <input type="range" id="screenAgeSlider" min="3" max="10" step="1" value="7">
            <span class="sac-val"><span id="screenAgeValue">7</span>&nbsp;ans et +</span>
        </div>
    </div>
    <div class="monitor-inventory" id="monitorInventory"></div>
</div>

<!-- v2.4.7 : repartition des modeles machine du parc -->
<div class="global-panel" id="modelPanel" style="display:none">
    <div class="monitor-panel-head"><h3>Mod&egrave;les de machines du parc</h3></div>
    <div id="modelInventory"></div>
</div>

<!-- v2.1.2 : Top Crashers deplace tout en bas de page -->
<div class="global-panel">
    <h3>Top Crashers parc global</h3>
    <div class="gc-list" id="globalCrashers"></div>
</div>

<script>
// ===== DONNEES =====
var pcData        = $jsonEmbed;
var scoreWeights  = $weightsEmbed;
var showSite      = $showSiteJs;
var seuilBootLong    = $SeuilBootLong;
var seuilCrashRecent = $SeuilCrashRecent;
// v2.1.2 : seuil d'age des ecrans secondaires (curseur), persiste en localStorage. Defaut 7 ans, borne 3-10.
var screenAgeThreshold = (function() { var v = parseInt(localStorage.getItem('pcpulse_screenAge'), 10); return (v >= 3 && v <= 10) ? v : 7; })();
var seuilDiskAlert   = $SeuilDiskAlert;
var seuilDiskWarning = $SeuilDiskWarning;
var generatedAt   = new Date('$($now.ToString("yyyy-MM-ddTHH:mm:ss"))');
// v2.5.3 : renouvellement. RENEWAL_MAX_YEAR = null => mode desactive (tag CPU d'age).
var RENEWAL_MAX_YEAR = $renewalMaxYearJs;
var RENEWAL_MODE = (RENEWAL_MAX_YEAR !== null);
function isRenewalCandidate(pc) { return RENEWAL_MODE && pc && pc.CPUYear && pc.CPUYear <= RENEWAL_MAX_YEAR; }
// v2.5.3 : match du filtre "CPU" (mode age) OU "Renouvellement" (mode candidat).
function cpuFilterMatch(p) {
    if (!state.cpuFilter) return true;
    if (RENEWAL_MODE) {
        if (state.cpuFilter === 'renew') return isRenewalCandidate(p.pc);
        if (state.cpuFilter === 'keep')  return !isRenewalCandidate(p.pc);
        return true;
    }
    return p.pc.CPUAgeCategory === state.cpuFilter;
}
// v2.1.12 : applis suivies (minuscules) - remontees dans la section "A investiguer en priorite".
var priorityApps  = $priorityAppsJson;
// v2.4.2 : registre decommission complet (stats/audit cycle de vie).
var decomRegistry = $decomRegistryJson;

// cockpit : mapping PC -> technicien assigne, tire du registre de decommission
// (seule source d'affectation nominative disponible cote client). Sert au
// filtre "Mon equipe" du rail. Ne couvre que les PC presents dans le registre.
var ckTechByPc = (function() {
    var m = {};
    (decomRegistry || []).forEach(function(e) {
        if (e && e.PC && e.AssignedTo) m[e.PC] = e.AssignedTo;
    });
    return m;
})();
// cockpit : metadonnees des familles (rempli par renderKpis) pour cabler le rail.
var ckFamilyMeta = {};

// ===== ETAT UI =====
var state = {
    days: 1,
    daysBtn: null,
    maskHealthy: $maskHealthyJs,
    siteFilter: '',
    cpuFilter: '',
    osFilter: '',
    modelFilter: '',
    vpnFilter: '',   // v2.5.2 : filtre par version de client VPN
    chassisFilter: null,   // v2.5.2 : 'laptop' | 'desktop' | 'aio' | null (filtre icones)
    kpiFilter: null,     // 'offline', 'crash', 'bsod', 'hw', 'bootLong', 'diskAlert', 'oldCpu', 'crashRecent'
    appFilter: null,     // v2.1.12 : nom d'appli (clic sur un crasher) -> filtre les PC qui l'ont en crash
    techFilter: null,    // cockpit : technicien du rail "Mon equipe" -> filtre les PC qui lui sont assignes
    sort: { col: 'score', dir: 'desc' },
    // Pagination : itemsPerPage peut valoir 20, 50, 100, ou 0 (tous)
    // Persistance via localStorage pour se souvenir du choix utilisateur
    itemsPerPage: (function() {
        try {
            var saved = localStorage.getItem('pcpulse_itemsPerPage');
            if (saved !== null) return parseInt(saved, 10);
        } catch (e) {}
        return 50;  // defaut
    })(),
    currentPage: 1
};

// ===== BUGCHECK MAPPING =====
// Stop codes BSOD les plus courants pour affichage symbolique
// Source : https://learn.microsoft.com/windows-hardware/drivers/debugger/bug-check-code-reference2
var BUGCHECKS = {
    '0x1':   'APC_INDEX_MISMATCH',
    '0xA':   'IRQL_NOT_LESS_OR_EQUAL',
    '0x1A':  'MEMORY_MANAGEMENT',
    '0x1E':  'KMODE_EXCEPTION_NOT_HANDLED',
    '0x3B':  'SYSTEM_SERVICE_EXCEPTION',
    '0x4E':  'PFN_LIST_CORRUPT',
    '0x50':  'PAGE_FAULT_IN_NONPAGED_AREA',
    '0x7A':  'KERNEL_DATA_INPAGE_ERROR',
    '0x7B':  'INACCESSIBLE_BOOT_DEVICE',
    '0x7E':  'SYSTEM_THREAD_EXCEPTION_NOT_HANDLED',
    '0x7F':  'UNEXPECTED_KERNEL_MODE_TRAP',
    '0x9F':  'DRIVER_POWER_STATE_FAILURE',
    '0xC1':  'SPECIAL_POOL_DETECTED_MEMORY_CORRUPTION',
    '0xC2':  'BAD_POOL_CALLER',
    '0xC4':  'DRIVER_VERIFIER_DETECTED_VIOLATION',
    '0xC5':  'DRIVER_CORRUPTED_EXPOOL',
    '0xD1':  'DRIVER_IRQL_NOT_LESS_OR_EQUAL',
    '0xDE':  'POOL_CORRUPTION_IN_FILE_AREA',
    '0xEF':  'CRITICAL_PROCESS_DIED',
    '0xF4':  'CRITICAL_OBJECT_TERMINATION',
    '0xF7':  'DRIVER_OVERRAN_STACK_BUFFER',
    '0x109': 'CRITICAL_STRUCTURE_CORRUPTION',
    '0x116': 'VIDEO_TDR_FAILURE',
    '0x124': 'WHEA_UNCORRECTABLE_ERROR',
    '0x133': 'DPC_WATCHDOG_VIOLATION',
    '0x139': 'KERNEL_SECURITY_CHECK_FAILURE',
    '0x154': 'UNEXPECTED_STORE_EXCEPTION',
    '0x1E1': 'VIDEO_DXGKRNL_FATAL_ERROR',
    '0xEF':  'CRITICAL_PROCESS_DIED'
};
function bugCheckName(stopCode) {
    if (!stopCode) return '';
    var key = stopCode.toUpperCase().replace(/^0X/, '0x');
    // Normalise : 0x0000007E -> 0x7E
    var num = key.replace(/^0x0+/, '0x');
    if (num === '0x') num = '0x0';
    return BUGCHECKS[num] || '';
}

// ===== UTILS =====
function parseDate(s) { return new Date(s.replace(' ', 'T')); }

// v2.3.0 : affichage utilisateur. Session live si presente, sinon dernier
// utilisateur connu (LastLoggedUser) en muted/italique - equivalent du
// "last seen" Nexthink, reduit a UNE entree. Les valeurs sont deja HTML-safe.
function userDisplay(pc) {
    var cur = pc.CurrentUser || '';
    if (cur && cur !== '(aucune session)') return cur;
    var last = pc.LastLoggedUser || '';
    if (last) {
        return '<span class="user-last" title="Aucune session active - dernier utilisateur connecte">'
             + last + '<span class="user-last-tag">dernier</span></span>';
    }
    return cur || '';   // '(aucune session)' si vraiment rien
}

// v5.8 : un ecran est-il identifiable ? Utilise le flag Identified (Collector
// >= 2.3.0) ; a defaut (ancien JSON) deduit la signature EDID nul (fabricant
// "@@@" ou vide + code produit nul + pas d'annee). Sert a libeller proprement
// les ecrans qui passent par un dock/adaptateur ne relayant pas l'EDID.
function monIsIdentified(mon) {
    if (mon.Identified === false) return false;
    if (mon.Identified === true)  return true;
    var mc = (mon.ManufacturerCode || '').trim();
    var pc = (mon.ProductCode || '').trim();
    var badManu = (mc === '' || /^@+$/.test(mc));
    var badProd = (pc === '' || /^0+$/.test(pc));
    var noYear  = !mon.YearOfManufacture;
    return !(badManu && badProd && noYear);
}

// v1.9 : decode un ID fabricant RAM JEDEC brut cote client (miroir de
// ConvertFrom-JedecManufacturer du Collector). Permet d'afficher un nom
// lisible meme sur les anciens JSON pas encore reguleres par le Collector 2.3.0.
function prettyRamManuf(raw) {
    if (!raw) return '';
    var s = ('' + raw).trim();
    if (!/^[0-9A-Fa-f]+$/.test(s)) return s;   // deja un nom lisible
    if (/^0+$/.test(s))            return '';   // BIOS non renseigne
    if (s.length % 2 !== 0)        return '';   // hex illisible -> on masque
    var map = { 1:'AMD', 11:'Nanya', 44:'Micron', 45:'SK Hynix', 78:'Samsung' };
    var cont = false;
    for (var i = 0; i < s.length; i += 2) {
        var b = parseInt(s.substr(i, 2), 16);
        if (b === 0x7F) { cont = true; continue; }
        if (b === 0x00) continue;
        if (cont) break;
        var code = b & 0x7F;
        if (code === 0) continue;
        if (map[code]) return map[code];
    }
    return '';   // code JEDEC inconnu : mieux vaut rien que du hex brut
}
// v2.5.0 : normalise le type memoire. Gere "DDR4" (deja propre), un code
// SMBIOS nu ("35") ou le repli du collector ("Type=35"). Table SMBIOS 7.18.2.
function cleanRamType(t) {
    if (!t) return '';
    var s = String(t);
    if (s.indexOf('=') === -1 && !/^\d+$/.test(s)) return s;   // deja lisible (DDR4...)
    var m = s.match(/(\d+)/);
    var code = m ? parseInt(m[1], 10) : null;
    var map = { 20:'DDR', 21:'DDR2', 24:'DDR3', 26:'DDR4', 34:'DDR5',
                28:'LPDDR', 29:'LPDDR2', 30:'LPDDR3', 31:'LPDDR4', 35:'LPDDR5' };
    return (code != null && map[code]) ? map[code] : '';
}
// v2.1.2 : formatage lisible d'une duree en secondes (throttling cumule par jour)
function fmtDuration(sec) {
    if (sec === null || sec === undefined || isNaN(sec) || sec <= 0) return '';
    sec = Math.round(sec);
    if (sec < 60) return sec + ' s';
    var m = Math.floor(sec / 60), s = sec % 60;
    if (m < 60) return m + ' min' + (s ? ' ' + s + ' s' : '');
    var h = Math.floor(m / 60); m = m % 60;
    return h + ' h' + (m ? ' ' + m + ' min' : '');
}
function timeAgo(dateStr) {
    var d = parseDate(dateStr);
    var diffMs = generatedAt - d;
    var diffMin = Math.floor(diffMs / 60000);
    var diffH = Math.floor(diffMin / 60);
    var diffJ = Math.floor(diffH / 24);
    if (diffJ > 0) return 'il y a ' + diffJ + 'j';
    if (diffH > 0) return 'il y a ' + diffH + 'h';
    if (diffMin > 0) return 'il y a ' + diffMin + 'min';
    return 'a l instant';
}
function frDate(s) {
    if (!s) return '';
    var m = String(s).match(/(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/);
    if (!m) return s;
    // v2.5.0 : format unique dans tout le dashboard -> "JJ-MM HH:MM". Les donnees
    // sont toujours recentes (<= 30 j) : l'annee est retiree, les secondes ignorees
    // (le regex ne capture que HH:MM). Jour + mois en gras.
    var r = '<b class="dd-strong">' + m[3] + '-' + m[2] + '</b>';
    if (m[4]) r += ' ' + m[4] + ':' + m[5];
    return r;
}

// ============================================================
// v1.6 : VERDICT GLOBAL (4 niveaux)
// Regle : on remonte tous les criteres observes et on prend le
// niveau le plus grave. Chaque critere contribue a la liste de
// raisons affichee en metadata.
// ============================================================
// ============================================================
// v1.7 : DEBRUITAGE DES TOP CRASHERS (Signal vs Bruit)
// ============================================================
// Blacklist HARD : processus ecartes completement, invisibles partout.
// Reserve aux crashers "inactionnables par construction" dont le volume
// pourrit la vue sans apporter aucune info exploitable.
var CRASHER_BLACKLIST_HARD = [
    // Le vrai nom du binaire Microsoft a une typo interne : "searchINbing"
    // (et non "searchBing" comme on pourrait le croire). On garde les 2 variantes
    // par securite au cas ou Microsoft corrigerait la typo un jour.
    'microsoftsearchinbing.exe',  // typo reelle du binaire (cas terrain)
    'microsoftsearchbing.exe'     // orthographe logique (si jamais corrige)
];

// Blacklist SOFT : processus toujours affiches, mais forces en section
// "Bruit ambient" meme si leur score les aurait remontes plus haut.
// Garde un oeil dessus sans les laisser polluer le top.
var CRASHER_BLACKLIST_SOFT = [
    'dellosd.exe',                       // Dell On-Screen Display
    'shellexperiencehost.exe',           // UI shell Windows, redemarre seul
    'gamebar.exe',                       // Xbox Game Bar
    'asussystemanalysis.exe',            // utilitaire ASUS
    'asusverifyjwt.exe',                 // utilitaire ASUS
    'dell.techhub.diagnostics.subagent'  // Dell Tech Hub
];

function isBlacklistedHard(name) {
    if (!name) return false;
    var n = String(name).toLowerCase();
    for (var i = 0; i < CRASHER_BLACKLIST_HARD.length; i++) {
        if (n.indexOf(CRASHER_BLACKLIST_HARD[i]) >= 0) return true;
    }
    return false;
}
function isBlacklistedSoft(name) {
    if (!name) return false;
    var n = String(name).toLowerCase();
    for (var i = 0; i < CRASHER_BLACKLIST_SOFT.length; i++) {
        if (n.indexOf(CRASHER_BLACKLIST_SOFT[i]) >= 0) return true;
    }
    return false;
}
// v2.1.12 : appli suivie (match EXACT, insensible a la casse). Les noms sont
// deja normalises par le Collector (Get-CrasherKey) et htmlsafe cote embed ;
// pour les libelles simples de la liste, la casse est la seule variation possible.
function isPriorityApp(name) {
    if (!name) return false;
    return priorityApps.indexOf(String(name).toLowerCase()) !== -1;
}

// v2.5.0 : filtre de bruit PARTAGE (panneau global + drill-down par PC). Le
// Collector capture parfois des fragments de message de log (traces Dell DDPM,
// exceptions .NET, namespaces, SID, {}, ...) comme si c'etaient des applis en
// echec. Un vrai process a un .exe, ou est un token propre sans point / espace /
// ponctuation. Une appli suivie (priorityApps) n'est jamais consideree bruit.
function isRealApp(n) {
    n = String(n);
    if (/[\s\[\]{}='":]/.test(n)) return false;   // espace / ponctuation => texte de log
    if (/\.exe/i.test(n)) return true;             // tout ce qui contient .exe est un process
    return n.indexOf('.') === -1;                  // token propre sans point OK ; namespace .NET rejete
}
// v2.5.0 : un "crasher" dont le nom est exactement la session ouverte est un
// artefact d'extraction (le nom d'utilisateur happe dans un message). On le
// masque en s'appuyant sur les identites connues du poste.
function isUserArtifact(name, pc) {
    if (!name || !pc) return false;
    var n = String(name).toLowerCase().trim();
    if (!n) return false;
    var users = [pc.CurrentUser, pc.LastLoggedUser];
    for (var i = 0; i < users.length; i++) {
        var u = String(users[i] || '').toLowerCase().trim();
        if (u && u !== '(aucune session)' && u === n) return true;
    }
    return false;
}

// Scoring : penalise la dispersion entre PC.
// score = (total/pcCount) / sqrt(pcCount) = moyenne_par_PC / sqrt(PC_impactes)
// Un processus concentre (8 crashs sur 1 PC) sort plus haut qu'un
// processus reparti (81 crashs sur 10 PC), qui est typiquement du bruit.
function computeCrasherScore(total, pcCount) {
    if (!pcCount || pcCount <= 0) return 0;
    var avgPerPc = total / pcCount;
    return avgPerPc / Math.sqrt(pcCount);
}

// Classification en 3 niveaux :
//   'local'   (score >= 3)  : concentre sur peu de PC, actionnable
//   'spread'  (2 <= score < 3) : reparti, possible bug app
//   'noise'   (score < 2)   : bruit ambient, pas actionnable
// Les entrees en blacklist SOFT sont forcees en 'noise' peu importe leur score.
function classifyCrasher(name, total, pcCount, type) {
    // v2.1.12 : une appli suivie remonte TOUJOURS dans sa section dediee, avant
    // tout autre classement (nature ou dispersion). On conserve son score pour
    // trier la section, mais il ne decide plus de sa place.
    if (isPriorityApp(name)) return { level: 'priority', score: computeCrasherScore(total, pcCount), forced: false };
    // v2.1.4 : la NATURE prime sur la dispersion. Un echec applicatif recurrent
    // (Type 'app_failure', ex Bing Wallpaper) n'est pas un crash : il part dans
    // sa categorie 'appfail' quelle que soit sa concentration, au lieu de
    // remonter en 'local' a tort. Type absent (JSON legacy) = traite en crash.
    if (type === 'app_failure') return { level: 'appfail', score: total, forced: false };
    var score = computeCrasherScore(total, pcCount);
    if (isBlacklistedSoft(name)) return { level: 'noise', score: score, forced: true };
    if (score >= 3)              return { level: 'local',  score: score, forced: false };
    if (score >= 2)              return { level: 'spread', score: score, forced: false };
    return                              { level: 'noise',  score: score, forced: false };
}

// v2.3.1 (#9) : traduction des crashers pour le tech de proximite - JAMAIS
// "hang" ni code brut a l'ecran. Origine deduite du module fautif, code
// exception en clair ; le brut (module + code) reste en tooltip pour l'admin.
function crasherOrigin(module) {
    if (!module) return '';
    var m = String(module).toLowerCase();
    if (m.indexOf('ntdll') !== -1 || m.indexOf('kernelbase') !== -1) return 'interne (m&eacute;moire/syst&egrave;me)';
    if (m.indexOf('clr') !== -1 || m.indexOf('mscor') !== -1)         return 'erreur .NET';
    if (/\.exe$/.test(m))                                             return 'vient de l\'application';
    if (/\.(dll|sys|ocx)$/.test(m))                                   return 'conflit de composant';
    return '';
}
function exceptionLabel(code) {
    if (!code) return '';
    var map = {
        '0xc0000005': 'acc&egrave;s m&eacute;moire invalide',
        '0xc0000374': 'corruption m&eacute;moire (tas)',
        '0xe0434352': 'exception .NET non g&eacute;r&eacute;e',
        '0xc0000409': 'd&eacute;passement de pile',
        '0xc00000fd': 'd&eacute;bordement de pile',
        '0x80000003': 'point d\'arr&ecirc;t (debug)',
        '0xc0000094': 'division par z&eacute;ro'
    };
    return map[String(code).toLowerCase()] || '';
}
// "X plantes / Y figes" a partir des compteurs (fallback : rien si absents/0)
function crasherNature(tc) {
    var parts = [];
    var err = +tc.ErrorCount || 0, hang = +tc.HangCount || 0;
    if (err > 0)  parts.push(err + ' plant&eacute;'  + (err > 1 ? 's' : ''));
    if (hang > 0) parts.push(hang + ' fig&eacute;'    + (hang > 1 ? 's' : ''));
    return parts.join(' / ');
}
// Ligne de detail lisible d'un crasher (nature + code clair + origine),
// avec le brut module/code en tooltip. Vide si rien de traduisible.
function crasherDetailHtml(tc) {
    var bits = [];
    var nature = crasherNature(tc);
    var code   = exceptionLabel(tc.ExceptionCode);
    var orig   = crasherOrigin(tc.FaultModule);
    if (nature) bits.push(nature);
    if (code)   bits.push(code);
    if (orig)   bits.push(orig);
    if (!bits.length) return '';
    var raw = [];
    if (tc.FaultModule)   raw.push('module ' + tc.FaultModule);
    if (tc.ExceptionCode) raw.push('code ' + tc.ExceptionCode);
    var titleAttr = raw.length ? ' title="' + raw.join(' &middot; ') + '"' : '';
    return '<div class="crasher-detail"' + titleAttr + '>' + bits.join(' &middot; ') + '</div>';
}

function computeVerdict(p) {
    // Donnees sur lesquelles on raisonne
    var crashCount    = p.crashCount        || 0;
    var bsodCount     = p.bsodCount         || 0;
    var hardCrash     = (p.pc && p.pc.TotalHardCrash) ? p.pc.TotalHardCrash : 0;  // v1.6
    var wheaFatal     = (p.wheaFatal    || []).length;
    var wheaCorrUniq  = (p.wheaCorrected || []).length;
    var wheaCorrTot   = p.wheaCorrectedTotal || 0;
    var cpuCat        = (p.pc && p.pc.CPUAgeCategory) || 'Inconnu';
    var battPct       = (p.battery && p.battery.HasBattery === true && typeof p.battery.HealthPercent === 'number') ? p.battery.HealthPercent : null;
    var smartWearMax  = 0;
    if (p.pc && Array.isArray(p.pc.DiskHealth)) {
        p.pc.DiskHealth.forEach(function(d) {
            if (typeof d.WearPct === 'number' && d.WearPct > smartWearMax) smartWearMax = d.WearPct;
        });
    }
    // Burst I/O massif (>= 50 events en un cluster)
    var hasBurstMassive = (p.warnings || []).some(function(w) { return w.IsBurst === true; });
    var hasBurstAny     = (p.warnings || []).some(function(w) { return typeof w.Count === 'number' && w.Count > 1; });

    var reasons = [];
    var level   = 'sain';

    // --- NIVEAU CRITIQUE ---
    if (wheaFatal > 0) {
        reasons.push(wheaFatal + ' erreur(s) WHEA fatale(s)');
        level = 'critical';
    }
    if (hardCrash >= 5) {
        reasons.push(hardCrash + ' hard crashs/30j');
        level = 'critical';
    }
    if (battPct !== null && battPct < 50) {
        reasons.push('batterie ' + battPct + '%');
        level = 'critical';
    }
    if (smartWearMax > 80) {
        reasons.push('SSD wear ' + smartWearMax + '%');
        level = 'critical';
    }
    // v2.5.3 : en mode renouvellement, l'axe age ne pollue plus le VERDICT sante
    // (le renouvellement est un axe de planning, montre via badge/filtre/KPI).
    // Aucun libelle "ancien/vieillissant" affiche. En fallback, comportement historique.
    if (!RENEWAL_MODE && cpuCat === 'Ancien' && crashCount >= 1) {
        reasons.push('CPU ancien + crashs');
        level = 'critical';
    }

    // --- NIVEAU INCIDENT PROBABLE ---
    if (level !== 'critical') {
        if (hasBurstMassive) {
            reasons.push('burst I/O massif (>=50 events)');
            level = 'incident';
        }
        if (wheaCorrUniq > 20 || wheaCorrTot > 50) {
            reasons.push(wheaCorrUniq + ' signatures WHEA corrigees');
            level = 'incident';
        }
        if (bsodCount >= 1) {
            reasons.push(bsodCount + ' BSOD/30j');
            level = 'incident';
        }
        if (hardCrash >= 2 && hardCrash < 5) {
            reasons.push(hardCrash + ' hard crashs/30j');
            level = 'incident';
        }
    }

    // --- NIVEAU A SURVEILLER ---
    if (level !== 'critical' && level !== 'incident') {
        if (!RENEWAL_MODE && cpuCat === 'Vieillissant') {
            reasons.push('CPU vieillissant');
            level = 'watch';
        }
        if (battPct !== null && battPct >= 50 && battPct < 70) {
            reasons.push('batterie ' + battPct + '%');
            level = 'watch';
        }
        if (smartWearMax >= 50 && smartWearMax <= 80) {
            reasons.push('SSD wear ' + smartWearMax + '%');
            level = 'watch';
        }
        if (hasBurstAny && !hasBurstMassive) {
            reasons.push('burst I/O isole');
            level = 'watch';
        }
    }

    // Libelles pour affichage
    var labels = {
        sain:     { cls: 'sain',     icon: '&#129001;', label: 'Sain' },              // green circle
        watch:    { cls: 'watch',    icon: '&#129000;', label: 'A surveiller' },      // yellow circle
        incident: { cls: 'incident', icon: '&#128992;', label: 'Incident probable' }, // orange circle
        critical: { cls: 'critical', icon: '&#128308;', label: 'Critique' }           // red circle
    };

    var info = labels[level];
    return {
        level:   level,
        cls:     info.cls,
        icon:    info.icon,
        label:   info.label,
        reasons: reasons
    };
}

// ============================================================
// v1.6 : DETECTION DE PATTERNS TEMPORELS (fenetre 10 min)
// 5 patterns qui croisent differentes sources pour donner un
// diagnostic qu'une seule source ne revele pas.
// ============================================================
function detectCorrelations(p) {
    var CORR_WINDOW_MS = 10 * 60 * 1000;  // 10 min
    var findings = [];

    function tsOf(s) { return parseDate(s).getTime(); }

    // -- Pattern 1 : Burst I/O (>=50 events) -> Hard crash dans les 10 min
    var bursts = (p.warnings || []).filter(function(w) { return w.IsBurst === true; });
    var crashTimestamps = (p.crashes || []).map(function(c) { return { ts: tsOf(c.Timestamp), cause: c.CrashCause || '', type: c.Type }; });
    bursts.forEach(function(b) {
        var bTs = tsOf(b.Timestamp);
        var match = crashTimestamps.filter(function(c) {
            return c.ts >= bTs && (c.ts - bTs) <= CORR_WINDOW_MS;
        });
        if (match.length > 0) {
            findings.push({
                severity: 'crit',
                icon: '&#128190;',
                title: 'Probable panne disque',
                detail: 'Burst I/O massif (' + b.Count + ' events) le ' + frDate(b.Timestamp) + ' suivi d un crash materiel dans les 10 min'
            });
        }
    });

    // -- Pattern 2 : WHEA Corrected PCIe -> crash systeme dans les 10 min
    var wheaPCIe = (p.wheaCorrected || []).filter(function(h) {
        return (h.ErrorSource || '').toUpperCase().indexOf('PCI') >= 0;
    });
    wheaPCIe.forEach(function(h) {
        var hTs = tsOf(h.LastSeen);
        var match = crashTimestamps.filter(function(c) {
            return Math.abs(c.ts - hTs) <= CORR_WINDOW_MS;
        });
        if (match.length > 0) {
            findings.push({
                severity: 'warn',
                icon: '&#128268;',
                title: 'Slot PCIe suspect',
                detail: 'WHEA PCIe (' + h.Count + ' occurrences) corrélée avec un crash dans les 10 min'
            });
        }
    });

    // -- Pattern 3 : Thermal event + BootTime > 2x moyenne historique
    var hasThermal = (p.thermal || []).length > 0;
    var avgBoot = (p.bootPerf && p.bootPerf.Stats && p.bootPerf.Stats.AvgBootTimeMs) ? p.bootPerf.Stats.AvgBootTimeMs : 0;
    var maxBoot = (p.bootPerf && p.bootPerf.Stats && p.bootPerf.Stats.MaxBootTimeMs) ? p.bootPerf.Stats.MaxBootTimeMs : 0;
    if (hasThermal && avgBoot > 0 && maxBoot > 2 * avgBoot) {
        findings.push({
            severity: 'warn',
            icon: '&#127777;',
            title: 'Refroidissement degrade',
            detail: 'Thermal event + pic de boot a ' + Math.round(maxBoot/1000) + 's (moyenne ' + Math.round(avgBoot/1000) + 's) : ventilo ou pate thermique a verifier'
        });
    }

    // -- Pattern 4 : BSOD >=2 avec meme BugCheck sur 7j
    var SEVEN_DAYS_MS = 7 * 24 * 3600 * 1000;
    var bsodByCode = {};
    (p.crashes || []).filter(function(c) { return c.Type === 'BSOD' && c.Detail; }).forEach(function(c) {
        var code = c.Detail;
        var ts = tsOf(c.Timestamp);
        if (!bsodByCode[code]) bsodByCode[code] = [];
        bsodByCode[code].push(ts);
    });
    Object.keys(bsodByCode).forEach(function(code) {
        var arr = bsodByCode[code].sort(function(a,b){ return b-a; });
        if (arr.length >= 2 && (arr[0] - arr[arr.length-1]) <= SEVEN_DAYS_MS) {
            var bugName = bugCheckName(code);
            findings.push({
                severity: 'warn',
                icon: '&#128165;',
                title: 'Crash recurrent',
                detail: arr.length + ' BSOD ' + code + (bugName ? ' (' + bugName + ')' : '') + ' en 7 jours'
            });
        }
    });

    // -- Pattern 5 : Hard crash >=2 en 24h
    var ONE_DAY_MS = 24 * 3600 * 1000;
    var hardCrashes = (p.crashes || []).filter(function(c) {
        return c.Type === 'Hard reset' && c.CrashCause && c.CrashCause !== 'UserForcedReset';
    }).map(function(c) { return tsOf(c.Timestamp); }).sort(function(a,b){ return b-a; });
    for (var i = 0; i < hardCrashes.length - 1; i++) {
        if ((hardCrashes[i] - hardCrashes[i+1]) <= ONE_DAY_MS) {
            findings.push({
                severity: 'warn',
                icon: '&#9889;',
                title: 'Instabilite marquee',
                detail: '>=2 hard crashs en moins de 24h : surveiller alim / thermique / drivers'
            });
            break;  // un seul match suffit
        }
    }

    return findings;
}
// Format humain de l'uptime a partir d'une valeur en jours (float).
// Regles :
//   < 1h       -> "Xmin"
//   1h a 24h   -> "Xh"
//   1j a 7j    -> "Xj Yh"   (ex: "1j 8h")
//   > 7j       -> "Xj"      (jours entiers, sans les heures)
function formatUptime(uptimeDays) {
    if (uptimeDays === null || uptimeDays === undefined) return 'N/A';
    var totalMinutes = Math.floor(uptimeDays * 24 * 60);
    if (totalMinutes < 60) return totalMinutes + 'min';
    var totalHours = Math.floor(totalMinutes / 60);
    if (totalHours < 24) return totalHours + 'h';
    var days = Math.floor(totalHours / 24);
    var remainingHours = totalHours - (days * 24);
    if (days < 7) return days + 'j ' + remainingHours + 'h';
    return days + 'j';
}
function freshnessClass(hoursAgo) {
    if (hoursAgo <= 2) return 'freshness-ok';
    if (hoursAgo <= 24) return 'freshness-warning';
    return 'freshness-danger';
}
function hoursSince(dateStr) {
    return (generatedAt - parseDate(dateStr)) / 3600000;
}

// ===== CALCUL SCORE SANTE =====
// Score compose a partir des poids de config. Plus c'est haut, plus c'est grave.
// v5.2 : on compte separement WHEA_Fatal, Thermal et GPU_TDR (plus de double
// comptage). Les WHEA corrected ne sont jamais comptees (c'est de la telemetrie).
function computeScore(p) {
    var s = 0;
    s += p.bsodCount          * scoreWeights.BSOD;
    s += p.wheaFatal.length   * scoreWeights.WHEA;
    s += p.crashCount         * scoreWeights.Crash;
    s += p.thermal.length     * scoreWeights.Thermal;
    s += p.gpuTDR.length      * scoreWeights.GPU_TDR;
    s += p.diskAlertCount     * scoreWeights.DiskAlert;
    s += p.bootLongCount      * scoreWeights.BootLong;
    if (p.pc.IsOffline)       s += scoreWeights.Offline;
    // v5.3 / v5.4 : nouvelles penalites (defensif si weights absents du config ancien)
    if (p.edrAlert)      s += (scoreWeights.EDRDown || 5);
    if (p.batteryAlert)       s += (scoreWeights.Battery      || 1);
    if (p.bootPerfAlert)      s += (scoreWeights.BootPerfSlow || 1);
    if (p.diskSmartAlert)     s += (scoreWeights.DiskHealth   || 3);
    return s;
}
function scoreClass(s) {
    if (s === 0) return 'score-ok';
    if (s <= 5)  return 'score-warn';
    if (s <= 15) return 'score-danger';
    return 'score-critic';
}

// ===== INIT SITES DROPDOWN =====
function initSiteDropdown() {
    if (!showSite) {
        document.getElementById('siteFilter').style.display = 'none';
        // masquer aussi le label du site
        var labels = document.querySelectorAll('.toolbar label');
        for (var i = 0; i < labels.length; i++) {
            if (labels[i].textContent.indexOf('Site') !== -1) labels[i].style.display = 'none';
        }
        return;
    }
    var sites = {};
    pcData.forEach(function(pc) { if (pc.Site) sites[pc.Site] = true; });
    var sorted = Object.keys(sites).sort();
    var sel = document.getElementById('siteFilter');
    sorted.forEach(function(s) {
        var opt = document.createElement('option');
        opt.value = s;
        opt.textContent = s;
        sel.appendChild(opt);
    });
}

// ===== INIT OS DROPDOWN =====
function initOsDropdown() {
    var oses = {};
    pcData.forEach(function(pc) { if (pc.OSProduct) oses[pc.OSProduct] = true; });
    var sorted = Object.keys(oses).sort();
    var sel = document.getElementById('osFilter');
    // Masquer le filtre s'il n'y a pas de diversite (0 ou 1 OS present) : inutile.
    if (sorted.length < 2) {
        if (sel) sel.style.display = 'none';
        var labels = document.querySelectorAll('.toolbar label');
        for (var i = 0; i < labels.length; i++) {
            if (labels[i].textContent.indexOf('OS') !== -1) labels[i].style.display = 'none';
        }
        return;
    }
    sorted.forEach(function(s) {
        var opt = document.createElement('option');
        opt.value = s;
        opt.textContent = s;
        sel.appendChild(opt);
    });
}

// ===== INIT MODELE DROPDOWN (v2.4.7) =====
// Liste les modeles machine presents dans le parc (Machine.Model, Collector >= 2.4.7).
function initModelDropdown() {
    var sel = document.getElementById('modelFilter');
    if (!sel) return;
    var models = {};
    pcData.forEach(function(pc) { if (pc.Model) models[pc.Model] = true; });
    var sorted = Object.keys(models).sort();
    // Masquer le filtre s'il n'y a pas de diversite (0 ou 1 modele connu) : inutile.
    if (sorted.length < 2) {
        sel.style.display = 'none';
        var labels = document.querySelectorAll('.toolbar label');
        for (var i = 0; i < labels.length; i++) {
            if (labels[i].textContent.indexOf('Mod') !== -1) labels[i].style.display = 'none';
        }
        return;
    }
    sorted.forEach(function(s) {
        var opt = document.createElement('option');
        opt.value = s;
        opt.textContent = s;
        sel.appendChild(opt);
    });
}

// ===== INIT VPN VERSION DROPDOWN (v2.5.2) =====
// Liste les versions de client VPN presentes (VpnClient.Version, Collector >= 2.5.1).
// But : traquer les vieilles versions (ex. FortiClient 6.x) sur un parc heterogene.
// Tri par version DECROISSANTE : les 7.x en haut, les 6.x regroupees en bas.
function initVpnDropdown() {
    var sel = document.getElementById('vpnFilter');
    if (!sel) return;
    var vers = {};
    pcData.forEach(function(pc) {
        if (pc.VpnClient && pc.VpnClient.Present && pc.VpnClient.Version) vers[pc.VpnClient.Version] = true;
    });
    var sorted = Object.keys(vers).sort(function(a, b) {
        var pa = a.split('.'), pb = b.split('.');
        for (var i = 0; i < Math.max(pa.length, pb.length); i++) {
            var na = parseInt(pa[i], 10) || 0, nb = parseInt(pb[i], 10) || 0;
            if (na !== nb) return nb - na;   // decroissant : 7.x avant 6.x
        }
        return 0;
    });
    // Masquer le filtre s'il n'y a pas de diversite (0 ou 1 version connue) : inutile.
    if (sorted.length < 2) {
        sel.style.display = 'none';
        var labels = document.querySelectorAll('.toolbar label');
        for (var i = 0; i < labels.length; i++) {
            if (labels[i].textContent.indexOf('VPN') !== -1) labels[i].style.display = 'none';
        }
        return;
    }
    sorted.forEach(function(s) {
        var opt = document.createElement('option');
        opt.value = s;
        opt.textContent = s;
        sel.appendChild(opt);
    });
}

// ===== THEME =====
function toggleTheme() {
    var html = document.documentElement;
    var now = html.getAttribute('data-theme');
    var next = (now === 'dark') ? 'light' : 'dark';
    html.setAttribute('data-theme', next);
    document.getElementById('themeToggle').innerHTML = next === 'dark' ? '&#9788;' : '&#9789;';
    try { localStorage.setItem('pcmon-theme', next); } catch(e) {}
}
(function() {
    try {
        var saved = localStorage.getItem('pcmon-theme');
        if (saved) {
            document.documentElement.setAttribute('data-theme', saved);
            document.getElementById('themeToggle').innerHTML = saved === 'dark' ? '&#9788;' : '&#9789;';
        }
    } catch(e) {}
})();

// ===== FILTRES =====
// v1.5 : wrapper appele par search/site/cpu qui reset la page avant render
// (evite de se retrouver sur une page qui n'existe plus apres filtrage)
function filterChanged() {
    state.currentPage = 1;
    render();
}

function setDays(days, btn) {
    state.days = days;
    state.daysBtn = btn;
    state.currentPage = 1;  // v1.5 : retour page 1 sur changement de filtre
    document.querySelectorAll('.range-btn').forEach(function(b) { b.classList.remove('active'); });
    if (btn) btn.classList.add('active');
    render();
}
// v2.5.2 : filtre chassis par icone (toggle). Un seul type actif a la fois ;
// re-cliquer l'icone active enleve le filtre. ChassisInfo null (Collector < 5.6)
// -> jamais retenu quand un type est selectionne.
function chassisMatch(pc, kind) {
    var ci = pc.ChassisInfo;
    if (!ci) return false;
    return (kind === 'laptop' && ci.IsLaptop) ||
           (kind === 'desktop' && ci.IsDesktop) ||
           (kind === 'aio' && ci.IsAIO);
}
function toggleChassisFilter(kind) {
    state.chassisFilter = (state.chassisFilter === kind) ? null : kind;
    state.currentPage = 1;
    ['laptop', 'desktop', 'aio'].forEach(function(k) {
        var b = document.getElementById('chassis' + k.charAt(0).toUpperCase() + k.slice(1));
        if (b) b.classList.toggle('active', state.chassisFilter === k);
    });
    render();
}
function toggleMaskHealthy() {
    state.maskHealthy = !state.maskHealthy;
    state.currentPage = 1;  // v1.5
    document.getElementById('maskHealthyBtn').classList.toggle('active', state.maskHealthy);
    render();
}

// v5.5 : bascule entre la vue compacte et la vue detaillee du tableau.
// Persiste le choix dans localStorage pour que l'utilisateur ne subisse
// pas un reset a chaque regeneration du dashboard.
function toggleAdvancedCols() {
    var on = document.body.classList.toggle('advanced-cols');
    document.getElementById('advancedColsBtn').classList.toggle('active', on);
    try { localStorage.setItem('pcmon-advcols', on ? '1' : '0'); } catch(e) {}
}

// v5.5 : changement d'onglet dans un drill-down PC. Le state est stocke
// en data-attribute sur le conteneur pour etre isole par ligne.
function selectDetailTab(idx, tab) {
    var root = document.getElementById('detail-' + idx);
    if (!root) return;
    var tabs = root.querySelectorAll('.detail-tab');
    var panels = root.querySelectorAll('.detail-tab-panel');
    for (var i = 0; i < tabs.length; i++) {
        tabs[i].classList.toggle('active', tabs[i].getAttribute('data-tab') === tab);
    }
    for (var j = 0; j < panels.length; j++) {
        panels[j].classList.toggle('active', panels[j].getAttribute('data-panel') === tab);
    }
}
function toggleKpiFilter(key) {
    state.kpiFilter = (state.kpiFilter === key) ? null : key;
    state.currentPage = 1;  // v1.5
    render();
}
function clearKpiFilter() {
    state.kpiFilter = null;
    state.currentPage = 1;  // v1.5
    render();
}
function clearSearch() {
    var i = document.getElementById('searchInput');
    if (i) { i.value = ''; i.focus(); }
    filterChanged();
}
// v2.1.12 : filtre par appli (clic sur un crasher dans le panneau Top Crashers).
function toggleAppFilter(name) {
    if (!name) return;
    state.appFilter = (state.appFilter === name) ? null : name;
    state.currentPage = 1;
    render();
}
function clearAppFilter() {
    state.appFilter = null;
    state.currentPage = 1;
    render();
}
// v2.1.14 : reinitialise TOUS les filtres d'un coup - les chips (KPI, appli, masquer sains)
// ET les controles qui n'ont pas de chip (menus site/CPU, champ recherche). Retour a l'etat
// d'ouverture : tout le parc visible.
function clearAllFilters() {
    state.kpiFilter   = null;
    state.appFilter   = null;
    state.techFilter  = null;   // cockpit
    state.maskHealthy = false;
    state.siteFilter  = '';
    state.cpuFilter   = '';
    state.osFilter    = '';
    state.modelFilter = '';
    state.vpnFilter   = '';
    state.chassisFilter = null;
    ['laptop', 'desktop', 'aio'].forEach(function(k) {
        var cb = document.getElementById('chassis' + k.charAt(0).toUpperCase() + k.slice(1));
        if (cb) cb.classList.remove('active');
    });
    var mb = document.getElementById('maskHealthyBtn'); if (mb) mb.classList.remove('active');
    var sf = document.getElementById('siteFilter');     if (sf) sf.value = '';
    var cf = document.getElementById('cpuFilter');      if (cf) cf.value = '';
    var of = document.getElementById('osFilter');        if (of) of.value = '';
    var mf = document.getElementById('modelFilter');    if (mf) mf.value = '';
    var vf = document.getElementById('vpnFilter');      if (vf) vf.value = '';
    var si = document.getElementById('searchInput');    if (si) si.value = '';
    state.currentPage = 1;
    render();
}

// ============================================================
// COCKPIT UI : helpers du rail + file d'actions.
// Ne fait que piloter les filtres EXISTANTS (toggleKpiFilter,
// maskHealthy, offline, clearAllFilters) + un filtre technicien.
// ============================================================
function ckRing(pct, color) {
    var r = 20, c = 2 * Math.PI * r, off = c * (1 - pct / 100);
    return '<svg class="ck-ring" width="50" height="50" viewBox="0 0 50 50" aria-hidden="true">' +
      '<circle cx="25" cy="25" r="' + r + '" fill="none" stroke="var(--border)" stroke-width="5"></circle>' +
      '<circle cx="25" cy="25" r="' + r + '" fill="none" stroke="' + color + '" stroke-width="5" stroke-linecap="round" stroke-dasharray="' + c + '" stroke-dashoffset="' + off + '" transform="rotate(-90 25 25)"></circle>' +
      '<text x="25" y="29" text-anchor="middle" font-size="12" font-weight="700" fill="var(--text-main)">' + pct + '</text></svg>';
}
function ckJsAttr(s) { return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'"); }
function ckSetMask(v) {
    state.maskHealthy = v;
    var b = document.getElementById('maskHealthyBtn');
    if (b) b.classList.toggle('active', v);
}
function ckTech(name) {
    if (!name) return;
    state.appFilter = null;
    state.kpiFilter = null;
    ckSetMask(false);
    state.techFilter = (state.techFilter === name) ? null : name;
    state.currentPage = 1;
    render();
}
function ckClearTech() {
    state.techFilter = null;
    state.currentPage = 1;
    render();
}
// File d'actions (droite) : cartes derivees des agregats deja calcules.
// Chaque carte applique un filtre KPI existant via toggleKpiFilter.
// Le bloc "Mon equipe" (techniciens du registre de decommission) est
// desormais rendu ici, en bas de la file, a la place de l'ancien rail.
function renderActionFile(agg, pcs) {
    var box = document.getElementById('ckActs');
    if (!box) return;
    var cards = [
        { k: 'edrDown',   c: 'var(--red)',    n: agg.totalEdrDown,     u: 'postes',  t: 'Agent EDR arr&ecirc;t&eacute; ou absent', why: 'Trou de s&eacute;curit&eacute; direct : relancer le service ou repousser le paquet.' },
        // v2.5.1 : "Silencieux / hors ligne" RETIRE de la file d'actions - un poste
        // eteint n'est pas une action a traiter (parc a moitie eteint l'ete). Le compte
        // reste visible dans la tuile "En ligne 24h", et le filtre 'offline' existe
        // toujours (tuile cliquable + matchKpiFilter) : on retire l'incitation a agir,
        // pas la visibilite.
        { k: 'osEol',     c: 'var(--orange)', n: agg.totalOsEol,       u: 'postes',  t: 'Windows 10 fin de support', why: 'A basculer en 11 ou a sortir du parc. Point de conformit&eacute;.' },
        { k: 'bsod',      c: 'var(--orange)', n: agg.totalBSOD,        u: '&eacute;v&eacute;n.', t: '&Eacute;crans bleus (BSOD)', why: 'Un utilisateur perd son travail. Prioriser les r&eacute;cidivistes.' },
        { k: 'diskAlert', c: 'var(--orange)', n: agg.totalDiskAlerts,  u: 'alertes', t: 'Disques satur&eacute;s', why: 'Cause n&deg;1 de lenteur et de blocage des mises a jour.' },
        { k: 'smart',     c: 'var(--pink)',   n: agg.totalSmartAlert,  u: 'postes',  t: 'Disques SMART en alerte', why: 'Signes de panne mat&eacute;rielle a venir : sauvegarder puis remplacer.' },
        { k: 'battery',   c: 'var(--yellow)', n: agg.totalBatteryWorn, u: 'postes',  t: 'Batteries us&eacute;es (&lt;60%)', why: 'Commande group&eacute;e plut&ocirc;t qu\'au coup par coup.' },
        { k: 'uptime30',  c: 'var(--cyan)',   n: agg.totalUptime30,    u: 'postes',  t: 'Jamais &eacute;teints (&gt;30 j)', why: 'Mises a jour non appliqu&eacute;es + consommation inutile.' },
        { k: 'decomLate', c: 'var(--red)',    n: agg.totalDecomLate,   u: 'postes',  t: 'D&eacute;com. en retard', why: '&Eacute;ch&eacute;ance de sortie d&eacute;pass&eacute;e : lib&egrave;re licences et budget.' },
        { k: 'decomTodo', c: 'var(--purple)', n: agg.totalDecomTodo,   u: 'postes',  t: '&Agrave; d&eacute;commissionner', why: 'Sortie d\'inventaire planifi&eacute;e : lib&egrave;re licences et support.' }
    ];
    var html = '';
    cards.forEach(function(a) {
        if (!a.n) return;
        var on = (state.kpiFilter === a.k) ? ' on' : '';
        html += '<div class="ck-act' + on + '" style="--ck-c:' + a.c + '" onclick="toggleKpiFilter(\'' + a.k + '\')">' +
                '<div class="ck-act-num">' + a.n + '<small>' + a.u + '</small></div>' +
                '<div class="ck-act-txt"><b>' + a.t + '</b><span>' + a.why + '</span></div>' +
                '</div>';
    });
    if (!html) html = '<div class="ck-act-empty">Rien d\'urgent &mdash; le parc est calme.</div>';

    // ===== Sous-section "Mon equipe" =====
    // Techniciens issus du registre de decommission, comptes sur la vue
    // courante (les PC hors registre n'y figurent pas). Clic -> filtre techFilter.
    var teamCount = {};
    (pcs || []).forEach(function(p) { var t = ckTechByPc[p.pc.PC]; if (t) teamCount[t] = (teamCount[t] || 0) + 1; });
    var names = Object.keys(teamCount).sort();
    if (names.length > 0) {
        html += '<div class="ck-team">' +
                '<div class="ck-team-head">Mon &eacute;quipe <small>postes assign&eacute;s</small></div>';
        names.forEach(function(n) {
            var on = (state.techFilter === n) ? ' on' : '';
            html += '<div class="ck-team-row' + on + '" onclick="ckTech(\'' + ckJsAttr(n) + '\')">' +
                    '<span class="ck-team-name">' + n + '</span>' +
                    '<span class="ck-team-n">' + teamCount[n] + '</span>' +
                    '</div>';
        });
        html += '</div>';
    }
    box.innerHTML = html;
}

// ===== ENRICHISSEMENT DES PC =====
// Calcule pour chaque PC les compteurs filtres par la periode + le score
function enrichPC(pc, cutoff) {
    var crashes     = (pc.Crashes || []).filter(function(c) { return parseDate(c.Timestamp) >= cutoff; });
    // v5.7 : on separe la LISTE filtree par periode (pour les compteurs et bootsLongs)
    // de la DERNIERE info boot connue (toujours affichee, meme si hors periode).
    // Avant : si aucun boot dans la periode 24h -> colonne "Boot" = N/A,
    // alors que le JSON contenait bien un dernier boot connu plus ancien.
    var allBoots    = pc.Boots || [];
    var boots       = allBoots.filter(function(b) { return parseDate(b.DateBoot) >= cutoff; });
    // v2.5.0 : redemarrages MAJ Windows (Event 1074) filtres sur la periode, pour
    // corerler les demarrages a leur cause dans l'onglet Demarrage.
    var reboots     = (pc.Reboots || []).filter(function(r) { return parseDate(r.Timestamp) >= cutoff; });
    var bsods       = (pc.BSODs   || []).filter(function(b) { return parseDate(b.Date) >= cutoff; });
    var warnings    = (pc.ResourceWarnings || []).filter(function(w) { return parseDate(w.Timestamp) >= cutoff; });
    var bootsLongs  = boots.filter(function(b) { return b.EstBootLong; });
    // Dernier boot connu : on prend dans allBoots, peu importe la periode
    var dernierBoot = allBoots.length > 0 ? allBoots[allBoots.length - 1] : null;
    var hw          = pc.HardwareHealth || { WHEA_Fatal: [], WHEA_Corrected: [], GPU_TDR: [], Thermal: [] };

    // WHEA_Fatal : filtre par periode
    var wheaFatal = (hw.WHEA_Fatal || []).filter(function(h) { return parseDate(h.Timestamp) >= cutoff; });
    // WHEA_Corrected : deja agrege, on filtre par LastSeen
    var wheaCorrected = (hw.WHEA_Corrected || []).filter(function(h) { return parseDate(h.LastSeen) >= cutoff; });
    // Somme des occurrences corrigees sur la periode
    var wheaCorrectedTotal = wheaCorrected.reduce(function(s, c) { return s + c.Count; }, 0);

    var gpuTDR  = (hw.GPU_TDR || []).filter(function(h) { return parseDate(h.Timestamp) >= cutoff; });
    var thermal = (hw.Thermal || []).filter(function(h) { return parseDate(h.Timestamp) >= cutoff; });

    // Decompose WHEA_Fatal par composant (pour affichage dans le drill-down)
    var fatalByComponent = { CPU: [], RAM: [], PCIe: [], GPU: [], Autre: [] };
    wheaFatal.forEach(function(h) {
        var c = h.Component || 'Autre';
        if (!fatalByComponent[c]) fatalByComponent[c] = [];
        fatalByComponent[c].push(h);
    });

    var diskAlerts = (pc.DiskInfo || []).filter(function(d) { return d.IsAlert; });

    // Boots par type (sur la periode filtree)
    var bootsByType = { ColdBoot: 0, FastStartup: 0, Resume: 0, Unknown: 0 };
    boots.forEach(function(b) {
        var t = b.BootType || 'Unknown';
        if (bootsByType[t] === undefined) bootsByType[t] = 0;
        bootsByType[t]++;
    });

    // ==============================================================
    // v5.3 : Batterie + Services EDR
    // v5.4 : Boot Performance + Disk Health SMART
    // Tous peuvent etre null/absents (retrocompat JSON v5.2 et anciens)
    // ==============================================================
    var battery = pc.BatteryInfo || null;
    var batteryAlert = !!(battery && battery.IsAlert);

    // v2.2 : services surveilles (liste pilotee par config). Le service Role='EDR'
    // pilote le score/badge ; la liste complete est montree dans le drill-down.
    var monitoredServices = (pc.ServicesHealth && pc.ServicesHealth.Monitored) ? pc.ServicesHealth.Monitored : [];
    var edr = null;
    for (var mi = 0; mi < monitoredServices.length; mi++) {
        if (monitoredServices[mi] && monitoredServices[mi].Role === 'EDR') { edr = monitoredServices[mi]; break; }
    }
    var edrAlert = !!(edr && edr.IsAlert);
    // v2.3.3 : on distingue EDR ARRETE (installe mais en alerte) vs ABSENT (pas
    // installe) - c'est plus actionnable dans la carte Securite. Somme = edrAlert.
    var edrStopped = !!(edr && edr.IsAlert && edr.Installed);
    var edrAbsent  = !!(edr && edr.IsAlert && !edr.Installed);

    // v2.5.1 : client VPN (present + version). Inventaire, PAS une alerte -> non
    // cable au score/badge. null si Collector < 2.5.1 (champ non collecte).
    var vpn = pc.VpnClient || null;

    var bootPerf = pc.BootPerformance || null;
    var bootPerfAlert = !!(bootPerf && bootPerf.IsAlert);

    var diskHealth = pc.DiskHealth || [];
    var diskSmartAlerts = diskHealth.filter(function(d) { return d.IsAlert; });
    var diskSmartAlert = diskSmartAlerts.length > 0;
    // WearPct max parmi les disques (null si tous null)
    var diskWorstWear = null;
    diskHealth.forEach(function(d) {
        if (d.WearPct !== null && d.WearPct !== undefined) {
            if (diskWorstWear === null || d.WearPct > diskWorstWear) diskWorstWear = d.WearPct;
        }
    });

    // v5.7 : moniteurs externes (peut etre absent si Collector < v5.5)
    var monitors = pc.Monitors || [];
    var SEUIL_MONITOR_OLD_YEARS = screenAgeThreshold;
    var oldMonitors = monitors.filter(function(m) {
        return m.AgeYears !== null && m.AgeYears !== undefined && m.AgeYears >= SEUIL_MONITOR_OLD_YEARS;
    });
    var oldMonitorAlert = oldMonitors.length > 0;

    // v2.1.13 : piste materielle "memoire". Co-occurrence (PAS coincidence a la minute,
    // qui serait bruyante : les WHEA corrigees sont un bruit de fond continu, les fatales
    // provoquent un BSOD pas un crash d'appli). Signal de fond du PC -> calcule sur l'ensemble
    // des donnees (hw.WHEA_* complet, TopCrashers global), independamment du filtre periode.
    // Garde-fou : uniquement les codes de NATURE memoire (acces invalide / corruption), jamais
    // .NET ni fige. C'est une PISTE (memtest a envisager), pas un diagnostic.
    var MEMORY_EXCEPTION_CODES = ['0xc0000005', '0xc0000374'];
    var memoryCrashers = (pc.TopCrashers || []).filter(function(c) {
        return c.ExceptionCode && MEMORY_EXCEPTION_CODES.indexOf(String(c.ExceptionCode).toLowerCase()) !== -1;
    });
    var ramWheaFatal     = (hw.WHEA_Fatal     || []).filter(function(h) { return h.Component === 'RAM'; });
    var ramWheaCorrected = (hw.WHEA_Corrected || []).filter(function(h) { return h.Component === 'RAM'; });
    var memoryPiste = {
        active: (memoryCrashers.length > 0) && ((ramWheaFatal.length + ramWheaCorrected.length) > 0),
        crashers: memoryCrashers,
        ramFatal: ramWheaFatal,
        ramCorrected: ramWheaCorrected
    };

    // v2.4.0 : cycle de vie (decommission). Registre separe, embed champ Decom.
    var decom = pc.Decom || null;
    var decomTodo = !!(decom && decom.Statut === 'A faire');
    var decomDone = !!(decom && decom.Statut === 'Fait');
    var decomDaysLeft = null, decomLate = false;
    if (decomTodo && decom.TargetDate) {
        decomDaysLeft = Math.ceil((parseDate(decom.TargetDate + ' 00:00:00') - generatedAt) / 86400000);
        decomLate = (decomDaysLeft < 0);
    }
    // marque "Fait" mais le poste remonte encore = incoherence a signaler
    var decomDoneOnline = !!(decomDone && !pc.IsOffline);

    var result = {
        pc: pc,
        crashes: crashes, boots: boots, reboots: reboots, bsods: bsods, warnings: warnings,
        topRAM: pc.TopRAM || [], diskInfo: pc.DiskInfo || [], diskAlerts: diskAlerts,
        topCrashers: pc.TopCrashers || [],
        anomalies: pc.Anomalies || [],
        decom: decom,
        decomTodo: decomTodo,
        decomDone: decomDone,
        decomLate: decomLate,
        decomDaysLeft: decomDaysLeft,
        decomDoneOnline: decomDoneOnline,
        bootsLongs: bootsLongs, dernierBoot: dernierBoot,
        bootsByType: bootsByType,
        crashCount: crashes.length,
        bsodCount: bsods.length,
        bsodClassifieCount: crashes.filter(function(c) { return c.Type === 'BSOD'; }).length,
        warningCount: warnings.length,
        bootLongCount: bootsLongs.length,
        diskAlertCount: diskAlerts.length,
        // v5.2 : WHEA separe Fatal vs Corrected
        wheaFatal: wheaFatal,
        wheaCorrected: wheaCorrected,
        wheaCorrectedTotal: wheaCorrectedTotal,
        fatalByComponent: fatalByComponent,
        memoryPiste: memoryPiste,
        gpuTDR: gpuTDR,
        thermal: thermal,
        // hwCount = seulement les alertes qui doivent impacter le score
        // (les corrected ne plombent pas un PC)
        hwCount: wheaFatal.length + gpuTDR.length + thermal.length,
        collectedHoursAgo: hoursSince(pc.CollectedAt),
        // v5.3 / v5.4
        battery: battery,
        batteryAlert: batteryAlert,
        monitoredServices: monitoredServices,
        edr: edr,
        edrAlert: edrAlert,
        edrStopped: edrStopped,
        edrAbsent: edrAbsent,
        vpn: vpn,
        bootPerf: bootPerf,
        bootPerfAlert: bootPerfAlert,
        diskHealth: diskHealth,
        diskSmartAlerts: diskSmartAlerts,
        diskSmartAlert: diskSmartAlert,
        diskWorstWear: diskWorstWear,
        // v5.7 : moniteurs externes
        monitors: monitors,
        oldMonitors: oldMonitors,
        oldMonitorAlert: oldMonitorAlert,
        // v1.8 : RAM/GPU/throttling (null ou [] si JSON < 1.8)
        memory: pc.MemoryInventory || null,
        gpuInventory: pc.GPUInventory || [],
        hardwareHealth: pc.HardwareHealth || {}
    };
    result.score = computeScore(result);
    return result;
}

// ===== TRI =====
function sortPCs(pcs) {
    var dir = state.sort.dir === 'asc' ? 1 : -1;
    pcs.sort(function(a, b) {
        switch (state.sort.col) {
            case 'name':   return dir * a.pc.PC.localeCompare(b.pc.PC);
            case 'site':   return dir * (a.pc.Site || '').localeCompare(b.pc.Site || '');
            case 'user':   return dir * (a.pc.CurrentUser || '').localeCompare(b.pc.CurrentUser || '');
            case 'cpu': {
                // Tri par annee du CPU (asc = plus ancien d'abord). CPU inconnus
                // (CPUYear null) toujours en bas, quel que soit le sens.
                var ay = a.pc.CPUYear, by = b.pc.CPUYear;
                var an = (ay === null || ay === undefined || ay === '');
                var bn = (by === null || by === undefined || by === '');
                if (an && bn) return 0;
                if (an) return 1;
                if (bn) return -1;
                return dir * (ay - by);
            }
            case 'uptime':
                var ua = a.pc.UptimeDays || 0, ub = b.pc.UptimeDays || 0;
                return dir * (ua - ub);
            case 'fresh':  return dir * (a.collectedHoursAgo - b.collectedHoursAgo);
            case 'boot':
                var ba = a.dernierBoot ? a.dernierBoot.DurationMin : 0;
                var bb = b.dernierBoot ? b.dernierBoot.DurationMin : 0;
                return dir * (ba - bb);
            case 'crash':  return dir * (a.crashCount - b.crashCount);
            case 'bsod':   return dir * (a.bsodCount - b.bsodCount);
            case 'hw':     return dir * (a.hwCount - b.hwCount);
            case 'disk':
                var da = a.diskInfo.length > 0 ? a.diskInfo.reduce(function(m, d) { return d.PctFree < m ? d.PctFree : m; }, 100) : 100;
                var db = b.diskInfo.length > 0 ? b.diskInfo.reduce(function(m, d) { return d.PctFree < m ? d.PctFree : m; }, 100) : 100;
                return dir * (da - db);
            case 'perf':   return dir * (a.warningCount - b.warningCount);
            case 'score':
            default:       return dir * (a.score - b.score);
        }
    });
}

// ===== FILTRE KPI =====
function matchKpiFilter(p, filter) {
    if (!filter) return true;
    var crashRecentCutoff = new Date(generatedAt);
    crashRecentCutoff.setDate(crashRecentCutoff.getDate() - seuilCrashRecent);
    switch (filter) {
        case 'offline':     return p.pc.IsOffline;
        case 'online':      return !p.pc.IsOffline;
        case 'healthy':     return p.score === 0;   // cockpit : vue "Postes sains"
        case 'crash':       return p.crashCount > 0;
        case 'bsod':        return p.bsodCount > 0;
        case 'hw':          return p.hwCount > 0;
        case 'wheaCorr':    return p.wheaCorrected.length > 0;
        case 'bootLong':    return p.bootLongCount > 0;
        case 'diskAlert':   return p.diskAlertCount > 0;
        case 'oldCpu':      return RENEWAL_MODE ? isRenewalCandidate(p.pc) : (p.pc.CPUAgeCategory === 'Ancien');
        case 'crashRecent': return (p.pc.Crashes || []).some(function(c) { return parseDate(c.Timestamp) >= crashRecentCutoff; });
        // v5.3 / v5.4
        case 'edrDown':     return p.edrAlert;
        case 'edrStopped':  return p.edrStopped;
        case 'edrAbsent':   return p.edrAbsent;
        case 'osEol':       return p.pc.OSProduct === 'Windows 10';
        case 'decomTodo':   return p.decomTodo;
        case 'decomLate':   return p.decomLate;
        case 'decomOnline': return p.decomDoneOnline;
        case 'decomDone':   return (p.decomDone && !p.decomDoneOnline); // faite + hors ligne = a purger
        case 'battery':     return p.batteryAlert;
        case 'bootPerf':    return p.bootPerfAlert;
        case 'smart':       return p.diskSmartAlert;
        // v5.7
        case 'oldMonitor':  return p.oldMonitorAlert;
        // v2.1.11 : PC ayant au moins une anomalie (donnee a verifier)
        case 'anomaly':     return (p.anomalies || []).length > 0;
        // filtres uptime : postes qui tournent longtemps sans reboot (jamais eteints)
        case 'uptime7':     return (p.pc.UptimeDays || 0) >= 7 && (p.pc.UptimeDays || 0) < 30;
        case 'uptime30':    return (p.pc.UptimeDays || 0) >= 30;
        default: return true;
    }
}

// ===== RENDER =====
function render() {
    var cutoff = new Date(generatedAt);
    cutoff.setDate(cutoff.getDate() - state.days);

    // v2.1.11 : etat visuel du bouton-filtre Anomalies (peut ne pas exister si 0 anomalie)
    var anomBtn = document.getElementById('anomalyFilterBtn');
    if (anomBtn) anomBtn.classList.toggle('active', state.kpiFilter === 'anomaly');

    var searchTerm = document.getElementById('searchInput').value.trim().toLowerCase();
    state.siteFilter = document.getElementById('siteFilter').value;
    state.cpuFilter  = document.getElementById('cpuFilter').value;
    state.osFilter   = document.getElementById('osFilter').value;
    var mfEl = document.getElementById('modelFilter'); state.modelFilter = mfEl ? mfEl.value : '';
    var vfEl = document.getElementById('vpnFilter'); state.vpnFilter = vfEl ? vfEl.value : '';

    // Enrichir tous les PCs
    var allEnriched = pcData.map(function(pc) { return enrichPC(pc, cutoff); });

    // KPIs : calcules sur l'ensemble (filtres de type "search/site" mais pas "maskHealthy/kpiFilter")
    // Pour que clicker un KPI filtre le tableau sans cacher le KPI lui-meme
    var baseFiltered = allEnriched.filter(function(p) {
        if (state.siteFilter && p.pc.Site !== state.siteFilter) return false;
        if (state.cpuFilter && !cpuFilterMatch(p)) return false;
        if (state.osFilter   && p.pc.OSProduct !== state.osFilter) return false;
        if (state.modelFilter && (p.pc.Model || '') !== state.modelFilter) return false;
        if (state.vpnFilter && (!p.vpn || p.vpn.Version !== state.vpnFilter)) return false;
        if (state.chassisFilter && !chassisMatch(p.pc, state.chassisFilter)) return false;
        if (searchTerm) {
            return p.pc.PC.toLowerCase().indexOf(searchTerm) !== -1 ||
                   (p.pc.CurrentUser || '').toLowerCase().indexOf(searchTerm) !== -1 ||
                   (p.pc.LastLoggedUser || '').toLowerCase().indexOf(searchTerm) !== -1 ||
                   (p.pc.SerialNumber || '').toLowerCase().indexOf(searchTerm) !== -1 ||
                   (p.pc.Model || '').toLowerCase().indexOf(searchTerm) !== -1 ||
                   (p.pc.Manufacturer || '').toLowerCase().indexOf(searchTerm) !== -1;
        }
        return true;
    });

    renderKpis(baseFiltered);

    // Vue tableau : on applique kpiFilter + maskHealthy
    var visible = baseFiltered.filter(function(p) {
        // cockpit : filtre "Mon equipe" (technicien assigne via le registre decom)
        if (state.techFilter && ckTechByPc[p.pc.PC] !== state.techFilter) return false;
        if (!matchKpiFilter(p, state.kpiFilter)) return false;
        // v2.1.11 : le filtre "anomaly" prime sur maskHealthy (axe orthogonal a la sante).
        // Sinon un PC sain-mais-anormal serait masque et le compteur du bouton mentirait.
        if (state.maskHealthy && state.kpiFilter !== 'anomaly' && p.score === 0) return false;
        // v2.1.12 : filtre par appli (clic sur un crasher) - garde les PC qui l'ont en crash.
        if (state.appFilter && !(p.topCrashers || []).some(function(c) { return String(c.AppName).toLowerCase() === state.appFilter.toLowerCase(); })) return false;
        return true;
    });

    sortPCs(visible);

    // v1.5 : Pagination
    // On slice "visible" selon la page courante et itemsPerPage.
    // Si un filtre change, il est possible que currentPage pointe sur une
    // page qui n'existe plus (ex: j'etais page 5, je filtre -> 40 PC -> 1 page).
    // On borne currentPage pour eviter un tableau vide.
    var totalPages = (state.itemsPerPage === 0) ? 1 : Math.max(1, Math.ceil(visible.length / state.itemsPerPage));
    if (state.currentPage > totalPages) state.currentPage = totalPages;
    if (state.currentPage < 1) state.currentPage = 1;

    var visibleSlice;
    if (state.itemsPerPage === 0) {
        visibleSlice = visible;  // Tout afficher
    } else {
        var startIdx = (state.currentPage - 1) * state.itemsPerPage;
        var endIdx   = startIdx + state.itemsPerPage;
        visibleSlice = visible.slice(startIdx, endIdx);
    }

    renderTable(visibleSlice);
    renderPagination(visible.length, totalPages);
    renderActiveFilters();
    renderGlobalCrashers(baseFiltered);
    renderBootBreakdown(baseFiltered);
    renderMonitorInventory(baseFiltered);
    renderModelInventory(baseFiltered);   // v2.4.7 : repartition des modeles machine
    renderDecomStats();   // v2.4.2 : audit cycle de vie (registre complet, independant des filtres)

    document.getElementById('pcCount').textContent = visible.length + ' / ' + allEnriched.length + ' appareil(s)';
}

// ===== v1.5 : PAGINATION =====
// Rendu du bloc de pagination (affiche en haut et en bas du tableau).
// totalItems = nombre total de PC apres tous les filtres (pour les compteurs)
// totalPages = calcule par render() pour coherence
function renderPagination(totalItems, totalPages) {
    var containers = ['paginationTop', 'paginationBottom'];
    var per = state.itemsPerPage;
    var cur = state.currentPage;

    // Etat actif/inactif pour les selecteurs
    function perActive(val) { return per === val ? ' active' : ''; }

    // Calculer les bornes d'affichage
    var startNum, endNum;
    if (per === 0) {
        startNum = totalItems > 0 ? 1 : 0;
        endNum   = totalItems;
    } else {
        startNum = totalItems > 0 ? (cur - 1) * per + 1 : 0;
        endNum   = Math.min(cur * per, totalItems);
    }

    // Construction des numeros de page (logique "1 ... 3 4 [5] 6 7 ... 20")
    var pagesHtml = '';
    if (per > 0 && totalPages > 1) {
        var range = [];
        if (totalPages <= 7) {
            // Peu de pages : on les affiche toutes
            for (var i = 1; i <= totalPages; i++) range.push(i);
        } else {
            // Beaucoup de pages : ellipses intelligentes
            range.push(1);
            if (cur > 3) range.push('...');
            for (var j = Math.max(2, cur - 1); j <= Math.min(totalPages - 1, cur + 1); j++) {
                range.push(j);
            }
            if (cur < totalPages - 2) range.push('...');
            range.push(totalPages);
        }
        range.forEach(function(p) {
            if (p === '...') {
                pagesHtml += '<span class="pagination-ellipsis">...</span>';
            } else {
                var active = p === cur ? ' active' : '';
                pagesHtml += '<button class="pagination-page' + active + '" onclick="goToPage(' + p + ')">' + p + '</button>';
            }
        });
    }

    // Bouton Precedent / Suivant
    var prevDisabled = (per === 0 || cur === 1) ? ' disabled' : '';
    var nextDisabled = (per === 0 || cur >= totalPages) ? ' disabled' : '';

    // Boutons "afficher plus" et "tout afficher"
    var canShowMore = (per > 0 && cur < totalPages);
    var showMoreBtn = canShowMore
        ? '<button class="pagination-showmore" onclick="showMorePage()" title="Afficher les ' + Math.min(per, totalItems - endNum) + ' PC suivants">+ Afficher ' + Math.min(per, totalItems - endNum) + ' de plus</button>'
        : '';
    var showAllBtn = (per !== 0 && totalItems > per)
        ? '<button class="pagination-showall" onclick="setItemsPerPage(0)" title="Desactiver la pagination et tout afficher">Tout afficher</button>'
        : '';

    var html =
        '<div class="pagination-bar">' +
            '<div class="pagination-info">' +
                'Affichage ' + startNum + '-' + endNum + ' sur ' + totalItems + ' PC' +
            '</div>' +
            '<div class="pagination-controls">' +
                '<span class="pagination-label">Par page :</span>' +
                '<button class="pagination-per' + perActive(20)  + '" onclick="setItemsPerPage(20)">20</button>' +
                '<button class="pagination-per' + perActive(50)  + '" onclick="setItemsPerPage(50)">50</button>' +
                '<button class="pagination-per' + perActive(100) + '" onclick="setItemsPerPage(100)">100</button>' +
                '<button class="pagination-per' + perActive(0)   + '" onclick="setItemsPerPage(0)">Tous</button>' +
                '<span class="pagination-sep"></span>' +
                '<button class="pagination-nav' + prevDisabled + '" onclick="goToPage(' + (cur - 1) + ')"' + prevDisabled + '>&larr; Prec.</button>' +
                pagesHtml +
                '<button class="pagination-nav' + nextDisabled + '" onclick="goToPage(' + (cur + 1) + ')"' + nextDisabled + '>Suiv. &rarr;</button>' +
                showMoreBtn +
                showAllBtn +
            '</div>' +
        '</div>';

    containers.forEach(function(id) {
        var el = document.getElementById(id);
        if (el) el.innerHTML = html;
    });
}

function setItemsPerPage(n) {
    state.itemsPerPage = n;
    state.currentPage = 1;
    try { localStorage.setItem('pcpulse_itemsPerPage', n); } catch (e) {}
    render();
}

function goToPage(n) {
    state.currentPage = n;
    render();
    // Scroll vers le haut du tableau pour le confort utilisateur
    var tableEl = document.getElementById('deviceTable');
    if (tableEl) tableEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// "Afficher + de N" : passe a la page suivante sans changer itemsPerPage
// (difference avec goToPage : ne scroll pas en haut, reste dans la continuite)
function showMorePage() {
    state.currentPage++;
    render();
}

// Reset currentPage a 1 des qu'un filtre change (sinon on peut se retrouver
// sur une page qui n'existe plus et voir un tableau vide le temps d'un render)
function resetPaginationToFirstPage() {
    state.currentPage = 1;
}

function renderActiveFilters() {
    var container = document.getElementById('activeFilters');
    var chips = [];
    if (state.kpiFilter) {
        var labels = {
            'offline': 'Offline', 'online': 'En ligne', 'healthy': 'Postes sains', 'crash': 'Crash/Freeze',
            'bsod': 'BSOD', 'hw': 'Erreurs fatales HW', 'wheaCorr': 'WHEA corrigees',
            'bootLong': 'Boots longs',
            'diskAlert': 'Disques critiques', 'oldCpu': (RENEWAL_MODE ? 'A renouveler' : 'CPU anciens'),
            'crashRecent': 'Crash recent (' + seuilCrashRecent + 'j)',
            // v5.3 / v5.4
            'edrDown': 'EDR en panne',
            'battery': 'Batterie usee',
            'bootPerf': 'Boots lents',
            'smart': 'SMART alerte',
            'oldMonitor': 'Ecrans anciens (>=7 ans)',
            'uptime7': 'Uptime 7-30 j', 'uptime30': 'Uptime > 30 j'
        };
        chips.push('<span class="active-filter-chip">' + labels[state.kpiFilter] + ' <span class="close" onclick="clearKpiFilter()">&times;</span></span>');
    }
    if (state.maskHealthy) {
        chips.push('<span class="active-filter-chip">PC sains masques <span class="close" onclick="toggleMaskHealthy()">&times;</span></span>');
    }
    if (state.appFilter) {
        chips.push('<span class="active-filter-chip">Appli&nbsp;: ' + state.appFilter + ' <span class="close" onclick="clearAppFilter()">&times;</span></span>');
    }
    if (state.techFilter) {
        chips.push('<span class="active-filter-chip">Technicien&nbsp;: ' + state.techFilter + ' <span class="close" onclick="ckClearTech()">&times;</span></span>');
    }
    // v2.1.14 : bouton "tout effacer" des qu'un filtre quelconque est actif - y compris les
    // menus site/CPU et la recherche, qui ne generent pas de chip a eux seuls.
    var searchActive = (document.getElementById('searchInput').value || '').trim().length > 0;
    var anyFilter = !!(state.kpiFilter || state.appFilter || state.techFilter || state.maskHealthy || state.siteFilter || state.cpuFilter || state.osFilter || state.modelFilter || state.vpnFilter || state.chassisFilter || searchActive);
    var resetBtn = document.getElementById('resetBtn');
    if (resetBtn) resetBtn.disabled = !anyFilter;
    if (anyFilter) {
        chips.push('<button class="clear-all-filters" onclick="clearAllFilters()" title="Reinitialiser tous les filtres">&times; Tout effacer</button>');
    }
    container.innerHTML = chips.join('');
}

function kpiCard(key, colorClass, value, label, cardClass) {
    var active = state.kpiFilter === key ? ' active' : '';
    // v5.5 : attenuer les cards "vides" (value=0) et emphasiser les critiques
    var quietCls = '';
    if ((value === 0 || value === '0') && key !== '') {
        quietCls = ' kpi-quiet';
    } else if (typeof value === 'number' && value > 0 && (cardClass === 'card-danger' || cardClass === 'card-hardware')) {
        quietCls = ' kpi-loud';
    }
    return '<div class="kpi-card ' + cardClass + quietCls + active + '" onclick="toggleKpiFilter(\'' + key + '\')">' +
           '<div class="kpi-value ' + colorClass + '">' + value + '</div>' +
           '<div class="kpi-label">' + label + '</div></div>';
}

function renderKpis(pcs) {
    var totalPC = pcs.length;
    var totalOffline = pcs.filter(function(p) { return p.pc.IsOffline; }).length;
    var totalCrash = pcs.reduce(function(s, p) { return s + p.crashCount; }, 0);
    var totalBSOD = pcs.reduce(function(s, p) { return s + p.bsodCount; }, 0);
    var totalBootsLongs = pcs.reduce(function(s, p) { return s + p.bootLongCount; }, 0);
    var totalDiskAlerts = pcs.reduce(function(s, p) { return s + p.diskAlertCount; }, 0);
    var totalHardware = pcs.reduce(function(s, p) { return s + p.hwCount; }, 0);
    var pcWithCorrected = pcs.filter(function(p) { return p.wheaCorrected.length > 0; }).length;
    var totalOldCPU = pcs.filter(function(p) { return RENEWAL_MODE ? isRenewalCandidate(p.pc) : (p.pc.CPUAgeCategory === 'Ancien'); }).length;

    var totalEdrDown = pcs.filter(function(p) { return p.edrAlert; }).length;
    var totalEdrStopped = pcs.filter(function(p) { return p.edrStopped; }).length;
    var totalEdrAbsent  = pcs.filter(function(p) { return p.edrAbsent; }).length;
    var totalOsEol      = pcs.filter(function(p) { return p.pc.OSProduct === 'Windows 10'; }).length;  // Win10 = fin de support (oct. 2025)
    var totalDecomTodo   = pcs.filter(function(p) { return p.decomTodo; }).length;
    var totalDecomLate   = pcs.filter(function(p) { return p.decomLate; }).length;
    var totalDecomOnline = pcs.filter(function(p) { return p.decomDoneOnline; }).length;
    var totalDecomDone   = pcs.filter(function(p) { return p.decomDone && !p.decomDoneOnline; }).length;
    var totalBatteryWorn = pcs.filter(function(p) { return p.batteryAlert; }).length;
    var totalBootPerfSlow = pcs.filter(function(p) { return p.bootPerfAlert; }).length;
    var totalSmartAlert = pcs.filter(function(p) { return p.diskSmartAlert; }).length;
    var totalUptime30    = pcs.filter(function(p) { return (p.pc.UptimeDays || 0) >= 30; }).length;
    var totalUptime7to30 = pcs.filter(function(p) { var u = p.pc.UptimeDays || 0; return u >= 7 && u < 30; }).length;

    // v5.7 : inventaire moniteurs externes
    var totalMonitors = pcs.reduce(function(s, p) { return s + (p.monitors ? p.monitors.length : 0); }, 0);
    var pcWithOldMonitor = pcs.filter(function(p) { return p.oldMonitorAlert; }).length;
    var pcWithMonitor = pcs.filter(function(p) { return p.monitors && p.monitors.length > 0; }).length;

    var crashRecentCutoff = new Date(generatedAt);
    crashRecentCutoff.setDate(crashRecentCutoff.getDate() - seuilCrashRecent);
    var pcCrashRecents = pcs.filter(function(p) {
        return (p.pc.Crashes || []).some(function(c) { return parseDate(c.Timestamp) >= crashRecentCutoff; });
    }).length;

    var pcHealthy = pcs.filter(function(p) { return p.score === 0; }).length;
    var pcWithIssues = totalPC - pcHealthy;
    var pctOnline = totalPC > 0 ? Math.round(((totalPC - totalOffline) / totalPC) * 100) : 0;
    var pctHealthy = totalPC > 0 ? Math.round((pcHealthy / totalPC) * 100) : 0;

    // ========================================================
    // SUMMARY BAR : 3 chiffres essentiels en haut de page
    // ========================================================
    var summary = document.getElementById('summaryBar');
    summary.innerHTML =
        '<div class="summary-tile">' +
          '<div class="summary-icon blue">&#128187;</div>' +
          '<div class="summary-content">' +
            '<div class="summary-value">' + totalPC + '</div>' +
            '<div class="summary-label">PC monitor&eacute;s</div>' +
            '<div class="summary-sub">dans la p&eacute;riode s&eacute;lectionn&eacute;e</div>' +
          '</div>' +
        '</div>' +
        '<div class="summary-tile clickable' + (state.kpiFilter === 'offline' ? ' active' : '') + '" onclick="toggleKpiFilter(\'offline\')" title="Cliquer pour filtrer les postes hors ligne">' +
          ckRing(pctOnline, 'var(--green)') +
          '<div class="summary-content">' +
            '<div class="summary-value">' + (totalPC - totalOffline) + ' <span style="font-size:14px;color:var(--text-muted);font-weight:500">/ ' + totalPC + '</span></div>' +
            '<div class="summary-label">En ligne (24h)</div>' +
            '<div class="summary-sub">' + pctOnline + '% joignable' + (totalOffline > 0 ? ' &middot; ' + totalOffline + ' hors ligne &rarr;' : '') + '</div>' +
          '</div>' +
        '</div>' +
        '<div class="summary-tile">' +
          ckRing(pctHealthy, (pctHealthy >= 80 ? 'var(--green)' : (pctHealthy >= 50 ? 'var(--orange)' : 'var(--red)'))) +
          '<div class="summary-content">' +
            '<div class="summary-value">' + pcHealthy + ' <span style="font-size:14px;color:var(--text-muted);font-weight:500">/ ' + totalPC + '</span></div>' +
            '<div class="summary-label">PC sans alerte</div>' +
            '<div class="summary-sub">' + pctHealthy + '% sains &middot; ' + pcWithIssues + ' &agrave; examiner</div>' +
          '</div>' +
        '</div>';

    // ========================================================
    // v5.6 : BARRE DE 4 BOUTONS DE GROUPES
    // ========================================================
    // Definition des 4 familles et de leurs sous-KPIs
    // Chaque sous-KPI = { filterKey, label, value, level('danger'|'warn'|'neutral') }
    var familySecurity = {
        key: 'security', family: 'family-security', icon: '&#128274;', title: 'S&eacute;curit&eacute;',
        subs: [
            { k: 'edrStopped', label: 'EDR arr&ecirc;t&eacute;',   v: totalEdrStopped, level: (totalEdrStopped > 0 ? 'danger' : 'neutral') },
            { k: 'edrAbsent',  label: 'EDR absent',                 v: totalEdrAbsent,  level: (totalEdrAbsent > 0 ? 'danger' : 'neutral') },
            { k: 'osEol',      label: 'OS fin de support',          v: totalOsEol,      level: (totalOsEol > 0 ? 'warn' : 'neutral') }
        ]
    };
    var familyStability = {
        key: 'stability', family: 'family-stability', icon: '&#128165;', title: 'Stabilit&eacute;',
        subs: [
            { k: 'crashRecent', label: 'Crash r&eacute;cent (' + seuilCrashRecent + 'j)', v: pcCrashRecents, level: (pcCrashRecents > 0 ? 'danger' : 'neutral') },
            { k: 'crash',       label: 'Crash/Freeze (' + state.days + 'j)',              v: totalCrash,     level: (totalCrash > 0 ? 'warn' : 'neutral') },
            { k: 'bsod',        label: 'BSOD (' + state.days + 'j)',                      v: totalBSOD,      level: (totalBSOD > 0 ? 'danger' : 'neutral') },
            { k: 'hw',          label: 'Erreurs fatales HW',                              v: totalHardware,  level: (totalHardware > 0 ? 'danger' : 'neutral') },
            { k: 'wheaCorr',    label: 'WHEA corrig&eacute;es',                           v: pcWithCorrected,level: 'neutral' }
        ]
    };
    var familyPerformance = {
        key: 'performance', family: 'family-performance', icon: '&#9889;', title: 'Boot &amp; uptime',
        subs: [
            { k: 'bootLong', label: 'Boots longs',          v: totalBootsLongs,   level: (totalBootsLongs > 0 ? 'warn' : 'neutral') },
            { k: 'bootPerf', label: 'Boots lents (&gt;90s)', v: totalBootPerfSlow, level: (totalBootPerfSlow > 0 ? 'warn' : 'neutral') },
            { k: 'uptime7',  label: 'Uptime 7-30 j',          v: totalUptime7to30,  level: (totalUptime7to30 > 0 ? 'warn' : 'neutral') },
            { k: 'uptime30', label: 'Uptime &gt; 30 j',       v: totalUptime30,     level: (totalUptime30 > 0 ? 'danger' : 'neutral') }
        ]
    };
    var familyMaterial = {
        key: 'material', family: 'family-material', icon: '&#128295;', title: 'Usure mat&eacute;rielle',
        subs: [
            { k: 'diskAlert',  label: 'Disques satur&eacute;s',           v: totalDiskAlerts,   level: (totalDiskAlerts > 0 ? 'warn' : 'neutral') },
            { k: 'smart',      label: 'SMART alerte',                     v: totalSmartAlert,   level: (totalSmartAlert > 0 ? 'warn' : 'neutral') },
            { k: 'battery',    label: 'Batterie us&eacute;e (&lt;60%)',   v: totalBatteryWorn,  level: (totalBatteryWorn > 0 ? 'warn' : 'neutral') },
            { k: 'oldCpu',     label: (RENEWAL_MODE ? 'A renouveler' : 'CPU anciens'), v: totalOldCPU, level: (totalOldCPU > 0 ? (RENEWAL_MODE ? 'warn' : 'danger') : 'neutral') },
            { k: 'oldMonitor', label: 'Ecrans &acirc;g&eacute;s (&ge;' + screenAgeThreshold + ' ans)', v: pcWithOldMonitor, level: (pcWithOldMonitor > 0 ? 'warn' : 'neutral') }
        ]
    };

    var familyLifecycle = {
        key: 'lifecycle', family: 'family-lifecycle', icon: '&#128260;', title: 'Cycle de vie',
        subs: [
            { k: 'decomTodo',   label: 'A d&eacute;commissionner', v: totalDecomTodo,   level: (totalDecomTodo > 0 ? 'warn' : 'neutral') },
            { k: 'decomLate',   label: 'En retard',                v: totalDecomLate,   level: (totalDecomLate > 0 ? 'danger' : 'neutral') },
            { k: 'decomOnline', label: 'Fait mais en ligne',       v: totalDecomOnline, level: (totalDecomOnline > 0 ? 'danger' : 'neutral') },
            { k: 'decomDone',   label: 'Faite (&agrave; purger)',  v: totalDecomDone,   level: 'neutral' }
        ]
    };

    var families = [familySecurity, familyStability, familyPerformance, familyMaterial, familyLifecycle];

    // Quels filtres KPI appartiennent a quelle famille (pour marquer
    // un bouton "active" si le filtre courant pointe vers cette famille)
    var keysOf = function(fam) { return fam.subs.map(function(s) { return s.k; }); };

    // Rendu des boutons
    var bar = document.getElementById('kpiGroupBar');
    var html = '';
    families.forEach(function(fam) {
        // Compter les alertes "chaudes" de la famille (warn + danger)
        // et celles qui sont "critiques" (danger)
        var hotCount = 0, criticalCount = 0;
        fam.subs.forEach(function(s) {
            if (s.level === 'danger' && s.v > 0) { hotCount += s.v; criticalCount += s.v; }
            else if (s.level === 'warn' && s.v > 0) { hotCount += s.v; }
        });

        var stateCls = 'calm';
        if (criticalCount > 0)   stateCls = 'danger';
        else if (hotCount > 0)   stateCls = 'warn';

        // "Active" = le filtre KPI courant est l'un des sous-KPIs de la famille
        var active = (state.kpiFilter && keysOf(fam).indexOf(state.kpiFilter) >= 0) ? ' active' : '';

        // Resume textuel sous le titre
        var sub;
        if (hotCount === 0) {
            sub = 'Aucune alerte';
        } else {
            var parts = [];
            fam.subs.forEach(function(s) {
                if (s.v > 0 && s.level !== 'neutral') parts.push(s.v + ' ' + s.label.replace(/&nbsp;/g,' '));
            });
            sub = parts.slice(0, 2).join(' &middot; ');
            if (parts.length > 2) sub += ' &hellip;';
        }

        // Popover : items cliquables pour filtrer sur un sous-KPI precis
        // onclick avec event.stopPropagation pour ne pas declencher le clic
        // du bouton parent (qui filtre sur la famille entiere)
        var itemsHtml = '';
        fam.subs.forEach(function(s) {
            var valCls = (s.v === 0) ? '' : (s.level === 'danger' ? 'danger' : (s.level === 'warn' ? 'warn' : ''));
            var itemActive = (state.kpiFilter === s.k) ? ' active' : '';
            itemsHtml += '<div class="kpi-popover-item' + itemActive + '" onclick="event.stopPropagation();toggleKpiFilter(\'' + s.k + '\')">' +
                         '<span class="pop-label">' + s.label + '</span>' +
                         '<span class="pop-value ' + valCls + '">' + s.v + '</span>' +
                         '</div>';
        });

        // Le clic du bouton filtre sur "crashRecent" pour stab, "edrDown" pour secu,
        // "bootPerf" pour perf, "diskAlert" pour usure : c'est le sous-KPI le plus
        // "important" de chaque famille. Ou alors on passe un filter meta "family"
        // -> plus simple : le clic toggle le filtre sur LE sous-KPI le plus critique
        // qui est > 0, sinon sur le premier de la famille.
        var primaryKey = fam.subs[0].k;
        for (var i = 0; i < fam.subs.length; i++) {
            if (fam.subs[i].v > 0 && fam.subs[i].level === 'danger') { primaryKey = fam.subs[i].k; break; }
        }
        if (primaryKey === fam.subs[0].k) {
            for (var j = 0; j < fam.subs.length; j++) {
                if (fam.subs[j].v > 0 && fam.subs[j].level === 'warn') { primaryKey = fam.subs[j].k; break; }
            }
        }
        // Aucune alerte chaude mais un compteur neutre > 0 (ex: "Faite a purger")
        // -> un clic direct sur la carte filtre ce sous-KPI plutot que rien.
        if (primaryKey === fam.subs[0].k && fam.subs[0].v === 0) {
            for (var n = 0; n < fam.subs.length; n++) {
                if (fam.subs[n].v > 0) { primaryKey = fam.subs[n].k; break; }
            }
        }

        // cockpit : memorise le sous-KPI representatif + le compteur "chaud"
        // de chaque famille pour cabler le rail de navigation "Par axe".
        ckFamilyMeta[fam.key] = { primary: primaryKey, hot: hotCount, keys: keysOf(fam) };

        html += '<button type="button" class="kpi-group-btn ' + stateCls + ' ' + fam.family + active + '" ' +
                'onclick="toggleKpiFilter(\'' + primaryKey + '\')">' +
                '<div class="kpi-group-btn-icon">' + fam.icon + '</div>' +
                '<div class="kpi-group-btn-content">' +
                  '<div class="kpi-group-btn-title">' + fam.title + '</div>' +
                  '<div class="kpi-group-btn-sub">' + sub + '</div>' +
                '</div>' +
                '<div class="kpi-group-btn-count">' + hotCount + '</div>' +
                '<div class="kpi-group-popover">' +
                  '<div class="kpi-popover-items">' + itemsHtml + '</div>' +
                '</div>' +
                '</button>';
    });
    bar.innerHTML = html;

    // cockpit : alimente la file d'actions a droite (+ sous-section "Mon equipe").
    var ckAgg = {
        totalPC: totalPC, pcWithIssues: pcWithIssues, totalOffline: totalOffline, pcHealthy: pcHealthy,
        totalEdrDown: totalEdrDown, totalOsEol: totalOsEol, totalBSOD: totalBSOD,
        totalDiskAlerts: totalDiskAlerts, totalSmartAlert: totalSmartAlert, totalBatteryWorn: totalBatteryWorn,
        totalUptime30: totalUptime30, totalDecomLate: totalDecomLate, totalDecomTodo: totalDecomTodo
    };
    renderActionFile(ckAgg, pcs);
}

// ===== AGREGATION PARC : REPARTITION DES DEMARRAGES (v5.2) =====
function renderBootBreakdown(pcs) {
    var container = document.getElementById('bootBreakdown');
    var totals = { ColdBoot: 0, FastStartup: 0, Resume: 0, Unknown: 0 };
    var pcCountWithBootData = 0;

    pcs.forEach(function(p) {
        var bt = p.bootsByType || {};
        var hasData = (bt.ColdBoot || 0) + (bt.FastStartup || 0) + (bt.Resume || 0) + (bt.Unknown || 0);
        if (hasData > 0) pcCountWithBootData++;
        totals.ColdBoot    += bt.ColdBoot    || 0;
        totals.FastStartup += bt.FastStartup || 0;
        totals.Resume      += bt.Resume      || 0;
        totals.Unknown     += bt.Unknown     || 0;
    });

    var grandTotal = totals.ColdBoot + totals.FastStartup + totals.Resume + totals.Unknown;

    if (grandTotal === 0) {
        container.innerHTML = '<div class="boot-breakdown-empty">Aucun demarrage detecte sur la periode ' +
            '(ou Collector en version anterieure a v5.2 : l\'Event 27 Kernel-Boot n\'etait pas encore collecte)</div>';
        return;
    }

    var pctCold    = (totals.ColdBoot    / grandTotal) * 100;
    var pctFast    = (totals.FastStartup / grandTotal) * 100;
    var pctResume  = (totals.Resume      / grandTotal) * 100;
    var pctUnknown = (totals.Unknown     / grandTotal) * 100;

    var barHtml = '<div class="boot-breakdown-bar">';
    if (pctCold > 0)    barHtml += '<div class="boot-breakdown-segment cold"    style="width:' + pctCold.toFixed(1)    + '%" title="ColdBoot : ' + totals.ColdBoot + '">'    + (pctCold    >= 5 ? totals.ColdBoot    : '') + '</div>';
    if (pctFast > 0)    barHtml += '<div class="boot-breakdown-segment fast"    style="width:' + pctFast.toFixed(1)    + '%" title="Fast Startup : ' + totals.FastStartup + '">' + (pctFast    >= 5 ? totals.FastStartup : '') + '</div>';
    if (pctResume > 0)  barHtml += '<div class="boot-breakdown-segment resume"  style="width:' + pctResume.toFixed(1)  + '%" title="Resume hibernation : ' + totals.Resume + '">' + (pctResume  >= 5 ? totals.Resume      : '') + '</div>';
    if (pctUnknown > 0) barHtml += '<div class="boot-breakdown-segment unknown" style="width:' + pctUnknown.toFixed(1) + '%" title="Inconnu : ' + totals.Unknown + '">' + (pctUnknown >= 5 ? totals.Unknown     : '') + '</div>';
    barHtml += '</div>';

    var legendHtml = '<div class="boot-breakdown-legend">' +
        '<div class="boot-breakdown-legend-item"><div class="boot-breakdown-legend-dot" style="background:#3a6cbf"></div>&#10052; ColdBoot : ' + totals.ColdBoot + ' (' + pctCold.toFixed(1) + '%)</div>' +
        '<div class="boot-breakdown-legend-item"><div class="boot-breakdown-legend-dot" style="background:var(--yellow)"></div>&#9889; Fast Startup : ' + totals.FastStartup + ' (' + pctFast.toFixed(1) + '%)</div>' +
        '<div class="boot-breakdown-legend-item"><div class="boot-breakdown-legend-dot" style="background:var(--purple)"></div>&#128164; Resume : ' + totals.Resume + ' (' + pctResume.toFixed(1) + '%)</div>';
    if (totals.Unknown > 0) {
        legendHtml += '<div class="boot-breakdown-legend-item"><div class="boot-breakdown-legend-dot" style="background:var(--text-faint)"></div>? Inconnu : ' + totals.Unknown + '</div>';
    }
    legendHtml += '</div>';

    // Note metier : expliquer ce que signifient ces chiffres
    var noteHtml = '';
    var fastDomine = (pctFast > 50);
    var coldDomine = (pctCold > 50);
    if (fastDomine) {
        noteHtml = '<div class="boot-breakdown-note">Le Fast Startup domine sur le parc (' + pctFast.toFixed(0) + '%). ' +
            'PCPulse detecte correctement les reprises d\'activite (Fast Startup et wake Modern Standby) : ' +
            'l\'uptime affiche reflete le temps depuis la derniere reprise utilisateur, pas depuis le dernier vrai cold boot.</div>';
    } else if (coldDomine && pctFast < 10) {
        noteHtml = '<div class="boot-breakdown-note">Le parc est majoritairement en ColdBoot (' + pctCold.toFixed(0) + '%). ' +
            'Fast Startup semble desactive sur la plupart des postes.</div>';
    }

    container.innerHTML = barHtml + legendHtml + noteHtml +
        '<div style="font-size:11px;color:var(--text-muted)">' + pcCountWithBootData + ' PC remontent des donnees de demarrage sur ' + pcs.length + ' (' + grandTotal + ' d&eacute;marrages au total)</div>';
}

// ===== AGREGATION PARC : TOP CRASHERS GLOBAL (v1.7 : 3 sections) =====
function renderGlobalCrashers(pcs) {
    var agg = {}, userSet = {};
    pcs.forEach(function(p) {
        // v2.5.0 : identites du poste, pour ecarter un nom d'utilisateur happe
        // par erreur comme une appli en echec.
        [p.pc && p.pc.CurrentUser, p.pc && p.pc.LastLoggedUser].forEach(function(u) {
            u = String(u || '').toLowerCase().trim();
            if (u && u !== '(aucune session)') userSet[u] = true;
        });
        (p.topCrashers || []).forEach(function(c) {
            var key = c.AppName;
            // v1.7 : blacklist HARD - on zappe completement a l'agregation
            if (isBlacklistedHard(key)) return;
            if (!agg[key]) agg[key] = { name: key, total: 0, pcCount: 0, pcSet: {}, isCrash: false };
            agg[key].total += c.CrashCount;
            // v2.1.4 : un vrai crash (ou Type absent = JSON legacy) prime sur l'agregat ;
            // une cle reste 'app_failure' seulement si TOUTES ses occurrences le sont.
            if (c.Type !== 'app_failure') agg[key].isCrash = true;
            if (!agg[key].pcSet[p.pc.PC]) {
                agg[key].pcCount++;
                agg[key].pcSet[p.pc.PC] = true;
            }
        });
    });

    var arr = Object.keys(agg).map(function(k) {
        var a = agg[k];
        var cls = classifyCrasher(a.name, a.total, a.pcCount, a.isCrash ? 'crash' : 'app_failure');
        a.score  = cls.score;
        a.level  = cls.level;
        a.forced = cls.forced;
        return a;
    });

    // v2.5 : pre-filtre "bruit" (isRealApp = helper global) + retrait des noms
    // d'utilisateur happes par erreur. Une appli suivie (priorityApps) reste
    // toujours, meme mal nommee.
    var hiddenCount = 0;
    arr = arr.filter(function(c) {
        if (isPriorityApp(c.name)) return true;
        if (userSet[String(c.name).toLowerCase().trim()]) { hiddenCount++; return false; }
        if (!isRealApp(c.name)) { hiddenCount++; return false; }
        return true;
    });

    var container = document.getElementById('globalCrashers');
    if (arr.length === 0) {
        container.innerHTML = '<div class="detail-empty">Aucun crash d application agrege</div>';
        return;
    }

    // v2.5 : seuils "Plus fort impact" (faciles a ajuster). Un crasher non suivi
    // remonte dans ce groupe s'il depasse l'un OU l'autre.
    var IMPACT_TOTAL_MIN  = 20;   // nb de plantages cumules
    var IMPACT_PC_MIN     = 5;    // nb de PC touches
    var WIDESPREAD_PC_MIN = 5;    // seuil du tag "repandu" vs "concentre"

    // Appli "importante" = suivie, declaree dans PriorityApps (config.psd1).
    // Le code ne nomme aucun produit : l'exploitant liste ses applis critiques
    // (dont l'agent EDR, ex. son process .exe) dans sa config, elles ressortent ici.
    function isImportant(c) { return isPriorityApp(c.name); }

    // Formatage : fine espace insecable comme separateur de milliers (>= 1000).
    function fmtN(v) {
        var s = String(v);
        return s.length > 3 ? s.replace(/\B(?=(\d{3})+(?!\d))/g, ' ') : s;
    }

    // Tri unique dans CHAQUE groupe : volume de plantages decroissant (jamais la dispersion).
    function byTotalDesc(a, b) { return b.total - a.total; }

    var important = arr.filter(isImportant).sort(byTotalDesc);
    var rest      = arr.filter(function(c) { return !isImportant(c); });
    function isImpact(c) { return c.total >= IMPACT_TOTAL_MIN || c.pcCount >= IMPACT_PC_MIN; }
    var impact = rest.filter(isImpact).sort(byTotalDesc);
    var noise  = rest.filter(function(c) { return !isImpact(c); }).sort(byTotalDesc);

    // Rendu d'une ligne crasheur : nom (mono) + un tag chip + metriques a droite.
    // Ligne cliquable -> filtre les PC concernes. Le nom vient de l'embed (deja
    // ConvertTo-HtmlSafe) : sur dans l'attribut, lu via dataset (aucune eval JS).
    function renderRow(c, group) {
        var isFail = !c.isCrash;   // agregat 'app_failure' (echec install/maj)
        var tag = '';
        if (isPriorityApp(c.name)) {
            tag = '<span class="gc-tag gc-tag-suivi" title="Appli suivie (config)">suivi</span>';
        } else if (isFail) {
            tag = '<span class="gc-tag gc-tag-fail" title="Installation ou mise a jour qui echoue en boucle">&eacute;chec install</span>';
        } else if (group === 'impact') {
            tag = (c.pcCount >= WIDESPREAD_PC_MIN)
                ? '<span class="gc-tag" title="Touche plusieurs PC">r&eacute;pandu</span>'
                : '<span class="gc-tag" title="Concentre sur peu de PC">concentr&eacute;</span>';
        }
        var word = isFail ? '&eacute;checs' : 'plant&eacute;s';
        var activeCls = (state.appFilter && state.appFilter === c.name) ? ' active' : '';
        var impCls = (group === 'important') ? ' gc-important' : '';
        return '<div class="global-crasher-row clickable' + impCls + activeCls + '" data-appname="' + String(c.name).replace(/&/g, '&amp;') + '">' +
                  '<span class="global-crasher-name" title="' + c.name + '">' + c.name + '</span>' +
                  tag +
                  '<span class="gc-spacer"></span>' +
                  '<span class="global-crasher-stats">' +
                    '<b>' + fmtN(c.total) + '</b> ' + word + ' &middot; <b>' + fmtN(c.pcCount) + '</b> PC' +
                  '</span>' +
                '</div>';
    }

    // Rendu d'un groupe : les TOP_N premiers, puis "voir les N autres" (replie).
    var TOP_N = 5;
    function renderGroup(list, group) {
        if (list.length <= TOP_N) {
            return list.map(function(c) { return renderRow(c, group); }).join('');
        }
        var head = list.slice(0, TOP_N).map(function(c) { return renderRow(c, group); }).join('');
        var more = list.slice(TOP_N).map(function(c) { return renderRow(c, group); }).join('');
        var n = list.length - TOP_N;
        return head +
            '<div class="gc-more" onclick="toggleMore(this)"><span class="gc-more-tri">&#9656;</span>voir les ' + n + ' autre' + (n > 1 ? 's' : '') + '</div>' +
            '<div class="gc-more-body" style="display:none">' + more + '</div>';
    }

    var html = '';

    // --- Groupe 1 : Importantes - a traiter (applis suivies / EDR, quel que soit le volume) ---
    html += '<div class="gc-group">Importantes &mdash; &agrave; traiter <span class="gc-group-count">' + important.length + '</span></div>';
    html += (important.length === 0)
        ? '<div class="gc-empty">aucune appli importante en plantage</div>'
        : renderGroup(important, 'important');

    // --- Groupe 2 : Plus fort impact sur le parc (volume ou nb de PC eleve) ---
    html += '<div class="gc-group">Plus fort impact sur le parc <span class="gc-group-count">' + impact.length + '</span></div>';
    html += (impact.length === 0)
        ? '<div class="gc-empty">aucun signal fort</div>'
        : renderGroup(impact, 'impact');

    // --- Groupe 3 : Bruit de fond (longue traine) - repli integral par defaut ---
    html += '<div class="gc-group">Bruit de fond <span class="gc-group-count">' + noise.length + '</span></div>';
    if (noise.length === 0) {
        html += '<div class="gc-empty">aucun bruit r&eacute;siduel</div>';
    } else {
        html += '<div class="gc-more" onclick="toggleMore(this)"><span class="gc-more-tri">&#9656;</span>voir les ' + noise.length + ' appli' + (noise.length > 1 ? 's' : '') + ' (bruit de fond)</div>' +
                '<div class="gc-more-body" style="display:none">' +
                noise.map(function(c) { return renderRow(c, 'noise'); }).join('') +
                '</div>';
    }

    // v2.5 : transparence - combien de fragments de trace non exploitables masques.
    if (hiddenCount > 0) {
        html += '<div class="gc-hidden">' + hiddenCount + ' entr&eacute;e(s) de trace non exploitable(s) masqu&eacute;e(s) (bruit Dell / .NET)</div>';
    }

    container.innerHTML = html;
}

// v2.5 : expander generique (top 5 + "voir les N autres" et bruit de fond).
// Bascule le conteneur .gc-more-body qui suit immediatement la ligne cliquee.
function toggleMore(el) {
    var body = el.nextElementSibling;
    if (!body || String(body.className).indexOf('gc-more-body') === -1) return;
    var tri = el.querySelector('.gc-more-tri');
    var hidden = (body.style.display === 'none');
    body.style.display = hidden ? '' : 'none';
    if (tri) tri.innerHTML = hidden ? '&#9662;' : '&#9656;';
}

// ===== AGREGATION PARC : STATS / AUDIT CYCLE DE VIE (v2.4.2) =====
// Lit le registre COMPLET (decomRegistry, y compris PC dont le JSON a ete purge)
// -> independant des filtres de la vue. Produit : compteurs fait/a faire/en
// retard, delai moyen marque->fait, repartition par technicien et par mois.
function renderDecomStats() {
    var panel = document.getElementById('decomPanel');
    var container = document.getElementById('decomStats');
    var reg = decomRegistry || [];
    if (reg.length === 0) { panel.style.display = 'none'; return; }
    panel.style.display = '';

    var today = new Date(generatedAt);
    var todayStr = today.getFullYear() + '-' +
                   ('0' + (today.getMonth() + 1)).slice(-2) + '-' +
                   ('0' + today.getDate()).slice(-2);

    var done = 0, todo = 0, late = 0;
    var delays = [];         // jours marque->fait (entrees Fait)
    var byTech = {};         // AssignedTo -> {todo, done}
    var byMonth = {};        // YYYY-MM (DoneAt) -> count

    reg.forEach(function(e) {
        var isDone = (e.Statut === 'Fait');
        var isTodo = (e.Statut === 'A faire');
        if (isDone) done++;
        if (isTodo) {
            todo++;
            if (e.TargetDate && e.TargetDate < todayStr) late++;
        }
        var tech = e.AssignedTo || '(non assigne)';
        if (!byTech[tech]) byTech[tech] = { todo: 0, done: 0 };
        if (isDone) byTech[tech].done++;
        else if (isTodo) byTech[tech].todo++;

        if (isDone && e.MarkedAt && e.DoneAt) {
            var d1 = parseDate(e.MarkedAt), d2 = parseDate(e.DoneAt);
            if (!isNaN(d1.getTime()) && !isNaN(d2.getTime())) {
                var dd = (d2 - d1) / 86400000;
                if (dd >= 0) delays.push(dd);
            }
            byMonth[e.DoneAt.slice(0, 7)] = (byMonth[e.DoneAt.slice(0, 7)] || 0) + 1;
        }
    });

    var avgDelay = delays.length ? (delays.reduce(function(a, b) { return a + b; }, 0) / delays.length) : null;

    function tile(val, label, cls) {
        return '<div class="decom-stat-tile ' + (cls || '') + '">' +
               '<div class="decom-stat-val">' + val + '</div>' +
               '<div class="decom-stat-lbl">' + label + '</div></div>';
    }
    var tiles = '<div class="decom-stat-row">' +
        tile(done, 'D&eacute;commissionn&eacute;s (fait)', 'ok') +
        tile(todo, '&Agrave; faire', (todo > 0 ? 'warn' : '')) +
        tile(late, 'En retard', (late > 0 ? 'danger' : '')) +
        tile(avgDelay != null ? (Math.round(avgDelay) + ' j') : '&ndash;', 'D&eacute;lai moyen marqu&eacute;&rarr;fait') +
        '</div>';

    var techNames = Object.keys(byTech).sort();
    var techRows = techNames.map(function(t) {
        return '<tr><td>' + t + '</td>' +
               '<td style="text-align:right">' + byTech[t].done + '</td>' +
               '<td style="text-align:right">' + byTech[t].todo + '</td></tr>';
    }).join('');
    var techTable = '<div class="decom-sub"><h5>Par technicien</h5>' +
        '<table class="decom-table"><thead><tr><th>Technicien</th>' +
        '<th style="text-align:right">Fait</th><th style="text-align:right">&Agrave; faire</th></tr></thead>' +
        '<tbody>' + techRows + '</tbody></table></div>';

    var months = Object.keys(byMonth).sort();
    var monthsShown = months.slice(-12);
    var maxM = monthsShown.reduce(function(m, k) { return Math.max(m, byMonth[k]); }, 0);
    var monthBars = monthsShown.map(function(k) {
        var w = maxM ? Math.round(byMonth[k] / maxM * 100) : 0;
        return '<div class="decom-month"><span class="decom-month-lbl">' + k + '</span>' +
               '<div class="decom-month-bar"><div class="decom-month-fill" style="width:' + w + '%"></div></div>' +
               '<span class="decom-month-val">' + byMonth[k] + '</span></div>';
    }).join('');
    var monthBlock = monthsShown.length
        ? '<div class="decom-sub"><h5>D&eacute;commissions r&eacute;alis&eacute;es par mois</h5>' + monthBars + '</div>'
        : '';

    container.innerHTML = tiles + '<div class="decom-sub-row">' + techTable + monthBlock + '</div>';
}

// ===== AGREGATION PARC : INVENTAIRE MONITORS (v5.7) =====
// Produit les stats d'inventaire des ecrans externes :
// - Nb total d'ecrans branches
// - Nb de PC avec au moins 1 ecran externe
// - Top fabricants avec compteur
// - Ecrans anciens (>= 7 ans)
function renderMonitorInventory(pcs) {
    var panel = document.getElementById('monitorPanel');
    var container = document.getElementById('monitorInventory');

    var allMonitors = [];
    var pcWithMonitor = 0;
    pcs.forEach(function(p) {
        if (p.monitors && p.monitors.length > 0) {
            pcWithMonitor++;
            p.monitors.forEach(function(m) { allMonitors.push({ mon: m, pc: p.pc.PC }); });
        }
    });

    // Si aucun moniteur remonte, on masque le panneau (ex: parc 100% Collector v5.4)
    if (allMonitors.length === 0) {
        panel.style.display = 'none';
        return;
    }
    panel.style.display = '';

    // Top fabricants
    var manufCount = {};
    var oldCount = 0;
    var ageSum = 0;
    var ageCount = 0;
    var unidentifiedCount = 0;   // v5.8 : ecrans a EDID non transmis
    allMonitors.forEach(function(item) {
        var m = item.mon;
        // v5.8 : les ecrans non identifies ne polluent pas le top fabricants
        // (ni "@@@" ni une entree vide) ; ils sont comptes a part.
        if (!monIsIdentified(m)) { unidentifiedCount++; return; }
        var key = m.Manufacturer || m.ManufacturerCode || '?';
        manufCount[key] = (manufCount[key] || 0) + 1;
        if (m.AgeYears !== null && m.AgeYears !== undefined) {
            if (m.AgeYears >= screenAgeThreshold) oldCount++;
            ageSum += m.AgeYears;
            ageCount++;
        }
    });
    var topManufs = Object.keys(manufCount).map(function(k) {
        return { name: k, count: manufCount[k] };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 5);

    var avgAge = ageCount > 0 ? (ageSum / ageCount).toFixed(1) : 'N/A';

    // Construction du HTML
    var html = '';

    html += '<div class="monitor-inv-tile">' +
            '<div class="monitor-inv-value">' + allMonitors.length + '</div>' +
            '<div class="monitor-inv-label">&Eacute;crans secondaires</div>' +
            '<div class="monitor-inv-sub">sur ' + pcWithMonitor + ' PC du parc</div>' +
            '</div>';

    html += '<div class="monitor-inv-tile">' +
            '<div class="monitor-inv-value">' + avgAge + (avgAge !== 'N/A' ? ' ans' : '') + '</div>' +
            '<div class="monitor-inv-label">&Acirc;ge moyen</div>' +
            '<div class="monitor-inv-sub">calcul sur ' + ageCount + ' &eacute;cran(s)</div>' +
            '</div>';

    html += '<div class="monitor-inv-tile">' +
            '<div class="monitor-inv-value" style="color:' + (oldCount > 0 ? 'var(--orange)' : 'var(--text-muted)') + '">' + oldCount + '</div>' +
            '<div class="monitor-inv-label">&Eacute;crans &ge; ' + screenAgeThreshold + ' ans</div>' +
            '<div class="monitor-inv-sub">candidats au renouvellement</div>' +
            '</div>';

    // v5.8 : tuile des ecrans non identifies (EDID non transmis), si presents.
    if (unidentifiedCount > 0) {
        html += '<div class="monitor-inv-tile">' +
                '<div class="monitor-inv-value" style="color:var(--text-muted)">' + unidentifiedCount + '</div>' +
                '<div class="monitor-inv-label">Non identifi&eacute;s</div>' +
                '<div class="monitor-inv-sub">EDID non transmis (dock / adaptateur)</div>' +
                '</div>';
    }

    var manufHtml = topManufs.map(function(m) {
        return '<span class="chip">' + m.name + ' : <strong>' + m.count + '</strong></span>';
    }).join('');
    html += '<div class="monitor-inv-tile" style="grid-column:span 2">' +
            '<div class="monitor-inv-label" style="margin-top:0">Top fabricants</div>' +
            '<div class="monitor-inv-sub">' + (manufHtml || 'Aucun') + '</div>' +
            '</div>';

    container.innerHTML = html;
}

// ===== AGREGATION PARC : REPARTITION DES MODELES (v2.4.7) =====
// Compte les machines par Machine.Model (Collector >= 2.4.7). Vue "coup d'oeil"
// PURE (non cliquable) : le filtrage se fait par le menu deroulant "Modele" en
// haut. Panneau masque si aucun modele connu (parc < 2.4.7 -> retro-compat).
function renderModelInventory(pcs) {
    var panel = document.getElementById('modelPanel');
    var container = document.getElementById('modelInventory');
    if (!panel || !container) return;

    var counts = {};
    var pcWithModel = 0;
    var unknownCount = 0;
    pcs.forEach(function(p) {
        var mdl = (p.pc && p.pc.Model) ? p.pc.Model : '';
        if (mdl) { counts[mdl] = (counts[mdl] || 0) + 1; pcWithModel++; }
        else { unknownCount++; }
    });

    var sorted = Object.keys(counts).map(function(k) { return { name: k, count: counts[k] }; })
                        .sort(function(a, b) { return b.count - a.count || (a.name < b.name ? -1 : 1); });

    // Aucun modele connu (parc pre-2.4.7) -> panneau masque.
    if (sorted.length === 0) { panel.style.display = 'none'; return; }
    panel.style.display = '';

    // Vision rapide : recap en une ligne + barres classees (la 1re barre = 100%).
    var maxCount = sorted[0].count;
    var html = '<div class="model-dist-head"><strong>' + sorted.length + '</strong> mod&egrave;les distincts &middot; ' +
               pcWithModel + ' machine(s) identifi&eacute;e(s)' +
               (unknownCount > 0 ? ' &middot; ' + unknownCount + ' sans mod&egrave;le' : '') + '</div>';
    html += '<div class="model-dist-list">';
    // Vue non cliquable. m.name est deja HTML-echappe (embed), insertion sure.
    sorted.forEach(function(m) {
        var pct = pcWithModel > 0 ? Math.round(m.count * 100 / pcWithModel) : 0;
        var w   = maxCount > 0 ? Math.round(m.count * 100 / maxCount) : 0;
        html += '<div class="model-bar-row">' +
                '<span class="model-bar-name">' + m.name + '</span>' +
                '<div class="model-bar-track"><div class="model-bar-fill" style="width:' + w + '%"></div></div>' +
                '<span class="model-bar-val">' + m.count + ' &middot; ' + pct + '%</span>' +
                '</div>';
    });
    html += '</div>';
    container.innerHTML = html;
}

// ===== TABLEAU =====
function sortArrow(col) {
    if (state.sort.col !== col) return '<span class="sort-arrow">&#9650;&#9660;</span>';
    return state.sort.dir === 'asc' ? '<span class="sort-arrow">&#9650;</span>' : '<span class="sort-arrow">&#9660;</span>';
}
function sortClass(col) {
    return 'sortable' + (state.sort.col === col ? ' sort-active' : '');
}
function sortColumn(col) {
    if (state.sort.col === col) {
        state.sort.dir = state.sort.dir === 'desc' ? 'asc' : 'desc';
    } else {
        state.sort.col = col;
        state.sort.dir = (col === 'name' || col === 'site' || col === 'user' || col === 'cpu') ? 'asc' : 'desc';
    }
    render();
}

function renderTable(pcs) {
    var siteHeader = showSite ? '<th class="' + sortClass('site') + '" onclick="sortColumn(\'site\')">Site ' + sortArrow('site') + '</th>' : '';
    // v5.5 : les colonnes "techniques" (crash/bsod/hw/disque/perf) sont
    // masquees par defaut via la classe .col-advanced. Le bouton toolbar
    // "Vue detaillee" bascule le body.advanced-cols pour les reveler.
    document.getElementById('tableHead').innerHTML =
        '<tr>' +
        '<th class="' + sortClass('score') + '" onclick="sortColumn(\'score\')">Sant&eacute; ' + sortArrow('score') + '</th>' +
        '<th class="' + sortClass('name')  + '" onclick="sortColumn(\'name\')">Nom PC '  + sortArrow('name')  + '</th>' +
        '<th>Statut</th>' + siteHeader +
        '<th>IP</th>' +
        '<th class="' + sortClass('user')  + '" onclick="sortColumn(\'user\')">Utilisateur ' + sortArrow('user') + '</th>' +
        '<th>Conn.</th>' +
        '<th class="' + sortClass('cpu') + '" onclick="sortColumn(\'cpu\')">CPU ' + sortArrow('cpu') + '</th>' +
        '<th>OS</th>' +
        '<th class="' + sortClass('uptime') + '" onclick="sortColumn(\'uptime\')">Uptime ' + sortArrow('uptime') + '</th>' +
        '<th class="' + sortClass('fresh')  + '" onclick="sortColumn(\'fresh\')">Vu '  + sortArrow('fresh')  + '</th>' +
        '<th class="col-advanced ' + sortClass('crash')  + '" onclick="sortColumn(\'crash\')">Crash ' + sortArrow('crash') + '</th>' +
        '<th class="col-advanced ' + sortClass('bsod')   + '" onclick="sortColumn(\'bsod\')">BSOD '  + sortArrow('bsod')   + '</th>' +
        '<th class="col-advanced ' + sortClass('hw')     + '" onclick="sortColumn(\'hw\')">HW '      + sortArrow('hw')     + '</th>' +
        '<th class="col-advanced ' + sortClass('disk')   + '" onclick="sortColumn(\'disk\')">Disque ' + sortArrow('disk')  + '</th>' +
        '<th class="col-advanced ' + sortClass('perf')   + '" onclick="sortColumn(\'perf\')">Perf '  + sortArrow('perf')   + '</th>' +
        '<th title="Indicateurs : Batterie / EDR / BootPerf / SMART">Ind.</th>' +
        '</tr>';

    var tbody = '';
    var colspan = showSite ? 17 : 16;
    pcs.forEach(function(p, idx) {
        var pc = p.pc;
        // v2.5.0-color : la teinte de ligne est neutralisee (voir CSS). La severite
        // est desormais portee uniquement par le badge de score, sur 4 niveaux.
        var rowClass = 'row-ok';

        // v2.5.0-color : severite du badge de score (1 seul signal couleur/ligne).
        //   gris  -> hors ligne (aucune donnee a jour, prime sur tout)
        //   vert  -> sain (score 0)
        //   rouge -> critique : reprend la classif "danger" du tint precedent
        //            (crash/BSOD/materiel fatal/EDR HS/SMART)
        //   orange-> a surveiller : tout autre score > 0
        var hasDanger = (p.crashCount > 0 || p.bsodCount > 0 || p.hwCount > 0 || p.edrAlert || p.diskSmartAlert);
        var sevClass = pc.IsOffline
            ? 'score-offline'
            : (p.score === 0 ? 'score-ok' : (hasDanger ? 'score-danger' : 'score-warn'));

        var scoreCell = '<span class="score-badge ' + sevClass + '" title="Score = BSOD:' + p.bsodCount + ' Crash:' + p.crashCount + ' WHEAFatal:' + p.wheaFatal.length + ' GPU:' + p.gpuTDR.length + ' Thermal:' + p.thermal.length + ' BootLong:' + p.bootLongCount + ' Disk:' + p.diskAlertCount + (p.pc.IsOffline ? ' +Offline' : '') + (p.edrAlert ? ' +EDR' : '') + (p.batteryAlert ? ' +Batt' : '') + (p.bootPerfAlert ? ' +BootSlow' : '') + (p.diskSmartAlert ? ' +SMART' : '') + '">' + p.score + '</span>';

        var statusBadge = pc.IsOffline
            ? '<span class="badge badge-offline">OFFLINE</span>'
            : '<span class="badge badge-online">ONLINE</span>';

        var siteCell = '';
        if (showSite) {
            var siteClass = pc.Site === 'Inconnu' ? 'site-badge inconnu' : 'site-badge';
            siteCell = '<td><span class="' + siteClass + '">' + pc.Site + '</span></td>';
        }

        var uptimeCell = 'N/A';
        if (pc.UptimeDays !== null && pc.UptimeDays !== undefined) {
            // v2.5.0-color : < 7 j = neutre (sain) ; 7-30 j = orange ; > 30 j = rouge.
            var uptimeClass = 'uptime-neutral';
            if (pc.UptimeDays > 30) uptimeClass = 'uptime-danger';
            else if (pc.UptimeDays >= 7) uptimeClass = 'uptime-warning';
            uptimeCell = '<span class="uptime-badge ' + uptimeClass + '">' + formatUptime(pc.UptimeDays) + '</span>';
        }

        var freshCell = '<span class="' + freshnessClass(p.collectedHoursAgo) + '">' + timeAgo(pc.CollectedAt) + '</span>';

        var connType = pc.ConnectionType || 'Inconnu';
        // v2.5.0-color : le type de connexion n'est pas une alerte -> neutre.
        // Ethernet / WiFi / Inconnu = neutre (conn-autre). Seul "Deconnecte" reste rouge.
        var connClass = 'conn-autre';
        if (connType === 'Deconnecte') connClass = 'conn-deconnecte';
        var connCell = '<span class="conn-badge ' + connClass + '">' + connType + '</span>';

        // v2.5.3 : en mode renouvellement, le tableau affiche "A renouveler" (candidat,
        // orange) ou l'annee CPU (parc courant, neutre) au lieu de Recent/Vieillissant/
        // Ancien - l'age reel reste dans l'onglet Materiel. Sinon, tag d'age historique.
        var cpuClass, cpuLabel;
        if (RENEWAL_MODE) {
            if (isRenewalCandidate(pc)) { cpuClass = 'cpu-vieillissant'; cpuLabel = 'A renouveler'; }
            else { cpuClass = 'cpu-inconnu'; cpuLabel = pc.CPUYear ? String(pc.CPUYear) : '—'; }
        } else {
            var cpuAgeCategory = pc.CPUAgeCategory || 'Inconnu';
            // v2.5.0-color : Recent / Vieillissant / Inconnu = neutre ; seul "Ancien" en orange.
            cpuClass = 'cpu-inconnu';
            if (cpuAgeCategory === 'Ancien') cpuClass = 'cpu-vieillissant';
            cpuLabel = cpuAgeCategory;
            if (pc.CPUYear) cpuLabel += ' (' + pc.CPUYear + ')';
        }

        // v5.8 / v2.5.0-ui : icone chassis (laptop/desktop/AIO) placee EN LIGNE
        // avant le libelle CPU. On garde le glyphe et les couleurs du badge
        // chassis existant, mais on retire le texte (repris en tooltip) pour
        // rester sur une seule ligne.
        var chassisIcon = '';
        if (pc.ChassisInfo && pc.ChassisInfo.ChassisLabel && pc.ChassisInfo.ChassisLabel !== 'Inconnu') {
            var chCls = 'chassis-autre', chIcon = '&#128221;';
            if (pc.ChassisInfo.IsLaptop)       { chCls = 'chassis-laptop';  chIcon = '&#128187;'; }
            else if (pc.ChassisInfo.IsAIO)     { chCls = 'chassis-aio';     chIcon = '&#128444;'; }
            else if (pc.ChassisInfo.IsDesktop) { chCls = 'chassis-desktop'; chIcon = '&#128421;'; }
            chassisIcon = '<span class="chassis-icon ' + chCls + '" title="' + pc.ChassisInfo.ChassisLabel + '">' + chIcon + '</span>';
        }
        var cpuCell = chassisIcon + '<span class="cpu-badge ' + cpuClass + '" title="' + (pc.CPUName || '') + '">' + cpuLabel + '</span>';

        // v2.2.1 : OS (Windows 10/11) - inventaire migration / EOL.
        var osProduct = pc.OSProduct || 'Inconnu';
        // v2.5.0-color : Windows 11 / Inconnu = neutre (os-inconnu).
        // Windows 10 reste en orange (os-win10) : fin de support = vraie alerte.
        var osCls = 'os-inconnu';
        if (osProduct === 'Windows 10') osCls = 'os-win10';
        var osTitle = 'Build ' + (pc.OSBuild || '?') + (pc.OSEdition ? ' - ' + pc.OSEdition : '');
        var osCell = '<span class="os-badge ' + osCls + '" title="' + osTitle + '">' + osProduct + '</span>';
        if (pc.OSDisplayVersion) {
            // v2.5.0-ui : build affiche EN LIGNE (mono-ligne), plus petit et attenue.
            osCell += '<span class="os-build">&middot; ' + pc.OSDisplayVersion + '</span>';
        }

        var diskCell = '';
        if (p.diskInfo.length > 0) {
            var worstDisk = p.diskInfo.reduce(function(a, b) { return a.PctFree < b.PctFree ? a : b; });
            var diskColor = 'color-green';
            if (worstDisk.IsAlert) diskColor = 'color-red';
            else if (worstDisk.PctFree < seuilDiskWarning) diskColor = 'color-orange';
            diskCell = '<span class="' + diskColor + '" style="font-weight:700">' + worstDisk.PctFree + '% libre</span><div style="font-size:10px;color:var(--text-faint)">' + worstDisk.FreeGB + ' GB</div>';
        } else {
            diskCell = '<span style="color:var(--text-ghost)">N/A</span>';
        }

        var hwCell = p.hwCount > 0
            ? '<span class="kpi-hardware">' + p.hwCount + '</span>'
            : '<span style="color:var(--text-ghost)">-</span>';

        // ==============================================================
        // COLONNE INDICATEURS (v5.3 / v5.4)
        // 4 pastilles compactes : Batt / EDR / Boot / SMART
        // Etats : ok (vert) / warn (orange) / ko (rouge) / na (gris)
        // ==============================================================
        var indB, indE, indBP, indS;
        // Batterie
        if (!p.battery || !p.battery.HasBattery) {
            indB = '<span class="indicator-dot na" title="Pas de batterie (desktop/VM)">B</span>';
        } else if (p.batteryAlert) {
            indB = '<span class="indicator-dot ko" title="Batterie us&eacute;e : ' + p.battery.HealthPercent + '% (' + p.battery.HealthCategory + ')">B</span>';
        } else {
            indB = '<span class="indicator-dot ok" title="Batterie OK : ' + p.battery.HealthPercent + '% (' + p.battery.HealthCategory + ')">B</span>';
        }
        // EDR (service Role='EDR', pilote par config)
        if (!p.edr) {
            indE = '<span class="indicator-dot na" title="Aucune donn&eacute;e EDR (poste en cours de MAJ Collector, ou aucun service EDR configur&eacute;)">E</span>';
        } else if (!p.edr.Installed) {
            indE = '<span class="indicator-dot ko" title="' + p.edr.DisplayName + ' NON INSTALLE">E</span>';
        } else if (p.edrAlert) {
            indE = '<span class="indicator-dot ko" title="' + p.edr.DisplayName + ' : ' + p.edr.Status + '">E</span>';
        } else {
            indE = '<span class="indicator-dot ok" title="' + p.edr.DisplayName + ' : Running">E</span>';
        }
        // Boot Performance
        if (!p.bootPerf || !p.bootPerf.LastBoot) {
            indBP = '<span class="indicator-dot na" title="Pas de donn&eacute;e boot perf (aucun cold boot r&eacute;cent ou Collector &lt; v5.4)">P</span>';
        } else if (p.bootPerfAlert) {
            var lb = p.bootPerf.LastBoot;
            indBP = '<span class="indicator-dot ko" title="Boot lent : ' + (lb.BootTimeMs/1000).toFixed(1) + 's (Post-boot ' + (lb.BootPostBootTimeMs/1000).toFixed(1) + 's)">P</span>';
        } else {
            indBP = '<span class="indicator-dot ok" title="Boot OK : ' + (p.bootPerf.LastBoot.BootTimeMs/1000).toFixed(1) + 's">P</span>';
        }
        // SMART
        if (p.diskHealth.length === 0) {
            indS = '<span class="indicator-dot na" title="Donn&eacute;es SMART absentes (Collector &lt; v5.4)">S</span>';
        } else if (p.diskSmartAlert) {
            var reasons = [];
            p.diskSmartAlerts.forEach(function(d) { reasons.push(d.FriendlyName + ' (' + d.AlertReasons.join(', ') + ')'); });
            indS = '<span class="indicator-dot ko" title="SMART alerte : ' + reasons.join(' | ') + '">S</span>';
        } else {
            indS = '<span class="indicator-dot ok" title="SMART OK (' + p.diskHealth.length + ' disque' + (p.diskHealth.length > 1 ? 's' : '') + ')">S</span>';
        }
        var indicatorsCell = '<div class="indicator-row">' + indB + indE + indBP + indS + '</div>';

        // v2.1.11 : badge anomalie (donnee a verifier) - cliquable pour filtrer,
        // detail au survol. stopPropagation pour ne pas ouvrir le drill-down au clic.
        var anomalyBadge = '';
        if (p.anomalies && p.anomalies.length > 0) {
            var anomTitle = p.anomalies.map(function(a) {
                return a.Reason + (a.Detail ? ' - ' + a.Detail : '');
            }).join(' | ');
            anomalyBadge = ' <span class="anomaly-badge" title="' + anomTitle +
                           '" onclick="event.stopPropagation(); toggleKpiFilter(\'anomaly\')">&#9888;</span>';
        }

        // v2.4.0 : badge decommission (cycle de vie). Cliquable -> filtre.
        var decomBadge = '';
        if (p.decom) {
            if (p.decomTodo) {
                var dLeft = (p.decomDaysLeft === null) ? '' :
                    (p.decomDaysLeft < 0 ? ' (' + Math.abs(p.decomDaysLeft) + 'j de retard)' : ' (J-' + p.decomDaysLeft + ')');
                var dTitle = 'A decommissionner - ' + (p.decom.Reason || '') +
                             ' | assigne: ' + (p.decom.AssignedTo || '?') +
                             ' | echeance: ' + (p.decom.TargetDate || '?') + dLeft;
                decomBadge = ' <span class="decom-badge' + (p.decomLate ? ' late' : '') + '" title="' + dTitle +
                             '" onclick="event.stopPropagation(); toggleKpiFilter(\'' + (p.decomLate ? 'decomLate' : 'decomTodo') + '\')">&#9851;</span>';
            } else if (p.decomDoneOnline) {
                decomBadge = ' <span class="decom-badge done-online" title="Marque FAIT mais le poste remonte encore - a verifier (fait par ' + (p.decom.DoneBy || '?') +
                             ')" onclick="event.stopPropagation(); toggleKpiFilter(\'decomOnline\')">&#9851;!</span>';
            } else if (p.decomDone) {
                decomBadge = ' <span class="decom-badge done" title="Decommission faite le ' + (p.decom.DoneAt || '?') + ' par ' + (p.decom.DoneBy || '?') + '">&#9851;</span>';
            }
        }

        tbody += '<tr class="' + rowClass + ' row-main" onclick="toggleDetail(\'detail-' + idx + '\')">' +
            '<td>' + scoreCell + '</td>' +
            '<td><span class="toggle-icon">&#9654;</span>' + pc.PC + anomalyBadge + decomBadge + '</td>' +
            '<td>' + statusBadge + '</td>' +
            siteCell +
            '<td>' + pc.IP + '</td>' +
            '<td>' + userDisplay(pc) + '</td>' +
            '<td>' + connCell + '</td>' +
            '<td>' + cpuCell + '</td>' +
            '<td>' + osCell + '</td>' +
            '<td>' + uptimeCell + '</td>' +
            '<td>' + freshCell + '</td>' +
            '<td class="col-advanced"><span class="kpi-crash">' + p.crashCount + '</span>' + (p.bsodClassifieCount > 0 ? ' <span class="kpi-bsod" style="font-size:11px">(' + p.bsodClassifieCount + ' BSOD)</span>' : '') + '</td>' +
            '<td class="col-advanced"><span class="kpi-bsod">' + p.bsodCount + '</span></td>' +
            '<td class="col-advanced">' + hwCell + '</td>' +
            '<td class="col-advanced">' + diskCell + '</td>' +
            '<td class="col-advanced">' + (p.warningCount > 0 ? '<span class="kpi-warning">' + p.warningCount + '</span>' : '-') + '</td>' +
            '<td>' + indicatorsCell + '</td>' +
            '</tr>';

        tbody += renderDetailRow(p, idx, colspan);
    });

    if (pcs.length === 0) {
        tbody = '<tr><td colspan="' + colspan + '" style="text-align:center;padding:30px;color:var(--text-ghost);font-style:italic">Aucun PC ne correspond aux filtres actifs</td></tr>';
    }
    document.getElementById('tableBody').innerHTML = tbody;
}

function renderDetailRow(p, idx, colspan) {
    // v2.5.0 (Essai A + option 3) : drill-down = liste de definitions mono-ligne
    // unifiee (une donnee = une ligne : gouttiere de libelle + valeur condensee).
    // Les sections riches sont condensees en une ligne composite ; les listes = une
    // ligne par item. Les champs N/A / vides / rares-par-defaut sont OMIS (option 3) ;
    // les champs rares mais renseignes sont replies dans la ligne. Aucune donnee
    // metier n'est modifiee : ce n'est qu'un reformatage de l'affichage.

    // --- petites fabriques de balisage (aucun contenu HTML n'est re-echappe :
    //     les valeurs proviennent deja echappees a l'embed du JSON) ---
    function ddG(name) { return '<div class="dd-group">' + name + '</div>'; }
    function ddR(label, val, vcls) {
        if (val === null || val === undefined || val === '') return '';
        return '<div class="dd-row"><span class="dd-l">' + label + '</span>' +
               '<span class="dd-v' + (vcls ? ' ' + vcls : '') + '">' + val + '</span></div>';
    }
    function mut(t)  { return '<span class="mut">' + t + '</span>'; }
    function mono(t) { return '<span class="mono">' + t + '</span>'; }
    function col(t, cls) { return cls ? '<span class="' + cls + '">' + t + '</span>' : t; }
    function tag(t, cls) { return '<span class="dd-tag' + (cls ? ' ' + cls : '') + '">' + t + '</span>'; }
    // secondaire muet, precede d'un separateur point median ; ignore vides/null
    function joinMut(parts) {
        var a = parts.filter(function(x) { return x !== null && x !== undefined && x !== ''; });
        return a.length ? ' ' + mut('&middot; ' + a.join(' &middot; ')) : '';
    }
    // v1.6 : libelles fins par CrashCause pour les Events 41 (fallback anciens JSON)
    function crashCauseLabel(cause) {
        if (cause === 'BSODSilent')        return 'BSOD silencieux';
        if (cause === 'SleepResumeFailed') return 'Reprise veille rat&eacute;e';
        if (cause === 'UserForcedReset')   return 'User bouton power';
        if (cause === 'PowerLoss')         return 'Coupure alim / thermal';
        if (cause === 'FreezeApp')         return 'Freeze applicatif';
        if (cause === 'FreezeUnknown')     return 'Freeze (cause inconnue)';
        return null;
    }

    // ========================================================
    // MATERIEL : groupes "Identite" (CPU / OS / Machine) puis "Materiel".
    // ========================================================
    // CPU : modele complet . annee . age  + tag categorie d'age
    var cpuVal = '';
    if (p.pc.CPUName) {
        var cpuSec = [];
        if (p.pc.CPUYear) cpuSec.push(p.pc.CPUYear);
        if (p.pc.CPUAge !== null && p.pc.CPUAge !== undefined) cpuSec.push(p.pc.CPUAge + ' ans');
        cpuVal = p.pc.CPUName + joinMut(cpuSec);
        // v2.5.3 : l'onglet Materiel garde TOUJOURS l'annee + l'age (cpuSec ci-dessus).
        // Le badge devient Candidat/Parc courant en mode renouvellement, sinon age.
        if (RENEWAL_MODE) {
            cpuVal += isRenewalCandidate(p.pc) ? tag('A renouveler', 'warn') : tag('Parc courant', 'ok');
        } else {
            var cpuCat = p.pc.CPUAgeCategory || '';
            if (cpuCat && cpuCat !== 'Inconnu') {
                var cpuTagCls = (cpuCat === 'Ancien') ? 'ko' : ((cpuCat === 'Vieillissant') ? 'warn' : 'ok');
                cpuVal += tag(cpuCat, cpuTagCls);
            }
        }
    }
    // OS : badge produit + version . edition . build
    var dOsProduct = p.pc.OSProduct || 'Inconnu';
    var dOsCls = (dOsProduct === 'Windows 11') ? 'os-win11' : ((dOsProduct === 'Windows 10') ? 'os-win10' : 'os-inconnu');
    var osSec = [];
    if (p.pc.OSDisplayVersion) osSec.push(p.pc.OSDisplayVersion);
    if (p.pc.OSEdition)        osSec.push(p.pc.OSEdition);
    if (p.pc.OSBuild)          osSec.push('build ' + p.pc.OSBuild);
    var osVal = '<span class="os-badge ' + dOsCls + '">' + dOsProduct + '</span>' + joinMut(osSec);
    // Machine : fabricant modele . S/N (mono)
    var dModel = [p.pc.Manufacturer || '', p.pc.Model || ''].filter(function(x) { return x; }).join(' ');
    var machVal = dModel || '';
    if (p.pc.SerialNumber) machVal += (machVal ? ' ' + mut('&middot; S/N ') : mut('S/N ')) + mono(p.pc.SerialNumber);

    // GPU : par adaptateur "Nom vX.Y" joints par point median
    var gpuList = p.gpuInventory || [];
    var gpuVal = gpuList.length ? gpuList.map(function(g) {
        var nm = g.Name || '(GPU inconnu)';
        return nm + (g.DriverVersion ? ' ' + mut('v' + g.DriverVersion) : '');
    }).join(' &middot; ') : '';

    // RAM : installee / max . slots libres . barrettes inline
    var ramVal = '';
    if (p.memory) {
        var mem = p.memory;
        ramVal = mem.TotalInstalledGB + ' / ' + mem.MaxCapacityGB + ' Go';
        var ramSec = [];
        if (mem.FreeSlots > 0)   ramSec.push(mem.FreeSlots + ' slot(s) libre(s)');
        else if (mem.TotalSlots) ramSec.push('0 slot libre');
        if (mem.Modules && mem.Modules.length) {
            // v2.5.0 : on regroupe les barrettes identiques (frequent sur RAM
            // soudee : plusieurs modules "Motherboard" strictement identiques ->
            // "4 x 8 Go LPDDR5 8000 MHz" au lieu de 4 lignes brutes repetees).
            // Le slot (DIMMA/Motherboard) est volontairement omis : peu utile ici
            // et il casse le regroupement.
            var seen = {}, order = [];
            mem.Modules.forEach(function(mod) {
                var mp = [mod.CapacityGB + ' Go'];
                var ty = cleanRamType(mod.Type);      if (ty) mp.push(ty);
                if (mod.SpeedMHz > 0)                 mp.push(mod.SpeedMHz + ' MHz');
                var rm = prettyRamManuf(mod.Manufacturer); if (rm) mp.push(rm);
                var key = mp.join(' ');
                if (seen[key] === undefined) { seen[key] = 0; order.push(key); }
                seen[key]++;
            });
            var mods = order.map(function(k) {
                return (seen[k] > 1 ? seen[k] + ' &times; ' : '') + k;
            }).join(', ');
            ramSec.push(mods);
        }
        ramVal += joinMut(ramSec);
    }

    // Disque : une ligne par volume, valeur coloree par seuil de remplissage
    var diskRows = '';
    (p.diskInfo || []).forEach(function(d) {
        var valCls = 'ok';
        if (d.IsAlert) valCls = 'ko';
        else if (d.PctFree < seuilDiskWarning) valCls = 'warn';
        diskRows += ddR('Disque ' + d.Drive, col(d.PctUsed + '%', valCls) + ' ' + mut('(' + d.FreeGB + ' Go libres)'));
    });

    // SMART : une ligne par disque physique (option 3 sur les champs rares)
    var smartRows = '';
    (p.diskHealth || []).forEach(function(d) {
        var hl = d.HealthStatus || 'Unknown';
        var hcls = (hl === 'Warning') ? 'warn' : ((hl !== 'Healthy') ? 'ko' : 'ok');
        var tempCls = '';
        if (d.TemperatureC != null) tempCls = (d.TemperatureC >= 70) ? 'ko' : ((d.TemperatureC >= 55) ? 'warn' : 'ok');
        var wearCls = '';
        if (d.WearPct != null) wearCls = (d.WearPct >= 70) ? 'ko' : ((d.WearPct >= 40) ? 'warn' : 'ok');
        var seg = [];
        var ctx = [];
        if (d.MediaType) ctx.push(d.MediaType);
        if (d.SizeGB)    ctx.push(d.SizeGB + ' Go');
        if (ctx.length) seg.push(mut(ctx.join(' ')));
        if (d.TemperatureC != null) seg.push(col(d.TemperatureC + '&deg;C', tempCls));
        if (d.WearPct != null)      seg.push('usure ' + col(d.WearPct + '%', wearCls));
        if (d.PowerOnHours != null) seg.push(d.PowerOnHours + ' h');
        seg.push('err ' + (d.ReadErrorsTotal || 0) + '/' + (d.WriteErrorsTotal || 0));
        // option 3 : champs rares uniquement si presents et hors valeur par defaut
        if (d.TemperatureMaxC != null) seg.push('max ' + d.TemperatureMaxC + '&deg;C');
        if (d.ReadErrorsUncorrected != null && d.ReadErrorsUncorrected > 0)  seg.push(col('lect.NC ' + d.ReadErrorsUncorrected, 'ko'));
        if (d.WriteErrorsUncorrected != null && d.WriteErrorsUncorrected > 0) seg.push(col('&eacute;cr.NC ' + d.WriteErrorsUncorrected, 'ko'));
        var sv = tag(hl, hcls) + ' ' + seg.join(' &middot; ');
        if (d.AlertReasons && d.AlertReasons.length) {
            sv += ' ' + d.AlertReasons.map(function(r) { return tag(r, 'ko'); }).join('');
        }
        smartRows += ddR(d.FriendlyName || 'Disque', sv);
    });

    // Batterie : sante . charge . capacite actuelle/design (+ cycles). Omise si absente.
    var battVal = '';
    if (p.battery && p.battery.HasBattery) {
        var b = p.battery;
        var bh = b.HealthPercent || 0;
        var bhcls = (bh < 60) ? 'ko' : ((bh < 80) ? 'warn' : 'ok');
        battVal = col(bh + '%', bhcls);
        if (bh < 80 && b.HealthCategory) battVal += tag(b.HealthCategory, bhcls);
        var bSec = [];
        if (b.CurrentChargePct != null) bSec.push(b.CurrentChargePct + '% (' + b.Status + ')');
        if (b.FullChargeCapacity || b.DesignCapacity) {
            var aw = b.FullChargeCapacity ? Math.round(b.FullChargeCapacity / 1000) + ' Wh' : '?';
            var dw = b.DesignCapacity ? Math.round(b.DesignCapacity / 1000) + ' Wh' : '?';
            bSec.push(aw + ' / ' + dw);
        }
        if (b.CycleCount != null) bSec.push(b.CycleCount + ' cycles');
        battVal += joinMut(bSec);
    }

    // Ecran(s) : une ligne par moniteur identifie (+ note si non identifie)
    var monRows = '';
    (p.monitors || []).forEach(function(mon) {
        var actTag = mon.Active ? tag('Actif', 'ok') : tag('D&eacute;branch&eacute;', '');
        if (!monIsIdentified(mon)) {
            monRows += ddR('&Eacute;cran', mut('non identifi&eacute; (EDID non transmis)') + actTag);
            return;
        }
        var name = mon.Model || mon.ProductCode || '(mod&egrave;le inconnu)';
        var mv = name + actTag;
        var isOld = (mon.AgeYears != null && mon.AgeYears >= screenAgeThreshold);
        if (isOld) mv += tag(mon.AgeYears + ' ans', 'warn');
        var mSec = [];
        var manuf = mon.Manufacturer || mon.ManufacturerCode;
        if (manuf) mSec.push(manuf);
        if (mon.YearOfManufacture) mSec.push(mon.YearOfManufacture);
        if (mon.SerialNumber) mSec.push('S/N ' + mon.SerialNumber);
        mv += joinMut(mSec);
        monRows += ddR('&Eacute;cran', mv);
    });

    // Throttling : "aucun detecte" (muet) ou synthese + journees
    var throttleList = (p.hardwareHealth && p.hardwareHealth.CPUThrottling) ? p.hardwareHealth.CPUThrottling : [];
    var thrVal, thrRows = '';
    if (!throttleList.length) {
        thrVal = mut('aucun d&eacute;tect&eacute;');
    } else {
        var thDays = throttleList.length;
        thrVal = (thDays >= 3)
            ? col('&#9888;&#65039; ' + thDays + ' jours de bridage firmware', 'warn')
            : thDays + ' jour(s) de bridage firmware';
        throttleList.slice(0, 5).forEach(function(t) {
            var typeShort = t.EventId === 55 ? 'Thermal reduction' : 'Firmware limit';
            var dv = fmtDuration(t.TotalSeconds);
            thrRows += ddR('Bridage', mono(t.Day) + ' ' + mut(typeShort + (dv ? ' &middot; cumul ' + dv : '') + ' &middot; &times;' + t.Count));
        });
        if (throttleList.length > 5) thrRows += ddR('Bridage', mut('+ ' + (throttleList.length - 5) + ' journ&eacute;es plus anciennes'));
    }

    var matContent =
        ddG('Identit&eacute;') +
        ddR('CPU', cpuVal) + ddR('OS', osVal) + ddR('Machine', machVal) +
        ddG('Mat&eacute;riel') +
        ddR('GPU', gpuVal) + ddR('RAM', ramVal) + diskRows + smartRows +
        ddR('Batterie', battVal) + monRows + ddR('Throttling', thrVal) + thrRows;

    // ========================================================
    // STABILITE : Crash/Freeze, erreurs materielles, top crashers, performance.
    // ========================================================
    var crashRows = '';
    if (p.crashes.length > 0) {
        // Replie les libelles d'evenement identiques consecutifs (ex. runs de
        // "Hard reset" / "BSOD") : le type n'est ecrit que sur la 1re ligne du
        // run, les occurrences suivantes gardent une gouttiere de libelle vide.
        // Une ligne par (type + detail) : les occurrences identiques (ex. runs de
        // "Hard reset - Coupure alim/thermal") sont regroupees, dates a la suite.
        var crashGroups = [];
        p.crashes.forEach(function(c) {
            var causeLbl = crashCauseLabel(c.CrashCause);
            var when = frDate(c.Timestamp) + ' (' + timeAgo(c.Timestamp) + ')';
            var label, detail;
            if (c.Type === 'BSOD') {
                label = 'BSOD';
                var bn = bugCheckName(c.Detail);
                detail = col(c.Detail || '', 'ko') + (bn ? ' ' + bn : '');
            } else if (c.Type === 'Hard reset') {
                label = 'Hard reset'; detail = (causeLbl ? causeLbl : '');
            } else if (c.Detail) {
                label = 'Freeze'; detail = c.Detail;
            } else {
                label = 'Freeze'; detail = (causeLbl ? causeLbl : '');
            }
            var last = crashGroups[crashGroups.length - 1];
            if (last && last.label === label && last.detail === detail) { last.whens.push(when); }
            else { crashGroups.push({ label: label, detail: detail, whens: [when] }); }
        });
        crashGroups.forEach(function(g) {
            crashRows += ddR(g.label, mut(g.whens.join(' &middot; ')) + (g.detail ? ' &middot; ' + g.detail : ''));
        });
    } else {
        crashRows = ddR('Crash / Freeze', mut('aucun sur la p&eacute;riode'));
    }

    // Piste memoire (correlation crash memoire + WHEA RAM), si active
    var pisteRow = '';
    if (p.memoryPiste && p.memoryPiste.active) {
        var mp = p.memoryPiste;
        var crashNames = mp.crashers.map(function(c) { return c.AppName + ' (' + (exceptionLabel(c.ExceptionCode) || c.ExceptionCode) + ')'; }).join(', ');
        var ramCount = mp.ramFatal.length + mp.ramCorrected.reduce(function(s, c) { return s + (c.Count || 1); }, 0);
        pisteRow = ddR('Piste', col('m&eacute;moire &agrave; v&eacute;rifier', 'warn') + ' ' +
            mut('plantages m&eacute;moire (' + crashNames + ') + WHEA RAM &times;' + ramCount + ' &middot; envisager un memtest'));
    }

    // Erreurs materielles : fatales (rouge) puis corrigees (telemetrie, muet)
    var hwRows = '';
    p.wheaFatal.forEach(function(h) {
        var detail = h.ErrorSource || '';
        if (h.BDF) detail += (detail ? ' &middot; ' : '') + h.BDF;
        hwRows += ddR(h.Component || 'HW', col(detail || (h.Component || '?'), 'ko') + ' ' + mut(frDate(h.Timestamp)));
    });
    p.gpuTDR.forEach(function(h) {
        hwRows += ddR('GPU', col('TDR ' + (h.Driver || ''), 'ko') + ' ' + mut(frDate(h.Timestamp)));
    });
    p.thermal.forEach(function(h) {
        var ti = (h.Temperature || '') + (h.Zone ? ' &middot; ' + h.Zone : '');
        hwRows += ddR('Thermal', col(h.AlertType, 'ko') + (ti ? ' ' + ti : '') + ' ' + mut(frDate(h.Timestamp)));
    });
    p.wheaCorrected.slice(0, 5).forEach(function(h) {
        var detail = h.ErrorSource || '';
        if (h.BDF) detail += (detail ? ' &middot; ' : '') + h.BDF;
        hwRows += ddR(h.Component || 'HW', mut((detail || 'corrig&eacute;e') + ' &times;' + h.Count + ' &middot; ') + frDate(h.LastSeen));
    });
    if (p.wheaCorrected.length > 5) hwRows += ddR('HW', mut('... et ' + (p.wheaCorrected.length - 5) + ' signature(s)'));
    if (!hwRows) hwRows = ddR('Hardware', mut('aucune erreur mat&eacute;rielle'));

    // Top crashers : label = appli, valeur = "<n> plantes . cause . origine"
    // v2.5.0 : meme filtre de bruit que le panneau global (isRealApp) + retrait du
    // nom d'utilisateur happe par erreur. Avant, le drill-down affichait tout brut
    // (traces Dell DDPM, un nom d'utilisateur, exceptions .NET...). Une appli suivie reste.
    var crasherRows = '';
    var crasherHidden = 0;
    var pcCrashers = (p.topCrashers || []).filter(function(tc) {
        if (isBlacklistedHard(tc.AppName)) { return false; }   // deja retire silencieusement (systeme)
        if (isPriorityApp(tc.AppName)) return true;
        if (isUserArtifact(tc.AppName, p.pc) || !isRealApp(tc.AppName)) { crasherHidden++; return false; }
        return true;
    });
    if (pcCrashers.length > 0) {
        pcCrashers.slice(0, 5).forEach(function(tc) {
            var v = tc.CrashCount + ' plant&eacute;' + (tc.CrashCount > 1 ? 's' : '');
            var extra = [];
            var code = exceptionLabel(tc.ExceptionCode);
            if (code) extra.push(code);
            var orig = crasherOrigin(tc.FaultModule);
            if (orig) extra.push(orig);
            v += joinMut(extra);
            if (tc.Type === 'app_failure') v += tag('&eacute;chec r&eacute;current', 'warn');
            else if (isBlacklistedSoft(tc.AppName)) v += tag('bruit', '');
            crasherRows += ddR(tc.AppName, v);
        });
        if (crasherHidden > 0) crasherRows += ddR('', mut(crasherHidden + ' entr&eacute;e(s) de trace non exploitable(s) masqu&eacute;e(s) (bruit de log)'));
    } else if (crasherHidden > 0) {
        crasherRows = ddR('Top crashers', mut('aucun crash app exploitable (' + crasherHidden + ' entr&eacute;e(s) de bruit masqu&eacute;e(s))'));
    } else {
        crasherRows = ddR('Top crashers', mut('aucun crash app'));
    }

    // Performance : warnings (RAM/CPU/Disque/IO) puis top consommateurs RAM
    var perfRows = '';
    p.warnings.forEach(function(w) {
        var bl = (w.Type === 'RAM exhaustion') ? 'RAM'
               : (w.Type === 'CPU throttling') ? 'CPU'
               : (w.Type === 'Disk full') ? 'Disque'
               : (w.Type === 'Disk slow') ? 'I/O' : 'Perf';
        var detailStr = w.Detail || '';
        var countStr = '';
        if (typeof w.Count === 'number' && w.Count > 1) {
            var durSec = 0;
            if (w.FirstSeen && w.LastSeen) { try { durSec = Math.round((parseDate(w.LastSeen) - parseDate(w.FirstSeen)) / 1000); } catch (e) {} }
            countStr = ' &times;' + w.Count + ' events';
            if (durSec <= 1) countStr += ' en 1s';
            else if (durSec < 60) countStr += ' en ' + durSec + 's';
            else countStr += ' en ' + Math.round(durSec / 60) + 'min';
        }
        var burst = (w.IsBurst === true) ? tag('&#128293; BURST', 'ko') : '';
        perfRows += ddR(bl, col(detailStr + countStr, 'warn') + burst + ' ' + mut(frDate(w.Timestamp)));
    });
    p.topRAM.forEach(function(proc) {
        var mb = proc.WorkingSetMB;
        var mcls = (mb >= 2000) ? 'ko' : ((mb >= 1000) ? 'warn' : '');
        perfRows += ddR(proc.Name, col(mb + ' MB', mcls));
    });
    if (!perfRows) perfRows = ddR('Performance', mut('aucun warning'));

    var stabContent =
        ddG('Crash / Freeze') + crashRows +
        ddG('Erreurs mat&eacute;rielles') + pisteRow + hwRows +
        ddG('Top crashers') + crasherRows +
        ddG('Performance') + perfRows;

    // ========================================================
    // DEMARRAGE : boots recents + boot performance.
    // ========================================================
    var bootRows = '';
    if (p.boots.length > 0) {
        // v2.5.0 : on classe chaque demarrage par CAUSE plutot que par type kernel.
        // Un "Cold x5" brut alarmait a tort (cf. cycle de MAJ Windows qui enchaine
        // plusieurs reboots en quelques minutes). On correle chaque boot :
        //   - reprise apres plantage  = un Event 41 (BSOD) tombe a +/- 3 min du boot
        //   - mise a jour Windows     = un Event 1074 "update" juste avant le boot
        //   - demarrage normal        = tout le reste (allumage, redemarrage manuel)
        var rb = p.reboots || [], cr = p.crashes || [];
        // Un Event 41 est journalise au 1er boot qui suit l'arret sale : son
        // horodatage colle donc au boot (+/- 90 s). Fenetre volontairement serree
        // pour ne pas happer un reboot de MAJ survenu 2-3 min apres le plantage.
        function nearCrash(tb) {
            for (var i = 0; i < cr.length; i++) {
                if (Math.abs((tb - parseDate(cr[i].Timestamp)) / 1000) <= 90) return true;
            }
            return false;
        }
        function nearUpdate(tb) {
            for (var i = 0; i < rb.length; i++) {
                if (!rb[i].IsUpdate) continue;
                var dt = (tb - parseDate(rb[i].Timestamp)) / 1000;   // boot apres le 1074
                if (dt >= -60 && dt <= 360) return true;
            }
            return false;
        }
        var grpMaj = [], grpCrash = [], grpNorm = [];   // grpNorm garde les DateBoot bruts
        p.boots.slice().reverse().forEach(function(bo) {   // plus recent -> plus ancien
            var tb = parseDate(bo.DateBoot);
            if (nearCrash(tb))       grpCrash.push(frDate(bo.DateBoot, true) + (bo.EstBootLong ? ' (long)' : ''));
            else if (nearUpdate(tb)) grpMaj.push(frDate(bo.DateBoot, true) + (bo.EstBootLong ? ' (long)' : ''));
            else                     grpNorm.push(bo.DateBoot);
        });
        // Listes datees pour ce qui compte une par une (MAJ, plantages).
        var listRow = function(label, arr) {
            if (!arr.length) return '';
            var shown = arr.slice(0, 12).join(' &middot; ');
            if (arr.length > 12) shown += ' &middot; +' + (arr.length - 12);
            return ddR(label, mut(shown));
        };
        bootRows += listRow('Mises &agrave; jour Windows' + (grpMaj.length > 1 ? ' (cycle)' : ''), grpMaj);
        bootRows += listRow('Reprises apr&egrave;s plantage', grpCrash);
        // v2.5.0 : demarrages normaux -> listes courtes en clair, listes longues
        // (rebooteur quotidien) resumees. Sinon 30 dates a la chaine = illisible.
        if (grpNorm.length) {
            if (grpNorm.length <= 6) {
                bootRows += ddR('D&eacute;marrages', mut(grpNorm.map(function(d) { return frDate(d, true); }).join(' &middot; ')));
            } else {
                var recent = grpNorm[0], oldest = grpNorm[grpNorm.length - 1];
                bootRows += ddR('D&eacute;marrages', mut(
                    '<b class="dd-strong">' + grpNorm.length + '</b> d&eacute;marrages &middot; du ' +
                    frDate(oldest, true) + ' au ' + frDate(recent, true)));
            }
        }
    } else {
        bootRows += ddR('D&eacute;marrage', mut('aucun d&eacute;marrage d&eacute;tect&eacute;'));
    }

    var bpRows = '';
    if (!p.bootPerf || !p.bootPerf.LastBoot) {
        bpRows = ddR('Boot perf', mut('aucune donn&eacute;e (aucun cold boot r&eacute;cent)'));
    } else {
        var lb = p.bootPerf.LastBoot;
        var secs = function(ms) { return (ms / 1000).toFixed(1) + ' s'; };
        var phaseRow = function(label, ms, slow) {
            var cls = (slow && ms > slow) ? 'warn' : ((ms < 10000) ? 'ok' : '');
            return ddR(label, col(secs(ms), cls));
        };
        // v2.5.0 : libelles friendly (les noms Microsoft MainPath/PostBoot etc.
        // ne parlaient pas aux techs).
        bpRows += phaseRow('Syst&egrave;me &amp; services', lb.MainPathBootTimeMs, 60000);
        bpRows += phaseRow('Finition en arri&egrave;re-plan', lb.BootPostBootTimeMs, 90000);
        bpRows += phaseRow('Chargement du profil', lb.UserProfileProcessingTimeMs, 10000);
        bpRows += phaseRow('Affichage du bureau', lb.ExplorerInitTimeMs, 15000);
        var summ = secs(lb.BootTimeMs) + joinMut([
            lb.NumStartupApps + ' apps',
            'niveau ' + (lb.Level || '?'),
            (p.bootPerf.Stats ? (p.bootPerf.Stats.SlowBootsCount + '/' + p.bootPerf.Stats.BootsAnalyzed + ' lents') : '')
        ]);
        if (lb.IsRebootAfterInstall) summ += tag('post-MAJ', 'warn');
        bpRows += ddR('Total', summ);
        if (p.bootPerf.History && p.bootPerf.History.length > 1) {
            // Toutes ces lignes portent le meme libelle "Boot" : on l'ecrit une
            // seule fois (1re ligne de l'historique), les suivantes s'empilent
            // avec une gouttiere de libelle vide.
            p.bootPerf.History.forEach(function(h, i) {
                if (i === 0) return;
                var histLabel = (i === 1) ? 'Boot' : '';
                bpRows += ddR(histLabel, mut(frDate(h.Timestamp, true)) + ' Syst&egrave;me ' + secs(h.MainPathBootTimeMs) +
                    ' &middot; Finition ' + secs(h.BootPostBootTimeMs) + ' &middot; ' + col('Total ' + secs(h.BootTimeMs), h.IsSlow ? 'warn' : ''));
            });
        }
        if (p.bootPerf.Stats) {
            var st = p.bootPerf.Stats;
            bpRows += ddR('Moyennes', mut('Syst&egrave;me ' + secs(st.AvgMainPathMs) + ' &middot; Finition ' + secs(st.AvgPostBootMs) + ' &middot; Total ' + secs(st.AvgBootTimeMs)));
        }
    }

    var bootContent = ddG('D&eacute;marrages r&eacute;cents') + bootRows + ddG('Boot performance') + bpRows;

    // ========================================================
    // SECURITE : services surveilles (EDR, AV, etc.), un par ligne.
    // ========================================================
    var edrRows = '';
    if (!p.monitoredServices || p.monitoredServices.length === 0) {
        edrRows = ddR('Services', mut('aucun service surveill&eacute;'));
    } else {
        p.monitoredServices.forEach(function(svc) {
            var st, cls;
            if (!svc.Installed)   { st = 'NON INSTALL&Eacute;'; cls = 'ko'; }
            else if (svc.IsAlert) { st = svc.Status;           cls = 'warn'; }
            else                  { st = 'Running';            cls = 'ok'; }
            var v = tag(st, cls);
            var sSec = [];
            if (svc.ServiceName) sSec.push(svc.ServiceName);
            if (svc.Role) sSec.push('[' + svc.Role + ']');
            if (svc.Installed && svc.StartType) sSec.push('d&eacute;marrage ' + svc.StartType);
            v += joinMut(sSec);
            edrRows += ddR(svc.DisplayName, v);
        });
    }
    // v2.5.1 : client VPN (present + version). Inventaire, pas d'alerte.
    var vpnRow;
    if (!p.vpn) {
        vpnRow = ddR('Client VPN', mut('non collect&eacute; (Collector &lt; 2.5.1)'));
    } else if (p.vpn.Present) {
        var vpnSec = [];
        if (p.vpn.Version) vpnSec.push('v' + p.vpn.Version);
        vpnRow = ddR(p.vpn.Product || 'VPN', tag('Pr&eacute;sent', 'ok') + joinMut(vpnSec));
    } else {
        vpnRow = ddR('Client VPN', mut('non d&eacute;tect&eacute;'));
    }
    var secContent = ddG('Services surveill&eacute;s') + edrRows + ddG('Client VPN') + vpnRow;

    // ========================================================
    // VUE D'ENSEMBLE : verdict + alertes actives + signaux croises (mono-ligne).
    // ========================================================
    function ovAlert(lvl, icon, title, meta) {
        var lbl = (lvl === 'critical') ? 'Critique' : ((lvl === 'warning') ? 'Alerte' : 'Info');
        var vcls = (lvl === 'critical') ? 'ko' : ((lvl === 'warning') ? 'warn' : '');
        var rank = (lvl === 'critical') ? 0 : ((lvl === 'warning') ? 1 : 2);
        var v = col(title, vcls) + (meta ? ' ' + mut('&middot; ' + meta) : '');
        return { lbl: lbl, v: v, rank: rank };
    }
    var overview = [];
    if (p.pc.IsOffline) {
        overview.push(ovAlert('critical', '&#128268;', '<strong>Machine hors-ligne</strong> depuis ' + timeAgo(p.pc.CollectedAt), ''));
    }
    if (p.edrAlert) {
        var edrStatus = p.edr && p.edr.Installed ? p.edr.Status : 'NON INSTALL&Eacute;';
        var edrName = p.edr ? p.edr.DisplayName : 'EDR';
        overview.push(ovAlert('critical', '&#128737;', '<strong>EDR (' + edrName + ')</strong> en probl&egrave;me', edrStatus));
    }
    if (p.bsodCount > 0) {
        overview.push(ovAlert('critical', '&#128165;', '<strong>' + p.bsodCount + ' BSOD</strong> sur la p&eacute;riode', ''));
    }
    if (p.crashCount > 0) {
        overview.push(ovAlert('warning', '&#9888;', '<strong>' + p.crashCount + ' crash/freeze</strong> sur la p&eacute;riode', ''));
    }
    if (p.hwCount > 0) {
        overview.push(ovAlert('critical', '&#128268;', '<strong>' + p.hwCount + ' erreur(s) mat&eacute;rielle(s) fatale(s)</strong>', ''));
    }
    if (p.diskSmartAlert) {
        var reasons = [];
        p.diskSmartAlerts.forEach(function(d) { reasons.push(d.FriendlyName + ' (' + d.AlertReasons.join(', ') + ')'); });
        overview.push(ovAlert('critical', '&#128190;', '<strong>SMART alerte</strong> sur disque', reasons.join(' | ')));
    }
    if (p.diskAlertCount > 0) {
        overview.push(ovAlert('warning', '&#128190;', '<strong>Disque(s) satur&eacute;(s)</strong>', p.diskAlertCount + ' volume(s)'));
    }
    if (p.bootPerfAlert) {
        var lb2 = (p.bootPerf && p.bootPerf.LastBoot) ? p.bootPerf.LastBoot : null;
        var bootMeta = lb2 ? ((lb2.BootTimeMs / 1000).toFixed(1) + 's au dernier cold boot') : '';
        overview.push(ovAlert('warning', '&#9201;', '<strong>Boots lents</strong> (&gt; 90s post-boot)', bootMeta));
    }
    if (p.bootLongCount > 0) {
        overview.push(ovAlert('warning', '&#9201;', '<strong>' + p.bootLongCount + ' boot(s) longs</strong> (&gt; seuil)', ''));
    }
    if (p.batteryAlert && p.battery) {
        overview.push(ovAlert('warning', '&#128267;', '<strong>Batterie us&eacute;e</strong>', p.battery.HealthPercent + '% de sa capacit&eacute; d\'origine'));
    }
    if (p.oldMonitorAlert && p.oldMonitors && p.oldMonitors.length > 0) {
        var oldest = p.oldMonitors.reduce(function(a, b) { return (a.AgeYears || 0) > (b.AgeYears || 0) ? a : b; });
        overview.push(ovAlert('warning', '&#128250;', '<strong>' + p.oldMonitors.length + ' &eacute;cran(s) &acirc;g&eacute;(s)</strong>', 'le plus vieux : ' + oldest.AgeYears + ' ans'));
    }
    if (p.wheaCorrected && p.wheaCorrected.length > 0) {
        overview.push(ovAlert('info', '&#9432;', p.wheaCorrectedTotal + ' erreur(s) mat&eacute;rielle(s) corrig&eacute;es (t&eacute;l&eacute;m&eacute;trie)', p.wheaCorrected.length + ' signature(s)'));
    }

    // Verdict global (une ligne)
    var verdict = computeVerdict(p);
    var verdictCls = (verdict.cls === 'critical') ? 'ko' : ((verdict.cls === 'incident' || verdict.cls === 'watch') ? 'warn' : 'ok');
    var verdictV = col(verdict.label, verdictCls) +
        (verdict.reasons.length > 0 ? ' ' + mut('&middot; ' + verdict.reasons.join(' &middot; ')) : '');

    var ovAlertsHTML;
    if (overview.length) {
        overview.sort(function(a, b) { return a.rank - b.rank; });
        var prevOv = null;
        ovAlertsHTML = overview.map(function(o) {
            var dl = (o.lbl === prevOv) ? '' : o.lbl;
            prevOv = o.lbl;
            return ddR(dl, o.v);
        }).join('');
    } else {
        ovAlertsHTML = ddR('&Eacute;tat', col('Tout va bien', 'ok') + ' ' + mut('aucune alerte active sur cette p&eacute;riode'));
    }

    // Signaux croises (correlations temporelles) : une ligne par correlation
    var correlations = detectCorrelations(p);
    var corrRows = '';
    correlations.forEach(function(f) {
        corrRows += ddR('Signal', col(f.title, f.severity === 'crit' ? 'ko' : 'warn') + ' ' + mut(f.detail));
    });

    var ovContent =
        ddG('Synth&egrave;se') + ddR('Verdict', verdictV) +
        ddG('Alertes actives') + ovAlertsHTML +
        (corrRows ? ddG('Signaux crois&eacute;s') + corrRows : '');

    // ========================================================
    // STRUCTURE EN ONGLETS (inchangee) : compteurs + boutons + panneaux.
    // ========================================================
    var cntStability = p.crashCount + p.bsodCount + p.hwCount + (p.wheaCorrected ? p.wheaCorrected.length : 0);
    var cntBoot      = p.bootLongCount + (p.bootPerfAlert ? 1 : 0);
    var cntMaterial  = p.diskAlertCount + (p.diskSmartAlert ? p.diskSmartAlerts.length : 0) + (p.batteryAlert ? 1 : 0);
    var cntSecurity  = (p.edrAlert ? 1 : 0);
    var cntOverview  = overview.length;

    function tabBtn(key, label, count, icon) {
        var active = (key === 'overview') ? ' active' : '';
        var badge = '';
        if (count > 0) {
            badge = '<span class="tab-badge">' + count + '</span>';
        } else if (count === 0) {
            if (key !== 'overview') badge = '<span class="tab-badge quiet">0</span>';
        }
        return '<button class="detail-tab' + active + '" data-tab="' + key + '" onclick="selectDetailTab(' + idx + ', \'' + key + '\')">' +
               '<span>' + icon + '</span>' + label + badge + '</button>';
    }

    var tabsHTML =
        tabBtn('overview',  'Vue d\'ensemble', cntOverview,  '&#128270;') +
        tabBtn('stability', 'Stabilit&eacute;',cntStability,'&#128165;') +
        tabBtn('boot',      'D&eacute;marrage',cntBoot,     '&#9889;')   +
        tabBtn('material',  'Mat&eacute;riel', cntMaterial, '&#128295;') +
        tabBtn('security',  'S&eacute;curit&eacute;',cntSecurity,'&#128274;');

    function panel(key, inner, active) {
        var cls = active ? 'detail-tab-panel active' : 'detail-tab-panel';
        return '<div class="' + cls + '" data-panel="' + key + '"><div class="dd-list">' + inner + '</div></div>';
    }

    return '<tr class="row-detail" id="detail-' + idx + '">' +
        '<td colspan="' + colspan + '" style="padding:0">' +
        '<div class="detail-tabs">' + tabsHTML + '</div>' +
        panel('overview',  ovContent,   true) +
        panel('stability', stabContent, false) +
        panel('boot',      bootContent, false) +
        panel('material',  matContent,  false) +
        panel('security',  secContent,  false) +
        '</td></tr>';
}

function toggleDetail(id) {
    var detailRow = document.getElementById(id);
    var mainRow = detailRow.previousElementSibling;
    detailRow.classList.toggle('visible');
    mainRow.classList.toggle('open');
}

// ===== EXPORT CSV =====
function exportCSV() {
    var cutoff = new Date(generatedAt);
    cutoff.setDate(cutoff.getDate() - state.days);
    var searchTerm = document.getElementById('searchInput').value.trim().toLowerCase();
    var allEnriched = pcData.map(function(pc) { return enrichPC(pc, cutoff); });
    var visible = allEnriched.filter(function(p) {
        if (state.siteFilter && p.pc.Site !== state.siteFilter) return false;
        if (state.cpuFilter && !cpuFilterMatch(p)) return false;
        if (state.osFilter   && p.pc.OSProduct !== state.osFilter) return false;
        if (state.modelFilter && (p.pc.Model || '') !== state.modelFilter) return false;
        if (state.vpnFilter && (!p.vpn || p.vpn.Version !== state.vpnFilter)) return false;
        if (state.chassisFilter && !chassisMatch(p.pc, state.chassisFilter)) return false;
        if (searchTerm) {
            if (p.pc.PC.toLowerCase().indexOf(searchTerm) === -1 &&
                (p.pc.CurrentUser || '').toLowerCase().indexOf(searchTerm) === -1 &&
                (p.pc.LastLoggedUser || '').toLowerCase().indexOf(searchTerm) === -1 &&
                (p.pc.SerialNumber || '').toLowerCase().indexOf(searchTerm) === -1 &&
                (p.pc.Model || '').toLowerCase().indexOf(searchTerm) === -1 &&
                (p.pc.Manufacturer || '').toLowerCase().indexOf(searchTerm) === -1) return false;
        }
        // cockpit : filtre "Mon equipe" applique aussi a l'export CSV
        if (state.techFilter && ckTechByPc[p.pc.PC] !== state.techFilter) return false;
        if (!matchKpiFilter(p, state.kpiFilter)) return false;
        if (state.maskHealthy && state.kpiFilter !== 'anomaly' && p.score === 0) return false;
        // v2.1.12 : filtre par appli (clic sur un crasher) - garde les PC qui l'ont en crash.
        if (state.appFilter && !(p.topCrashers || []).some(function(c) { return String(c.AppName).toLowerCase() === state.appFilter.toLowerCase(); })) return false;
        return true;
    });
    sortPCs(visible);

    var headers = ['PC', 'Site', 'IP', 'NumeroSerie', 'Fabricant', 'Modele', 'Utilisateur', 'CollectorRunAs', 'Statut', 'Connexion', 'CPU', 'CategorieCPU',
                   'AnneeCPU', 'CandidatRenouvellement', 'OS', 'OSBuild', 'OSVersion', 'OSEdition',
                   'UptimeJours', 'DerniereActivite', 'Score',
                   'Crash', 'BSOD',
                   'WHEA_Fatal', 'WHEA_Corrected_Occurrences', 'WHEA_Corrected_Signatures',
                   'GPU_TDR', 'Thermal',
                   'BootsLongs',
                   'Boots_ColdBoot', 'Boots_FastStartup', 'Boots_Resume',
                   'DisquesCritiques', 'PctLibreMin', 'Warnings',
                   // v5.3 / v5.4
                   'BatterieSantePct', 'BatterieCategorie', 'BatterieCycles', 'BatterieAlerte',
                   'EDRStatus', 'EDRInstalle', 'EDRAlerte',
                   'BootTimeMs', 'PostBootMs', 'BootPerfAlerte',
                   'DiskWorstWearPct', 'DiskSmartAlerte', 'DiskSmartRaisons',
                   // v5.7
                   'MonitorsCount', 'MonitorsOldCount', 'MonitorsList',
                   // v2.5.1 : client VPN (present + version)
                   'VpnPresent', 'VpnProduct', 'VpnVersion'];
    var rows = [headers.join(';')];

    visible.forEach(function(p) {
        var pc = p.pc;
        var minPctFree = p.diskInfo.length > 0
            ? p.diskInfo.reduce(function(m, d) { return d.PctFree < m ? d.PctFree : m; }, 100)
            : '';
        var bt = p.bootsByType || {};

        // v5.3 / v5.4 helpers
        var battPct   = (p.battery && p.battery.HasBattery) ? p.battery.HealthPercent : '';
        var battCat   = (p.battery && p.battery.HasBattery) ? p.battery.HealthCategory : '';
        var battCycle = (p.battery && p.battery.HasBattery && p.battery.CycleCount != null) ? p.battery.CycleCount : '';
        var battAlert = p.batteryAlert ? 'OUI' : 'NON';
        var edrStat    = p.edr ? p.edr.Status : '';
        var edrInst    = p.edr ? (p.edr.Installed ? 'OUI' : 'NON') : '';
        var edrAlertStr = p.edrAlert ? 'OUI' : 'NON';
        var btMs      = (p.bootPerf && p.bootPerf.LastBoot) ? p.bootPerf.LastBoot.BootTimeMs : '';
        var pbMs      = (p.bootPerf && p.bootPerf.LastBoot) ? p.bootPerf.LastBoot.BootPostBootTimeMs : '';
        var bpAlert   = p.bootPerfAlert ? 'OUI' : 'NON';
        var worstWear = (p.diskWorstWear != null) ? p.diskWorstWear : '';
        var smartAlert= p.diskSmartAlert ? 'OUI' : 'NON';
        var smartRaisons = '';
        if (p.diskSmartAlerts && p.diskSmartAlerts.length > 0) {
            smartRaisons = p.diskSmartAlerts.map(function(d) {
                return d.FriendlyName + ':' + d.AlertReasons.join('+');
            }).join(' | ');
        }

        // v5.7 helpers : moniteurs
        var monCount = (p.monitors || []).length;
        var monOldCount = (p.oldMonitors || []).length;
        var monList = '';
        if (monCount > 0) {
            monList = p.monitors.map(function(m) {
                if (!monIsIdentified(m)) return 'Ecran non identifie (EDID non transmis)';
                var s = (m.Manufacturer || '?') + ' ' + (m.Model || m.ProductCode || '?');
                if (m.SerialNumber) s += ' [SN:' + m.SerialNumber + ']';
                if (m.YearOfManufacture) s += ' (' + m.YearOfManufacture + ')';
                return s;
            }).join(' | ');
        }

        // v2.5.1 : client VPN (present + version). '' si Collector < 2.5.1.
        var vpnPresent = p.vpn ? (p.vpn.Present ? 'OUI' : 'NON') : '';
        var vpnProduct = (p.vpn && p.vpn.Present) ? p.vpn.Product : '';
        var vpnVersion = (p.vpn && p.vpn.Present) ? p.vpn.Version : '';

        var csvUser = (pc.CurrentUser && pc.CurrentUser !== '(aucune session)')
            ? pc.CurrentUser
            : (pc.LastLoggedUser ? pc.LastLoggedUser + ' (dernier)' : (pc.CurrentUser || ''));
        var row = [
            pc.PC, pc.Site, pc.IP, pc.SerialNumber || '', pc.Manufacturer || '', pc.Model || '', csvUser, pc.CollectorRunAs || '',
            pc.IsOffline ? 'OFFLINE' : 'OK',
            pc.ConnectionType || '',
            pc.CPUName || '', pc.CPUAgeCategory || '', pc.CPUYear || '',
            (RENEWAL_MODE ? (isRenewalCandidate(pc) ? 'OUI' : 'NON') : ''),
            pc.OSProduct || '', (pc.OSBuild != null ? pc.OSBuild : ''), pc.OSDisplayVersion || '', pc.OSEdition || '',
            pc.UptimeDays != null ? pc.UptimeDays : '',
            pc.CollectedAt, p.score,
            p.crashCount, p.bsodCount,
            p.wheaFatal.length, p.wheaCorrectedTotal, p.wheaCorrected.length,
            p.gpuTDR.length, p.thermal.length,
            p.bootLongCount,
            bt.ColdBoot || 0, bt.FastStartup || 0, bt.Resume || 0,
            p.diskAlertCount, minPctFree, p.warningCount,
            // v5.3 / v5.4
            battPct, battCat, battCycle, battAlert,
            edrStat, edrInst, edrAlertStr,
            btMs, pbMs, bpAlert,
            worstWear, smartAlert, smartRaisons,
            // v5.7
            monCount, monOldCount, monList,
            // v2.5.1
            vpnPresent, vpnProduct, vpnVersion
        ];
        rows.push(row.map(function(v) {
            var s = String(v).replace(/"/g, '""');
            return (s.indexOf(';') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1) ? '"' + s + '"' : s;
        }).join(';'));
    });

    var csv = '\uFEFF' + rows.join('\n');   // BOM UTF-8 pour Excel FR
    var blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    var dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    a.download = 'PCPulse-Export-' + dateStr + '.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// ===== INIT =====
initSiteDropdown();
initOsDropdown();
initModelDropdown();
initVpnDropdown();
var initBtn = document.querySelectorAll('.range-btn')[0];
state.daysBtn = initBtn;
initBtn.classList.add('active');

if (state.maskHealthy) {
    document.getElementById('maskHealthyBtn').classList.add('active');
}

// v5.5 : restaurer la preference "Vue detaillee" depuis localStorage
try {
    if (localStorage.getItem('pcmon-advcols') === '1') {
        document.body.classList.add('advanced-cols');
        document.getElementById('advancedColsBtn').classList.add('active');
    }
} catch(e) {}

// v2.1.2 : curseur de seuil d'age des ecrans secondaires (persiste en localStorage)
(function initScreenAgeSlider() {
    var sl  = document.getElementById('screenAgeSlider');
    var lbl = document.getElementById('screenAgeValue');
    if (!sl) return;
    sl.value = screenAgeThreshold;
    if (lbl) lbl.textContent = screenAgeThreshold;
    sl.addEventListener('input', function() {
        screenAgeThreshold = parseInt(sl.value, 10);
        if (lbl) lbl.textContent = screenAgeThreshold;
        try { localStorage.setItem('pcpulse_screenAge', screenAgeThreshold); } catch(e) {}
        render();   // re-enrichit tous les PC avec le nouveau seuil -> carte + KPI + detail realignes
    });
})();

// v2.1.12 : delegation de clic sur le panneau Top Crashers. Le container
// #globalCrashers est statique (seul son innerHTML change au render), donc un
// seul listener suffit et survit a tous les re-rendus.
(function() {
    var gc = document.getElementById('globalCrashers');
    if (gc) {
        gc.addEventListener('click', function(e) {
            var row = (e.target && e.target.closest) ? e.target.closest('.global-crasher-row') : null;
            if (row && row.dataset && row.dataset.appname) {
                toggleAppFilter(row.dataset.appname);
            }
        });
    }
})();

try {
    render();
} catch(e) {
    // v5.5 : on affiche l'erreur dans la zone des groupes KPI (plus visible)
    var errTarget = document.getElementById('kpiGroups') || document.getElementById('kpiGrid');
    errTarget.innerHTML =
        '<div style="color:var(--red);padding:24px;font-size:14px;background:var(--bg-danger);border-radius:10px">' +
        '<strong>Erreur JS :</strong> ' + e.message + '<br>' + e.stack + '</div>';
}

window.addEventListener('scroll', function() {
    var b = document.getElementById('backToTop');
    if (b) b.style.display = (window.scrollY > 400) ? 'flex' : 'none';
});
</script>

<button id="backToTop" onclick="window.scrollTo({top:0,behavior:'smooth'})" title="Remonter en haut" aria-label="Remonter en haut">&#8593;</button>

</body>
</html>
"@

# ============================================================
# EXPORT ET OUVERTURE
# ============================================================
# v2.1.3 : ecriture atomique (.tmp puis rename) - une consultation ne tombe jamais
# sur un HTML a moitie ecrit pendant la generation (important en mode publie/tache).
$tmpOut = "$OutputHTML.tmp"
Set-Content -Path $tmpOut -Value $html -Encoding UTF8
Move-Item -Path $tmpOut -Destination $OutputHTML -Force
Write-Host "[+] Dashboard genere : $OutputHTML" -ForegroundColor Green

# v2.4.7 : RETENTION des dashboards horodates. En mode interactif (sans -OutputPath),
# chaque run cree un "PCPulse-Dashboard-<horodatage>.html" -> ils s'accumulaient (30+).
# On ne garde que les $KeepDashboards plus recents (celui qu'on vient d'ecrire inclus).
# En mode -OutputPath (tache/publication) il n'y a qu'un fichier fixe : rien a purger.
if (-not $OutputPath) {
    $KeepDashboards = 10
    try {
        $outDir = Split-Path -Parent $OutputHTML
        @(Get-ChildItem -Path $outDir -Filter 'PCPulse-Dashboard-*.html' -File -ErrorAction Stop |
            Sort-Object LastWriteTime -Descending |
            Select-Object -Skip $KeepDashboards) |
            ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
    } catch {
        Write-Host "[!] Nettoyage des anciens dashboards ignore : $_" -ForegroundColor DarkYellow
    }
}

# v2.1.3 : pas d'ouverture navigateur en mode tache (-NoLaunch). En interactif, on ouvre comme avant.
if (-not $NoLaunch) {
    Write-Host "[*] Ouverture dans le navigateur..." -ForegroundColor Cyan
    Start-Process -FilePath $OutputHTML
}
